"""
title: SharePoint Tool
author: OpenWebUI User
version: 2.3.1
required_open_webui_version: 0.5.0
requirements: aiohttp, msal, cryptography, pypdf
description: >
    Search and read Microsoft SharePoint sites, pages, and documents via the
    Microsoft Graph API using delegated permissions.

    Authentication priority (automatic):
      1. MSAL per-user token cache  (survives SSO token expiry, auto-refreshes)
      2. OpenWebUI SSO token        (__oauth_token__ from Microsoft login)

    When both are missing or expired, type "connect my microsoft account" in chat
    to trigger a device-code sign-in link — authenticating once covers the
    Calendar, Mail, and SharePoint tools simultaneously.

    ADMIN SETUP (one-time):
      1. Azure App Registration > Authentication > Allow public client flows = Yes
      2. API Permissions > Delegated: Sites.Read.All, Files.Read.All (already granted)
      3. Set Valves: azure_client_id, azure_tenant_id, token_cache_dir
"""

from __future__ import annotations

import asyncio
import base64
import html
import io
import json
import os
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional
from datetime import timedelta, datetime, timezone

import ssl

import aiohttp
import msal
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from pydantic import BaseModel, Field

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SEARCH_ENDPOINT = f"{GRAPH_BASE}/search/query"

# Combined scopes for all three tools — one sign-in covers everything.
_ALL_TOOL_SCOPES = [
    "offline_access",
    "Mail.Read",
    "Mail.Send",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Sites.Read.All",
    "Files.Read.All",
]
# Scopes required specifically by this tool (used for silent cache lookups).
_TOOL_SCOPES = ["Sites.Read.All", "Files.Read.All"]


class _HTMLStripper(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []
        self._skip_tags = {"script", "style"}
        self._current_skip = 0

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in self._skip_tags:
            self._current_skip += 1

    def handle_endtag(self, tag: str) -> None:
        if tag in self._skip_tags and self._current_skip > 0:
            self._current_skip -= 1

    def handle_data(self, data: str) -> None:
        if self._current_skip == 0:
            stripped = data.strip()
            if stripped:
                self._parts.append(stripped)

    def get_text(self) -> str:
        return " ".join(self._parts)


def _strip_html(raw: str) -> str:
    try:
        s = _HTMLStripper()
        s.feed(html.unescape(raw or ""))
        return s.get_text()
    except Exception:
        return raw


class Tools:
    # ------------------------------------------------------------------
    # Valves
    # ------------------------------------------------------------------
    class Valves(BaseModel):
        azure_client_id: str = Field(
            default="",
            description="Azure App Registration Client ID. Required for device-code fallback auth.",
        )
        azure_tenant_id: str = Field(
            default="common",
            description="Azure Tenant ID. Use your Directory ID for single-tenant apps.",
        )
        token_cache_dir: str = Field(
            default="/app/backend/data/outlook_tokens",
            description=(
                "Server directory for per-user MSAL token caches. "
                "Must match across all three tools (Mail, Calendar, SharePoint) "
                "so one sign-in covers all."
            ),
        )
        max_results: int = Field(
            default=20,
            description="Default maximum number of results returned per request.",
        )
        max_page_content_chars: int = Field(
            default=8000,
            description="Maximum characters of page text returned to the LLM.",
        )
        request_timeout_seconds: int = Field(
            default=30,
            description="HTTP timeout in seconds for Graph API calls.",
        )
        openwebui_base_url: str = Field(
            default="",
            description=(
                "Base URL of this OpenWebUI instance, e.g. 'http://localhost:3000'. "
                "Leave empty to auto-detect from the incoming request."
            ),
        )
        max_document_size_mb: int = Field(
            default=20,
            description="Maximum SharePoint file size (MB) allowed for content extraction.",
        )
        verify_ssl: bool = Field(
            default=True,
            description=(
                "Verify TLS certificates on outbound HTTPS calls (Graph API and the "
                "OpenWebUI Files API). Turn OFF only on internal networks where a "
                "corporate TLS-inspection proxy presents certs that are not in the "
                "container's trust store (symptom: 'Cannot connect to host ... ssl:default')."
            ),
        )

    # ------------------------------------------------------------------
    # Init
    # ------------------------------------------------------------------
    def __init__(self) -> None:
        self.valves = self.Valves()

    # ------------------------------------------------------------------
    # aiohttp connector (honours the verify_ssl valve)
    # ------------------------------------------------------------------
    def _connector(self) -> aiohttp.TCPConnector:
        if self.valves.verify_ssl:
            return aiohttp.TCPConnector()
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return aiohttp.TCPConnector(ssl=ctx)

    # ------------------------------------------------------------------
    # MSAL token cache helpers
    # ------------------------------------------------------------------
    def _cache_path(self, user_id: str) -> Path:
        d = Path(self.valves.token_cache_dir)
        d.mkdir(parents=True, exist_ok=True)
        return d / f"{user_id}.json"

    def _get_fernet(self) -> Optional[Fernet]:
        key_material = os.environ.get("OPENWEBUI_TOKEN_CACHE_KEY", "").strip()
        if not key_material:
            return None
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=b"openwebui_m365_token_cache_v1",
            iterations=100_000,
        )
        return Fernet(base64.urlsafe_b64encode(kdf.derive(key_material.encode())))

    def _load_cache(self, user_id: str) -> msal.SerializableTokenCache:
        cache = msal.SerializableTokenCache()
        p = self._cache_path(user_id)
        if not p.exists():
            return cache
        raw = p.read_bytes()
        fernet = self._get_fernet()
        if fernet:
            try:
                data = fernet.decrypt(raw).decode("utf-8")
            except Exception:
                try:  # migration: file was written unencrypted before key was set
                    data = raw.decode("utf-8")
                    cache.deserialize(data)
                    self._save_cache(user_id, cache)  # re-save encrypted
                    return cache
                except Exception:
                    return cache
        else:
            try:
                data = raw.decode("utf-8")
            except Exception:
                return cache
        try:
            cache.deserialize(data)
        except Exception:
            pass
        return cache

    def _save_cache(self, user_id: str, cache: msal.SerializableTokenCache) -> None:
        if cache.has_state_changed:
            serialized = cache.serialize()
            fernet = self._get_fernet()
            if fernet:
                self._cache_path(user_id).write_bytes(
                    fernet.encrypt(serialized.encode("utf-8"))
                )
            else:
                self._cache_path(user_id).write_text(serialized, encoding="utf-8")

    def _msal_app(
        self, cache: msal.SerializableTokenCache
    ) -> msal.PublicClientApplication:
        if not self.valves.azure_client_id:
            raise ValueError(
                "azure_client_id is not set in tool Valves. "
                "An administrator must configure it."
            )
        return msal.PublicClientApplication(
            client_id=self.valves.azure_client_id,
            authority=f"https://login.microsoftonline.com/{self.valves.azure_tenant_id}",
            token_cache=cache,
        )

    def _token_from_cache(self, user_id: str) -> Optional[str]:
        try:
            cache = self._load_cache(user_id)
            app = self._msal_app(cache)
            accounts = app.get_accounts()
            if not accounts:
                return None
            result = app.acquire_token_silent(_TOOL_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache(user_id, cache)
                return result["access_token"]
        except Exception:
            pass
        return None

    # ------------------------------------------------------------------
    # Background device-code polling
    # ------------------------------------------------------------------
    async def _poll_device_flow(
        self,
        user_id: str,
        flow: dict,
        cache: msal.SerializableTokenCache,
    ) -> None:
        loop = asyncio.get_event_loop()
        app = self._msal_app(cache)
        try:
            result = await loop.run_in_executor(
                None, lambda: app.acquire_token_by_device_flow(flow)
            )
            if result and "access_token" in result:
                self._save_cache(user_id, cache)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Header builder — MSAL cache → SSO token → error
    # ------------------------------------------------------------------
    def _get_headers(self, user_id: Optional[str], oauth_token: Optional[dict]) -> dict:
        access_token: Optional[str] = None

        if user_id:
            access_token = self._token_from_cache(user_id)

        if not access_token and oauth_token:
            access_token = oauth_token.get("access_token")

        if not access_token:
            raise ValueError(
                "Not authenticated with Microsoft. "
                "Type 'connect my microsoft account' to sign in."
            )

        return {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    # ------------------------------------------------------------------
    # Other private helpers (unchanged)
    # ------------------------------------------------------------------
    def _format_search_hit(self, hit: dict) -> dict:
        resource = hit.get("resource", {})
        odata_type = resource.get("@odata.type", "")
        result = {
            "rank": hit.get("rank", 0),
            "summary": hit.get("summary", ""),
            "type": odata_type.split(".")[-1] if odata_type else "unknown",
        }
        if "site" in odata_type:
            result.update(
                {
                    "id": resource.get("id", ""),
                    "name": resource.get("displayName", ""),
                    "url": resource.get("webUrl", ""),
                    "description": resource.get("description", ""),
                }
            )
        elif "driveItem" in odata_type:
            result.update(
                {
                    "id": resource.get("id", ""),
                    "name": resource.get("name", ""),
                    "url": resource.get("webUrl", ""),
                    "size": resource.get("size", 0),
                    "last_modified": resource.get("lastModifiedDateTime", ""),
                    "site_id": resource.get("parentReference", {}).get("siteId", ""),
                    "drive_id": resource.get("parentReference", {}).get("driveId", ""),
                }
            )
        elif "listItem" in odata_type:
            fields = resource.get("fields", {})
            result.update(
                {
                    "id": resource.get("id", ""),
                    "name": fields.get("Title", fields.get("FileLeafRef", "")),
                    "url": resource.get("webUrl", ""),
                    "last_modified": resource.get("lastModifiedDateTime", ""),
                    "site_id": resource.get("parentReference", {}).get("siteId", ""),
                }
            )
        return result

    def _format_drive_item(self, item: dict) -> dict:
        is_folder = "folder" in item
        parent = item.get("parentReference", {})
        result = {
            "id": item.get("id", ""),
            "name": item.get("name", ""),
            "url": item.get("webUrl", ""),
            "is_folder": is_folder,
            "size_bytes": item.get("size", 0),
            "created": item.get("createdDateTime", ""),
            "last_modified": item.get("lastModifiedDateTime", ""),
            "created_by": item.get("createdBy", {})
            .get("user", {})
            .get("displayName", ""),
            "modified_by": item.get("lastModifiedBy", {})
            .get("user", {})
            .get("displayName", ""),
            "drive_id": parent.get("driveId", ""),
        }
        if not is_folder:
            result["mime_type"] = item.get("file", {}).get("mimeType", "")
        if is_folder:
            result["child_count"] = item.get("folder", {}).get("childCount", 0)
        return result

    async def _resolve_site_url(
        self, session: aiohttp.ClientSession, headers: dict, site_id: str
    ) -> str:
        async with session.get(
            f"{GRAPH_BASE}/sites/{site_id}",
            headers=headers,
            params={"$select": "webUrl"},
        ) as resp:
            if resp.ok:
                return (await resp.json()).get("webUrl", "")
        return ""

    async def _search_post(
        self,
        session: aiohttp.ClientSession,
        headers: dict,
        payload: dict,
    ) -> tuple[int, object]:
        """POST to the search endpoint with one automatic retry on 429 (rate limit)."""
        for attempt in range(2):
            async with session.post(
                SEARCH_ENDPOINT, headers=headers, json=payload
            ) as resp:
                if resp.status == 429 and attempt == 0:
                    retry_after = min(int(resp.headers.get("Retry-After", "10")), 30)
                    await asyncio.sleep(retry_after)
                    continue
                if resp.ok:
                    return resp.status, await resp.json()
                return resp.status, await resp.text()
        return 429, "Search service is rate-limited. Please try again in a moment."

    def _owui_base_url(self, request: Optional[object]) -> str:
        """Resolve the OpenWebUI base URL from Valve, then request, then localhost fallback."""
        if self.valves.openwebui_base_url:
            return self.valves.openwebui_base_url.rstrip("/")
        if request:
            try:
                return str(request.base_url).rstrip("/")
            except Exception:
                pass
        return "http://localhost:3000"

    def _owui_user_token(self, user: Optional[dict], request: Optional[object]) -> str:
        """Resolve the logged-in user's OpenWebUI bearer credential.

        Accepts both JWT session tokens and ``sk-`` API keys — OpenWebUI's
        ``get_verified_user`` dependency honours both.
        """
        token = ((user or {}).get("token") or "").strip()
        if token:
            return token

        if request:
            try:
                auth_header = request.headers.get("authorization", "")
                if auth_header.lower().startswith("bearer "):
                    token = auth_header[7:].strip()
                    if token:
                        return token
            except Exception:
                pass

            try:
                token = (request.cookies.get("token") or "").strip()
                if token:
                    return token
            except Exception:
                pass

        return ""

    async def _delete_owui_file(
        self, base_url: str, headers: dict, file_id: str
    ) -> None:
        """Delete a temporary file from the OpenWebUI Files API (best-effort cleanup)."""
        try:
            async with aiohttp.ClientSession(
                timeout=aiohttp.ClientTimeout(total=10),
                connector=self._connector(),
            ) as session:
                await session.delete(
                    f"{base_url}/api/v1/files/{file_id}", headers=headers
                )
        except Exception:
            pass

    def _extract_pdf_text(self, file_bytes: bytes) -> str:
        """Best-effort local PDF text extraction fallback for text-based PDFs."""
        try:
            from pypdf import PdfReader

            reader = PdfReader(io.BytesIO(file_bytes))
            pages: list[str] = []
            for page in reader.pages:
                text = (page.extract_text() or "").strip()
                if text:
                    pages.append(text)
            return "\n\n".join(pages)
        except Exception:
            return ""

    def _owui_error_message(self, response_text: str) -> str:
        """Extract a useful error message from an OpenWebUI JSON or text response."""
        try:
            payload = json.loads(response_text)
            if isinstance(payload, dict):
                return str(
                    payload.get("detail")
                    or payload.get("error")
                    or payload.get("message")
                    or response_text
                )
        except Exception:
            pass
        return response_text

    # ------------------------------------------------------------------
    # PUBLIC TOOL FUNCTIONS
    # ------------------------------------------------------------------

    async def authenticate_with_microsoft(
        self,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Connect your Microsoft account to all Outlook tools (Mail, Calendar, SharePoint).
        Call this when you see an authentication error or when asked to sign in.
        Authenticating once here covers all three tools automatically.

        Returns:
            A sign-in URL and short code. Open the URL in any browser, enter the code,
            and sign in with your Microsoft account. Then repeat your original request.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")

            if self._token_from_cache(user_id):
                return json.dumps(
                    {
                        "status": "already_authenticated",
                        "message": "Your Microsoft account is already connected. You can use all Outlook tools.",
                    }
                )

            cache = self._load_cache(user_id)
            loop = asyncio.get_event_loop()
            app = self._msal_app(cache)

            flow = await loop.run_in_executor(
                None, lambda: app.initiate_device_flow(scopes=_ALL_TOOL_SCOPES)
            )

            if "user_code" not in flow:
                err = flow.get("error_description", flow.get("error", "Unknown error"))
                return json.dumps({"error": f"Could not start sign-in: {err}"})

            asyncio.create_task(self._poll_device_flow(user_id, flow, cache))

            expires_min = flow.get("expires_in", 900) // 60
            return json.dumps(
                {
                    "status": "authentication_required",
                    "message": (
                        f"**Sign in to Microsoft:**\n\n"
                        f"1. Open: {flow['verification_uri']}\n"
                        f"2. Enter code: **{flow['user_code']}**\n\n"
                        f"Code expires in {expires_min} minutes. "
                        f"After signing in, repeat your original request."
                    ),
                    "sign_in_url": flow["verification_uri"],
                    "user_code": flow["user_code"],
                },
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except Exception as exc:
            return json.dumps({"error": f"Authentication setup failed: {exc}"})

    async def disconnect_microsoft_account(
        self,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Disconnect your Microsoft account by deleting the saved login session.
        Use this to sign out, switch accounts, or revoke cached credentials.
        To fully revoke app permissions, also visit https://myapps.microsoft.com.

        Returns:
            JSON indicating whether a session was found and deleted.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            p = self._cache_path(user_id)
            if p.exists():
                p.unlink()
                return json.dumps(
                    {
                        "success": True,
                        "message": (
                            "Microsoft account disconnected. Token cache deleted from server. "
                            "To fully revoke app permissions, visit https://myapps.microsoft.com"
                        ),
                    }
                )
            return json.dumps(
                {"success": True, "message": "No active Microsoft session found."}
            )
        except Exception as exc:
            return json.dumps({"error": f"Failed to disconnect: {exc}"})

    async def search_sharepoint(
        self,
        query: str,
        content_type: str = "all",
        max_results: int = 20,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Search across all SharePoint sites, pages, and documents the user has access to.
        Use this for "Where on SharePoint can I find X?" or "Is there a document about Y?".

        Args:
            query: Search keywords or phrase.
            content_type: 'all' (default), 'sites', 'documents', or 'pages'.
            max_results: Maximum results to return (default 20).

        Returns:
            JSON with ranked results including name, URL, summary, and site ID.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            top = min(max_results, self.valves.max_results * 2)

            type_map = {
                "all": ["site", "driveItem", "listItem"],
                "sites": ["site"],
                "documents": ["driveItem"],
                "pages": ["listItem"],
            }
            entity_types = type_map.get(content_type.lower(), type_map["all"])

            payload = {
                "requests": [
                    {
                        "entityTypes": entity_types,
                        "query": {"queryString": query},
                        "fields": [
                            "id",
                            "name",
                            "displayName",
                            "webUrl",
                            "description",
                            "lastModifiedDateTime",
                            "size",
                            "summary",
                            "parentReference",
                            "fields",
                        ],
                        "size": top,
                    }
                ]
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout, connector=self._connector()) as session:
                status, data = await self._search_post(session, headers, payload)

            if status == 401:
                return json.dumps(
                    {
                        "error": "Session expired. Type 'connect my microsoft account' to re-authenticate."
                    }
                )
            if status == 403:
                return json.dumps(
                    {
                        "error": "Permission denied. Sites.Read.All and Files.Read.All are required."
                    }
                )
            if status == 429:
                return json.dumps(
                    {
                        "error": "Microsoft Search is temporarily rate-limited. Please wait a moment and try again."
                    }
                )
            if status != 200:
                return json.dumps({"error": f"Graph API error {status}: {data}"})

            hits = [
                self._format_search_hit(hit)
                for rv in data.get("value", [])
                for container in rv.get("hitsContainers", [])
                for hit in container.get("hits", [])
            ]

            return json.dumps(
                {
                    "query": query,
                    "content_type": content_type,
                    "count": len(hits),
                    "results": hits,
                },
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def get_site_pages(
        self,
        site_id: str,
        max_results: int = 30,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        List all pages published on a SharePoint site.

        Args:
            site_id: SharePoint site ID (from search_sharepoint results).
            max_results: Maximum pages to return (default 30).

        Returns:
            JSON with page titles, URLs, dates, and page IDs for use with get_page_content.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout, connector=self._connector()) as session:
                async with session.get(
                    f"{GRAPH_BASE}/sites/{site_id}/pages",
                    headers=headers,
                    params={
                        "$select": "id,title,webUrl,lastModifiedDateTime,createdDateTime,publishingState",
                        "$top": str(min(max_results, 100)),
                        "$orderby": "lastModifiedDateTime desc",
                    },
                ) as resp:
                    if resp.status == 401:
                        return json.dumps(
                            {
                                "error": "Session expired. Type 'connect my microsoft account' to re-authenticate."
                            }
                        )
                    if resp.status == 403:
                        return json.dumps(
                            {"error": "Permission denied. Sites.Read.All required."}
                        )
                    if resp.status == 404:
                        return json.dumps({"error": f"Site '{site_id}' not found."})
                    if not resp.ok:
                        return json.dumps(
                            {
                                "error": f"Graph API error {resp.status}: {await resp.text()}"
                            }
                        )
                    data = await resp.json()

            pages = [
                {
                    "id": p.get("id", ""),
                    "title": p.get("title", "(untitled)"),
                    "url": p.get("webUrl", ""),
                    "last_modified": p.get("lastModifiedDateTime", ""),
                    "created": p.get("createdDateTime", ""),
                    "published": p.get("publishingState", {}).get("level", "")
                    == "published",
                }
                for p in data.get("value", [])
            ]

            return json.dumps(
                {"site_id": site_id, "count": len(pages), "pages": pages},
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def get_page_content(
        self,
        site_id: str,
        page_id: str,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Read the full text content of a specific SharePoint page.
        Use this for "What does page X say about Y?" or "Is Z mentioned on this page?".

        Args:
            site_id: SharePoint site ID (from search_sharepoint).
            page_id: Page ID (from get_site_pages).

        Returns:
            JSON with page title, URL, and full extracted plain-text content.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout, connector=self._connector()) as session:
                async with session.get(
                    f"{GRAPH_BASE}/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage",
                    headers=headers,
                    params={"$expand": "canvasLayout"},
                ) as resp:
                    if resp.status == 401:
                        return json.dumps(
                            {
                                "error": "Session expired. Type 'connect my microsoft account' to re-authenticate."
                            }
                        )
                    if resp.status == 403:
                        return json.dumps(
                            {"error": "Permission denied. Sites.Read.All required."}
                        )
                    if resp.status == 404:
                        return json.dumps(
                            {
                                "error": f"Page '{page_id}' not found on site '{site_id}'."
                            }
                        )
                    if not resp.ok:
                        return json.dumps(
                            {
                                "error": f"Graph API error {resp.status}: {await resp.text()}"
                            }
                        )
                    page = await resp.json()

            text_parts = [
                _strip_html(wp.get("innerHtml", ""))
                for section in page.get("canvasLayout", {}).get(
                    "horizontalSections", []
                )
                for col in section.get("columns", [])
                for wp in col.get("webparts", [])
                if wp.get("innerHtml", "").strip()
            ]

            full_text = "\n\n".join(filter(None, text_parts))
            if len(full_text) > self.valves.max_page_content_chars:
                full_text = (
                    full_text[: self.valves.max_page_content_chars]
                    + "\n\n[... content truncated ...]"
                )

            return json.dumps(
                {
                    "site_id": site_id,
                    "page_id": page_id,
                    "title": page.get("title", ""),
                    "url": page.get("webUrl", ""),
                    "last_modified": page.get("lastModifiedDateTime", ""),
                    "content": full_text,
                    "content_length": len(full_text),
                },
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def list_documents(
        self,
        site_id: str,
        folder_path: Optional[str] = None,
        max_results: int = 50,
        newest_first: bool = True,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        List documents and files in a SharePoint site's document library.
        Use this for "Which documents on site X are new?" or "What files are in folder Y?".

        Args:
            site_id: SharePoint site ID (from search_sharepoint).
            folder_path: Subfolder path, e.g. 'General' or 'Projects/2024'. Empty = root.
            max_results: Maximum items to return (default 50).
            newest_first: Sort by most recently modified (default True).

        Returns:
            JSON list of files/folders with name, URL, size, dates, and author.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            top = min(max_results, 200)

            url = (
                f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{folder_path.strip('/')}:/children"
                if folder_path
                else f"{GRAPH_BASE}/sites/{site_id}/drive/root/children"
            )

            params = {
                "$select": "id,name,webUrl,size,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,file,folder",
                "$top": str(top),
            }
            if newest_first:
                params["$orderby"] = "lastModifiedDateTime desc"

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout, connector=self._connector()) as session:
                async with session.get(url, headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps(
                            {
                                "error": "Session expired. Type 'connect my microsoft account' to re-authenticate."
                            }
                        )
                    if resp.status == 403:
                        return json.dumps(
                            {"error": "Permission denied. Files.Read.All required."}
                        )
                    if resp.status == 404:
                        return json.dumps(
                            {
                                "error": f"Site or folder not found. site_id='{site_id}', folder='{folder_path}'."
                            }
                        )
                    if not resp.ok:
                        return json.dumps(
                            {
                                "error": f"Graph API error {resp.status}: {await resp.text()}"
                            }
                        )
                    data = await resp.json()

            items = [self._format_drive_item(i) for i in data.get("value", [])]
            return json.dumps(
                {
                    "site_id": site_id,
                    "folder": folder_path or "(root)",
                    "count": len(items),
                    "newest_first": newest_first,
                    "items": items,
                },
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def search_in_site(
        self,
        site_id: str,
        query: str,
        content_type: str = "all",
        max_results: int = 20,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Search for pages or documents within a specific SharePoint site.
        Use this for "Is there anything about Z on site X?" or "Find documents about Y on site X".

        Args:
            site_id: SharePoint site ID to search within (from search_sharepoint).
            query: Search keywords or phrase.
            content_type: 'all' (default), 'documents', or 'pages'.
            max_results: Maximum results to return (default 20).

        Returns:
            JSON with search results scoped to the specified site.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            top = min(max_results, 50)

            type_map = {
                "all": ["driveItem", "listItem"],
                "documents": ["driveItem"],
                "pages": ["listItem"],
            }
            entity_types = type_map.get(content_type.lower(), type_map["all"])

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout, connector=self._connector()) as session:
                site_url = await self._resolve_site_url(session, headers, site_id)
                if not site_url:
                    return json.dumps(
                        {"error": f"Could not resolve URL for site '{site_id}'."}
                    )

                # Graph Search only accepts `contentSources` for entity type
                # `externalItem` (Microsoft Search connectors). For native
                # driveItem/listItem we must scope the query via a KQL
                # `Path:` clause in the query string itself.
                scoped_query = f'({query}) AND Path:"{site_url}"'

                payload = {
                    "requests": [
                        {
                            "entityTypes": entity_types,
                            "query": {"queryString": scoped_query},
                            "fields": [
                                "id",
                                "name",
                                "webUrl",
                                "lastModifiedDateTime",
                                "size",
                                "parentReference",
                                "fields",
                            ],
                            "size": top,
                        }
                    ]
                }

                status, data = await self._search_post(session, headers, payload)

            if status == 401:
                return json.dumps(
                    {
                        "error": "Session expired. Type 'connect my microsoft account' to re-authenticate."
                    }
                )
            if status == 403:
                return json.dumps(
                    {"error": "Permission denied. Sites.Read.All required."}
                )
            if status == 429:
                return json.dumps(
                    {
                        "error": "Microsoft Search is temporarily rate-limited. Please wait a moment and try again."
                    }
                )
            if status != 200:
                return json.dumps({"error": f"Graph API error {status}: {data}"})

            hits = [
                self._format_search_hit(hit)
                for rv in data.get("value", [])
                for container in rv.get("hitsContainers", [])
                for hit in container.get("hits", [])
            ]

            return json.dumps(
                {
                    "site_id": site_id,
                    "site_url": site_url,
                    "query": query,
                    "content_type": content_type,
                    "count": len(hits),
                    "results": hits,
                },
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def view_file(
        self,
        item_id: str,
        file_name: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Download a SharePoint document (PDF, DOCX, XLSX, PPTX, etc.) and add it to this
        chat as an attachment, then return its extracted plain-text content.
        Use this when the user wants to read, summarise, or ask questions about a specific file.

        How to obtain the arguments:
          - item_id  : the 'id' field of any driveItem from list_documents or search_sharepoint.
          - file_name: the 'name' field of that same item (including extension).
          - drive_id : the 'drive_id' field from search_sharepoint / list_documents results (preferred).
          - site_id  : the SharePoint site ID — required when drive_id is not available.

        Args:
            item_id:   Drive item ID of the file.
            file_name: File name including extension, e.g. 'Q3_Report.pdf'.
            site_id:   SharePoint site ID (use when drive_id is unavailable).
            drive_id:  Drive ID (preferred — use when returned by search_sharepoint or list_documents).

        Returns:
            JSON with id, filename and full extracted content. The function name
            'view_file' is recognised by OpenWebUI as a citation source, so the
            document will appear as a citation card in the chat response.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            try:
                graph_headers = self._get_headers(user_id, __oauth_token__)
            except ValueError as exc:
                if "Not authenticated with Microsoft" in str(exc):
                    return await self.authenticate_with_microsoft(
                        __user__=__user__,
                        __request__=__request__,
                    )
                raise

            if not drive_id and not site_id:
                return json.dumps(
                    {
                        "error": "Provide at least one of drive_id or site_id to locate the file."
                    }
                )

            # ----------------------------------------------------------
            # Step 1: Download file bytes from SharePoint
            # ----------------------------------------------------------
            download_url = (
                f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
                if drive_id
                else f"{GRAPH_BASE}/sites/{site_id}/drive/items/{item_id}/content"
            )

            dl_timeout = aiohttp.ClientTimeout(
                total=max(self.valves.request_timeout_seconds * 2, 60)
            )
            async with aiohttp.ClientSession(
                timeout=dl_timeout, connector=self._connector()
            ) as session:
                async with session.get(
                    download_url, headers=graph_headers, allow_redirects=True
                ) as resp:
                    if resp.status == 401:
                        return json.dumps(
                            {
                                "error": "Session expired. Type 'connect my microsoft account' to re-authenticate."
                            }
                        )
                    if resp.status == 403:
                        return json.dumps(
                            {"error": "Permission denied. Files.Read.All required."}
                        )
                    if resp.status == 404:
                        return json.dumps(
                            {
                                "error": f"File '{file_name}' not found (item_id='{item_id}')."
                            }
                        )
                    if not resp.ok:
                        return json.dumps(
                            {
                                "error": f"Graph API error {resp.status}: {await resp.text()}"
                            }
                        )
                    file_bytes = await resp.read()
                    content_type = resp.content_type or "application/octet-stream"

            if not file_bytes:
                return json.dumps({"error": "Downloaded file is empty."})

            size_mb = len(file_bytes) / (1024 * 1024)
            if size_mb > self.valves.max_document_size_mb:
                return json.dumps(
                    {
                        "error": (
                            f"File is {size_mb:.1f} MB, which exceeds the {self.valves.max_document_size_mb} MB limit. "
                            "Increase max_document_size_mb in Valves or use a smaller file."
                        )
                    }
                )

            # ----------------------------------------------------------
            # Step 2: Resolve OpenWebUI base URL and the caller's bearer token
            # ----------------------------------------------------------
            owui_token = self._owui_user_token(__user__, __request__)
            if not owui_token:
                return json.dumps(
                    {
                        "error": (
                            "No OpenWebUI bearer token is available for document processing. "
                            "Please sign in again and retry."
                        )
                    }
                )

            base_url = self._owui_base_url(__request__)
            owui_headers = {
                "Authorization": f"Bearer {owui_token}",
                "Accept": "application/json",
            }

            # ----------------------------------------------------------
            # Step 3: Upload to OpenWebUI and process synchronously in one call.
            # process=true runs the Loader + chunker + embedder; background=false
            # makes the response wait for completion and populates data.content.
            # ----------------------------------------------------------
            form = aiohttp.FormData()
            form.add_field(
                "file", file_bytes, filename=file_name, content_type=content_type
            )

            upload_url = (
                f"{base_url}/api/v1/files/"
                "?process=true&process_in_background=false"
            )
            async with aiohttp.ClientSession(
                timeout=aiohttp.ClientTimeout(total=300),
                connector=self._connector(),
            ) as session:
                async with session.post(
                    upload_url, headers=owui_headers, data=form
                ) as resp:
                    if resp.status in (401, 403):
                        return json.dumps(
                            {
                                "error": (
                                    "Unauthorised when uploading to OpenWebUI Files API. "
                                    "Please sign in again and retry."
                                )
                            }
                        )
                    if not resp.ok:
                        return json.dumps(
                            {
                                "error": f"OpenWebUI file upload failed ({resp.status}): "
                                + self._owui_error_message(await resp.text())
                            }
                        )
                    upload_result = await resp.json()

            file_id = upload_result.get("id")
            if not file_id:
                return json.dumps(
                    {"error": "OpenWebUI did not return a file ID after upload."}
                )

            # ----------------------------------------------------------
            # Step 4: Extract text from the upload response, or fetch it
            # from the file record if not inlined.
            # ----------------------------------------------------------
            file_data = upload_result.get("data") or {}
            content = (file_data.get("content") or "").strip()
            processing_error = file_data.get("error")

            if not content or file_data.get("status") == "failed":
                async with aiohttp.ClientSession(
                    timeout=aiohttp.ClientTimeout(total=30),
                    connector=self._connector(),
                ) as session:
                    async with session.get(
                        f"{base_url}/api/v1/files/{file_id}",
                        headers=owui_headers,
                    ) as resp:
                        if resp.ok:
                            file_record = await resp.json()
                            file_data = file_record.get("data") or {}
                            content = (file_data.get("content") or "").strip()
                            processing_error = (
                                file_data.get("error") or processing_error
                            )

            # ----------------------------------------------------------
            # Step 5: Last-resort client-side PDF extraction.
            # ----------------------------------------------------------
            if not content and file_name.lower().endswith(".pdf"):
                content = self._extract_pdf_text(file_bytes)

            if not content:
                await self._delete_owui_file(base_url, owui_headers, file_id)
                detail = (
                    f": {processing_error}"
                    if processing_error
                    else ". The file may be image-only, password-protected, or in an unsupported format."
                )
                return json.dumps(
                    {
                        "error": f"No text could be extracted from this document{detail}"
                    }
                )

            max_chars = self.valves.max_page_content_chars
            truncated = len(content) > max_chars
            display_content = (
                content[:max_chars] + "\n\n[... content truncated ...]"
                if truncated
                else content
            )

            # NOTE: we intentionally do NOT delete the OWUI file here — keeping it
            # makes the document available in the chat's attachments and lets
            # OpenWebUI render it as a citation card (via the 'view_file' tool-name
            # hook in middleware.get_citation_source_from_tool_result).
            return json.dumps(
                {
                    "id": file_id,
                    "filename": file_name,
                    "content": display_content,
                    "item_id": item_id,
                    "size_mb": round(size_mb, 2),
                    "content_length": len(content),
                    "truncated": truncated,
                    "file_url": f"{base_url}/api/v1/files/{file_id}/content",
                },
                ensure_ascii=False,
            )

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    # Backwards-compatible alias. The canonical tool name is ``view_file`` — it
    # is the one hooked into OpenWebUI's citation renderer. Keep the old name
    # as a thin wrapper so existing prompts/bookmarks keep working.
    async def get_document_content(
        self,
        item_id: str,
        file_name: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Deprecated alias for ``view_file``. Prefer calling ``view_file`` directly
        so the result is rendered as a citation card in the chat.
        """
        return await self.view_file(
            item_id=item_id,
            file_name=file_name,
            site_id=site_id,
            drive_id=drive_id,
            __oauth_token__=__oauth_token__,
            __user__=__user__,
            __request__=__request__,
        )
