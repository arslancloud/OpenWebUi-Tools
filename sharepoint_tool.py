"""
title: SharePoint Tool
author: OpenWebUI User
version: 2.0.0
required_open_webui_version: 0.5.0
requirements: aiohttp, msal
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
import html
import json
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional
from datetime import timedelta, datetime, timezone

import aiohttp
import msal
from pydantic import BaseModel, Field


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SEARCH_ENDPOINT = f"{GRAPH_BASE}/search/query"

# Combined scopes for all three tools — one sign-in covers everything.
_ALL_TOOL_SCOPES = [
    "Mail.Read", "Mail.Send", "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Sites.Read.All", "Files.Read.All",
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

    # ------------------------------------------------------------------
    # Init
    # ------------------------------------------------------------------
    def __init__(self) -> None:
        self.valves = self.Valves()

    # ------------------------------------------------------------------
    # MSAL token cache helpers
    # ------------------------------------------------------------------
    def _cache_path(self, user_id: str) -> Path:
        d = Path(self.valves.token_cache_dir)
        d.mkdir(parents=True, exist_ok=True)
        return d / f"{user_id}.json"

    def _load_cache(self, user_id: str) -> msal.SerializableTokenCache:
        cache = msal.SerializableTokenCache()
        p = self._cache_path(user_id)
        if p.exists():
            cache.deserialize(p.read_text(encoding="utf-8"))
        return cache

    def _save_cache(self, user_id: str, cache: msal.SerializableTokenCache) -> None:
        if cache.has_state_changed:
            self._cache_path(user_id).write_text(cache.serialize(), encoding="utf-8")

    def _msal_app(self, cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
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
            result.update({
                "id": resource.get("id", ""),
                "name": resource.get("displayName", ""),
                "url": resource.get("webUrl", ""),
                "description": resource.get("description", ""),
            })
        elif "driveItem" in odata_type:
            result.update({
                "id": resource.get("id", ""),
                "name": resource.get("name", ""),
                "url": resource.get("webUrl", ""),
                "size": resource.get("size", 0),
                "last_modified": resource.get("lastModifiedDateTime", ""),
                "site_id": resource.get("parentReference", {}).get("siteId", ""),
                "drive_id": resource.get("parentReference", {}).get("driveId", ""),
            })
        elif "listItem" in odata_type:
            fields = resource.get("fields", {})
            result.update({
                "id": resource.get("id", ""),
                "name": fields.get("Title", fields.get("FileLeafRef", "")),
                "url": resource.get("webUrl", ""),
                "last_modified": resource.get("lastModifiedDateTime", ""),
                "site_id": resource.get("parentReference", {}).get("siteId", ""),
            })
        return result

    def _format_drive_item(self, item: dict) -> dict:
        is_folder = "folder" in item
        result = {
            "id": item.get("id", ""),
            "name": item.get("name", ""),
            "url": item.get("webUrl", ""),
            "is_folder": is_folder,
            "size_bytes": item.get("size", 0),
            "created": item.get("createdDateTime", ""),
            "last_modified": item.get("lastModifiedDateTime", ""),
            "created_by": item.get("createdBy", {}).get("user", {}).get("displayName", ""),
            "modified_by": item.get("lastModifiedBy", {}).get("user", {}).get("displayName", ""),
        }
        if not is_folder:
            result["mime_type"] = item.get("file", {}).get("mimeType", "")
        if is_folder:
            result["child_count"] = item.get("folder", {}).get("childCount", 0)
        return result

    async def _resolve_site_url(self, session: aiohttp.ClientSession, headers: dict, site_id: str) -> str:
        async with session.get(
            f"{GRAPH_BASE}/sites/{site_id}",
            headers=headers,
            params={"$select": "webUrl"},
        ) as resp:
            if resp.ok:
                return (await resp.json()).get("webUrl", "")
        return ""

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
                return json.dumps({
                    "status": "already_authenticated",
                    "message": "Your Microsoft account is already connected. You can use all Outlook tools.",
                })

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
            return json.dumps({
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
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except Exception as exc:
            return json.dumps({"error": f"Authentication setup failed: {exc}"})

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
                "requests": [{
                    "entityTypes": entity_types,
                    "query": {"queryString": query},
                    "fields": ["id", "name", "displayName", "webUrl", "description",
                               "lastModifiedDateTime", "size", "summary", "parentReference", "fields"],
                    "size": top,
                }]
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(SEARCH_ENDPOINT, headers=headers, json=payload) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Sites.Read.All and Files.Read.All are required."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            hits = [
                self._format_search_hit(hit)
                for rv in data.get("value", [])
                for container in rv.get("hitsContainers", [])
                for hit in container.get("hits", [])
            ]

            return json.dumps({"query": query, "content_type": content_type,
                               "count": len(hits), "results": hits}, ensure_ascii=False)

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
            async with aiohttp.ClientSession(timeout=timeout) as session:
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
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Sites.Read.All required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Site '{site_id}' not found."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            pages = [
                {
                    "id": p.get("id", ""),
                    "title": p.get("title", "(untitled)"),
                    "url": p.get("webUrl", ""),
                    "last_modified": p.get("lastModifiedDateTime", ""),
                    "created": p.get("createdDateTime", ""),
                    "published": p.get("publishingState", {}).get("level", "") == "published",
                }
                for p in data.get("value", [])
            ]

            return json.dumps({"site_id": site_id, "count": len(pages), "pages": pages},
                              ensure_ascii=False)

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
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(
                    f"{GRAPH_BASE}/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage",
                    headers=headers,
                    params={"$expand": "canvasLayout"},
                ) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Sites.Read.All required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Page '{page_id}' not found on site '{site_id}'."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    page = await resp.json()

            text_parts = [
                _strip_html(wp.get("innerHtml", ""))
                for section in page.get("canvasLayout", {}).get("horizontalSections", [])
                for col in section.get("columns", [])
                for wp in col.get("webparts", [])
                if wp.get("innerHtml", "").strip()
            ]

            full_text = "\n\n".join(filter(None, text_parts))
            if len(full_text) > self.valves.max_page_content_chars:
                full_text = full_text[:self.valves.max_page_content_chars] + "\n\n[... content truncated ...]"

            return json.dumps({
                "site_id": site_id,
                "page_id": page_id,
                "title": page.get("title", ""),
                "url": page.get("webUrl", ""),
                "last_modified": page.get("lastModifiedDateTime", ""),
                "content": full_text,
                "content_length": len(full_text),
            }, ensure_ascii=False)

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
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(url, headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Files.Read.All required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Site or folder not found. site_id='{site_id}', folder='{folder_path}'."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            items = [self._format_drive_item(i) for i in data.get("value", [])]
            return json.dumps({
                "site_id": site_id, "folder": folder_path or "(root)",
                "count": len(items), "newest_first": newest_first, "items": items,
            }, ensure_ascii=False)

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
            async with aiohttp.ClientSession(timeout=timeout) as session:
                site_url = await self._resolve_site_url(session, headers, site_id)
                if not site_url:
                    return json.dumps({"error": f"Could not resolve URL for site '{site_id}'."})

                payload = {
                    "requests": [{
                        "entityTypes": entity_types,
                        "query": {"queryString": query},
                        "contentSources": [site_url],
                        "fields": ["id", "name", "webUrl", "lastModifiedDateTime",
                                   "size", "parentReference", "fields"],
                        "size": top,
                    }]
                }

                async with session.post(SEARCH_ENDPOINT, headers=headers, json=payload) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Sites.Read.All required."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            hits = [
                self._format_search_hit(hit)
                for rv in data.get("value", [])
                for container in rv.get("hitsContainers", [])
                for hit in container.get("hits", [])
            ]

            return json.dumps({
                "site_id": site_id, "site_url": site_url,
                "query": query, "content_type": content_type,
                "count": len(hits), "results": hits,
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})
