"""
title: Outlook Mail Tool
author: OpenWebUI User
version: 2.0.0
required_open_webui_version: 0.5.0
requirements: aiohttp, msal
description: >
    Read, search, and send Microsoft Outlook email via the Microsoft Graph API
    using delegated permissions. Includes built-in authentication via Microsoft
    Device Code Flow — no separate OAuth configuration required.

    SETUP (admin, one-time):
      1. In Azure Portal > App Registrations > [your app] > Authentication:
         Enable "Allow public client flows" = Yes
      2. In API Permissions add Delegated: Mail.Read, Mail.Send, Mail.ReadWrite
         then click "Grant admin consent"
      3. In the tool Valves, set:
         - azure_client_id  (Application/Client ID from App Overview)
         - azure_tenant_id  (Directory/Tenant ID from App Overview)

    USER AUTHENTICATION:
      Type "connect my outlook" or "authenticate outlook" in chat.
      The tool will reply with a short URL and 9-character code.
      Open https://microsoft.com/devicelogin, enter the code, and sign in.
      Done — your token is saved and used automatically from then on.
"""

from __future__ import annotations

import asyncio
import json
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

import aiohttp
import msal
from pydantic import BaseModel, Field


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
MAIL_SCOPES = ["Mail.Read", "Mail.Send", "Mail.ReadWrite"]


class Tools:
    # ------------------------------------------------------------------
    # Valves — admin-configurable settings
    # ------------------------------------------------------------------
    class Valves(BaseModel):
        azure_client_id: str = Field(
            default="",
            description="Azure App Registration Client ID (Application ID). Required for authentication.",
        )
        azure_tenant_id: str = Field(
            default="common",
            description="Azure Tenant ID. Use 'common' for multi-tenant or paste your Directory ID.",
        )
        token_cache_dir: str = Field(
            default="/app/backend/data/outlook_tokens",
            description="Server directory where per-user token caches are stored. Must be writable by the OpenWebUI process.",
        )
        max_emails: int = Field(
            default=50,
            description="Hard upper limit on emails returned per request.",
        )
        request_timeout_seconds: int = Field(
            default=30,
            description="HTTP timeout in seconds for Microsoft Graph API calls.",
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
        cache_dir = Path(self.valves.token_cache_dir)
        cache_dir.mkdir(parents=True, exist_ok=True)
        return cache_dir / f"{user_id}.json"

    def _load_cache(self, user_id: str) -> msal.SerializableTokenCache:
        cache = msal.SerializableTokenCache()
        p = self._cache_path(user_id)
        if p.exists():
            cache.deserialize(p.read_text(encoding="utf-8"))
        return cache

    def _save_cache(self, user_id: str, cache: msal.SerializableTokenCache) -> None:
        if cache.has_state_changed:
            self._cache_path(user_id).write_text(
                cache.serialize(), encoding="utf-8"
            )

    def _msal_app(self, cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
        if not self.valves.azure_client_id:
            raise ValueError(
                "Azure Client ID is not configured. "
                "An administrator must set 'azure_client_id' in the tool Valves."
            )
        return msal.PublicClientApplication(
            client_id=self.valves.azure_client_id,
            authority=f"https://login.microsoftonline.com/{self.valves.azure_tenant_id}",
            token_cache=cache,
        )

    def _access_token_from_cache(self, user_id: str) -> Optional[str]:
        """Return a valid access token from the MSAL cache, or None if absent/expired."""
        try:
            cache = self._load_cache(user_id)
            app = self._msal_app(cache)
            accounts = app.get_accounts()
            if not accounts:
                return None
            result = app.acquire_token_silent(MAIL_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache(user_id, cache)
                return result["access_token"]
        except Exception:
            pass
        return None

    # ------------------------------------------------------------------
    # Header builder — tries MSAL cache first, then OpenWebUI OAuth token
    # ------------------------------------------------------------------
    def _get_headers(
        self,
        user_id: Optional[str],
        oauth_token: Optional[dict],
    ) -> dict:
        access_token: Optional[str] = None

        # 1. Try the MSAL per-user cache (device-code auth or prior session)
        if user_id:
            access_token = self._access_token_from_cache(user_id)

        # 2. Fall back to the token OpenWebUI injected from its own OAuth session
        if not access_token and oauth_token:
            access_token = oauth_token.get("access_token")

        if not access_token:
            raise ValueError(
                "Not authenticated with Microsoft Outlook. "
                "Type 'connect my outlook' to sign in."
            )

        return {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    # ------------------------------------------------------------------
    # Background polling task for device code flow
    # ------------------------------------------------------------------
    async def _poll_device_flow(
        self,
        user_id: str,
        flow: dict,
        cache: msal.SerializableTokenCache,
    ) -> None:
        """Runs in the background. Polls Microsoft until the user authenticates."""
        loop = asyncio.get_event_loop()
        app = self._msal_app(cache)
        try:
            result = await loop.run_in_executor(
                None,
                lambda: app.acquire_token_by_device_flow(flow),
            )
            if result and "access_token" in result:
                self._save_cache(user_id, cache)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Email formatting helper
    # ------------------------------------------------------------------
    def _parse_time_period(self, time_period: str) -> str:
        now_utc = datetime.now(timezone.utc)
        today_midnight = now_utc.replace(hour=0, minute=0, second=0, microsecond=0)
        period = time_period.strip().lower()

        if period == "today":
            return f"receivedDateTime gt {today_midnight.strftime('%Y-%m-%dT%H:%M:%SZ')}"
        if period == "yesterday":
            s = today_midnight - timedelta(days=1)
            e = today_midnight
            return (
                f"receivedDateTime gt {s.strftime('%Y-%m-%dT%H:%M:%SZ')}"
                f" and receivedDateTime lt {e.strftime('%Y-%m-%dT%H:%M:%SZ')}"
            )
        if period == "last_7_days":
            s = today_midnight - timedelta(days=7)
            return f"receivedDateTime gt {s.strftime('%Y-%m-%dT%H:%M:%SZ')}"
        if period == "last_30_days":
            s = today_midnight - timedelta(days=30)
            return f"receivedDateTime gt {s.strftime('%Y-%m-%dT%H:%M:%SZ')}"

        for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S"):
            try:
                parsed = datetime.strptime(time_period.strip(), fmt).replace(tzinfo=timezone.utc)
                return f"receivedDateTime gt {parsed.strftime('%Y-%m-%dT%H:%M:%SZ')}"
            except ValueError:
                continue

        raise ValueError(
            f"Unrecognised time period '{time_period}'. "
            "Use 'today', 'yesterday', 'last_7_days', 'last_30_days', or 'YYYY-MM-DD'."
        )

    def _format_email(self, msg: dict, include_body: bool = False) -> dict:
        sender = msg.get("from", {}).get("emailAddress", {})
        result = {
            "id": msg.get("id", ""),
            "conversation_id": msg.get("conversationId", ""),
            "subject": msg.get("subject", "(no subject)"),
            "from_name": sender.get("name", ""),
            "from_email": sender.get("address", ""),
            "to": [r.get("emailAddress", {}).get("address", "") for r in msg.get("toRecipients", [])],
            "received": msg.get("receivedDateTime", ""),
            "is_read": msg.get("isRead", True),
            "preview": msg.get("bodyPreview", ""),
        }
        if include_body:
            body = msg.get("body", {})
            result["body_type"] = body.get("contentType", "text")
            result["body"] = body.get("content", "")
        return result

    # ------------------------------------------------------------------
    # PUBLIC TOOL FUNCTIONS
    # ------------------------------------------------------------------

    async def authenticate_with_microsoft(
        self,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Connect your Microsoft Outlook account to this tool using a one-time sign-in code.
        Call this function whenever you see an authentication error, or when you want to
        connect or reconnect your Outlook account.

        Returns:
            A sign-in URL and a short code to enter at that URL. Open the URL in any browser,
            enter the code, and sign in with your Microsoft account. The tool will automatically
            detect when you have signed in.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")

            # Already authenticated?
            existing = self._access_token_from_cache(user_id)
            if existing:
                return json.dumps({
                    "status": "already_authenticated",
                    "message": (
                        "Your Outlook account is already connected. "
                        "You can now ask about your emails."
                    ),
                })

            cache = self._load_cache(user_id)
            loop = asyncio.get_event_loop()
            app = self._msal_app(cache)

            flow = await loop.run_in_executor(
                None,
                lambda: app.initiate_device_flow(scopes=MAIL_SCOPES),
            )

            if "user_code" not in flow:
                err = flow.get("error_description", flow.get("error", "Unknown error"))
                return json.dumps({"error": f"Could not start sign-in flow: {err}"})

            expires_minutes = flow.get("expires_in", 900) // 60

            # Start background polling — detects when user completes sign-in
            asyncio.create_task(self._poll_device_flow(user_id, flow, cache))

            return json.dumps({
                "status": "authentication_required",
                "message": (
                    f"**To connect your Outlook account:**\n\n"
                    f"1. Open this link in your browser: {flow['verification_uri']}\n"
                    f"2. Enter this code: **{flow['user_code']}**\n\n"
                    f"The code expires in {expires_minutes} minutes. "
                    f"After you sign in, simply repeat your original request."
                ),
                "sign_in_url": flow["verification_uri"],
                "user_code": flow["user_code"],
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except Exception as exc:
            return json.dumps({"error": f"Authentication setup failed: {exc}"})

    async def get_emails(
        self,
        time_period: str,
        filter_unread: bool = False,
        max_results: int = 20,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Retrieve emails from a specific time period, optionally filtered to unread only.

        Args:
            time_period: When to retrieve emails from. Accepts:
                         'today', 'yesterday', 'last_7_days', 'last_30_days',
                         or an ISO date string like '2024-03-15'.
            filter_unread: If True, only return unread emails. Default False.
            max_results: Maximum number of emails to return (default 20).

        Returns:
            JSON string with a list of emails including sender, subject, preview, date, and read status.
            If not authenticated, returns instructions to connect the Outlook account.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            top = min(max_results, self.valves.max_emails)

            time_filter = self._parse_time_period(time_period)
            filters = [time_filter]
            if filter_unread:
                filters.append("isRead eq false")

            params = {
                "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,conversationId",
                "$orderby": "receivedDateTime desc",
                "$top": str(top),
                "$filter": " and ".join(filters),
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(f"{GRAPH_BASE}/me/messages", headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my outlook' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Mail.Read permission is required. Contact your administrator."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            emails = [self._format_email(m) for m in data.get("value", [])]
            return json.dumps({
                "count": len(emails),
                "time_period": time_period,
                "filter_unread": filter_unread,
                "emails": emails,
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def search_emails(
        self,
        query: str,
        sender_email: Optional[str] = None,
        max_results: int = 20,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Search emails by keyword, subject, or sender.

        Args:
            query: Free-text search query (searches subject and body).
                   Supports KQL: 'subject:budget', 'from:john@company.com'.
                   May be empty string when filtering by sender_email only.
            sender_email: Optional — filter results to emails from this address.
            max_results: Maximum number of results (default 20).

        Returns:
            JSON string with matching emails.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            top = min(max_results, self.valves.max_emails)

            params = {
                "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,conversationId",
                "$top": str(top),
            }

            # Graph cannot combine $search + $filter for messages
            if sender_email and query:
                params["$search"] = f'"{query}" from:{sender_email}'
                headers["ConsistencyLevel"] = "eventual"
            elif sender_email:
                params["$filter"] = f"from/emailAddress/address eq '{sender_email}'"
                params["$orderby"] = "receivedDateTime desc"
            else:
                params["$search"] = f'"{query}"'
                headers["ConsistencyLevel"] = "eventual"

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(f"{GRAPH_BASE}/me/messages", headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my outlook' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Mail.Read permission required."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            emails = [self._format_email(m) for m in data.get("value", [])]
            return json.dumps({
                "count": len(emails),
                "query": query,
                "sender_filter": sender_email,
                "emails": emails,
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def get_email_details(
        self,
        email_id: str,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Retrieve the full content of a specific email including the complete body.
        Also marks the email as read.

        Args:
            email_id: The unique message ID (from get_emails or search_emails).

        Returns:
            JSON string with the complete email including full body, CC list,
            importance level, and whether it has attachments.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            params = {
                "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,body,conversationId,importance,hasAttachments",
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(f"{GRAPH_BASE}/me/messages/{email_id}", headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my outlook' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Mail.Read permission required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Email '{email_id}' not found."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    msg = await resp.json()

            # Mark as read — silent fail if Mail.ReadWrite not granted
            try:
                async with aiohttp.ClientSession(timeout=timeout) as session:
                    await session.patch(
                        f"{GRAPH_BASE}/me/messages/{email_id}",
                        headers=headers,
                        json={"isRead": True},
                    )
            except Exception:
                pass

            formatted = self._format_email(msg, include_body=True)
            formatted["cc"] = [r.get("emailAddress", {}).get("address", "") for r in msg.get("ccRecipients", [])]
            formatted["importance"] = msg.get("importance", "normal")
            formatted["has_attachments"] = msg.get("hasAttachments", False)

            return json.dumps(formatted, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def get_email_thread(
        self,
        conversation_id: str,
        max_results: int = 50,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Retrieve all messages in an email conversation thread, ordered oldest to newest.

        Args:
            conversation_id: The conversation ID (from any email's conversation_id field).
            max_results: Maximum number of messages to return (default 50).

        Returns:
            JSON string with all emails in the thread in chronological order, including full bodies.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            top = min(max_results, self.valves.max_emails)

            params = {
                "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,body,conversationId",
                "$filter": f"conversationId eq '{conversation_id}'",
                "$orderby": "receivedDateTime asc",
                "$top": str(top),
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(f"{GRAPH_BASE}/me/messages", headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my outlook' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Mail.Read permission required."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            thread = [self._format_email(m, include_body=True) for m in data.get("value", [])]
            return json.dumps({
                "conversation_id": conversation_id,
                "message_count": len(thread),
                "thread": thread,
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def send_email(
        self,
        to_address: str,
        subject: str,
        body: str,
        cc_address: Optional[str] = None,
        is_html: bool = False,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Send a new email to one or more recipients.

        Args:
            to_address: Recipient email address. Separate multiple with semicolons.
            subject: Email subject line.
            body: Email body text (plain text, or HTML if is_html is True).
            cc_address: Optional CC addresses, semicolon-separated.
            is_html: True if body contains HTML markup. Default False.

        Returns:
            JSON string indicating success or describing any error.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            def build_recipients(addr_string: str) -> list:
                return [
                    {"emailAddress": {"address": a.strip()}}
                    for a in addr_string.split(";")
                    if a.strip()
                ]

            message_payload: dict = {
                "subject": subject,
                "body": {"contentType": "HTML" if is_html else "Text", "content": body},
                "toRecipients": build_recipients(to_address),
            }
            if cc_address:
                message_payload["ccRecipients"] = build_recipients(cc_address)

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(
                    f"{GRAPH_BASE}/me/sendMail",
                    headers=headers,
                    json={"message": message_payload, "saveToSentItems": True},
                ) as resp:
                    if resp.status == 202:
                        return json.dumps({"success": True, "message": f"Email sent to {to_address}."})
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my outlook' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Mail.Send permission required."})
                    return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def reply_to_email(
        self,
        email_id: str,
        reply_body: str,
        reply_all: bool = False,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Reply to an existing email. Optionally reply to all recipients.

        Args:
            email_id: The message ID of the email to reply to (from get_emails or search_emails).
            reply_body: The text of the reply.
            reply_all: If True, reply to all original recipients. Default False.

        Returns:
            JSON string indicating success or describing any error.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            endpoint = "replyAll" if reply_all else "reply"
            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)

            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(
                    f"{GRAPH_BASE}/me/messages/{email_id}/{endpoint}",
                    headers=headers,
                    json={"comment": reply_body},
                ) as resp:
                    if resp.status == 202:
                        action = "Reply-all" if reply_all else "Reply"
                        return json.dumps({"success": True, "message": f"{action} sent successfully."})
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my outlook' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Mail.Send permission required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Email '{email_id}' not found."})
                    return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})
