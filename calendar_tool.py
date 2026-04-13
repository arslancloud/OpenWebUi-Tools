"""
title: Outlook Calendar Tool
author: OpenWebUI User
version: 2.0.0
required_open_webui_version: 0.5.0
requirements: aiohttp, msal
description: >
    Read, search, and manage Microsoft Outlook calendar events via the Microsoft
    Graph API using delegated permissions.

    Authentication priority (automatic):
      1. MSAL per-user token cache  (survives SSO token expiry, auto-refreshes)
      2. OpenWebUI SSO token        (__oauth_token__ from Microsoft login)

    When both are missing or expired, type "connect my microsoft account" in chat
    to trigger a device-code sign-in link — authenticating once covers the
    Calendar, Mail, and SharePoint tools simultaneously.

    ADMIN SETUP (one-time):
      1. Azure App Registration > Authentication > Allow public client flows = Yes
      2. API Permissions > Delegated: Calendars.ReadWrite (already granted)
      3. Set Valves: azure_client_id, azure_tenant_id, token_cache_dir
"""

from __future__ import annotations

import asyncio
import json
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

import aiohttp
import msal
from pydantic import BaseModel, Field


GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Combined scopes for all three tools — one sign-in covers everything.
_ALL_TOOL_SCOPES = [
    "Mail.Read", "Mail.Send", "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Sites.Read.All", "Files.Read.All",
]
# Scopes required specifically by this tool (used for silent cache lookups).
_TOOL_SCOPES = ["Calendars.ReadWrite"]


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
        default_timezone: str = Field(
            default="UTC",
            description="IANA timezone for event times, e.g. 'Europe/Berlin'.",
        )
        max_events: int = Field(
            default=50,
            description="Hard upper limit on events returned per request.",
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
        """Return a valid access token from the MSAL cache, or None."""
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
            "Prefer": f'outlook.timezone="{self.valves.default_timezone}"',
        }

    # ------------------------------------------------------------------
    # Other private helpers (unchanged)
    # ------------------------------------------------------------------
    def _parse_time_period(self, time_period: str) -> tuple[str, str]:
        now_utc = datetime.now(timezone.utc)
        today = now_utc.replace(hour=0, minute=0, second=0, microsecond=0)
        period = time_period.strip().lower()

        if period == "today":
            return today.strftime("%Y-%m-%dT%H:%M:%SZ"), (today + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
        if period == "tomorrow":
            s = today + timedelta(days=1)
            return s.strftime("%Y-%m-%dT%H:%M:%SZ"), (s + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
        if period == "this_week":
            monday = today - timedelta(days=today.weekday())
            return monday.strftime("%Y-%m-%dT%H:%M:%SZ"), (monday + timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
        if period == "next_week":
            monday = today + timedelta(days=7 - today.weekday())
            return monday.strftime("%Y-%m-%dT%H:%M:%SZ"), (monday + timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
        if "/" in time_period:
            parts = time_period.split("/", 1)
            try:
                s = datetime.strptime(parts[0].strip(), "%Y-%m-%d").replace(tzinfo=timezone.utc)
                e = datetime.strptime(parts[1].strip(), "%Y-%m-%d").replace(tzinfo=timezone.utc) + timedelta(days=1)
                return s.strftime("%Y-%m-%dT%H:%M:%SZ"), e.strftime("%Y-%m-%dT%H:%M:%SZ")
            except ValueError:
                pass
        for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ"):
            try:
                parsed = datetime.strptime(time_period.strip(), fmt).replace(tzinfo=timezone.utc)
                return parsed.strftime("%Y-%m-%dT%H:%M:%SZ"), (parsed + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
            except ValueError:
                continue
        raise ValueError(
            f"Unrecognised time_period '{time_period}'. "
            "Use 'today', 'tomorrow', 'this_week', 'next_week', 'YYYY-MM-DD', or 'YYYY-MM-DD/YYYY-MM-DD'."
        )

    def _duration_to_iso(self, minutes: int) -> str:
        h, m = divmod(minutes, 60)
        if h and m:
            return f"PT{h}H{m}M"
        return f"PT{h}H" if h else f"PT{m}M"

    def _format_event(self, event: dict) -> dict:
        start = event.get("start", {})
        end = event.get("end", {})
        organizer = event.get("organizer", {}).get("emailAddress", {})
        return {
            "id": event.get("id", ""),
            "subject": event.get("subject", "(no subject)"),
            "start": start.get("dateTime", ""),
            "end": end.get("dateTime", ""),
            "timezone": start.get("timeZone", self.valves.default_timezone),
            "organizer_name": organizer.get("name", ""),
            "organizer_email": organizer.get("address", ""),
            "attendees": [
                {
                    "name": a.get("emailAddress", {}).get("name", ""),
                    "email": a.get("emailAddress", {}).get("address", ""),
                    "response": a.get("status", {}).get("response", "none"),
                    "type": a.get("type", "required"),
                }
                for a in event.get("attendees", [])
            ],
            "location": event.get("location", {}).get("displayName", ""),
            "is_online_meeting": event.get("isOnlineMeeting", False),
            "online_meeting_url": event.get("onlineMeetingUrl", ""),
            "preview": event.get("bodyPreview", ""),
            "is_cancelled": event.get("isCancelled", False),
            "is_all_day": event.get("isAllDay", False),
            "response_status": event.get("responseStatus", {}).get("response", ""),
        }

    def _build_attendee_list(self, emails: str) -> list:
        return [
            {"emailAddress": {"address": e.strip()}, "type": "required"}
            for e in emails.split(";") if e.strip()
        ]

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

    async def get_events(
        self,
        time_period: str,
        max_results: int = 20,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Retrieve calendar events for a given time period.

        Args:
            time_period: 'today', 'tomorrow', 'this_week', 'next_week',
                         'YYYY-MM-DD', or 'YYYY-MM-DD/YYYY-MM-DD'.
            max_results: Maximum number of events to return (default 20).

        Returns:
            JSON list of events with subject, start/end times, attendees, location,
            and Teams link if present.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)
            start_dt, end_dt = self._parse_time_period(time_period)
            top = min(max_results, self.valves.max_events)

            params = {
                "startDateTime": start_dt,
                "endDateTime": end_dt,
                "$select": "id,subject,start,end,organizer,attendees,location,isOnlineMeeting,onlineMeetingUrl,bodyPreview,isCancelled,isAllDay,responseStatus",
                "$orderby": "start/dateTime",
                "$top": str(top),
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(f"{GRAPH_BASE}/me/calendarView", headers=headers, params=params) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Calendars.ReadWrite permission required."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            events = [self._format_event(e) for e in data.get("value", [])]
            return json.dumps({"count": len(events), "time_period": time_period,
                               "start": start_dt, "end": end_dt, "events": events},
                              ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def get_event_details(
        self,
        event_id: str,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Retrieve the full details of a specific calendar event including body/description.

        Args:
            event_id: The unique event ID (from get_events).

        Returns:
            JSON with complete event details including description, all attendee responses,
            and recurrence pattern if applicable.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.get(
                    f"{GRAPH_BASE}/me/events/{event_id}",
                    headers=headers,
                    params={"$select": "id,subject,start,end,organizer,attendees,location,isOnlineMeeting,onlineMeetingUrl,body,isCancelled,isAllDay,responseStatus,recurrence,importance"},
                ) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Calendars.ReadWrite permission required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Event '{event_id}' not found."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    event = await resp.json()

            formatted = self._format_event(event)
            body = event.get("body", {})
            formatted["body"] = body.get("content", "")
            formatted["body_type"] = body.get("contentType", "text")
            formatted["importance"] = event.get("importance", "normal")
            formatted["is_recurring"] = event.get("recurrence") is not None
            return json.dumps(formatted, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def find_available_meeting_times(
        self,
        attendee_emails: str,
        duration_minutes: int = 60,
        search_days: int = 5,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Find meeting times when all given attendees are available.
        Use this for "When can I meet with X and Y?" or "Schedule a meeting when everyone is free."

        Args:
            attendee_emails: Semicolon-separated email addresses (current user added automatically).
            duration_minutes: Required meeting length in minutes (default 60).
            search_days: How many working days ahead to search (default 5).

        Returns:
            JSON with up to 5 suggested meeting slots ordered by earliest availability.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            now_utc = datetime.now(timezone.utc)
            search_start = now_utc.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
            search_end = search_start + timedelta(days=search_days)

            payload = {
                "attendees": self._build_attendee_list(attendee_emails),
                "locationConstraint": {"isRequired": False, "suggestLocation": False},
                "timeConstraint": {
                    "activityDomain": "work",
                    "timeslots": [{
                        "start": {"dateTime": search_start.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
                        "end": {"dateTime": search_end.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
                    }],
                },
                "meetingDuration": self._duration_to_iso(duration_minutes),
                "returnSuggestionReasons": True,
                "minimumAttendeePercentage": 100,
                "maxCandidates": 5,
            }

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(f"{GRAPH_BASE}/me/findMeetingTimes", headers=headers, json=payload) as resp:
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Calendars.ReadWrite permission required."})
                    if not resp.ok:
                        return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})
                    data = await resp.json()

            suggestions = [
                {
                    "start": s.get("meetingTimeSlot", {}).get("start", {}).get("dateTime", ""),
                    "end": s.get("meetingTimeSlot", {}).get("end", {}).get("dateTime", ""),
                    "timezone": s.get("meetingTimeSlot", {}).get("start", {}).get("timeZone", "UTC"),
                    "confidence": s.get("confidence", 0),
                    "attendee_availability": [
                        {"email": a.get("attendee", {}).get("emailAddress", {}).get("address", ""),
                         "availability": a.get("availability", "unknown")}
                        for a in s.get("attendeeAvailability", [])
                    ],
                }
                for s in data.get("meetingTimeSuggestions", [])
            ]

            return json.dumps({
                "duration_minutes": duration_minutes,
                "search_days": search_days,
                "attendees": attendee_emails,
                "suggestions_count": len(suggestions),
                "suggestions": suggestions,
                "empty_reason": data.get("emptySuggestionsReason", "") if not suggestions else "",
            }, ensure_ascii=False)

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def create_event(
        self,
        subject: str,
        start_datetime: str,
        end_datetime: str,
        attendee_emails: Optional[str] = None,
        body: Optional[str] = None,
        location: Optional[str] = None,
        is_online_meeting: bool = True,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Create a new calendar event and send invitations to attendees.

        Args:
            subject: Meeting title.
            start_datetime: Start time ISO string, e.g. '2024-03-15T10:00:00'.
            end_datetime: End time ISO string, e.g. '2024-03-15T11:00:00'.
            attendee_emails: Optional semicolon-separated emails to invite.
            body: Optional meeting description or agenda.
            location: Optional location name or room.
            is_online_meeting: Creates a Teams link if True (default True).

        Returns:
            JSON with created event ID, subject, start/end, and Teams link.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            event_payload: dict = {
                "subject": subject,
                "start": {"dateTime": start_datetime, "timeZone": self.valves.default_timezone},
                "end": {"dateTime": end_datetime, "timeZone": self.valves.default_timezone},
                "isOnlineMeeting": is_online_meeting,
            }
            if is_online_meeting:
                event_payload["onlineMeetingProvider"] = "teamsForBusiness"
            if attendee_emails:
                event_payload["attendees"] = self._build_attendee_list(attendee_emails)
            if body:
                event_payload["body"] = {"contentType": "Text", "content": body}
            if location:
                event_payload["location"] = {"displayName": location}

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(f"{GRAPH_BASE}/me/events", headers=headers, json=event_payload) as resp:
                    if resp.status == 201:
                        created = await resp.json()
                        return json.dumps({
                            "success": True,
                            "id": created.get("id", ""),
                            "subject": created.get("subject", ""),
                            "start": created.get("start", {}).get("dateTime", ""),
                            "end": created.get("end", {}).get("dateTime", ""),
                            "online_meeting_url": created.get("onlineMeetingUrl", ""),
                            "message": f"Event '{subject}' created."
                                       + (" Invitations sent." if attendee_emails else ""),
                        }, ensure_ascii=False)
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Calendars.ReadWrite permission required."})
                    return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def update_event(
        self,
        event_id: str,
        subject: Optional[str] = None,
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
        location: Optional[str] = None,
        body: Optional[str] = None,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Update an existing calendar event. Only provided fields are changed.

        Args:
            event_id: The event ID to update (from get_events).
            subject: New title (optional).
            start_datetime: New start ISO string (optional).
            end_datetime: New end ISO string (optional).
            location: New location (optional).
            body: New description (optional).

        Returns:
            JSON indicating success or error.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            patch: dict = {}
            if subject:
                patch["subject"] = subject
            if start_datetime:
                patch["start"] = {"dateTime": start_datetime, "timeZone": self.valves.default_timezone}
            if end_datetime:
                patch["end"] = {"dateTime": end_datetime, "timeZone": self.valves.default_timezone}
            if location:
                patch["location"] = {"displayName": location}
            if body:
                patch["body"] = {"contentType": "Text", "content": body}
            if not patch:
                return json.dumps({"error": "No fields provided to update."})

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.patch(f"{GRAPH_BASE}/me/events/{event_id}", headers=headers, json=patch) as resp:
                    if resp.status == 200:
                        return json.dumps({"success": True, "message": "Event updated successfully."})
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Calendars.ReadWrite permission required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Event '{event_id}' not found."})
                    return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})

    async def delete_event(
        self,
        event_id: str,
        __oauth_token__: Optional[dict] = None,
        __user__: Optional[dict] = None,
        __request__: Optional[object] = None,
    ) -> str:
        """
        Delete (cancel) a calendar event. Attendees receive a cancellation notice.

        Args:
            event_id: The event ID to delete (from get_events).

        Returns:
            JSON indicating success or error.
        """
        try:
            user_id = (__user__ or {}).get("id", "anonymous")
            headers = self._get_headers(user_id, __oauth_token__)

            timeout = aiohttp.ClientTimeout(total=self.valves.request_timeout_seconds)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.delete(f"{GRAPH_BASE}/me/events/{event_id}", headers=headers) as resp:
                    if resp.status == 204:
                        return json.dumps({"success": True, "message": "Event deleted. Cancellation sent to attendees."})
                    if resp.status == 401:
                        return json.dumps({"error": "Session expired. Type 'connect my microsoft account' to re-authenticate."})
                    if resp.status == 403:
                        return json.dumps({"error": "Permission denied. Calendars.ReadWrite permission required."})
                    if resp.status == 404:
                        return json.dumps({"error": f"Event '{event_id}' not found."})
                    return json.dumps({"error": f"Graph API error {resp.status}: {await resp.text()}"})

        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        except aiohttp.ClientError as exc:
            return json.dumps({"error": f"Network error: {exc}"})
