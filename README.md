# Microsoft 365 Tools for OpenWebUI

A set of three OpenWebUI tools that connect your LLM to **Microsoft Outlook Mail**, **Outlook Calendar**, and **SharePoint** using the Microsoft Graph API with delegated user permissions. All data access is scoped to the signed-in user — the tools never access other users' accounts.

---

## Tools

| Tool | File | What it does |
|---|---|---|
| Outlook Mail | `outlook_tool.py` | Read, search, and send emails |
| Outlook Calendar | `calendar_tool.py` | Read, create, and manage calendar events |
| SharePoint | `sharepoint_tool.py` | Search and read SharePoint sites, pages, and documents |

---

## Example prompts

```
"Summarize the emails I received today."
"Do I have any unread emails from Sarah?"
"Reply to John's email and say I'll be there at 3pm."
"What meetings do I have tomorrow?"
"Schedule a 1-hour meeting with x@company.com when everyone is free this week."
"Where on SharePoint can I find the onboarding documentation?"
"Which documents on the HR site were uploaded this week?"
"Is there anything about the Q3 budget on the Finance SharePoint page?"
```

---

## Prerequisites

- A running [OpenWebUI](https://github.com/open-webui/open-webui) instance (≥ 0.5.0)
- A Microsoft Azure App Registration with the permissions listed below
- Users must sign into OpenWebUI via **Microsoft SSO** (or use the built-in device-code login)

---

## Azure App Registration setup

### 1. Create or locate your App Registration

Go to **Azure Portal → Azure Active Directory → App registrations** and open (or create) the app used by OpenWebUI.

### 2. Add Delegated API permissions

Under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add the following and then click **Grant admin consent**:

| Permission | Required by |
|---|---|
| `openid` | Sign-in |
| `email` | Sign-in |
| `profile` | Sign-in |
| `Mail.Read` | Outlook Mail Tool |
| `Mail.Send` | Outlook Mail Tool |
| `Mail.ReadWrite` | Outlook Mail Tool (mark as read) |
| `Calendars.ReadWrite` | Calendar Tool |
| `Sites.Read.All` | SharePoint Tool |
| `Files.Read.All` | SharePoint Tool |

> **Note:** These are all **delegated** permissions. The app can only access data on behalf of a user who is actively signed in — it cannot access any mailbox, calendar, or site without user authentication.

### 3. Enable public client flows

Under **Authentication → Advanced settings**, set **Allow public client flows → Yes**.

This enables the device-code sign-in fallback used when the SSO token expires.

### 4. Add the OpenWebUI redirect URI

Under **Authentication → Redirect URIs**, add:

```
https://<your-openwebui-domain>/oauth/oidc/callback
```

---

## OpenWebUI configuration

Add the following to your OpenWebUI `.env` file (or container environment), then **restart OpenWebUI**:

```env
MICROSOFT_CLIENT_ID=<your-azure-app-client-id>
MICROSOFT_CLIENT_SECRET=<your-azure-app-client-secret>
MICROSOFT_CLIENT_TENANT_ID=<your-azure-tenant-id>
MICROSOFT_OAUTH_SCOPE=openid email profile Mail.Read Mail.Send Mail.ReadWrite Calendars.ReadWrite Sites.Read.All Files.Read.All

# Optional — encrypt token cache files at rest (recommended)
OPENWEBUI_TOKEN_CACHE_KEY=<your-strong-passphrase>
```

> **Admin Panel alternative:** If you have already started OpenWebUI once, the scope value may be stored in the database. Update it via **Admin Panel → Settings → OAuth → Microsoft OAuth Scope** instead of (or in addition to) the `.env` file.

### Token cache encryption (optional but recommended)

When `OPENWEBUI_TOKEN_CACHE_KEY` is set, each user's token cache file is encrypted at rest using **AES-256 via Fernet** (key derived with PBKDF2-HMAC-SHA256). Without this variable the files are stored as plain JSON (same behaviour as before).

- Use any strong passphrase — it never leaves the server.
- All three tools share the same key and the same `token_cache_dir`, so existing plain-text cache files are **automatically re-encrypted** on the next successful token refresh — no manual migration needed.
- If you remove the key or change it, users will need to sign in again via `authenticate_with_microsoft`.

---

## Installation

1. Copy the contents of each `.py` file.
2. In OpenWebUI, go to **Workspace → Tools → Create Tool**.
3. Paste the file content and click **Save**. OpenWebUI will automatically install the required Python packages (`aiohttp`, `msal`).
4. Repeat for each of the three tools.

### Configure the Valves

After saving each tool, click the **gear icon** and set these values (same for all three tools):

| Valve | Description | Example |
|---|---|---|
| `azure_client_id` | Azure App Registration Client ID | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| `azure_tenant_id` | Azure Directory (Tenant) ID | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| `token_cache_dir` | Server path for token cache files | `/app/backend/data/outlook_tokens` |

> **Important:** Use the **same `token_cache_dir`** in all three tools. This allows a single sign-in to authenticate all tools at once.

The Calendar tool has one additional Valve:

| Valve | Description | Default |
|---|---|---|
| `default_timezone` | IANA timezone for event times | `UTC` — change to e.g. `Europe/Berlin` |

The SharePoint tool has additional Valves for document content extraction:

| Valve | Description | Default |
|---|---|---|
| `openwebui_base_url` | Base URL of this OpenWebUI instance | Auto-detected from request |
| `max_document_size_mb` | Maximum file size for document extraction | `20` |

> Document extraction now always uses the logged-in user's own OpenWebUI session, so uploaded temporary files are created, read, and deleted under that user's identity.

---

## Authentication

### Primary: Microsoft SSO (recommended)

When users sign into OpenWebUI using **Sign in with Microsoft**, OpenWebUI stores their access token automatically. The tools use this token on every call — no additional steps needed.

Users must sign out and sign back in after the `MICROSOFT_OAUTH_SCOPE` is updated, so a new token with the correct permissions is issued.

### Fallback: Device-code login

If the SSO token is absent or has expired, the tools fall back to a device-code flow. The LLM will prompt the user automatically, or you can trigger it manually:

> *"Connect my Microsoft account"*

The tool responds with a URL and a short code:

```
1. Open: https://microsoft.com/devicelogin
2. Enter code: XXXXXXXXX
```

Open the link in any browser, enter the code, and sign in. **One authentication covers all three tools** — the token is cached per user in `token_cache_dir` and refreshed automatically for up to 90 days.

---

## Tool reference

### Outlook Mail (`outlook_tool.py`)

| Function | Description |
|---|---|
| `authenticate_with_microsoft` | Trigger device-code sign-in |
| `disconnect_microsoft_account` | Delete the saved token cache and sign out |
| `get_emails` | List emails by time period with optional unread filter |
| `search_emails` | Full-text and sender search |
| `get_email_details` | Fetch full body of a single email |
| `get_email_thread` | Retrieve a complete conversation thread |
| `send_email` | Compose and send a new email |
| `reply_to_email` | Reply or reply-all to an existing email |

### Outlook Calendar (`calendar_tool.py`)

| Function | Description |
|---|---|
| `authenticate_with_microsoft` | Trigger device-code sign-in |
| `disconnect_microsoft_account` | Delete the saved token cache and sign out |
| `get_events` | List events for today, this week, a date, or a date range |
| `get_event_details` | Full details of a single event including body and attendee responses |
| `find_available_meeting_times` | Find free slots across multiple attendees |
| `create_event` | Create a new meeting with Teams link and invitations |
| `update_event` | Change title, time, location, or description of an event |
| `delete_event` | Cancel an event and notify attendees |
| `list_categories` | List all Outlook color categories |
| `set_event_categories` | Assign categories to an event |
| `create_category` | Create a new color category |

### SharePoint (`sharepoint_tool.py`)

| Function | Description |
|---|---|
| `authenticate_with_microsoft` | Trigger device-code sign-in |
| `disconnect_microsoft_account` | Delete the saved token cache and sign out |
| `search_sharepoint` | Full-text search across all sites, pages, and documents |
| `get_site_pages` | List all pages on a SharePoint site |
| `get_page_content` | Extract plain-text content from a SharePoint page |
| `list_documents` | List files in a document library, sorted by newest |
| `search_in_site` | Scope a keyword search to a specific site |
| `get_document_content` | Download a file and extract its text via the OpenWebUI pipeline (PDF, DOCX, XLSX, PPTX, …) |

#### Document content extraction workflow

```
1. search_sharepoint / list_documents  →  find the file, note id, name, drive_id
2. get_document_content(item_id=..., file_name=..., drive_id=...)  →  extracted text
```

Example prompts:
- *"Summarise the Q3 budget spreadsheet on SharePoint"*
- *"What does the onboarding PDF say about equipment requests?"*
- *"Find the project charter and list its key milestones"*

---

## Security notes

- All permissions are **delegated** — the app only accesses data on behalf of the currently signed-in user.
- Token cache files are stored server-side, keyed by OpenWebUI user ID. Each user's tokens are isolated in a separate file.
- No application-level (app-only) permissions are used. The app cannot access any data without a user actively authenticating.
- Removing a user's cache file at `<token_cache_dir>/<user_id>.json` immediately revokes their stored session.
- When `OPENWEBUI_TOKEN_CACHE_KEY` is set, cache files are encrypted with AES-256-GCM (Fernet). An attacker with filesystem access cannot read tokens without the key.
- Users can sign themselves out at any time by saying *"disconnect my Microsoft account"* — this deletes their cache file from the server.

---

## License

MIT
