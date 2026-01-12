# Outlook Email Integration

Connect your workspace to Microsoft Outlook to monitor emails, send messages, and manage your email workflow.

## Features

- **Email Sync**: Automatically sync important emails from your inbox and sent items
- **Send Emails**: Compose and send emails directly through the integration
- **Draft Management**: Create email drafts
- **Email Operations**: Reply, forward, delete, and move emails
- **Folder Management**: List and organize emails across folders
- **Search**: Search emails by content

## Setup

### 1. Create a Microsoft Azure App Registration

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: Your app name (e.g., "Core Outlook Integration")
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Add your callback URL (e.g., `https://your-domain.com/api/v1/oauth/callback`)
5. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**
3. Add these permissions:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
   - `offline_access`
4. Click **Grant admin consent** (if required by your organization)

### 3. Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and set expiration
4. Copy the secret value (you won't be able to see it again)

### 4. Configure the Integration

Add the following to your integration definition:

```json
{
  "name": "Outlook Email",
  "slug": "outlook",
  "config": {
    "clientId": "YOUR_CLIENT_ID",
    "clientSecret": "YOUR_CLIENT_SECRET"
  }
}
```

## Available MCP Tools

| Tool | Description |
|------|-------------|
| `send_email` | Send a new email |
| `create_draft` | Create an email draft |
| `read_email` | Read a specific email by ID |
| `search_emails` | Search emails by text query |
| `list_emails` | List emails from a folder |
| `reply_email` | Reply to an email |
| `forward_email` | Forward an email |
| `delete_email` | Delete an email |
| `mark_as_read` | Mark email as read/unread |
| `move_email` | Move email to a folder |
| `list_folders` | List all mail folders |

## Sync Behavior

The integration syncs every 15 minutes and retrieves:
- High importance emails from your inbox
- Recent sent emails

Activities are created for each synced email with the full content converted to markdown.

## OAuth Scopes

The integration requires these Microsoft Graph API scopes:
- `Mail.Read` - Read user mail
- `Mail.ReadWrite` - Read and write user mail
- `Mail.Send` - Send mail as the user
- `User.Read` - Read user profile
- `offline_access` - Maintain access for refresh tokens
