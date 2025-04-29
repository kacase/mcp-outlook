# Outlook MCP (Model Context Protocol) Server

A Model Context Protocol server that integrates with Microsoft Outlook through Microsoft Graph API, allowing Claude and other LLMs to check calendar events, schedule new ones, read emails, and send messages.

## Features

- 📅 **Calendar Integration**: View, list, create, update, and delete calendar events
- 📧 **Email Integration**: Read, send, draft, and manage emails from your Outlook account
- 🔁 **Model Context Protocol**: Follows MCP standards for LLM tool integration
- 🛡️ **Type Safety**: Full TypeScript implementation with Zod validation

## Prerequisites

- Node.js 18+
- Microsoft 365 account with appropriate permissions
- Microsoft Azure App Registration with Graph API permissions (Calendar and Mail)

## Setup


1. Register an application in Azure Active Directory:
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to "App registrations"
   - Create a new registration with a redirect URI of type "Public client/native (mobile & desktop)"
     - Register `http://localhost` as the redirect URI

   - Configure API permissions:
     - Choose Microsoft Graph and type delegated, as we will act on the users behalf
     - For Calendar: "Calendars.Read" and "Calendars.ReadWrite"
     - For Email: "Mail.Read", "Mail.ReadWrite" and "Mail.Send"

2. Note the values from your Azure app registration (Overview) to use for the MCP config as environment variables:
  - Client ID (Application (client) ID)
  - Tenant ID (Directory (tenant) ID)

3. Register the MCP server
For Claude Desktop, create or update your configuration in `~/.claude/config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": [
        "@kacase/mcp-outlook"
      ],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "MCP_SERVER_NAME": "outlook-mcp",
        "MCP_SERVER_VERSION": "1.0.0"
      }
    }
  }
}
```

Make sure to replace the path and environment variables with your actual values.

## Available Tools

### Calendar Tools
- **listCalendarEvents**: Lists the user's calendar events for a specified time range
- **createCalendarEvent**: Creates a new calendar event
- **getCalendarEvent**: Gets details of a specific calendar event
- **updateCalendarEvent**: Updates an existing calendar event
- **deleteCalendarEvent**: Deletes a calendar event

### Email Tools
- **listEmails**: Lists emails from a specified folder (inbox, sent, drafts, etc.)
- **getEmail**: Gets details of a specific email message
- **sendEmail**: Sends a new email message
- **createDraft**: Creates a draft email message without sending it
- **markEmailAsRead**: Marks an email message as read
- **markEmailAsUnread**: Marks an email message as unread
- **deleteEmail**: Deletes an email message

### Resources
- **calendar**: Resource containing calendar events data
- **inbox**: Resource containing inbox messages data

## Development

Run in development mode with live reloading:
```
npm run dev
```

Run linting:
```
npm run lint
```

Configure your MCP locally
```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": [
        "/ABSOLUTE/PATH/TO/outlook_mcp/build/index.js"
      ],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "MCP_SERVER_NAME": "outlook-mcp",
        "MCP_SERVER_VERSION": "1.0.0"
      }
    }
  }
}
```
