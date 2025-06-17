# Outlook MCP Remote Server (Cloudflare Workers)

This is a remote MCP (Model Context Protocol) server that runs on Cloudflare Workers, providing access to Microsoft Outlook calendar and email functionality.

## Architecture

- **Authentication**: GitHub OAuth for user authentication
- **Authorization**: Username-based access control
- **Backend**: Microsoft Graph API for Outlook integration
- **Deployment**: Cloudflare Workers with KV storage

## Setup Instructions

### 1. GitHub OAuth App Setup

Create a GitHub OAuth App with these settings:

**For Local Development:**
- Application name: `Outlook MCP Local`
- Homepage URL: `http://localhost:8787`
- Authorization callback URL: `http://localhost:8787/token`

**For Production:**
- Application name: `Outlook MCP Production`
- Homepage URL: `https://outlook-mcp-remote.<your-account>.workers.dev`
- Authorization callback URL: `https://outlook-mcp-remote.<your-account>.workers.dev/token`

### 2. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Create a new registration:
   - Name: `Outlook MCP Worker`
   - Supported account types: Accounts in this organizational directory only
   - Redirect URI: Not needed for client credentials flow
3. Note down:
   - Application (client) ID
   - Directory (tenant) ID
4. Create a client secret in "Certificates & secrets"
5. Grant API permissions:
   - Microsoft Graph > Application permissions
   - Add: `Calendars.ReadWrite`, `Mail.ReadWrite`, `People.Read`
   - Admin consent required

### 3. Cloudflare Setup

1. Create a KV namespace:
   ```bash
   wrangler kv:namespace create "AUTH_STATE"
   wrangler kv:namespace create "AUTH_STATE" --preview
   ```

2. Update `wrangler.toml` with your KV namespace IDs

3. Set environment variables:
   ```bash
   # Public variables
   wrangler secret put OAUTH_CLIENT_SECRET
   wrangler secret put AZURE_CLIENT_ID  
   wrangler secret put AZURE_CLIENT_SECRET
   wrangler secret put AZURE_TENANT_ID
   ```

4. Update `wrangler.toml`:
   ```toml
   OAUTH_CLIENT_ID = "your-github-client-id"
   ALLOWED_USERNAMES = "your-github-username,another-user"
   ```

### 4. Local Development

1. Copy environment variables:
   ```bash
   cp .dev.vars.example .dev.vars
   # Fill in your actual values
   ```

2. Start development server:
   ```bash
   npm run dev:worker
   ```

3. Access at: `http://localhost:8787`

### 5. Production Deployment

```bash
npm run build:worker
npm run deploy
```

## Usage

### MCP Client Configuration

Add to your MCP client configuration:

```json
{
  "mcpServers": {
    "outlook-remote": {
      "command": "npx",
      "args": ["@modelcontextprotocol/server-remote", "https://outlook-mcp-remote.<your-account>.workers.dev/sse"]
    }
  }
}
```

### Available Tools

- **Calendar**: `listCalendarEvents`, `createCalendarEvent`
- **Email**: `listEmails`, `sendEmail`
- **People**: `searchPeople`

### Authentication Flow

1. Visit your Worker URL
2. Click "Authorize with GitHub"
3. Complete OAuth flow
4. Access is granted based on `ALLOWED_USERNAMES`

## Security

- GitHub OAuth provides user authentication
- Username-based authorization controls access
- Azure client credentials for Graph API access
- All secrets stored in Cloudflare Workers secrets

## Troubleshooting

- Check Cloudflare Workers logs: `wrangler tail`
- Verify OAuth callback URLs match exactly
- Ensure Azure app has proper permissions and admin consent
- Check KV namespace bindings in wrangler.toml