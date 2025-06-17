# Outlook MCP Remote Server

A secure remote MCP (Model Context Protocol) server that runs on Cloudflare Workers, providing AI assistants with access to Microsoft Outlook calendar and email functionality through native Microsoft authentication.

## ✨ Features

- **📅 Calendar Management**: List, create, update, and manage calendar events
- **📧 Email Operations**: Read, send, and manage emails 
- **👥 People Search**: Find and access contact information
- **🔐 Secure Authentication**: Native Microsoft Entra OAuth with delegated permissions
- **🛡️ Access Control**: Email-based user authorization
- **☁️ Serverless**: Runs on Cloudflare Workers for global performance

## 🏗️ Architecture

```
User → Microsoft OAuth → Cloudflare Worker → Microsoft Graph API (Delegated Permissions)
```

**Benefits of Pure Entra Authentication:**
- ✅ Native Microsoft 365 experience
- ✅ User accesses their own data (delegated permissions)
- ✅ Supports Conditional Access policies
- ✅ MFA enforcement
- ✅ Device compliance checks
- ✅ No third-party OAuth providers

## 🚀 Quick Start

### 1. Azure App Registration Setup

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Click "New registration":
   - **Name**: `Outlook MCP Remote Server`
   - **Supported account types**: `Accounts in any organizational directory and personal Microsoft accounts`
   - **Redirect URI**: 
     - Type: `Web`
     - Local: `http://localhost:8787/auth/callback`
     - Production: `https://outlook-mcp-remote.<your-account>.workers.dev/auth/callback`

3. Note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

4. Create a client secret:
   - Go to "Certificates & secrets" > "New client secret"
   - Copy the secret value (you won't see it again!)

5. Configure API permissions:
   - Go to "API permissions" > "Add a permission"
   - Choose "Microsoft Graph" > "Delegated permissions"
   - Add these permissions:
     - `Calendars.ReadWrite` - Manage user's calendar
     - `Mail.ReadWrite` - Read and send emails
     - `People.Read` - Access contact information
     - `User.Read` - Read user profile
   - Click "Grant admin consent" (if you're an admin)

### 2. Cloudflare Setup

1. **Create KV Namespace**:
   ```bash
   wrangler kv:namespace create "AUTH_STATE"
   wrangler kv:namespace create "AUTH_STATE" --preview
   ```

2. **Update wrangler.toml** with your KV namespace IDs:
   ```toml
   [[kv_namespaces]]
   binding = "AUTH_STATE"
   id = "your-production-kv-id"
   preview_id = "your-preview-kv-id"
   ```

3. **Set Environment Variables**:
   ```bash
   # Public variables (in wrangler.toml)
   AZURE_CLIENT_ID = "your-app-client-id"
   AZURE_TENANT_ID = "your-tenant-id"
   ALLOWED_USER_EMAILS = "user1@domain.com,user2@domain.com"

   # Secret variables (via wrangler)
   wrangler secret put AZURE_CLIENT_SECRET
   # Enter your client secret when prompted
   ```

### 3. Local Development

1. **Setup environment**:
   ```bash
   cp .dev.vars.example .dev.vars
   # Edit .dev.vars with your Azure client secret
   ```

2. **Start development server**:
   ```bash
   npm run build:worker
   npm run dev:worker
   ```

3. **Test authentication**:
   - Visit: `http://localhost:8787`
   - Click "Sign in with Microsoft"
   - Complete OAuth flow

### 4. Production Deployment

1. **Deploy to Cloudflare**:
   ```bash
   npm run build:worker
   npm run deploy
   ```

2. **Update Azure App Registration**:
   - Add production redirect URI: `https://outlook-mcp-remote.<your-account>.workers.dev/auth/callback`

## 🔧 MCP Client Configuration

### Claude Desktop

Add to your MCP configuration file (`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS):

```json
{
  "mcpServers": {
    "outlook-remote": {
      "command": "npx",
      "args": [
        "@modelcontextprotocol/server-remote",
        "https://outlook-mcp-remote.<your-account>.workers.dev/sse"
      ]
    }
  }
}
```

### Other MCP Clients

Use the MCP server URL: `https://outlook-mcp-remote.<your-account>.workers.dev/sse`

## 🛠️ Available Tools

### Calendar Tools
- **`listCalendarEvents`** - List events for a date range
- **`createCalendarEvent`** - Create new calendar events

### Email Tools  
- **`listEmails`** - List emails from folders (inbox, sent, etc.)
- **`sendEmail`** - Send new email messages

### People Tools
- **`searchPeople`** - Search for contacts and colleagues

## 🔐 Security Features

### Authentication Flow
1. User visits worker URL
2. Redirected to Microsoft OAuth
3. User authenticates with their Microsoft account
4. Token exchanged and validated
5. User email checked against allowlist
6. Session created with delegated permissions

### Access Control
- **Email Allowlist**: Only specified users can access the server
- **Delegated Permissions**: Users access only their own data
- **Token Expiration**: Sessions automatically expire
- **Secure Storage**: Sessions stored in Cloudflare KV

### Optional: Conditional Access
You can enhance security by configuring Conditional Access policies in Azure AD:

1. Go to Azure AD > Security > Conditional Access
2. Create new policy targeting your MCP app
3. Configure conditions (device compliance, location, MFA)
4. Set access controls (require MFA, compliant device, etc.)

## 🔄 Advanced Configuration

### Group-Based Access Control

Instead of email allowlists, use Azure AD groups:

1. Create a security group in Azure AD
2. Add authorized users to the group  
3. Update the worker code to check group membership:

```typescript
// Check group membership instead of email allowlist
const groupsResponse = await graphHelper.makeGraphRequest('/me/memberOf');
const groups = groupsResponse.value || [];
const allowedGroupId = 'your-mcp-users-group-id';
const hasAccess = groups.some(group => group.id === allowedGroupId);
```

### Multi-Tenant Support

For supporting multiple organizations:

1. Change account type to "Multitenant" in Azure app registration
2. Update tenant ID in auth URL to "common":
   ```typescript
   const authUrl = new URL('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
   ```

## 🐛 Troubleshooting

### Common Issues

**Authentication Failures**:
- Verify redirect URIs match exactly (including protocol)
- Check client secret is correct and not expired
- Ensure API permissions are granted

**Access Denied**:
- Verify user email is in `ALLOWED_USER_EMAILS`
- Check if user account is enabled
- Confirm permissions are granted for the application

**Token Expired**:
- Sessions expire automatically (default 1 hour)
- Users need to re-authenticate
- Consider implementing token refresh (advanced)

### Debugging

1. **Check Cloudflare Workers logs**:
   ```bash
   wrangler tail
   ```

2. **Verify KV storage**:
   ```bash
   wrangler kv:key list --binding=AUTH_STATE
   ```

3. **Test Graph API access**:
   - Use [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
   - Test with same permissions as your app

## 📦 Dependencies

- `@modelcontextprotocol/sdk` - MCP protocol implementation
- `@cloudflare/workers-types` - TypeScript types for Workers
- `zod` - Schema validation

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

ISC License - see package.json for details

## 🆘 Support

For issues and questions:
1. Check the troubleshooting section
2. Review Cloudflare Workers docs
3. Check Microsoft Graph API documentation
4. Open an issue in the repository