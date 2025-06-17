# Entra Authentication Options for Outlook MCP

## Current Implementation: Hybrid Auth
- **User Auth**: GitHub OAuth  
- **API Access**: Azure Client Credentials
- **User Control**: GitHub username allowlist

## Option 1: Pure Entra OAuth with Delegated Permissions (Recommended)

### Benefits
- ✅ Single identity provider (Microsoft)
- ✅ User's own data access (delegated permissions)
- ✅ Native integration with Microsoft 365
- ✅ Conditional Access policy support
- ✅ MFA enforcement
- ✅ Device compliance checks

### Implementation
```typescript
// User Flow:
User → Microsoft OAuth → Access Token → Graph API (as user)

// Permissions Required:
- Calendars.ReadWrite (delegated)
- Mail.ReadWrite (delegated)  
- People.Read (delegated)
- User.Read (delegated)
```

### Configuration
```bash
# Azure App Registration Settings
- Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
- Redirect URIs: https://your-worker.workers.dev/auth/callback
- API Permissions: Microsoft Graph (Delegated)
  - Calendars.ReadWrite
  - Mail.ReadWrite
  - People.Read
  - User.Read
```

## Option 2: Entra with Application Permissions + Conditional Access

### Benefits
- ✅ Centralized admin control
- ✅ Conditional Access policies
- ✅ Application-level permissions
- ✅ Admin consent workflow

### Implementation
```typescript
// User Flow:
User → Entra Auth → Admin Consent → App Permissions → Graph API

// Permissions Required:
- Calendars.ReadWrite (application)
- Mail.ReadWrite (application)
- People.Read.All (application)
```

## Option 3: Azure AD B2C (For External Users)

### Benefits
- ✅ Support for external identities
- ✅ Custom user journeys
- ✅ Social identity providers
- ✅ Custom policies

## Comparison Table

| Feature | Current (Hybrid) | Option 1 (Delegated) | Option 2 (Application) | Option 3 (B2C) |
|---------|------------------|----------------------|------------------------|-----------------|
| Identity Provider | GitHub + Azure | Microsoft Only | Microsoft Only | Microsoft + Others |
| Data Access | Application-level | User-level | Application-level | User-level |
| Conditional Access | ❌ | ✅ | ✅ | ✅ |
| MFA Support | ❌ | ✅ | ✅ | ✅ |
| Admin Control | GitHub usernames | Email allowlist | Azure AD groups | B2C policies |
| User Experience | 2-step auth | Native M365 | Native M365 | Custom |

## Recommendation: Option 1 (Pure Entra OAuth)

**Why?**
1. **Native Microsoft Experience**: Users authenticate with their Microsoft accounts
2. **Delegated Permissions**: Access user's own data (more secure)
3. **Conditional Access**: Leverage existing organizational policies
4. **MFA Support**: Automatic enforcement of security policies
5. **Simpler Architecture**: Single authentication flow

## Migration Steps

### 1. Update Azure App Registration
```bash
# Add redirect URIs
https://your-worker.workers.dev/auth/callback

# Configure delegated permissions
Calendars.ReadWrite (delegated)
Mail.ReadWrite (delegated)
People.Read (delegated)
User.Read (delegated)
```

### 2. Update Cloudflare Environment
```bash
# Remove GitHub OAuth variables
# Keep Azure variables
AZURE_CLIENT_ID="your-app-id"
AZURE_CLIENT_SECRET="your-app-secret"  
AZURE_TENANT_ID="your-tenant-id"
ALLOWED_USER_EMAILS="user1@domain.com,user2@domain.com"
```

### 3. Update Worker Code
- Replace GitHub OAuth with Microsoft OAuth
- Use delegated permissions instead of client credentials
- Implement token refresh logic
- Add user email validation

### 4. Optional: Add Conditional Access
```bash
# In Azure AD, create Conditional Access policy:
1. Target: Your MCP app
2. Conditions: Device compliance, location, etc.
3. Access controls: Require MFA, compliant device, etc.
```

## Advanced: Group-Based Access Control

Instead of email allowlists, use Azure AD groups:

```typescript
// Check group membership
const groupsResponse = await fetch('https://graph.microsoft.com/v1.0/me/memberOf', {
  headers: { 'Authorization': `Bearer ${userToken}` }
});

const groups = await groupsResponse.json();
const allowedGroupId = 'your-mcp-users-group-id';
const hasAccess = groups.value.some(group => group.id === allowedGroupId);
```

This provides better admin control and scales with your organization.