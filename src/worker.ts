import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { 
  CreateEventSchema, 
  SendEmailSchema,
  ListEventsQuerySchema,
  ListEmailsQuerySchema,
  SearchPeopleQuerySchema
} from './types.js';

// Import Cloudflare Workers types
/// <reference types="@cloudflare/workers-types" />

// Worker environment interface
interface Env {
  AUTH_STATE: KVNamespace;
  AZURE_CLIENT_ID: string;
  AZURE_CLIENT_SECRET: string;
  AZURE_TENANT_ID: string;
  ALLOWED_USER_EMAILS: string; // comma-separated list of allowed emails
}

// Session data interface
interface SessionData {
  userId: string;
  email: string;
  name: string;
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
}

// Microsoft Graph helper with delegated permissions
class GraphHelper {
  private accessToken: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
  }

  async makeGraphRequest(url: string, options: RequestInit = {}): Promise<any> {
    const response = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
      ...options,
      headers: {
        'Authorization': `Bearer ${this.accessToken}`,
        'Content-Type': 'application/json',
        ...options.headers
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph API request failed: ${response.status} - ${errorText}`);
    }

    if (response.status === 204) return null;
    return response.json();
  }
}

// MCP Server class
class OutlookMCPServer {
  private server: McpServer;

  constructor() {
    this.server = new McpServer({
      name: 'outlook-mcp-remote',
      version: '1.0.0'
    });

    this.setupTools();
  }

  private setupTools(): void {
    // ============= Calendar Tools =============
    
    this.server.tool(
      'listCalendarEvents',
      'Lists the user\'s calendar events for a specified time range',
      ListEventsQuerySchema.shape,
      async (params: any, { env, sessionData }: { env: Env; sessionData: SessionData }) => {
        const graphHelper = new GraphHelper(sessionData.accessToken);
        let url = '/me/calendar/events';
        const queryParams = new URLSearchParams();

        if (params.startDateTime || params.endDateTime) {
          const filters = [];
          if (params.startDateTime) filters.push(`start/dateTime ge '${params.startDateTime}'`);
          if (params.endDateTime) filters.push(`end/dateTime le '${params.endDateTime}'`);
          queryParams.append('$filter', filters.join(' and '));
        }
        if (params.top) queryParams.append('$top', params.top.toString());
        if (params.orderBy) queryParams.append('$orderby', params.orderBy);

        if (queryParams.toString()) url += `?${queryParams.toString()}`;

        const response = await graphHelper.makeGraphRequest(url);
        const events = response.value || [];

        const formattedEvents = events.map((event: any) => ({
          id: event.id,
          subject: event.subject,
          start: new Date(event.start.dateTime).toLocaleString(),
          end: new Date(event.end.dateTime).toLocaleString(),
          timeZone: event.start.timeZone,
          location: event.location?.displayName || 'No location',
          isAllDay: event.isAllDay || false,
          attendees: event.attendees?.map((a: any) => a.emailAddress.address).join(', ') || 'No attendees',
          preview: event.bodyPreview || ''
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(formattedEvents, null, 2)
            }
          ]
        };
      }
    );

    this.server.tool(
      'createCalendarEvent',
      'Creates a new calendar event',
      CreateEventSchema.shape,
      async (params: any, { env, sessionData }: { env: Env; sessionData: SessionData }) => {
        const graphHelper = new GraphHelper(sessionData.accessToken);
        const eventData = {
          subject: params.subject,
          body: { contentType: 'text', content: params.body || '' },
          start: { dateTime: params.startDateTime, timeZone: params.timeZone || 'UTC' },
          end: { dateTime: params.endDateTime, timeZone: params.timeZone || 'UTC' },
          attendees: params.attendees?.map((email: string) => ({
            emailAddress: { address: email, name: email },
            type: 'required'
          })) || [],
          location: params.location ? { displayName: params.location } : undefined,
          isAllDay: params.isAllDay || false
        };

        const createdEvent = await graphHelper.makeGraphRequest('/me/calendar/events', {
          method: 'POST',
          body: JSON.stringify(eventData)
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(createdEvent, null, 2)
            }
          ]
        };
      }
    );

    // ============= Email Tools =============

    this.server.tool(
      'listEmails',
      'Lists the user\'s emails from a specified folder',
      ListEmailsQuerySchema.shape,
      async (params: any, { env, sessionData }: { env: Env; sessionData: SessionData }) => {
        const graphHelper = new GraphHelper(sessionData.accessToken);
        let url = `/me/mailFolders/${params.folder || 'inbox'}/messages`;
        const queryParams = new URLSearchParams();

        if (params.top) queryParams.append('$top', params.top.toString());
        if (params.orderBy) queryParams.append('$orderby', params.orderBy);
        if (params.filter) queryParams.append('$filter', params.filter);
        if (params.select) queryParams.append('$select', params.select);

        if (queryParams.toString()) url += `?${queryParams.toString()}`;

        const response = await graphHelper.makeGraphRequest(url);
        const emails = response.value || [];

        const formattedEmails = emails.map((email: any) => ({
          id: email.id,
          subject: email.subject || '(No Subject)',
          from: email.from?.emailAddress.address || 'Unknown',
          fromName: email.from?.emailAddress.name,
          received: email.receivedDateTime ? new Date(email.receivedDateTime).toLocaleString() : 'Unknown',
          isRead: email.isRead,
          importance: email.importance || 'normal',
          hasAttachments: email.hasAttachments || false,
          preview: email.bodyPreview || ''
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(formattedEmails, null, 2)
            }
          ]
        };
      }
    );

    this.server.tool(
      'sendEmail',
      'Sends a new email message',
      SendEmailSchema.shape,
      async (params: any, { env, sessionData }: { env: Env; sessionData: SessionData }) => {
        const graphHelper = new GraphHelper(sessionData.accessToken);
        const emailData = {
          message: {
            subject: params.subject,
            body: { contentType: params.bodyType || 'text', content: params.body },
            toRecipients: params.to.map((email: string) => ({ emailAddress: { address: email } })),
            ccRecipients: params.cc?.map((email: string) => ({ emailAddress: { address: email } })) || [],
            bccRecipients: params.bcc?.map((email: string) => ({ emailAddress: { address: email } })) || []
          }
        };

        await graphHelper.makeGraphRequest('/me/sendMail', {
          method: 'POST',
          body: JSON.stringify(emailData)
        });

        return {
          content: [
            {
              type: "text",
              text: 'Email sent successfully'
            }
          ]
        };
      }
    );

    // ============= People Tools =============

    this.server.tool(
      'searchPeople',
      'Searches for people relevant to the current user',
      SearchPeopleQuerySchema.shape,
      async (params: any, { env, sessionData }: { env: Env; sessionData: SessionData }) => {
        const graphHelper = new GraphHelper(sessionData.accessToken);
        let url = '/me/people';
        const queryParams = new URLSearchParams();

        if (params.searchTerm) queryParams.append('$search', `"${params.searchTerm}"`);
        if (params.filter) queryParams.append('$filter', params.filter);
        if (params.select) queryParams.append('$select', params.select);
        if (params.top) queryParams.append('$top', params.top.toString());

        if (queryParams.toString()) url += `?${queryParams.toString()}`;

        const response = await graphHelper.makeGraphRequest(url);
        const people = response.value || [];

        const formattedPeople = people.map((person: any) => {
          const primaryEmail = person.scoredEmailAddresses?.[0]?.address || '';
          const personClass = person.personType?.class || 'Unknown';
          const personSubclass = person.personType?.subclass || '';
          
          return {
            id: person.id,
            displayName: person.displayName || 'Unknown',
            email: primaryEmail,
            jobTitle: person.jobTitle || '',
            department: person.department || '',
            type: `${personClass}${personSubclass ? ` (${personSubclass})` : ''}`
          };
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(formattedPeople, null, 2)
            }
          ]
        };
      }
    );
  }

  getServer(): McpServer {
    return this.server;
  }
}

// Main Worker export
export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    
    // Handle authentication routes
    if (url.pathname === '/') {
      return new Response(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Outlook MCP Remote Server</title>
          <style>
            body { 
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
              max-width: 800px; 
              margin: 0 auto; 
              padding: 20px;
              background: #f8f9fa;
            }
            .container { 
              text-align: center; 
              margin-top: 50px; 
              background: white;
              padding: 40px;
              border-radius: 8px;
              box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            .btn { 
              display: inline-block; 
              padding: 12px 24px; 
              background: #0078d4; 
              color: white; 
              text-decoration: none; 
              border-radius: 4px; 
              margin: 10px;
              font-weight: 500;
            }
            .btn:hover { 
              background: #106ebe; 
            }
            .features {
              margin-top: 30px;
              text-align: left;
            }
            .feature {
              margin: 10px 0;
              padding: 10px;
              background: #f0f9ff;
              border-left: 4px solid #0078d4;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>🎯 Outlook MCP Remote Server</h1>
            <p>This is a remote MCP server that provides AI assistants with secure access to Microsoft Outlook calendar and email functionality.</p>
            
            <div class="features">
              <h3>Available Features:</h3>
              <div class="feature">📅 <strong>Calendar Management</strong> - List, create, and manage calendar events</div>
              <div class="feature">📧 <strong>Email Operations</strong> - Read, send, and manage emails</div>
              <div class="feature">👥 <strong>People Search</strong> - Find and access contact information</div>
            </div>
            
            <p style="margin-top: 30px;"><strong>Sign in with your Microsoft account to continue.</strong></p>
            <a href="/auth/login" class="btn">🔐 Sign in with Microsoft</a>
          </div>
        </body>
        </html>
      `, {
        headers: { 'Content-Type': 'text/html' }
      });
    }

    // Microsoft OAuth login initiation
    if (url.pathname === '/auth/login') {
      const state = crypto.randomUUID();
      const nonce = crypto.randomUUID();
      
      // Store state in KV for validation
      await env.AUTH_STATE.put(`state:${state}`, nonce, { expirationTtl: 600 }); // 10 minutes
      
      const authUrl = new URL(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/authorize`);
      authUrl.searchParams.set('client_id', env.AZURE_CLIENT_ID);
      authUrl.searchParams.set('response_type', 'code');
      authUrl.searchParams.set('redirect_uri', `${url.origin}/auth/callback`);
      authUrl.searchParams.set('scope', 'openid profile email https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/People.Read');
      authUrl.searchParams.set('state', state);
      authUrl.searchParams.set('nonce', nonce);
      
      return Response.redirect(authUrl.toString(), 302);
    }

    // Microsoft OAuth callback
    if (url.pathname === '/auth/callback') {
      const code = url.searchParams.get('code');
      const state = url.searchParams.get('state');
      const error = url.searchParams.get('error');
      
      if (error) {
        return new Response(`Authentication failed: ${error}`, { status: 400 });
      }
      
      if (!code || !state) {
        return new Response('Missing code or state parameter', { status: 400 });
      }
      
      // Validate state
      const storedNonce = await env.AUTH_STATE.get(`state:${state}`);
      if (!storedNonce) {
        return new Response('Invalid or expired state', { status: 400 });
      }
      
      try {
        // Exchange code for tokens
        const tokenResponse = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            client_id: env.AZURE_CLIENT_ID,
            client_secret: env.AZURE_CLIENT_SECRET,
            code: code,
            grant_type: 'authorization_code',
            redirect_uri: `${url.origin}/auth/callback`,
          })
        });
        
        if (!tokenResponse.ok) {
          const errorData = await tokenResponse.text();
          return new Response(`Token exchange failed: ${errorData}`, { status: 400 });
        }
        
        const tokens = await tokenResponse.json() as any;
        
        // Get user info from Microsoft Graph
        const userResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: { 'Authorization': `Bearer ${tokens.access_token}` }
        });
        
        if (!userResponse.ok) {
          return new Response('Failed to get user info', { status: 400 });
        }
        
        const user = await userResponse.json() as any;
        
        // Check if user is authorized
        const allowedEmails = env.ALLOWED_USER_EMAILS.split(',').map(email => email.trim().toLowerCase());
        const userEmail = (user.mail || user.userPrincipalName || '').toLowerCase();
        if (!allowedEmails.includes(userEmail)) {
          return new Response('Access denied: User not authorized', { status: 403 });
        }
        
        // Store user session
        const sessionId = crypto.randomUUID();
        const sessionData: SessionData = {
          userId: user.id,
          email: user.mail || user.userPrincipalName,
          name: user.displayName,
          accessToken: tokens.access_token,
          refreshToken: tokens.refresh_token,
          expiresAt: Date.now() + (tokens.expires_in * 1000)
        };
        
        await env.AUTH_STATE.put(`session:${sessionId}`, JSON.stringify(sessionData), { 
          expirationTtl: tokens.expires_in || 3600 
        });
        
        // Set session cookie and redirect
        return new Response(`
          <!DOCTYPE html>
          <html>
          <head>
            <title>Authentication Successful</title>
            <style>
              body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; text-align: center; padding: 50px; }
              .success { color: #28a745; font-size: 24px; margin-bottom: 20px; }
              .info { background: #e9ecef; padding: 20px; border-radius: 8px; margin: 20px 0; }
              code { background: #f8f9fa; padding: 4px 8px; border-radius: 4px; font-family: monospace; }
            </style>
          </head>
          <body>
            <div class="success">✅ Authentication Successful!</div>
            <div class="info">
              <h3>Welcome, ${user.displayName}!</h3>
              <p>Email: ${userEmail}</p>
              <p>You can now use this MCP server in your AI clients.</p>
            </div>
            <p><strong>MCP Server URL:</strong></p>
            <p><code>${url.origin}/sse</code></p>
            <p><small>Session expires in ${Math.round((tokens.expires_in || 3600) / 60)} minutes</small></p>
          </body>
          </html>
        `, {
          headers: {
            'Content-Type': 'text/html',
            'Set-Cookie': `session=${sessionId}; HttpOnly; Secure; SameSite=Strict; Max-Age=${tokens.expires_in || 3600}`
          }
        });
        
      } catch (error) {
        return new Response(`Authentication error: ${error}`, { status: 500 });
      }
    }

    // MCP SSE endpoint
    if (url.pathname === '/sse') {
      // Extract session from cookie or Authorization header
      const sessionId = extractSessionId(request);
      if (!sessionId) {
        return new Response('Unauthorized: No session found', { status: 401 });
      }

      // Get session data
      const sessionDataRaw = await env.AUTH_STATE.get(`session:${sessionId}`);
      if (!sessionDataRaw) {
        return new Response('Unauthorized: Invalid or expired session', { status: 401 });
      }

      const sessionData: SessionData = JSON.parse(sessionDataRaw);
      
      // Check if token is expired
      if (Date.now() >= sessionData.expiresAt) {
        return new Response('Unauthorized: Token expired', { status: 401 });
      }

      // Handle MCP protocol
      const mcpServer = new OutlookMCPServer();
      return handleMCPRequest(request, mcpServer.getServer(), { env, sessionData });
    }

    return new Response('Not Found', { status: 404 });
  }
};

// Helper function to extract session ID
function extractSessionId(request: Request): string | null {
  // Try cookie first
  const cookies = request.headers.get('Cookie');
  if (cookies) {
    const sessionMatch = cookies.match(/session=([^;]+)/);
    if (sessionMatch) return sessionMatch[1];
  }
  
  // Try Authorization header
  const auth = request.headers.get('Authorization');
  if (auth && auth.startsWith('Bearer ')) {
    return auth.slice(7);
  }
  
  return null;
}

// Helper function to handle MCP requests
async function handleMCPRequest(
  request: Request, 
  server: McpServer, 
  context: { env: Env; sessionData: SessionData }
): Promise<Response> {
  if (request.method === 'POST') {
    try {
      const body = await request.json();
      
      // Handle MCP method calls
      if (body.method === 'tools/call') {
        const toolName = body.params.name;
        const toolArgs = body.params.arguments || {};
        
        // Find and call the tool
        const result = await server.tool(toolName, '', {}, async () => {
          // This is a simplified implementation
          // In a real implementation, you'd need to properly handle the MCP protocol
          return { content: [{ type: "text", text: "Tool execution not fully implemented yet" }] };
        })(toolArgs, context);
        
        return new Response(JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result
        }), {
          headers: { 'Content-Type': 'application/json' }
        });
      }
      
      return new Response(JSON.stringify({
        jsonrpc: '2.0',
        id: body.id,
        error: { code: -32601, message: 'Method not found' }
      }), {
        headers: { 'Content-Type': 'application/json' }
      });
      
    } catch (error) {
      return new Response(JSON.stringify({
        jsonrpc: '2.0',
        error: { code: -32603, message: 'Internal error' }
      }), {
        status: 500,
        headers: { 'Content-Type': 'application/json' }
      });
    }
  }
  
  return new Response('Method not allowed', { status: 405 });
}