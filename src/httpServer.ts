#!/usr/bin/env node

import express from 'express';
import cors from 'cors';
import { createServer } from 'http';
import { tokenValidator } from './oauthMiddleware.js';
import { setupMcpServer } from './mcpServer.js';
import { randomBytes } from 'crypto';

interface ServerOptions {
  port: number;
  host: string;
  oauth: boolean;
}

export class OutlookMcpHttpServer {
  private app: express.Application;
  private server: ReturnType<typeof createServer>;
  private mcpServer: ReturnType<typeof setupMcpServer>;
  private options: ServerOptions;
  private sessions: Map<string, any> = new Map();

  constructor(options: ServerOptions) {
    this.options = options;
    this.app = express();
    this.server = createServer(this.app);
    this.mcpServer = setupMcpServer();
    this.setupMiddleware();
    this.setupRoutes();
  }

  private generateSecureId(): string {
    return randomBytes(32).toString('hex');
  }

  private setupMiddleware(): void {
    // CORS for localhost only (security requirement)
    this.app.use(cors({
      origin: (origin: string | undefined, callback: (err: Error | null, allow?: boolean) => void) => {
        // Allow requests with no origin (like mobile apps, curl requests)
        if (!origin) return callback(null, true);
        
        // Only allow localhost origins for security
        const allowedOrigins = [
          'http://localhost:3000',
          'http://127.0.0.1:3000',
          'https://localhost:3000',
          'https://127.0.0.1:3000'
        ];
        
        if (allowedOrigins.includes(origin)) {
          callback(null, true);
        } else {
          callback(new Error('Not allowed by CORS policy'));
        }
      },
      credentials: true
    }));

    this.app.use(express.json({ limit: '10mb' }));
    this.app.use(express.urlencoded({ extended: true }));

    // Validate Origin header (MCP security requirement)
    this.app.use((req, res, next) => {
      const origin = req.headers.origin;
      if (origin && !origin.includes('localhost') && !origin.includes('127.0.0.1')) {
        res.status(403).json({
          error: 'forbidden',
          error_description: 'Invalid origin'
        });
        return;
      }
      next();
    });
  }

  private setupRoutes(): void {
    // OAuth metadata endpoint
    this.app.get('/.well-known/oauth-authorization-server', (_req, res) => {
      res.json({
        issuer: `https://login.microsoftonline.com/${process.env.AUTHORITY || 'common'}/v2.0`,
        authorization_endpoint: `https://login.microsoftonline.com/${process.env.AUTHORITY || 'common'}/oauth2/v2.0/authorize`,
        token_endpoint: `https://login.microsoftonline.com/${process.env.AUTHORITY || 'common'}/oauth2/v2.0/token`,
        userinfo_endpoint: 'https://graph.microsoft.com/oidc/userinfo',
        jwks_uri: `https://login.microsoftonline.com/${process.env.AUTHORITY || 'common'}/discovery/v2.0/keys`,
        response_types_supported: ['code'],
        grant_types_supported: ['authorization_code', 'refresh_token'],
        subject_types_supported: ['public'],
        id_token_signing_alg_values_supported: ['RS256'],
        scopes_supported: [
          'openid',
          'profile',
          'email',
          'offline_access',
          'User.Read',
          'Calendars.Read',
          'Calendars.ReadWrite',
          'Mail.Send',
          'Mail.ReadWrite',
          'Mail.Read',
          'People.Read'
        ]
      });
    });

    // MCP endpoint handles both POST and GET (Streamable HTTP)
    const mcpHandler = this.options.oauth 
      ? [tokenValidator.createMiddleware(['User.Read']), this.createMcpHandler()]
      : [this.createMcpHandler()];

    this.app.post('/mcp', mcpHandler as any);
    this.app.get('/mcp', mcpHandler as any);
    
    // Session termination endpoint
    this.app.delete('/mcp', (req, res) => {
      const sessionId = req.headers['mcp-session-id'] as string;
      if (sessionId && this.sessions.has(sessionId)) {
        this.sessions.delete(sessionId);
        res.status(200).json({ message: 'Session terminated' });
      } else {
        res.status(404).json({ error: 'Session not found' });
      }
    });

    // Health check
    this.app.get('/health', (_req, res) => {
      res.json({ status: 'ok', oauth: this.options.oauth });
    });
  }

  private createMcpHandler() {
    return async (req: express.Request, res: express.Response) => {
      // Set MCP protocol version header
      res.setHeader('MCP-Protocol-Version', '2025-03-26');

      if (req.method === 'POST') {
        // Handle JSON-RPC request
        try {
          let sessionId = req.headers['mcp-session-id'] as string;
          
          // Handle session management for initialize requests
          if (req.body.method === 'initialize') {
            sessionId = this.generateSecureId();
            res.setHeader('Mcp-Session-Id', sessionId);
            this.sessions.set(sessionId, { created: Date.now() });
          }

          // Create a simple request handler that mimics the MCP server behavior
          const response = await this.handleMcpRequest(req.body);
          
          // For now, return single JSON response
          // In future, could implement SSE streaming for long-running operations
          res.json(response);
          
        } catch (error) {
          res.status(500).json({
            error: 'internal_error',
            message: error instanceof Error ? error.message : 'Unknown error'
          });
        }
        
      } else if (req.method === 'GET') {
        // Handle SSE stream for server-initiated communication
        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');
        res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
        res.setHeader('Access-Control-Allow-Credentials', 'true');
        
        // Send initial connection event
        const eventId = this.generateSecureId();
        res.write(`id: ${eventId}\n`);
        res.write(`event: connection\n`);
        res.write(`data: {"type":"connection","timestamp":"${new Date().toISOString()}"}\n\n`);
        
        // Keep connection alive with heartbeat
        const keepAlive = setInterval(() => {
          const heartbeatId = this.generateSecureId();
          res.write(`id: ${heartbeatId}\n`);
          res.write(`event: heartbeat\n`);
          res.write(`data: {"type":"heartbeat","timestamp":"${new Date().toISOString()}"}\n\n`);
        }, 30000);
        
        req.on('close', () => {
          clearInterval(keepAlive);
        });
      }
    };
  }

  private async handleMcpRequest(request: any): Promise<any> {
    // This is a simplified request handler that routes to the appropriate MCP server tools
    // In a full implementation, this would use the MCP SDK's request handling
    
    try {
      if (request.method === 'initialize') {
        return {
          jsonrpc: '2.0',
          id: request.id,
          result: {
            protocolVersion: '2025-03-26',
            capabilities: {
              tools: {},
              resources: {},
              prompts: {}
            },
            serverInfo: {
              name: 'outlook-mcp',
              version: '1.0.0'
            }
          }
        };
      }

      if (request.method === 'tools/list') {
        // Return list of available tools
        const tools = [
          { name: 'listCalendarEvents', description: 'Lists calendar events' },
          { name: 'createCalendarEvent', description: 'Creates a new calendar event' },
          { name: 'listEmails', description: 'Lists emails from folders' },
          { name: 'sendEmail', description: 'Sends an email' },
          // Add other tools...
        ];
        
        return {
          jsonrpc: '2.0',
          id: request.id,
          result: { tools }
        };
      }

      if (request.method === 'tools/call') {
        // This would call the actual tool implementation
        // For now, return a placeholder
        return {
          jsonrpc: '2.0',
          id: request.id,
          result: {
            content: [
              {
                type: 'text',
                text: `Tool ${request.params.name} called with params: ${JSON.stringify(request.params.arguments, null, 2)}`
              }
            ]
          }
        };
      }

      // Default response for unknown methods
      return {
        jsonrpc: '2.0',
        id: request.id,
        error: {
          code: -32601,
          message: 'Method not found',
          data: request.method
        }
      };

    } catch (error) {
      return {
        jsonrpc: '2.0',
        id: request.id,
        error: {
          code: -32603,
          message: 'Internal error',
          data: error instanceof Error ? error.message : String(error)
        }
      };
    }
  }

  async start(): Promise<void> {
    return new Promise((resolve, reject) => {
      this.server.listen(this.options.port, this.options.host, () => {
        console.log(`Outlook MCP server listening on ${this.options.host}:${this.options.port}`);
        console.log(`OAuth enabled: ${this.options.oauth}`);
        console.log(`Health check: http://${this.options.host}:${this.options.port}/health`);
        console.log(`MCP endpoint: http://${this.options.host}:${this.options.port}/mcp`);
        resolve();
      });

      this.server.on('error', reject);
    });
  }

  async stop(): Promise<void> {
    return new Promise((resolve) => {
      this.server.close(() => resolve());
    });
  }
}