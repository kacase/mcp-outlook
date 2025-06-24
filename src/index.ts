#!/usr/bin/env node

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { OutlookMcpHttpServer } from './httpServer.js';
import { setupMcpServer } from './mcpServer.js';

// Parse command line arguments
const args = process.argv.slice(2);
const useHttp = args.includes('--http');
const useOAuth = args.includes('--oauth');
const port = parseInt(args.find(arg => arg.startsWith('--port='))?.split('=')[1] || '3000');
const host = args.find(arg => arg.startsWith('--host='))?.split('=')[1] || '127.0.0.1';

if (useHttp) {
  // Start HTTP server
  const httpServer = new OutlookMcpHttpServer({
    port,
    host, 
    oauth: useOAuth
  });
  
  httpServer.start().catch(console.error);
  
  // Handle graceful shutdown
  process.on('SIGINT', async () => {
    console.log('\nShutting down...');
    await httpServer.stop();
    process.exit(0);
  });
} else {
  // Create server instance for stdio
  const server = setupMcpServer();
  
  // Connect the server to stdio transport
  server.connect(new StdioServerTransport());
}