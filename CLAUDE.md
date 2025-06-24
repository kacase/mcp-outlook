# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build Commands
- `npm run build` - Build the TypeScript project
- `npm test` - Run tests (when implemented)
- `npm run lint` - Lint the code (when implemented)
- `npm start` - Start the MCP server (when implemented)

## Code Style Guidelines
- **TypeScript**: Use strict typing with explicit return types
- **Formatting**: Follow 2-space indentation, trailing commas
- **Imports**: Group by external packages first, then internal modules
- **Naming**: camelCase for variables/functions, PascalCase for classes/types
- **Error Handling**: Use typed error responses when possible
- **Modules**: Use ES modules (type: "module" is set in package.json)
- **SDK Usage**: Follow @modelcontextprotocol/sdk patterns for tools and resources

## Project Structure
- `/src` - TypeScript source files
- `/build` - Compiled JavaScript output

This project is a model context protocol server for Microsoft Outlook. It allows Claude to:

1. **Calendar functionality**:
   - Check calendar events
   - Schedule new events
   - Update existing events
   - Delete events

2. **Email functionality**:
   - Read emails from inbox and other folders
   - Send new emails
   - Create draft emails
   - Mark emails as read/unread
   - Delete emails

The server uses the Microsoft Graph API to interact with Outlook's calendar and email systems.

## OAuth and Authentication Notes
- Reference implementation: https://raw.githubusercontent.com/modelcontextprotocol/typescript-sdk/refs/heads/main/src/examples/server/simpleStreamableHttp.ts
- Investigating Microsoft OAuth middleware and Entra ID integration
- Discussion reference: https://github.com/modelcontextprotocol/modelcontextprotocol/pull/284

## MSAL and Authorization Specifications
- MSAL Tutorial: https://learn.microsoft.com/en-us/entra/identity-platform/tutorial-v2-nodejs-webapp-msal
- Authorization Specification: https://modelcontextprotocol.io/specification/2025-03-26/basic/authorization
- Use official Microsoft packages for lean implementation
- Implement streamable HTTP as per: https://modelcontextprotocol.io/specification/draft/basic/transports#streamable-http