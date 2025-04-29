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

