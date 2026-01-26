#!/usr/bin/env node
/**
 * MCP Microsoft 365 Server v2
 * Built with Spec-Driven Development
 * 
 * @author Mahmoud Alkhatib <iam@malkhatib.com>
 * @license MIT
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { logger, logError } from './core/logger.js';
import { formatError } from './core/errors.js';
import { allTools, getHandler } from './tools/index.js';

// Create MCP Server
const server = new Server(
  {
    name: 'mcp-microsoft365',
    version: '2.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  logger.info('Listing tools', { count: allTools.length });
  return { tools: allTools };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;
  
  try {
    const handler = getHandler(name);
    
    if (!handler) {
      throw new Error(`Unknown tool: ${name}`);
    }

    const result = await handler(args as Record<string, unknown>);
    
    return {
      content: [{ type: 'text', text: result }],
    };
    
  } catch (error) {
    const err = error as Error;
    logError(`tool_${name}`, err);
    
    return {
      content: [{ type: 'text', text: formatError(err) }],
      isError: true,
    };
  }
});

// Main
async function main() {
  logger.info('Starting MCP Microsoft 365 Server v2', { tools: allTools.length });
  
  const transport = new StdioServerTransport();
  await server.connect(transport);
  
  logger.info('Server running on stdio');
}

main().catch((error) => {
  logError('startup', error as Error);
  process.exit(1);
});
