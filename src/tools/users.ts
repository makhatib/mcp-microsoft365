/**
 * Users Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { User, UserSummary } from '../types/users.js';

// === Schemas ===

const usersListSchema = z.object({
  top: z.number().min(1).max(100).default(20),
  filter: z.string().optional(),
});

const usersGetSchema = z.object({
  user: z.string(),
});

const usersSearchSchema = z.object({
  query: z.string(),
  top: z.number().min(1).max(50).default(10),
});

// === Tool Definitions ===

export const usersTools: Tool[] = [
  {
    name: 'm365_users_list',
    description: 'List organization users',
    inputSchema: {
      type: 'object',
      properties: {
        top: { type: 'number', description: 'Number of users', default: 20 },
        filter: { type: 'string', description: 'OData filter' },
      },
    },
  },
  {
    name: 'm365_users_get',
    description: 'Get user profile information',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email or ID' },
      },
      required: ['user'],
    },
  },
  {
    name: 'm365_users_search',
    description: 'Search for users',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query (name or email)' },
        top: { type: 'number', description: 'Number of results', default: 10 },
      },
      required: ['query'],
    },
  },
];

// === Handlers ===

async function usersList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = usersListSchema.parse(args);
  
  logToolCall('m365_users_list', { top: input.top });

  const response = await graphClient.get<GraphListResponse<User>>(
    '/users',
    {
      $top: input.top,
      $select: 'id,displayName,mail,userPrincipalName,jobTitle',
      $filter: input.filter,
    }
  );

  const result: UserSummary[] = response.value.map(u => ({
    id: u.id,
    displayName: u.displayName,
    email: u.mail || u.userPrincipalName,
    jobTitle: u.jobTitle,
  }));

  logToolResult('m365_users_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function usersGet(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = usersGetSchema.parse(args);
  
  logToolCall('m365_users_get', { user: input.user });

  const user = await graphClient.get<User>(
    `/users/${input.user}`,
    { $select: 'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone' }
  );

  logToolResult('m365_users_get', true, Date.now() - start);
  return JSON.stringify(user, null, 2);
}

async function usersSearch(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = usersSearchSchema.parse(args);
  
  logToolCall('m365_users_search', { query: input.query });

  // Search by displayName or mail containing the query
  const filter = `startswith(displayName,'${input.query}') or startswith(mail,'${input.query}')`;
  
  const response = await graphClient.get<GraphListResponse<User>>(
    '/users',
    {
      $top: input.top,
      $filter: filter,
      $select: 'id,displayName,mail,userPrincipalName,jobTitle',
    }
  );

  const result: UserSummary[] = response.value.map(u => ({
    id: u.id,
    displayName: u.displayName,
    email: u.mail || u.userPrincipalName,
    jobTitle: u.jobTitle,
  }));

  logToolResult('m365_users_search', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

// === Export ===

export const usersHandlers: Record<string, ToolHandler> = {
  'm365_users_list': usersList,
  'm365_users_get': usersGet,
  'm365_users_search': usersSearch,
};
