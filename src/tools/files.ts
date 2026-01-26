/**
 * OneDrive Files Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { DriveItem, DriveItemSummary } from '../types/files.js';

// === Schemas ===

const filesListSchema = z.object({
  user: z.string().optional(),
  path: z.string().default(''),
  top: z.number().min(1).max(100).default(20),
});

const filesGetSchema = z.object({
  user: z.string().optional(),
  itemId: z.string(),
});

const filesReadSchema = z.object({
  user: z.string().optional(),
  itemId: z.string(),
});

const filesSearchSchema = z.object({
  user: z.string().optional(),
  query: z.string(),
  top: z.number().min(1).max(100).default(10),
});

const filesDeleteSchema = z.object({
  user: z.string().optional(),
  itemId: z.string(),
});

// === Tool Definitions ===

export const filesTools: Tool[] = [
  {
    name: 'm365_files_list',
    description: 'List files and folders in OneDrive',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        path: { type: 'string', description: 'Folder path (empty for root)', default: '' },
        top: { type: 'number', description: 'Number of items', default: 20 },
      },
    },
  },
  {
    name: 'm365_files_get',
    description: 'Get file/folder metadata',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        itemId: { type: 'string', description: 'Item ID' },
      },
      required: ['itemId'],
    },
  },
  {
    name: 'm365_files_read',
    description: 'Read file content (text files only)',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        itemId: { type: 'string', description: 'File item ID' },
      },
      required: ['itemId'],
    },
  },
  {
    name: 'm365_files_search',
    description: 'Search for files in OneDrive',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        query: { type: 'string', description: 'Search query' },
        top: { type: 'number', description: 'Number of results', default: 10 },
      },
      required: ['query'],
    },
  },
  {
    name: 'm365_files_delete',
    description: 'Delete a file or folder',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        itemId: { type: 'string', description: 'Item ID to delete' },
      },
      required: ['itemId'],
    },
  },
];

// === Handlers ===

async function filesList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = filesListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_files_list', { path: input.path, top: input.top });

  const pathPart = input.path ? `:/${input.path}:` : '';
  const response = await graphClient.get<GraphListResponse<DriveItem>>(
    `/users/${user}/drive/root${pathPart}/children`,
    {
      $top: input.top,
      $select: 'id,name,size,folder,file,webUrl,lastModifiedDateTime',
    }
  );

  const result: DriveItemSummary[] = response.value.map(f => ({
    id: f.id,
    name: f.name,
    size: f.size,
    isFolder: !!f.folder,
    mimeType: f.file?.mimeType,
    webUrl: f.webUrl,
    modified: f.lastModifiedDateTime,
  }));

  logToolResult('m365_files_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function filesGet(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = filesGetSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_files_get', { itemId: input.itemId });

  const item = await graphClient.get<DriveItem>(
    `/users/${user}/drive/items/${input.itemId}`
  );

  logToolResult('m365_files_get', true, Date.now() - start);
  return JSON.stringify(item, null, 2);
}

async function filesRead(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = filesReadSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_files_read', { itemId: input.itemId });

  // Get file metadata with download URL
  const item = await graphClient.get<DriveItem>(
    `/users/${user}/drive/items/${input.itemId}`
  );

  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error('Cannot read this file - no download URL');
  }

  const content = await graphClient.download(downloadUrl);
  
  // Limit content size
  const truncated = content.substring(0, 50000);

  logToolResult('m365_files_read', true, Date.now() - start);
  return truncated;
}

async function filesSearch(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = filesSearchSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_files_search', { query: input.query });

  const response = await graphClient.get<GraphListResponse<DriveItem>>(
    `/users/${user}/drive/root/search(q='${encodeURIComponent(input.query)}')`,
    { $top: input.top }
  );

  const result = response.value.map(f => ({
    id: f.id,
    name: f.name,
    path: f.parentReference?.path,
    webUrl: f.webUrl,
  }));

  logToolResult('m365_files_search', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function filesDelete(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = filesDeleteSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_files_delete', { itemId: input.itemId });

  await graphClient.delete(`/users/${user}/drive/items/${input.itemId}`);

  logToolResult('m365_files_delete', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'File deleted' });
}

// === Export ===

export const filesHandlers: Record<string, ToolHandler> = {
  'm365_files_list': filesList,
  'm365_files_get': filesGet,
  'm365_files_read': filesRead,
  'm365_files_search': filesSearch,
  'm365_files_delete': filesDelete,
};
