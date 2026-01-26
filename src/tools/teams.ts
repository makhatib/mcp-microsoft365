/**
 * Teams Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { Chat, ChatMessage, ChatSummary, MessageSummary } from '../types/teams.js';

// === Schemas ===

const teamsChatsSchema = z.object({
  user: z.string().optional(),
  top: z.number().min(1).max(50).default(20),
});

const teamsMessagesSchema = z.object({
  user: z.string().optional(),
  chatId: z.string(),
  top: z.number().min(1).max(50).default(20),
});

const teamsSendSchema = z.object({
  user: z.string().optional(),
  chatId: z.string(),
  message: z.string(),
});

// === Tool Definitions ===

export const teamsTools: Tool[] = [
  {
    name: 'm365_teams_chats',
    description: 'List user\'s Teams chats',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        top: { type: 'number', description: 'Number of chats', default: 20 },
      },
    },
  },
  {
    name: 'm365_teams_messages',
    description: 'Get messages from a Teams chat',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        chatId: { type: 'string', description: 'Chat ID' },
        top: { type: 'number', description: 'Number of messages', default: 20 },
      },
      required: ['chatId'],
    },
  },
  {
    name: 'm365_teams_send',
    description: 'Send a message to a Teams chat',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        chatId: { type: 'string', description: 'Chat ID' },
        message: { type: 'string', description: 'Message content' },
      },
      required: ['chatId', 'message'],
    },
  },
];

// === Handlers ===

async function teamsChats(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = teamsChatsSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_teams_chats', { top: input.top });

  const response = await graphClient.get<GraphListResponse<Chat>>(
    `/users/${user}/chats`,
    {
      $top: input.top,
      $expand: 'members',
    }
  );

  const result: ChatSummary[] = response.value.map(c => ({
    id: c.id,
    topic: c.topic,
    chatType: c.chatType,
    members: c.members?.map(m => m.displayName),
  }));

  logToolResult('m365_teams_chats', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function teamsMessages(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = teamsMessagesSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_teams_messages', { chatId: input.chatId });

  const response = await graphClient.get<GraphListResponse<ChatMessage>>(
    `/users/${user}/chats/${input.chatId}/messages`,
    { $top: input.top }
  );

  const result: MessageSummary[] = response.value.map(m => ({
    id: m.id,
    from: m.from?.user?.displayName,
    content: m.body?.content?.replace(/<[^>]+>/g, '').substring(0, 500),
    sent: m.createdDateTime,
  }));

  logToolResult('m365_teams_messages', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function teamsSend(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = teamsSendSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_teams_send', { chatId: input.chatId });

  const response = await graphClient.post<{ id: string }>(
    `/users/${user}/chats/${input.chatId}/messages`,
    { body: { content: input.message } }
  );

  logToolResult('m365_teams_send', true, Date.now() - start);
  return JSON.stringify({ success: true, id: response.id });
}

// === Export ===

export const teamsHandlers: Record<string, ToolHandler> = {
  'm365_teams_chats': teamsChats,
  'm365_teams_messages': teamsMessages,
  'm365_teams_send': teamsSend,
};
