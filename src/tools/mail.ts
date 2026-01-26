/**
 * Mail Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { MailMessage, MailMessageSummary } from '../types/mail.js';

// === Schemas ===

const mailListSchema = z.object({
  user: z.string().optional(),
  folder: z.string().default('inbox'),
  top: z.number().min(1).max(100).default(10),
  filter: z.string().optional(),
});

const mailReadSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
});

const mailSendSchema = z.object({
  user: z.string().optional(),
  to: z.string(),
  subject: z.string(),
  body: z.string(),
  cc: z.string().optional(),
});

const mailSearchSchema = z.object({
  user: z.string().optional(),
  query: z.string(),
  top: z.number().min(1).max(100).default(10),
});

const mailDeleteSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
});

// === Tool Definitions ===

export const mailTools: Tool[] = [
  {
    name: 'm365_mail_list',
    description: 'List emails from a user\'s mailbox',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        folder: { type: 'string', description: 'Folder name (inbox, sentitems, drafts)', default: 'inbox' },
        top: { type: 'number', description: 'Number of emails to return (1-100)', default: 10 },
        filter: { type: 'string', description: 'OData filter (e.g., "isRead eq false")' },
      },
    },
  },
  {
    name: 'm365_mail_read',
    description: 'Read a specific email by ID',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        messageId: { type: 'string', description: 'Message ID' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'm365_mail_send',
    description: 'Send an email',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'Sender email' },
        to: { type: 'string', description: 'Recipient email(s), comma-separated' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body (HTML supported)' },
        cc: { type: 'string', description: 'CC recipients, comma-separated' },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  {
    name: 'm365_mail_search',
    description: 'Search emails',
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
    name: 'm365_mail_delete',
    description: 'Delete an email',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        messageId: { type: 'string', description: 'Message ID to delete' },
      },
      required: ['messageId'],
    },
  },
];

// === Handlers ===

async function mailList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_mail_list', { folder: input.folder, top: input.top });

  const response = await graphClient.get<GraphListResponse<MailMessage>>(
    `/users/${user}/mailFolders/${input.folder}/messages`,
    {
      $top: input.top,
      $select: 'id,subject,from,receivedDateTime,isRead,bodyPreview,importance,hasAttachments',
      $orderby: 'receivedDateTime desc',
      $filter: input.filter,
    }
  );

  const result: MailMessageSummary[] = response.value.map(m => ({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address || 'unknown',
    received: m.receivedDateTime,
    isRead: m.isRead,
    preview: m.bodyPreview?.substring(0, 200),
  }));

  logToolResult('m365_mail_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function mailRead(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailReadSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_mail_read', { messageId: input.messageId });

  const message = await graphClient.get<MailMessage>(
    `/users/${user}/messages/${input.messageId}`,
    { $select: 'id,subject,from,toRecipients,ccRecipients,body,receivedDateTime,importance,hasAttachments' }
  );

  const result = {
    id: message.id,
    subject: message.subject,
    from: message.from?.emailAddress,
    to: message.toRecipients?.map(r => r.emailAddress),
    cc: message.ccRecipients?.map(r => r.emailAddress),
    received: message.receivedDateTime,
    importance: message.importance,
    body: message.body?.content,
  };

  logToolResult('m365_mail_read', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function mailSend(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailSendSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_mail_send', { to: input.to, subject: input.subject });

  const toRecipients = input.to.split(',').map(email => ({
    emailAddress: { address: email.trim() },
  }));

  const message: Record<string, unknown> = {
    subject: input.subject,
    body: { contentType: 'HTML', content: input.body },
    toRecipients,
  };

  if (input.cc) {
    message.ccRecipients = input.cc.split(',').map(email => ({
      emailAddress: { address: email.trim() },
    }));
  }

  await graphClient.post(`/users/${user}/sendMail`, { message });

  logToolResult('m365_mail_send', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Email sent successfully' });
}

async function mailSearch(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailSearchSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_mail_search', { query: input.query });

  const response = await graphClient.get<GraphListResponse<MailMessage>>(
    `/users/${user}/messages`,
    {
      $search: `"${input.query}"`,
      $top: input.top,
      $select: 'id,subject,from,receivedDateTime,bodyPreview',
    }
  );

  const result: MailMessageSummary[] = response.value.map(m => ({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address || 'unknown',
    received: m.receivedDateTime,
    isRead: true,
    preview: m.bodyPreview?.substring(0, 200),
  }));

  logToolResult('m365_mail_search', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function mailDelete(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailDeleteSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_mail_delete', { messageId: input.messageId });

  await graphClient.delete(`/users/${user}/messages/${input.messageId}`);

  logToolResult('m365_mail_delete', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Email deleted' });
}

// === Export ===

export const mailHandlers: Record<string, ToolHandler> = {
  'm365_mail_list': mailList,
  'm365_mail_read': mailRead,
  'm365_mail_send': mailSend,
  'm365_mail_search': mailSearch,
  'm365_mail_delete': mailDelete,
};
