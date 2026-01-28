/**
 * Mail Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { MailMessage, MailMessageSummary, MailFolder } from '../types/mail.js';

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

const mailReplySchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
  body: z.string(),
  replyAll: z.boolean().default(false),
});

const mailForwardSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
  to: z.string(),
  comment: z.string().optional(),
});

const mailMoveSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
  destinationFolder: z.string(),
});

const mailFlagSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
  flagStatus: z.enum(['flagged', 'complete', 'notFlagged']).default('flagged'),
});

const mailMarkReadSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
  isRead: z.boolean().default(true),
});

const mailFoldersListSchema = z.object({
  user: z.string().optional(),
  top: z.number().min(1).max(100).default(25),
});

const mailCreateDraftSchema = z.object({
  user: z.string().optional(),
  to: z.string(),
  subject: z.string(),
  body: z.string(),
  cc: z.string().optional(),
});

const mailDeleteSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
});

const mailAttachmentsListSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
});

const mailAttachmentGetSchema = z.object({
  user: z.string().optional(),
  messageId: z.string(),
  attachmentId: z.string(),
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
    name: 'm365_mail_reply',
    description: 'Reply to an email by message ID',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        messageId: { type: 'string', description: 'ID of the message to reply to' },
        body: { type: 'string', description: 'Reply body (HTML supported)' },
        replyAll: { type: 'boolean', description: 'Reply to all recipients (default: false)' },
      },
      required: ['messageId', 'body'],
    },
  },
  {
    name: 'm365_mail_forward',
    description: 'Forward an email to another recipient',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        messageId: { type: 'string', description: 'ID of the message to forward' },
        to: { type: 'string', description: 'Recipient email(s), comma-separated' },
        comment: { type: 'string', description: 'Optional comment to include with forwarded message' },
      },
      required: ['messageId', 'to'],
    },
  },
  {
    name: 'm365_mail_move',
    description: 'Move an email to a different folder',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        messageId: { type: 'string', description: 'ID of the message to move' },
        destinationFolder: { type: 'string', description: 'Destination folder name or ID (e.g., archive, deleteditems, inbox, junkemail, drafts, sentitems)' },
      },
      required: ['messageId', 'destinationFolder'],
    },
  },
  {
    name: 'm365_mail_flag',
    description: 'Flag or unflag an email for follow-up',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        messageId: { type: 'string', description: 'ID of the message to flag' },
        flagStatus: { type: 'string', enum: ['flagged', 'complete', 'notFlagged'], description: 'Flag status', default: 'flagged' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'm365_mail_mark_read',
    description: 'Mark an email as read or unread',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        messageId: { type: 'string', description: 'ID of the message' },
        isRead: { type: 'boolean', description: 'true = mark as read, false = mark as unread', default: true },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'm365_mail_folders',
    description: 'List mail folders',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
        top: { type: 'number', description: 'Number of folders to return', default: 25 },
      },
    },
  },
  {
    name: 'm365_mail_create_draft',
    description: 'Create a draft email without sending',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (default: configured user)' },
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
  {
    name: 'm365_mail_attachments_list',
    description: 'List attachments for an email',
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
    name: 'm365_mail_attachment_get',
    description: 'Get attachment details and content (base64 for files)',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        messageId: { type: 'string', description: 'Message ID' },
        attachmentId: { type: 'string', description: 'Attachment ID' },
      },
      required: ['messageId', 'attachmentId'],
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

async function mailReply(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailReplySchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  const action = input.replyAll ? 'replyAll' : 'reply';

  logToolCall('m365_mail_reply', { messageId: input.messageId, replyAll: input.replyAll });

  await graphClient.post(
    `/users/${user}/messages/${input.messageId}/${action}`,
    { message: {}, comment: input.body }
  );

  logToolResult('m365_mail_reply', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Reply sent successfully' });
}

async function mailForward(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailForwardSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_forward', { messageId: input.messageId, to: input.to });

  const toRecipients = input.to.split(',').map(email => ({
    emailAddress: { address: email.trim() },
  }));

  await graphClient.post(
    `/users/${user}/messages/${input.messageId}/forward`,
    { toRecipients, comment: input.comment || '' }
  );

  logToolResult('m365_mail_forward', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Email forwarded successfully' });
}

async function mailMove(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailMoveSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_move', { messageId: input.messageId, destination: input.destinationFolder });

  // Map common folder names to well-known folder names
  const folderMap: Record<string, string> = {
    'inbox': 'inbox',
    'archive': 'archive',
    'trash': 'deleteditems',
    'deleted': 'deleteditems',
    'deleteditems': 'deleteditems',
    'junk': 'junkemail',
    'spam': 'junkemail',
    'junkemail': 'junkemail',
    'drafts': 'drafts',
    'sent': 'sentitems',
    'sentitems': 'sentitems',
  };

  const destinationId = folderMap[input.destinationFolder.toLowerCase()] || input.destinationFolder;

  await graphClient.post(
    `/users/${user}/messages/${input.messageId}/move`,
    { destinationId }
  );

  logToolResult('m365_mail_move', true, Date.now() - start);
  return JSON.stringify({ success: true, message: `Email moved to ${input.destinationFolder}` });
}

async function mailFlag(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailFlagSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_flag', { messageId: input.messageId, flagStatus: input.flagStatus });

  await graphClient.patch(
    `/users/${user}/messages/${input.messageId}`,
    { flag: { flagStatus: input.flagStatus } }
  );

  logToolResult('m365_mail_flag', true, Date.now() - start);
  return JSON.stringify({ success: true, message: `Email flag set to ${input.flagStatus}` });
}

async function mailMarkRead(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailMarkReadSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_mark_read', { messageId: input.messageId, isRead: input.isRead });

  await graphClient.patch(
    `/users/${user}/messages/${input.messageId}`,
    { isRead: input.isRead }
  );

  logToolResult('m365_mail_mark_read', true, Date.now() - start);
  return JSON.stringify({ success: true, message: `Email marked as ${input.isRead ? 'read' : 'unread'}` });
}

async function mailFoldersList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailFoldersListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_folders', { top: input.top });

  const response = await graphClient.get<GraphListResponse<MailFolder>>(
    `/users/${user}/mailFolders`,
    {
      $top: input.top,
      $select: 'id,displayName,totalItemCount,unreadItemCount',
    }
  );

  logToolResult('m365_mail_folders', true, Date.now() - start);
  return JSON.stringify(response.value, null, 2);
}

async function mailCreateDraft(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailCreateDraftSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_create_draft', { to: input.to, subject: input.subject });

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

  const draft = await graphClient.post<MailMessage>(`/users/${user}/messages`, message);

  logToolResult('m365_mail_create_draft', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Draft created', id: draft.id, subject: draft.subject });
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

interface Attachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
  contentBytes?: string;
}

async function mailAttachmentsList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailAttachmentsListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_attachments_list', { messageId: input.messageId });

  const response = await graphClient.get<GraphListResponse<Attachment>>(
    `/users/${user}/messages/${input.messageId}/attachments`,
    { $select: 'id,name,contentType,size,isInline' }
  );

  const result = response.value.map(a => ({
    id: a.id,
    name: a.name,
    contentType: a.contentType,
    size: a.size,
    isInline: a.isInline,
  }));

  logToolResult('m365_mail_attachments_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function mailAttachmentGet(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = mailAttachmentGetSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_mail_attachment_get', { messageId: input.messageId, attachmentId: input.attachmentId });

  const attachment = await graphClient.get<Attachment>(
    `/users/${user}/messages/${input.messageId}/attachments/${input.attachmentId}`
  );

  // Limit content size for large attachments
  const result: Record<string, unknown> = {
    id: attachment.id,
    name: attachment.name,
    contentType: attachment.contentType,
    size: attachment.size,
  };
  
  if (attachment.contentBytes) {
    // Include base64 content but truncate if very large (>1MB)
    if (attachment.contentBytes.length > 1_400_000) {
      result.contentBytes = attachment.contentBytes.substring(0, 1_400_000);
      result.truncated = true;
    } else {
      result.contentBytes = attachment.contentBytes;
    }
  }

  logToolResult('m365_mail_attachment_get', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

// === Export ===

export const mailHandlers: Record<string, ToolHandler> = {
  'm365_mail_list': mailList,
  'm365_mail_read': mailRead,
  'm365_mail_send': mailSend,
  'm365_mail_reply': mailReply,
  'm365_mail_forward': mailForward,
  'm365_mail_move': mailMove,
  'm365_mail_flag': mailFlag,
  'm365_mail_mark_read': mailMarkRead,
  'm365_mail_folders': mailFoldersList,
  'm365_mail_create_draft': mailCreateDraft,
  'm365_mail_search': mailSearch,
  'm365_mail_delete': mailDelete,
  'm365_mail_attachments_list': mailAttachmentsList,
  'm365_mail_attachment_get': mailAttachmentGet,
};
