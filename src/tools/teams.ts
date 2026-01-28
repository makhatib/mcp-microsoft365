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

const meetingsListSchema = z.object({
  user: z.string().optional(),
  top: z.number().min(1).max(50).default(20),
});

const meetingTranscriptsSchema = z.object({
  user: z.string().optional(),
  meetingId: z.string(),
});

const meetingTranscriptContentSchema = z.object({
  user: z.string().optional(),
  meetingId: z.string(),
  transcriptId: z.string(),
});

const callRecordsSchema = z.object({
  fromDateTime: z.string().optional(),
  toDateTime: z.string().optional(),
  top: z.number().min(1).max(100).default(20),
});

const meetingByJoinUrlSchema = z.object({
  user: z.string().optional(),
  joinWebUrl: z.string(),
});

const teamsCreateChatSchema = z.object({
  user: z.string().optional(),
  members: z.string(),
  topic: z.string().optional(),
  message: z.string().optional(),
});

const teamsCreateOnlineMeetingSchema = z.object({
  user: z.string().optional(),
  subject: z.string(),
  startDateTime: z.string(),
  endDateTime: z.string(),
  participants: z.string().optional(),
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
  {
    name: 'm365_teams_create_chat',
    description: 'Create a new Teams chat (1:1 or group)',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        members: { type: 'string', description: 'Member emails, comma-separated (include yourself for group chats)' },
        topic: { type: 'string', description: 'Chat topic (for group chats)' },
        message: { type: 'string', description: 'Optional first message to send' },
      },
      required: ['members'],
    },
  },
  {
    name: 'm365_teams_create_online_meeting',
    description: 'Create a Teams online meeting (standalone, not linked to calendar)',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email (organizer)' },
        subject: { type: 'string', description: 'Meeting subject' },
        startDateTime: { type: 'string', description: 'Start datetime (ISO format)' },
        endDateTime: { type: 'string', description: 'End datetime (ISO format)' },
        participants: { type: 'string', description: 'Participant emails, comma-separated' },
      },
      required: ['subject', 'startDateTime', 'endDateTime'],
    },
  },
  {
    name: 'm365_meetings_list',
    description: 'List online meetings (Teams meetings)',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        top: { type: 'number', description: 'Number of meetings', default: 20 },
      },
    },
  },
  {
    name: 'm365_meeting_transcripts',
    description: 'List available transcripts for a Teams meeting',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        meetingId: { type: 'string', description: 'Online meeting ID' },
      },
      required: ['meetingId'],
    },
  },
  {
    name: 'm365_meeting_transcript_content',
    description: 'Get the transcript content (text) of a Teams meeting',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        meetingId: { type: 'string', description: 'Online meeting ID' },
        transcriptId: { type: 'string', description: 'Transcript ID' },
      },
      required: ['meetingId', 'transcriptId'],
    },
  },
  {
    name: 'm365_call_records',
    description: 'List call records (Teams meetings and calls). Requires CallRecords.Read.All permission.',
    inputSchema: {
      type: 'object',
      properties: {
        fromDateTime: { type: 'string', description: 'Filter calls from this datetime (ISO format)' },
        toDateTime: { type: 'string', description: 'Filter calls to this datetime (ISO format)' },
        top: { type: 'number', description: 'Number of records', default: 20 },
      },
    },
  },
  {
    name: 'm365_meeting_by_join_url',
    description: 'Get online meeting details by JoinWebUrl. Use this to find the meeting ID for transcript access.',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        joinWebUrl: { type: 'string', description: 'Teams meeting JoinWebUrl' },
      },
      required: ['joinWebUrl'],
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

// Meeting types
interface OnlineMeeting {
  id: string;
  subject: string;
  startDateTime: string;
  endDateTime: string;
  joinWebUrl: string;
  participants?: {
    organizer?: { identity?: { user?: { displayName: string } } };
    attendees?: Array<{ identity?: { user?: { displayName: string } } }>;
  };
}

interface Transcript {
  id: string;
  createdDateTime: string;
  meetingId: string;
  meetingOrganizerId: string;
}

async function meetingsList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = meetingsListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_meetings_list', { top: input.top });

  // Try beta API for better support of Teams meetings
  const response = await graphClient.getBeta<GraphListResponse<OnlineMeeting>>(
    `/users/${user}/onlineMeetings`,
    { $top: input.top }
  );

  const result = response.value.map(m => ({
    id: m.id,
    subject: m.subject,
    start: m.startDateTime,
    end: m.endDateTime,
    joinUrl: m.joinWebUrl,
    organizer: m.participants?.organizer?.identity?.user?.displayName,
  }));

  logToolResult('m365_meetings_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function meetingTranscripts(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = meetingTranscriptsSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_meeting_transcripts', { meetingId: input.meetingId });

  // Requires OnlineMeetingTranscript.Read.All for app-only auth
  const response = await graphClient.get<GraphListResponse<Transcript>>(
    `/users/${user}/onlineMeetings/${input.meetingId}/transcripts`
  );

  const result = response.value.map(t => ({
    id: t.id,
    created: t.createdDateTime,
    meetingId: t.meetingId,
  }));

  logToolResult('m365_meeting_transcripts', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function meetingTranscriptContent(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = meetingTranscriptContentSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_meeting_transcript_content', { meetingId: input.meetingId, transcriptId: input.transcriptId });

  // Get transcript content as VTT format
  // Requires OnlineMeetingTranscript.Read.All for app-only auth
  const response = await graphClient.getText(
    `/users/${user}/onlineMeetings/${input.meetingId}/transcripts/${input.transcriptId}/content`,
    'text/vtt'
  );

  logToolResult('m365_meeting_transcript_content', true, Date.now() - start);
  
  // Parse VTT to extract just the text
  const lines = response.split('\n');
  const textOnly: string[] = [];
  
  for (const line of lines) {
    // Skip WEBVTT header, timestamps, and empty lines
    if (line.startsWith('WEBVTT') || line.match(/^\d{2}:\d{2}/) || line.trim() === '') {
      continue;
    }
    // Extract speaker and text from <v Speaker>text</v> format
    const match = line.match(/<v ([^>]+)>(.+)<\/v>/);
    if (match) {
      textOnly.push(`${match[1]}: ${match[2]}`);
    } else if (line.trim()) {
      textOnly.push(line.trim());
    }
  }

  return JSON.stringify({
    raw: response,
    text: textOnly.join('\n'),
  }, null, 2);
}

// === Call Records ===

interface CallRecord {
  id: string;
  type: string;
  modalities: string[];
  startDateTime: string;
  endDateTime: string;
  joinWebUrl?: string;
  organizer?: {
    user?: {
      id: string;
      displayName: string;
    };
  };
}

async function callRecords(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = callRecordsSchema.parse(args);
  
  logToolCall('m365_call_records', { fromDateTime: input.fromDateTime, toDateTime: input.toDateTime });

  const params: Record<string, string | number | boolean | undefined> = {};
  
  // Build filter if dates provided (required for call records API)
  const filters: string[] = [];
  if (input.fromDateTime) {
    filters.push(`startDateTime ge ${input.fromDateTime}`);
  }
  if (input.toDateTime) {
    filters.push(`startDateTime le ${input.toDateTime}`);
  }
  if (filters.length > 0) {
    params.$filter = filters.join(' and ');
  }

  const response = await graphClient.get<GraphListResponse<CallRecord>>(
    `/communications/callRecords`,
    params
  );

  const result = response.value.map(r => ({
    id: r.id,
    type: r.type,
    modalities: r.modalities,
    start: r.startDateTime,
    end: r.endDateTime,
    joinWebUrl: r.joinWebUrl,
    organizer: r.organizer?.user?.displayName,
  }));

  logToolResult('m365_call_records', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

// === Meeting by JoinWebUrl ===

async function meetingByJoinUrl(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = meetingByJoinUrlSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_meeting_by_join_url', { joinWebUrl: input.joinWebUrl.substring(0, 50) + '...' });

  // Use filter to find meeting by JoinWebUrl
  const response = await graphClient.get<GraphListResponse<OnlineMeeting>>(
    `/users/${user}/onlineMeetings`,
    { $filter: `JoinWebUrl eq '${input.joinWebUrl}'` }
  );

  if (response.value.length === 0) {
    return JSON.stringify({ error: 'Meeting not found with the provided JoinWebUrl' });
  }

  const meeting = response.value[0];
  const result = {
    id: meeting.id,
    subject: meeting.subject,
    start: meeting.startDateTime,
    end: meeting.endDateTime,
    joinUrl: meeting.joinWebUrl,
    organizer: meeting.participants?.organizer?.identity?.user?.displayName,
  };

  logToolResult('m365_meeting_by_join_url', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

// === Create Chat ===

async function teamsCreateChat(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = teamsCreateChatSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_teams_create_chat', { members: input.members });

  const memberEmails = input.members.split(',').map(e => e.trim());
  const chatType = memberEmails.length > 1 ? 'group' : 'oneOnOne';

  // Build members array with conversation member format
  const members = memberEmails.map(email => ({
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    roles: ['owner'],
    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${email}')`,
  }));

  // Add the requesting user as owner too
  if (!memberEmails.includes(user)) {
    members.push({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user}')`,
    });
  }

  const chatBody: Record<string, unknown> = {
    chatType,
    members,
  };
  if (input.topic && chatType === 'group') {
    chatBody.topic = input.topic;
  }

  const chat = await graphClient.post<Chat>('/chats', chatBody);

  // Send initial message if provided
  if (input.message && chat.id) {
    await graphClient.post(`/chats/${chat.id}/messages`, {
      body: { content: input.message },
    });
  }

  logToolResult('m365_teams_create_chat', true, Date.now() - start);
  return JSON.stringify({
    success: true,
    chatId: chat.id,
    chatType: chat.chatType,
    topic: chat.topic,
    messageSent: !!input.message,
  }, null, 2);
}

// === Create Online Meeting ===

async function teamsCreateOnlineMeeting(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = teamsCreateOnlineMeetingSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_teams_create_online_meeting', { subject: input.subject });

  const meetingBody: Record<string, unknown> = {
    subject: input.subject,
    startDateTime: input.startDateTime,
    endDateTime: input.endDateTime,
  };

  if (input.participants) {
    meetingBody.participants = {
      attendees: input.participants.split(',').map(email => ({
        identity: {
          user: { id: email.trim() },
        },
        upn: email.trim(),
      })),
    };
  }

  const meeting = await graphClient.post<OnlineMeeting>(
    `/users/${user}/onlineMeetings`,
    meetingBody
  );

  logToolResult('m365_teams_create_online_meeting', true, Date.now() - start);
  return JSON.stringify({
    success: true,
    id: meeting.id,
    subject: meeting.subject,
    start: meeting.startDateTime,
    end: meeting.endDateTime,
    joinUrl: meeting.joinWebUrl,
  }, null, 2);
}

// === Export ===

export const teamsHandlers: Record<string, ToolHandler> = {
  'm365_teams_chats': teamsChats,
  'm365_teams_messages': teamsMessages,
  'm365_teams_send': teamsSend,
  'm365_teams_create_chat': teamsCreateChat,
  'm365_teams_create_online_meeting': teamsCreateOnlineMeeting,
  'm365_meetings_list': meetingsList,
  'm365_meeting_transcripts': meetingTranscripts,
  'm365_meeting_transcript_content': meetingTranscriptContent,
  'm365_call_records': callRecords,
  'm365_meeting_by_join_url': meetingByJoinUrl,
};

