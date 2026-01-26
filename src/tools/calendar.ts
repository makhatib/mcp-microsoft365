/**
 * Calendar Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { CalendarEvent, CalendarEventSummary, ScheduleInformation } from '../types/calendar.js';

// === Schemas ===

const calendarListSchema = z.object({
  user: z.string().optional(),
  startDateTime: z.string().optional(),
  endDateTime: z.string().optional(),
  top: z.number().min(1).max(100).default(10),
});

const calendarGetSchema = z.object({
  user: z.string().optional(),
  eventId: z.string(),
});

const calendarCreateSchema = z.object({
  user: z.string().optional(),
  subject: z.string(),
  start: z.string(),
  end: z.string(),
  body: z.string().optional(),
  location: z.string().optional(),
  attendees: z.string().optional(),
  isOnline: z.boolean().default(false),
});

const calendarDeleteSchema = z.object({
  user: z.string().optional(),
  eventId: z.string(),
});

const calendarAvailabilitySchema = z.object({
  user: z.string().optional(),
  users: z.string(),
  startDateTime: z.string(),
  endDateTime: z.string(),
});

// === Tool Definitions ===

export const calendarTools: Tool[] = [
  {
    name: 'm365_calendar_list',
    description: 'List calendar events',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        startDateTime: { type: 'string', description: 'Start date (ISO format)' },
        endDateTime: { type: 'string', description: 'End date (ISO format)' },
        top: { type: 'number', description: 'Number of events', default: 10 },
      },
    },
  },
  {
    name: 'm365_calendar_get',
    description: 'Get a specific calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        eventId: { type: 'string', description: 'Event ID' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'm365_calendar_create',
    description: 'Create a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        subject: { type: 'string', description: 'Event subject' },
        start: { type: 'string', description: 'Start datetime (ISO format)' },
        end: { type: 'string', description: 'End datetime (ISO format)' },
        body: { type: 'string', description: 'Event description' },
        location: { type: 'string', description: 'Event location' },
        attendees: { type: 'string', description: 'Attendee emails, comma-separated' },
        isOnline: { type: 'boolean', description: 'Create Teams meeting', default: false },
      },
      required: ['subject', 'start', 'end'],
    },
  },
  {
    name: 'm365_calendar_delete',
    description: 'Delete a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        eventId: { type: 'string', description: 'Event ID to delete' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'm365_calendar_availability',
    description: 'Check user availability / free-busy status',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User making the request' },
        users: { type: 'string', description: 'User emails to check, comma-separated' },
        startDateTime: { type: 'string', description: 'Start datetime (ISO format)' },
        endDateTime: { type: 'string', description: 'End datetime (ISO format)' },
      },
      required: ['users', 'startDateTime', 'endDateTime'],
    },
  },
];

// === Handlers ===

async function calendarList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_calendar_list', { top: input.top });

  let response: GraphListResponse<CalendarEvent>;
  
  if (input.startDateTime && input.endDateTime) {
    response = await graphClient.get<GraphListResponse<CalendarEvent>>(
      `/users/${user}/calendarView`,
      {
        startDateTime: input.startDateTime,
        endDateTime: input.endDateTime,
        $top: input.top,
        $select: 'id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl',
        $orderby: 'start/dateTime',
      }
    );
  } else {
    response = await graphClient.get<GraphListResponse<CalendarEvent>>(
      `/users/${user}/events`,
      {
        $top: input.top,
        $select: 'id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl',
        $orderby: 'start/dateTime',
      }
    );
  }

  const result: CalendarEventSummary[] = response.value.map(e => ({
    id: e.id,
    subject: e.subject,
    start: e.start.dateTime,
    end: e.end.dateTime,
    location: e.location?.displayName,
    isOnline: e.isOnlineMeeting,
    meetingUrl: e.onlineMeetingUrl,
  }));

  logToolResult('m365_calendar_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function calendarGet(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarGetSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_calendar_get', { eventId: input.eventId });

  const event = await graphClient.get<CalendarEvent>(
    `/users/${user}/events/${input.eventId}`
  );

  logToolResult('m365_calendar_get', true, Date.now() - start);
  return JSON.stringify(event, null, 2);
}

async function calendarCreate(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarCreateSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_calendar_create', { subject: input.subject });

  const event: Record<string, unknown> = {
    subject: input.subject,
    start: { dateTime: input.start, timeZone: 'UTC' },
    end: { dateTime: input.end, timeZone: 'UTC' },
  };

  if (input.body) {
    event.body = { contentType: 'HTML', content: input.body };
  }
  if (input.location) {
    event.location = { displayName: input.location };
  }
  if (input.attendees) {
    event.attendees = input.attendees.split(',').map(email => ({
      emailAddress: { address: email.trim() },
      type: 'required',
    }));
  }
  if (input.isOnline) {
    event.isOnlineMeeting = true;
  }

  const created = await graphClient.post<CalendarEvent>(`/users/${user}/events`, event);

  const result = {
    id: created.id,
    subject: created.subject,
    webLink: created.webLink,
    onlineMeetingUrl: created.onlineMeetingUrl,
  };

  logToolResult('m365_calendar_create', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function calendarDelete(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarDeleteSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_calendar_delete', { eventId: input.eventId });

  await graphClient.delete(`/users/${user}/events/${input.eventId}`);

  logToolResult('m365_calendar_delete', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Event deleted' });
}

async function calendarAvailability(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarAvailabilitySchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_calendar_availability', { users: input.users });

  const schedules = input.users.split(',').map(u => u.trim());
  
  const response = await graphClient.post<{ value: ScheduleInformation[] }>(
    `/users/${user}/calendar/getSchedule`,
    {
      schedules,
      startTime: { dateTime: input.startDateTime, timeZone: 'UTC' },
      endTime: { dateTime: input.endDateTime, timeZone: 'UTC' },
    }
  );

  logToolResult('m365_calendar_availability', true, Date.now() - start);
  return JSON.stringify(response.value, null, 2);
}

// === Export ===

export const calendarHandlers: Record<string, ToolHandler> = {
  'm365_calendar_list': calendarList,
  'm365_calendar_get': calendarGet,
  'm365_calendar_create': calendarCreate,
  'm365_calendar_delete': calendarDelete,
  'm365_calendar_availability': calendarAvailability,
};
