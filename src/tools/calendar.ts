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

const calendarUpdateSchema = z.object({
  user: z.string().optional(),
  eventId: z.string(),
  subject: z.string().optional(),
  start: z.string().optional(),
  end: z.string().optional(),
  body: z.string().optional(),
  location: z.string().optional(),
  attendees: z.string().optional(),
  isOnline: z.boolean().optional(),
});

const calendarRespondSchema = z.object({
  user: z.string().optional(),
  eventId: z.string(),
  response: z.enum(['accept', 'tentativelyAccept', 'decline']),
  comment: z.string().optional(),
  sendResponse: z.boolean().default(true),
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

const calendarFindMeetingTimesSchema = z.object({
  user: z.string().optional(),
  attendees: z.string(),
  durationMinutes: z.number().default(60),
  startDateTime: z.string(),
  endDateTime: z.string(),
  maxCandidates: z.number().default(5),
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
    description: 'Create a calendar event with optional Teams meeting link and attendees',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        subject: { type: 'string', description: 'Event subject' },
        start: { type: 'string', description: 'Start datetime (ISO format)' },
        end: { type: 'string', description: 'End datetime (ISO format)' },
        body: { type: 'string', description: 'Event description (HTML supported)' },
        location: { type: 'string', description: 'Event location' },
        attendees: { type: 'string', description: 'Attendee emails, comma-separated' },
        isOnline: { type: 'boolean', description: 'Create Teams meeting link (default: false)', default: false },
      },
      required: ['subject', 'start', 'end'],
    },
  },
  {
    name: 'm365_calendar_update',
    description: 'Update an existing calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        eventId: { type: 'string', description: 'Event ID to update' },
        subject: { type: 'string', description: 'New event subject' },
        start: { type: 'string', description: 'New start datetime (ISO format)' },
        end: { type: 'string', description: 'New end datetime (ISO format)' },
        body: { type: 'string', description: 'New event description (HTML supported)' },
        location: { type: 'string', description: 'New location' },
        attendees: { type: 'string', description: 'Updated attendee emails, comma-separated' },
        isOnline: { type: 'boolean', description: 'Enable/disable Teams meeting' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'm365_calendar_respond',
    description: 'Accept, tentatively accept, or decline a calendar event/meeting invitation',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        eventId: { type: 'string', description: 'Event ID to respond to' },
        response: { type: 'string', enum: ['accept', 'tentativelyAccept', 'decline'], description: 'Response type' },
        comment: { type: 'string', description: 'Optional comment with response' },
        sendResponse: { type: 'boolean', description: 'Send response to organizer (default: true)', default: true },
      },
      required: ['eventId', 'response'],
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
  {
    name: 'm365_calendar_find_meeting_times',
    description: 'Find available meeting times that work for all attendees',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        attendees: { type: 'string', description: 'Attendee emails, comma-separated' },
        durationMinutes: { type: 'number', description: 'Meeting duration in minutes (default: 60)', default: 60 },
        startDateTime: { type: 'string', description: 'Search window start (ISO format)' },
        endDateTime: { type: 'string', description: 'Search window end (ISO format)' },
        maxCandidates: { type: 'number', description: 'Max number of suggestions (default: 5)', default: 5 },
      },
      required: ['attendees', 'startDateTime', 'endDateTime'],
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
    event.onlineMeetingProvider = 'teamsForBusiness';
  }

  const created = await graphClient.post<CalendarEvent>(`/users/${user}/events`, event);

  const result: Record<string, unknown> = {
    id: created.id,
    subject: created.subject,
    start: created.start,
    end: created.end,
    webLink: created.webLink,
  };
  if (created.onlineMeetingUrl) {
    result.onlineMeetingUrl = created.onlineMeetingUrl;
  }

  logToolResult('m365_calendar_create', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function calendarUpdate(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarUpdateSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_calendar_update', { eventId: input.eventId });

  const updates: Record<string, unknown> = {};

  if (input.subject) updates.subject = input.subject;
  if (input.start) updates.start = { dateTime: input.start, timeZone: 'UTC' };
  if (input.end) updates.end = { dateTime: input.end, timeZone: 'UTC' };
  if (input.body) updates.body = { contentType: 'HTML', content: input.body };
  if (input.location) updates.location = { displayName: input.location };
  if (input.attendees) {
    updates.attendees = input.attendees.split(',').map(email => ({
      emailAddress: { address: email.trim() },
      type: 'required',
    }));
  }
  if (input.isOnline !== undefined) {
    updates.isOnlineMeeting = input.isOnline;
    if (input.isOnline) updates.onlineMeetingProvider = 'teamsForBusiness';
  }

  const updated = await graphClient.patch<CalendarEvent>(
    `/users/${user}/events/${input.eventId}`,
    updates
  );

  const result: Record<string, unknown> = {
    id: updated.id,
    subject: updated.subject,
    start: updated.start,
    end: updated.end,
  };
  if (updated.onlineMeetingUrl) {
    result.onlineMeetingUrl = updated.onlineMeetingUrl;
  }

  logToolResult('m365_calendar_update', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function calendarRespond(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarRespondSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_calendar_respond', { eventId: input.eventId, response: input.response });

  await graphClient.post(
    `/users/${user}/events/${input.eventId}/${input.response}`,
    {
      comment: input.comment || '',
      sendResponse: input.sendResponse,
    }
  );

  logToolResult('m365_calendar_respond', true, Date.now() - start);
  return JSON.stringify({ success: true, message: `Event ${input.response}ed successfully` });
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

async function calendarFindMeetingTimes(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = calendarFindMeetingTimesSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();

  logToolCall('m365_calendar_find_meeting_times', { attendees: input.attendees });

  const attendees = input.attendees.split(',').map(email => ({
    emailAddress: { address: email.trim() },
    type: 'required' as const,
  }));

  const response = await graphClient.post<{
    meetingTimeSuggestions: Array<{
      confidence: number;
      meetingTimeSlot: { start: { dateTime: string; timeZone: string }; end: { dateTime: string; timeZone: string } };
      attendeeAvailability: Array<{ attendee: { emailAddress: { address: string } }; availability: string }>;
    }>;
  }>(
    `/users/${user}/findMeetingTimes`,
    {
      attendees,
      timeConstraint: {
        timeslots: [{
          start: { dateTime: input.startDateTime, timeZone: 'UTC' },
          end: { dateTime: input.endDateTime, timeZone: 'UTC' },
        }],
      },
      meetingDuration: `PT${input.durationMinutes}M`,
      maxCandidates: input.maxCandidates,
    }
  );

  const suggestions = response.meetingTimeSuggestions?.map(s => ({
    confidence: s.confidence,
    start: s.meetingTimeSlot.start.dateTime,
    end: s.meetingTimeSlot.end.dateTime,
    attendeeAvailability: s.attendeeAvailability?.map(a => ({
      email: a.attendee.emailAddress.address,
      availability: a.availability,
    })),
  })) || [];

  logToolResult('m365_calendar_find_meeting_times', true, Date.now() - start);
  return JSON.stringify(suggestions, null, 2);
}

// === Export ===

export const calendarHandlers: Record<string, ToolHandler> = {
  'm365_calendar_list': calendarList,
  'm365_calendar_get': calendarGet,
  'm365_calendar_create': calendarCreate,
  'm365_calendar_update': calendarUpdate,
  'm365_calendar_respond': calendarRespond,
  'm365_calendar_delete': calendarDelete,
  'm365_calendar_availability': calendarAvailability,
  'm365_calendar_find_meeting_times': calendarFindMeetingTimes,
};
