/**
 * Calendar Types
 */

import type { EmailAddress } from './common.js';

export interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface Attendee {
  emailAddress: EmailAddress;
  type: 'required' | 'optional' | 'resource';
  status?: {
    response: 'none' | 'organizer' | 'accepted' | 'declined' | 'tentativelyAccepted';
  };
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: DateTimeTimeZone;
  end: DateTimeTimeZone;
  location?: { displayName: string };
  body?: { contentType: string; content: string };
  attendees?: Attendee[];
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  webLink?: string;
  organizer?: { emailAddress: EmailAddress };
}

export interface CalendarEventSummary {
  id: string;
  subject: string;
  start: string;
  end: string;
  location?: string;
  isOnline: boolean;
  meetingUrl?: string;
}

export interface ScheduleItem {
  status: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown';
  start: DateTimeTimeZone;
  end: DateTimeTimeZone;
}

export interface ScheduleInformation {
  scheduleId: string;
  availabilityView: string;
  scheduleItems: ScheduleItem[];
}
