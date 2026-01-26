/**
 * Mail Types
 */

import type { EmailAddress } from './common.js';

export interface MailMessage {
  id: string;
  subject: string;
  from: { emailAddress: EmailAddress };
  toRecipients: { emailAddress: EmailAddress }[];
  ccRecipients?: { emailAddress: EmailAddress }[];
  body: { contentType: 'html' | 'text'; content: string };
  bodyPreview?: string;
  receivedDateTime: string;
  sentDateTime?: string;
  isRead: boolean;
  importance: 'low' | 'normal' | 'high';
  hasAttachments: boolean;
}

export interface MailFolder {
  id: string;
  displayName: string;
  totalItemCount: number;
  unreadItemCount: number;
}

export interface MailMessageSummary {
  id: string;
  subject: string;
  from: string;
  received: string;
  isRead: boolean;
  preview?: string;
}
