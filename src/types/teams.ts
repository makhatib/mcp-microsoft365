/**
 * Teams Types
 */

export interface Chat {
  id: string;
  topic?: string;
  chatType: 'oneOnOne' | 'group' | 'meeting';
  createdDateTime: string;
  members?: ChatMember[];
}

export interface ChatMember {
  id: string;
  displayName: string;
  email?: string;
}

export interface ChatMessage {
  id: string;
  messageType: 'message' | 'chatEvent' | 'typing';
  createdDateTime: string;
  from?: {
    user?: { displayName: string; id: string };
  };
  body: { contentType: string; content: string };
}

export interface ChatSummary {
  id: string;
  topic?: string;
  chatType: string;
  members?: string[];
}

export interface MessageSummary {
  id: string;
  from?: string;
  content: string;
  sent: string;
}
