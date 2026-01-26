/**
 * Tasks (To-Do) Types
 */

export interface TaskList {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
}

export interface TodoTask {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred';
  importance: 'low' | 'normal' | 'high';
  body?: { contentType: string; content: string };
  dueDateTime?: { dateTime: string; timeZone: string };
  completedDateTime?: { dateTime: string; timeZone: string };
  createdDateTime: string;
  lastModifiedDateTime: string;
}

export interface TaskSummary {
  id: string;
  title: string;
  status: string;
  importance: string;
  dueDate?: string;
}
