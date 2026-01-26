/**
 * Tasks (To-Do) Tools
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { graphClient } from '../core/graph-client.js';
import { logToolCall, logToolResult } from '../core/logger.js';
import type { ToolHandler, GraphListResponse } from '../types/common.js';
import type { TaskList, TodoTask, TaskSummary } from '../types/tasks.js';

// === Schemas ===

const tasksListsSchema = z.object({
  user: z.string().optional(),
});

const tasksListSchema = z.object({
  user: z.string().optional(),
  listId: z.string(),
  top: z.number().min(1).max(100).default(20),
});

const tasksCreateSchema = z.object({
  user: z.string().optional(),
  listId: z.string(),
  title: z.string(),
  body: z.string().optional(),
  dueDateTime: z.string().optional(),
  importance: z.enum(['low', 'normal', 'high']).default('normal'),
});

const tasksCompleteSchema = z.object({
  user: z.string().optional(),
  listId: z.string(),
  taskId: z.string(),
});

const tasksDeleteSchema = z.object({
  user: z.string().optional(),
  listId: z.string(),
  taskId: z.string(),
});

// === Tool Definitions ===

export const tasksTools: Tool[] = [
  {
    name: 'm365_tasks_lists',
    description: 'List all task lists (To-Do)',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
      },
    },
  },
  {
    name: 'm365_tasks_list',
    description: 'List tasks in a task list',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        listId: { type: 'string', description: 'Task list ID' },
        top: { type: 'number', description: 'Number of tasks', default: 20 },
      },
      required: ['listId'],
    },
  },
  {
    name: 'm365_tasks_create',
    description: 'Create a new task',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        listId: { type: 'string', description: 'Task list ID' },
        title: { type: 'string', description: 'Task title' },
        body: { type: 'string', description: 'Task description' },
        dueDateTime: { type: 'string', description: 'Due date (ISO format)' },
        importance: { type: 'string', description: 'low, normal, or high', default: 'normal' },
      },
      required: ['listId', 'title'],
    },
  },
  {
    name: 'm365_tasks_complete',
    description: 'Mark a task as completed',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        listId: { type: 'string', description: 'Task list ID' },
        taskId: { type: 'string', description: 'Task ID' },
      },
      required: ['listId', 'taskId'],
    },
  },
  {
    name: 'm365_tasks_delete',
    description: 'Delete a task',
    inputSchema: {
      type: 'object',
      properties: {
        user: { type: 'string', description: 'User email' },
        listId: { type: 'string', description: 'Task list ID' },
        taskId: { type: 'string', description: 'Task ID' },
      },
      required: ['listId', 'taskId'],
    },
  },
];

// === Handlers ===

async function tasksLists(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = tasksListsSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_tasks_lists', {});

  const response = await graphClient.get<GraphListResponse<TaskList>>(
    `/users/${user}/todo/lists`
  );

  const result = response.value.map(l => ({
    id: l.id,
    displayName: l.displayName,
    isOwner: l.isOwner,
  }));

  logToolResult('m365_tasks_lists', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function tasksList(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = tasksListSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_tasks_list', { listId: input.listId });

  const listId = encodeURIComponent(input.listId);
  const response = await graphClient.get<GraphListResponse<TodoTask>>(
    `/users/${user}/todo/lists/${listId}/tasks`,
    { $top: input.top }
  );

  const result: TaskSummary[] = response.value.map(t => ({
    id: t.id,
    title: t.title,
    status: t.status,
    importance: t.importance,
    dueDate: t.dueDateTime?.dateTime,
  }));

  logToolResult('m365_tasks_list', true, Date.now() - start);
  return JSON.stringify(result, null, 2);
}

async function tasksCreate(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = tasksCreateSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_tasks_create', { title: input.title });

  const listId = encodeURIComponent(input.listId);
  const task: Record<string, unknown> = {
    title: input.title,
    importance: input.importance,
  };

  if (input.body) {
    task.body = { contentType: 'text', content: input.body };
  }
  if (input.dueDateTime) {
    task.dueDateTime = { dateTime: input.dueDateTime, timeZone: 'UTC' };
  }

  const created = await graphClient.post<TodoTask>(
    `/users/${user}/todo/lists/${listId}/tasks`,
    task
  );

  logToolResult('m365_tasks_create', true, Date.now() - start);
  return JSON.stringify({
    id: created.id,
    title: created.title,
    status: created.status,
  }, null, 2);
}

async function tasksComplete(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = tasksCompleteSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_tasks_complete', { taskId: input.taskId });

  const listId = encodeURIComponent(input.listId);
  const taskId = encodeURIComponent(input.taskId);

  await graphClient.patch(
    `/users/${user}/todo/lists/${listId}/tasks/${taskId}`,
    { status: 'completed' }
  );

  logToolResult('m365_tasks_complete', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Task marked as completed' });
}

async function tasksDelete(args: Record<string, unknown>): Promise<string> {
  const start = Date.now();
  const input = tasksDeleteSchema.parse(args);
  const user = input.user || graphClient.getDefaultUser();
  
  logToolCall('m365_tasks_delete', { taskId: input.taskId });

  const listId = encodeURIComponent(input.listId);
  const taskId = encodeURIComponent(input.taskId);

  await graphClient.delete(
    `/users/${user}/todo/lists/${listId}/tasks/${taskId}`
  );

  logToolResult('m365_tasks_delete', true, Date.now() - start);
  return JSON.stringify({ success: true, message: 'Task deleted' });
}

// === Export ===

export const tasksHandlers: Record<string, ToolHandler> = {
  'm365_tasks_lists': tasksLists,
  'm365_tasks_list': tasksList,
  'm365_tasks_create': tasksCreate,
  'm365_tasks_complete': tasksComplete,
  'm365_tasks_delete': tasksDelete,
};
