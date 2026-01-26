/**
 * Tools Registry
 * Combines all tools and handlers
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { ToolHandler } from '../types/common.js';

import { mailTools, mailHandlers } from './mail.js';
import { calendarTools, calendarHandlers } from './calendar.js';
import { filesTools, filesHandlers } from './files.js';
import { tasksTools, tasksHandlers } from './tasks.js';
import { teamsTools, teamsHandlers } from './teams.js';
import { usersTools, usersHandlers } from './users.js';

// Combine all tools
export const allTools: Tool[] = [
  ...mailTools,
  ...calendarTools,
  ...filesTools,
  ...tasksTools,
  ...teamsTools,
  ...usersTools,
];

// Combine all handlers
export const allHandlers: Record<string, ToolHandler> = {
  ...mailHandlers,
  ...calendarHandlers,
  ...filesHandlers,
  ...tasksHandlers,
  ...teamsHandlers,
  ...usersHandlers,
};

// Get handler by tool name
export function getHandler(toolName: string): ToolHandler | undefined {
  return allHandlers[toolName];
}
