/**
 * Common Types
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';

export interface EmailAddress {
  address: string;
  name?: string;
}

export interface GraphListResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export type ToolHandler = (args: Record<string, unknown>) => Promise<string>;

export interface ToolModule {
  tools: Tool[];
  handlers: Record<string, ToolHandler>;
}
