/**
 * Logger Module
 * Structured logging with winston
 */

import winston from 'winston';
import { config } from './config.js';

const { combine, timestamp, json } = winston.format;

// Custom format for stderr (MCP servers use stdout for protocol)
const stderrFormat = combine(
  timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
  json()
);

// Create logger instance
export const logger = winston.createLogger({
  level: config.logLevel,
  format: stderrFormat,
  defaultMeta: { service: 'mcp-m365' },
  transports: [
    // Log to stderr (stdout is reserved for MCP protocol)
    new winston.transports.Stream({
      stream: process.stderr,
    }),
  ],
});

// Helper functions for structured logging
export function logToolCall(tool: string, args: Record<string, unknown>): void {
  logger.info('tool_call', { tool, args });
}

export function logToolResult(tool: string, success: boolean, duration: number): void {
  logger.info('tool_result', { tool, success, duration });
}

export function logError(context: string, error: Error, extra?: Record<string, unknown>): void {
  logger.error('error', {
    context,
    message: error.message,
    stack: config.debug ? error.stack : undefined,
    ...extra,
  });
}

export function logDebug(message: string, data?: Record<string, unknown>): void {
  logger.debug(message, data);
}
