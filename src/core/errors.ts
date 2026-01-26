/**
 * Custom Error Classes
 */

/**
 * Error from Microsoft Graph API
 */
export class GraphApiError extends Error {
  constructor(
    message: string,
    public readonly statusCode: number,
    public readonly code: string,
    public readonly requestId?: string
  ) {
    super(message);
    this.name = 'GraphApiError';
  }

  toJSON() {
    return {
      name: this.name,
      message: this.message,
      statusCode: this.statusCode,
      code: this.code,
      requestId: this.requestId,
    };
  }
}

/**
 * Authentication error
 */
export class AuthError extends Error {
  constructor(message: string, public readonly details?: string) {
    super(message);
    this.name = 'AuthError';
  }
}

/**
 * Input validation error
 */
export class ValidationError extends Error {
  constructor(
    message: string,
    public readonly field: string,
    public readonly value?: unknown
  ) {
    super(message);
    this.name = 'ValidationError';
  }
}

/**
 * Format error for MCP response
 */
export function formatError(error: Error): string {
  if (error instanceof GraphApiError) {
    return `Graph API Error (${error.statusCode}): ${error.message} [${error.code}]`;
  }
  if (error instanceof AuthError) {
    return `Authentication Error: ${error.message}`;
  }
  if (error instanceof ValidationError) {
    return `Validation Error: ${error.message} (field: ${error.field})`;
  }
  return `Error: ${error.message}`;
}
