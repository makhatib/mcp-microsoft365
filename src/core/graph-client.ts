/**
 * Microsoft Graph API Client
 * Handles all HTTP requests to Graph API
 */

import { getAccessToken } from './auth.js';
import { GraphApiError } from './errors.js';
import { logDebug } from './logger.js';
import { config } from './config.js';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

interface RequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  body?: unknown;
  params?: Record<string, string | number | boolean | undefined>;
}

interface GraphErrorResponse {
  error?: {
    code: string;
    message: string;
    innerError?: {
      'request-id'?: string;
    };
  };
}

/**
 * Build URL with query parameters
 */
function buildUrl(endpoint: string, params?: Record<string, string | number | boolean | undefined>): string {
  // Ensure endpoint starts with /
  const normalizedEndpoint = endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
  const fullUrl = `${GRAPH_BASE_URL}${normalizedEndpoint}`;
  const url = new URL(fullUrl);
  
  if (params) {
    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== '') {
        url.searchParams.append(key, String(value));
      }
    });
  }
  
  return url.toString();
}

/**
 * Make a request to Microsoft Graph API
 */
async function request<T>(endpoint: string, options: RequestOptions = {}): Promise<T> {
  const { method = 'GET', body, params } = options;
  const url = buildUrl(endpoint, params);
  const token = await getAccessToken();

  logDebug('Graph API request', { method, endpoint, params });

  const fetchOptions: RequestInit = {
    method,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
  };

  if (body && (method === 'POST' || method === 'PATCH')) {
    fetchOptions.body = JSON.stringify(body);
  }

  const startTime = Date.now();
  const response = await fetch(url, fetchOptions);
  const duration = Date.now() - startTime;

  // Handle no-content responses
  if (response.status === 204 || response.status === 202) {
    logDebug('Graph API response', { status: response.status, duration });
    return { success: true } as T;
  }

  const text = await response.text();
  
  if (!text) {
    logDebug('Graph API response (empty)', { status: response.status, duration });
    return { success: true } as T;
  }

  let data: T & GraphErrorResponse;
  try {
    data = JSON.parse(text);
  } catch {
    throw new GraphApiError(
      `Invalid JSON response: ${text.substring(0, 100)}`,
      response.status,
      'INVALID_JSON'
    );
  }

  // Handle error responses
  if (!response.ok || data.error) {
    const error = data.error || { code: 'UNKNOWN', message: 'Unknown error' };
    throw new GraphApiError(
      error.message,
      response.status,
      error.code,
      error.innerError?.['request-id']
    );
  }

  logDebug('Graph API response', { status: response.status, duration });
  return data;
}

/**
 * Graph Client with convenience methods
 */
export const graphClient = {
  /**
   * GET request
   */
  async get<T>(endpoint: string, params?: Record<string, string | number | boolean | undefined>): Promise<T> {
    return request<T>(endpoint, { method: 'GET', params });
  },

  /**
   * POST request
   */
  async post<T>(endpoint: string, body?: unknown, params?: Record<string, string | number | boolean | undefined>): Promise<T> {
    return request<T>(endpoint, { method: 'POST', body, params });
  },

  /**
   * PATCH request
   */
  async patch<T>(endpoint: string, body?: unknown): Promise<T> {
    return request<T>(endpoint, { method: 'PATCH', body });
  },

  /**
   * DELETE request
   */
  async delete<T>(endpoint: string): Promise<T> {
    return request<T>(endpoint, { method: 'DELETE' });
  },

  /**
   * Download file content
   */
  async download(downloadUrl: string): Promise<string> {
    const response = await fetch(downloadUrl);
    if (!response.ok) {
      throw new GraphApiError(
        'Failed to download file',
        response.status,
        'DOWNLOAD_FAILED'
      );
    }
    return response.text();
  },

  /**
   * Get default user from config
   */
  getDefaultUser(): string {
    if (!config.defaultUser) {
      throw new GraphApiError(
        'No user specified and DEFAULT_USER not configured',
        400,
        'NO_USER'
      );
    }
    return config.defaultUser;
  },
};
