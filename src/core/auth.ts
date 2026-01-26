/**
 * Authentication Module
 * Handles OAuth2 token management for Microsoft Graph API
 */

import { config } from './config.js';
import { AuthError } from './errors.js';
import { logger } from './logger.js';

const TOKEN_URL = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

// Token cache
interface TokenCache {
  accessToken: string;
  expiresAt: number;
}

let tokenCache: TokenCache | null = null;

/**
 * Get a valid access token, refreshing if necessary
 */
export async function getAccessToken(): Promise<string> {
  // Check if cached token is still valid (with 60s buffer)
  if (tokenCache && tokenCache.expiresAt > Date.now() + 60000) {
    logger.debug('Using cached token', { expiresIn: Math.floor((tokenCache.expiresAt - Date.now()) / 1000) });
    return tokenCache.accessToken;
  }

  logger.debug('Requesting new token');
  
  try {
    const response = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: config.clientId,
        client_secret: config.clientSecret,
        scope: GRAPH_SCOPE,
        grant_type: 'client_credentials',
      }),
    });

    const data = await response.json() as Record<string, unknown>;

    if (!response.ok || data.error) {
      const errorDesc = data.error_description || data.error || 'Unknown error';
      throw new AuthError(`Token request failed: ${errorDesc}`);
    }

    const accessToken = data.access_token as string;
    const expiresIn = data.expires_in as number;

    // Cache the token
    tokenCache = {
      accessToken,
      expiresAt: Date.now() + (expiresIn * 1000),
    };

    logger.debug('Token acquired', { expiresIn });
    return accessToken;

  } catch (error) {
    if (error instanceof AuthError) {
      throw error;
    }
    throw new AuthError(`Failed to authenticate: ${(error as Error).message}`);
  }
}

/**
 * Clear the token cache (useful for testing or force refresh)
 */
export function clearTokenCache(): void {
  tokenCache = null;
}
