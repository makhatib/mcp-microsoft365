/**
 * Configuration Module
 * Loads and validates environment variables
 */

import { config as loadEnv } from 'dotenv';
import { z } from 'zod';

// Load .env file
loadEnv();

// Configuration schema
const configSchema = z.object({
  tenantId: z.string().min(1, 'TENANT_ID is required'),
  clientId: z.string().min(1, 'CLIENT_ID is required'),
  clientSecret: z.string().min(1, 'CLIENT_SECRET is required'),
  defaultUser: z.string().email('DEFAULT_USER must be a valid email').optional(),
  logLevel: z.enum(['debug', 'info', 'warn', 'error']).default('info'),
  debug: z.boolean().default(false),
});

export type Config = z.infer<typeof configSchema>;

function loadConfig(): Config {
  const rawConfig = {
    tenantId: process.env.TENANT_ID,
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    defaultUser: process.env.DEFAULT_USER,
    logLevel: process.env.LOG_LEVEL || 'info',
    debug: process.env.DEBUG === 'true',
  };

  const result = configSchema.safeParse(rawConfig);
  
  if (!result.success) {
    const errors = result.error.errors
      .map(e => `  - ${e.path.join('.')}: ${e.message}`)
      .join('\n');
    throw new Error(`Configuration error:\n${errors}`);
  }

  return result.data;
}

export const config = loadConfig();
