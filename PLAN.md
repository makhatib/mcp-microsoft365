# 🏗️ Technical Plan - MCP Microsoft 365 Server v2

## 📁 Project Structure

```
mcp-microsoft365-v2/
├── 📄 CONSTITUTION.md          # مبادئ المشروع
├── 📄 SPEC.md                  # المواصفات
├── 📄 PLAN.md                  # الخطة التقنية (هذا الملف)
├── 📄 TASKS.md                 # قائمة المهام
├── 📄 README.md                # التوثيق
├── 📄 package.json
├── 📄 tsconfig.json
├── 📄 .env.example
├── 📄 .gitignore
│
├── 📁 src/
│   ├── 📄 index.ts             # Entry point - MCP server setup
│   ├── 📄 server.ts            # Server configuration
│   │
│   ├── 📁 core/
│   │   ├── 📄 graph-client.ts  # Microsoft Graph API client
│   │   ├── 📄 auth.ts          # Authentication & token management
│   │   ├── 📄 logger.ts        # Logging utility
│   │   ├── 📄 errors.ts        # Custom error classes
│   │   └── 📄 config.ts        # Configuration management
│   │
│   ├── 📁 tools/
│   │   ├── 📄 index.ts         # Tool registry & exports
│   │   ├── 📄 mail.ts          # Mail tools (6 tools)
│   │   ├── 📄 calendar.ts      # Calendar tools (6 tools)
│   │   ├── 📄 files.ts         # OneDrive tools (6 tools)
│   │   ├── 📄 tasks.ts         # Tasks tools (7 tools)
│   │   ├── 📄 teams.ts         # Teams tools (3 tools)
│   │   └── 📄 users.ts         # Users tools (3 tools)
│   │
│   ├── 📁 types/
│   │   ├── 📄 index.ts         # Type exports
│   │   ├── 📄 mail.ts          # Mail types
│   │   ├── 📄 calendar.ts      # Calendar types
│   │   ├── 📄 files.ts         # Files types
│   │   ├── 📄 tasks.ts         # Tasks types
│   │   ├── 📄 teams.ts         # Teams types
│   │   ├── 📄 users.ts         # Users types
│   │   └── 📄 graph.ts         # Graph API common types
│   │
│   └── 📁 utils/
│       ├── 📄 validators.ts    # Input validation with Zod
│       └── 📄 formatters.ts    # Response formatters
│
└── 📁 tests/
    ├── 📄 setup.ts             # Test configuration
    ├── 📁 unit/
    │   ├── 📄 mail.test.ts
    │   ├── 📄 calendar.test.ts
    │   ├── 📄 files.test.ts
    │   ├── 📄 tasks.test.ts
    │   ├── 📄 teams.test.ts
    │   └── 📄 users.test.ts
    └── 📁 mocks/
        └── 📄 graph-api.ts     # Mocked responses
```

---

## 🔧 Technology Stack

| Layer | Technology | السبب |
|-------|------------|-------|
| Language | TypeScript 5.x | Type safety |
| Runtime | Node.js 20+ | LTS, native fetch |
| MCP SDK | @modelcontextprotocol/sdk | Official SDK |
| Validation | Zod | Runtime type checking |
| Logging | winston | Flexible, structured |
| Testing | Vitest | Fast, native ESM |
| Linting | ESLint + Prettier | Code quality |

---

## 🏛️ Architecture

### Component Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                      MCP Client (AI Agent)                  │
└─────────────────────────┬───────────────────────────────────┘
                          │ stdio
┌─────────────────────────▼───────────────────────────────────┐
│                      MCP Server                              │
│  ┌─────────────────────────────────────────────────────┐    │
│  │                   Tool Registry                      │    │
│  │  ┌─────┐ ┌────────┐ ┌─────┐ ┌─────┐ ┌─────┐ ┌─────┐│    │
│  │  │Mail │ │Calendar│ │Files│ │Tasks│ │Teams│ │Users││    │
│  │  └──┬──┘ └───┬────┘ └──┬──┘ └──┬──┘ └──┬──┘ └──┬──┘│    │
│  └─────┼────────┼─────────┼───────┼───────┼───────┼───┘    │
│        └────────┴─────────┴───────┴───────┴───────┘        │
│                           │                                  │
│  ┌────────────────────────▼────────────────────────────┐    │
│  │                  Graph Client                        │    │
│  │  ┌──────────┐  ┌───────────┐  ┌──────────────────┐  │    │
│  │  │   Auth   │  │  Request  │  │  Error Handler   │  │    │
│  │  │ Manager  │  │  Handler  │  │                  │  │    │
│  │  └──────────┘  └───────────┘  └──────────────────┘  │    │
│  └─────────────────────────────────────────────────────┘    │
└─────────────────────────┬───────────────────────────────────┘
                          │ HTTPS
┌─────────────────────────▼───────────────────────────────────┐
│              Microsoft Graph API                             │
└─────────────────────────────────────────────────────────────┘
```

### Data Flow

```
1. MCP Client sends tool request
         │
         ▼
2. Server routes to appropriate tool handler
         │
         ▼
3. Tool validates input (Zod)
         │
         ▼
4. Tool calls GraphClient
         │
         ▼
5. GraphClient checks/refreshes token
         │
         ▼
6. GraphClient makes API request
         │
         ▼
7. Response formatted & returned
```

---

## 🔐 Authentication Flow

```
┌─────────────────────────────────────────────────────────────┐
│                    Token Management                          │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│   ┌──────────────┐     ┌──────────────┐                     │
│   │ Check Cache  │────▶│ Token Valid? │                     │
│   └──────────────┘     └──────┬───────┘                     │
│                               │                              │
│              ┌────────────────┼────────────────┐            │
│              │ Yes            │ No             │            │
│              ▼                ▼                             │
│   ┌──────────────┐     ┌──────────────┐                     │
│   │ Return Token │     │ Refresh/New  │                     │
│   └──────────────┘     │    Token     │                     │
│                        └──────┬───────┘                     │
│                               │                              │
│                        ┌──────▼───────┐                     │
│                        │ Cache Token  │                     │
│                        └──────────────┘                     │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

---

## 📝 Tool Implementation Pattern

كل tool يتبع نفس النمط:

```typescript
// src/tools/mail.ts

import { z } from 'zod';
import { Tool, ToolHandler } from '../types';
import { graphClient } from '../core/graph-client';
import { logger } from '../core/logger';

// 1. Input Schema
const mailListSchema = z.object({
  user: z.string().email().optional(),
  folder: z.string().default('inbox'),
  top: z.number().min(1).max(100).default(10),
  filter: z.string().optional(),
});

// 2. Tool Definition
export const mailListTool: Tool = {
  name: 'm365_mail_list',
  description: 'List emails from a user\'s mailbox',
  inputSchema: {
    type: 'object',
    properties: {
      user: { type: 'string', description: 'User email' },
      folder: { type: 'string', description: 'Folder name' },
      top: { type: 'number', description: 'Number of emails' },
      filter: { type: 'string', description: 'OData filter' },
    },
  },
};

// 3. Handler Implementation
export const mailListHandler: ToolHandler = async (args) => {
  // Validate input
  const input = mailListSchema.parse(args);
  
  // Log request
  logger.info('mail_list', { folder: input.folder, top: input.top });
  
  // Make API call
  const response = await graphClient.get(
    `/users/${input.user}/mailFolders/${input.folder}/messages`,
    { $top: input.top, $filter: input.filter }
  );
  
  // Format & return
  return formatMailList(response.value);
};

// 4. Export
export const mailTools = {
  tools: [mailListTool, /* ... */],
  handlers: {
    'm365_mail_list': mailListHandler,
    /* ... */
  },
};
```

---

## 🧪 Testing Strategy

### Unit Tests
- Mock GraphClient
- Test each tool handler independently
- Test validation schemas

### Integration Tests (Optional)
- Real Graph API calls (with test tenant)
- End-to-end tool execution

### Test Example
```typescript
// tests/unit/mail.test.ts
import { describe, it, expect, vi } from 'vitest';
import { mailListHandler } from '../../src/tools/mail';
import { graphClient } from '../../src/core/graph-client';

vi.mock('../../src/core/graph-client');

describe('mailListHandler', () => {
  it('should return formatted email list', async () => {
    vi.mocked(graphClient.get).mockResolvedValue({
      value: [{ id: '1', subject: 'Test' }]
    });

    const result = await mailListHandler({ folder: 'inbox', top: 10 });
    
    expect(result).toHaveLength(1);
    expect(result[0].subject).toBe('Test');
  });
});
```

---

## ⚠️ Error Handling

```typescript
// src/core/errors.ts

export class GraphApiError extends Error {
  constructor(
    message: string,
    public statusCode: number,
    public code: string
  ) {
    super(message);
    this.name = 'GraphApiError';
  }
}

export class AuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AuthError';
  }
}

export class ValidationError extends Error {
  constructor(message: string, public field: string) {
    super(message);
    this.name = 'ValidationError';
  }
}
```

---

## 📊 Logging Format

```typescript
// Structured logging with winston
logger.info('tool_executed', {
  tool: 'm365_mail_list',
  duration: 234,
  user: 'user@domain.com',
  success: true
});

logger.error('graph_api_error', {
  tool: 'm365_mail_send',
  statusCode: 403,
  error: 'Insufficient permissions'
});
```

---

## 🚀 Deployment

### Environment Variables
```env
# Required
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
DEFAULT_USER=user@domain.com

# Optional
LOG_LEVEL=info
DEBUG=false
```

### Build & Run
```bash
npm install
npm run build
npm start
```

### mcporter Integration
```bash
mcporter config add m365 --stdio "node /path/to/dist/index.js"
```

---

*تم إنشاء هذا الملف وفق منهجية Spec-Kit*
*التاريخ: 2026-01-26*
