# MCP Microsoft 365 Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.5-blue.svg)](https://www.typescriptlang.org/)
[![MCP](https://img.shields.io/badge/MCP-1.0-green.svg)](https://modelcontextprotocol.io/)

An MCP (Model Context Protocol) server for Microsoft 365 integration. Built with **Spec-Driven Development** methodology.

## ✨ Features

### 📧 Mail (Outlook)
- `m365_mail_list` - List emails from inbox/folders
- `m365_mail_read` - Read specific email
- `m365_mail_send` - Send emails (HTML supported)
- `m365_mail_reply` - Reply to an email
- `m365_mail_forward` - Forward an email
- `m365_mail_move` - Move email to folder
- `m365_mail_flag` - Flag/unflag for follow-up
- `m365_mail_mark_read` - Mark as read/unread
- `m365_mail_folders` - List mail folders
- `m365_mail_create_draft` - Create draft email
- `m365_mail_search` - Search emails
- `m365_mail_delete` - Delete emails
- `m365_mail_attachments_list` - List email attachments
- `m365_mail_attachment_get` - Get attachment content

### 📅 Calendar
- `m365_calendar_list` - List events
- `m365_calendar_get` - Get event details
- `m365_calendar_create` - Create events (Teams meeting support)
- `m365_calendar_update` - Update events
- `m365_calendar_respond` - Accept/decline invitations
- `m365_calendar_delete` - Delete events
- `m365_calendar_availability` - Check free/busy status
- `m365_calendar_find_meeting_times` - Find available meeting slots

### 📁 OneDrive
- `m365_files_list` - List files/folders
- `m365_files_get` - Get file metadata
- `m365_files_read` - Read file content
- `m365_files_search` - Search files
- `m365_files_upload` - Upload files
- `m365_files_delete` - Delete files
- `m365_files_share` - Create sharing link
- `m365_files_create_folder` - Create new folder
- `m365_files_move` - Move file/folder
- `m365_files_copy` - Copy file/folder

### ✅ Tasks (To-Do)
- `m365_tasks_lists` - List task lists
- `m365_tasks_list` - List tasks
- `m365_tasks_create` - Create tasks
- `m365_tasks_update` - Update tasks
- `m365_tasks_complete` - Mark task complete
- `m365_tasks_delete` - Delete tasks
- `m365_tasks_create_list` - Create new task list

### 💬 Teams
- `m365_teams_chats` - List chats
- `m365_teams_messages` - Read chat messages
- `m365_teams_send` - Send messages
- `m365_teams_create_chat` - Create new chat
- `m365_teams_create_online_meeting` - Create Teams meeting
- `m365_meetings_list` - List online meetings
- `m365_meeting_transcripts` - List meeting transcripts
- `m365_meeting_transcript_content` - Get transcript content
- `m365_call_records` - List call records
- `m365_meeting_by_join_url` - Find meeting by URL

### 📇 Contacts
- `m365_contacts_list` - List contacts
- `m365_contacts_get` - Get contact details
- `m365_contacts_create` - Create contact
- `m365_contacts_update` - Update contact
- `m365_contacts_search` - Search contacts
- `m365_contacts_delete` - Delete contact

### 👥 Users
- `m365_users_list` - List organization users
- `m365_users_get` - Get user profile
- `m365_users_search` - Search users

## 🚀 Quick Start

### 1. Azure Entra ID Setup

Create an app in [Azure Portal](https://portal.azure.com):

1. Go to **Microsoft Entra ID** → **App registrations** → **New registration**
2. Add **Application permissions** for Microsoft Graph:
   - `Mail.Read`, `Mail.Send`, `Mail.ReadWrite`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read.All`, `Files.ReadWrite.All`
   - `Sites.Read.All`
   - `Tasks.Read.All`, `Tasks.ReadWrite.All`
   - `Chat.Read.All`, `Chat.ReadWrite.All`
   - `Contacts.Read`, `Contacts.ReadWrite`
   - `User.Read.All`
   - `OnlineMeetings.Read.All`, `OnlineMeetings.ReadWrite.All` (for Teams meetings)
   - `OnlineMeetingTranscript.Read.All` (for transcripts)
3. **Grant admin consent** for your organization

### 2. Installation

```bash
git clone https://github.com/malkhatib/mcp-microsoft365.git
cd mcp-microsoft365
npm install
```

### 3. Configuration

Create a `.env` file:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
DEFAULT_USER=user@yourdomain.com
```

### 4. Build & Run

```bash
npm run build
npm start
```

### 5. Use with mcporter

```bash
# Add to mcporter config
mcporter config add m365 --stdio "node /path/to/dist/index.js"

# Test
mcporter call m365.m365_mail_list top:5
```

## 📖 Usage Examples

```bash
# List recent emails
mcporter call m365.m365_mail_list top:10

# Send an email
mcporter call m365.m365_mail_send \
  to:"someone@example.com" \
  subject:"Hello" \
  body:"<p>Hi there!</p>"

# List calendar events
mcporter call m365.m365_calendar_list top:5

# Create event with Teams meeting
mcporter call m365.m365_calendar_create \
  subject:"Meeting" \
  start:"2024-01-30T10:00:00" \
  end:"2024-01-30T11:00:00" \
  isOnline:true

# Search files
mcporter call m365.m365_files_search query:"report" top:10

# Create a task
mcporter call m365.m365_tasks_create \
  listId:"xxx" \
  title:"Review document" \
  importance:"high"
```

## 🏗️ Project Structure

```
mcp-microsoft365/
├── src/
│   ├── index.ts          # Entry point
│   ├── core/
│   │   ├── auth.ts       # OAuth2 token management
│   │   ├── config.ts     # Environment configuration
│   │   ├── errors.ts     # Custom error classes
│   │   ├── graph-client.ts # Microsoft Graph API client
│   │   └── logger.ts     # Structured logging
│   ├── tools/
│   │   ├── mail.ts       # Mail tools
│   │   ├── calendar.ts   # Calendar tools
│   │   ├── files.ts      # OneDrive tools
│   │   ├── tasks.ts      # Tasks tools
│   │   ├── teams.ts      # Teams tools
│   │   └── users.ts      # Users tools
│   └── types/            # TypeScript interfaces
├── CONSTITUTION.md       # Project principles
├── SPEC.md              # Specifications
├── PLAN.md              # Technical plan
└── TASKS.md             # Task breakdown
```

## 🔧 Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Watch mode
npm run dev

# Run tests
npm test

# Lint
npm run lint
```

## 📋 Built with Spec-Driven Development

This project follows the [Spec-Kit](https://github.com/github/spec-kit) methodology:

1. **Constitution** - Project principles and standards
2. **Specification** - What we're building and why
3. **Plan** - Technical architecture
4. **Tasks** - Implementation breakdown
5. **Implementation** - Clean, modular code

## 🔒 Security

- No hardcoded secrets
- Environment variables for configuration
- Token caching with automatic refresh
- Input validation with Zod

## 📄 License

MIT License - see [LICENSE](LICENSE) file.

## 👤 Author

**Mahmoud Alkhatib**
- Website: [malkhatib.com](https://malkhatib.com)
- YouTube: [@malkhatib](https://youtube.com/@malkhatib)

## 🙏 Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Microsoft Graph API](https://docs.microsoft.com/graph/)
- [Spec-Kit by GitHub](https://github.com/github/spec-kit)
