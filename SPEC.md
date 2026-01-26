# 📝 Specification - MCP Microsoft 365 Server v2

## 🎯 Overview

### What are we building?
MCP (Model Context Protocol) Server يتيح للـ AI agents التفاعل مع Microsoft 365 APIs.

### Why?
- تمكين AI من قراءة/إرسال الإيميلات
- إدارة التقويم والمواعيد
- الوصول لملفات OneDrive
- إدارة المهام (To-Do)
- التفاعل مع Teams

### For whom?
- Clawdbot و AI agents أخرى
- المطورين الذين يريدون تكامل M365

---

## 🔧 Functional Requirements

### 📧 FR1: Mail (Outlook)

| Tool | الوظيفة | المدخلات | المخرجات |
|------|---------|----------|----------|
| `m365_mail_list` | عرض الإيميلات | user?, folder?, top?, filter? | قائمة إيميلات |
| `m365_mail_read` | قراءة إيميل | messageId | تفاصيل الإيميل كاملة |
| `m365_mail_send` | إرسال إيميل | to, subject, body, cc? | تأكيد الإرسال |
| `m365_mail_search` | البحث | query, top? | نتائج البحث |
| `m365_mail_delete` | حذف إيميل | messageId | تأكيد الحذف |
| `m365_mail_move` | نقل إيميل | messageId, folderId | تأكيد النقل |

### 📅 FR2: Calendar

| Tool | الوظيفة | المدخلات | المخرجات |
|------|---------|----------|----------|
| `m365_calendar_list` | عرض المواعيد | user?, start?, end?, top? | قائمة المواعيد |
| `m365_calendar_get` | تفاصيل موعد | eventId | تفاصيل كاملة |
| `m365_calendar_create` | إنشاء موعد | subject, start, end, attendees?, isOnline? | الموعد الجديد |
| `m365_calendar_update` | تحديث موعد | eventId, updates | الموعد المحدث |
| `m365_calendar_delete` | حذف موعد | eventId | تأكيد الحذف |
| `m365_calendar_availability` | التوفر | users, start, end | جدول التوفر |

### 📁 FR3: OneDrive Files

| Tool | الوظيفة | المدخلات | المخرجات |
|------|---------|----------|----------|
| `m365_files_list` | عرض الملفات | path?, top? | قائمة الملفات |
| `m365_files_get` | تفاصيل ملف | itemId | metadata |
| `m365_files_read` | قراءة محتوى | itemId | المحتوى |
| `m365_files_search` | البحث | query, top? | نتائج البحث |
| `m365_files_upload` | رفع ملف | path, content | الملف الجديد |
| `m365_files_delete` | حذف ملف | itemId | تأكيد الحذف |

### ✅ FR4: Tasks (To-Do)

| Tool | الوظيفة | المدخلات | المخرجات |
|------|---------|----------|----------|
| `m365_tasks_lists` | قوائم المهام | user? | القوائم |
| `m365_tasks_list` | مهام قائمة | listId, top? | المهام |
| `m365_tasks_get` | تفاصيل مهمة | listId, taskId | التفاصيل |
| `m365_tasks_create` | إنشاء مهمة | listId, title, body?, due?, importance? | المهمة الجديدة |
| `m365_tasks_update` | تحديث مهمة | listId, taskId, updates | المهمة المحدثة |
| `m365_tasks_complete` | إكمال مهمة | listId, taskId | تأكيد |
| `m365_tasks_delete` | حذف مهمة | listId, taskId | تأكيد |

### 💬 FR5: Teams

| Tool | الوظيفة | المدخلات | المخرجات |
|------|---------|----------|----------|
| `m365_teams_chats` | قائمة المحادثات | top? | المحادثات |
| `m365_teams_messages` | رسائل محادثة | chatId, top? | الرسائل |
| `m365_teams_send` | إرسال رسالة | chatId, message | تأكيد |

### 👥 FR6: Users

| Tool | الوظيفة | المدخلات | المخرجات |
|------|---------|----------|----------|
| `m365_users_list` | قائمة المستخدمين | top?, filter? | المستخدمين |
| `m365_users_get` | تفاصيل مستخدم | userId | الملف الشخصي |
| `m365_users_search` | البحث عن مستخدم | query | النتائج |

---

## ⚙️ Non-Functional Requirements

### NFR1: Performance
- Response time < 2 seconds للعمليات العادية
- Token caching لتقليل auth requests
- Connection pooling للـ HTTP

### NFR2: Reliability
- Retry logic للـ transient failures
- Graceful error handling
- Meaningful error messages

### NFR3: Security
- Environment variables للـ secrets
- Token refresh قبل انتهاء الصلاحية
- Input sanitization

### NFR4: Maintainability
- Modular code structure
- Comprehensive logging
- TypeScript strict mode

### NFR5: Testability
- Unit tests لكل tool
- Mocked Graph API للـ tests
- CI/CD integration ready

---

## 📊 Data Models

### MailMessage
```typescript
interface MailMessage {
  id: string;
  subject: string;
  from: EmailAddress;
  to: EmailAddress[];
  cc?: EmailAddress[];
  body: { contentType: 'html' | 'text'; content: string };
  receivedDateTime: string;
  isRead: boolean;
  importance: 'low' | 'normal' | 'high';
  hasAttachments: boolean;
}
```

### CalendarEvent
```typescript
interface CalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName: string };
  attendees?: Attendee[];
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  body?: { contentType: string; content: string };
}
```

### DriveItem
```typescript
interface DriveItem {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  lastModifiedDateTime: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  parentReference?: { path: string };
}
```

### TodoTask
```typescript
interface TodoTask {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed';
  importance: 'low' | 'normal' | 'high';
  dueDateTime?: { dateTime: string; timeZone: string };
  body?: { contentType: string; content: string };
  completedDateTime?: { dateTime: string; timeZone: string };
}
```

---

## 🔌 External Dependencies

| Dependency | الغرض | الإصدار |
|------------|-------|---------|
| @modelcontextprotocol/sdk | MCP Server SDK | ^1.0.0 |
| dotenv | Environment config | ^16.0.0 |
| zod | Input validation | ^3.22.0 |
| winston | Logging | ^3.11.0 |

---

## 📈 Success Metrics

| المقياس | الهدف |
|---------|-------|
| عدد الأدوات | 28 tool |
| Test coverage | > 80% |
| Documentation | 100% |
| Error rate | < 1% |
| Avg response time | < 1.5s |

---

*تم إنشاء هذا الملف وفق منهجية Spec-Kit*
*التاريخ: 2026-01-26*
