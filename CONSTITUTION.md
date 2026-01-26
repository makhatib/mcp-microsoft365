# 📜 Constitution - MCP Microsoft 365 Server v2

## 🎯 Mission
بناء MCP Server احترافي للتكامل مع Microsoft 365، يكون:
- **موثوق** - يعمل بدون أخطاء
- **قابل للصيانة** - كود نظيف ومنظم
- **قابل للتوسع** - سهل إضافة أدوات جديدة
- **موثق** - كل شي واضح ومشروح

---

## 📋 Core Principles (المبادئ الأساسية)

### 1. 🏗️ Modular Architecture
- كل مجموعة أدوات في ملف منفصل (mail, calendar, files, etc.)
- فصل المنطق عن التكوين
- استخدام dependency injection

### 2. 🛡️ Type Safety
- TypeScript strict mode
- تعريف types لكل API response
- لا `any` إلا للضرورة القصوى

### 3. 🔒 Security First
- لا hardcoded secrets
- Token caching مع expiry
- Input validation لكل tool

### 4. 📝 Comprehensive Logging
- Log كل request/response
- Error logging مع context
- Debug mode للتطوير

### 5. ✅ Testability
- Unit tests لكل tool
- Integration tests للـ Graph API
- Mocking للـ external calls

### 6. 📚 Documentation
- JSDoc لكل function
- README شامل
- أمثلة استخدام

---

## 🚫 Anti-Patterns (ممنوعات)

- ❌ ملف واحد لكل الكود
- ❌ تجاهل الأخطاء (silent failures)
- ❌ secrets في الكود
- ❌ any types
- ❌ كود بدون tests
- ❌ magic strings/numbers

---

## 📊 Quality Standards

| المعيار | الهدف |
|---------|-------|
| Test Coverage | > 80% |
| TypeScript Strict | ✅ Enabled |
| ESLint Errors | 0 |
| Documentation | كل public function |
| Error Handling | كل external call |

---

## 🎨 Code Style

- **Naming**: camelCase للمتغيرات، PascalCase للـ types
- **Files**: kebab-case (e.g., `graph-client.ts`)
- **Exports**: Named exports فقط (لا default)
- **Comments**: بالإنجليزية للكود، بالعربية للتوثيق العام

---

## 👥 Target Users

1. **Clawdbot/AI Agents** - استخدام آلي للأدوات
2. **Developers** - تكامل مع تطبيقاتهم
3. **Power Users** - استخدام عبر mcporter CLI

---

*تم إنشاء هذا الملف وفق منهجية Spec-Kit*
*التاريخ: 2026-01-26*
