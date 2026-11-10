# Pattern: Database Connection + Security

Ang blueprint na ginamit sa **LCC Calendar Notebook**. Kopyahin ito sa bagong
project — Next.js (App Router) + Google Sheets bilang database, na pinapagitnaan
ng Apps Script web app.

---

## 1. Arkitektura

```
Browser (React)
   │  fetch("/api/...")        ← walang secret sa client
   ▼
Next.js Route Handler  (server only: env vars, session, validation)
   │  POST ?action=...&key=...
   ▼
Google Apps Script Web App  (API_KEY sa Script Properties)
   │
   ▼
Google Sheets  =  database (isa kang tab = isang table)
```

**Prinsipyo:** ang browser ay **hindi kailanman** kakausap sa Apps Script
nang direkta. Lahat ng request ay dumadaan sa Next.js server, kaya:

- ang `APPS_SCRIPT_SECRET` ay hindi lumalabas sa bundle,
- ang bawat action ay may sariling auth check,
- at ang validation ay nangyayari bago pa maabot ang database.

---

## 2. Database layer (Sheets as database)

### 2.1 Schema bilang constants

```js
const TABS = { USERS: 'Users', ENTRIES: 'Entries', PHOTOS: 'Photos', ADMIN_LOGS: 'AdminLogs' };

const SCHEMAS = {
  Users:      ['id','email','name','role','password_hash','created_at','is_active','updated_at'],
  Entries:    ['id','control_number','date', /* ... */ 'photo_urls','created_at','user_email','updated_at'],
  Photos:     ['id','entry_id','photo_url','photo_name','uploaded_at'],
  AdminLogs:  ['id','action','user_email','target_entry_id','timestamp'],
};
```

Isang lugar lang ang nagde-define ng columns. Ang `initializeDatabase()` ang
gumagawa ng kulang na tabs at headers — pwede mong i-run ulit nang walang
side effect.

### 2.2 Row → object conversion

```js
function getSheetData(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const timeZone = getTimeZone_(sheet);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = normalizeCellValue_(row[i], h, timeZone); });
    return obj;
  });
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); setupSheetHeaders(sheet, sheetName); }
  return sheet;
}
```

### 2.3 Request router

```js
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const action = e.parameter.action || e.parameters.action;
  const body = e.postData ? JSON.parse(e.postData.contents) : {};

  // public
  if (action === 'validateLogin') return validateLogin(body);

  // everything else needs the server-held key
  const auth = checkAuth(e.parameter);
  if (!auth.ok) return json({ ok: false, error: auth.error });

  switch (action) {
    case 'saveEntry': return saveEntry(body);
    // ...
    default: return json({ ok: false, error: 'Unknown action: ' + action });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### 2.4 Client-side proxy helper

```ts
// lib/appsScript.ts — server only
export async function callAppsScript(action: string, payload: Record<string, unknown> = {}) {
  const config = getAppsScriptConfig();           // env vars
  if (!config) throw new AppsScriptConfigError();

  const target = `${config.url}?action=${encodeURIComponent(action)}&key=${encodeURIComponent(config.key)}`;
  const res = await fetch(target, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    cache: "no-store",
  });
  if (!res.ok) throw new Error(`Apps Script request failed with status ${res.status}`);
  return (await res.json()) as Record<string, any>;
}
```

---

## 3. Security — pitong layers

### Layer 1 — Signed session cookie

Bakit: noong plain JSON ang cookie, pwede itong i-edit sa DevTools at gawing
`role: "admin"`. Ngayon: `base64url(payload).base64url(HMAC-SHA256)`.

```ts
// lib/session.ts — Web Crypto lang, kaya gumagana sa route handler AT middleware
export const SESSION_COOKIE = "lcc_session";

export async function signSession(payload: SessionPayload): Promise<string> {
  const body = toBase64Url(encoder.encode(JSON.stringify(payload)));
  const signature = await crypto.subtle.sign("HMAC", await getKey(), encoder.encode(body));
  return `${body}.${toBase64Url(new Uint8Array(signature))}`;
}

export async function verifySession(value: string | null): Promise<SessionPayload | null> {
  // verify signature → parse → check expiresAt → return payload o null
}
```

**Rules na huwag kalimutan:**
- `httpOnly: true` — hindi mababasa ng JS.
- `sameSite: "lax"`, `secure` sa production.
- `expiresAt` ay nasa loob ng signed payload, kaya hindi ma-extend ng client.

### Layer 2 — Route protection sa middleware

```ts
const PUBLIC_ROUTES = ["/login", "/terms", "/privacy", "/dataprotection", "/_next", "/icon"];
const ADMIN_ROUTES  = ["/admin"];

export async function middleware(request: NextRequest) {
  if (PUBLIC_ROUTES.some(r => pathname.startsWith(r))) return NextResponse.next();
  const session = await verifySession(request.cookies.get(SESSION_COOKIE)?.value);
  if (!session) return NextResponse.redirect(new URL("/login", request.url));
  if (ADMIN_ROUTES.some(r => pathname.startsWith(r)) && session.role !== "admin")
    return NextResponse.redirect(new URL("/admin/login", request.url));
  return NextResponse.next();
}
```

⚠️ Ang `matcher` ay nag-e-exclude ng `/api`, kaya **bawat API route ay may
sariling session check**. Huwag umasa sa middleware para sa API.

### Layer 3 — Action allowlists sa API route

Huwag hayaan ang browser na pumili ng kahit anong Apps Script action.

```ts
const USER_ACTIONS  = new Set(["saveEntry", "getEntries", "getEntriesByMonth"]);
const ADMIN_ACTIONS = new Set(["getAllEntries", "deleteEntry", "updateEntry", "listUsers"]);

if (!USER_ACTIONS.has(action) && !ADMIN_ACTIONS.has(action)) return 400;
if (ADMIN_ACTIONS.has(action) && session.role !== "admin") return 403;
```

### Layer 4 — Server-side validation, at huwag magtiwala sa client identity

```ts
const parsed = entryPayloadSchema.safeParse(payload);   // zod
if (!parsed.success) return NextResponse.json({ ok:false, error: parsed.error.issues[0].message }, { status: 400 });

// Ang email ay galing sa session, HINDI sa request body.
payload = { ...parsed.data, user_email: session.email };
```

Ganito rin sa admin actions: `payload = { ...payload, user_email: session.email }`
para totoo ang audit trail.

### Layer 5 — Login throttling

```ts
const MAX_ATTEMPTS = 8;
const WINDOW_MS = 10 * 60 * 1000;
const attempts = new Map<string, { count: number; resetAt: number }>();

if (isThrottled(key)) return NextResponse.json({ ok:false, error:"Too many attempts." }, { status: 429 });
```

Key = `IP + email`. In-memory ito (per instance, nawawala sa redeploy) — sapat
laban sa scripted guessing, pero kung multi-instance, ilipat sa Redis/DB.

### Layer 6 — Server-to-server secret

- Ang `API_KEY` ay nasa **Apps Script → Script Properties**.
- Ang `APPS_SCRIPT_SECRET` ay nasa **Vercel/hosting env**.
- Pareho ang value; wala sa browser, wala sa git.

```js
function checkAuth(params) {
  if (!API_KEY) return { ok: false, error: 'API key not configured' };
  if (String((params && params.key) || '') === API_KEY) return { ok: true };
  return { ok: false, error: 'Unauthorized' };
}
```

### Layer 7 — Password hashing + audit log

```js
const PASSWORD_HASH_ROUNDS = 500;
function hashPassword(password, salt) { /* salted SHA-256, iterated, 'sha256$salt$hex' */ }
```

Stored format: `sha256$<salt>$<hex>`. May legacy plaintext support para sa
migration — **alisin ito** kapag tapos na ang migration.

Bawat sensitive action: `logAudit('DELETE_ENTRY', user_email, targetId)`.

> Kung bago ang project: mas mainam ang proper password KDF (bcrypt/argon2) o
> hosted auth provider. Ang pattern na ito ay minimum na gumagana sa Sheets.

---

## 4. Config / env

| Pangalan | Saan naka-store | Nakikita ng browser? |
|---|---|---|
| `GOOGLE_APPS_SCRIPT_URL` | hosting env | ❌ (server only) |
| `APPS_SCRIPT_SECRET` | hosting env ↔ Apps Script `API_KEY` | ❌ |
| `SESSION_SECRET` | hosting env | ❌ |
| `TELEGRAM_BOT_TOKEN` / `TELEGRAM_CHAT_ID` | Apps Script Script Properties | ❌ |
| `DRIVE_FOLDER_ID` | Apps Script Script Properties | ❌ |

Wala kahit isang secret sa client bundle. I-verify: `grep -r "bot" .next/static`
dapat walang token.

---

## 5. Checklist para sa bagong project

1. **Sheet**: gumawa ng spreadsheet, tapusin ang `TABS` + `SCHEMAS`, i-run ang
   `initializeDatabase()`.
2. **Apps Script**: i-paste ang `Code.gs` → Deploy as Web app → *Execute as: Me*,
   *Who has access: Anyone* → i-set ang `API_KEY` sa Script Properties.
3. **Env**: `GOOGLE_APPS_SCRIPT_URL`, `APPS_SCRIPT_SECRET`, `SESSION_SECRET`.
4. **`lib/appsScript.ts`**: `callAppsScript()` + `AppsScriptConfigError`.
5. **`lib/session.ts`**: `signSession` / `verifySession` (Web Crypto, para gumana
   sa middleware at route handlers).
6. **`middleware.ts`**: `PUBLIC_ROUTES`, `ADMIN_ROUTES`, matcher na nag-e-exclude
   ng `/api`.
7. **`app/api/<resource>/route.ts`**: session check → action allowlist → role
   check → zod validate → `callAppsScript`.
8. **`app/api/login/route.ts`**: throttle + `signSession` + `httpOnly` cookie.
9. **Audit**: `logAudit()` sa bawat mutating action, email mula sa session.

---

## 6. Mga pitfalls na aktwal na na-encounter

| Problema | Dahilan | Solusyon |
|---|---|---|
| `output.setHeader is not a function` | walang CORS API ang `ContentService` | huwag mag-set ng header; server-only ang tawag, walang preflight |
| Date matching laging fail | Sheets ibinabalik ang `"2026-11-08"` bilang `Date` | `normalizeCellValue_` + `dateKey_` |
| Forged admin cookie | plain JSON cookie | HMAC-signed session |
| "Walang entries" na blangko | API nagre-return ng `{ok:false}`, hindi nag-throw | i-check ang `ok` bago mag-set ng state |
| Delete hindi nagpapaluwag ng Drive | Trash files still count sa quota | huwag magsulat sa Drive; o permanent delete kung kailangan |
| Telegram photo delete fail | bot `deleteMessage` limit: **<48 hours** lang | huwag mag-delete sa Telegram; archive na lang |
| Photos nawala pagkatapos i-delete ang entry | file_id lang ang hawak ng app | tanggapin bilang disenyo: Telegram chat = archive |

---

## 7. Kailan hindi na Sheets ang tamang database

Lumipat kapag umabot sa alinman sa mga ito:

- Sabay-sabay na writes (Sheets ay walang transaction/locking).
- Mahigit ~50k rows, o kailangan ng indexed queries at joins.
- Kailangan ng real-time o multiple writers.
- Mahigit ~5 min ang isang operation (Apps Script execution limit).

Sa puntong iyon: Postgres (Neon/Supabase) + isang tunay na API layer, pero
**panatilihin ang parehong layers 1–7** — session signing, allowlists, zod,
throttle, server-only secrets, audit log.

---

## 8. Reference files sa repo na ito

| File | Layer |
|---|---|
| `lib/session.ts` | 1 |
| `middleware.ts` | 2 |
| `app/api/apps-script/route.ts` | 3, 4 |
| `app/api/login/route.ts` | 5 |
| `scripts/apps-script/Code.gs` | 6, 7 + database |
| `lib/appsScript.ts` | server↔backend |
| `lib/schema.ts` | shared validation types |
