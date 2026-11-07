/**
 * ZeeMoApp demo back end (generated: do not edit mock-api.js by hand).
 *
 * Runs the real Apps Script back-end code in the browser against in-memory
 * stand-ins for Sheets, Cache, Drive and friends, so the demo behaves exactly
 * like a deployment. Used only when there is no apiUrl and no google.script.run.
 * Edit tools/demo_runtime.js, then run: python3 tools/build_preview.py
 */

const STORE_KEY = 'zm.demo.v2';

/* ---------- SHA-256 and HMAC (synchronous, as Apps Script's Utilities are) ---------- */
const K256 = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
]);

function sha256(bytes) {
  const len = bytes.length;
  const padded = new Uint8Array(((len + 9 + 63) >> 6) << 6);
  padded.set(bytes);
  padded[len] = 0x80;
  const bits = len * 8;
  const dv = new DataView(padded.buffer);
  dv.setUint32(padded.length - 8, Math.floor(bits / 0x100000000));
  dv.setUint32(padded.length - 4, bits >>> 0);
  const H = new Uint32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19]);
  const W = new Uint32Array(64);
  const rotr = (x, n) => (x >>> n) | (x << (32 - n));
  for (let off = 0; off < padded.length; off += 64) {
    for (let i = 0; i < 16; i++) W[i] = dv.getUint32(off + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(W[i - 15], 7) ^ rotr(W[i - 15], 18) ^ (W[i - 15] >>> 3);
      const s1 = rotr(W[i - 2], 17) ^ rotr(W[i - 2], 19) ^ (W[i - 2] >>> 10);
      W[i] = (W[i - 16] + s0 + W[i - 7] + s1) >>> 0;
    }
    let [a, b, c, d, e, f, g, h] = H;
    for (let i = 0; i < 64; i++) {
      const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const t1 = (h + S1 + ((e & f) ^ (~e & g)) + K256[i] + W[i]) >>> 0;
      const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) >>> 0;
      h = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    H[0] += a; H[1] += b; H[2] += c; H[3] += d; H[4] += e; H[5] += f; H[6] += g; H[7] += h;
  }
  const out = new Uint8Array(32);
  const odv = new DataView(out.buffer);
  for (let i = 0; i < 8; i++) odv.setUint32(i * 4, H[i]);
  return out;
}

function hmacSha256(key, msg) {
  let k = key.length > 64 ? sha256(key) : key;
  const block = new Uint8Array(64);
  block.set(k);
  const inner = new Uint8Array(64 + msg.length);
  const outer = new Uint8Array(64 + 32);
  for (let i = 0; i < 64; i++) { inner[i] = block[i] ^ 0x36; outer[i] = block[i] ^ 0x5c; }
  inner.set(msg, 64);
  outer.set(sha256(inner), 64);
  return sha256(outer);
}

const utf8 = s => new TextEncoder().encode(String(s));
const signed = u8 => Array.from(u8, b => (b > 127 ? b - 256 : b));

/* ---------- Sheets ---------- */
class DemoRange {
  constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const src = this.sh.data[this.r - 1 + i] || [];
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(src[this.c - 1 + j] ?? '');
      out.push(row);
    }
    return out;
  }
  setValues(v) {
    v.forEach((row, i) => row.forEach((val, j) => {
      const ri = this.r - 1 + i;
      while (this.sh.data.length <= ri) this.sh.data.push([]);
      this.sh.data[ri][this.c - 1 + j] = val;
    }));
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  getRow() { return this.r; }
  createTextFinder(text) {
    const self = this;
    return {
      matchEntireCell() { return this; },
      matchCase() { return this; },
      findAll() {
        const hits = [];
        for (let i = 0; i < self.nr; i++) {
          const src = self.sh.data[self.r - 1 + i] || [];
          for (let j = 0; j < self.nc; j++) {
            if (String(src[self.c - 1 + j] ?? '') === text) hits.push(new DemoRange(self.sh, self.r + i, self.c + j, 1, 1));
          }
        }
        return hits;
      },
      findNext() { return this.findAll()[0] || null; }
    };
  }
}

class DemoSheet {
  constructor(name, data, maxCols) { this.name = name; this.data = data || []; this.maxCols = maxCols || 26; }
  getName() { return this.name; }
  getLastRow() {
    let n = this.data.length;
    while (n && !(this.data[n - 1] || []).some(v => v !== '' && v !== undefined && v !== null)) n--;
    return n;
  }
  getLastColumn() { return Math.max(0, ...this.data.map(r => (r || []).length)); }
  getMaxRows() { return 1000; }
  getMaxColumns() { return this.maxCols; }
  getRange(r, c, nr = 1, nc = 1) { return new DemoRange(this, r, c, nr, nc); }
  clear() { this.data = []; }
  appendRow(row) { this.data.length = this.getLastRow(); this.data.push(row.slice()); }
  setFrozenRows() {}
  deleteColumns(start) { this.maxCols = start - 1; }
  insertColumnsAfter(after, n) { this.maxCols += n; }
}

class DemoSpreadsheet {
  constructor(sheets) { this.sheets = sheets || [new DemoSheet('Sheet1')]; }
  getId() { return 'demo-db'; }
  getUrl() { return 'demo'; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n, idx) { const s = new DemoSheet(n); if (idx === 0) this.sheets.unshift(s); else this.sheets.push(s); return s; }
  getSheets() { return this.sheets; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  setSpreadsheetTimeZone() {}
}

/* ---------- The real back end, wrapped ---------- */
function loadBackend(SpreadsheetApp, PropertiesService, CacheService, LockService, Utilities, DriveApp,
  MailApp, UrlFetchApp, ContentService, HtmlService, ScriptApp, Session, Date) {
/* ===== Config.gs ===== */
/**
 * ZeeMoApp back end: configuration and schema.
 *
 * Secrets and environment values live in Script Properties (Project Settings),
 * never in code:
 *   SPREADSHEET_ID     set by setup()
 *   MEDIA_FOLDER_ID    set by setup()
 *   OTP_PEPPER         set by setup()
 *   SIGNIN_METHOD      "phone" (default), "email" or "both": how people join
 *   TERMII_API_KEY     needed when phone sign-in is on
 *   TERMII_BASE_URL    optional, defaults to https://api.ng.termii.com
 *   TERMII_SENDER_ID   optional, defaults to CONFIG.SMS_SENDER_ID
 *   DEV_MODE           "true" returns OTP codes in the API response and skips SMS
 *
 * Security settings for the platform owner (setup() adds them with these defaults):
 *   REQUIRE_PASSCODE            "true": every user must set a passcode before chatting
 *   LOCK_AFTER_MINUTES          idle minutes before the app locks; "0" turns idle lock off
 *   MINIMISE_ACTION             "lock" (close chat and lock), "close" (close chat only) or "none"
 *   DELETE_FOR_EVERYONE_HOURS   how long senders can delete a message for everyone; "0" = no limit
 *   SHOW_LAST_SEEN              "false" hides everyone's last-seen time from other users
 *   PASSCODE_MIN_LENGTH         shortest passcode allowed, 4 to 6 digits; shorter existing ones must be changed
 *   LOCK_ON_BLUR                "true": on computers, lock as soon as the app window loses focus
 *
 * Housekeeping settings (setup() adds them):
 *   BACKUP_KEEP                 weekly database copies to keep in the Backups folder
 *   MEDIA_RETENTION_DAYS        delete chat photos and voice notes older than this; "0" keeps them
 *
 * To block someone, type TRUE in their "blocked" cell in the Users sheet.
 *
 * Shown on the Terms, Privacy and Delete-account pages (setup() adds them):
 *   OWNER_NAME                  who runs this service, e.g. "Zee Publishers"
 *   CONTACT_EMAIL               where people write about privacy, reports and deletion
 */

const CONFIG = Object.freeze({
  APP_NAME: 'ZeeMoApp',

  OTP_LENGTH: 6,
  OTP_TTL_SEC: 300,            // code validity
  OTP_MAX_ATTEMPTS: 5,         // wrong guesses before the code is burnt
  OTP_MAX_REQUESTS: 3,         // codes per phone number per window
  OTP_WINDOW_SEC: 900,

  SESSION_TTL_DAYS: 30,
  SESSION_CACHE_SEC: 21600,    // 6 h, the CacheService maximum

  MAX_MESSAGE_LENGTH: 4000,
  MAX_GROUP_MEMBERS: 256,
  RECENT_CACHE_SIZE: 50,       // messages kept per conversation in CacheService
  CACHE_VALUE_LIMIT: 90000,    // stay under the 100 KB CacheService value cap
  SYNC_MAX_MESSAGES: 200,
  HISTORY_MAX_MONTHS: 6,
  LAST_SEEN_WRITE_SEC: 60,     // throttle Users.lastSeen writes

  LOCK_WAIT_MS: 10000,

  MEDIA_MAX_BYTES: 5 * 1024 * 1024,
  MEDIA_MIME_TYPES: [
    'image/jpeg', 'image/png', 'image/webp',
    'audio/webm', 'audio/ogg', 'audio/mpeg', 'audio/mp4'
  ],

  SIGNIN_METHODS: ['phone', 'email'],
  DEFAULT_SIGNIN_METHOD: 'phone',

  SECURITY_DEFAULTS: {
    REQUIRE_PASSCODE: 'true',
    LOCK_AFTER_MINUTES: '5',
    MINIMISE_ACTION: 'lock',
    DELETE_FOR_EVERYONE_HOURS: '48',
    SHOW_LAST_SEEN: 'true',
    PASSCODE_MIN_LENGTH: '6',
    LOCK_ON_BLUR: 'false'
  },
  MINIMISE_ACTIONS: ['lock', 'close', 'none'],
  OWNER_DEFAULTS: {
    BACKUP_KEEP: '4',
    MEDIA_RETENTION_DAYS: '0',
    OWNER_NAME: '',
    CONTACT_EMAIL: ''
  },
  TERMS_VERSION: '1',          // raise when the terms change: everyone is asked to accept again
  REPORT_REASONS: ['spam', 'harassment', 'hate', 'sexual', 'scam', 'other'],
  REPORTS_PER_DAY: 20,
  DELETION_GRACE_DAYS: 7,     // web deletion requests without the passcode wait this long, so the owner of the account can cancel
  AVATAR_MAX_BYTES: 400 * 1024,
  AVATAR_MIME_TYPES: ['image/jpeg', 'image/png', 'image/webp'],
  BLOCK_CACHE_SEC: 120,
  PASSCODE_PATTERN: /^[0-9]{4,6}$/,
  PASSCODE_MAX_ATTEMPTS: 5,

  // Disappearing-message timers offered in each chat, in seconds. 0 is off.
  DISAPPEAR_OPTIONS: [0, 86400, 604800, 7776000],

  SMS_SENDER_ID: 'ZeeMo',
  DEFAULT_COUNTRY_CODE: '234',
  MAX_REQUEST_CHARS: 8 * 1024 * 1024
});

/** Column A of every sheet is its unique key. Never delete or sort rows. */
const SHEETS = Object.freeze({
  USERS: {
    name: 'Users',
    headers: ['userId', 'phone', 'displayName', 'avatarFileId', 'about', 'createdAt', 'lastSeen', 'email', 'passcode', 'blocked', 'lastSeenPrivacy',
              'termsVersion', 'termsAcceptedAt', 'deletedAt', 'deletionRequestedAt']
  },
  SESSIONS: {
    name: 'Sessions',
    headers: ['tokenHash', 'userId', 'createdAt', 'expiresAt', 'device', 'revoked']
  },
  CONVERSATIONS: {
    name: 'Conversations',
    headers: ['convId', 'type', 'title', 'directKey', 'createdBy', 'createdAt',
              'lastMsgId', 'lastMsgAt', 'lastMsgSenderId', 'lastMsgPreview', 'disappearSec', 'lastMsgExpiresAt', 'avatarFileId', 'lastMsgType']
  },
  PARTICIPANTS: {
    name: 'Participants',
    headers: ['key', 'convId', 'userId', 'role', 'joinedAt', 'lastReadMsgId', 'lastDeliveredMsgId', 'leftAt']
  },
  MEDIA: {
    name: 'Media',
    // convId holds the chat, or "avatar:<userId>" / "group:<convId>" for profile and group photos
    headers: ['fileId', 'convId', 'uploaderId', 'mimeType', 'bytes', 'createdAt', 'trashedAt']
  },
  BLOCKS: {
    name: 'Blocks',
    headers: ['key', 'blockerId', 'blockedId', 'createdAt', 'active']
  },
  REPORTS: {
    // Owner: type REMOVE, BLOCK or DISMISS in "action"; applyReportActions (hourly) carries it out.
    name: 'Reports',
    headers: ['reportId', 'createdAt', 'reporterId', 'reportedUserId', 'convId', 'msgId', 'reason', 'details',
              'messageText', 'status', 'action', 'handledAt']
  }
});

const MESSAGES_PREFIX = 'Messages_';
const MESSAGE_HEADERS = Object.freeze([
  'msgId', 'convId', 'senderId', 'type', 'body', 'mediaFileId', 'replyTo', 'clientId', 'createdAt', 'deleted', 'expiresAt', 'hiddenFor'
]);
const MESSAGE_TYPES = Object.freeze(['text', 'image', 'audio']);   // types a user may send; 'system' is server-only
const ACTIVE_ROLES = Object.freeze(['member', 'admin']);            // 'left' and 'removed' rows are kept, not deleted

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/* ===== Util.gs ===== */
/**
 * ZeeMoApp back end: shared helpers.
 */

class ApiError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

function nowIso_() {
  return new Date().toISOString();
}

/** UUID v4 is the only cryptographically random source in Apps Script. */
function randomHex_(chars) {
  let out = '';
  while (out.length < chars) out += Utilities.getUuid().replace(/-/g, '');
  return out.substring(0, chars);
}

function randomDigits_(n) {
  let out = '';
  while (out.length < n) {
    const hex = Utilities.getUuid().replace(/-/g, '');
    for (let i = 0; i + 2 <= hex.length && out.length < n; i += 2) {
      const byte = parseInt(hex.substr(i, 2), 16);
      if (byte < 250) out += String(byte % 10);   // reject 250..255 to avoid modulo bias
    }
  }
  return out;
}

/**
 * Time-sortable ID: prefix + 9 base-36 chars of epoch ms + random suffix.
 * Lexical order equals creation order for IDs with the same prefix.
 */
function newId_(prefix, ts) {
  const t = (ts || Date.now()).toString(36).padStart(9, '0');
  const r = parseInt(randomHex_(6), 16).toString(36).padStart(5, '0').substring(0, 5);
  return prefix + t + r;
}

/** Unique change stamp for list and receipt versions. Clients compare for inequality only. */
function newVersion_() {
  return newId_('v');
}

function tsFromId_(id) {
  return parseInt(String(id).substring(1, 10), 36);
}

function monthKey_(ts) {
  return Utilities.formatDate(new Date(ts), 'UTC', 'yyyyMM');
}

function isMsgId_(id) {
  return typeof id === 'string' && /^m[0-9a-z]{14}$/.test(id);
}

function bytesToHex_(bytes) {
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function sha256Hex_(text) {
  return bytesToHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}

function hmacHex_(text) {
  const pepper = prop_('OTP_PEPPER');
  if (!pepper) throw new ApiError('NOT_CONFIGURED', 'Run setup() first.');
  return bytesToHex_(Utilities.computeHmacSha256Signature(text, pepper));
}

/** Constant-time comparison for equal-length hex strings. */
function safeEqual_(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/** Strip control characters (keep \n and \t), trim, cap length. Rendering must still escape HTML. */
function cleanText_(value, max) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim()
    .substring(0, max);
}

function requireString_(value, field, max) {
  const v = cleanText_(value, max || 200);
  if (!v) throw new ApiError('BAD_REQUEST', `Missing ${field}.`);
  return v;
}

/** Normalise to E.164. Nigerian local numbers (0803...) become +234803... */
function normalisePhone_(raw) {
  let d = String(raw || '').replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '0') d = CONFIG.DEFAULT_COUNTRY_CODE + d.substring(1);
  if (d.length < 10 || d.length > 15) throw new ApiError('INVALID_PHONE', 'Enter a valid phone number.');
  return '+' + d;
}

/** Lower-case and validate an email address. */
function normaliseEmail_(raw) {
  const e = String(raw || '').trim().toLowerCase();
  if (e.length > 254 || !/^[^\s@<>()",;:]+@[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)+$/.test(e)) {
    throw new ApiError('INVALID_EMAIL', 'Enter a valid email address.');
  }
  return e;
}

/* ---------- CacheService helpers ---------- */

function cache_() {
  return CacheService.getScriptCache();
}

function cacheGetJson_(key) {
  const v = cache_().get(key);
  if (v === null) return null;
  try { return JSON.parse(v); } catch (e) { return null; }
}

function cachePutJson_(key, value, ttlSec) {
  const s = JSON.stringify(value);
  if (s.length > CONFIG.CACHE_VALUE_LIMIT) return false;
  cache_().put(key, s, Math.min(ttlSec || 21600, 21600));
  return true;
}

const CK = Object.freeze({
  session: h => 'ss:' + h,
  otp: p => 'otp:' + p,
  otpRate: p => 'otr:' + p,
  convLast: c => 'cl:' + c,
  recent: c => 'cm:' + c,
  receiptVer: c => 'rv:' + c,
  userConvs: u => 'uc2:' + u,
  listVer: u => 'lv:' + u,
  lastSeen: u => 'ls:' + u,
  dedupe: (u, id) => 'dd:' + u + ':' + id,
  deletions: c => 'dl:' + c,
  deleteVer: c => 'dv:' + c,
  unlocked: tokenHash => 'ul:' + tokenHash,
  passAttempts: u => 'pa:' + u,
  hasPasscode: u => 'hp:' + u,
  blocked: () => 'blk',
  blocksBy: u => 'bk:' + u,
  terms: u => 'tv:' + u,
  reportCount: u => 'rc:' + u,
  msgClock: () => 'mts'
});

/* ===== Crypto.gs ===== */
/**
 * ZeeMoApp back end: encryption of message text at rest.
 *
 * Message text and chat-list previews are stored in the sheet encrypted, so the
 * spreadsheet, its version history and its backups are unreadable without keys
 * that live only in Script Properties (which have no version history).
 *
 * Cipher: HMAC-SHA256 as a pseudo-random function in counter mode (keystream),
 * with an HMAC-SHA256 tag over nonce and ciphertext (encrypt-then-MAC).
 * Stored form:  e1:<epoch>:<nonce>:<base64 ciphertext>:<tag>
 *
 * Keys ("epochs") are random 256-bit values in Script Properties named EK_<epoch>:
 *   m<yyyyMM>    messages that do not disappear, one key per month
 *   x<yyyyMMdd>  disappearing messages, one key per expiry day
 * The hourly clean-up deletes x-keys once their day has passed, so expired messages
 * cannot be recovered from version history or backups by anyone, the owner included.
 *
 * Rows written before encryption was added have no "e1:" prefix and are read as-is;
 * run encryptExistingMessages() once to encrypt them.
 */

const CRYPTO_PREFIX = 'e1:';
const KEY_PREFIX = 'EK_';

/* ---------- UTF-8 (Apps Script has no TextEncoder) ---------- */

function utf8Encode_(str) {
  const out = [];
  for (let i = 0; i < str.length; i++) {
    let c = str.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < str.length) {
      const d = str.charCodeAt(i + 1);
      if (d >= 0xdc00 && d <= 0xdfff) { c = 0x10000 + ((c - 0xd800) << 10) + (d - 0xdc00); i++; }
    }
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}

function utf8Decode_(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length;) {
    const b = bytes[i++] & 255;
    let c;
    if (b < 0x80) c = b;
    else if (b < 0xe0) c = ((b & 31) << 6) | (bytes[i++] & 63);
    else if (b < 0xf0) c = ((b & 15) << 12) | ((bytes[i++] & 63) << 6) | (bytes[i++] & 63);
    else c = ((b & 7) << 18) | ((bytes[i++] & 63) << 12) | ((bytes[i++] & 63) << 6) | (bytes[i++] & 63);
    if (c >= 0x10000) { c -= 0x10000; out += String.fromCharCode(0xd800 + (c >> 10), 0xdc00 + (c & 1023)); }
    else out += String.fromCharCode(c);
  }
  return out;
}

/* ---------- Keys ---------- */

/** Keys cached for this execution. Script Properties are read once. */
function keyRing_() {
  if (!keyRing_.keys) {
    const all = PropertiesService.getScriptProperties().getProperties();
    keyRing_.keys = {};
    Object.keys(all).forEach(k => { if (k.indexOf(KEY_PREFIX) === 0) keyRing_.keys[k.substring(KEY_PREFIX.length)] = all[k]; });
  }
  return keyRing_.keys;
}

/** Existing key, or a new one. Creating must happen inside withLock_ so two runs cannot race. */
function epochKey_(epoch, create) {
  const ring = keyRing_();
  if (ring[epoch]) return ring[epoch];
  if (!create) return null;
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(KEY_PREFIX + epoch);   // another run may have just made it
  if (!key) {
    key = randomHex_(64);
    props.setProperty(KEY_PREFIX + epoch, key);
  }
  ring[epoch] = key;
  return key;
}

/** Which key a message uses: its expiry day if it disappears, otherwise the month it was sent. */
function epochFor_(createdAtIso, expiresAtIso) {
  if (expiresAtIso) return 'x' + expiresAtIso.substring(0, 10).replace(/-/g, '');
  return 'm' + createdAtIso.substring(0, 7).replace('-', '');
}

/* ---------- Seal and open ---------- */

function keystreamXor_(bytes, key, nonce) {
  const out = new Array(bytes.length);
  for (let block = 0; block * 32 < bytes.length; block++) {
    const ks = Utilities.computeHmacSha256Signature('k|' + nonce + '|' + block, key);
    const base = block * 32;
    for (let j = 0; j < 32 && base + j < bytes.length; j++) out[base + j] = ((bytes[base + j] ^ ks[j]) << 24) >> 24;
  }
  return out;
}

function tag_(key, nonce, ct) {
  return bytesToHex_(Utilities.computeHmacSha256Signature('t|' + nonce + '|' + ct, key)).substring(0, 32);
}

/** Encrypt text with the epoch's key. Empty text stays empty. Call inside withLock_. */
function sealText_(plain, epoch) {
  if (!plain) return '';
  const key = epochKey_(epoch, true);
  const nonce = randomHex_(24);
  const ct = Utilities.base64Encode(keystreamXor_(utf8Encode_(String(plain)), key, nonce));
  return [CRYPTO_PREFIX.slice(0, -1), epoch, nonce, ct, tag_(key, nonce, ct)].join(':');
}

/** Decrypt. Text without the prefix is returned as-is; a shredded key or bad tag gives ''. */
function openText_(stored) {
  if (!stored || String(stored).indexOf(CRYPTO_PREFIX) !== 0) return stored || '';
  const parts = String(stored).split(':');
  if (parts.length !== 5) return '';
  const epoch = parts[1], nonce = parts[2], ct = parts[3], tag = parts[4];
  const key = epochKey_(epoch, false);
  if (!key) return '';
  if (!safeEqual_(tag, tag_(key, nonce, ct))) {
    console.warn('Message failed its integrity check; it was changed outside the app.');
    return '';
  }
  return utf8Decode_(keystreamXor_(Utilities.base64Decode(ct), key, nonce));
}

function isSealed_(stored) {
  return String(stored || '').indexOf(CRYPTO_PREFIX) === 0;
}

/** Delete keys of disappearing messages whose expiry day has passed. Returns how many. */
function shredExpiredKeys_() {
  const props = PropertiesService.getScriptProperties();
  const yesterday = Utilities.formatDate(new Date(Date.now() - 86400000), 'UTC', 'yyyyMMdd');
  let n = 0;
  Object.keys(props.getProperties()).forEach(k => {
    if (k.indexOf(KEY_PREFIX + 'x') !== 0) return;
    if (k.substring(KEY_PREFIX.length + 1) < yesterday) {
      props.deleteProperty(k);
      n++;
    }
  });
  keyRing_.keys = null;
  return n;
}

/* ===== DataStore.gs ===== */
/**
 * ZeeMoApp back end: the only file that touches Sheets directly.
 * Replace this layer (same function names) to migrate to Firestore or Supabase.
 *
 * Invariants:
 *  - Rows are never deleted or sorted, so a row index stays valid forever.
 *    That is what makes single-cell updates outside the lock safe.
 *  - All cells are plain text ('@'), so IDs and phone numbers are never coerced to numbers.
 *  - Appends always run inside withLock_().
 */

function ss_() {
  if (!ss_.instance) {
    const id = prop_('SPREADSHEET_ID');
    if (!id) throw new ApiError('NOT_CONFIGURED', 'Run setup() first.');
    ss_.instance = SpreadsheetApp.openById(id);
  }
  return ss_.instance;
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new ApiError('NOT_CONFIGURED', `Missing sheet ${name}. Run setup().`);
  return sh;
}

function headersFor_(name) {
  if (name.indexOf(MESSAGES_PREFIX) === 0) return MESSAGE_HEADERS;
  const def = Object.keys(SHEETS).map(k => SHEETS[k]).filter(s => s.name === name)[0];
  if (!def) throw new ApiError('SERVER_ERROR', `Unknown sheet ${name}.`);
  return def.headers;
}

function rowToObj_(headers, row) {
  const o = {};
  headers.forEach((h, i) => { o[h] = row[i] === undefined || row[i] === null ? '' : String(row[i]); });
  return o;
}

function objToRow_(headers, obj) {
  return headers.map(h => (obj[h] === undefined || obj[h] === null ? '' : String(obj[h])));
}

/** Serialise writes. Throws BUSY if the lock cannot be taken in time. */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) throw new ApiError('BUSY', 'Server busy. Please retry.');
  try {
    const result = fn();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function appendObj_(sheetName, obj) {
  sheet_(sheetName).appendRow(objToRow_(headersFor_(sheetName), obj));
}

function appendObjs_(sheetName, objs) {
  if (!objs.length) return;
  const headers = headersFor_(sheetName);
  const sh = sheet_(sheetName);
  const rows = objs.map(o => objToRow_(headers, o));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/** Exact match on a column using TextFinder. Returns [{rowIdx, obj}]. */
function findRows_(sheetName, column, value, firstOnly) {
  const sh = sheet_(sheetName);
  const headers = headersFor_(sheetName);
  const col = headers.indexOf(column) + 1;
  if (col < 1) throw new ApiError('SERVER_ERROR', `Unknown column ${column}.`);
  const last = sh.getLastRow();
  if (last < 2 || value === '' || value === null || value === undefined) return [];

  const finder = sh.getRange(2, col, last - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).matchCase(true);
  const cells = firstOnly ? [finder.findNext()].filter(Boolean) : finder.findAll();
  const width = Math.min(headers.length, sh.getMaxColumns());
  return cells.map(c => {
    const rowIdx = c.getRow();
    const row = sh.getRange(rowIdx, 1, 1, width).getValues()[0];
    return { rowIdx: rowIdx, obj: rowToObj_(headers, row) };
  });
}

function findOne_(sheetName, column, value) {
  return findRows_(sheetName, column, value, true)[0] || null;
}

function readAll_(sheetName) {
  const sh = sheet_(sheetName);
  const headers = headersFor_(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = Math.min(headers.length, sh.getMaxColumns());
  return sh.getRange(2, 1, last - 1, width).getValues().map(r => rowToObj_(headers, r));
}

/** Update named columns of one row. Safe outside the lock because rows never move. */
function updateFields_(sheetName, rowIdx, fields) {
  const sh = sheet_(sheetName);
  const headers = headersFor_(sheetName);
  Object.keys(fields).forEach(k => {
    const col = headers.indexOf(k) + 1;
    if (col < 1) throw new ApiError('SERVER_ERROR', `Unknown column ${k}.`);
    sh.getRange(rowIdx, col).setValue(String(fields[k]));
  });
}

/* ---------- Monthly message shards ---------- */

function messagesSheetName_(monthKey) {
  return MESSAGES_PREFIX + monthKey;
}

/** Get a month shard; create it (caller must hold the lock) when create is true. */
function messagesSheet_(monthKey, create) {
  const name = messagesSheetName_(monthKey);
  let sh = ss_().getSheetByName(name);
  if (!sh && create) sh = createSheetWithHeaders_(name, MESSAGE_HEADERS);
  return sh;
}

function createSheetWithHeaders_(name, headers) {
  const sh = ss_().insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  const extra = sh.getMaxColumns() - headers.length;
  if (extra > 0) sh.deleteColumns(headers.length + 1, extra);
  return sh;
}

/** Month keys from fromTs to toTs inclusive, oldest first. */
function monthRange_(fromTs, toTs) {
  const keys = [];
  const d = new Date(fromTs);
  d.setUTCDate(1);
  d.setUTCHours(0, 0, 0, 0);
  while (d.getTime() <= toTs) {
    keys.push(monthKey_(d.getTime()));
    d.setUTCMonth(d.getUTCMonth() + 1);
  }
  return keys;
}

/** All messages of one conversation in one month shard, oldest first. */
function readConvMessagesInMonth_(convId, monthKey) {
  const sh = messagesSheet_(monthKey, false);
  if (!sh) return [];
  // One bulk read beats per-row TextFinder fetches once a conversation has more than a few messages.
  return readAll_(sh.getName())
    .filter(m => m.convId === convId)
    .sort((a, b) => (a.msgId < b.msgId ? -1 : 1));
}

/* ---------- Schema upgrades ---------- */

/** Bump when a column is added. Existing sheets are upgraded automatically on the next request. */
const SCHEMA_VERSION = '6';

/** Upgrade sheet headers once per schema version. Cheap after the first call (cache hit). */
function ensureSchema_() {
  if (cache_().get('schema') === SCHEMA_VERSION) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA_VERSION') !== SCHEMA_VERSION) {
    withLock_(() => {
      if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION) return;
      upgradeSheets_();
      props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION);
    });
  }
  cache_().put('schema', SCHEMA_VERSION, 21600);
}

function upgradeSheets_() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(k => {
    const def = SHEETS[k];
    const sh = ss.getSheetByName(def.name);
    if (sh) ensureHeaders_(sh, def.headers);
    else createSheetWithHeaders_(def.name, def.headers);
  });
  ss.getSheets().forEach(sh => {
    if (sh.getName().indexOf(MESSAGES_PREFIX) === 0) ensureHeaders_(sh, MESSAGE_HEADERS);
  });
}

/** Append any header columns missing from an existing sheet. Existing data is untouched. */
function ensureHeaders_(sh, headers) {
  const width = Math.max(Math.min(sh.getLastColumn(), sh.getMaxColumns()), 1);
  const current = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
  headers.forEach((h, i) => {
    if (current[i] === h) return;
    if (current[i]) throw new Error(`Sheet ${sh.getName()} column ${i + 1} is "${current[i]}", expected "${h}". Fix it by hand.`);
    if (sh.getMaxColumns() < i + 1) sh.insertColumnsAfter(sh.getMaxColumns(), i + 1 - sh.getMaxColumns());
    sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, i + 1).setValue(h).setFontWeight('bold');
  });
}

/* ===== Auth.gs ===== */
/**
 * ZeeMoApp back end: one-time-code sign-in by phone (SMS) or email, and sessions.
 * The platform owner chooses the method with the SIGNIN_METHOD Script Property:
 *   phone  codes by SMS through Termii (default)
 *   email  codes by email from the script owner's Google account (MailApp)
 *   both   people choose on the sign-in screen
 * A phone account and an email account are separate accounts; they are not merged.
 */

function signinMethods_() {
  const v = String(prop_('SIGNIN_METHOD') || CONFIG.DEFAULT_SIGNIN_METHOD).trim().toLowerCase();
  if (v === 'both') return CONFIG.SIGNIN_METHODS.slice();
  return CONFIG.SIGNIN_METHODS.indexOf(v) >= 0 ? [v] : [CONFIG.DEFAULT_SIGNIN_METHOD];
}

/** Public: tells the sign-in screen which methods to show. */
function authConfig() {
  return {
    appName: CONFIG.APP_NAME,
    methods: signinMethods_(),
    devMode: prop_('DEV_MODE') === 'true',
    security: securitySettings_(),
    termsVersion: CONFIG.TERMS_VERSION,
    links: legalLinks_()
  };
}

/** Resolve {method, phone|email} into a checked, normalised identifier. */
function identity_(p) {
  const allowed = signinMethods_();
  const method = p.method || (p.email ? 'email' : 'phone');
  if (allowed.indexOf(method) === -1) {
    throw new ApiError('METHOD_DISABLED', method === 'email'
      ? 'Email sign-in is turned off. Use your phone number.'
      : 'Phone sign-in is turned off. Use your email address.');
  }
  const value = method === 'email' ? normaliseEmail_(p.email) : normalisePhone_(p.phone);
  return { method: method, value: value, key: method.charAt(0) + ':' + value };
}

const BLOCKED_MESSAGE = 'This account has been blocked. Contact the administrator.';

function requestOtp(p) {
  const id = identity_(p);
  if (isBlockedIdentity_(id)) throw new ApiError('BLOCKED', BLOCKED_MESSAGE);

  const rateKey = CK.otpRate(id.key);
  const count = Number(cache_().get(rateKey) || 0);
  if (count >= CONFIG.OTP_MAX_REQUESTS) {
    throw new ApiError('RATE_LIMITED', 'Too many codes requested. Try again in 15 minutes.');
  }

  const code = randomDigits_(CONFIG.OTP_LENGTH);
  const dev = prop_('DEV_MODE') === 'true';
  if (!dev) {
    if (id.method === 'email') sendOtpEmail_(id.value, code);
    else sendOtpSms_(id.value, code);
  }

  // Store only after sending succeeds, so a failed send does not use up a request.
  cachePutJson_(CK.otp(id.key), { h: hmacHex_(id.key + ':' + code), a: 0 }, CONFIG.OTP_TTL_SEC);
  cache_().put(rateKey, String(count + 1), CONFIG.OTP_WINDOW_SEC);

  const out = { sent: true, method: id.method, expiresInSec: CONFIG.OTP_TTL_SEC };
  out[id.method] = id.value;
  if (dev) out.devCode = code;   // never enable DEV_MODE in production
  return out;
}

/** Check a sign-in code for an identity; it is used up on success. */
function verifyCode_(id, rawCode) {
  const code = String(rawCode || '').replace(/\D/g, '');
  const key = CK.otp(id.key);
  const rec = cacheGetJson_(key);
  if (!rec) throw new ApiError('OTP_EXPIRED', 'Code expired. Request a new one.');
  if (rec.a >= CONFIG.OTP_MAX_ATTEMPTS) {
    cache_().remove(key);
    throw new ApiError('OTP_LOCKED', 'Too many wrong attempts. Request a new code.');
  }
  if (!safeEqual_(rec.h, hmacHex_(id.key + ':' + code))) {
    rec.a += 1;
    cachePutJson_(key, rec, CONFIG.OTP_TTL_SEC);
    throw new ApiError('OTP_INVALID', 'Incorrect code.');
  }
  cache_().remove(key);
}

function verifyOtp(p) {
  const id = identity_(p);
  if (isBlockedIdentity_(id)) throw new ApiError('BLOCKED', BLOCKED_MESSAGE);
  verifyCode_(id, p.code);

  const device = cleanText_(p.device, 80);
  const displayName = cleanText_(p.displayName, 40);

  const result = withLock_(() => {
    const found = findOne_(SHEETS.USERS.name, id.method, id.value);
    let user;
    let isNew = false;
    if (found) {
      user = found.obj;
    } else {
      isNew = true;
      user = {
        userId: newId_('u'),
        phone: id.method === 'phone' ? id.value : '',
        email: id.method === 'email' ? id.value : '',
        displayName: displayName || defaultName_(id),
        avatarFileId: '',
        about: '',
        createdAt: nowIso_(),
        lastSeen: nowIso_()
      };
      appendObj_(SHEETS.USERS.name, user);
    }

    const session = createSession_(user.userId, device);
    return { token: session.token, expiresAt: session.expiresAt, isNewUser: isNew, user: privateUser_(user) };
  });
  // Wrong-passcode counts are kept: a sign-in code alone does not reset them.
  return result;
}

function defaultName_(id) {
  if (id.method === 'phone') return 'User ' + id.value.slice(-4);
  const local = id.value.split('@')[0].replace(/[._-]+/g, ' ').trim();
  return local ? local.replace(/\b\w/g, c => c.toUpperCase()).substring(0, 40) : 'New user';
}

/** Caller must hold the lock. Only the SHA-256 hash of the token is stored. */
function createSession_(userId, device) {
  const token = randomHex_(64);                  // 256 bits
  const tokenHash = sha256Hex_(token);
  const expiresAt = new Date(Date.now() + CONFIG.SESSION_TTL_DAYS * 86400000).toISOString();
  appendObj_(SHEETS.SESSIONS.name, {
    tokenHash: tokenHash, userId: userId, createdAt: nowIso_(),
    expiresAt: expiresAt, device: device, revoked: 'FALSE'
  });
  cache_().put(CK.session(tokenHash), userId, CONFIG.SESSION_CACHE_SEC);
  return { token: token, expiresAt: expiresAt };
}

/** Returns the userId for a valid token or throws UNAUTHORISED. */
function requireSession_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/.test(token)) {
    throw new ApiError('UNAUTHORISED', 'Please sign in.');
  }
  const tokenHash = sha256Hex_(token);
  const cached = cache_().get(CK.session(tokenHash));
  if (cached) {
    if (blockList_().ids.indexOf(cached) >= 0) throw new ApiError('UNAUTHORISED', BLOCKED_MESSAGE);
    return cached;
  }

  const found = findOne_(SHEETS.SESSIONS.name, 'tokenHash', tokenHash);
  if (!found || found.obj.revoked === 'TRUE') throw new ApiError('UNAUTHORISED', 'Please sign in.');

  const remainingSec = Math.floor((new Date(found.obj.expiresAt).getTime() - Date.now()) / 1000);
  if (remainingSec <= 0) throw new ApiError('UNAUTHORISED', 'Session expired. Please sign in.');

  if (blockList_().ids.indexOf(found.obj.userId) >= 0) throw new ApiError('UNAUTHORISED', BLOCKED_MESSAGE);
  cache_().put(CK.session(tokenHash), found.obj.userId, Math.min(CONFIG.SESSION_CACHE_SEC, remainingSec));
  return found.obj.userId;
}

/** Sign out every device of a user, except the session whose hash is kept. Returns how many. */
function revokeAllSessions_(userId, keepHash) {
  const sh = sheet_(SHEETS.SESSIONS.name);
  const col = SHEETS.SESSIONS.headers.indexOf('revoked') + 1;
  let n = 0;
  readAll_(SHEETS.SESSIONS.name).forEach((s, i) => {
    if (s.userId !== userId || s.revoked === 'TRUE' || s.tokenHash === keepHash) return;
    sh.getRange(i + 2, col).setValue('TRUE');
    cache_().remove(CK.session(s.tokenHash));
    cache_().remove(CK.unlocked(s.tokenHash));
    n++;
  });
  return n;
}

/* ---------- Blocking ---------- */

/**
 * Users whose "blocked" cell says TRUE (also yes or 1). Read at most every two minutes,
 * so a block takes effect within that time; run applyBlocksNow() for immediate effect.
 */
function blockList_() {
  const cached = cacheGetJson_(CK.blocked());
  if (cached) return cached;
  const list = { ids: [], keys: [] };
  readAll_(SHEETS.USERS.name).forEach(u => {
    if (['true', 'yes', '1'].indexOf(String(u.blocked).trim().toLowerCase()) === -1) return;
    list.ids.push(u.userId);
    if (u.phone) list.keys.push('p:' + u.phone);
    if (u.email) list.keys.push('e:' + u.email);
  });
  cachePutJson_(CK.blocked(), list, CONFIG.BLOCK_CACHE_SEC);
  return list;
}

function isBlockedIdentity_(id) {
  return blockList_().keys.indexOf(id.key) >= 0;
}

function logout(ctx) {
  const tokenHash = sha256Hex_(ctx.token);
  const found = findOne_(SHEETS.SESSIONS.name, 'tokenHash', tokenHash);
  if (found) updateFields_(SHEETS.SESSIONS.name, found.rowIdx, { revoked: 'TRUE' });
  cache_().remove(CK.session(tokenHash));
  return { loggedOut: true };
}

/**
 * Email code from the script owner's Google account.
 * Daily limit: about 100 recipients on a free Gmail account, 1,500 on Google Workspace.
 */
function sendOtpEmail_(email, code) {
  if (MailApp.getRemainingDailyQuota() < 1) {
    throw new ApiError('EMAIL_LIMIT', 'Sign-in by email is unavailable for the rest of today. Try again tomorrow.');
  }
  const minutes = Math.round(CONFIG.OTP_TTL_SEC / 60);
  const text = `Your ${CONFIG.APP_NAME} sign-in code is ${code}.\n\nIt expires in ${minutes} minutes. If you did not ask for it, ignore this email.`;
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:16px;color:#151935">' +
    `<p>Your ${CONFIG.APP_NAME} sign-in code is:</p>` +
    `<p style="font-size:30px;font-weight:bold;letter-spacing:6px;margin:12px 0">${code}</p>` +
    `<p>It expires in ${minutes} minutes. If you did not ask for it, ignore this email.</p></div>`;
  try {
    MailApp.sendEmail({ to: email, subject: `${code} is your ${CONFIG.APP_NAME} code`, body: text, htmlBody: html, name: CONFIG.APP_NAME });
  } catch (err) {
    console.error('MailApp error', err);
    throw new ApiError('EMAIL_FAILED', 'Could not send the code. Check the address and try again.');
  }
}

/**
 * Termii SMS. Base URL is account-specific in some regions: copy it from the Termii dashboard.
 * The sender ID must be approved by Termii; the "dnd" route reaches numbers on Do-Not-Disturb.
 */
function sendOtpSms_(phone, code) {
  const apiKey = prop_('TERMII_API_KEY');
  if (!apiKey) throw new ApiError('NOT_CONFIGURED', 'SMS provider is not configured.');
  const base = (prop_('TERMII_BASE_URL') || 'https://api.ng.termii.com').replace(/\/+$/, '');

  const res = UrlFetchApp.fetch(base + '/api/sms/send', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      to: phone.replace('+', ''),
      from: prop_('TERMII_SENDER_ID') || CONFIG.SMS_SENDER_ID,
      sms: `Your ${CONFIG.APP_NAME} code is ${code}. It expires in ${Math.round(CONFIG.OTP_TTL_SEC / 60)} minutes. Do not share it.`,
      type: 'plain',
      channel: 'dnd',
      api_key: apiKey
    })
  });

  if (res.getResponseCode() >= 300) {
    console.error('Termii error', res.getResponseCode(), res.getContentText());
    throw new ApiError('SMS_FAILED', 'Could not send the code. Try again shortly.');
  }
}

/* ===== Users.gs ===== */
/**
 * ZeeMoApp back end: profiles and presence.
 */

/**
 * What other people may see. Last seen is shown only to people who share a chat with
 * this user, only if the user allows it, and only if the owner has not hidden it for all.
 * showLastSeen: pass true only when the viewer shares a chat with u.
 */
function publicUser_(u, showLastSeen) {
  const visible = showLastSeen === true && u.lastSeenPrivacy !== 'nobody';
  return {
    userId: u.userId,
    displayName: u.displayName,
    avatarFileId: u.avatarFileId,
    about: u.about,
    lastSeen: visible ? (cache_().get(CK.lastSeen(u.userId)) || u.lastSeen) : ''
  };
}

function privateUser_(u) {
  const o = publicUser_(u, true);
  o.lastSeen = cache_().get(CK.lastSeen(u.userId)) || u.lastSeen;
  o.lastSeenPrivacy = u.lastSeenPrivacy === 'nobody' ? 'nobody' : 'contacts';
  o.phone = u.phone;
  o.email = u.email;
  o.hasPasscode = !!u.passcode;   // the passcode itself never leaves the server
  o.termsAccepted = String(u.termsVersion || '') === CONFIG.TERMS_VERSION;
  o.termsVersion = CONFIG.TERMS_VERSION;
  o.deletionRequestedAt = u.deletionRequestedAt || '';
  return o;
}

function getUser_(userId) {
  const found = findOne_(SHEETS.USERS.name, 'userId', userId);
  if (!found) throw new ApiError('NOT_FOUND', 'User not found.');
  return found;
}

function me(ctx) {
  return privateUser_(getUser_(ctx.userId).obj);
}

function updateProfile(ctx, p) {
  const found = getUser_(ctx.userId);
  const fields = {};
  if (p.displayName !== undefined) fields.displayName = requireString_(p.displayName, 'display name', 40);
  if (p.about !== undefined) fields.about = cleanText_(p.about, 140);
  if (p.lastSeenPrivacy !== undefined) fields.lastSeenPrivacy = p.lastSeenPrivacy === 'nobody' ? 'nobody' : 'contacts';
  if (Object.keys(fields).length) updateFields_(SHEETS.USERS.name, found.rowIdx, fields);
  return privateUser_(Object.assign(found.obj, fields));
}

/**
 * Look up a contact by phone or email: { phone } or { email }, or { query } holding either.
 * Enumeration is acceptable for a closed community but should be rate-limited before a public launch.
 */
function findUser(ctx, p) {
  const q = String(p.query || p.email || p.phone || '').trim();
  const isEmail = q.indexOf('@') >= 0;
  const found = isEmail
    ? findOne_(SHEETS.USERS.name, 'email', normaliseEmail_(q))
    : findOne_(SHEETS.USERS.name, 'phone', normalisePhone_(q));
  return found ? publicUser_(found.obj, false) : null;
}

/**
 * Look up a contact by phone. Kept for older clients; use findUser. This allows number enumeration, which is acceptable
 * for a closed community but should be rate-limited before any public launch.
 */
function findUserByPhone(ctx, p) {
  const found = findOne_(SHEETS.USERS.name, 'phone', normalisePhone_(p.phone));
  return found ? publicUser_(found.obj, false) : null;
}

/** Presence lives in cache; the sheet is written at most once a minute per user. */
function touchLastSeen_(userId) {
  const now = nowIso_();
  const key = CK.lastSeen(userId);
  cache_().put(key, now, 21600);
  if (cache_().get(key + ':w')) return;          // sheet written within the last minute
  const found = findOne_(SHEETS.USERS.name, 'userId', userId);
  if (found) updateFields_(SHEETS.USERS.name, found.rowIdx, { lastSeen: now });
  cache_().put(key + ':w', '1', CONFIG.LAST_SEEN_WRITE_SEC);
}

/* ===== Conversations.gs ===== */
/**
 * ZeeMoApp back end: direct chats, groups and membership.
 */

function participantKey_(convId, userId) {
  return convId + ':' + userId;
}

function isActive_(participant) {
  return ACTIVE_ROLES.indexOf(participant.role) >= 0;
}

/**
 * The user's active memberships as { convId: joinedAt }, cached for 6 h and
 * invalidated on any membership change. joinedAt hides messages from before a member joined.
 */
function userMemberships_(userId) {
  const cached = cacheGetJson_(CK.userConvs(userId));
  if (cached) return cached;
  const map = {};
  readAll_(SHEETS.PARTICIPANTS.name)
    .filter(p => p.userId === userId && isActive_(p))
    .forEach(p => { map[p.convId] = p.joinedAt; });
  cachePutJson_(CK.userConvs(userId), map, 21600);
  return map;
}

function userConvIds_(userId) {
  return Object.keys(userMemberships_(userId));
}

/** Messages a member may see: none from before they (last) joined. */
function visibleSince_(convId, userId) {
  return userMemberships_(userId)[convId] || '';
}

function invalidateMembership_(userIds) {
  const ver = newVersion_();
  userIds.forEach(u => {
    cache_().remove(CK.userConvs(u));
    cache_().put(CK.listVer(u), ver, 21600);
  });
}

function listVersion_(userId) {
  let v = cache_().get(CK.listVer(userId));
  if (!v) {
    v = newVersion_();
    cache_().put(CK.listVer(userId), v, 21600);
  }
  return v;
}

function assertParticipant_(convId, userId) {
  if (!Object.prototype.hasOwnProperty.call(userMemberships_(userId), convId)) {
    throw new ApiError('FORBIDDEN', 'You are not a member of this conversation.');
  }
}

function createDirect(ctx, p) {
  const otherId = requireString_(p.otherUserId, 'otherUserId', 40);
  if (otherId === ctx.userId) throw new ApiError('BAD_REQUEST', 'You cannot start a chat with yourself.');
  getUser_(otherId);

  const directKey = [ctx.userId, otherId].sort().join(':');
  const result = withLock_(() => {
    const existing = findOne_(SHEETS.CONVERSATIONS.name, 'directKey', directKey);
    if (existing) return { conv: existing.obj, created: false };

    const now = nowIso_();
    const conv = {
      convId: newId_('c'), type: 'direct', title: '', directKey: directKey,
      createdBy: ctx.userId, createdAt: now,
      lastMsgId: '', lastMsgAt: now, lastMsgSenderId: '', lastMsgPreview: '',
      disappearSec: '0', lastMsgExpiresAt: ''
    };
    appendObj_(SHEETS.CONVERSATIONS.name, conv);
    appendObjs_(SHEETS.PARTICIPANTS.name, [ctx.userId, otherId].map(u => ({
      key: participantKey_(conv.convId, u), convId: conv.convId, userId: u,
      role: 'member', joinedAt: now, lastReadMsgId: '', lastDeliveredMsgId: '', leftAt: ''
    })));
    return { conv: conv, created: true };
  });

  if (result.created) invalidateMembership_([ctx.userId, otherId]);
  return { convId: result.conv.convId, created: result.created };
}

function createGroup(ctx, p) {
  const title = requireString_(p.title, 'group title', 60);
  const requested = Array.isArray(p.memberIds) ? p.memberIds.map(String) : [];
  const members = Array.from(new Set([ctx.userId].concat(requested)));
  if (members.length < 2) throw new ApiError('BAD_REQUEST', 'Add at least one member.');
  if (members.length > CONFIG.MAX_GROUP_MEMBERS) throw new ApiError('BAD_REQUEST', 'Group is too large.');

  const known = new Set(readAll_(SHEETS.USERS.name).map(u => u.userId));
  const unknown = members.filter(m => !known.has(m));
  if (unknown.length) throw new ApiError('BAD_REQUEST', 'Some members do not exist.');

  const now = nowIso_();
  const conv = {
    convId: newId_('c'), type: 'group', title: title, directKey: '',
    createdBy: ctx.userId, createdAt: now,
    lastMsgId: '', lastMsgAt: now, lastMsgSenderId: '', lastMsgPreview: '',
    disappearSec: '0', lastMsgExpiresAt: ''
  };

  withLock_(() => {
    appendObj_(SHEETS.CONVERSATIONS.name, conv);
    appendObjs_(SHEETS.PARTICIPANTS.name, members.map(u => ({
      key: participantKey_(conv.convId, u), convId: conv.convId, userId: u,
      role: u === ctx.userId ? 'admin' : 'member', joinedAt: now,
      lastReadMsgId: '', lastDeliveredMsgId: '', leftAt: ''
    })));
  });

  invalidateMembership_(members);
  return { convId: conv.convId, created: true };
}

/** Chat list, newest activity first. Direct chats take the other person's name as title. */
function listConversations(ctx) {
  const mine = new Set(userConvIds_(ctx.userId));
  if (!mine.size) return { listVersion: listVersion_(ctx.userId), conversations: [] };

  const convs = readAll_(SHEETS.CONVERSATIONS.name).filter(c => mine.has(c.convId));
  const parts = readAll_(SHEETS.PARTICIPANTS.name).filter(p => mine.has(p.convId) && isActive_(p));
  const users = {};
  const nowIso = nowIso_();
  const showSeen = securitySettings_().showLastSeen;
  readAll_(SHEETS.USERS.name).forEach(u => { users[u.userId] = u; });
  const iBlocked = new Set(blocksBy_(ctx.userId));

  const list = convs.map(c => {
    const members = parts.filter(p => p.convId === c.convId);
    const self = members.filter(p => p.userId === ctx.userId)[0] || {};
    let title = c.title;
    let avatarFileId = c.type === 'group' ? (c.avatarFileId || '') : '';
    let peer = null;
    if (c.type === 'direct') {
      // From directKey, so a chat with someone who deleted their account still has a peer.
      const peerId = String(c.directKey || '').split(':').filter(id => id !== ctx.userId)[0];
      const u = peerId && users[peerId];
      if (u) {
        // Neither side of a block sees the other's last seen.
        const blockedMe = blocksBy_(peerId).indexOf(ctx.userId) >= 0;
        peer = publicUser_(u, showSeen && !blockedMe && !iBlocked.has(peerId));
        peer.blockedByMe = iBlocked.has(peerId);
        peer.deleted = !!u.deletedAt;
        title = u.displayName;
        avatarFileId = u.avatarFileId;
      }
    }
    return {
      convId: c.convId,
      type: c.type,
      title: title,
      avatarFileId: avatarFileId,
      peer: peer,
      memberCount: members.length,
      myRole: self.role || 'member',
      lastMsgId: c.lastMsgId,
      lastMsgAt: c.lastMsgAt,
      lastMsgSenderId: c.lastMsgSenderId,
      // A preview of a message that has since disappeared is never shown.
      lastMsgPreview: c.lastMsgExpiresAt && c.lastMsgExpiresAt <= nowIso ? '' : openText_(c.lastMsgPreview),
      lastMsgExpiresAt: c.lastMsgExpiresAt || '',
      lastMsgType: c.lastMsgType || '',
      disappearSec: Number(c.disappearSec) || 0,
      ...unreadFor_(c, ctx.userId, self.lastReadMsgId || '')
    };
  }).sort((a, b) => (a.lastMsgAt < b.lastMsgAt ? 1 : -1));

  return { listVersion: listVersion_(ctx.userId), conversations: list };
}

/**
 * Unread messages from others since the reader's last read pointer, counted from the
 * 50-message cache. unreadMore means there may be more than unreadCount.
 */
function unreadFor_(conv, userId, lastRead) {
  if (!conv.lastMsgId || conv.lastMsgId <= lastRead) return { hasUnread: false, unreadCount: 0, unreadMore: false };
  const recent = recentMessages_(conv.convId, false);
  const visible = visibleTo_(conv.convId, userId);
  const count = recent.filter(m => m.msgId > lastRead && m.senderId !== userId && m.type !== 'system' && !m.deleted && visible(m)).length;
  const more = recent.length >= CONFIG.RECENT_CACHE_SIZE && recent[0].msgId > lastRead;
  return { hasUnread: count > 0, unreadCount: count, unreadMore: more && count > 0 };
}

/** Members of a conversation with public profiles, for sender names and group info. */
function listMembers(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  const parts = readAll_(SHEETS.PARTICIPANTS.name).filter(x => x.convId === convId && isActive_(x));
  const ids = new Set(parts.map(x => x.userId));
  const users = {};
  readAll_(SHEETS.USERS.name).forEach(u => { if (ids.has(u.userId)) users[u.userId] = u; });
  const showSeen = securitySettings_().showLastSeen;
  const iBlocked = new Set(blocksBy_(ctx.userId));
  return parts.filter(x => users[x.userId]).map(x => {
    const blockedMe = x.userId !== ctx.userId && blocksBy_(x.userId).indexOf(ctx.userId) >= 0;
    return Object.assign(publicUser_(users[x.userId], showSeen && !blockedMe && !iBlocked.has(x.userId)), { role: x.role, blockedByMe: iBlocked.has(x.userId) });
  });
}

/** Latest message ID of a conversation: cache first, sheet on miss. '' means no messages. */
function getConvLast_(convId) {
  const cached = cache_().get(CK.convLast(convId));
  if (cached !== null) return cached === '-' ? '' : cached;
  const found = findOne_(SHEETS.CONVERSATIONS.name, 'convId', convId);
  const last = found ? found.obj.lastMsgId : '';
  cache_().put(CK.convLast(convId), last || '-', 21600);
  return last;
}

/* ===== Groups.gs ===== */
/**
 * ZeeMoApp back end: group management and per-chat disappearing-message timers.
 *
 * Participant rows are never deleted. Leaving sets role "left", removal sets "removed";
 * re-adding someone reactivates their row with a new joinedAt, so they see only
 * messages sent after they rejoined.
 */

function participant_(convId, userId) {
  return findOne_(SHEETS.PARTICIPANTS.name, 'key', participantKey_(convId, userId));
}

function groupOrThrow_(convId) {
  const conv = findOne_(SHEETS.CONVERSATIONS.name, 'convId', convId);
  if (!conv) throw new ApiError('NOT_FOUND', 'Conversation not found.');
  if (conv.obj.type !== 'group') throw new ApiError('BAD_REQUEST', 'This only works in groups.');
  return conv;
}

function isGroupAdmin_(convId, userId) {
  const p = participant_(convId, userId);
  return !!p && p.obj.role === 'admin';
}

function requireAdmin_(convId, userId) {
  assertParticipant_(convId, userId);
  groupOrThrow_(convId);
  if (!isGroupAdmin_(convId, userId)) throw new ApiError('FORBIDDEN', 'Only group admins can do that.');
}

function activeParticipants_(convId) {
  return readAll_(SHEETS.PARTICIPANTS.name).filter(p => p.convId === convId && isActive_(p));
}

function renameGroup(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  requireAdmin_(convId, ctx.userId);
  const title = requireString_(p.title, 'group name', 60);
  const conv = groupOrThrow_(convId);
  if (conv.obj.title === title) return { convId: convId, title: title };

  withLock_(() => updateFields_(SHEETS.CONVERSATIONS.name, conv.rowIdx, { title: title }));
  invalidateMembership_(activeParticipants_(convId).map(x => x.userId));
  postSystem_(convId, ctx.userId, 'renamed', [], title);
  return { convId: convId, title: title };
}

function addMembers(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  requireAdmin_(convId, ctx.userId);
  const requested = Array.from(new Set((Array.isArray(p.memberIds) ? p.memberIds : []).map(String)));
  if (!requested.length) throw new ApiError('BAD_REQUEST', 'Choose at least one person to add.');

  const known = new Set(readAll_(SHEETS.USERS.name).map(u => u.userId));
  if (requested.some(id => !known.has(id))) throw new ApiError('BAD_REQUEST', 'Some people do not exist.');

  const added = withLock_(() => {
    const active = new Set(activeParticipants_(convId).map(x => x.userId));
    const toAdd = requested.filter(id => !active.has(id));
    if (active.size + toAdd.length > CONFIG.MAX_GROUP_MEMBERS) throw new ApiError('BAD_REQUEST', 'Group is too large.');
    // Take a tick from the message clock so every earlier message is strictly older than the join.
    const now = new Date(tsFromId_(nextMsgId_())).toISOString();
    const fresh = [];
    toAdd.forEach(id => {
      const row = participant_(convId, id);
      if (row) {
        updateFields_(SHEETS.PARTICIPANTS.name, row.rowIdx, {
          role: 'member', joinedAt: now, leftAt: '', lastReadMsgId: '', lastDeliveredMsgId: ''
        });
      } else {
        fresh.push({
          key: participantKey_(convId, id), convId: convId, userId: id, role: 'member',
          joinedAt: now, lastReadMsgId: '', lastDeliveredMsgId: '', leftAt: ''
        });
      }
    });
    appendObjs_(SHEETS.PARTICIPANTS.name, fresh);
    return toAdd;
  });

  if (!added.length) return { convId: convId, added: [] };
  invalidateMembership_(activeParticipants_(convId).map(x => x.userId));
  postSystem_(convId, ctx.userId, 'added', added);
  return { convId: convId, added: added };
}

function removeMember(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  const target = requireString_(p.userId, 'userId', 40);
  requireAdmin_(convId, ctx.userId);
  if (target === ctx.userId) throw new ApiError('BAD_REQUEST', 'Use "Leave group" to remove yourself.');

  const before = activeParticipants_(convId).map(x => x.userId);
  withLock_(() => {
    const row = participant_(convId, target);
    if (!row || !isActive_(row.obj)) throw new ApiError('NOT_FOUND', 'That person is not in this group.');
    updateFields_(SHEETS.PARTICIPANTS.name, row.rowIdx, { role: 'removed', leftAt: nowIso_() });
  });
  invalidateMembership_(before);
  postSystem_(convId, ctx.userId, 'removed', [target]);
  return { convId: convId, removed: target };
}

/** Leave a group. If the last admin leaves, the longest-standing member becomes admin. */
function leaveGroup(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  groupOrThrow_(convId);

  const before = activeParticipants_(convId).map(x => x.userId);
  const promoted = withLock_(() => {
    const row = participant_(convId, ctx.userId);
    updateFields_(SHEETS.PARTICIPANTS.name, row.rowIdx, { role: 'left', leftAt: nowIso_() });
    const rest = activeParticipants_(convId);
    if (!rest.length || rest.some(x => x.role === 'admin')) return '';
    const next = rest.sort((a, b) => (a.joinedAt < b.joinedAt ? -1 : 1))[0];
    const nextRow = participant_(convId, next.userId);
    updateFields_(SHEETS.PARTICIPANTS.name, nextRow.rowIdx, { role: 'admin' });
    return next.userId;
  });

  invalidateMembership_(before);
  postSystem_(convId, ctx.userId, 'left', []);
  if (promoted) postSystem_(convId, ctx.userId, 'promoted', [promoted]);
  return { convId: convId, left: true, promoted: promoted };
}

/** Make someone an admin, or dismiss them as admin. A group always keeps at least one admin. */
function setMemberRole(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  const target = requireString_(p.userId, 'userId', 40);
  const role = p.role === 'admin' ? 'admin' : 'member';
  requireAdmin_(convId, ctx.userId);

  const changed = withLock_(() => {
    const row = participant_(convId, target);
    if (!row || !isActive_(row.obj)) throw new ApiError('NOT_FOUND', 'That person is not in this group.');
    if (row.obj.role === role) return false;
    if (role === 'member' && activeParticipants_(convId).filter(x => x.role === 'admin').length < 2) {
      throw new ApiError('BAD_REQUEST', 'A group needs at least one admin.');
    }
    updateFields_(SHEETS.PARTICIPANTS.name, row.rowIdx, { role: role });
    return true;
  });

  if (changed) {
    invalidateMembership_(activeParticipants_(convId).map(x => x.userId));
    postSystem_(convId, ctx.userId, role === 'admin' ? 'admin' : 'unadmin', [target]);
  }
  return { convId: convId, userId: target, role: role };
}

/**
 * Per-chat disappearing-message timer. Applies to messages sent after the change.
 * Direct chats: either person. Groups: admins only.
 */
function setDisappearing(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  const seconds = Number(p.seconds) || 0;
  if (CONFIG.DISAPPEAR_OPTIONS.indexOf(seconds) === -1) throw new ApiError('BAD_REQUEST', 'Choose one of the timer options.');

  const conv = findOne_(SHEETS.CONVERSATIONS.name, 'convId', convId);
  if (conv.obj.type === 'group' && !isGroupAdmin_(convId, ctx.userId)) {
    throw new ApiError('FORBIDDEN', 'Only group admins can change disappearing messages.');
  }
  if ((Number(conv.obj.disappearSec) || 0) === seconds) return { convId: convId, disappearSec: seconds };

  withLock_(() => updateFields_(SHEETS.CONVERSATIONS.name, conv.rowIdx, { disappearSec: String(seconds) }));
  invalidateMembership_(activeParticipants_(convId).map(x => x.userId));
  postSystem_(convId, ctx.userId, 'disappear', [], seconds);
  return { convId: convId, disappearSec: seconds };
}

/* ===== Messages.gs ===== */
/**
 * ZeeMoApp back end: sending, history, receipts, deletion, disappearing messages
 * and the cache-backed sync endpoint.
 *
 * Read path: a poll first compares the client's cursor with the conversation's
 * latest message ID held in CacheService. Only when they differ does it read
 * messages, and those normally come from a per-conversation cache of the last
 * 50 messages, so most polls never touch the spreadsheet.
 *
 * Every message leaving the server passes visibleTo_(): expired disappearing
 * messages and messages from before the reader joined are never returned.
 */

/** Monotonic message ID. Must be called inside the lock. */
function nextMsgId_() {
  const last = Number(cache_().get(CK.msgClock()) || 0);
  const ts = Math.max(Date.now(), last + 1);
  cache_().put(CK.msgClock(), String(ts), 21600);
  return newId_('m', ts);
}

/** The form clients receive. Decrypts stored text (rows already decrypted pass through unchanged). */
function publicMessage_(m) {
  const deleted = m.deleted === 'TRUE' || m.deleted === true;
  return {
    msgId: m.msgId, convId: m.convId, senderId: m.senderId, type: m.type,
    body: deleted ? '' : (m.type === 'system' ? m.body : openText_(m.body)),
    mediaFileId: deleted ? '' : m.mediaFileId,
    replyTo: m.replyTo, clientId: m.clientId, createdAt: m.createdAt,
    deleted: deleted,
    expiresAt: m.expiresAt || '',
    hiddenFor: m.hiddenFor || ''      // internal; removed by clientMessage_
  };
}

/**
 * What leaves the server. hiddenFor is dropped so a sender can never tell that the
 * recipient has blocked them.
 */
function clientMessage_(m) {
  const o = Object.assign({}, m);
  delete o.hiddenFor;
  return o;
}

function isExpired_(m, nowIso) {
  return !!m.expiresAt && m.expiresAt <= (nowIso || nowIso_());
}

/** Filter for one reader in one conversation. */
function visibleTo_(convId, userId) {
  const since = visibleSince_(convId, userId);
  const nowIso = nowIso_();
  return m => !isExpired_(m, nowIso) && (!since || m.createdAt >= since) && m.hiddenFor !== userId;
}

/* ---------- System notices (group changes, timer changes) ---------- */

/**
 * System message bodies are JSON so each reader's app can say "You" for itself:
 *   { k: kind, a: [actorId, actorName], t: [[userId, name], ...], v: value }
 */
function systemBody_(kind, actorId, targetIds, value) {
  const ids = [actorId].concat(targetIds || []);
  const names = {};
  ids.forEach(id => {
    const u = findOne_(SHEETS.USERS.name, 'userId', id);
    names[id] = u ? u.obj.displayName : 'Someone';
  });
  return JSON.stringify({
    k: kind,
    a: [actorId, names[actorId]],
    t: (targetIds || []).map(id => [id, names[id]]),
    v: value === undefined ? '' : value
  });
}

function timerLabel_(sec) {
  sec = Number(sec) || 0;
  if (!sec) return 'off';
  if (sec % 86400 === 0) {
    const d = sec / 86400;
    return d === 1 ? '24 hours' : d + ' days';
  }
  return Math.round(sec / 3600) + ' hours';
}

/** Plain-text version, used for chat-list previews. */
function systemText_(body) {
  let s;
  try { s = JSON.parse(body); } catch (e) { return ''; }
  const a = s.a ? s.a[1] : 'Someone';
  const t = (s.t || []).map(x => x[1]).join(', ');
  switch (s.k) {
    case 'added': return `${a} added ${t}`;
    case 'removed': return `${a} removed ${t}`;
    case 'left': return `${a} left`;
    case 'renamed': return `${a} changed the group name to "${s.v}"`;
    case 'admin': return `${a} made ${t} an admin`;
    case 'unadmin': return `${a} removed ${t} as admin`;
    case 'promoted': return `${t} is now an admin`;
    case 'photo': return `${a} changed the group photo`;
    case 'photoRemoved': return `${a} removed the group photo`;
    case 'disappear': return Number(s.v) ? `${a} set messages to disappear after ${timerLabel_(s.v)}` : `${a} turned off disappearing messages`;
    default: return '';
  }
}

function previewFor_(m) {
  if (m.type === 'system') return systemText_(m.body);
  if (m.type === 'image') return '📷 Photo';
  if (m.type === 'audio') return '🎤 Voice note';
  return cleanText_(m.body, 80).replace(/\s+/g, ' ');
}

function postSystem_(convId, actorId, kind, targetIds, value) {
  return appendMessage_(convId, actorId, { type: 'system', body: systemBody_(kind, actorId, targetIds, value) });
}

/* ---------- Recent-message cache ---------- */

function putRecent_(convId, msgs) {
  let list = msgs.slice(-CONFIG.RECENT_CACHE_SIZE);
  while (list.length && !cachePutJson_(CK.recent(convId), list, 21600)) list = list.slice(1);
}

/** Last N messages of a conversation from cache, rebuilt from the sheet on a miss. Unfiltered. */
function recentMessages_(convId, rebuild) {
  if (!rebuild) {
    const cached = cacheGetJson_(CK.recent(convId));
    if (cached) return cached;
  }
  const now = Date.now();
  const prevMonth = new Date(now);
  prevMonth.setUTCMonth(prevMonth.getUTCMonth() - 1);
  let rows = [];
  monthRange_(prevMonth.getTime(), now).forEach(k => {
    rows = rows.concat(readConvMessagesInMonth_(convId, k));
  });
  const msgs = rows.slice(-CONFIG.RECENT_CACHE_SIZE).map(publicMessage_);   // decrypt only what is kept
  putRecent_(convId, msgs);
  return msgs;
}

/** Messages strictly after cursor, oldest first, capped at limit. Unfiltered. */
function messagesAfter_(convId, cursor, limit) {
  const last = getConvLast_(convId);
  if (!last || last <= cursor) return { messages: [], hasMore: false };

  let recent = recentMessages_(convId, false);
  // Consistency check: a stale cache written by a racing rebuild is detected and rebuilt.
  if (!recent.length || recent[recent.length - 1].msgId !== last) recent = recentMessages_(convId, true);

  const covered = recent.length && (cursor === '' || cursor >= recent[0].msgId);
  let msgs;
  if (covered) {
    msgs = recent.filter(m => m.msgId > cursor);
  } else if (!cursor) {
    // No cursor and nothing recent: history is fetched with getMessages, not sync.
    msgs = [];
  } else {
    // The client fell behind the cache window: scan month shards from the cursor's month.
    let rows = [];
    monthRange_(tsFromId_(cursor), Date.now()).forEach(k => {
      if (rows.length > limit) return;
      rows = rows.concat(readConvMessagesInMonth_(convId, k).filter(m => m.msgId > cursor));
    });
    msgs = rows.slice(0, limit + 1).map(publicMessage_);
  }
  return { messages: msgs.slice(0, limit), hasMore: msgs.length > limit };
}

/* ---------- Write path ---------- */

/**
 * Append one message and update the conversation and caches.
 * Takes the lock itself: never call it from inside withLock_().
 */
function appendMessage_(convId, senderId, fields) {
  const msg = withLock_(() => {
    const conv = findOne_(SHEETS.CONVERSATIONS.name, 'convId', convId);
    if (!conv) throw new ApiError('NOT_FOUND', 'Conversation not found.');

    const msgId = nextMsgId_();
    const ts = tsFromId_(msgId);
    const timer = Number(conv.obj.disappearSec) || 0;
    const m = {
      msgId: msgId, convId: convId, senderId: senderId, type: fields.type, body: fields.body || '',
      mediaFileId: fields.mediaFileId || '', replyTo: fields.replyTo || '', clientId: fields.clientId || '',
      createdAt: new Date(ts).toISOString(), deleted: 'FALSE',
      // Notices stay; everything else follows the chat's timer at the moment it was sent.
      expiresAt: timer && fields.type !== 'system' ? new Date(ts + timer * 1000).toISOString() : '',
      hiddenFor: fields.hiddenFor || ''
    };

    // The sheet gets encrypted text; notices (names only) stay readable.
    const epoch = epochFor_(m.createdAt, m.expiresAt);
    const stored = Object.assign({}, m, { body: m.type === 'system' ? m.body : sealText_(m.body, epoch) });
    messagesSheet_(monthKey_(ts), true).appendRow(objToRow_(MESSAGE_HEADERS, stored));

    // A message the recipient has blocked must not show up in their chat list.
    if (!m.hiddenFor) {
      updateFields_(SHEETS.CONVERSATIONS.name, conv.rowIdx, {
        lastMsgId: msgId, lastMsgAt: m.createdAt, lastMsgSenderId: senderId,
        lastMsgPreview: m.type === 'system' ? previewFor_(m) : sealText_(previewFor_(m), epoch),
        lastMsgExpiresAt: m.expiresAt, lastMsgType: m.type
      });
    }

    // Cache updates stay inside the lock so concurrent sends cannot overwrite each other.
    const pub = publicMessage_(m);
    const recent = recentMessages_(convId, false).filter(x => x.msgId !== msgId);
    recent.push(pub);
    putRecent_(convId, recent);
    cache_().put(CK.convLast(convId), msgId, 21600);
    return pub;
  });

  advanceReceipt_(convId, senderId, 'lastReadMsgId', msg.msgId);
  if (msg.type !== 'system') queuePush_(convId, senderId, msg);
  return msg;
}

function sendMessage(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);

  const type = MESSAGE_TYPES.indexOf(p.type) >= 0 ? p.type : 'text';
  const body = cleanText_(p.body, CONFIG.MAX_MESSAGE_LENGTH);
  const clientId = cleanText_(p.clientId, 64);
  const replyTo = isMsgId_(p.replyTo) ? p.replyTo : '';
  let mediaFileId = '';

  if (type === 'text' && !body) throw new ApiError('BAD_REQUEST', 'Message is empty.');
  if (type !== 'text') {
    mediaFileId = requireString_(p.mediaFileId, 'mediaFileId', 100);
    const media = findOne_(SHEETS.MEDIA.name, 'fileId', mediaFileId);
    if (!media || media.obj.convId !== convId || media.obj.uploaderId !== ctx.userId) {
      throw new ApiError('BAD_REQUEST', 'Attachment not found.');
    }
  }

  // Idempotency: a retried send with the same clientId returns the original message.
  if (clientId) {
    const dup = cacheGetJson_(CK.dedupe(ctx.userId, clientId));
    if (dup) return dup;
  }

  // Blocking applies to direct chats: you cannot message someone you blocked, and
  // messages to someone who blocked you are kept from them (you are not told).
  let hiddenFor = '';
  const peer = directPeer_(convId, ctx.userId);
  if (peer) {
    if (blocksBy_(ctx.userId).indexOf(peer) >= 0) {
      throw new ApiError('BLOCKED_CONTACT', 'You blocked this person. Unblock them to send messages.');
    }
    if (blocksBy_(peer).indexOf(ctx.userId) >= 0) hiddenFor = peer;
  }

  const msg = clientMessage_(appendMessage_(convId, ctx.userId, {
    type: type, body: body, mediaFileId: mediaFileId, replyTo: replyTo, clientId: clientId, hiddenFor: hiddenFor
  }));
  if (clientId) cachePutJson_(CK.dedupe(ctx.userId, clientId), msg, 600);
  return msg;
}

/**
 * Placeholder for push notifications. Sending FCM requests inline would add
 * 300 to 800 ms to every send, so queue here and flush from a 1-minute time trigger.
 */
function queuePush_(convId, senderId, msg) {
  // TODO: append {convId, senderId, msgId} to a queue and process it in a trigger.
}

/** Older messages for scrolling back, newest page first in time order. */
function getMessages(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  const limit = Math.min(Math.max(Number(p.limit) || 30, 1), 100);
  const before = isMsgId_(p.before) ? p.before : null;
  const visible = visibleTo_(convId, ctx.userId);
  const since = visibleSince_(convId, ctx.userId);

  if (!before) {
    const recent = recentMessages_(convId, false);
    const shown = recent.filter(visible);
    return { messages: shown.slice(-limit).map(clientMessage_), hasMore: recent.length >= limit && shown.length >= limit };
  }

  const floor = since ? Math.max(new Date(since).getTime(), 0) : 0;
  const from = Math.max(tsFromId_(before) - (CONFIG.HISTORY_MAX_MONTHS - 1) * 31 * 86400000, floor);
  const keys = monthRange_(from, tsFromId_(before)).reverse();
  let out = [];
  for (let i = 0; i < keys.length && out.length < limit + 1; i++) {
    const page = readConvMessagesInMonth_(convId, keys[i]).filter(m => m.msgId < before).filter(visible);
    out = page.concat(out);
  }
  return { messages: out.slice(-limit).map(publicMessage_).map(clientMessage_), hasMore: out.length > limit };
}

/* ---------- Delete for everyone ---------- */

function findMessage_(msgId) {
  const sh = messagesSheet_(monthKey_(tsFromId_(msgId)), false);
  return sh ? findOne_(sh.getName(), 'msgId', msgId) : null;
}

/**
 * The sender may delete within DELETE_FOR_EVERYONE_HOURS; a group admin may delete any message.
 * The text or file is erased, leaving a "This message was deleted" notice.
 */
function deleteMessage(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  if (!isMsgId_(p.msgId)) throw new ApiError('BAD_REQUEST', 'Invalid msgId.');

  const found = findMessage_(p.msgId);
  const m = found && found.obj;
  if (!m || m.convId !== convId || !visibleTo_(convId, ctx.userId)(m)) throw new ApiError('NOT_FOUND', 'Message not found.');
  if (m.type === 'system') throw new ApiError('BAD_REQUEST', 'Notices cannot be deleted.');
  if (m.deleted === 'TRUE') return clientMessage_(publicMessage_(m));

  const isSender = m.senderId === ctx.userId;
  const hours = securitySettings_().deleteWindowHours;
  const inWindow = !hours || Date.now() - new Date(m.createdAt).getTime() <= hours * 3600000;
  if (!(isSender && inWindow) && !isGroupAdmin_(convId, ctx.userId)) {
    throw new ApiError('FORBIDDEN', isSender
      ? `Messages can only be deleted for everyone within ${hours} hours of sending.`
      : 'Only the sender or a group admin can delete this message.');
  }

  return clientMessage_(eraseMessage_(found));
}

/**
 * Erase a message's text and file for everyone, leaving "This message was deleted".
 * Used by delete for everyone, report actions and account deletion. found = findMessage_().
 */
function eraseMessage_(found) {
  const m = found.obj;
  const convId = m.convId;
  const sheetName = messagesSheetName_(monthKey_(tsFromId_(m.msgId)));
  const tomb = withLock_(() => {
    updateFields_(sheetName, found.rowIdx, { body: '', mediaFileId: '', deleted: 'TRUE' });
    const pub = publicMessage_(Object.assign({}, m, { deleted: 'TRUE' }));

    const recent = recentMessages_(convId, false);
    const i = recent.map(x => x.msgId).indexOf(m.msgId);
    if (i >= 0) { recent[i] = pub; putRecent_(convId, recent); }

    const conv = findOne_(SHEETS.CONVERSATIONS.name, 'convId', convId);
    if (conv && conv.obj.lastMsgId === m.msgId) {
      updateFields_(SHEETS.CONVERSATIONS.name, conv.rowIdx, { lastMsgPreview: 'This message was deleted' });
    }

    const list = (cacheGetJson_(CK.deletions(convId)) || []).concat([m.msgId]).slice(-200);
    cachePutJson_(CK.deletions(convId), list, 21600);
    cache_().put(CK.deleteVer(convId), newVersion_(), 21600);
    return pub;
  });

  if (m.mediaFileId) trashFile_(m.mediaFileId);
  return tomb;
}


/* ---------- Disappearing-message clean-up (hourly trigger) ---------- */

/**
 * Erases the text and files of expired messages. Rows stay (row positions are pointers);
 * readers already never see expired messages, this removes the content itself.
 */
function purgeExpired() {
  const started = Date.now();
  const nowIso = nowIso_();
  const maxTimer = Math.max.apply(null, CONFIG.DISAPPEAR_OPTIONS);
  const keys = monthRange_(Date.now() - (maxTimer * 1000) - 62 * 86400000, Date.now());
  const touched = {};
  let erased = 0;

  for (let k = 0; k < keys.length; k++) {
    const sh = messagesSheet_(keys[k], false);
    if (!sh) continue;
    const rows = readAll_(sh.getName());
    const bodyCol = MESSAGE_HEADERS.indexOf('body') + 1;          // body and mediaFileId are adjacent
    const delCol = MESSAGE_HEADERS.indexOf('deleted') + 1;
    for (let i = 0; i < rows.length; i++) {
      const m = rows[i];
      if (!isExpired_(m, nowIso) || (m.deleted === 'TRUE' && !m.body && !m.mediaFileId)) continue;
      sh.getRange(i + 2, bodyCol, 1, 2).setValues([['', '']]);
      sh.getRange(i + 2, delCol).setValue('TRUE');
      if (m.mediaFileId) trashFile_(m.mediaFileId);
      touched[m.convId] = true;
      erased++;
      if (Date.now() - started > 270000) break;   // stay well inside the 6-minute limit; the next run continues
    }
  }

  const convSheet = sheet_(SHEETS.CONVERSATIONS.name);
  const previewCol = SHEETS.CONVERSATIONS.headers.indexOf('lastMsgPreview') + 1;
  readAll_(SHEETS.CONVERSATIONS.name).forEach((c, i) => {
    if (c.lastMsgPreview && isExpired_({ expiresAt: c.lastMsgExpiresAt }, nowIso)) {
      convSheet.getRange(i + 2, previewCol).setValue('');
      touched[c.convId] = true;
    }
  });

  Object.keys(touched).forEach(c => cache_().remove(CK.recent(c)));
  const shredded = shredExpiredKeys_();
  console.log(`purgeExpired: erased ${erased} messages in ${Object.keys(touched).length} chats; destroyed ${shredded} expired keys`);
  return { erased: erased, chats: Object.keys(touched).length, keysDestroyed: shredded };
}

/* ---------- Receipts ---------- */

/** Move a read or delivered pointer forward only. Single-cell write, safe outside the lock. */
function advanceReceipt_(convId, userId, field, msgId) {
  const found = findOne_(SHEETS.PARTICIPANTS.name, 'key', participantKey_(convId, userId));
  if (!found || found.obj[field] >= msgId) return false;
  const fields = {};
  fields[field] = msgId;
  if (field === 'lastReadMsgId' && found.obj.lastDeliveredMsgId < msgId) fields.lastDeliveredMsgId = msgId;
  updateFields_(SHEETS.PARTICIPANTS.name, found.rowIdx, fields);
  cache_().put(CK.receiptVer(convId), newVersion_(), 21600);
  return true;
}

function markRead(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  if (!isMsgId_(p.msgId)) throw new ApiError('BAD_REQUEST', 'Invalid msgId.');
  const last = getConvLast_(convId);
  const msgId = p.msgId > last ? last : p.msgId;   // cannot read beyond the latest message
  return { updated: msgId ? advanceReceipt_(convId, ctx.userId, 'lastReadMsgId', msgId) : false };
}

function receiptVersion_(convId) {
  let v = cache_().get(CK.receiptVer(convId));
  if (!v) {
    v = newVersion_();
    cache_().put(CK.receiptVer(convId), v, 21600);
  }
  return v;
}

/* ---------- Sync ---------- */

/**
 * One call returns everything the client needs since its last poll.
 * Request payload:
 *   cursors:          { convId: lastMsgIdSeen }
 *   receiptVersions:  { convId: versionSeen }
 *   deleteVersions:   { convId: versionSeen }
 *   listVersion:      string
 * Response cursors[convId] is the newest message ID scanned, including ones this
 * reader may not see, so the client can move past them.
 */
function sync(ctx, p) {
  const cursors = p.cursors && typeof p.cursors === 'object' ? p.cursors : {};
  const seenReceipts = p.receiptVersions && typeof p.receiptVersions === 'object' ? p.receiptVersions : {};
  const seenDeletes = p.deleteVersions && typeof p.deleteVersions === 'object' ? p.deleteVersions : {};
  const convIds = userConvIds_(ctx.userId);

  const out = {
    serverTime: nowIso_(),
    listVersion: listVersion_(ctx.userId),
    messages: {},
    cursors: {},
    receipts: {},
    deletions: {},
    hasMore: false
  };
  out.listChanged = String(p.listVersion || '') !== out.listVersion;

  let budget = CONFIG.SYNC_MAX_MESSAGES;
  const changedReceipts = [];

  convIds.forEach(convId => {
    const cursor = isMsgId_(cursors[convId]) ? cursors[convId] : '';
    if (budget > 0) {
      const r = messagesAfter_(convId, cursor, budget);
      if (r.messages.length) {
        out.cursors[convId] = r.messages[r.messages.length - 1].msgId;
        budget -= r.messages.length;
        const shown = r.messages.filter(visibleTo_(convId, ctx.userId));
        if (shown.length) {
          out.messages[convId] = shown.map(clientMessage_);
          const newestFromOthers = shown.filter(m => m.senderId !== ctx.userId).pop();
          if (newestFromOthers) advanceReceipt_(convId, ctx.userId, 'lastDeliveredMsgId', newestFromOthers.msgId);
        }
      }
      if (r.hasMore) out.hasMore = true;
    } else {
      out.hasMore = true;
    }

    const ver = receiptVersion_(convId);
    if (String(seenReceipts[convId] || '') !== ver) changedReceipts.push({ convId: convId, ver: ver });

    const dv = cache_().get(CK.deleteVer(convId));
    if (dv && String(seenDeletes[convId] || '') !== dv) {
      out.deletions[convId] = { ver: dv, msgIds: cacheGetJson_(CK.deletions(convId)) || [] };
    }
  });

  if (changedReceipts.length) {
    const wanted = new Set(changedReceipts.map(c => c.convId));
    const parts = readAll_(SHEETS.PARTICIPANTS.name).filter(x => wanted.has(x.convId) && isActive_(x));
    changedReceipts.forEach(c => {
      out.receipts[c.convId] = {
        ver: c.ver,
        members: parts.filter(x => x.convId === c.convId).map(x => ({
          userId: x.userId, lastReadMsgId: x.lastReadMsgId, lastDeliveredMsgId: x.lastDeliveredMsgId
        }))
      };
    });
  }

  touchLastSeen_(ctx.userId);
  return out;
}

/* ===== Media.gs ===== */
/**
 * ZeeMoApp back end: private media in Drive.
 * Files stay private to the script owner and are served only to people allowed to see them:
 *   chat files      members of that chat
 *   profile photos  any signed-in user
 *   group photos    members of that group
 * Compress on the client before upload (images about 1280 px, JPEG quality 0.7; photos 256 px).
 */

function decodeUpload_(p, allowedTypes, maxBytes) {
  const mimeType = String(p.mimeType || '').toLowerCase();
  if (allowedTypes.indexOf(mimeType) === -1) throw new ApiError('BAD_REQUEST', 'File type not allowed.');
  if (typeof p.base64 !== 'string' || !p.base64) throw new ApiError('BAD_REQUEST', 'Missing file data.');
  let bytes;
  try {
    bytes = Utilities.base64Decode(p.base64.replace(/^data:[^,]*,/, ''));
  } catch (e) {
    throw new ApiError('BAD_REQUEST', 'File data is not valid base64.');
  }
  if (bytes.length > maxBytes) {
    const limit = maxBytes >= 1048576 ? `${Math.round(maxBytes / 1048576)} MB` : `${Math.round(maxBytes / 1024)} KB`;
    throw new ApiError('TOO_LARGE', `File is larger than ${limit}.`);
  }
  return { mimeType: mimeType, bytes: bytes };
}

/** Save to the media folder and record it. scope is a convId, "avatar:<userId>" or "group:<convId>". */
function saveFile_(upload, scope, uploaderId) {
  const folderId = prop_('MEDIA_FOLDER_ID');
  if (!folderId) throw new ApiError('NOT_CONFIGURED', 'Run setup() first.');
  const ext = upload.mimeType.split('/')[1].replace('mpeg', 'mp3');
  const name = `${scope.replace(':', '-')}_${newId_('f')}.${ext}`;
  const file = DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(upload.bytes, upload.mimeType, name));
  const record = {
    fileId: file.getId(), convId: scope, uploaderId: uploaderId,
    mimeType: upload.mimeType, bytes: String(upload.bytes.length), createdAt: nowIso_(), trashedAt: ''
  };
  withLock_(() => appendObj_(SHEETS.MEDIA.name, record));
  return record;
}

function uploadMedia(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  const upload = decodeUpload_(p, CONFIG.MEDIA_MIME_TYPES, CONFIG.MEDIA_MAX_BYTES);
  const record = saveFile_(upload, convId, ctx.userId);
  return { fileId: record.fileId, mimeType: upload.mimeType, bytes: upload.bytes.length };
}

/** Returns the file as base64. */
function getMedia(ctx, p) {
  const fileId = requireString_(p.fileId, 'fileId', 100);
  const media = findOne_(SHEETS.MEDIA.name, 'fileId', fileId);
  if (!media) throw new ApiError('NOT_FOUND', 'File not found.');
  const scope = media.obj.convId;
  if (scope.indexOf('group:') === 0) assertParticipant_(scope.substring(6), ctx.userId);
  else if (scope.indexOf('avatar:') !== 0) assertParticipant_(scope, ctx.userId);
  if (media.obj.trashedAt) throw new ApiError('NOT_FOUND', 'This file was deleted.');

  const file = DriveApp.getFileById(fileId);
  if (file.isTrashed()) throw new ApiError('NOT_FOUND', 'This file was deleted.');
  return { fileId: fileId, mimeType: media.obj.mimeType, base64: Utilities.base64Encode(file.getBlob().getBytes()) };
}

/**
 * Delete a file permanently (it skips the Drive bin, where it would stay 30 days), and mark
 * its record so storage figures stay accurate. Falls back to the bin if the Drive API refuses.
 */
function trashFile_(fileId) {
  if (!deletePermanently_(fileId)) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { console.warn('Could not trash ' + fileId, e); }
  }
  const rec = findOne_(SHEETS.MEDIA.name, 'fileId', fileId);
  if (rec && !rec.obj.trashedAt) updateFields_(SHEETS.MEDIA.name, rec.rowIdx, { trashedAt: nowIso_() });
}

/** Drive API delete: the file is removed for good. Uses the script owner's existing Drive permission. */
function deletePermanently_(fileId) {
  try {
    const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId), {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    return code === 204 || code === 200 || code === 404;
  } catch (e) {
    return false;
  }
}

/* ---------- Profile and group photos ---------- */

/** Tell everyone who shares a chat with these users to refresh their chat list. */
function bumpListsOf_(userIds) {
  const wanted = new Set(userIds);
  const convs = new Set();
  const parts = readAll_(SHEETS.PARTICIPANTS.name).filter(isActive_);
  parts.forEach(p => { if (wanted.has(p.userId)) convs.add(p.convId); });
  const people = new Set(userIds);
  parts.forEach(p => { if (convs.has(p.convId)) people.add(p.userId); });
  const ver = newVersion_();
  people.forEach(u => cache_().put(CK.listVer(u), ver, 21600));
}

function setAvatar(ctx, p) {
  const user = getUser_(ctx.userId);
  const old = user.obj.avatarFileId;
  let fileId = '';
  if (!p.remove) {
    const upload = decodeUpload_(p, CONFIG.AVATAR_MIME_TYPES, CONFIG.AVATAR_MAX_BYTES);
    fileId = saveFile_(upload, 'avatar:' + ctx.userId, ctx.userId).fileId;
  }
  updateFields_(SHEETS.USERS.name, user.rowIdx, { avatarFileId: fileId });
  if (old) trashFile_(old);
  bumpListsOf_([ctx.userId]);
  return privateUser_(Object.assign(user.obj, { avatarFileId: fileId }));
}

function setGroupAvatar(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  requireAdmin_(convId, ctx.userId);
  const conv = groupOrThrow_(convId);
  const old = conv.obj.avatarFileId;
  let fileId = '';
  if (!p.remove) {
    const upload = decodeUpload_(p, CONFIG.AVATAR_MIME_TYPES, CONFIG.AVATAR_MAX_BYTES);
    fileId = saveFile_(upload, 'group:' + convId, ctx.userId).fileId;
  } else if (!old) {
    return { convId: convId, avatarFileId: '' };
  }
  withLock_(() => updateFields_(SHEETS.CONVERSATIONS.name, conv.rowIdx, { avatarFileId: fileId }));
  if (old) trashFile_(old);
  invalidateMembership_(activeParticipants_(convId).map(x => x.userId));
  postSystem_(convId, ctx.userId, fileId ? 'photo' : 'photoRemoved', []);
  return { convId: convId, avatarFileId: fileId };
}

/* ===== Security.gs ===== */
/**
 * ZeeMoApp back end: owner security settings, user passcodes and screen lock.
 *
 * The passcode is stored only as an HMAC of a per-user salt and the passcode, keyed
 * with OTP_PEPPER (a Script Property, never in the sheet).
 *
 * Every session is LOCKED unless the server holds an unlock window for it. Unlocking
 * with the passcode opens the window; it closes after the owner's idle time without use
 * (or 6 hours if idle lock is off), when the app locks, and when the token is copied
 * somewhere else and simply left. A sign-in code does not unlock: a new device must
 * also know the passcode. A locked session can call only routes marked lockSafe.
 *
 * Wrong guesses are counted per user for 6 hours. Every PASSCODE_MAX_ATTEMPTS wrong
 * guesses sign that device out; after twice that, guessing stops until the count
 * expires or the owner resets the passcode (resetPasscode in Owner.gs).
 */

function securitySettings_() {
  const d = CONFIG.SECURITY_DEFAULTS;
  const get = k => {
    const v = prop_(k);
    return v === null || String(v).trim() === '' ? d[k] : String(v).trim().toLowerCase();
  };
  const minutes = Math.max(0, Math.floor(Number(get('LOCK_AFTER_MINUTES'))));
  const hours = Math.max(0, Number(get('DELETE_FOR_EVERYONE_HOURS')));
  const action = get('MINIMISE_ACTION');
  return {
    requirePasscode: get('REQUIRE_PASSCODE') !== 'false',
    lockAfterMinutes: isFinite(minutes) ? minutes : Number(d.LOCK_AFTER_MINUTES),
    minimiseAction: CONFIG.MINIMISE_ACTIONS.indexOf(action) >= 0 ? action : d.MINIMISE_ACTION,
    deleteWindowHours: isFinite(hours) ? hours : Number(d.DELETE_FOR_EVERYONE_HOURS),
    showLastSeen: get('SHOW_LAST_SEEN') !== 'false',
    passcodeMinLength: Math.min(6, Math.max(4, Math.floor(Number(get('PASSCODE_MIN_LENGTH'))) || 6)),
    lockOnBlur: get('LOCK_ON_BLUR') === 'true',
    disappearOptions: CONFIG.DISAPPEAR_OPTIONS.slice()
  };
}

function hasPasscode_(userId) {
  const cached = cache_().get(CK.hasPasscode(userId));
  if (cached !== null) return cached === '1';
  const u = findOne_(SHEETS.USERS.name, 'userId', userId);
  const has = !!(u && u.obj.passcode);
  cache_().put(CK.hasPasscode(userId), has ? '1' : '0', 300);   // short, so an owner reset applies quickly
  return has;
}

function passcodeDigest_(salt, passcode) {
  return hmacHex_('pc:' + salt + ':' + passcode);
}

function passcodeRule_() {
  const min = securitySettings_().passcodeMinLength;
  return { min: min, text: min === 6 ? 'Use 6 digits.' : `Use ${min} to 6 digits.` };
}

/** New passcodes must meet the owner's minimum length. */
function checkPasscodeFormat_(passcode) {
  const v = String(passcode === undefined || passcode === null ? '' : passcode);
  const rule = passcodeRule_();
  if (!CONFIG.PASSCODE_PATTERN.test(v) || v.length < rule.min) throw new ApiError('BAD_REQUEST', rule.text);
  return v;
}

/**
 * Check a passcode against the stored value, counting wrong guesses for this user.
 */
function verifyPasscode_(ctx, userRow, passcode) {
  const key = CK.passAttempts(ctx.userId);
  const max = CONFIG.PASSCODE_MAX_ATTEMPTS;
  const before = Number(cache_().get(key) || 0);
  if (before >= max * 2) {
    throw new ApiError('PASSCODE_COOLDOWN', 'Too many wrong passcodes. Try again in a few hours, or ask the administrator to reset your passcode.');
  }
  const parts = String(userRow.passcode || '').split('$');
  const ok = parts.length === 2 && safeEqual_(parts[1], passcodeDigest_(parts[0], String(passcode || '')));
  if (ok) {
    cache_().remove(key);
    return true;
  }
  const attempts = before + 1;
  cache_().put(key, String(attempts), 21600);
  if (attempts % max === 0) {
    logout(ctx);
    throw new ApiError('PASSCODE_LOCKED', 'Too many wrong passcodes. This device has been signed out.');
  }
  const left = max - (attempts % max);
  throw new ApiError('PASSCODE_INVALID', `Wrong passcode. ${left} ${left === 1 ? 'try' : 'tries'} left.`);
}

function unlockWindowSec_() {
  const mins = securitySettings_().lockAfterMinutes;
  return mins > 0 ? Math.min((mins + 2) * 60, 21600) : 21600;
}

function openUnlockWindow_(ctx) {
  cache_().put(CK.unlocked(ctx.tokenHash), '1', unlockWindowSec_());
}

/**
 * Set or change the passcode. Changing always needs the current passcode; someone who
 * has forgotten it asks the owner to run resetPasscode.
 */
function setPasscode(ctx, p) {
  const passcode = checkPasscodeFormat_(p.passcode);
  const user = getUser_(ctx.userId);
  if (user.obj.passcode) {
    if (p.currentPasscode === undefined || p.currentPasscode === '') {
      throw new ApiError('CURRENT_PASSCODE_REQUIRED', 'Enter your current passcode.');
    }
    verifyPasscode_(ctx, user.obj, p.currentPasscode);
  }
  const salt = randomHex_(16);
  updateFields_(SHEETS.USERS.name, user.rowIdx, { passcode: salt + '$' + passcodeDigest_(salt, passcode) });
  cache_().put(CK.hasPasscode(ctx.userId), '1', 300);
  cache_().remove(CK.passAttempts(ctx.userId));
  openUnlockWindow_(ctx);
  return { hasPasscode: true };
}

/** Lock this session now. The app calls it on open, when idle and when minimised. */
function lockSession(ctx) {
  cache_().remove(CK.unlocked(ctx.tokenHash));
  return { locked: hasPasscode_(ctx.userId) };
}

function unlock(ctx, p) {
  const user = getUser_(ctx.userId);
  if (user.obj.passcode) verifyPasscode_(ctx, user.obj, p.passcode);
  openUnlockWindow_(ctx);
  // A passcode shorter than the owner's minimum still opens the app, but must be changed now.
  const tooShort = !!user.obj.passcode && String(p.passcode || '').length < passcodeRule_().min;
  return { unlocked: true, mustChange: tooShort };
}

/**
 * Called by the router before every signed-in action.
 * Passive routes (polling, read receipts, pictures) keep the window open only when the
 * app reports recent taps or typing (payload.active), so an unattended app still locks.
 */
function enforceSecurity_(ctx, route, payload) {
  if (route.lockSafe) return;
  if (!route.termsSafe && !hasAcceptedTerms_(ctx.userId)) throw new ApiError('TERMS_REQUIRED', 'Accept the terms of use to continue.');
  if (!hasPasscode_(ctx.userId)) {
    if (securitySettings_().requirePasscode) throw new ApiError('PASSCODE_NOT_SET', 'Set a passcode to continue.');
    return;
  }
  if (!cache_().get(CK.unlocked(ctx.tokenHash))) throw new ApiError('LOCKED', 'Enter your passcode to continue.');
  if (!route.passive || (payload && payload.active === true)) openUnlockWindow_(ctx);
}

/* ---------- Devices ---------- */

/** This user's signed-in devices, newest first. id is a short, non-secret label. */
function listSessions(ctx) {
  const now = nowIso_();
  return readAll_(SHEETS.SESSIONS.name)
    .filter(s => s.userId === ctx.userId && s.revoked !== 'TRUE' && s.expiresAt > now)
    .map(s => ({ id: s.tokenHash.substring(0, 16), device: s.device, createdAt: s.createdAt, current: s.tokenHash === ctx.tokenHash }))
    .sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));
}

/** { others: true } signs out every other device; { id } signs out one. */
function revokeSessions(ctx, p) {
  if (p.others) return { signedOut: revokeAllSessions_(ctx.userId, ctx.tokenHash) };
  const id = requireString_(p.id, 'id', 16);
  const sh = sheet_(SHEETS.SESSIONS.name);
  const col = SHEETS.SESSIONS.headers.indexOf('revoked') + 1;
  let n = 0;
  readAll_(SHEETS.SESSIONS.name).forEach((s, i) => {
    if (s.userId !== ctx.userId || s.revoked === 'TRUE' || s.tokenHash.substring(0, 16) !== id) return;
    sh.getRange(i + 2, col).setValue('TRUE');
    cache_().remove(CK.session(s.tokenHash));
    cache_().remove(CK.unlocked(s.tokenHash));
    n++;
  });
  if (!n) throw new ApiError('NOT_FOUND', 'That device is already signed out.');
  return { signedOut: n };
}

/* ===== Owner.gs ===== */
/**
 * ZeeMoApp back end: tools for the platform owner.
 *
 *   refreshDashboard()  rewrites the Dashboard sheet (daily trigger; run it any time)
 *   backupDatabase()    copies the database to the Backups folder (weekly trigger)
 *   cleanOldMedia()     deletes chat photos and voice notes older than MEDIA_RETENTION_DAYS
 *   applyBlocksNow()    applies changes in the Users "blocked" column immediately
 *   resetPasscode()     clears the passcode of the person named in RESET_PASSCODE_FOR
 *   encryptExistingMessages()  encrypts rows written before encryption was added
 *   moveToFreshCopy()   moves the database to a new copy that has no version history
 */

function ownerSettings_() {
  const num = (k, min) => {
    const v = Math.floor(Number(prop_(k)));
    return isFinite(v) && v >= min ? v : Number(CONFIG.OWNER_DEFAULTS[k]);
  };
  return { backupKeep: num('BACKUP_KEEP', 1), mediaRetentionDays: num('MEDIA_RETENTION_DAYS', 0) };
}

/** Runs daily: media clean-up, then the dashboard. */
function dailyMaintenance() {
  processDeletionRequests();
  cleanOldMedia();
  refreshDashboard();
}

function applyBlocksNow() {
  cache_().remove(CK.blocked());
  const list = blockList_();
  console.log(`${list.ids.length} blocked account(s). They are signed out on their next request.`);
}

/* ---------- Dashboard ---------- */

function dayKey_(ts) {
  return Utilities.formatDate(new Date(ts), 'Africa/Lagos', 'yyyy-MM-dd');
}

function refreshDashboard() {
  const now = Date.now();
  const day = 86400000;
  const isTrue = v => ['true', 'yes', '1'].indexOf(String(v).trim().toLowerCase()) >= 0;

  const users = readAll_(SHEETS.USERS.name).filter(u => !u.deletedAt);
  const deletedCount = readAll_(SHEETS.USERS.name).length - users.length;
  const reports = readAll_(SHEETS.REPORTS.name);
  const seen = u => new Date(cache_().get(CK.lastSeen(u.userId)) || u.lastSeen || 0).getTime();
  const convs = readAll_(SHEETS.CONVERSATIONS.name);
  const media = readAll_(SHEETS.MEDIA.name).filter(m => !m.trashedAt);
  const mediaBytes = media.reduce((sum, m) => sum + (Number(m.bytes) || 0), 0);

  // Messages per day for the last 30 days (notices excluded).
  const perDay = {};
  for (let i = 0; i < 30; i++) perDay[dayKey_(now - i * day)] = 0;
  let last24 = 0;
  monthRange_(now - 31 * day, now).forEach(k => {
    const sh = messagesSheet_(k, false);
    if (!sh) return;
    readAll_(sh.getName()).forEach(m => {
      if (m.type === 'system' || !m.createdAt) return;
      const t = new Date(m.createdAt).getTime();
      const key = dayKey_(t);
      if (key in perDay) perDay[key]++;
      if (now - t <= day) last24++;
    });
  });
  const days = Object.keys(perDay).sort().reverse();
  const sum = n => days.slice(0, n).reduce((s, k) => s + perDay[k], 0);
  const sec = securitySettings_();
  const own = ownerSettings_();

  const rows = [
    ['ZeeMoApp dashboard', ''],
    ['Updated', Utilities.formatDate(new Date(now), 'Africa/Lagos', 'yyyy-MM-dd HH:mm')],
    ['', ''],
    ['People', ''],
    ['Users', users.length],
    ['Active in the last 24 hours', users.filter(u => now - seen(u) <= day).length],
    ['Active in the last 7 days', users.filter(u => now - seen(u) <= 7 * day).length],
    ['Blocked', users.filter(u => isTrue(u.blocked)).length],
    ['Deleted accounts', deletedCount],
    ['Waiting to be deleted (web requests)', users.filter(u => u.deletionRequestedAt).length],
    ['', ''],
    ['Safety', ''],
    ['Open reports (see the Reports sheet)', reports.filter(r => r.status === 'open').length],
    ['Reports in the last 7 days', reports.filter(r => now - new Date(r.createdAt).getTime() <= 7 * day).length],
    ['', ''],
    ['Chats', ''],
    ['Direct chats', convs.filter(c => c.type === 'direct').length],
    ['Groups', convs.filter(c => c.type === 'group').length],
    ['Messages in the last 24 hours', last24],
    ['Messages in the last 7 days', sum(7)],
    ['Messages in the last 30 days', sum(30)],
    ['Photos and voice notes stored', media.length],
    ['Storage used (MB)', Math.round(mediaBytes / 1048576 * 10) / 10],
    ['', ''],
    ['Settings', ''],
    ['Sign-in method', signinMethods_().join(' and ')],
    ['Passcode required', sec.requirePasscode ? 'Yes' : 'No'],
    ['Lock after (minutes, 0 = off)', sec.lockAfterMinutes],
    ['When minimised', sec.minimiseAction],
    ['Delete for everyone within (hours, 0 = any time)', sec.deleteWindowHours],
    ['Backups kept', own.backupKeep],
    ['Delete media after (days, 0 = never)', own.mediaRetentionDays],
    ['Last backup', prop_('LAST_BACKUP_AT') || 'None yet'],
    ['', ''],
    ['Messages per day', '']
  ].concat(days.slice(0, 14).map(k => [k, perDay[k]]));

  const ss = ss_();
  let sh = ss.getSheetByName('Dashboard');
  if (!sh) sh = ss.insertSheet('Dashboard', 0);
  sh.clear();
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  [1, 4, 10, 19, 29].forEach(r => sh.getRange(r, 1).setFontWeight('bold'));
  try {
    sh.getRange(1, 1).setFontSize(14);
    sh.setColumnWidth(1, 300);
    sh.setColumnWidth(2, 160);
    sh.getCharts().forEach(c => sh.removeChart(c));
    sh.insertChart(sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(30, 1, 14, 2))
      .setOption('title', 'Messages per day (last 14 days)')
      .setOption('legend', { position: 'none' })
      .setPosition(2, 4, 0, 0)
      .build());
  } catch (e) {
    console.warn('Dashboard formatting skipped', e);
  }
  return rows;
}

/* ---------- Weekly backup ---------- */

function backupFolder_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('BACKUP_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { id = null; }
  }
  const folder = DriveApp.createFolder(`${CONFIG.APP_NAME} Backups`);
  props.setProperty('BACKUP_FOLDER_ID', folder.getId());
  return folder;
}

/** Copies the database spreadsheet and keeps only the newest BACKUP_KEEP copies. */
function backupDatabase() {
  const folder = backupFolder_();
  const stamp = Utilities.formatDate(new Date(), 'Africa/Lagos', 'yyyy-MM-dd HHmm');
  DriveApp.getFileById(prop_('SPREADSHEET_ID')).makeCopy(`${CONFIG.APP_NAME} DB backup ${stamp}`, folder);
  PropertiesService.getScriptProperties().setProperty('LAST_BACKUP_AT', stamp);

  const copies = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(`${CONFIG.APP_NAME} DB backup`) === 0) copies.push(f);
  }
  copies.sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime());
  copies.slice(ownerSettings_().backupKeep).forEach(f => { if (!deletePermanently_(f.getId())) f.setTrashed(true); });
  console.log(`Backup saved: ${stamp}. Keeping ${Math.min(copies.length, ownerSettings_().backupKeep)} copies.`);
  return { saved: stamp, kept: Math.min(copies.length, ownerSettings_().backupKeep) };
}

/* ---------- Media retention ---------- */

/** Deletes chat photos and voice notes older than MEDIA_RETENTION_DAYS. Profile and group photos stay. */
function cleanOldMedia() {
  const days = ownerSettings_().mediaRetentionDays;
  if (!days) return { deleted: 0 };
  const cutoff = new Date(Date.now() - days * 86400000).toISOString();
  const started = Date.now();
  let deleted = 0;
  readAll_(SHEETS.MEDIA.name).forEach(m => {
    if (Date.now() - started > 240000) return;
    if (m.trashedAt || m.createdAt >= cutoff) return;
    if (m.convId.indexOf('avatar:') === 0 || m.convId.indexOf('group:') === 0) return;
    trashFile_(m.fileId);
    deleted++;
  });
  console.log(`cleanOldMedia: deleted ${deleted} files older than ${days} days`);
  return { deleted: deleted };
}

/* ---------- Passcode reset (for someone who forgot theirs) ---------- */

/**
 * 1. In Script Properties, add RESET_PASSCODE_FOR with the person's phone number or email.
 * 2. Run resetPasscode.
 * Their passcode is cleared and every device they use is signed out. At their next code
 * sign-in they choose a new passcode. The property is removed afterwards.
 */
function resetPasscode() {
  const props = PropertiesService.getScriptProperties();
  const who = String(props.getProperty('RESET_PASSCODE_FOR') || '').trim();
  if (!who) throw new Error('Add RESET_PASSCODE_FOR (phone number or email) in Script Properties first.');
  const found = who.indexOf('@') >= 0
    ? findOne_(SHEETS.USERS.name, 'email', normaliseEmail_(who))
    : findOne_(SHEETS.USERS.name, 'phone', normalisePhone_(who));
  if (!found) throw new Error('No user with ' + who);

  updateFields_(SHEETS.USERS.name, found.rowIdx, { passcode: '' });
  cache_().remove(CK.hasPasscode(found.obj.userId));
  cache_().remove(CK.passAttempts(found.obj.userId));
  const n = revokeAllSessions_(found.obj.userId, '');
  props.deleteProperty('RESET_PASSCODE_FOR');
  console.log(`Passcode cleared for ${found.obj.displayName}; ${n} device(s) signed out. They set a new one at next sign-in.`);
}

/* ---------- Encryption tools ---------- */

/**
 * Encrypts message text and previews written before encryption was added.
 * Safe to run more than once; it continues where it stopped if it runs out of time.
 */
function encryptExistingMessages() {
  const started = Date.now();
  let done = 0, left = 0;
  const bodyCol = MESSAGE_HEADERS.indexOf('body') + 1;
  ss_().getSheets().filter(sh => sh.getName().indexOf(MESSAGES_PREFIX) === 0).forEach(sh => {
    if (Date.now() - started > 240000) { left++; return; }
    withLock_(() => {
      readAll_(sh.getName()).forEach((m, i) => {
        if (!m.body || m.type === 'system' || isSealed_(m.body)) return;
        sh.getRange(i + 2, bodyCol).setValue(sealText_(m.body, epochFor_(m.createdAt, m.expiresAt)));
        done++;
      });
    });
  });
  withLock_(() => {
    const convSheet = sheet_(SHEETS.CONVERSATIONS.name);
    const col = SHEETS.CONVERSATIONS.headers.indexOf('lastMsgPreview') + 1;
    readAll_(SHEETS.CONVERSATIONS.name).forEach((c, i) => {
      const p = c.lastMsgPreview;
      if (!p || isSealed_(p) || c.lastMsgType === 'system' || p === 'This message was deleted') return;
      convSheet.getRange(i + 2, col).setValue(sealText_(p, epochFor_(c.lastMsgAt || nowIso_(), c.lastMsgExpiresAt)));
      done++;
    });
  });
  console.log(left
    ? `Encrypted ${done} items. Time ran out with ${left} message sheet(s) left: run encryptExistingMessages again.`
    : `Encrypted ${done} items. All message text is now encrypted. Next, run moveToFreshCopy to leave old version history behind.`);
  return { encrypted: done, sheetsLeft: left };
}

/**
 * Google Sheets keeps a version history that still shows text from before it was
 * encrypted or erased. A copy starts with no history. This:
 *   1. copies the database, 2. switches the app to the copy, 3. permanently deletes the
 *   old file, 4. deletes old backups (they may hold unencrypted text) and makes a new one.
 * Run it at a quiet time, after encryptExistingMessages.
 */
function moveToFreshCopy() {
  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty('SPREADSHEET_ID');
  const bound = (() => { try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { return null; } })();
  if (bound && bound.getId() === oldId) throw new Error('This script is attached to the database file, so it cannot replace it.');

  const result = withLock_(() => {
    const oldFile = DriveApp.getFileById(oldId);
    const parents = oldFile.getParents();
    const name = oldFile.getName();
    const copy = parents.hasNext() ? oldFile.makeCopy(name, parents.next()) : oldFile.makeCopy(name);
    const fresh = SpreadsheetApp.openById(copy.getId());
    if (!fresh.getSheetByName(SHEETS.USERS.name)) throw new Error('The copy looks incomplete; nothing was changed.');
    props.setProperty('SPREADSHEET_ID', copy.getId());
    ss_.instance = null;
    const removed = deletePermanently_(oldId);
    if (!removed) oldFile.setTrashed(true);
    return { newId: copy.getId(), oldRemoved: removed };
  });

  // Old backups were made before this point and may contain unencrypted text.
  const folder = backupFolder_();
  const it = folder.getFiles();
  let cleared = 0;
  while (it.hasNext()) {
    const f = it.next();
    if (!deletePermanently_(f.getId())) f.setTrashed(true);
    cleared++;
  }
  backupDatabase();
  refreshDashboard();
  console.log(`Database moved to a fresh copy (${result.newId}). Old file ${result.oldRemoved ? 'deleted permanently' : 'moved to the bin: empty the Drive bin to finish'}. ` +
    `${cleared} old backup(s) removed and a new backup made. Bookmark the new ZeeMoApp DB file.`);
  return result;
}

/* ===== Safety.gs ===== */
/**
 * ZeeMoApp back end: terms of use, blocking, reporting and account deletion.
 * These are what Google Play's User Generated Content and account-deletion policies ask for.
 */

/* ---------- Terms of use ---------- */

function hasAcceptedTerms_(userId) {
  const cached = cache_().get(CK.terms(userId));
  if (cached !== null) return cached === CONFIG.TERMS_VERSION;
  const u = findOne_(SHEETS.USERS.name, 'userId', userId);
  const v = u ? String(u.obj.termsVersion || '') : '';
  cache_().put(CK.terms(userId), v || '-', 300);
  return v === CONFIG.TERMS_VERSION;
}

function acceptTerms(ctx) {
  const user = getUser_(ctx.userId);
  updateFields_(SHEETS.USERS.name, user.rowIdx, { termsVersion: CONFIG.TERMS_VERSION, termsAcceptedAt: nowIso_() });
  cache_().put(CK.terms(ctx.userId), CONFIG.TERMS_VERSION, 300);
  return { termsAccepted: true, version: CONFIG.TERMS_VERSION };
}

/* ---------- Blocking (between users) ---------- */

/** IDs this user has blocked. */
function blocksBy_(userId) {
  const cached = cacheGetJson_(CK.blocksBy(userId));
  if (cached) return cached;
  const ids = readAll_(SHEETS.BLOCKS.name).filter(b => b.blockerId === userId && b.active === 'TRUE').map(b => b.blockedId);
  cachePutJson_(CK.blocksBy(userId), ids, 21600);
  return ids;
}

/** The other person in a direct chat, or '' for a group. */
function directPeer_(convId, userId) {
  const conv = findOne_(SHEETS.CONVERSATIONS.name, 'convId', convId);
  if (!conv || conv.obj.type !== 'direct') return '';
  return conv.obj.directKey.split(':').filter(id => id !== userId)[0] || '';
}

function setBlock_(blockerId, blockedId, active) {
  const key = blockerId + ':' + blockedId;
  withLock_(() => {
    const row = findOne_(SHEETS.BLOCKS.name, 'key', key);
    if (row) updateFields_(SHEETS.BLOCKS.name, row.rowIdx, { active: active ? 'TRUE' : 'FALSE', createdAt: nowIso_() });
    else if (active) appendObj_(SHEETS.BLOCKS.name, { key: key, blockerId: blockerId, blockedId: blockedId, createdAt: nowIso_(), active: 'TRUE' });
  });
  cache_().remove(CK.blocksBy(blockerId));
  cache_().put(CK.listVer(blockerId), newVersion_(), 21600);
}

function blockUser(ctx, p) {
  const target = requireString_(p.userId, 'userId', 40);
  if (target === ctx.userId) throw new ApiError('BAD_REQUEST', 'You cannot block yourself.');
  getUser_(target);
  setBlock_(ctx.userId, target, true);
  return { userId: target, blocked: true };
}

function unblockUser(ctx, p) {
  const target = requireString_(p.userId, 'userId', 40);
  setBlock_(ctx.userId, target, false);
  return { userId: target, blocked: false };
}

function listBlocked(ctx) {
  const ids = new Set(blocksBy_(ctx.userId));
  if (!ids.size) return [];
  return readAll_(SHEETS.USERS.name).filter(u => ids.has(u.userId)).map(u => publicUser_(u, false));
}

/* ---------- Reporting ---------- */

/**
 * Report a message, a person or a group to the owner. The reported message's text is
 * copied into the Reports sheet so the owner can review it. Optionally blocks the person.
 */
function reportContent(ctx, p) {
  const convId = requireString_(p.convId, 'convId', 40);
  assertParticipant_(convId, ctx.userId);
  const reason = CONFIG.REPORT_REASONS.indexOf(p.reason) >= 0 ? p.reason : '';
  if (!reason) throw new ApiError('BAD_REQUEST', 'Choose a reason.');

  const countKey = CK.reportCount(ctx.userId);
  const count = Number(cache_().get(countKey) || 0);
  if (count >= CONFIG.REPORTS_PER_DAY) throw new ApiError('RATE_LIMITED', 'You have sent many reports today. Try again tomorrow.');

  let reportedUserId = '';
  let msgId = '';
  let messageText = '';
  if (p.msgId) {
    if (!isMsgId_(p.msgId)) throw new ApiError('BAD_REQUEST', 'Invalid msgId.');
    const found = findMessage_(p.msgId);
    const m = found && found.obj;
    if (!m || m.convId !== convId || !visibleTo_(convId, ctx.userId)(m) || m.type === 'system') throw new ApiError('NOT_FOUND', 'Message not found.');
    if (m.senderId === ctx.userId) throw new ApiError('BAD_REQUEST', 'You cannot report your own message.');
    msgId = m.msgId;
    reportedUserId = m.senderId;
    const pub = publicMessage_(m);
    messageText = pub.deleted ? '(already deleted)'
      : (pub.type === 'text' ? pub.body : `[${pub.type === 'image' ? 'photo' : 'voice note'}: Drive file ${pub.mediaFileId}] ${pub.body}`);
  } else if (p.userId) {
    reportedUserId = requireString_(p.userId, 'userId', 40);
    if (reportedUserId === ctx.userId) throw new ApiError('BAD_REQUEST', 'You cannot report yourself.');
    const member = readAll_(SHEETS.PARTICIPANTS.name).some(x => x.convId === convId && x.userId === reportedUserId);
    if (!member) throw new ApiError('NOT_FOUND', 'That person is not in this chat.');
  }

  withLock_(() => appendObj_(SHEETS.REPORTS.name, {
    reportId: newId_('r'), createdAt: nowIso_(), reporterId: ctx.userId, reportedUserId: reportedUserId,
    convId: convId, msgId: msgId, reason: reason, details: cleanText_(p.details, 500),
    messageText: cleanText_(messageText, 4000), status: 'open', action: '', handledAt: ''
  }));
  cache_().put(countKey, String(count + 1), 21600);

  let blocked = false;
  if (p.block && reportedUserId) {
    setBlock_(ctx.userId, reportedUserId, true);
    blocked = true;
  }
  return { reported: true, blocked: blocked };
}

/**
 * Owner moderation. In the Reports sheet, type in the "action" cell:
 *   REMOVE   erase the reported message for everyone
 *   BLOCK    block the reported person's account (Users "blocked" = TRUE)
 *   DISMISS  close the report without action
 * Runs hourly; run it yourself to apply straight away.
 */
function applyReportActions() {
  const sh = sheet_(SHEETS.REPORTS.name);
  const h = SHEETS.REPORTS.headers;
  const done = { REMOVE: 0, BLOCK: 0, DISMISS: 0 };
  readAll_(SHEETS.REPORTS.name).forEach((r, i) => {
    const action = String(r.action || '').trim().toUpperCase();
    if (r.status !== 'open' || !(action in done)) return;
    if (action === 'REMOVE' && r.msgId) {
      const found = findMessage_(r.msgId);
      if (found && found.obj.deleted !== 'TRUE') eraseMessage_(found);
    }
    if (action === 'BLOCK' && r.reportedUserId) {
      const u = findOne_(SHEETS.USERS.name, 'userId', r.reportedUserId);
      if (u) updateFields_(SHEETS.USERS.name, u.rowIdx, { blocked: 'TRUE' });
      cache_().remove(CK.blocked());
    }
    sh.getRange(i + 2, h.indexOf('status') + 1).setValue(action === 'DISMISS' ? 'dismissed' : 'done');
    sh.getRange(i + 2, h.indexOf('handledAt') + 1).setValue(nowIso_());
    done[action]++;
  });
  console.log(`Reports handled: ${done.REMOVE} removed, ${done.BLOCK} blocked, ${done.DISMISS} dismissed.`);
  return done;
}

/* ---------- Account deletion ---------- */

/** In the app: delete my account. Needs the passcode if one is set. */
function deleteAccount(ctx, p) {
  if (String(p.confirm || '') !== 'DELETE') throw new ApiError('BAD_REQUEST', 'Confirm by sending confirm: "DELETE".');
  const user = getUser_(ctx.userId);
  if (user.obj.passcode) verifyPasscode_(ctx, user.obj, p.passcode);
  const summary = eraseAccount_(ctx.userId);
  return Object.assign({ deleted: true }, summary);
}

/** A signed-in user cancels a deletion requested on the web page. */
function cancelDeletion(ctx) {
  const user = getUser_(ctx.userId);
  if (user.obj.deletionRequestedAt) updateFields_(SHEETS.USERS.name, user.rowIdx, { deletionRequestedAt: '' });
  return { cancelled: true };
}

/**
 * From the public Delete-account page, without the app.
 * The sign-in code proves the phone number or email. With the passcode too, the account is
 * deleted at once; without it, deletion waits DELETION_GRACE_DAYS so the real owner of the
 * account can cancel by signing in (protects against someone holding a stolen SIM).
 */
function requestDeletion(p) {
  const id = identity_(p);
  verifyCode_(id, p.code);
  const found = findOne_(SHEETS.USERS.name, id.method, id.value);
  if (!found || found.obj.deletedAt) return { status: 'none' };   // say nothing about whether an account exists

  if (found.obj.passcode && p.passcode) {
    // Same wrong-guess limits as the app; token is empty so no device is signed out.
    verifyPasscode_({ userId: found.obj.userId, token: '', tokenHash: '' }, found.obj, p.passcode);
    eraseAccount_(found.obj.userId);
    return { status: 'deleted' };
  }
  if (!found.obj.passcode) {
    eraseAccount_(found.obj.userId);
    return { status: 'deleted' };
  }
  updateFields_(SHEETS.USERS.name, found.rowIdx, { deletionRequestedAt: nowIso_() });
  return { status: 'scheduled', days: CONFIG.DELETION_GRACE_DAYS };
}

/** Daily: carry out web deletion requests whose waiting period has passed. */
function processDeletionRequests() {
  const cutoff = new Date(Date.now() - CONFIG.DELETION_GRACE_DAYS * 86400000).toISOString();
  let n = 0;
  readAll_(SHEETS.USERS.name).forEach(u => {
    if (u.deletionRequestedAt && !u.deletedAt && u.deletionRequestedAt <= cutoff) {
      eraseAccount_(u.userId);
      n++;
    }
  });
  if (n) console.log(`Deleted ${n} account(s) after their waiting period.`);
  return n;
}

/**
 * Delete a person's data: their messages' text and files, their photos and uploads,
 * their profile, passcode, contact details and blocks. They leave every chat and are
 * signed out everywhere. The user row stays as "Deleted account" so chats remain readable
 * for others. Reports they made or received are kept for safety records.
 */
function eraseAccount_(userId) {
  const user = findOne_(SHEETS.USERS.name, 'userId', userId);
  if (!user) throw new ApiError('NOT_FOUND', 'User not found.');
  const started = Date.now();
  let messages = 0;

  // 1. Text and files of everything they sent.
  ss_().getSheets().filter(sh => sh.getName().indexOf(MESSAGES_PREFIX) === 0).forEach(sh => {
    readAll_(sh.getName()).forEach((m, i) => {
      if (m.senderId !== userId || m.type === 'system') return;
      if (m.deleted === 'TRUE' && !m.body && !m.mediaFileId) return;
      eraseMessage_({ rowIdx: i + 2, obj: m });
      messages++;
    });
  });

  // 2. Other uploads (photos never sent, profile photo).
  readAll_(SHEETS.MEDIA.name).forEach(f => { if (f.uploaderId === userId && !f.trashedAt) trashFile_(f.fileId); });

  // 3. Leave every chat.
  const memberships = readAll_(SHEETS.PARTICIPANTS.name).filter(x => x.userId === userId && isActive_(x));
  memberships.forEach(x => {
    const conv = findOne_(SHEETS.CONVERSATIONS.name, 'convId', x.convId);
    if (conv && conv.obj.type === 'group') {
      leaveGroup({ userId: userId }, { convId: x.convId });
    } else {
      const row = participant_(x.convId, userId);
      updateFields_(SHEETS.PARTICIPANTS.name, row.rowIdx, { role: 'left', leftAt: nowIso_() });
      invalidateMembership_(readAll_(SHEETS.PARTICIPANTS.name).filter(y => y.convId === x.convId).map(y => y.userId));
    }
  });

  // 4. Blocks they made or that name them.
  const bsh = sheet_(SHEETS.BLOCKS.name);
  const activeCol = SHEETS.BLOCKS.headers.indexOf('active') + 1;
  readAll_(SHEETS.BLOCKS.name).forEach((b, i) => {
    if ((b.blockerId === userId || b.blockedId === userId) && b.active === 'TRUE') {
      bsh.getRange(i + 2, activeCol).setValue('FALSE');
      cache_().remove(CK.blocksBy(b.blockerId));
    }
  });

  // 5. Profile and sign-in details.
  updateFields_(SHEETS.USERS.name, user.rowIdx, {
    displayName: 'Deleted account', phone: '', email: '', avatarFileId: '', about: '', passcode: '',
    lastSeen: '', lastSeenPrivacy: '', termsVersion: '', termsAcceptedAt: '', deletionRequestedAt: '',
    deletedAt: nowIso_()
  });
  [CK.hasPasscode(userId), CK.passAttempts(userId), CK.lastSeen(userId), CK.terms(userId), CK.userConvs(userId)]
    .forEach(k => cache_().remove(k));

  // 6. Sign out everywhere.
  revokeAllSessions_(userId, '');
  console.log(`Account ${userId} deleted: ${messages} messages erased in ${Math.round((Date.now() - started) / 1000)} s.`);
  return { messagesErased: messages };
}

/* ===== Pages.gs ===== */
/**
 * ZeeMoApp back end: public pages for Google Play.
 *   <web app URL>?page=terms           Terms of use
 *   <web app URL>?page=privacy         Privacy notice
 *   <web app URL>?page=delete-account  Delete your account without the app
 * The texts use OWNER_NAME and CONTACT_EMAIL from Script Properties. Read them through and
 * change them to suit your service before you publish; they describe how this code behaves.
 */

function ownerInfo_() {
  return {
    name: String(prop_('OWNER_NAME') || '').trim() || 'the operator of this service',
    email: String(prop_('CONTACT_EMAIL') || '').trim() || Session.getEffectiveUser().getEmail() || ''
  };
}

function backupWeeks_() {
  const v = Number(prop_('BACKUP_KEEP'));
  return isFinite(v) && v >= 0 ? v : Number(CONFIG.OWNER_DEFAULTS.BACKUP_KEEP);
}

/** Links shown in the app. Empty when the web app URL is not known (for example in tests). */
function legalLinks_() {
  let base = '';
  try { base = ScriptApp.getService().getUrl() || ''; } catch (err) { base = ''; }
  if (!base) return { terms: '', privacy: '', deleteAccount: '' };
  return { terms: base + '?page=terms', privacy: base + '?page=privacy', deleteAccount: base + '?page=delete-account' };
}

/** Sections as [heading, [paragraphs]]. Plain text; escaped when turned into HTML. */
function legalSections_(which) {
  const o = ownerInfo_();
  const app = CONFIG.APP_NAME;
  const weeks = backupWeeks_();
  const reasons = 'spam, harassment or bullying, hate speech, sexual content, scams or fraud';

  if (which === 'terms') {
    return {
      title: `${app} terms of use`,
      version: CONFIG.TERMS_VERSION,
      sections: [
        ['About these terms', [
          `${app} is a messaging service run by ${o.name}. By creating an account or using ${app} you agree to these terms. If you do not agree, do not use the service.`
        ]],
        ['Who can use it', [
          'You must be at least 13 years old, or older if the law where you live requires it.',
          'You must use your own phone number or email address and keep your passcode to yourself.'
        ]],
        ['What you must not do', [
          `There is no tolerance for objectionable content or abusive users. Do not send ${reasons}; threats or content that encourages violence or self-harm; content that sexualises children; anything illegal; or anything that breaks someone else's rights, such as copyright or privacy.`,
          'Do not impersonate others, send bulk or automated messages, or try to break, overload or get around the security of the service.'
        ]],
        ['Reporting and blocking', [
          'You can report a message or a person from inside the app, and block anyone so they can no longer reach you.',
          `Reports go to ${o.name}, who reviews them, normally within 24 hours, and may remove content or close accounts that break these terms, with or without notice.`
        ]],
        ['Your content', [
          `You own what you send. You give ${o.name} permission to store and deliver it only so the service can work.`,
          'You are responsible for what you send.'
        ]],
        ['Closing your account', [
          `You can delete your account at any time in the app (Profile, Delete my account) or on the delete-account page. ${o.name} may suspend or close accounts that break these terms.`
        ]],
        ['No warranty', [
          `The service is provided as it is, without guarantees that it will always be available or free of errors. To the extent the law allows, ${o.name} is not liable for losses that come from using it.`
        ]],
        ['Changes', [
          'These terms may change. When they do, the app asks you to accept the new version before you continue.'
        ]],
        ['Contact', [o.email ? `Questions: ${o.email}` : `Questions: contact ${o.name}.`]]
      ]
    };
  }

  if (which === 'privacy') {
    return {
      title: `${app} privacy notice`,
      version: CONFIG.TERMS_VERSION,
      sections: [
        ['Who we are', [
          `${app} is run by ${o.name}${o.email ? ` (${o.email})` : ''}. This notice explains what the service stores, why, and for how long.`
        ]],
        ['What we store', [
          'Account: your phone number or email address, display name, optional photo and "about" text, and when you were last seen.',
          'Security: a one-way hash of your passcode (the passcode itself is never stored), a hash of each sign-in token, and the device name of each sign-in.',
          'Messages: text, photos and voice notes you send, who sent them, when, and delivery and read status.',
          'Safety: people you block and reports you make, including a copy of a reported message.'
        ]],
        ['Where it is stored', [
          `Data is kept in a Google Sheet and Google Drive folder owned by ${o.name}, on Google's servers. Message text is encrypted in the sheet; the keys are held in the service's settings, so this is not end-to-end encryption.`,
          'All traffic between the app and the service uses HTTPS.'
        ]],
        ['Who else receives data', [
          'Sign-in codes by text message are sent through the SMS provider Termii, which receives your phone number and the code. Sign-in codes by email are sent from a Google account.',
          'Data is not sold, not used for advertising, and not shared with anyone else unless the law requires it.'
        ]],
        ['How long it is kept', [
          'Messages are kept until they are deleted for everyone, or until they disappear if the chat has disappearing messages turned on.',
          `Weekly backups are kept for ${weeks} week${weeks === 1 ? '' : 's'}, so deleted data can remain in a backup for up to that long before it is removed.`,
          'Reports are kept for safety records.'
        ]],
        ['On your device', [
          'The app keeps only your sign-in token on the device. Chats are held in memory and cleared when the app locks or closes.'
        ]],
        ['Deleting your account', [
          'In the app: Profile, then Delete my account. Without the app: use the delete-account page.',
          'Deleting erases your profile, phone number or email, passcode, and the text, photos and voice notes you sent; you leave every chat and are signed out on every device. Chats show you as "Deleted account". Reports you made or received are kept, and copies in backups expire as described above.'
        ]],
        ['Children', [`${app} is not meant for children under 13.`]],
        ['Your rights', [
          `You can ask for a copy of your data, a correction, or deletion${o.email ? ` by writing to ${o.email}` : ''}. You may also complain to your data protection authority; in Nigeria this is the Nigeria Data Protection Commission.`
        ]],
        ['Changes', ['If this notice changes, the new version is posted here.']]
      ]
    };
  }
  return null;
}

/** Route: the texts as JSON, for the app's terms screen. */
function legalText(p) {
  const which = p && p.page === 'privacy' ? 'privacy' : 'terms';
  return Object.assign(legalSections_(which), { links: legalLinks_() });
}

function escHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

const PAGE_STYLE_ =
  'body{font:16px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;margin:0;background:#f4f6fb;color:#151935}' +
  'main{max-width:720px;margin:0 auto;padding:24px 16px 48px}h1{font-size:1.6rem;margin:.2em 0 .6em}h2{font-size:1.1rem;margin:1.6em 0 .4em}' +
  'p{margin:.4em 0}.card{background:#fff;border-radius:14px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08)}' +
  'nav a{margin-right:14px;color:#3b4bd8}label{display:block;margin:14px 0 4px;font-weight:600}' +
  'input,select{width:100%;box-sizing:border-box;padding:12px;border:1px solid #c9cde0;border-radius:10px;font:inherit}' +
  'button{margin-top:16px;padding:12px 18px;border:0;border-radius:10px;background:#3b4bd8;color:#fff;font:inherit;font-weight:600;cursor:pointer}' +
  'button.danger{background:#c62828}button:disabled{opacity:.6}.msg{margin-top:14px;padding:12px;border-radius:10px;background:#eef1ff}' +
  '.err{background:#fdecea;color:#8a1c1c}.muted{color:#5b6080;font-size:.9rem}[hidden]{display:none!important}' +
  '@media(prefers-color-scheme:dark){body{background:#0f1222;color:#e8eaf6}.card{background:#191d33;box-shadow:none}' +
  'input,select{background:#0f1222;color:#e8eaf6;border-color:#3a4060}.msg{background:#232a4d}.err{background:#4a1d1d;color:#ffd7d7}' +
  'nav a{color:#9aa6ff}.muted{color:#a3a8c8}}';

function pageShell_(title, body) {
  const links = legalLinks_();
  const nav = links.terms
    ? `<nav><a href="${escHtml_(links.terms)}" target="_top">Terms</a><a href="${escHtml_(links.privacy)}" target="_top">Privacy</a>` +
      `<a href="${escHtml_(links.deleteAccount)}" target="_top">Delete account</a></nav>`
    : '';
  return HtmlService.createHtmlOutput(
    `<!doctype html><html lang="en-GB"><head><meta charset="utf-8"><title>${escHtml_(title)}</title>` +
    `<style>${PAGE_STYLE_}</style></head><body><main>${nav}${body}</main></body></html>`)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** doGet ?page=terms | privacy | delete-account */
function legalPage_(page) {
  if (page === 'delete-account') return deletePage_();
  const doc = legalSections_(page === 'privacy' ? 'privacy' : 'terms');
  const body = `<div class="card"><h1>${escHtml_(doc.title)}</h1>` +
    `<p class="muted">Version ${escHtml_(doc.version)}</p>` +
    doc.sections.map(s => `<h2>${escHtml_(s[0])}</h2>` + s[1].map(t => `<p>${escHtml_(t)}</p>`).join('')).join('') +
    '</div>';
  return pageShell_(doc.title, body);
}

function deletePage_() {
  const o = ownerInfo_();
  const app = CONFIG.APP_NAME;
  const methods = signinMethods_();
  const days = CONFIG.DELETION_GRACE_DAYS;
  const choose = methods.length > 1
    ? '<label for="method">Sign-in method</label><select id="method"><option value="phone">Phone number</option><option value="email">Email</option></select>'
    : `<input type="hidden" id="method" value="${methods[0]}">`;
  const body = `<div class="card"><h1>Delete your ${escHtml_(app)} account</h1>
<p>You can also do this in the app: Profile, then <b>Delete my account</b>.</p>
<p>Deleting erases your profile, your phone number or email, your passcode, and the text, photos and voice notes you sent. You leave every chat and are signed out on every device. Others see you as "Deleted account". Reports you made or received are kept for safety records, and copies in weekly backups expire after ${backupWeeks_()} weeks.</p>
<p>If you enter your passcode, the account is deleted straight away. Without it, the account is deleted after ${days} days, so the real owner can stop a request made by someone else by opening the app.</p>
<form id="f1">${choose}
<div id="phoneRow"><label for="phone">Phone number</label><input id="phone" type="tel" autocomplete="tel" placeholder="0803 000 0000"></div>
<div id="emailRow" hidden><label for="email">Email</label><input id="email" type="email" autocomplete="email"></div>
<button id="send" type="submit">Send me a code</button></form>
<form id="f2" hidden><label for="code">Code</label><input id="code" inputmode="numeric" autocomplete="one-time-code" maxlength="6">
<label for="passcode">Passcode (optional)</label><input id="passcode" type="password" inputmode="numeric" maxlength="6">
<button id="del" class="danger" type="submit">Delete my account</button></form>
<div id="out" class="msg" role="status" hidden></div>
${o.email ? `<p class="muted">Need help? ${escHtml_(o.email)}</p>` : ''}</div>
<script>
(function(){
  var $=function(id){return document.getElementById(id)};
  var m=$('method');
  function show(){var e=m.value==='email';$('emailRow').hidden=!e;$('phoneRow').hidden=e;}
  if(m.tagName==='SELECT'){m.addEventListener('change',show);} show();
  function who(){var p={method:m.value};if(m.value==='email')p.email=$('email').value;else p.phone=$('phone').value;return p;}
  function say(t,bad){var o=$('out');o.textContent=t;o.className='msg'+(bad?' err':'');o.hidden=false;}
  function call(action,payload,done,btn){
    btn.disabled=true;
    google.script.run.withSuccessHandler(function(raw){btn.disabled=false;var r=JSON.parse(raw);if(r.ok)done(r.data);else say(r.error.message,true);})
      .withFailureHandler(function(){btn.disabled=false;say('Could not reach the service. Try again.',true);})
      .api(JSON.stringify({action:action,payload:payload}));
  }
  $('f1').addEventListener('submit',function(e){e.preventDefault();
    call('requestOtp',who(),function(d){$('f2').hidden=false;$('code').focus();
      say(d.devCode?'Test mode code: '+d.devCode:'We sent a code. It expires in '+Math.round(d.expiresInSec/60)+' minutes.');},$('send'));});
  $('f2').addEventListener('submit',function(e){e.preventDefault();
    if(!confirm('Delete your account? This cannot be undone.'))return;
    var p=who();p.code=$('code').value;p.passcode=$('passcode').value;
    call('requestDeletion',p,function(d){
      $('f1').hidden=true;$('f2').hidden=true;
      if(d.status==='scheduled')say('Request received. The account will be deleted in '+d.days+' days unless it is cancelled in the app.');
      else say('Done. If an account used these details, it has been deleted.');},$('del'));});
})();
</script>`;
  return pageShell_(`Delete your ${app} account`, body);
}

/* ===== Router.gs ===== */
/**
 * ZeeMoApp back end: HTTP entry points.
 *
 * Two ways in, same JSON contract:
 *  - api(raw) via google.script.run, used by the bundled Index.html (recommended).
 *  - doPost for a front end hosted elsewhere: a POST with Content-Type text/plain and a JSON body:
 *   { "action": "sync", "token": "<session token>", "payload": { ... } }
 * text/plain avoids the CORS preflight that Apps Script web apps cannot answer.
 *
 * Every response is JSON:
 *   { "ok": true,  "data": ... }
 *   { "ok": false, "error": { "code": "...", "message": "..." } }
 * Apps Script cannot set HTTP status codes, so clients must check "ok".
 *
 * Handlers are wrapped in arrow functions so names resolve at call time,
 * regardless of the order in which Apps Script loads the files.
 */

/*
 * auth:     needs a signed-in session
 * lockSafe: allowed while the screen is locked or before a required passcode is set
 * termsSafe: allowed before the terms of use are accepted
 * passive:  automatic calls; they keep the unlock window open only with payload.active
 */
const ROUTES = {
  ping:              { auth: false, fn: () => ({ pong: true, serverTime: nowIso_() }) },
  authConfig:        { auth: false, fn: () => authConfig() },
  requestOtp:        { auth: false, fn: (ctx, p) => requestOtp(p) },
  verifyOtp:         { auth: false, fn: (ctx, p) => verifyOtp(p) },
  legalText:         { auth: false, fn: (ctx, p) => legalText(p) },
  requestDeletion:   { auth: false, fn: (ctx, p) => requestDeletion(p) },

  logout:            { auth: true, lockSafe: true, fn: (ctx) => logout(ctx) },
  me:                { auth: true, lockSafe: true, fn: (ctx) => me(ctx) },
  setPasscode:       { auth: true, lockSafe: true, fn: (ctx, p) => setPasscode(ctx, p) },
  lock:              { auth: true, lockSafe: true, fn: (ctx) => lockSession(ctx) },
  unlock:            { auth: true, lockSafe: true, fn: (ctx, p) => unlock(ctx, p) },
  acceptTerms:       { auth: true, lockSafe: true, fn: (ctx) => acceptTerms(ctx) },
  listSessions:      { auth: true, fn: (ctx) => listSessions(ctx) },
  revokeSessions:    { auth: true, fn: (ctx, p) => revokeSessions(ctx, p) },

  updateProfile:     { auth: true, fn: (ctx, p) => updateProfile(ctx, p) },
  setAvatar:         { auth: true, fn: (ctx, p) => setAvatar(ctx, p) },
  findUser:          { auth: true, fn: (ctx, p) => findUser(ctx, p) },
  findUserByPhone:   { auth: true, fn: (ctx, p) => findUserByPhone(ctx, p) },

  listConversations: { auth: true, passive: true, fn: (ctx) => listConversations(ctx) },
  createDirect:      { auth: true, fn: (ctx, p) => createDirect(ctx, p) },
  createGroup:       { auth: true, fn: (ctx, p) => createGroup(ctx, p) },
  listMembers:       { auth: true, passive: true, fn: (ctx, p) => listMembers(ctx, p) },
  renameGroup:       { auth: true, fn: (ctx, p) => renameGroup(ctx, p) },
  addMembers:        { auth: true, fn: (ctx, p) => addMembers(ctx, p) },
  removeMember:      { auth: true, fn: (ctx, p) => removeMember(ctx, p) },
  leaveGroup:        { auth: true, fn: (ctx, p) => leaveGroup(ctx, p) },
  setMemberRole:     { auth: true, fn: (ctx, p) => setMemberRole(ctx, p) },
  setDisappearing:   { auth: true, fn: (ctx, p) => setDisappearing(ctx, p) },
  setGroupAvatar:    { auth: true, fn: (ctx, p) => setGroupAvatar(ctx, p) },

  sendMessage:       { auth: true, fn: (ctx, p) => sendMessage(ctx, p) },
  getMessages:       { auth: true, fn: (ctx, p) => getMessages(ctx, p) },
  deleteMessage:     { auth: true, fn: (ctx, p) => deleteMessage(ctx, p) },
  markRead:          { auth: true, passive: true, fn: (ctx, p) => markRead(ctx, p) },
  sync:              { auth: true, passive: true, fn: (ctx, p) => sync(ctx, p) },

  uploadMedia:       { auth: true, fn: (ctx, p) => uploadMedia(ctx, p) },
  getMedia:          { auth: true, passive: true, fn: (ctx, p) => getMedia(ctx, p) },

  blockUser:         { auth: true, fn: (ctx, p) => blockUser(ctx, p) },
  unblockUser:       { auth: true, fn: (ctx, p) => unblockUser(ctx, p) },
  listBlocked:       { auth: true, fn: (ctx) => listBlocked(ctx) },
  reportContent:     { auth: true, fn: (ctx, p) => reportContent(ctx, p) },
  deleteAccount:     { auth: true, termsSafe: true, fn: (ctx, p) => deleteAccount(ctx, p) },
  cancelDeletion:    { auth: true, termsSafe: true, fn: (ctx) => cancelDeletion(ctx) }
};

/** HTTP entry point for a separately hosted front end. */
function doPost(e) {
  const raw = e && e.postData ? e.postData.contents : '';
  return json_(handle_(raw));
}

/**
 * Entry point for the bundled Index.html, called with google.script.run.
 * Takes and returns JSON strings, so the response shape is identical to doPost.
 */
function api(raw) {
  return JSON.stringify(handle_(raw));
}

/**
 * Serves the app. ?health=1 returns a JSON status check instead.
 * Requires an HTML file named "Index" in the project.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.health) {
    return json_({ ok: true, data: { app: CONFIG.APP_NAME, status: 'up', serverTime: nowIso_() } });
  }
  if (e && e.parameter && e.parameter.page) return legalPage_(String(e.parameter.page));
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ZeeMo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function handle_(raw) {
  let action = 'unknown';
  try {
    const req = parseRequest_(raw);
    action = req.action;

    const route = Object.prototype.hasOwnProperty.call(ROUTES, action) ? ROUTES[action] : null;
    if (!route) throw new ApiError('UNKNOWN_ACTION', `Unknown action "${action}".`);

    ensureSchema_();
    const ctx = { token: req.token };
    if (route.auth) {
      ctx.userId = requireSession_(req.token);
      ctx.tokenHash = sha256Hex_(req.token);
    }
    const payload = req.payload && typeof req.payload === 'object' ? req.payload : {};
    if (route.auth) enforceSecurity_(ctx, route, payload);
    return { ok: true, data: route.fn(ctx, payload) };
  } catch (err) {
    const known = err instanceof ApiError;
    if (!known) console.error(`[${action}]`, err && err.stack ? err.stack : err);
    return {
      ok: false,
      error: {
        code: known ? err.code : 'SERVER_ERROR',
        message: known ? err.message : 'Unexpected server error.'
      }
    };
  }
}

function parseRequest_(raw) {
  if (typeof raw !== 'string' || !raw) throw new ApiError('BAD_REQUEST', 'Empty request.');
  if (raw.length > CONFIG.MAX_REQUEST_CHARS) throw new ApiError('TOO_LARGE', 'Request too large.');
  let req;
  try {
    req = JSON.parse(raw);
  } catch (err) {
    throw new ApiError('BAD_REQUEST', 'Request body must be JSON.');
  }
  if (!req || typeof req.action !== 'string') throw new ApiError('BAD_REQUEST', 'Missing action.');
  return req;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ===== Setup.gs ===== */
/**
 * ZeeMoApp back end: one-time setup. Run setup() from the editor once, then deploy.
 * Safe to run again: it only creates what is missing.
 */

function setup() {
  const props = PropertiesService.getScriptProperties();

  let ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) {
    ssId = SpreadsheetApp.create(`${CONFIG.APP_NAME} DB`).getId();
    props.setProperty('SPREADSHEET_ID', ssId);
  }
  const ss = SpreadsheetApp.openById(ssId);
  ss_.instance = ss;
  ss.setSpreadsheetTimeZone('Africa/Lagos');

  upgradeSheets_();
  props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION);
  messagesSheet_(monthKey_(Date.now()), true);

  const placeholder = ss.getSheetByName('Sheet1');
  if (placeholder && ss.getSheets().length > 1) ss.deleteSheet(placeholder);

  if (!props.getProperty('MEDIA_FOLDER_ID')) {
    props.setProperty('MEDIA_FOLDER_ID', DriveApp.createFolder(`${CONFIG.APP_NAME} Media`).getId());
  }
  // Security settings appear in Script Properties so the owner can see and edit them.
  [CONFIG.SECURITY_DEFAULTS, CONFIG.OWNER_DEFAULTS].forEach(defaults => Object.keys(defaults).forEach(k => {
    if (props.getProperty(k) === null) props.setProperty(k, defaults[k]);
  }));
  if (!props.getProperty('SIGNIN_METHOD')) props.setProperty('SIGNIN_METHOD', CONFIG.DEFAULT_SIGNIN_METHOD);
  if (!props.getProperty('OTP_PEPPER')) props.setProperty('OTP_PEPPER', randomHex_(64));

  installTriggers();
  refreshDashboard();

  console.log('Spreadsheet: ' + ss.getUrl());
  console.log('Media folder: ' + props.getProperty('MEDIA_FOLDER_ID'));
  console.log('Sign-in method: ' + props.getProperty('SIGNIN_METHOD') + ' (change with setSigninMethodPhone/Email/Both).');
  console.log('Security: ' + JSON.stringify(securitySettings_()));
  console.log('Next: set TERMII_API_KEY for phone sign-in (or DEV_MODE=true for testing), then Deploy > New deployment > Web app.');
}

/** Pre-create next month's shard so the first message of a month does not pay the creation cost. */
function createNextMonthShard() {
  const d = new Date();
  d.setUTCDate(1);
  d.setUTCMonth(d.getUTCMonth() + 1);
  withLock_(() => messagesSheet_(monthKey_(d.getTime()), true));
}

/**
 * Installs all scheduled jobs. Safe to re-run.
 *   purgeExpired          hourly    erase expired disappearing messages
 *   applyReportActions    hourly    carry out REMOVE / BLOCK / DISMISS typed in the Reports sheet
 *   dailyMaintenance      03:00     web deletion requests, media clean-up and dashboard
 *   backupDatabase        Sunday    weekly copy of the database
 *   createNextMonthShard  25th      next month's message sheet
 */
function installTriggers() {
  const wanted = ['createNextMonthShard', 'purgeExpired', 'dailyMaintenance', 'backupDatabase', 'applyReportActions'];
  ScriptApp.getProjectTriggers()
    .filter(t => wanted.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('createNextMonthShard').timeBased().onMonthDay(25).atHour(2).create();
  ScriptApp.newTrigger('purgeExpired').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('applyReportActions').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyMaintenance').timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger('backupDatabase').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
}

/** Logs the current security settings. Edit them in Project Settings > Script Properties. */
function showSecuritySettings() {
  console.log(JSON.stringify(securitySettings_(), null, 2));
}

function setDevMode(on) {
  PropertiesService.getScriptProperties().setProperty('DEV_MODE', on ? 'true' : 'false');
}

/* Sign-in method switches: select one in the editor and click Run. Takes effect immediately. */
function setSigninMethodPhone() { setSigninMethod_('phone'); }
function setSigninMethodEmail() { setSigninMethod_('email'); }
function setSigninMethodBoth() { setSigninMethod_('both'); }

function setSigninMethod_(method) {
  PropertiesService.getScriptProperties().setProperty('SIGNIN_METHOD', method);
  console.log('Sign-in method is now: ' + method);
  if (method !== 'email' && !prop_('TERMII_API_KEY') && prop_('DEV_MODE') !== 'true') {
    console.warn('Phone sign-in needs TERMII_API_KEY in Script Properties before real users can join.');
  }
}

  return {
    api, setup, setDevMode, setSigninMethodBoth, purgeExpired,
    sendMessage, markRead, createGroup, createDirect, setDisappearing,
    withLock_, appendObj_, findOne_, readAll_, newId_, nowIso_, cache_, CK, SHEETS, ss_
  };
}

/* ---------- Demo contacts ---------- */
const BOTS = [
  { phone: '+2348030000101', displayName: 'Hauwa Musa', about: 'Head of Department, Sciences' },
  { phone: '+2348030000102', displayName: 'Tunde Bakare', about: 'Cooperative secretary' },
  { phone: '+2348030000103', displayName: 'Mrs Ngozi Okafor', about: 'Exams officer' },
  { phone: '+2348030000104', displayName: 'Ibrahim Sule', about: 'ICT unit' }
];
const REPLIES = {
  'Hauwa Musa': ['Noted. I will bring the practical scores tomorrow morning.', 'Thank you. Please share the timetable when it is ready.', 'Okay, see you at the staff meeting.'],
  'Tunde Bakare': ['Received. The contribution list is updated.', 'I will confirm with the treasurer and get back to you.', 'Alright sir, well done.'],
  'Mrs Ngozi Okafor': ['Good. The question papers must be submitted by Friday.', 'Please remind your class teachers.'],
  'Ibrahim Sule': ['The lab computers are ready for the CBT practice.', 'I have reset the password. Try again now.']
};

export function createMockTransport() {
  /* Persistent state: sheets, properties and cache (media files live in memory only). */
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch { saved = null; }
  try { localStorage.removeItem('zm.demo.db.v1'); } catch { /* old demo data */ }

  const ss = new DemoSpreadsheet(saved ? saved.sheets.map(s => new DemoSheet(s.name, s.data, s.maxCols)) : null);
  const props = saved ? saved.props : {};
  const cacheStore = saved ? saved.cache : {};
  const files = {};
  let clockOffset = 0;

  class DemoDate extends Date {
    constructor(...a) { if (a.length) super(...a); else super(Date.now() + clockOffset); }
    static now() { return Date.now() + clockOffset; }
  }

  const cache = {
    get(k) {
      const e = cacheStore[k];
      if (!e) return null;
      if (e.x && e.x < Date.now()) { delete cacheStore[k]; return null; }
      return e.v;
    },
    put(k, v, ttl) { cacheStore[k] = { v: String(v), x: Date.now() + Math.min(ttl || 600, 21600) * 1000 }; },
    remove(k) { delete cacheStore[k]; }
  };

  let uuidCounter = 0;
  const uuid = () => (crypto.randomUUID ? crypto.randomUUID()
    : Array.from(crypto.getRandomValues(new Uint8Array(16)), b => b.toString(16).padStart(2, '0')).join('') + (uuidCounter++));

  const G = {
    SpreadsheetApp: { create: () => ss, openById: () => ss, flush() {}, getActiveSpreadsheet: () => null },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (Object.prototype.hasOwnProperty.call(props, k) ? props[k] : null),
      setProperty: (k, v) => { props[k] = String(v); },
      getProperties: () => Object.assign({}, props),
      deleteProperty: k => { delete props[k]; }
    }) },
    CacheService: { getScriptCache: () => cache },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: {
      getUuid: uuid,
      DigestAlgorithm: { SHA_256: 'sha256' },
      Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, text) => signed(sha256(utf8(text))),
      computeHmacSha256Signature: (text, key) => signed(hmacSha256(utf8(key), utf8(text))),
      formatDate: (d, tz, f) => {
        const iso = d.toISOString();
        return f.replace('yyyy', iso.slice(0, 4)).replace('MM', iso.slice(5, 7)).replace('dd', iso.slice(8, 10))
          .replace('HH', iso.slice(11, 13)).replace('mm', iso.slice(14, 16));
      },
      base64Decode: s => signed(Uint8Array.from(atob(s), c => c.charCodeAt(0))),
      base64Encode: b => {
        let bin = '';
        for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.slice(i, i + 0x8000).map(x => (x + 256) % 256));
        return btoa(bin);
      },
      newBlob: (bytes, mime, name) => ({ bytes, mime, name })
    },
    DriveApp: {
      createFolder: () => ({ getId: () => 'demo-folder' }),
      getFolderById: () => ({ createFile: blob => { const id = 'demo-file-' + uuid().slice(0, 12); files[id] = blob; return { getId: () => id }; } }),
      getFileById: id => {
        const f = files[id];
        if (!f) throw new Error('Demo files are kept only until the page is reloaded.');
        return { getBlob: () => ({ getBytes: () => f.bytes }), isTrashed: () => !!f.trashed, setTrashed: v => { f.trashed = v; } };
      }
    },
    MailApp: { getRemainingDailyQuota: () => 100, sendEmail() {} },
    UrlFetchApp: { fetch() { throw new Error('No network in the demo.'); } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType: () => s }) },
    HtmlService: {},
    ScriptApp: (() => {
      const chain = { timeBased: () => chain, onMonthDay: () => chain, atHour: () => chain, everyHours: () => chain, everyDays: () => chain, onWeekDay: () => chain, create() {} };
      return { WeekDay: { SUNDAY: 'SUNDAY' }, getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => chain, getOAuthToken: () => '', getService: () => ({ getUrl: () => '' }) };
    })(),
    Session: { getEffectiveUser: () => ({ getEmail: () => '' }) },
    Date: DemoDate
  };

  const B = loadBackend(G.SpreadsheetApp, G.PropertiesService, G.CacheService, G.LockService, G.Utilities, G.DriveApp,
    G.MailApp, G.UrlFetchApp, G.ContentService, G.HtmlService, G.ScriptApp, G.Session, G.Date);

  const quiet = fn => {
    const log = console.log, warn = console.warn;
    console.log = () => {}; console.warn = () => {};
    try { return fn(); } finally { console.log = log; console.warn = warn; }
  };

  function save() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({
        sheets: ss.sheets.map(s => ({ name: s.name, data: s.data, maxCols: s.maxCols })),
        props, cache: cacheStore
      }));
    } catch { /* storage full or blocked: the demo still works for this visit */ }
  }

  /* First run: create the database, demo settings and contacts. */
  if (!saved) {
    quiet(() => B.setup());
    B.setDevMode(true);
    quiet(() => B.setSigninMethodBoth());
    B.withLock_(() => BOTS.forEach(b => B.appendObj_(B.SHEETS.USERS.name, {
      userId: B.newId_('u'), phone: b.phone, email: '', displayName: b.displayName, avatarFileId: '', about: b.about,
      createdAt: B.nowIso_(), lastSeen: B.nowIso_(), passcode: ''
    })));
    props.DEMO_SEEDED = '';
    save();
  }

  const botByName = name => B.findOne_(B.SHEETS.USERS.name, 'displayName', name).obj;
  const isBot = userId => BOTS.some(b => b.displayName === (B.findOne_(B.SHEETS.USERS.name, 'userId', userId) || { obj: {} }).obj.displayName);
  const as = userId => ({ userId });

  /** A believable starting state for each new demo user, written in the past. */
  function seedFor(me) {
    const seeded = (props.DEMO_SEEDED || '').split(',');
    if (seeded.includes(me)) return;
    props.DEMO_SEEDED = seeded.concat(me).filter(Boolean).join(',');

    const h = botByName('Hauwa Musa').userId, t = botByName('Tunde Bakare').userId;
    const n = botByName('Mrs Ngozi Okafor').userId, i = botByName('Ibrahim Sule').userId;
    const at = (minAgo, fn) => { clockOffset = -minAgo * 60000; try { return fn(); } finally { clockOffset = 0; } };
    cache.remove('mts');   // let seeded messages take past timestamps

    const coop = at(1560, () => B.createGroup(as(t), { title: 'Staff Cooperative', memberIds: [me, h] }).convId);
    at(1500, () => B.sendMessage(as(t), { convId: coop, body: 'Reminder: monthly contributions close on the 25th.' }));
    at(1460, () => B.sendMessage(as(h), { convId: coop, body: 'Paid. Please confirm.' }));

    const staff = at(200, () => B.createGroup(as(n), { title: 'Staff Room', memberIds: [me, h, i] }).convId);
    at(190, () => B.sendMessage(as(n), { convId: staff, body: 'Good morning, colleagues. Second term continuous assessment scores are due on Friday.' }));
    at(176, () => B.sendMessage(as(i), { convId: staff, body: 'The ICT lab is free on Wednesday afternoon for anyone who needs to enter scores.' }));
    at(160, () => B.sendMessage(as(h), { convId: staff, body: 'Thank you. Sciences will use the Wednesday slot.' }));

    const dh = at(100, () => B.createDirect(as(h), { otherUserId: me }).convId);
    at(95, () => B.sendMessage(as(h), { convId: dh, body: 'Salam. Did you get the revised scheme of work?' }));
    const mine = at(90, () => B.sendMessage(as(me), { convId: dh, body: 'Yes, I have it. I will adjust week 6 and send it back.' }));
    at(89, () => B.markRead(as(h), { convId: dh, msgId: mine.msgId }));
    at(12, () => B.sendMessage(as(h), { convId: dh, body: 'Perfect. Please do before the HOD meeting.' }));

    // Older chats are read; the latest direct message stays unread.
    [coop, staff].forEach(c => {
      const last = B.findOne_(B.SHEETS.CONVERSATIONS.name, 'convId', c).obj.lastMsgId;
      B.markRead(as(me), { convId: c, msgId: last });
    });
    cache.remove('mts');
  }

  /** Demo contacts read and answer. */
  function afterSend(senderId, msg) {
    const conv = B.findOne_(B.SHEETS.CONVERSATIONS.name, 'convId', msg.convId);
    if (!conv) return;
    const members = B.readAll_(B.SHEETS.PARTICIPANTS.name)
      .filter(p => p.convId === msg.convId && p.userId !== senderId && (p.role === 'member' || p.role === 'admin'));
    const bots = members.filter(p => isBot(p.userId));
    bots.forEach((p, idx) => {
      setTimeout(() => { try { B.markRead(as(p.userId), { convId: msg.convId, msgId: msg.msgId }); save(); } catch { /* removed meanwhile */ } }, 1800 + idx * 500);
    });
    if (conv.obj.type === 'direct' && bots.length && msg.type !== 'system') {
      const bot = B.findOne_(B.SHEETS.USERS.name, 'userId', bots[0].userId).obj;
      const lines = REPLIES[bot.displayName] || ['Okay.'];
      const count = Number(props['DEMO_REPLIES_' + bot.userId] || 0);
      setTimeout(() => {
        try {
          B.sendMessage(as(bot.userId), { convId: msg.convId, body: lines[count % lines.length] });
          props['DEMO_REPLIES_' + bot.userId] = String(count + 1);
          save();
        } catch { /* chat changed meanwhile */ }
      }, 4200);
    }
  }

  function botsOnline() {
    BOTS.forEach(b => {
      const u = B.findOne_(B.SHEETS.USERS.name, 'phone', b.phone);
      if (u) cache.put(B.CK.lastSeen(u.obj.userId), B.nowIso_(), 21600);
    });
  }

  quiet(() => B.purgeExpired());
  setInterval(() => { quiet(() => B.purgeExpired()); save(); }, 10 * 60000);

  return async function transport(action, token, payload) {
    await new Promise(r => setTimeout(r, 120 + Math.random() * 240));   // network latency
    botsOnline();
    const out = JSON.parse(quiet(() => B.api(JSON.stringify({ action, token, payload }))));
    if (out.ok && action === 'verifyOtp') seedFor(out.data.user.userId);
    if (out.ok && action === 'sendMessage') afterSend(out.data.senderId, out.data);
    save();
    return out;
  };
}
