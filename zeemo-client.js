/**
 * ZeeMoApp browser client for the Apps Script API.
 * Host the PWA on your own origin and import this module.
 *
 *   import { ZeeMoClient } from './zeemo-client.js';
 *   const api = new ZeeMoClient('https://script.google.com/macros/s/DEPLOYMENT_ID/exec');
 *   api.addEventListener('messages', e => render(e.detail));
 *   api.start();
 *
 * Events: messages, receipts, deletions, edits, listchanged, pending, sent, failed,
 *         signedout, locked, needpasscode, needterms, synced, syncerror
 * Pass { transport } to replace HTTP (used by the in-browser demo and google.script.run).
 *
 * Privacy: only the session token is saved on the device. Cursors and versions are
 * kept in memory and rebuilt after each start, and message content is never stored.
 */

/** A readable device name such as "Chrome on Android", saved with each sign-in. */
export function deviceLabel(ua) {
  ua = String(ua || (typeof navigator !== 'undefined' ? navigator.userAgent : ''));
  const browser = /Edg\//.test(ua) ? 'Edge' : /OPR\/|Opera/.test(ua) ? 'Opera' : /SamsungBrowser/.test(ua) ? 'Samsung Internet'
    : /Firefox\//.test(ua) ? 'Firefox' : /Chrome\/|CriOS/.test(ua) ? 'Chrome' : /Safari\//.test(ua) ? 'Safari' : 'Browser';
  const os = /Android/.test(ua) ? 'Android' : /iPhone/.test(ua) ? 'iPhone' : /iPad/.test(ua) ? 'iPad' : /CrOS/.test(ua) ? 'Chromebook'
    : /Windows/.test(ua) ? 'Windows' : /Mac OS X|Macintosh/.test(ua) ? 'Mac' : /Linux/.test(ua) ? 'Linux' : '';
  return os ? `${browser} on ${os}` : browser;
}

export class ZeeMoError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

const store = {
  get(key, fallback) {
    try { const v = localStorage.getItem(key); return v === null ? fallback : JSON.parse(v); }
    catch { return fallback; }
  },
  set(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* storage unavailable */ }
  },
  remove(key) {
    try { localStorage.removeItem(key); } catch { /* storage unavailable */ }
  }
};

export class ZeeMoClient extends EventTarget {
  constructor(endpoint, options = {}) {
    super();
    this.endpoint = endpoint;
    this.activeInterval = options.activeInterval ?? 2500;   // a chat is open
    this.idleInterval = options.idleInterval ?? 10000;      // chat list only
    this.maxBackoff = options.maxBackoff ?? 60000;
    this.transport = options.transport || null;   // optional (action, token, payload) => Promise<response JSON>
    this.isActive = options.isActive || (() => true);   // were there taps or typing since the last poll?

    this.token = store.get('zm.token', null);
    // Earlier versions saved these; remove them so nothing about chats stays on the device.
    ['zm.cursors', 'zm.receipts', 'zm.listVersion'].forEach(k => store.remove(k));
    this.cursors = {};
    this.receiptVersions = {};
    this.deleteVersions = {};
    this.editVersions = {};
    this.listVersion = '';

    this.activeConvId = null;
    this.timer = null;
    this.failures = 0;
    this.running = false;
    this.inFlight = false;

    document.addEventListener('visibilitychange', () => {
      if (!this.running) return;
      if (document.hidden) this._clearTimer();
      else this._schedule(0);
    });
  }

  /* ---------- transport ---------- */

  async call(action, payload = {}) {
    if (this.transport) return this._unwrap(await this.transport(action, this.token, payload));
    let res;
    try {
      res = await fetch(this.endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, token: this.token, payload }),
        redirect: 'follow'
      });
    } catch (err) {
      throw new ZeeMoError('NETWORK', 'No connection.');
    }
    if (!res.ok) throw new ZeeMoError('NETWORK', `HTTP ${res.status}`);

    return this._unwrap(await res.json());
  }

  _unwrap(json) {
    if (!json.ok) {
      const code = json.error.code;
      if (code === 'UNAUTHORISED' || code === 'PASSCODE_LOCKED') this._signOutLocal(json.error.message);
      else if (code === 'LOCKED') { this.stop(); this._emit('locked'); }
      else if (code === 'PASSCODE_NOT_SET') { this.stop(); this._emit('needpasscode'); }
      else if (code === 'TERMS_REQUIRED') { this.stop(); this._emit('needterms'); }
      throw new ZeeMoError(code, json.error.message);
    }
    return json.data;
  }

  /* ---------- auth ---------- */

  get signedIn() {
    return !!this.token;
  }

  /** Which sign-in methods the platform owner has switched on: ['phone'], ['email'] or both. */
  authConfig() {
    return this.call('authConfig');
  }

  /** identity: { method: 'phone', phone } or { method: 'email', email }. A bare string is a phone number. */
  requestOtp(identity) {
    return this.call('requestOtp', typeof identity === 'string' ? { phone: identity } : identity);
  }

  async verifyOtp(identity, code, displayName) {
    const id = typeof identity === 'string' ? { phone: identity } : identity;
    const data = await this.call('verifyOtp', { ...id, code, displayName, device: deviceLabel() });
    this.token = data.token;
    store.set('zm.token', this.token);
    return data;
  }

  async logout() {
    try { await this.call('logout'); } finally { this._signOutLocal(); }
  }

  /** Forget everything this client holds about chats (used on lock and sign-out). */
  wipe() {
    this.stop();
    this.activeConvId = null;
    this.cursors = {};
    this.receiptVersions = {};
    this.deleteVersions = {};
    this.editVersions = {};
    this.listVersion = '';
  }

  /** Forget the session on this device only (the server has already ended it). */
  signOutHere(reason) { this._signOutLocal(reason); }

  _signOutLocal(reason) {
    if (!this.token) return;
    this.wipe();
    this.token = null;
    try {
      Object.keys(localStorage).filter(k => k.indexOf('zm.') === 0 && k.indexOf('zm.demo') !== 0).forEach(k => localStorage.removeItem(k));
      sessionStorage.clear();
    } catch { /* storage unavailable */ }
    this._emit('signedout', { reason: reason || '' });
  }

  /* ---------- passcode and lock ---------- */

  setPasscode(passcode, currentPasscode) {
    return this.call('setPasscode', currentPasscode ? { passcode, currentPasscode } : { passcode });
  }

  /** Lock on the server too, so no messages reach a locked screen. */
  lock() {
    this.wipe();
    this._locking = this.call('lock').catch(() => ({ locked: true }));
    return this._locking;
  }

  /** Waits for any lock request still in flight, so it cannot land after the unlock. */
  async unlock(passcode) {
    if (this._locking) await this._locking;
    return this.call('unlock', { passcode });
  }

  /* ---------- conversations ---------- */

  async listConversations() {
    const data = await this.call('listConversations');
    this.listVersion = data.listVersion;
    // New conversations start their cursor at the latest message; history comes from openChat().
    data.conversations.forEach(c => {
      if (!(c.convId in this.cursors)) this.cursors[c.convId] = c.lastMsgId || '';
    });
    return data.conversations;
  }

  /** Mark a chat as open: loads recent history and switches to fast polling. */
  async openChat(convId) {
    this.activeConvId = convId;
    const data = await this.call('getMessages', { convId, limit: 30 });
    const last = data.messages[data.messages.length - 1];
    if (last) {
      this._advanceCursor(convId, last.msgId);
      this.markRead(convId, last.msgId);
    }
    this._schedule(0);
    return data;
  }

  closeChat() {
    this.activeConvId = null;
  }

  loadOlder(convId, beforeMsgId) {
    return this.call('getMessages', { convId, before: beforeMsgId, limit: 30 });
  }

  markRead(convId, msgId) {
    return this.call('markRead', { convId, msgId }).catch(() => {});
  }

  deleteMessage(convId, msgId) {
    return this.call('deleteMessage', { convId, msgId });
  }

  /* ---------- sending with optimistic UI ---------- */

  async sendText(convId, body, replyTo) {
    return this._send({ convId, type: 'text', body, replyTo });
  }

  async sendMedia(convId, type, file, caption) {
    const base64 = await blobToBase64(file);
    const up = await this.call('uploadMedia', { convId, mimeType: file.type, base64 });
    return this._send({ convId, type, body: caption || '', mediaFileId: up.fileId });
  }

  async _send(payload) {
    const clientId = (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2));
    const pending = { ...payload, clientId, status: 'pending', createdAt: new Date().toISOString() };
    this._emit('pending', pending);

    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        // Same clientId on every retry: the server de-duplicates.
        const msg = await this.call('sendMessage', { ...payload, clientId });
        this._advanceCursor(payload.convId, msg.msgId);
        this._emit('sent', { clientId, message: msg });
        return msg;
      } catch (err) {
        const retryable = err.code === 'NETWORK' || err.code === 'BUSY' || err.code === 'SERVER_ERROR';
        if (!retryable || attempt === 2) {
          this._emit('failed', { clientId, error: err });
          throw err;
        }
        await sleep(800 * 2 ** attempt);
      }
    }
  }

  /* ---------- sync loop ---------- */

  start() {
    if (!this.token || this.running) return;
    this.running = true;
    this._schedule(0);
  }

  stop() {
    this.running = false;
    this._clearTimer();
  }

  async syncOnce() {
    if (this.inFlight) return;
    this.inFlight = true;
    try {
      const data = await this.call('sync', {
        cursors: this.cursors,
        receiptVersions: this.receiptVersions,
        deleteVersions: this.deleteVersions,
        editVersions: this.editVersions,
        listVersion: this.listVersion,
        active: this.isActive()   // lets the server keep the unlock window open only while someone is using the app
      });
      if (!this.running) return data;   // locked or signed out while the request was in flight

      Object.entries(data.cursors || {}).forEach(([convId, id]) => this._advanceCursor(convId, id));
      Object.entries(data.messages).forEach(([convId, msgs]) => {
        if (!msgs.length) return;
        this._advanceCursor(convId, msgs[msgs.length - 1].msgId);
        this._emit('messages', { convId, messages: msgs });
        if (convId === this.activeConvId && !document.hidden) {
          this.markRead(convId, msgs[msgs.length - 1].msgId);
        }
      });

      Object.entries(data.receipts).forEach(([convId, r]) => {
        this.receiptVersions[convId] = r.ver;
        this._emit('receipts', { convId, members: r.members });
      });

      Object.entries(data.deletions || {}).forEach(([convId, d]) => {
        this.deleteVersions[convId] = d.ver;
        this._emit('deletions', { convId, msgIds: d.msgIds });
      });

      Object.entries(data.edits || {}).forEach(([convId, d]) => {
        this.editVersions[convId] = d.ver;
        if (d.messages.length) this._emit('edits', { convId, messages: d.messages });
      });

      if (data.listChanged) {
        this.listVersion = data.listVersion;
        this._emit('listchanged');
      }

      this.failures = 0;
      this._emit('synced', data);
      return data;
    } finally {
      this.inFlight = false;
    }
  }

  _schedule(delay) {
    this._clearTimer();
    if (!this.running || document.hidden) return;
    this.timer = setTimeout(async () => {
      let next;
      try {
        const data = await this.syncOnce();
        next = data && data.hasMore ? 0 : (this.activeConvId ? this.activeInterval : this.idleInterval);
      } catch (err) {
        if (['UNAUTHORISED', 'PASSCODE_LOCKED', 'LOCKED', 'PASSCODE_NOT_SET', 'TERMS_REQUIRED'].includes(err.code)) return;
        this.failures += 1;
        next = Math.min(this.maxBackoff, 2000 * 2 ** this.failures);
        this._emit('syncerror', err);
      }
      this._schedule(next);
    }, delay);
  }

  _clearTimer() {
    if (this.timer) clearTimeout(this.timer);
    this.timer = null;
  }

  _advanceCursor(convId, msgId) {
    if (!this.cursors[convId] || msgId > this.cursors[convId]) this.cursors[convId] = msgId;
  }

  _emit(type, detail) {
    this.dispatchEvent(new CustomEvent(type, { detail }));
  }
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

/** Resize and re-encode an image before upload. */
export async function compressImage(file, maxWidth = 1280, quality = 0.7) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, maxWidth / bitmap.width);
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', quality));
  return new File([blob], 'photo.jpg', { type: 'image/jpeg' });
}
