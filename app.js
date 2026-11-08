/**
 * ZeeMoApp front end. Plain DOM, no framework.
 * All user text is inserted with textContent, never innerHTML.
 *
 * Privacy: messages, photos and voice notes live only in memory (state below) and
 * are wiped on lock, on sign-out and when the page is hidden or closed. Only the
 * session token (guarded by the passcode) and the owner's settings are kept on the device.
 */
import { ZeeMoClient, compressImage, deviceLabel } from './zeemo-client.js';

const CONFIG = window.ZEEMO_CONFIG || {};
// Inside Apps Script (Index.html served by doGet), calls go through google.script.run.
const IN_APPS_SCRIPT = !!(window.google && google.script && google.script.run);
const DEMO = !CONFIG.apiUrl && !IN_APPS_SCRIPT;
const api = new ZeeMoClient(CONFIG.apiUrl || 'local:', {
  transport: IN_APPS_SCRIPT ? appsScriptTransport : DEMO ? demoTransport() : null,
  isActive: () => activeSinceLastPoll()
});

/** The demo back end (about 150 KB) is downloaded only when the demo is actually used. */
function demoTransport() {
  let ready = null;
  return (action, token, payload) => {
    if (!ready) ready = loadDemo().then(m => m.createMockTransport());
    return ready.then(t => t(action, token, payload));
  };
}
const loadDemo = () => import('./mock-api.js');   // build: demo-loader

/** Same request and response JSON as doPost, carried by google.script.run. */
function appsScriptTransport(action, token, payload) {
  return new Promise(resolve => {
    google.script.run
      .withSuccessHandler(text => {
        try { resolve(JSON.parse(text)); }
        catch { resolve({ ok: false, error: { code: 'SERVER_ERROR', message: 'Unreadable server response.' } }); }
      })
      .withFailureHandler(() => resolve({ ok: false, error: { code: 'NETWORK', message: 'No connection.' } }))
      .api(JSON.stringify({ action, token, payload }));
  });
}

/* ---------- Icons (24px, stroke) ---------- */
const PATHS = {
  back: '<path d="M15 5l-7 7 7 7"/>',
  more: '<circle cx="12" cy="5" r="1.4" fill="currentColor"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/><circle cx="12" cy="19" r="1.4" fill="currentColor"/>',
  compose: '<path d="M4 20l1.2-4.2L15.6 5.4a2 2 0 012.8 0l.2.2a2 2 0 010 2.8L8.2 18.8 4 20z"/><path d="M13.5 7.5l3 3"/>',
  send: '<path d="M4 12l16-8-6 16-2.5-6.5L4 12z"/>',
  mic: '<rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5.5 11a6.5 6.5 0 0013 0M12 17.5V21"/>',
  stop: '<rect x="7" y="7" width="10" height="10" rx="2" fill="currentColor"/>',
  image: '<rect x="3.5" y="5" width="17" height="14" rx="2.5"/><circle cx="9" cy="10" r="1.6"/><path d="M4 17l5-4.5 4 3.5 3-2.5 4 3.5"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>',
  reply: '<path d="M10 7L4 12l6 5"/><path d="M4.5 12H14a6 6 0 016 6v1"/>',
  group: '<circle cx="9" cy="9" r="3.2"/><path d="M3.5 19a5.5 5.5 0 0111 0"/><circle cx="17" cy="9.5" r="2.4"/><path d="M15.8 14.2A4.5 4.5 0 0120.5 19"/>',
  tick: '<path d="M5 12.5l4 4L19 7"/>',
  ticks: '<path d="M2.5 12.5l4 4L16 7"/><path d="M11.5 15.5l1 1L22 7"/>',
  clock: '<circle cx="12" cy="12" r="7.5"/><path d="M12 8v4.5l3 1.8"/>',
  alert: '<circle cx="12" cy="12" r="8"/><path d="M12 8v5M12 16h.01"/>',
  timer: '<circle cx="12" cy="13" r="7.5"/><path d="M12 9v4l2.5 1.5M9.5 3h5"/>',
  trash: '<path d="M4.5 7h15M10 11v6M14 11v6M6.5 7l1 12.5a1.5 1.5 0 001.5 1.5h6a1.5 1.5 0 001.5-1.5L17.5 7M9.5 7V4.5h5V7"/>',
  lock: '<rect x="5" y="10.5" width="14" height="10" rx="2.5"/><path d="M8 10.5V8a4 4 0 018 0v2.5"/>',
  smile: '<circle cx="12" cy="12" r="8.5"/><path d="M8.5 14a4.2 4.2 0 007 0"/><circle cx="9.3" cy="10" r=".9" fill="currentColor" stroke="none"/><circle cx="14.7" cy="10" r=".9" fill="currentColor" stroke="none"/>',
  flag: '<path d="M5.5 21V4.5"/><path d="M5.5 4.5h11l-2.2 4 2.2 4h-11"/>',
  edit: '<path d="M4 20l1.2-4.2L15.6 5.4a2 2 0 012.8 0l.2.2a2 2 0 010 2.8L8.2 18.8 4 20z"/>',
  camera: '<path d="M4 8.5A1.5 1.5 0 015.5 7h2.2l1.5-2h5.6l1.5 2h2.2A1.5 1.5 0 0120 8.5v9a1.5 1.5 0 01-1.5 1.5h-13A1.5 1.5 0 014 17.5z"/><circle cx="12" cy="13" r="3.4"/>',
  keyboard: '<rect x="3" y="6.5" width="18" height="11" rx="2"/><path d="M7 10h.01M10 10h.01M13 10h.01M16 10h.01M8 14h8"/>'
};
function icon(name) {
  const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  s.setAttribute('viewBox', '0 0 24 24');
  s.setAttribute('fill', 'none');
  s.setAttribute('stroke', 'currentColor');
  s.setAttribute('stroke-width', '1.9');
  s.setAttribute('stroke-linecap', 'round');
  s.setAttribute('stroke-linejoin', 'round');
  s.setAttribute('aria-hidden', 'true');
  s.innerHTML = PATHS[name];   // static, trusted strings only
  return s;
}

/* ---------- DOM helper ---------- */
function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === null || v === undefined || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k === 'text') el.textContent = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return el;
}
const $ = id => document.getElementById(id);

/** Focus a field once it is on screen, unless the person has already moved to another field. */
function focusSoon(el) {
  requestAnimationFrame(() => {
    const a = document.activeElement;
    if (el.isConnected && (!a || a === document.body || !a.matches('input, textarea, select'))) el.focus();
  });
}
/** append() that skips null and false, which DOM append would render as text. */
function appendAll(el, ...nodes) {
  el.append(...nodes.flat().filter(n => n !== null && n !== undefined && n !== false));
}

/* ---------- State ---------- */
const state = {
  me: null,
  convs: [],
  msgs: {},          // convId -> Map(msgId -> msg)
  pending: {},       // convId -> Map(clientId -> pending msg)
  receipts: {},      // convId -> [{userId, lastReadMsgId, lastDeliveredMsgId}]
  activeId: null,
  replyTo: null,
  filter: '',
  media: {},         // fileId -> object URL | 'loading' | 'missing'
  names: {},         // userId -> displayName
  members: {},       // convId -> [{userId, displayName, role}]
  security: { requirePasscode: false, lockAfterMinutes: 0, minimiseAction: 'none', deleteWindowHours: 48, disappearOptions: [0, 86400, 604800, 7776000] },
  links: {},         // public Terms, Privacy and Delete-account pages
  locked: false
};

/* ---------- Formatting ---------- */
const AVATAR_HUES = ['#2B3C96', '#8A4B12', '#1F6F5C', '#7A2E57', '#4B5A12', '#205C8A', '#6B3FA0', '#A0431F'];
function hueFor(id) {
  let n = 0;
  for (const ch of String(id)) n = (n * 31 + ch.charCodeAt(0)) >>> 0;
  return AVATAR_HUES[n % AVATAR_HUES.length];
}
/** First name for short labels, without a title such as Mrs or Dr. */
function firstName(name) {
  return String(name || '').replace(/^(mr|mrs|ms|dr|prof)\.?\s+/i, '').trim().split(/\s+/)[0] || 'them';
}
function initials(name) {
  const parts = String(name || '?').replace(/^(mr|mrs|ms|dr|prof)\.?\s+/i, '').trim().split(/\s+/);
  return ((parts[0] || '?')[0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}
function avatar(conv, small) {
  const a = h('span', { class: 'avatar' + (small ? ' sm' : ''), 'aria-hidden': 'true' });
  a.style.background = hueFor(conv.type === 'group' ? conv.convId : (conv.peer ? conv.peer.userId : conv.convId));
  if (conv.type === 'group') a.append(icon('group'));
  else a.textContent = initials(conv.title);
  const fileId = conv.avatarFileId || (conv.peer && conv.peer.avatarFileId);
  if (fileId) a.append(photoImg(fileId));
  return a;
}

/* Profile and group photos: fetched once, kept in memory only, wiped on lock. */
const photos = {};   // fileId -> object URL | 'loading' | 'missing'

function photoImg(fileId) {
  const img = h('img', { class: 'avatar-img', alt: '', 'data-photo': fileId });
  const st = photos[fileId];
  if (st && st !== 'loading' && st !== 'missing') img.src = st;
  else if (!st) loadPhoto(fileId);
  if (st === 'missing') img.hidden = true;
  return img;
}

async function loadPhoto(fileId) {
  photos[fileId] = 'loading';
  try {
    const r = await api.call('getMedia', { fileId });
    if (state.locked || !state.me) { delete photos[fileId]; return; }
    const bytes = Uint8Array.from(atob(r.base64), c => c.charCodeAt(0));
    photos[fileId] = URL.createObjectURL(new Blob([bytes], { type: r.mimeType }));
  } catch {
    photos[fileId] = 'missing';
  }
  document.querySelectorAll(`img[data-photo="${CSS.escape(fileId)}"]`).forEach(img => {
    if (photos[fileId] === 'missing') img.hidden = true;
    else img.src = photos[fileId];
  });
}

function forgetPhotos() {
  Object.keys(photos).forEach(k => {
    if (typeof photos[k] === 'string' && photos[k].startsWith('blob:')) URL.revokeObjectURL(photos[k]);
    delete photos[k];
  });
}

/** Square-crop and shrink a picture for a profile or group photo. */
async function squarePhoto(file, size = 320) {
  const bitmap = await createImageBitmap(file);
  const side = Math.min(bitmap.width, bitmap.height);
  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = size;
  canvas.getContext('2d').drawImage(bitmap, (bitmap.width - side) / 2, (bitmap.height - side) / 2, side, side, 0, 0, size, size);
  const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.82));
  const base64 = await new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result).split(',')[1]);
    fr.onerror = () => reject(fr.error);
    fr.readAsDataURL(blob);
  });
  return { mimeType: 'image/jpeg', base64 };
}

/** "Change photo" and "Remove photo" controls. save(upload|null) does the API call. */
function photoControls(current, save) {
  const input = h('input', { type: 'file', accept: 'image/*', hidden: true });
  const change = h('button', { class: 'btn small secondary', type: 'button', text: current ? 'Change photo' : 'Add photo' });
  const remove = current ? h('button', { class: 'btn small danger', type: 'button', text: 'Remove photo' }) : null;
  change.addEventListener('click', () => { pickerOpened(); input.click(); });
  input.addEventListener('change', async () => {
    pickerClosed();
    const f = input.files[0];
    input.value = '';
    if (!f) return;
    change.disabled = true;
    try { await save(await squarePhoto(f)); }
    catch (err) { toast(errorText(err)); change.disabled = false; }
  });
  if (remove) remove.addEventListener('click', async () => {
    remove.disabled = true;
    try { await save(null); } catch (err) { toast(errorText(err)); remove.disabled = false; }
  });
  return h('div', { class: 'photo-controls' }, change, remove, input);
}
const timeFmt = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });
const dayFmt = new Intl.DateTimeFormat(undefined, { weekday: 'long' });
const dateFmt = new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
function sameDay(a, b) { return a.toDateString() === b.toDateString(); }
function listTime(iso) {
  if (!iso) return '';
  const d = new Date(iso), now = new Date();
  if (sameDay(d, now)) return timeFmt.format(d);
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (sameDay(d, y)) return 'Yesterday';
  if (now - d < 6 * 86400000) return dayFmt.format(d);
  return dateFmt.format(d);
}
function dayLabel(d) {
  const now = new Date();
  if (sameDay(d, now)) return 'Today';
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (sameDay(d, y)) return 'Yesterday';
  return dateFmt.format(d);
}
function lastSeenLabel(iso) {
  if (!iso) return '';
  const diff = Date.now() - new Date(iso).getTime();
  if (diff < 90000) return 'online';
  if (diff < 3600000) return `last seen ${Math.round(diff / 60000)} min ago`;
  return `last seen ${listTime(iso).toLowerCase() === 'yesterday' ? 'yesterday' : listTime(iso)}`;
}
function previewOf(m) {
  if (m.deleted) return 'This message was deleted';
  if (m.type === 'system') return systemText(m);
  if (m.type === 'image') return 'Photo' + (m.body ? ': ' + m.body : '');
  if (m.type === 'audio') return 'Voice note';
  return m.body;
}

function timerLabel(sec) {
  sec = Number(sec) || 0;
  if (!sec) return 'Off';
  if (sec % 86400 === 0) return sec === 86400 ? '24 hours' : `${sec / 86400} days`;
  return `${Math.round(sec / 3600)} hours`;
}

/** System notices carry names and IDs so each reader sees "You" for themselves. */
function systemText(m) {
  let s;
  try { s = JSON.parse(m.body); } catch { return ''; }
  const me = state.me && state.me.userId;
  const who = p => (p && p[0] === me ? 'You' : (p && p[1]) || 'Someone');
  const a = who(s.a);
  const t = (s.t || []).map(who).join(', ');
  switch (s.k) {
    case 'added': return `${a} added ${t}`;
    case 'removed': return `${a} removed ${t}`;
    case 'left': return `${a} left`;
    case 'renamed': return `${a} changed the group name to “${s.v}”`;
    case 'admin': return `${a} made ${t} an admin`;
    case 'unadmin': return `${a} removed ${t} as admin`;
    case 'promoted': return t === 'You' ? 'You are now an admin' : `${t} is now an admin`;
    case 'photo': return `${a} changed the group photo`;
    case 'photoRemoved': return `${a} removed the group photo`;
    case 'disappear': return Number(s.v)
      ? `${a} set messages to disappear after ${timerLabel(s.v).toLowerCase()}`
      : `${a} turned off disappearing messages`;
    default: return '';
  }
}

function isExpired(m) {
  return !!m.expiresAt && new Date(m.expiresAt).getTime() <= Date.now();
}

/* ---------- Toast ---------- */
let toastTimer;
function toast(message) {
  const t = $('toast');
  t.textContent = message;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 3600);
}
function errorText(e) {
  if (e && e.code === 'NETWORK') return 'No connection. Check your data or Wi-Fi and try again.';
  if (e && e.code === 'BUSY') return 'The server is busy. Try again in a moment.';
  if (e && e.code === 'SERVER_ERROR') return 'Something went wrong on the server. Try again.';
  return (e && e.message) || 'Something went wrong.';
}

/* ---------- Sign-in ---------- */
// methods comes from the back end (SIGNIN_METHOD): ['phone'], ['email'] or both.
const auth = { step: 'start', methods: null, method: 'phone', value: '', sentTo: '' };

const METHOD_COPY = {
  phone: { tab: 'Phone', label: 'Phone number', type: 'tel', mode: 'tel', ac: 'tel', ph: '0803 123 4567', where: 'sent by SMS to' },
  email: { tab: 'Email', label: 'Email address', type: 'email', mode: 'email', ac: 'email', ph: 'name@example.com', where: 'emailed to' }
};

const CFG_KEY = 'zm.cfg';   // owner settings only: no personal data

function applyAuthConfig(cfg) {
  auth.methods = cfg.methods && cfg.methods.length ? cfg.methods : ['phone'];
  if (cfg.security) state.security = cfg.security;
  if (cfg.links) state.links = cfg.links;
  if (!auth.methods.includes(auth.method)) auth.method = auth.methods[0];
}

/** Settings from the last visit are used at once; a fresh copy is fetched alongside. */
function cachedAuthConfig() {
  try { const c = JSON.parse(localStorage.getItem(CFG_KEY) || 'null'); if (c) applyAuthConfig(c); return !!c; } catch { return false; }
}

async function loadAuthConfig(force) {
  if (auth.methods && !force) return;
  try {
    const cfg = await api.authConfig();
    applyAuthConfig(cfg);
    try { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); } catch { /* storage unavailable */ }
  } catch {
    if (!auth.methods) applyAuthConfig({ methods: ['phone'] });   // older back end, or offline
  }
}

function lede() {
  const m = auth.methods;
  if (m.length > 1) return 'Messages for your school, office or cooperative. Sign in with your phone number or email.';
  return `Messages for your school, office or cooperative. Sign in with your ${m[0] === 'email' ? 'email address' : 'phone number'}.`;
}

async function renderAuth(errorMsg, demoCode) {
  const root = $('auth');
  root.hidden = false;
  root.dataset.mode = 'signin';
  await loadAuthConfig();
  root.replaceChildren();
  const card = h('div', { class: 'auth-card' },
    h('div', { class: 'wordmark', 'aria-label': 'ZeeMo' }, h('b', { text: 'Zee' }), h('span', { text: 'Mo' })),
    DEMO ? h('span', { class: 'demo-pill', text: 'Demo mode: no real SMS or email is sent' }) : null
  );

  if (auth.step === 'start') {
    const copy = METHOD_COPY[auth.method];
    const input = h('input', { id: 'auth-id', class: 'input', type: copy.type, inputmode: copy.mode, autocomplete: copy.ac, placeholder: copy.ph, required: true, value: auth.value });
    const btn = h('button', { class: 'btn', type: 'submit', text: 'Send code' });

    const tabs = auth.methods.length > 1 ? h('div', { class: 'seg', role: 'radiogroup', 'aria-label': 'Sign in with' },
      ...auth.methods.map(m => h('button', {
        type: 'button', role: 'radio', class: 'seg-btn', 'aria-checked': String(m === auth.method),
        text: METHOD_COPY[m].tab,
        onclick: () => { if (m !== auth.method) { auth.method = m; auth.value = ''; renderAuth(); } }
      }))) : null;

    appendAll(card,
      h('p', { class: 'auth-lede', text: lede() }),
      tabs,
      h('form', {
        onsubmit: async e => {
          e.preventDefault();
          btn.disabled = true; btn.textContent = 'Sending…';
          auth.value = input.value;
          try {
            const r = await api.requestOtp({ method: auth.method, [auth.method]: input.value });
            auth.sentTo = r[auth.method] || input.value;
            auth.step = 'code';
            renderAuth('', r.devCode);
          } catch (err) {
            renderAuth(errorText(err));
          }
        }
      },
        h('label', { for: 'auth-id', text: copy.label }), input,
        h('p', { class: 'error', role: 'alert', text: errorMsg || '' }),
        btn),
      legalLinks()
    );
    focusSoon(input);
  } else {
    const copy = METHOD_COPY[auth.method];
    const code = h('input', { id: 'auth-code', class: 'input code', inputmode: 'numeric', autocomplete: 'one-time-code', maxlength: '6', pattern: '[0-9]{6}', placeholder: '••••••', required: true });

    const name = h('input', { id: 'auth-name', class: 'input', autocomplete: 'name', maxlength: '40', placeholder: 'e.g. Aisha Bello' });
    const btn = h('button', { class: 'btn', type: 'submit', text: 'Verify and continue' });
    appendAll(card,
      h('p', { class: 'auth-lede', text: `Enter the 6-digit code ${copy.where} ${auth.sentTo}.` }),
      auth.method === 'email' && !demoCode ? h('p', { class: 'note', text: 'It can take a minute to arrive. Check your spam or promotions folder too.' }) : null,
      demoCode ? h('p', { class: 'demo-code', text: DEMO ? `Demo code: ${demoCode}` : `Test code: ${demoCode} (DEV_MODE is on, so nothing was sent)` }) : null,
      h('form', {
        onsubmit: async e => {
          e.preventDefault();
          btn.disabled = true; btn.textContent = 'Checking…';
          try {
            const r = await api.verifyOtp({ method: auth.method, [auth.method]: auth.sentTo }, code.value, name.value);
            state.me = r.user;
            auth.step = 'start';
            loadAuthConfig(true);   // refresh in the background
            // A sign-in code is not enough on its own: a device must also know the passcode.
            nextStep('code');
          } catch (err) {
            btn.disabled = false; btn.textContent = 'Verify and continue';
            $('auth-error').textContent = errorText(err);
            if (err.code === 'OTP_EXPIRED' || err.code === 'OTP_LOCKED') { auth.step = 'start'; renderAuth(errorText(err)); }
          }
        }
      },
        h('label', { for: 'auth-code', text: 'Code' }), code,
        h('label', { for: 'auth-name', text: 'Your name (first sign-in only)' }), name,
        h('p', { id: 'auth-error', class: 'error', role: 'alert', text: errorMsg || '' }),
        btn,
        h('button', { class: 'link', type: 'button', text: auth.method === 'email' ? 'Use a different email' : 'Use a different number', onclick: () => { auth.step = 'start'; renderAuth(); } })
      ),
      legalLinks()
    );
    focusSoon(code);
  }
  root.append(card);
}

/* ---------- Chat list ---------- */
function unreadBadge(c) {
  const n = Number(c.unreadCount) || 0;
  if (!n) return h('span', { class: 'dot', 'aria-label': 'Unread' });
  const label = c.unreadMore ? `${n}+` : String(n > 99 ? '99+' : n);
  return h('span', { class: 'count', 'aria-label': `${label} unread` , text: label });
}

function renderList() {
  const list = $('list');
  const q = state.filter.trim().toLowerCase();
  const rows = state.convs.filter(c => !q || c.title.toLowerCase().includes(q) || (c.lastMsgPreview || '').toLowerCase().includes(q));
  list.replaceChildren();
  if (!state.convs.length) {
    list.append(h('div', { class: 'empty' },
      h('h3', { text: 'No chats yet' }),
      h('p', { text: 'Start a chat with someone by their phone number, or create a group.' }),
      h('button', { class: 'btn', text: 'New chat', onclick: openNewChat })));
    return;
  }
  if (!rows.length) {
    list.append(h('div', { class: 'empty' }, h('p', { text: `No chats match “${state.filter}”.` })));
    return;
  }
  for (const c of rows) {
    const mine = c.lastMsgSenderId && state.me && c.lastMsgSenderId === state.me.userId;
    const who = c.type === 'group' && c.lastMsgSenderId && !mine ? (state.names[c.lastMsgSenderId] || '').split(' ')[0] : '';
    const notice = c.lastMsgType === 'system';
    const preview = (notice ? '' : mine ? 'You: ' : who ? who + ': ' : '') + (c.lastMsgPreview || (c.type === 'group' ? 'Group created' : 'Say hello'));
    list.append(h('button', {
      class: 'row' + (c.hasUnread ? ' unread' : ''),
      'aria-current': c.convId === state.activeId ? 'true' : null,
      onclick: () => openChat(c.convId)
    },
      avatar(c),
      h('span', { class: 'row-main' },
        h('span', { class: 'row-top' }, h('span', { class: 'row-title', text: c.title }), h('span', { class: 'row-time', text: listTime(c.lastMsgAt) })),
        h('span', { class: 'row-bottom' },
          h('span', { class: 'row-preview', text: preview }),
          c.hasUnread ? unreadBadge(c) : null))
    ));
  }
}

const QUIET_ERRORS = ['UNAUTHORISED', 'LOCKED', 'PASSCODE_NOT_SET', 'PASSCODE_LOCKED', 'TERMS_REQUIRED'];

async function loadConversations() {
  if (state.locked) return;
  try {
    state.convs = await api.listConversations();
    if (state.locked) { state.convs = []; return; }
    state.convs.forEach(c => { if (c.peer) state.names[c.peer.userId] = c.peer.displayName; });
    renderList();
    state.convs.filter(c => c.type === 'group' && c.lastMsgSenderId && !state.names[c.lastMsgSenderId]).forEach(resolveGroupNames);
    if (state.activeId && !activeConv()) {
      // Removed from a group, or the chat is gone.
      closeChat();
      closeSheet();
      toast('You are no longer in that chat.');
    } else if (state.activeId) {
      renderChatHeader();
      if (composerKey(activeConv()) !== state.composerKey) renderComposer();
      if (activeConv().type === 'group') resolveGroupNames(activeConv());
    }
    if (!$('sheet').hidden && $('sheet').dataset.convId) openChatInfo($('sheet').dataset.convId, true);
  } catch (err) {
    if (!QUIET_ERRORS.includes(err.code)) toast(errorText(err));
  }
}

/* ---------- Chat view ---------- */
function activeConv() {
  return state.convs.find(c => c.convId === state.activeId) || null;
}

function mergeMessages(convId, list) {
  const map = state.msgs[convId] || (state.msgs[convId] = new Map());
  list.forEach(m => {
    map.set(m.msgId, m);
    if (m.senderId === (state.me && state.me.userId) && m.clientId && state.pending[convId]) state.pending[convId].delete(m.clientId);
  });
}

async function openChat(convId) {
  state.activeId = convId;
  state.replyTo = null;
  state.editing = null;
  document.body.classList.add('chat-open');
  const conv = activeConv();
  if (conv) { conv.hasUnread = false; conv.unreadCount = 0; conv.unreadMore = false; }
  renderList();
  renderChatShell();
  try {
    const r = await api.openChat(convId);
    if (state.locked || state.activeId !== convId) return;
    mergeMessages(convId, r.messages);
    if (state.activeId === convId) renderThread(true);
    if (conv && conv.type === 'group') resolveGroupNames(conv);
  } catch (err) {
    toast(errorText(err));
  }
}

/** Close the open chat and drop its messages and media from memory. */
function closeChat() {
  const id = state.activeId;
  if (rec.recorder) stopRecording(false);
  closeActions();
  state.activeId = null;
  state.replyTo = null;
  state.editing = null;
  api.closeChat();
  if (id) {
    delete state.msgs[id];
    delete state.pending[id];
    forgetMedia();
  }
  document.body.classList.remove('chat-open');
  renderList();
  renderNoChat();
}

function forgetMedia() {
  Object.values(state.media).forEach(u => { if (typeof u === 'string' && u.startsWith('blob:')) URL.revokeObjectURL(u); });
  state.media = {};
}

function renderNoChat() {
  $('chat').replaceChildren(h('div', { class: 'no-chat' },
    h('div', {}, h('h2', { text: 'Pick a chat' }), h('p', { text: 'Your conversations open here.' }))));
}

function renderChatShell() {
  const pane = $('chat');
  pane.replaceChildren(
    h('header', { class: 'bar', id: 'chat-bar' }),
    h('div', { class: 'thread', id: 'thread', role: 'log', 'aria-live': 'polite' }),
    h('div', { class: 'composer-wrap', id: 'composer-wrap' })
  );
  renderChatHeader();
  renderComposer();
  renderThread(true);
}

function renderChatHeader() {
  const bar = $('chat-bar');
  const conv = activeConv();
  if (!bar || !conv) return;
  let sub = '';
  if (conv.type === 'group') sub = `${conv.memberCount} members`;
  else if (conv.peer && conv.peer.deleted) sub = '';
  else if (conv.peer && conv.peer.blockedByMe) sub = 'Blocked';
  else if (conv.peer) sub = lastSeenLabel(conv.peer.lastSeen) || conv.peer.about || '';
  if (conv.disappearSec) sub = `Disappearing: ${timerLabel(conv.disappearSec).toLowerCase()}` + (sub ? ` · ${sub}` : '');
  bar.replaceChildren();
  appendAll(bar,
    h('button', { class: 'icon-btn back-btn', 'aria-label': 'Back to chats', onclick: closeChat }, icon('back')),
    h('button', { class: 'chat-head grow', 'aria-label': `Chat info for ${conv.title}`, onclick: () => openChatInfo(conv.convId) },
      avatar(conv, true),
      h('span', { class: 'chat-title' }, h('strong', { text: conv.title }), sub ? h('small', { text: sub }) : null)),
    conv.disappearSec ? h('span', { class: 'head-timer', title: 'Disappearing messages on' }, icon('timer')) : null
  );
}

function tickFor(m) {
  const me = state.me.userId;
  const others = (state.receipts[m.convId] || []).filter(r => r.userId !== me);
  if (!others.length) return h('span', { 'aria-label': 'Sent' }, icon('tick'));
  if (others.every(r => r.lastReadMsgId >= m.msgId)) return h('span', { class: 'read', 'aria-label': 'Read' }, icon('ticks'));
  if (others.every(r => r.lastDeliveredMsgId >= m.msgId)) return h('span', { 'aria-label': 'Delivered' }, icon('ticks'));
  return h('span', { 'aria-label': 'Sent' }, icon('tick'));
}

function mediaNode(m) {
  if (!m.mediaFileId) return null;
  const st = state.media[m.mediaFileId];
  if (st === 'missing') return h('div', { class: 'media-missing', text: m.type === 'image' ? 'Photo not available' : 'Voice note not available' });
  if (!st) loadMedia(m.mediaFileId);
  if (m.type === 'image') return h('img', { class: 'media-img', alt: 'Photo', src: st && st !== 'loading' ? st : null });
  if (m.type === 'audio') return st && st !== 'loading' ? h('audio', { class: 'media-audio', controls: true, preload: 'metadata', src: st }) : h('div', { class: 'media-missing', text: 'Loading voice note…' });
  return null;
}

async function loadMedia(fileId) {
  state.media[fileId] = 'loading';
  try {
    const r = await api.call('getMedia', { fileId });
    if (state.locked || !state.activeId) { delete state.media[fileId]; return; }
    const bytes = Uint8Array.from(atob(r.base64), c => c.charCodeAt(0));
    state.media[fileId] = URL.createObjectURL(new Blob([bytes], { type: r.mimeType }));
  } catch {
    state.media[fileId] = 'missing';
  }
  if (state.activeId) renderThread(false);
}

function quoteNode(msgId, convId) {
  const q = state.msgs[convId] && state.msgs[convId].get(msgId);
  if (!q || isExpired(q)) return null;
  const who = q.senderId === state.me.userId ? 'You' : (state.names[q.senderId] || 'Member');
  return h('div', { class: 'quote' }, h('b', { text: who }), h('span', { text: previewOf(q) }));
}

let lastRenderedId = '';
function renderThread(forceBottom) {
  const thread = $('thread');
  const conv = activeConv();
  if (!thread || !conv) return;
  const nearBottom = thread.scrollHeight - thread.scrollTop - thread.clientHeight < 120;
  const me = state.me.userId;
  const sent = Array.from((state.msgs[conv.convId] || new Map()).values())
    .filter(m => !isExpired(m))
    .sort((a, b) => (a.msgId < b.msgId ? -1 : 1));
  const pend = Array.from((state.pending[conv.convId] || new Map()).values());
  const all = sent.concat(pend);

  const frag = document.createDocumentFragment();
  if (!all.length) frag.append(h('div', { class: 'day', text: conv.type === 'group' ? 'Group created. Say hello.' : 'No messages yet. Say hello.' }));
  let prevDay = '', prevSender = '';
  const newest = sent.length ? sent[sent.length - 1].msgId : '';
  for (const m of all) {
    const d = new Date(m.createdAt);
    const dk = d.toDateString();
    if (dk !== prevDay) { frag.append(h('div', { class: 'day', text: dayLabel(d) })); prevSender = ''; }
    prevDay = dk;
    if (m.type === 'system') {
      frag.append(h('div', { class: 'sys', role: 'note', text: systemText(m) }));
      prevSender = '';
      continue;
    }
    const out = m.senderId === me || m.status;
    const first = m.senderId !== prevSender;
    prevSender = m.senderId;
    const fresh = m.msgId && m.msgId > lastRenderedId && lastRenderedId !== '';
    const failed = m.status === 'failed';

    const meta = h('span', { class: 'meta' },
      m.editedAt ? h('span', { class: 'edited', text: 'edited' }) : null,
      timeFmt.format(d),
      out ? (m.status === 'pending' ? h('span', { 'aria-label': 'Sending' }, icon('clock'))
        : failed ? h('span', { 'aria-label': 'Not sent' }, icon('alert'))
        : tickFor(m)) : null);

    const bubble = m.deleted
      ? h('div', { class: 'bubble deleted' },
          !out && first && conv.type === 'group' ? h('div', { class: 'sender', text: state.names[m.senderId] || 'Member' }) : null,
          h('span', { class: 'text' }, icon('trash'), out ? 'You deleted this message' : 'This message was deleted'),
          h('span', { class: 'meta' }, timeFmt.format(d)))
      : h('div', { class: 'bubble' },
          !out && first && conv.type === 'group' ? h('div', { class: 'sender', text: state.names[m.senderId] || 'Member' }) : null,
          m.replyTo ? quoteNode(m.replyTo, conv.convId) : null,
          mediaNode(m),
          m.body ? h('span', { class: 'text', text: m.body }) : null,
          meta,
          m.msgId ? h('button', { class: 'more-btn', 'aria-label': 'Message options', 'aria-haspopup': 'dialog', onclick: () => openActions(m) }, icon('more')) : null
        );
    if (!out && first && conv.type === 'group') bubble.querySelector('.sender').style.setProperty('--sender', hueFor(m.senderId));
    if (m.expiresAt) bubble.querySelector('.meta').prepend(h('span', { class: 'exp', 'aria-label': `Disappears ${dateFmt.format(new Date(m.expiresAt))}` }, icon('timer')));

    frag.append(h('div', { class: `msg ${out ? 'out' : 'in'}${first ? ' first' : ''}${fresh ? ' fresh' : ''}${failed ? ' failed' : ''}` },
      bubble,
      failed ? h('button', { class: 'retry', text: 'Not sent. Tap to retry', onclick: () => retry(m) }) : null));
  }
  thread.replaceChildren(frag);
  lastRenderedId = newest;
  if (forceBottom || nearBottom) thread.scrollTop = thread.scrollHeight;
}

/* ---------- Emoji picker ---------- */
// Common emoji that display on older phones too (nothing newer than Emoji 13).
const EMOJI = [
  ['Expressions', '😀 😃 😄 😁 😆 😅 😂 🤣 😊 😇 🙂 🙃 😉 😌 😋 😛 😜 🤪 😝 🤗 🤭 🤫 🤔 🤐 🤨 😐 😑 😶 😏 😒 🙄 😬 😮 😯 😲 😳 🥺 😦 😧 😨 😰 😥 😢 😭 😱 😖 😣 😞 😓 😩 😫 🥱 😤 😡 😠 🤬 😎 🤓 🧐 🥳 🤩 😴 🤤 😪 😵 🤯 🥴 🤢 🤮 🤧 😷 🤒 🤕 🥵 🥶'],
  ['Love', '❤️ 🧡 💛 💚 💙 💜 🤎 🖤 🤍 💕 💞 💓 💗 💖 💘 💝 💟 ❣️ 💔 😍 🥰 😘 😗 😙 😚 😻 💋 💌 💑 💏 🌹 💐 🤗'],
  ['Gestures', '👍 👎 👏 🙌 👐 🤲 🙏 🤝 👋 🤚 ✋ 🖐️ 👌 ✌️ 🤞 🤟 🤘 🤙 👈 👉 👆 👇 ☝️ ✊ 👊 🤛 🤜 💪 🙇 🙋 🙆 🙅 🤷 🤦'],
  ['Celebration', '🎉 🎊 🥳 🎂 🍰 🎁 🎈 🎀 🏆 🥇 🥈 🥉 🏅 🎖️ 🎓 ✨ 🌟 ⭐ 🔥 💯 🥂 🍾 🎶 🎵'],
  ['Nature', '🌞 🌙 🌈 ☁️ 🌧️ ⛈️ ❄️ 🌊 🌴 🌳 🌵 🌷 🌻 🌺 🌸 🍀 🍁 🐶 🐱 🐰 🦊 🐻 🐼 🦁 🐯 🐮 🐔 🐐 🐑 🐴 🐘 🦒 🐒 🐦 🐟 🦋 🐝'],
  ['Food', '🍚 🍛 🍲 🥘 🍝 🍜 🍞 🥚 🍳 🍗 🍖 🥩 🍤 🥗 🌯 🍔 🍟 🍕 🌽 🥕 🍅 🍆 🥔 🧅 🌶️ 🥜 🍌 🍍 🥭 🍉 🍊 🍎 🍇 🥥 🍫 🍬 🍪 🍩 🍦 ☕ 🍵 🥤 🧃 🍹'],
  ['Activities', '⚽ 🏀 🏈 ⚾ 🎾 🏐 🏓 🏸 🥊 🥋 🏃 🚶 🏊 🚴 🧘 💃 🕺 🎮 🎲 ♟️ 🎯 🎤 🎧 🎸 🥁 🎹 🎬 📚 ✍️ 🎨'],
  ['Objects', '📱 💻 🖥️ ⌨️ 🖨️ 📷 📞 ☎️ 📺 📻 ⏰ ⌛ 💡 🔦 🔋 🔌 💰 💵 💳 🧾 📝 📄 📅 📌 📎 ✂️ 🔒 🔑 🧰 🚗 🚌 🏍️ ✈️ 🏠 🏫 🏥 🕌 ⛪ 🛒'],
  ['Symbols', '✅ ☑️ ✔️ ❌ ❎ ⚠️ ❗ ❓ ‼️ ⁉️ 🔴 🟢 🟡 🔵 ⚪ ⚫ ➡️ ⬅️ ⬆️ ⬇️ 🔁 🔄 ➕ ➖ ✖️ ➗ 🆗 🆕 🆓 🔝 ℹ️ ♻️ 🚫 ⛔ 📢 🔔 🔕 💤']
].map(([name, list]) => ({ name, items: list.split(' ') }));

// Recently used emoji are kept in memory only and forgotten when the app locks.
const emojiState = { open: false, tab: 'Expressions', recent: [] };

function emojiPanel(onPick) {
  const cats = (emojiState.recent.length ? [{ name: 'Recent', items: emojiState.recent }] : []).concat(EMOJI);
  if (!cats.some(c => c.name === emojiState.tab)) emojiState.tab = cats[0].name;
  const grid = h('div', { class: 'emoji-grid', role: 'tabpanel', id: 'emoji-grid' });
  const tabs = h('div', { class: 'emoji-tabs', role: 'tablist', 'aria-label': 'Emoji categories' });
  const show = name => {
    emojiState.tab = name;
    const cat = cats.find(c => c.name === name);
    tabs.querySelectorAll('[role="tab"]').forEach(t => t.setAttribute('aria-selected', String(t.dataset.cat === name)));
    grid.setAttribute('aria-label', name);
    grid.replaceChildren(...cat.items.map(e => h('button', { type: 'button', class: 'emoji', 'aria-label': e, text: e, onclick: () => onPick(e) })));
    grid.scrollTop = 0;
  };
  cats.forEach(c => tabs.append(h('button', {
    type: 'button', role: 'tab', class: 'emoji-tab', 'data-cat': c.name, 'aria-controls': 'emoji-grid', text: c.name,
    onclick: () => show(c.name)
  })));
  const panel = h('div', { class: 'emoji-panel', id: 'emoji-panel' }, tabs, grid);
  show(emojiState.tab);
  return panel;
}

function rememberEmoji(e) {
  emojiState.recent = [e].concat(emojiState.recent.filter(x => x !== e)).slice(0, 24);
}

/* ---------- Composer ---------- */
const rec = { recorder: null, chunks: [], started: 0, timer: null };

function renderComposer() {
  const wrap = $('composer-wrap');
  if (!wrap) return;
  wrap.replaceChildren();
  state.composerKey = composerKey(activeConv());
  const stopped = blockedComposer(activeConv());
  if (stopped) { emojiState.open = false; wrap.append(stopped); return; }
  const editing = state.editing && state.editing.convId === state.activeId ? state.editing : null;
  if (editing) {
    wrap.append(h('div', { class: 'replying editing' },
      h('div', { class: 'quote' }, h('b', { text: 'Edit message' }), h('span', { text: editing.original })),
      h('button', { class: 'icon-btn', 'aria-label': 'Cancel edit', onclick: cancelEdit }, icon('close'))));
  } else if (state.replyTo) {
    wrap.append(h('div', { class: 'replying' }, quoteNode(state.replyTo, state.activeId),
      h('button', { class: 'icon-btn', 'aria-label': 'Cancel reply', onclick: () => { state.replyTo = null; renderComposer(); } }, icon('close'))));
  }

  if (rec.recorder) {
    const status = h('span', { class: 'rec-status', id: 'rec-status', text: 'Recording 0:00' });
    wrap.append(h('div', { class: 'composer' },
      h('button', { class: 'icon-btn', 'aria-label': 'Discard voice note', onclick: () => stopRecording(false) }, icon('close')),
      status,
      h('button', { class: 'send recording', 'aria-label': 'Send voice note', onclick: () => stopRecording(true) }, icon('send'))));
    return;
  }

  const file = h('input', { type: 'file', accept: 'image/*', hidden: true, onchange: () => { pickerClosed(); if (file.files[0]) sendPhoto(file.files[0]); file.value = ''; } });
  // Phones: a separate button opens the camera straight away.
  const cam = h('input', { type: 'file', accept: 'image/*', capture: 'environment', hidden: true, onchange: () => { pickerClosed(); if (cam.files[0]) sendPhoto(cam.files[0]); cam.value = ''; } });
  const ta = h('textarea', { id: 'composer-text', rows: '1', placeholder: 'Message', 'aria-label': 'Message', maxlength: '4000', autocomplete: 'off' });
  const touch = !window.matchMedia('(hover: hover) and (pointer: fine)').matches;
  let sel = { start: 0, end: 0 };
  const keepSel = () => { sel = { start: ta.selectionStart, end: ta.selectionEnd }; };
  ['keyup', 'click', 'select', 'input'].forEach(ev => ta.addEventListener(ev, keepSel));
  const emojiBtn = h('button', {
    class: 'icon-btn emoji-btn', type: 'button',
    'aria-label': emojiState.open ? 'Show keyboard' : 'Emoji', 'aria-expanded': String(emojiState.open), 'aria-controls': 'emoji-panel'
  }, icon(emojiState.open ? 'keyboard' : 'smile'));
  emojiBtn.addEventListener('click', () => {
    emojiState.open = !emojiState.open;
    const draft = ta.value;
    renderComposer();
    const fresh = $('composer-text');
    fresh.value = draft;
    fresh.dispatchEvent(new Event('input'));
    if (!emojiState.open || !touch) fresh.focus();
  });
  // On phones, typing in the box brings the keyboard back instead of the emoji.
  ta.addEventListener('focus', () => {
    if (touch && emojiState.open) {
      emojiState.open = false;
      const panel = $('emoji-panel');
      if (panel) panel.remove();
      emojiBtn.replaceChildren(icon('smile'));
      emojiBtn.setAttribute('aria-label', 'Emoji');
      emojiBtn.setAttribute('aria-expanded', 'false');
    }
  });
  const pick = e => {
    const max = Number(ta.maxLength) || 4000;
    if (ta.value.length + e.length > max) return;
    const start = Math.min(sel.start, ta.value.length), end = Math.min(sel.end, ta.value.length);
    ta.value = ta.value.slice(0, start) + e + ta.value.slice(end);
    sel = { start: start + e.length, end: start + e.length };
    if (!touch) { ta.focus(); ta.setSelectionRange(sel.start, sel.end); }
    ta.dispatchEvent(new Event('input'));
    rememberEmoji(e);
  };
  const sendBtn = h('button', { class: 'send', 'aria-label': 'Voice note' }, icon('mic'));
  const sync = () => {
    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight, 140) + 'px';
    const hasText = ta.value.trim().length > 0;
    if (editing) {
      sendBtn.replaceChildren(icon('tick'));
      sendBtn.setAttribute('aria-label', 'Save edit');
      sendBtn.disabled = !hasText;
      return;
    }
    sendBtn.replaceChildren(icon(hasText ? 'send' : 'mic'));
    sendBtn.setAttribute('aria-label', hasText ? 'Send' : 'Record voice note');
  };
  ta.addEventListener('input', sync);
  ta.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey && window.matchMedia('(hover: hover)').matches) { e.preventDefault(); submit(); }
    if (e.key === 'Escape' && editing) { e.preventDefault(); cancelEdit(); }
  });
  const submit = () => {
    const text = ta.value.trim();
    if (editing) { if (text) saveEdit(editing, text); return; }
    if (!text) { startRecording(); return; }
    ta.value = '';
    sync();
    sendText(text);
    ta.focus();
  };
  sendBtn.addEventListener('click', submit);

  wrap.append(h('div', { class: 'composer' },
    h('div', { class: 'field' }, emojiBtn, ta,
      h('button', { class: 'icon-btn', 'aria-label': 'Send a photo from the gallery', title: 'Photo from gallery', onclick: () => { pickerOpened(); file.click(); } }, icon('image')), file,
      touch && !editing ? h('button', { class: 'icon-btn', 'aria-label': 'Take a photo', title: 'Camera', onclick: () => { pickerOpened(); cam.click(); } }, icon('camera')) : null,
      touch && !editing ? cam : null),
    sendBtn));
  if (emojiState.open) wrap.append(emojiPanel(pick));
  sync();
}

function startEdit(m) {
  state.replyTo = null;
  state.editing = { convId: m.convId, msgId: m.msgId, original: m.body };
  renderComposer();
  const ta = $('composer-text');
  if (!ta) return;
  ta.value = m.body;
  ta.dispatchEvent(new Event('input'));
  ta.focus();
  ta.setSelectionRange(ta.value.length, ta.value.length);
}

function cancelEdit() {
  state.editing = null;
  renderComposer();
}

/** Show the new text at once; put the old text back if the server refuses. */
async function saveEdit(editing, text) {
  state.editing = null;
  renderComposer();
  const map = state.msgs[editing.convId];
  const m = map && map.get(editing.msgId);
  if (!m || text === m.body) return;
  const before = { body: m.body, editedAt: m.editedAt };
  Object.assign(m, { body: text, editedAt: new Date().toISOString() });
  if (editing.convId === state.activeId) renderThread(false);
  try {
    const fresh = await api.call('editMessage', { convId: editing.convId, msgId: editing.msgId, body: text });
    applyEdits(editing.convId, [fresh]);
  } catch (err) {
    Object.assign(m, before);
    if (editing.convId === state.activeId) renderThread(false);
    toast(errorText(err));
  }
}

/** Edited messages from this device or from sync. */
function applyEdits(convId, list) {
  const map = state.msgs[convId];
  let changed = false;
  list.forEach(e => {
    const m = map && map.get(e.msgId);
    if (m && !m.deleted && (m.body !== e.body || m.editedAt !== e.editedAt)) { Object.assign(m, { body: e.body, editedAt: e.editedAt }); changed = true; }
  });
  const conv = state.convs.find(c => c.convId === convId);
  const last = conv && list.find(e => e.msgId === conv.lastMsgId);
  if (last) { conv.lastMsgPreview = previewOf(last); renderList(); }
  if (changed && convId === state.activeId) renderThread(false);
}

function startReply(m) {
  state.editing = null;
  state.replyTo = m.msgId;
  renderComposer();
  const ta = $('composer-text');
  if (ta) ta.focus();
}

function sendText(text) {
  const convId = state.activeId;
  const replyTo = state.replyTo;
  state.replyTo = null;
  renderComposer();
  api.sendText(convId, text, replyTo).catch(err => toast(errorText(err)));
}

async function sendPhoto(f) {
  const convId = state.activeId;
  try {
    const small = await compressImage(f);
    if (small.size > 5 * 1024 * 1024) { toast('That photo is still over 5 MB after compressing. Choose a smaller one.'); return; }
    toast('Sending photo…');
    await api.sendMedia(convId, 'image', small, '');
  } catch (err) {
    toast(errorText(err));
  }
}

async function startRecording() {
  if (!navigator.mediaDevices || !window.MediaRecorder) { toast('Voice notes need a browser with microphone support.'); return; }
  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const type = ['audio/webm', 'audio/ogg', 'audio/mp4'].find(t => MediaRecorder.isTypeSupported(t)) || '';
    rec.recorder = new MediaRecorder(stream, type ? { mimeType: type } : undefined);
    rec.chunks = [];
    rec.recorder.ondataavailable = e => { if (e.data.size) rec.chunks.push(e.data); };
    rec.recorder.start();
    rec.started = Date.now();
    renderComposer();
    rec.timer = setInterval(() => {
      const s = Math.floor((Date.now() - rec.started) / 1000);
      const el = $('rec-status');
      if (el) el.textContent = `Recording ${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
      if (s >= 120) stopRecording(true);
    }, 500);
  } catch {
    toast('Microphone access was blocked. Allow it in your browser settings to send voice notes.');
  }
}

function stopRecording(send) {
  const r = rec.recorder;
  if (!r) return;
  clearInterval(rec.timer);
  r.onstop = async () => {
    r.stream.getTracks().forEach(t => t.stop());
    const mime = (r.mimeType || 'audio/webm').split(';')[0];
    const blob = new Blob(rec.chunks, { type: mime });
    rec.recorder = null;
    renderComposer();
    if (send && blob.size) {
      try { await api.sendMedia(state.activeId, 'audio', new File([blob], 'voice', { type: mime })); }
      catch (err) { toast(errorText(err)); }
    }
  };
  r.stop();
}

function retry(m) {
  state.pending[m.convId].delete(m.clientId);
  renderThread(false);
  if (m.type === 'text') api.sendText(m.convId, m.body, m.replyTo).catch(err => toast(errorText(err)));
}

/* ---------- Names for group senders ---------- */
async function resolveGroupNames(conv) {
  try {
    const members = await api.call('listMembers', { convId: conv.convId });
    if (state.locked) return;
    state.members[conv.convId] = members;
    members.forEach(u => { state.names[u.userId] = u.displayName; });
  } catch { return; }
  if (state.activeId === conv.convId) renderThread(false);
  renderList();
}

/* ---------- Sheets ---------- */
function closeSheet() {
  const s = $('sheet');
  s.hidden = true;
  delete s.dataset.convId;
  s.replaceChildren();
}
function sheet(title, ...body) {
  const s = $('sheet');
  delete s.dataset.convId;   // openChatInfo sets it again after this
  s.replaceChildren(
    h('header', { class: 'bar' },
      h('button', { class: 'icon-btn', 'aria-label': 'Close', onclick: closeSheet }, icon('back')),
      h('h2', { class: 'grow', text: title })),
    h('div', { class: 'sheet-body' }, ...body));
  s.hidden = false;
}

function lookupBlock(onFound, label, noFocus) {
  const input = h('input', { id: 'lookup-phone', class: 'input', type: 'text', autocomplete: 'off', autocapitalize: 'off', placeholder: '0803 123 4567 or name@example.com' });
  const result = h('div', { 'aria-live': 'polite' });
  const btn = h('button', { class: 'btn secondary', type: 'submit', text: 'Find' });
  const form = h('form', {
    class: 'field-block',
    onsubmit: async e => {
      e.preventDefault();
      btn.disabled = true;
      result.replaceChildren();
      try {
        const u = await api.call('findUser', { query: input.value });
        if (!u) result.append(h('p', { class: 'hint', text: 'Nobody has joined with that number or email yet. Ask them to sign in to ZeeMo first.' }));
        else if (u.userId === state.me.userId) result.append(h('p', { class: 'hint', text: 'That is you.' }));
        else { state.names[u.userId] = u.displayName; onFound(u, result, input); }
      } catch (err) {
        result.append(h('p', { class: 'hint', text: errorText(err) }));
      }
      btn.disabled = false;
    }
  },
    h('label', { for: 'lookup-phone', text: label }),
    h('div', { class: 'inline' }, input, btn),
    DEMO ? h('p', { class: 'hint', text: 'Demo contacts: 0803 000 0101 to 0803 000 0104.' }) : null,
    h('p', { class: 'hint', text: 'Use the phone number or email they signed in with.' }),
    result);
  if (!noFocus) focusSoon(input);
  return form;
}

function openNewChat() {
  closeMenu();
  sheet('New chat',
    lookupBlock((u, result) => {
      result.append(h('div', { class: 'found' },
        avatar({ type: 'direct', title: u.displayName, peer: u, convId: u.userId }, true),
        h('div', { class: 'grow' }, h('strong', { text: u.displayName }), u.about ? h('small', { text: u.about }) : null),
        h('button', {
          class: 'btn', text: 'Message', onclick: async () => {
            try {
              const r = await api.call('createDirect', { otherUserId: u.userId });
              await loadConversations();
              closeSheet();
              openChat(r.convId);
            } catch (err) { toast(errorText(err)); }
          }
        })));
    }, 'Their phone number or email'),
    h('button', { class: 'btn secondary', text: 'Create a group instead', onclick: openNewGroup }));
}

function openNewGroup() {
  closeMenu();
  const members = new Map();
  const title = h('input', { id: 'group-title', class: 'input', maxlength: '60', placeholder: 'e.g. JSS 2 Class Teachers' });
  const chips = h('div', { class: 'chips' });
  const create = h('button', { class: 'btn', text: 'Create group', disabled: true });
  const known = state.convs.filter(c => c.type === 'direct' && c.peer).map(c => c.peer);

  const refresh = () => {
    chips.replaceChildren(...Array.from(members.values()).map(u => h('span', { class: 'chip' }, u.displayName,
      h('button', { 'aria-label': `Remove ${u.displayName}`, onclick: () => { members.delete(u.userId); refresh(); } }, icon('close')))));
    create.disabled = !members.size || !title.value.trim();
    picks.querySelectorAll('input').forEach(i => { i.checked = members.has(i.value); });
  };
  const picks = h('div', {}, ...known.map(u => h('label', { class: 'pick' },
    h('input', { type: 'checkbox', value: u.userId, onchange: e => { if (e.target.checked) members.set(u.userId, u); else members.delete(u.userId); refresh(); } }),
    avatar({ type: 'direct', title: u.displayName, peer: u, convId: u.userId }, true), u.displayName)));
  title.addEventListener('input', refresh);

  create.addEventListener('click', async () => {
    create.disabled = true;
    try {
      const r = await api.call('createGroup', { title: title.value, memberIds: Array.from(members.keys()) });
      await loadConversations();
      closeSheet();
      openChat(r.convId);
      toast('Group created');
    } catch (err) { toast(errorText(err)); refresh(); }
  });

  sheet('New group',
    h('div', { class: 'field-block' }, h('label', { for: 'group-title', text: 'Group name' }), title),
    h('div', { class: 'field-block' }, h('span', { class: 'section-title', text: 'Members' }), chips),
    known.length ? h('div', { class: 'field-block' }, h('span', { class: 'section-title', text: 'From your chats' }), picks) : null,
    lookupBlock((u, result, input) => { members.set(u.userId, u); input.value = ''; refresh(); result.append(h('p', { class: 'hint', text: `${u.displayName} added.` })); }, 'Add by phone number or email'),
    create);
}

function openProfile() {
  closeMenu();
  const name = h('input', { id: 'profile-name', class: 'input', maxlength: '40', value: state.me.displayName, autocomplete: 'off' });
  const about = h('textarea', { id: 'profile-about', class: 'input', maxlength: '140', placeholder: 'e.g. Mathematics teacher, JSS 3' });
  about.value = state.me.about || '';
  const save = h('button', { class: 'btn', text: 'Save profile' });
  save.addEventListener('click', async () => {
    save.disabled = true;
    try {
      state.me = await api.call('updateProfile', { displayName: name.value, about: about.value });
      toast('Profile saved');
    } catch (err) { toast(errorText(err)); }
    save.disabled = false;
  });
  const meAsConv = { type: 'direct', title: state.me.displayName, peer: state.me, convId: state.me.userId };
  sheet('Profile',
    h('div', { class: 'info-head' }, avatar(meAsConv),
      photoControls(state.me.avatarFileId, async upload => {
        state.me = await api.call('setAvatar', upload || { remove: true });
        toast(upload ? 'Photo updated' : 'Photo removed');
        openProfile();
      })),
    h('div', { class: 'field-block' }, h('label', { for: 'profile-name', text: 'Name' }), name),
    h('div', { class: 'field-block' }, h('label', { for: 'profile-about', text: 'About' }), about),
    state.me.phone ? h('div', { class: 'field-block' }, h('span', { class: 'section-title', text: 'Phone number' }), h('span', { text: state.me.phone })) : null,
    state.me.email ? h('div', { class: 'field-block' }, h('span', { class: 'section-title', text: 'Email' }), h('span', { text: state.me.email })) : null,
    save,
    lastSeenSection(),
    passcodeSection(),
    devicesSection(),
    blockedSection(),
    h('button', { class: 'btn danger', text: 'Sign out', onclick: async () => { closeSheet(); await api.logout().catch(() => {}); } }),
    deleteAccountSection(),
    legalLinks(),
    DEMO ? h('p', { class: 'hint', text: 'Demo mode keeps data in this browser only. Set apiUrl in config.js to use your Apps Script back end.' }) : null);
}

/* ---------- Message actions ---------- */
function myRole(convId) {
  const conv = state.convs.find(c => c.convId === convId);
  return conv ? conv.myRole : 'member';
}

/** Own text messages, within the owner's edit window. */
function canEdit(m) {
  const minutes = Number(state.security.editMinutes) || 0;
  return !!(minutes && m.msgId && !m.deleted && m.type === 'text' && m.senderId === state.me.userId &&
    Date.now() - new Date(m.createdAt).getTime() <= minutes * 60000);
}

function canDelete(m) {
  if (!m.msgId || m.deleted || m.type === 'system') return false;
  const conv = state.convs.find(c => c.convId === m.convId);
  if (conv && conv.type === 'group' && myRole(m.convId) === 'admin') return true;
  if (m.senderId !== state.me.userId) return false;
  const hours = Number(state.security.deleteWindowHours) || 0;
  return !hours || Date.now() - new Date(m.createdAt).getTime() <= hours * 3600000;
}

function closeActions() {
  const a = $('actions');
  if (a) { a.hidden = true; a.replaceChildren(); }
}

function openActions(m, confirming) {
  const a = $('actions');
  const preview = h('div', { class: 'quote' },
    h('b', { text: m.senderId === state.me.userId ? 'You' : (state.names[m.senderId] || 'Member') }),
    h('span', { text: previewOf(m) }));
  const buttons = confirming
    ? [
        h('p', { class: 'actions-note', text: m.senderId === state.me.userId
          ? 'Delete this message for everyone in the chat? It is erased and cannot be recovered.'
          : 'As a group admin, delete this message for everyone? It is erased and cannot be recovered.' }),
        h('button', { class: 'act danger', text: 'Delete for everyone', onclick: () => deleteForEveryone(m) }),
        h('button', { class: 'act', text: 'Cancel', onclick: closeActions })
      ]
    : [
        h('button', { class: 'act', onclick: () => { closeActions(); startReply(m); } }, icon('reply'), 'Reply'),
        canEdit(m) ? h('button', { class: 'act', onclick: () => { closeActions(); startEdit(m); } }, icon('edit'), 'Edit') : null,
        canDelete(m) ? h('button', { class: 'act danger', onclick: () => openActions(m, true) }, icon('trash'), 'Delete for everyone') : null,
        m.senderId !== state.me.userId && m.msgId && !m.deleted && m.type !== 'system'
          ? h('button', { class: 'act danger', onclick: () => openReport({ convId: m.convId, msgId: m.msgId, userId: m.senderId, name: state.names[m.senderId] || 'Member', alreadyBlocked: blockedIds().has(m.senderId) }) }, icon('flag'), 'Report')
          : null,
        h('button', { class: 'act', text: 'Cancel', onclick: closeActions })
      ];
  a.replaceChildren(
    h('div', { class: 'actions-backdrop', onclick: closeActions }),
    h('div', { class: 'actions-panel', role: 'dialog', 'aria-label': 'Message options' }, preview, ...buttons.filter(Boolean)));
  a.hidden = false;
  const first = a.querySelector('.act');
  if (first) first.focus();
}

async function deleteForEveryone(m) {
  closeActions();
  try {
    const tomb = await api.deleteMessage(m.convId, m.msgId);
    mergeMessages(m.convId, [tomb]);
    const conv = state.convs.find(c => c.convId === m.convId);
    if (conv && conv.lastMsgId === m.msgId) conv.lastMsgPreview = 'This message was deleted';
    if (m.mediaFileId && state.media[m.mediaFileId]) {
      URL.revokeObjectURL(state.media[m.mediaFileId]);
      delete state.media[m.mediaFileId];
    }
    renderList();
    if (state.activeId === m.convId) renderThread(false);
    toast('Message deleted for everyone');
  } catch (err) {
    toast(errorText(err));
  }
}

/* ---------- Chat info: group management and disappearing messages ---------- */
function confirmButton(label, confirmLabel, action, cls) {
  const btn = h('button', { class: cls || 'btn danger', text: label });
  let armed = false;
  btn.addEventListener('click', () => {
    if (!armed) {
      armed = true;
      btn.textContent = confirmLabel;
      btn.classList.add('armed');
      setTimeout(() => { if (btn.isConnected && armed) { armed = false; btn.textContent = label; btn.classList.remove('armed'); } }, 4000);
      return;
    }
    armed = false;
    btn.disabled = true;
    action();
  });
  return btn;
}

async function runAction(promise, done) {
  try {
    await promise;
    if (done) toast(done);
    await loadConversations();
  } catch (err) {
    toast(errorText(err));
    const id = $('sheet').dataset.convId;
    if (id && state.convs.some(c => c.convId === id)) openChatInfo(id, true);
  }
}

async function openChatInfo(convId, refresh) {
  const conv = state.convs.find(c => c.convId === convId);
  if (!conv) { closeSheet(); return; }
  const isGroup = conv.type === 'group';
  const admin = isGroup && conv.myRole === 'admin';
  const canTimer = !isGroup || admin;

  if (isGroup && (!state.members[convId] || refresh)) {
    try { state.members[convId] = await api.call('listMembers', { convId }); } catch (err) { if (!QUIET_ERRORS.includes(err.code)) toast(errorText(err)); return; }
    state.members[convId].forEach(u => { state.names[u.userId] = u.displayName; });
  }
  if (state.locked) return;

  const body = [];
  body.push(h('div', { class: 'info-head' },
    avatar(conv),
    h('div', { class: 'grow' },
      h('strong', { text: conv.title }),
      h('small', { text: isGroup ? `Group · ${conv.memberCount} members` : ((conv.peer && conv.peer.about) || 'Direct chat') }))));
  if (admin) {
    body.push(photoControls(conv.avatarFileId, upload => runAction(
      api.call('setGroupAvatar', Object.assign({ convId }, upload || { remove: true })),
      upload ? 'Group photo updated' : 'Group photo removed')));
  }

  if (admin) {
    const title = h('input', { id: 'info-title', class: 'input', maxlength: '60', value: conv.title, autocomplete: 'off' });
    const save = h('button', { class: 'btn secondary', type: 'submit', text: 'Rename' });
    body.push(h('form', {
      class: 'field-block',
      onsubmit: e => {
        e.preventDefault();
        if (!title.value.trim() || title.value.trim() === conv.title) return;
        save.disabled = true;
        runAction(api.call('renameGroup', { convId, title: title.value }), 'Group renamed');
      }
    }, h('label', { for: 'info-title', text: 'Group name' }), h('div', { class: 'inline' }, title, save)));
  }

  // Disappearing messages
  const opts = state.security.disappearOptions || [0, 86400, 604800, 7776000];
  const current = Number(conv.disappearSec) || 0;
  body.push(h('fieldset', { class: 'field-block timer-set', disabled: !canTimer },
    h('legend', { class: 'section-title', text: 'Disappearing messages' }),
    h('p', { class: 'hint', text: canTimer
      ? 'New messages disappear from this chat for everyone after the time you choose. Messages already sent are not affected.'
      : 'New messages disappear from this chat after this time. Only admins can change it.' }),
    h('div', { class: 'radios', role: 'radiogroup' },
      ...opts.map(sec => h('label', { class: 'radio' },
        h('input', {
          type: 'radio', name: 'timer', value: String(sec), checked: sec === current,
          onchange: () => runAction(api.call('setDisappearing', { convId, seconds: sec }),
            sec ? `Messages will disappear after ${timerLabel(sec).toLowerCase()}` : 'Disappearing messages off')
        }),
        h('span', { text: timerLabel(sec) }))))));

  if (!isGroup) body.push(...peerSafetyNodes(conv));

  if (isGroup) {
    const members = (state.members[convId] || []).slice().sort((a, b) =>
      (a.userId === state.me.userId ? -1 : b.userId === state.me.userId ? 1 : 0) ||
      (a.role === b.role ? a.displayName.localeCompare(b.displayName) : a.role === 'admin' ? -1 : 1));
    body.push(h('div', { class: 'field-block' },
      h('span', { class: 'section-title', text: `${members.length} members` }),
      h('ul', { class: 'members' }, ...members.map(u => {
        const me = u.userId === state.me.userId;
        const actions = admin && !me ? h('div', { class: 'member-actions' },
          h('button', {
            class: 'btn small secondary',
            text: u.role === 'admin' ? 'Dismiss as admin' : 'Make admin',
            onclick: e => { e.target.disabled = true; runAction(api.call('setMemberRole', { convId, userId: u.userId, role: u.role === 'admin' ? 'member' : 'admin' })); }
          }),
          confirmButton('Remove', `Remove ${u.displayName.split(' ')[0]}?`, () => runAction(api.call('removeMember', { convId, userId: u.userId }), `${u.displayName} removed`), 'btn small danger')
        ) : null;
        return h('li', { class: 'member' },
          avatar({ type: 'direct', title: u.displayName, peer: u, convId: u.userId }, true),
          h('div', { class: 'grow' },
            h('strong', { text: me ? `${u.displayName} (you)` : u.displayName }),
            u.about ? h('small', { text: u.about }) : null),
          u.role === 'admin' ? h('span', { class: 'badge', text: 'Admin' }) : null,
          actions,
          !me ? h('button', { class: 'icon-btn report-btn', 'aria-label': `Report ${u.displayName}`, title: 'Report',
            onclick: () => openReport({ convId, userId: u.userId, name: u.displayName, alreadyBlocked: u.blockedByMe }) }, icon('flag')) : null);
      }))));

    if (admin) {
      const inGroup = new Set(members.map(u => u.userId));
      const known = state.convs.filter(c => c.type === 'direct' && c.peer && !inGroup.has(c.peer.userId)).map(c => c.peer);
      const picked = new Map();
      const addBtn = h('button', { class: 'btn', text: 'Add to group', disabled: true });
      const chips = h('div', { class: 'chips' });
      const refreshPicks = () => {
        chips.replaceChildren(...Array.from(picked.values()).map(u => h('span', { class: 'chip' }, u.displayName,
          h('button', { 'aria-label': `Remove ${u.displayName}`, onclick: () => { picked.delete(u.userId); refreshPicks(); } }, icon('close')))));
        addBtn.disabled = !picked.size;
        addBtn.textContent = picked.size > 1 ? `Add ${picked.size} people` : 'Add to group';
      };
      addBtn.addEventListener('click', () => {
        addBtn.disabled = true;
        runAction(api.call('addMembers', { convId, memberIds: Array.from(picked.keys()) }), 'Added to group');
      });
      body.push(h('details', { class: 'add-members' },
        h('summary', { text: 'Add members' }),
        known.length ? h('div', { class: 'field-block' },
          h('span', { class: 'section-title', text: 'From your chats' }),
          ...known.map(u => h('label', { class: 'pick' },
            h('input', { type: 'checkbox', value: u.userId, onchange: e => { if (e.target.checked) picked.set(u.userId, u); else picked.delete(u.userId); refreshPicks(); } }),
            avatar({ type: 'direct', title: u.displayName, peer: u, convId: u.userId }, true), u.displayName))) : null,
        lookupBlock((u, result, input) => {
          if (inGroup.has(u.userId)) { result.append(h('p', { class: 'hint', text: `${u.displayName} is already in this group.` })); return; }
          picked.set(u.userId, u); input.value = ''; refreshPicks();
          result.append(h('p', { class: 'hint', text: `${u.displayName} will be added.` }));
        }, 'Add by phone number or email', true),
        chips,
        addBtn));
    }

    body.push(confirmButton('Leave group', 'Tap again to leave', async () => {
      try {
        await api.call('leaveGroup', { convId });
        closeSheet();
        closeChat();
        toast(`You left ${conv.title}`);
        await loadConversations();
      } catch (err) { toast(errorText(err)); }
    }));
  }

  const sheetEl = $('sheet');
  const scroll = refresh ? (sheetEl.querySelector('.sheet-body') || {}).scrollTop : 0;
  sheet(isGroup ? 'Group info' : 'Chat info', ...body);
  sheetEl.dataset.convId = convId;
  if (scroll) sheetEl.querySelector('.sheet-body').scrollTop = scroll;
}

/* ---------- Terms, reporting, blocking and account deletion ---------- */
const REPORT_REASONS = [
  ['spam', 'Spam'], ['harassment', 'Harassment or bullying'], ['hate', 'Hate speech'],
  ['sexual', 'Sexual content'], ['scam', 'Scam or fraud'], ['other', 'Something else']
];
const legalCache = {};

async function legalDoc(page) {
  if (!legalCache[page]) legalCache[page] = await api.call('legalText', { page });
  return legalCache[page];
}

function legalNodes(doc) {
  return doc.sections.map(([title, paras]) => h('section', {},
    h('h3', { text: title }), ...paras.map(t => h('p', { text: t }))));
}

/** Open the Terms or Privacy page: the public web page when there is one, otherwise in the app. */
async function openLegal(page) {
  const links = state.links || {};
  const url = page === 'privacy' ? links.privacy : links.terms;
  if (url) { window.open(url, '_blank', 'noopener'); return; }
  try {
    const doc = await legalDoc(page);
    sheet(doc.title, h('div', { class: 'legal' }, ...legalNodes(doc)));
  } catch (err) { toast(errorText(err)); }
}

function legalLinks() {
  return h('p', { class: 'legal-links' },
    h('button', { class: 'link', type: 'button', text: 'Terms of use', onclick: () => openLegal('terms') }),
    h('span', { text: ' · ' }),
    h('button', { class: 'link', type: 'button', text: 'Privacy notice', onclick: () => openLegal('privacy') }));
}

/** After sign-in or on opening: terms first, then the passcode, then the chats. */
function nextStep(how) {
  if (!state.me.termsAccepted) { renderTerms(how); return; }
  if (how === 'resume') { $('auth').hidden = true; boot(); return; }
  if (state.me.hasPasscode) {
    if (how === 'code') { state.locked = true; wipeScreen(); renderLock(); }
    else lockApp({});
  } else if (state.security.requirePasscode) renderSetPasscode();
  else { $('auth').hidden = true; boot(); }
}

/** how: 'code' after a code sign-in, 'start' on opening, 'resume' when new terms appear mid-session. */
async function renderTerms(how) {
  const root = $('auth');
  root.hidden = false;
  root.dataset.mode = 'terms';
  const box = h('div', { class: 'terms-box', tabindex: '0', 'aria-label': 'Terms of use' }, h('p', { text: 'Loading…' }));
  const agree = h('input', { type: 'checkbox', id: 'terms-agree' });
  const btn = h('button', { class: 'btn', type: 'submit', text: 'Agree and continue', disabled: true });
  const err = h('p', { class: 'error', role: 'alert' });
  agree.addEventListener('change', () => { btn.disabled = !agree.checked; });
  root.replaceChildren(h('div', { class: 'auth-card terms-card' },
    h('div', { class: 'wordmark', 'aria-label': 'ZeeMo' }, h('b', { text: 'Zee' }), h('span', { text: 'Mo' })),
    h('p', { class: 'auth-lede', text: how === 'resume' ? 'The terms of use have changed. Please read and accept them to carry on.' : 'Before you start, please read the terms of use.' }),
    box,
    h('form', {
      onsubmit: async e => {
        e.preventDefault();
        if (!agree.checked) return;
        btn.disabled = true; btn.textContent = 'Saving…';
        try {
          await api.call('acceptTerms');
          state.me.termsAccepted = true;
          nextStep(how);
        } catch (ex) {
          btn.disabled = false; btn.textContent = 'Agree and continue';
          err.textContent = errorText(ex);
        }
      }
    },
      h('label', { class: 'check', for: 'terms-agree' }, agree,
        h('span', { text: 'I am at least 13 years old and I agree to the terms of use. Abusive content and users are not tolerated.' })),
      h('button', { class: 'link', type: 'button', text: 'Read the privacy notice', onclick: () => showPrivacyInline() }),
      err, btn,
      h('button', { class: 'link', type: 'button', text: 'Sign out', onclick: () => api.logout().catch(() => {}) }))));

  async function showPrivacyInline() {
    const links = state.links || {};
    if (links.privacy) { window.open(links.privacy, '_blank', 'noopener'); return; }
    try { box.replaceChildren(...legalNodes(await legalDoc('privacy'))); box.scrollTop = 0; } catch (ex) { err.textContent = errorText(ex); }
  }
  try {
    box.replaceChildren(...legalNodes(await legalDoc('terms')));
  } catch (ex) {
    box.replaceChildren(h('p', { text: 'The terms could not be loaded. ' + errorText(ex) }));
  }
}

/** People this user has blocked, as far as the chat list knows. */
function blockedIds() {
  return new Set(state.convs.filter(c => c.peer && c.peer.blockedByMe).map(c => c.peer.userId));
}

/** Report a message (target.msgId) or a person (target.userId). */
function openReport(target) {
  closeActions();
  const first = target.name ? firstName(target.name) : 'this person';
  let reason = '';
  const send = h('button', { class: 'btn', text: 'Send report', disabled: true });
  const details = h('textarea', { id: 'report-details', class: 'input', maxlength: '500', placeholder: 'Anything that helps the administrator (optional)' });
  const block = h('input', { type: 'checkbox', checked: true });
  send.addEventListener('click', async () => {
    send.disabled = true; send.textContent = 'Sending…';
    try {
      const payload = { convId: target.convId, reason, details: details.value, block: !!(block.checked && target.userId) };
      if (target.msgId) payload.msgId = target.msgId; else payload.userId = target.userId;
      const r = await api.call('reportContent', payload);
      closeSheet();
      toast(r.blocked ? `Report sent and ${first} blocked` : 'Report sent. The administrator will review it.');
      if (r.blocked) await loadConversations();
    } catch (err) {
      send.disabled = false; send.textContent = 'Send report';
      toast(errorText(err));
    }
  });
  sheet(target.msgId ? 'Report message' : `Report ${first}`,
    h('p', { class: 'hint', text: target.msgId
      ? `The administrator will see this message and who sent it. ${first} is not told who reported it.`
      : `The administrator will review ${first}'s account. ${first} is not told who reported it.` }),
    h('fieldset', { class: 'field-block timer-set' },
      h('legend', { class: 'section-title', text: 'What is wrong?' }),
      h('div', { class: 'radios one' }, ...REPORT_REASONS.map(([value, label]) => h('label', { class: 'radio' },
        h('input', { type: 'radio', name: 'report-reason', value, onchange: () => { reason = value; send.disabled = false; } }),
        h('span', { text: label }))))),
    h('div', { class: 'field-block' }, h('label', { for: 'report-details', text: 'Details' }), details),
    target.userId && !target.alreadyBlocked ? h('label', { class: 'pick' }, block, `Also block ${first}`) : null,
    send);
}

async function setBlocked(userId, blocked, name) {
  try {
    await api.call(blocked ? 'blockUser' : 'unblockUser', { userId });
    toast(blocked ? `${name} blocked` : `${name} unblocked`);
    await loadConversations();
    return true;
  } catch (err) { toast(errorText(err)); return false; }
}

/** Direct-chat info: block or unblock, and report. */
function peerSafetyNodes(conv) {
  const p = conv.peer;
  if (!p || p.deleted) return [];
  const first = firstName(p.displayName);
  return [
    h('div', { class: 'field-block safety' },
      h('span', { class: 'section-title', text: 'Safety' }),
      p.blockedByMe
        ? h('button', { class: 'btn secondary', text: `Unblock ${first}`, onclick: async () => { if (await setBlocked(p.userId, false, first)) openChatInfo(conv.convId); } })
        : confirmButton(`Block ${first}`, `Tap again to block ${first}`, async () => { if (await setBlocked(p.userId, true, first)) openChatInfo(conv.convId); }),
      h('p', { class: 'hint', text: p.blockedByMe
        ? `${first} cannot reach you. Messages they send while blocked are never shown to you.`
        : `Blocked people cannot message you or see when you were last online. They are not told.` }),
      h('button', { class: 'btn danger', text: `Report ${first}`, onclick: () => openReport({ convId: conv.convId, userId: p.userId, name: p.displayName, alreadyBlocked: p.blockedByMe }) }))
  ];
}

function composerKey(conv) {
  const p = conv && conv.type === 'direct' ? conv.peer : null;
  return p ? (p.deleted ? 'deleted' : p.blockedByMe ? 'blocked' : '') : '';
}

/** In place of the message box when you blocked the person, or their account was deleted. */
function blockedComposer(conv) {
  const p = conv && conv.type === 'direct' ? conv.peer : null;
  if (!p) return null;
  if (p.deleted) return h('div', { class: 'composer-note' }, h('span', { text: 'This account has been deleted. You can no longer send messages here.' }));
  if (!p.blockedByMe) return null;
  const first = firstName(p.displayName);
  return h('div', { class: 'composer-note' },
    h('span', { class: 'grow', text: `You blocked ${first}.` }),
    h('button', { class: 'btn small', text: 'Unblock', onclick: () => setBlocked(p.userId, false, first) }));
}

/** Profile: people you have blocked. */
function blockedSection() {
  const list = h('ul', { class: 'members' });
  const wrap = h('details', { class: 'add-members' }, h('summary', { text: 'Blocked people' }), list);
  wrap.addEventListener('toggle', async () => {
    if (!wrap.open) return;
    list.replaceChildren(h('li', { class: 'hint', text: 'Loading…' }));
    try {
      const people = await api.call('listBlocked');
      list.replaceChildren(...(people.length ? people.map(u => h('li', { class: 'member' },
        avatar({ type: 'direct', title: u.displayName, peer: u, convId: u.userId }, true),
        h('div', { class: 'grow' }, h('strong', { text: u.displayName })),
        h('button', { class: 'btn small secondary', text: 'Unblock', onclick: async e => {
          e.target.disabled = true;
          if (await setBlocked(u.userId, false, u.displayName)) e.target.closest('li').remove();
          else e.target.disabled = false;
        } }))) : [h('li', { class: 'hint', text: 'You have not blocked anyone.' })]));
    } catch (err) { list.replaceChildren(h('li', { class: 'hint', text: errorText(err) })); }
  });
  return wrap;
}

/** Profile: delete my account. */
function deleteAccountSection() {
  const wrap = h('details', { class: 'add-members danger-zone' }, h('summary', { text: 'Delete my account' }));
  const err = h('p', { class: 'error-dark', role: 'alert' });
  let pass = null;
  if (state.me.hasPasscode) {
    pass = h('input', { id: 'delete-pass', class: 'input', type: 'password', inputmode: 'numeric', maxlength: '6', autocomplete: 'off', placeholder: 'Your passcode' });
  }
  const go = confirmButton('Delete my account', 'Tap again to delete for good', async () => {
    err.textContent = '';
    try {
      await api.call('deleteAccount', { confirm: 'DELETE', passcode: pass ? pass.value : '' });
      wipeScreen();
      state.me = null;
      api.signOutHere('Your account has been deleted.');
    } catch (ex) {
      err.textContent = errorText(ex);
      go.disabled = false; go.textContent = 'Delete my account'; go.classList.remove('armed');
      if (pass) { pass.value = ''; pass.focus(); }
    }
  });
  appendAll(wrap,
    h('p', { class: 'hint', text: 'This erases your profile, your phone number or email, your passcode and every message, photo and voice note you sent. You leave all your chats and are signed out everywhere. It cannot be undone.' }),
    pass ? h('div', { class: 'field-block' }, h('label', { for: 'delete-pass', text: 'Passcode' }), pass) : null,
    err, go);
  return wrap;
}

/** Chat list banner when someone asked on the web page to delete this account. */
function deletionNotice() {
  const at = state.me && state.me.deletionRequestedAt;
  if (!at) return false;
  const when = new Date(new Date(at).getTime() + 7 * 86400000);
  const box = $('notice');
  box.replaceChildren(
    h('span', { class: 'grow', text: `Someone asked to delete this account. It will be deleted on ${when.toLocaleDateString('en-GB', { day: 'numeric', month: 'long' })} unless you cancel.` }),
    h('button', { class: 'btn small', text: 'Cancel deletion', onclick: async e => {
      e.target.disabled = true;
      try { await api.call('cancelDeletion'); state.me.deletionRequestedAt = ''; box.hidden = true; toast('Deletion cancelled'); }
      catch (err) { e.target.disabled = false; toast(errorText(err)); }
    } }));
  box.hidden = false;
  return true;
}

/* ---------- Security: passcode, lock, idle and minimise ---------- */
const guard = { lastActivity: Date.now(), activeSincePoll: true, pickerOpen: false, pickerTimer: null };

function markActive() { guard.lastActivity = Date.now(); guard.activeSincePoll = true; }

/** Reported with each poll so the server only keeps an attended app unlocked. */
function activeSinceLastPoll() {
  if (!(Number(state.security.lockAfterMinutes) > 0)) return true;   // idle lock off: stay unlocked while open
  const was = guard.activeSincePoll;
  guard.activeSincePoll = false;
  return !!was;
}

/** The file picker hides the page on phones; that must not lock the app mid-upload. */
function pickerOpened() {
  guard.pickerOpen = true;
  clearTimeout(guard.pickerTimer);
  guard.pickerTimer = setTimeout(() => { guard.pickerOpen = false; }, 180000);
}
function pickerClosed() {
  clearTimeout(guard.pickerTimer);
  setTimeout(() => { guard.pickerOpen = false; }, 800);
}

/** Remove every message, preview and photo from the screen and from memory. */
function wipeScreen() {
  if (rec.recorder) stopRecording(false);
  api.wipe();
  closeActions();
  closeSheet();
  closeMenu();
  state.activeId = null;
  state.replyTo = null;
  state.editing = null;
  state.convs = [];
  state.msgs = {};
  state.pending = {};
  state.receipts = {};
  state.members = {};
  state.filter = '';
  emojiState.open = false;
  emojiState.recent = [];
  forgetMedia();
  forgetPhotos();
  document.body.classList.remove('chat-open');
  $('search-input').value = '';
  $('list').replaceChildren();
  $('toast').hidden = true;
  $('notice').hidden = true;
  renderNoChat();
}

/** Lock: wipe everything and require the passcode. Without a passcode it closes the chat instead. */
function lockApp(opts) {
  if (!state.me) return;
  if (!state.me.hasPasscode) {
    if (state.activeId) closeChat();
    return;
  }
  const already = state.locked;
  state.locked = true;
  wipeScreen();
  if (!opts.serverAlreadyLocked && !already) api.lock();
  renderLock();
}

/** The owner's passcode rule, e.g. 6 digits or 4 to 6 digits. */
function passRule() {
  const min = Math.min(6, Math.max(4, Number(state.security.passcodeMinLength) || 4));
  return { min, label: min === 6 ? '6 digits' : `${min} to 6 digits`, re: new RegExp(`^[0-9]{${min},6}$`) };
}

function passField(id, label, autoFocus) {
  const input = h('input', {
    id, class: 'input code', type: 'password', inputmode: 'numeric', pattern: '[0-9]*', maxlength: '6',
    autocomplete: 'off', placeholder: '••••••', required: true
  });
  if (autoFocus) focusSoon(input);
  return [h('label', { for: id, text: label }), input];
}

function renderLock(errorMsg) {
  const root = $('auth');
  root.hidden = false;
  root.dataset.mode = 'lock';
  const [label, input] = passField('lock-pass', 'Passcode', true);
  const btn = h('button', { class: 'btn', type: 'submit', text: 'Unlock' });
  const err = h('p', { class: 'error', role: 'alert', text: errorMsg || '' });
  const forgot = h('div', { class: 'forgot', hidden: true },
    h('p', { class: 'note', text: 'Ask your administrator to reset your passcode. Then sign in again with a code and choose a new one. Your chats stay safe on the server.' }),
    h('div', { class: 'inline' },
      h('button', { class: 'btn', type: 'button', text: 'Sign out', onclick: async () => {
        state.locked = false;
        await api.logout().catch(() => {});
      } }),
      h('button', { class: 'btn ghost', type: 'button', text: 'Cancel', onclick: () => { forgot.hidden = true; input.focus(); } })));

  root.replaceChildren(h('div', { class: 'auth-card' },
    h('div', { class: 'lock-mark' }, icon('lock')),
    h('div', { class: 'wordmark', 'aria-label': 'ZeeMo' }, h('b', { text: 'Zee' }), h('span', { text: 'Mo' })),
    h('p', { class: 'auth-lede', text: `Locked. Enter your passcode to open your chats, ${state.me.displayName.split(' ')[0]}.` }),
    h('form', {
      onsubmit: async e => {
        e.preventDefault();
        btn.disabled = true; btn.textContent = 'Checking…';
        try {
          const typed = input.value;
          const r = await api.unlock(typed);
          if (r && r.mustChange) renderLongerPasscode(typed);
          else await boot();
        } catch (ex) {
          if (!api.signedIn) return;   // too many wrong tries: signed out
          btn.disabled = false; btn.textContent = 'Unlock';
          err.textContent = errorText(ex);
          input.value = '';
          input.focus();
        }
      }
    }, label, input, err, btn,
      h('button', { class: 'link', type: 'button', text: 'Forgot passcode?', onclick: () => { forgot.hidden = false; } })),
    forgot));
}

/** The owner raised the minimum length: the person must replace a shorter passcode before going on. */
function renderLongerPasscode(current) {
  const root = $('auth');
  root.hidden = false;
  root.dataset.mode = 'passcode';
  const rule = passRule();
  const [l1, p1] = passField('long-pass', `New passcode (${rule.label})`, true);
  const [l2, p2] = passField('long-pass-2', 'Enter it again');
  const err = h('p', { class: 'error', role: 'alert' });
  const btn = h('button', { class: 'btn', type: 'submit', text: 'Save passcode' });
  root.replaceChildren(h('div', { class: 'auth-card' },
    h('div', { class: 'lock-mark' }, icon('lock')),
    h('div', { class: 'wordmark', 'aria-label': 'ZeeMo' }, h('b', { text: 'Zee' }), h('span', { text: 'Mo' })),
    h('p', { class: 'auth-lede', text: `Your administrator now asks for longer passcodes. Choose a new one with ${rule.label} to continue.` }),
    h('form', {
      onsubmit: async e => {
        e.preventDefault();
        if (!rule.re.test(p1.value)) { err.textContent = `Use ${rule.label}.`; p1.focus(); return; }
        if (p1.value !== p2.value) { err.textContent = 'The two passcodes do not match.'; p2.value = ''; p2.focus(); return; }
        btn.disabled = true; btn.textContent = 'Saving…';
        try {
          await api.setPasscode(p1.value, current);
          current = '';
          await boot();
          toast('Passcode saved');
        } catch (ex) {
          btn.disabled = false; btn.textContent = 'Save passcode';
          err.textContent = errorText(ex);
        }
      }
    }, l1, p1, l2, p2, err, btn)));
}

/** Shown when a passcode is required and none is set (new users, or after an administrator reset). */
function renderSetPasscode() {
  const root = $('auth');
  root.hidden = false;
  root.dataset.mode = 'passcode';
  const rule = passRule();
  const [l1, p1] = passField('new-pass', `New passcode (${rule.label})`, true);
  const [l2, p2] = passField('new-pass-2', 'Enter it again');
  const err = h('p', { class: 'error', role: 'alert' });
  const btn = h('button', { class: 'btn', type: 'submit', text: 'Save passcode' });
  const lede = 'Set a passcode. You will need it to open ZeeMo, whenever it locks, and when you sign in on another phone or computer, so nobody else can read your chats.';

  root.replaceChildren(h('div', { class: 'auth-card' },
    h('div', { class: 'lock-mark' }, icon('lock')),
    h('div', { class: 'wordmark', 'aria-label': 'ZeeMo' }, h('b', { text: 'Zee' }), h('span', { text: 'Mo' })),
    h('p', { class: 'auth-lede', text: lede }),
    h('form', {
      onsubmit: async e => {
        e.preventDefault();
        if (!rule.re.test(p1.value)) { err.textContent = `Use ${rule.label}.`; p1.focus(); return; }
        if (p1.value !== p2.value) { err.textContent = 'The two passcodes do not match.'; p2.value = ''; p2.focus(); return; }
        btn.disabled = true; btn.textContent = 'Saving…';
        try {
          await api.setPasscode(p1.value);
          state.me.hasPasscode = true;
          await boot();
          toast('Passcode saved');
        } catch (ex) {
          btn.disabled = false; btn.textContent = 'Save passcode';
          err.textContent = errorText(ex);
        }
      }
    }, l1, p1, l2, p2, err, btn,
      h('button', { class: 'link', type: 'button', text: 'Sign out', onclick: () => api.logout().catch(() => {}) }))));
}

/* ---------- Privacy: last seen ---------- */
function lastSeenSection() {
  const wrap = h('fieldset', { class: 'field-block timer-set' },
    h('legend', { class: 'section-title', text: 'Who can see when I was last online' }));
  if (state.security.showLastSeen === false) {
    wrap.append(h('p', { class: 'hint', text: 'The administrator has hidden last seen for everyone.' }));
    return wrap;
  }
  const current = state.me.lastSeenPrivacy === 'nobody' ? 'nobody' : 'contacts';
  wrap.append(
    h('div', { class: 'radios', role: 'radiogroup' },
      ...[['contacts', 'People I chat with'], ['nobody', 'Nobody']].map(([value, label]) => h('label', { class: 'radio' },
        h('input', {
          type: 'radio', name: 'lastseen', value, checked: value === current,
          onchange: async () => {
            try {
              state.me = await api.call('updateProfile', { lastSeenPrivacy: value });
              toast(value === 'nobody' ? 'Last seen hidden' : 'Last seen shown to people you chat with');
            } catch (err) { toast(errorText(err)); }
          }
        }),
        h('span', { text: label })))),
    h('p', { class: 'hint', text: 'People who look up your number or email never see it.' }));
  return wrap;
}

/* ---------- Devices ---------- */
const sessionTime = new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' });

function deviceName(d) {
  // Older sign-ins stored the browser's raw description.
  return /Mozilla|AppleWebKit/.test(d.device || '') ? deviceLabel(d.device) : (d.device || 'Unknown device');
}

function devicesSection() {
  const wrap = h('details', { class: 'add-members devices' }, h('summary', { text: 'Devices signed in' }));
  const list = h('ul', { class: 'members' });
  const load = async () => {
    list.replaceChildren(h('li', { class: 'hint', text: 'Loading…' }));
    try {
      const devices = await api.call('listSessions');
      list.replaceChildren(...devices.map(d => h('li', { class: 'member' },
        h('span', { class: 'avatar sm device-icon', 'aria-hidden': 'true' }, icon('lock')),
        h('div', { class: 'grow' },
          h('strong', { text: deviceName(d) }),
          h('small', { text: (d.current ? 'This device · ' : '') + 'signed in ' + sessionTime.format(new Date(d.createdAt)) })),
        d.current ? null : h('button', {
          class: 'btn small danger', text: 'Sign out',
          onclick: async e => {
            e.target.disabled = true;
            try { await api.call('revokeSessions', { id: d.id }); toast('Device signed out'); load(); }
            catch (err) { toast(errorText(err)); e.target.disabled = false; }
          }
        }))));
      if (devices.filter(d => !d.current).length > 1) {
        list.append(h('li', {}, confirmButton('Sign out all other devices', 'Tap again to sign them out', async () => {
          try { await api.call('revokeSessions', { others: true }); toast('Other devices signed out'); load(); }
          catch (err) { toast(errorText(err)); }
        })));
      }
    } catch (err) {
      list.replaceChildren(h('li', { class: 'hint', text: errorText(err) }));
    }
  };
  wrap.addEventListener('toggle', () => { if (wrap.open) load(); });
  wrap.append(h('p', { class: 'hint', text: 'Every phone or computer where your account is open. Sign out any you do not recognise.' }), list);
  return wrap;
}

/** After unlocking: warn about other devices that signed in during the last week. */
async function checkNewSignIns() {
  if (deletionNotice()) return;
  try {
    const devices = await api.call('listSessions');
    if (state.locked) return;
    const week = Date.now() - 7 * 86400000;
    const others = devices.filter(d => !d.current && new Date(d.createdAt).getTime() > week);
    const box = $('notice');
    if (!others.length) { box.hidden = true; return; }
    const newest = others[0];
    box.replaceChildren(
      h('span', { class: 'grow', text: others.length === 1
        ? `Your account was also opened on ${deviceName(newest)} on ${sessionTime.format(new Date(newest.createdAt))}. Not you?`
        : `Your account was opened on ${others.length} other devices this week. Not you?` }),
      h('button', { class: 'btn small', text: 'Review', onclick: () => { box.hidden = true; openProfile(); const d = document.querySelector('#sheet details.devices'); if (d) { d.open = true; d.scrollIntoView(); } } }),
      h('button', { class: 'icon-btn', 'aria-label': 'Dismiss', onclick: () => { box.hidden = true; } }, icon('close')));
    box.hidden = false;
  } catch { /* not important enough to interrupt */ }
}

/** Profile: change passcode. */
function passcodeSection() {
  const wrap = h('details', { class: 'add-members' }, h('summary', { text: state.me.hasPasscode ? 'Change passcode' : 'Set a passcode' }));
  const [l0, cur] = passField('cur-pass', 'Current passcode');
  cur.required = false;
  const rule = passRule();
  const [l1, p1] = passField('chg-pass', `New passcode (${rule.label})`);
  const [l2, p2] = passField('chg-pass-2', 'Enter it again');
  const err = h('p', { class: 'error-dark', role: 'alert' });
  const btn = h('button', { class: 'btn', type: 'submit', text: 'Save passcode' });
  wrap.append(h('form', {
    class: 'field-block',
    onsubmit: async e => {
      e.preventDefault();
      err.textContent = '';
      if (!rule.re.test(p1.value)) { err.textContent = `Use ${rule.label}.`; return; }
      if (p1.value !== p2.value) { err.textContent = 'The two passcodes do not match.'; return; }
      btn.disabled = true;
      try {
        await api.setPasscode(p1.value, cur.value);
        state.me.hasPasscode = true;
        [cur, p1, p2].forEach(i => { i.value = ''; });
        wrap.open = false;
        toast('Passcode changed');
      } catch (ex) {
        err.textContent = errorText(ex);
      }
      btn.disabled = false;
    }
  },
    state.me.hasPasscode ? h('div', { class: 'field-block' }, l0, cur,
      h('p', { class: 'hint', text: 'Forgot it? Ask your administrator to reset it.' })) : null,
    l1, p1, l2, p2, err, btn));
  return wrap;
}

/** Remove disappearing messages from the screen the moment they expire. */
function expiryCheck() {
  if (!state.me || state.locked) return;
  let listChanged = false;
  state.convs.forEach(c => {
    if (c.lastMsgPreview && c.lastMsgExpiresAt && isExpired({ expiresAt: c.lastMsgExpiresAt })) { c.lastMsgPreview = ''; listChanged = true; }
  });
  if (listChanged) renderList();
  const map = state.activeId && state.msgs[state.activeId];
  if (!map) return;
  let gone = false;
  map.forEach((m, id) => { if (isExpired(m)) { map.delete(id); gone = true; } });
  if (gone) renderThread(false);
}

function idleCheck() {
  if (!state.me || state.locked || !$('auth').hidden) return;
  const mins = Number(state.security.lockAfterMinutes) || 0;
  if (mins > 0 && Date.now() - guard.lastActivity >= mins * 60000) lockApp({});
}

/**
 * Privacy shield. Phones keep a picture of the screen for the recent-apps view; the shield
 * covers the chats the moment the app starts to leave the foreground, so that picture shows
 * only the ZeeMo logo. It is lifted when the app comes back, behind the lock screen if locked.
 */
function shield(on) {
  document.documentElement.classList.toggle('shielded', !!on);
}
function wireShield() {
  const phone = !(window.matchMedia && window.matchMedia('(hover: hover) and (pointer: fine)').matches);
  const back = () => { if (!document.hidden) shield(false); };
  // On phones, focus is lost first when the recent-apps view or home screen opens: cover at once.
  window.addEventListener('blur', () => { if (phone && state.me) shield(true); });
  window.addEventListener('focus', back);
  document.addEventListener('visibilitychange', () => { if (document.hidden) { if (state.me) shield(true); } else back(); }, true);
  window.addEventListener('pagehide', () => shield(true), true);
  document.addEventListener('freeze', () => shield(true));
  window.addEventListener('pageshow', back);
  // Never trap anyone behind the shield if a browser misses the focus event.
  $('shield').addEventListener('pointerdown', () => { if (!document.hidden) shield(false); });
}

function wireSecurity() {
  wireShield();
  ['pointerdown', 'keydown', 'touchstart', 'wheel', 'input'].forEach(ev =>
    document.addEventListener(ev, markActive, { capture: true, passive: true }));
  setInterval(idleCheck, 10000);
  setInterval(expiryCheck, 15000);

  document.addEventListener('visibilitychange', () => {
    if (!state.me || !$('auth').hidden) return;
    if (document.hidden) {
      if (guard.pickerOpen) return;
      const action = state.security.minimiseAction;
      if (action === 'lock') lockApp({});
      else if (action === 'close' && state.activeId) closeChat();
    } else {
      // Timers are slowed while hidden, so check the idle time as soon as the page is back.
      idleCheck();
    }
  });

  // Optional, computers only: lock as soon as the window loses focus.
  const finePointer = window.matchMedia && window.matchMedia('(hover: hover) and (pointer: fine)').matches;
  window.addEventListener('blur', () => {
    if (!finePointer || !state.security.lockOnBlur || !state.me || state.locked || !$('auth').hidden || guard.pickerOpen) return;
    setTimeout(() => { if (!document.hasFocus() && !guard.pickerOpen) lockApp({}); }, 800);
  });

  // Never leave chats in the back/forward cache; returning starts again from the lock screen.
  window.addEventListener('pagehide', () => {
    if (!state.me) return;
    state.locked = false;
    wipeScreen();
  });
  window.addEventListener('pageshow', e => { if (e.persisted) start(); });
}

/* ---------- Menu ---------- */
function toggleMenu() {
  const m = $('menu');
  $('menu-lock').hidden = !(state.me && state.me.hasPasscode);
  m.hidden = !m.hidden;
  if (!m.hidden) m.querySelector('button').focus();
}
function closeMenu() { $('menu').hidden = true; }

/* ---------- Client events ---------- */
api.addEventListener('pending', e => {
  const p = e.detail;
  const map = state.pending[p.convId] || (state.pending[p.convId] = new Map());
  map.set(p.clientId, { ...p, senderId: state.me.userId, msgId: '', status: 'pending' });
  if (p.convId === state.activeId) renderThread(true);
});
api.addEventListener('sent', e => {
  const m = e.detail.message;
  mergeMessages(m.convId, [m]);
  const conv = state.convs.find(c => c.convId === m.convId);
  if (conv) Object.assign(conv, { lastMsgId: m.msgId, lastMsgAt: m.createdAt, lastMsgSenderId: m.senderId, lastMsgPreview: previewOf(m), lastMsgExpiresAt: m.expiresAt || '', lastMsgType: m.type });
  state.convs.sort((a, b) => (a.lastMsgAt < b.lastMsgAt ? 1 : -1));
  renderList();
  if (m.convId === state.activeId) renderThread(true);
});
api.addEventListener('failed', e => {
  for (const map of Object.values(state.pending)) {
    const p = map.get(e.detail.clientId);
    if (p) p.status = 'failed';
  }
  if (state.activeId) renderThread(false);
});
api.addEventListener('messages', e => {
  const { convId, messages } = e.detail;
  mergeMessages(convId, messages);
  const conv = state.convs.find(c => c.convId === convId);
  const last = messages[messages.length - 1];
  if (conv && last && last.msgId >= (conv.lastMsgId || '')) {
    Object.assign(conv, { lastMsgId: last.msgId, lastMsgAt: last.createdAt, lastMsgSenderId: last.senderId, lastMsgPreview: previewOf(last), lastMsgExpiresAt: last.expiresAt || '', lastMsgType: last.type });
    if (convId !== state.activeId) {
      const fresh = messages.filter(m => m.senderId !== state.me.userId && m.type !== 'system').length;
      if (fresh) { conv.hasUnread = true; conv.unreadCount = (Number(conv.unreadCount) || 0) + fresh; }
    }
  }
  state.convs.sort((a, b) => (a.lastMsgAt < b.lastMsgAt ? 1 : -1));
  renderList();
  if (convId === state.activeId) renderThread(false);
});
api.addEventListener('deletions', e => {
  const { convId, msgIds } = e.detail;
  const map = state.msgs[convId];
  let changed = false;
  msgIds.forEach(id => {
    const m = map && map.get(id);
    if (m && !m.deleted) { Object.assign(m, { deleted: true, body: '', mediaFileId: '' }); changed = true; }
  });
  const conv = state.convs.find(c => c.convId === convId);
  if (conv && msgIds.includes(conv.lastMsgId)) { conv.lastMsgPreview = 'This message was deleted'; renderList(); }
  if (changed && convId === state.activeId) renderThread(false);
});
api.addEventListener('edits', e => applyEdits(e.detail.convId, e.detail.messages));
api.addEventListener('receipts', e => {
  state.receipts[e.detail.convId] = e.detail.members;
  if (e.detail.convId === state.activeId) renderThread(false);
});
api.addEventListener('listchanged', loadConversations);
api.addEventListener('synced', () => { $('offline').hidden = true; });
api.addEventListener('syncerror', e => { if (e.detail.code === 'NETWORK') $('offline').hidden = false; });
api.addEventListener('signedout', e => {
  wipeScreen();
  state.me = null;
  state.locked = false;
  renderAuth(e.detail && e.detail.reason);
});
api.addEventListener('locked', () => { if (state.me && !state.locked) lockApp({ serverAlreadyLocked: true }); });
api.addEventListener('needpasscode', () => { if (state.me) { wipeScreen(); renderSetPasscode(); } });
api.addEventListener('needterms', () => { if (state.me) { wipeScreen(); state.me.termsAccepted = false; renderTerms('resume'); } });

/* ---------- Boot ---------- */
/** Show the chats. Only called once the user is signed in and unlocked. */
async function boot() {
  if (!state.me) state.me = await api.call('me');
  state.locked = false;
  $('auth').hidden = true;
  markActive();
  await loadConversations();
  if (!state.activeId) renderNoChat();
  api.start();
  checkNewSignIns();
}

/** Opening the app (or returning to it from the back/forward cache). */
async function start() {
  const hadConfig = cachedAuthConfig();
  if (!api.signedIn) {
    if (!hadConfig) { renderAuth(); return; }
    renderAuth();
    // If the owner has changed the sign-in methods since the last visit, redraw once.
    const before = auth.methods.join();
    loadAuthConfig(true).then(() => {
      const root = $('auth');
      if (auth.methods.join() !== before && auth.step === 'start' && !root.hidden && root.dataset.mode === 'signin') {
        const typed = $('auth-id');
        if (typed) auth.value = typed.value;
        renderAuth();
      }
    });
    return;
  }
  try {
    // Settings and profile together; with saved settings, only the profile is waited for.
    const fresh = loadAuthConfig(true);
    const [me] = await Promise.all([api.call('me'), hadConfig ? null : fresh]);
    state.me = me;
  } catch (err) {
    if (err.code === 'NETWORK') { renderAuth('No connection. Check your data or Wi-Fi and try again.'); return; }
    if (!api.signedIn) return;   // the signedout handler has shown the sign-in screen
    renderAuth(errorText(err));
    return;
  }
  nextStep('start');
}

function wireShell() {
  $('menu-btn').addEventListener('click', toggleMenu);
  $('menu-new-group').addEventListener('click', openNewGroup);
  $('menu-profile').addEventListener('click', openProfile);
  $('menu-signout').addEventListener('click', () => { closeMenu(); api.logout().catch(() => {}); });
  $('menu-lock').addEventListener('click', () => { closeMenu(); lockApp({}); });
  $('fab').addEventListener('click', openNewChat);
  $('search-input').addEventListener('input', e => { state.filter = e.target.value; renderList(); });
  document.addEventListener('click', e => {
    const m = $('menu');
    if (!m.hidden && !m.contains(e.target) && e.target.closest('#menu-btn') === null) closeMenu();
  });
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    if (!$('menu').hidden) closeMenu();
    else if (!$('sheet').hidden) closeSheet();
    else if (state.activeId && window.innerWidth < 900) closeChat();
  });
}

wireShell();
wireSecurity();
renderNoChat();
start();

// Offline app shell. Skipped in the single-file preview, where service workers are unavailable.
if ('serviceWorker' in navigator && location.protocol === 'https:' && !window.ZEEMO_PREVIEW && !IN_APPS_SCRIPT) {
  navigator.serviceWorker.register('sw.js').catch(() => { /* offline shell unavailable */ });
}
