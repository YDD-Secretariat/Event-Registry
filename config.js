// ============================================================================
//  Event Registry — shared client (one codebase for every programme)
//  1) Paste your Apps Script Web App URL below (see README, step 3).
//  2) Everything else (name, groups, labels, which modules show) comes from
//     the "Programs" sheet in the database workbook — no code changes needed.
// ============================================================================
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxQOW946XRkT60JSci-J015NajxRT_5EDXjtZvWs6Ps7aotsz3vvokuDAw_oblCFCIK/exec";

const QUEUE_KEY = 'registry_offline_queue_v2';
const STATION_KEY = 'registry_station';
const PROGRAM_KEY = 'registry_program';
const PROGRAM_CACHE = 'registry_program_cache_';

const DEFAULT_PROGRAM = {
  id: '', name: 'Event Registry', tagline: '',
  groups: ['Group 1'], otherGroup: 'Others', visitorsInMembers: false,
  modules: { schools: true, members: true, visitors: true, workersBulk: false, individuals: true, uploads: true },
  labels: { member: 'Members', worker: 'Workers', school: 'School', student: 'Students', group: 'Group', stationHint: 'e.g. Desk 3' },
  accent: '#4f46e5'
};

const Registry = {
  // ---------- state ----------
  program: DEFAULT_PROGRAM,
  isSyncing: false,
  lastSyncError: '',
  _listeners: [],

  // ---------- programme selection ----------
  programId() {
    const fromUrl = new URLSearchParams(location.search).get('program');
    if (fromUrl) { try { localStorage.setItem(PROGRAM_KEY, fromUrl); } catch (_) {} return fromUrl; }
    try { return localStorage.getItem(PROGRAM_KEY) || ''; } catch (_) { return ''; }
  },
  setProgram(id) { try { localStorage.setItem(PROGRAM_KEY, id || ''); } catch (_) {} },
  clearProgram() { try { localStorage.removeItem(PROGRAM_KEY); } catch (_) {} },
  configured() { return APPS_SCRIPT_URL && !/PASTE_YOUR/.test(APPS_SCRIPT_URL); },

  // ---------- station / desk ----------
  getStation() { try { return localStorage.getItem(STATION_KEY) || ''; } catch (_) { return ''; } },
  setStation(name) { try { localStorage.setItem(STATION_KEY, name || ''); } catch (_) {} },

  // ---------- network + offline queue ----------
  isOnline() { return navigator.onLine; },
  getQueue() { try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]'); } catch (_) { return []; } },
  saveQueue(items) { localStorage.setItem(QUEUE_KEY, JSON.stringify(items)); this.notifyStatusChange(); },
  getQueueCount() { return this.getQueue().length; },
  onStatusChange(fn) { this._listeners.push(fn); fn(this.status()); },
  status() { return { online: this.isOnline(), queueCount: this.getQueueCount(), syncing: this.isSyncing, error: this.lastSyncError }; },
  notifyStatusChange() { this._listeners.forEach(fn => { try { fn(this.status()); } catch (_) {} }); },

  async get(action, params = {}) {
    if (!this.configured()) throw new Error('App is not connected to a database yet (set APPS_SCRIPT_URL in config.js).');
    const url = new URL(APPS_SCRIPT_URL);
    url.searchParams.set('action', action);
    const withProgram = Object.assign({ programId: this.programId() }, params);
    Object.entries(withProgram).forEach(([k, v]) => { if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, v); });
    const response = await fetch(url.toString());
    if (!response.ok) throw new Error('Server HTTP ' + response.status);
    const data = await response.json();
    if (data.error) throw new Error(data.error);
    return data;
  },

  async send(body) {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    });
    if (!response.ok) throw new Error('Server HTTP ' + response.status);
    return response.json();
  },

  // Writes are confirmed only when the server says so. Otherwise they are queued on this
  // device WITH their programme, so a late sync can never land in the wrong event.
  async post(body) {
    const payload = Object.assign({}, body, {
      programId: body.programId || this.programId(),
      clientRequestId: body.clientRequestId || 'req_' + Date.now() + '_' + Math.random().toString(36).slice(2)
    });
    if (!this.configured()) return { success: false, error: 'App is not connected to a database yet.' };
    try {
      const result = await this.send(payload);
      if (!result || result.success !== true) return result || { success: false, error: 'Server did not confirm the write.' };
      return result;
    } catch (err) {
      const queue = this.getQueue();
      queue.push({ id: payload.clientRequestId, queuedAt: new Date().toISOString(), programId: payload.programId, payload, attempts: 0 });
      this.saveQueue(queue);
      return { success: false, queued: true, error: 'Write not confirmed. Saved on this device for retry. Check the sync badge.' };
    }
  },

  async processNextQueueItem() {
    if (this.isSyncing || !this.isOnline() || !this.configured()) return;
    const queue = this.getQueue();
    if (!queue.length) return;
    this.isSyncing = true; this.notifyStatusChange();
    const item = queue[0];
    try {
      const result = await this.send(Object.assign({}, item.payload, { clientRequestId: item.payload.clientRequestId || item.id, fromQueue: true }));
      if (!result || (result.success !== true && !result.duplicate)) {
        this.lastSyncError = (result && result.error) || 'Server rejected a pending write.';
        return;
      }
      queue.shift(); this.saveQueue(queue);
      this.lastSyncError = '';
    } catch (err) {
      this.lastSyncError = 'Pending write could not reach server.';
    } finally {
      this.isSyncing = false; this.notifyStatusChange();
    }
  },
  syncQueue() { return this.processNextQueueItem(); },

  // ---------- programme config ----------
  async loadProgram() {
    const id = this.programId();
    let cfg = null;
    try {
      const res = await this.get('getProgram', { programId: id });
      cfg = res.program;
      try { localStorage.setItem(PROGRAM_CACHE + id, JSON.stringify(cfg)); } catch (_) {}
    } catch (e) {
      try { cfg = JSON.parse(localStorage.getItem(PROGRAM_CACHE + id) || 'null'); } catch (_) {}
      if (!cfg) throw e;
    }
    this.program = {
      ...DEFAULT_PROGRAM, ...cfg,
      modules: { ...DEFAULT_PROGRAM.modules, ...(cfg.modules || {}) },
      labels: { ...DEFAULT_PROGRAM.labels, ...(cfg.labels || {}) }
    };
    return this.program;
  },

  get labels() {
    const l = this.program.labels;
    return { ...l, schools: /s$/i.test(l.school) ? l.school : l.school + 's' };
  },

  // Fill <span data-l="member"> etc. and program name/tagline placeholders.
  fill(root = document) {
    const L = this.labels;
    root.querySelectorAll('[data-l]').forEach(el => { const v = L[el.dataset.l]; if (v) el.textContent = v; });
    root.querySelectorAll('[data-program-name]').forEach(el => el.textContent = this.program.name);
    root.querySelectorAll('[data-program-tagline]').forEach(el => el.textContent = this.program.tagline || '');
    root.querySelectorAll('[data-placeholder="station"]').forEach(el => el.placeholder = L.stationHint);
    root.querySelectorAll('[data-mod]').forEach(el => {
      const mods = el.dataset.mod.split(',');
      el.style.display = mods.some(m => this.moduleOn(m)) ? '' : 'none';
    });
  },

  moduleOn(key) {
    const m = this.program.modules;
    if (key === 'uploads') return !!(m.uploads && m.schools);
    return !!m[key];
  },

  // ---------- navigation ----------
  navItems() {
    const L = this.labels, p = this.program;
    return [
      ['schools', 'register.html', L.schools],
      ['members', 'register-members.html', `${L.member} (Bulk)`],
      ['workersBulk', 'register-workers.html', `${L.worker} (Bulk)`],
      ['individuals', 'register-worker.html', p.modules.workersBulk ? `${L.worker} (Individual)` : L.worker],
      ['visitors', 'register-visitor.html', 'Visitors'],
      ['uploads', 'upload.html', 'Uploads'],
      ['dashboard', 'dashboard.html', 'Dashboard']
    ].filter(([key]) => key === 'dashboard' || this.moduleOn(key));
  },

  renderNav() {
    const nav = document.getElementById('nav');
    if (!nav) return;
    const here = location.pathname.split('/').pop() || 'index.html';
    nav.innerHTML = this.navItems().map(([, href, label]) =>
      `<a href="${href}" ${href === here ? 'class="active"' : ''}>${label}</a>`).join('');
  },

  injectChrome() {
    const header = document.querySelector('header');
    if (header && !document.getElementById('registry-prog')) {
      const pill = document.createElement('a');
      pill.id = 'registry-prog'; pill.href = 'programs.html';
      pill.title = 'Switch programme';
      pill.style.cssText = 'display:inline-block;margin:6px 0 0;font-size:.72rem;font-weight:700;padding:3px 10px;border-radius:20px;background:rgba(255,255,255,.18);color:inherit;text-decoration:none;';
      pill.textContent = '📍 ' + this.program.name + ' · switch';
      const h1 = header.querySelector('h1');
      (h1 || header).insertAdjacentElement('afterend', pill);
    }
    if (header && !document.getElementById('registry-sync-badge')) {
      const badge = document.createElement('button');
      badge.id = 'registry-sync-badge'; badge.type = 'button';
      badge.style.cssText = 'border:0;border-radius:20px;padding:5px 12px;background:#334155;color:white;cursor:pointer;font-size:.75rem;';
      this.onStatusChange(st => {
        badge.textContent = st.error ? `⚠ Sync error · ${st.queueCount} pending` :
          st.syncing ? `Syncing · ${st.queueCount} pending` :
          st.queueCount ? `${st.queueCount} pending — tap to retry` : '✓ Server confirmed';
        badge.title = st.error || 'Entries are confirmed only after the server saves them.';
      });
      badge.onclick = () => this.syncQueue();
      header.appendChild(badge);
    }
    if (!this.configured() && !document.getElementById('registry-setup-warning')) {
      const w = document.createElement('div');
      w.id = 'registry-setup-warning';
      w.style.cssText = 'background:#fef3c7;color:#92400e;padding:10px 14px;font:600 .85rem sans-serif;text-align:center;';
      w.textContent = 'Not connected: open config.js and paste your Apps Script Web App URL.';
      document.body.insertAdjacentElement('afterbegin', w);
    }
  },

  // ---------- page boot ----------
  // <body data-module="members"> gates the page; data-page="picker" skips the programme requirement.
  async init() {
    const body = document.body;
    if (body.dataset.page === 'picker') { this.injectChrome(); return null; }
    if (!this.programId()) { location.replace('programs.html'); return new Promise(() => {}); }
    try {
      await this.loadProgram();
    } catch (e) {
      this.clearProgram();
      location.replace('programs.html?error=' + encodeURIComponent(e.message || 'Programme not found'));
      return new Promise(() => {});
    }
    const mod = body.dataset.module;
    if (mod && !this.moduleOn(mod)) { location.replace('index.html'); return new Promise(() => {}); }
    document.title = document.title.replace(/\s*[—·]\s*.*$/, '') + ' — ' + this.program.name;
    this.injectChrome(); this.renderNav(); this.fill();
    return this.program;
  },

  esc(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
};

Registry.ready = new Promise(resolve => {
  const go = () => resolve(Registry.init());
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', go); else go();
});

if (typeof window !== 'undefined') {
  window.addEventListener('online', () => Registry.syncQueue());
  setTimeout(() => Registry.syncQueue(), 1500);
  setInterval(() => Registry.syncQueue(), 15000);
}
