/* UI wiring: file loading, Brex column mapping, results table, upload tracking, saved sessions, downloads. */
(function () {
  const M = window.GrabMatch, R = window.GrabReceipt;
  const $ = id => document.getElementById(id);
  const state = { grabFiles: [], grabs: [], brexRows: [], brexHeaders: [], brexName: '', map: {}, card: '1521', matches: [], rows: [], filter: 'All',
                  hideUploaded: false, sessionId: null, sessionName: '' };

  const FIELD_LABELS = { id: 'Expense ID', time: 'Swipe time (UTC)', date: 'Transaction date', amount: 'Purchased amount (original currency)',
    currency: 'Purchased currency', usd: 'USD amount', descriptor: 'Statement descriptor (has Grab booking ID)', merchant: 'Merchant name',
    memo: 'Memo', user: 'Location (employee)', erp: 'Integration ERP link (QuickBooks)', link: 'Brex dashboard link', receipt: 'Receipt URL', status: 'Expense status', last4: 'Card last 4', gl: 'GL account (Integration Debit Line Category)', cls: 'Integration class' };
  const LINK_COLS = new Set(['Integration ERP', 'Brex Dashboard']);
  const COPY_COLS = new Set(['Memo', 'Location', 'Class', 'GL Account']); // fields typed into Brex

  const ICON_COPY = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><rect x="5.25" y="5.25" width="8.5" height="8.5" rx="1.75" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10.75 2.75v-.25a1.25 1.25 0 0 0-1.25-1.25H3.5A2.25 2.25 0 0 0 1.25 3.5v6c0 .7.55 1.25 1.25 1.25h.25" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
  const ICON_CHECK = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path d="M3 8.5l3.25 3.25L13 5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  async function readWorkbook(file) {
    const buf = await file.arrayBuffer();
    // CSV: decode as UTF-8 ourselves (memos contain "—"; a file without a BOM would otherwise be read as Latin-1)
    // and keep cells as text (raw) so dates/IDs aren't reinterpreted
    if (/\.csv$/i.test(file.name)) return XLSX.read(new TextDecoder('utf-8').decode(buf), { type: 'string', raw: true });
    return XLSX.read(buf, { type: 'array' });
  }

  /* ---------- Grab files (accumulate; same name replaces) ---------- */
  function grabEntry(name, matrix) {
    try {
      const trips = M.parseGrabMatrix(matrix, name);
      return { name, matrix, trips, category: trips[0]?.category || '?' };
    } catch (e) {
      return { name, matrix, trips: [], error: e.message };
    }
  }

  async function loadGrab(files) {
    for (const f of files) {
      const wb = await readWorkbook(f);
      const matrix = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
      const entry = grabEntry(f.name, matrix);
      state.grabFiles = state.grabFiles.filter(x => x.name !== f.name).concat(entry);
    }
    refreshGrab();
  }

  function refreshGrab() {
    state.grabs = state.grabFiles.flatMap(f => f.trips);
    state.grabs.forEach((g, i) => (g.idx = i));
    $('grab-files').innerHTML = state.grabFiles.map((f, i) => `<li class="${f.error ? 'err' : ''}">
        <span class="cat">${esc(f.error ? 'Error' : f.category)}</span>
        <span class="fname" title="${esc(f.name)}">${esc(f.name)}</span>
        <span class="muted">${f.error ? esc(f.error) : `${f.trips.length} rows`}</span>
        <button class="x" data-i="${i}" title="Remove">×</button></li>`).join('');
    $('grab-files').querySelectorAll('.x').forEach(b => (b.onclick = () => { state.grabFiles.splice(+b.dataset.i, 1); refreshGrab(); }));
    const byCat = state.grabs.reduce((c, g) => ((c[g.category] = (c[g.category] || 0) + 1), c), {});
    const n = state.grabFiles.filter(f => !f.error).length;
    if (!n) { $('status-grab').textContent = 'No file yet'; $('drop-grab').classList.remove('done'); $('run').disabled = true; return; }
    setStatus('grab', Object.entries(byCat).map(([c, k]) => `${k} ${c}`).join(' · ') + ` from ${n} file${n > 1 ? 's' : ''}. Drop more to add.`);
  }

  /* ---------- Brex file ---------- */
  function setBrex(name, rows, map) {
    state.brexName = name;
    state.brexRows = rows;
    state.brexHeaders = rows.length ? Object.keys(rows[0]) : [];
    state.map = map || M.autoMapBrex(state.brexHeaders);
    renderMapping();
    setStatus('brex', `${rows.length} transactions`);
  }

  async function loadBrex(file) {
    const wb = await readWorkbook(file);
    setBrex(file.name, XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { raw: true, defval: '' }));
  }

  function setStatus(which, text, err) {
    $('status-' + which).textContent = text;
    $('status-' + which).style.color = err ? 'var(--bad)' : '';
    $('drop-' + which).classList.toggle('done', !err);
    $('run').disabled = !state.grabs.length;
  }

  function renderMapping() {
    const grid = $('map-grid'); grid.innerHTML = '';
    for (const [field, label] of Object.entries(FIELD_LABELS)) {
      const opts = ['<option value="">(none)</option>', ...state.brexHeaders.map(h =>
        `<option ${state.map[field] === h ? 'selected' : ''}>${esc(h)}</option>`)].join('');
      const el = document.createElement('label');
      el.innerHTML = `${label}<select data-field="${field}">${opts}</select>`;
      el.querySelector('select').onchange = e => { state.map[field] = e.target.value || undefined; };
      grid.appendChild(el);
    }
    const card = document.createElement('label');
    card.innerHTML = `Only match card ending in<input id="card" inputmode="numeric" maxlength="4" value="${esc(state.card)}">
      <span class="hint">Brex charges on other cards are skipped. Leave it empty to match all cards.</span>`;
    card.querySelector('input').oninput = e => { state.card = e.target.value.trim(); };
    grid.appendChild(card);
    $('mapping').hidden = false;
    if (!state.map.amount) $('mapping').open = true;
  }

  /* ---------- Matching ---------- */
  function run(opts = {}) {
    if (!opts.keepSession) { state.sessionId = null; state.sessionName = ''; }
    const brex = state.brexRows.length ? M.normalizeBrex(state.brexRows, state.map, $('link-tpl').value.trim()) : [];
    const card = brex.length && state.map.last4 ? state.card : '';
    state.matches = M.matchAll(state.grabs, brex, { dateWindowDays: +$('win').value || 2, card });
    const skipped = card ? brex.filter(b => b.last4 && b.last4 !== card.replace(/\D/g, '')).length : 0;
    state.rows = M.toRows(state.matches, R.receiptFileName);
    state.filter = 'All';
    $('results').hidden = false;
    $('msg').textContent = !brex.length ? 'No Brex file loaded, so only receipts and the masterlist were generated.'
      : skipped ? `Matched on card ending ${card} only. Skipped ${skipped} Brex charges on other cards.` : '';
    const memoIssues = state.rows.filter(r => r['Name Check'] && r['Name Check'] !== 'OK').length;
    if (memoIssues) $('msg').textContent += ` ${memoIssues} matched rows have a Brex memo naming a different person (see Name Check).`;
    renderResults();
    renderSaveState();
    if (!opts.quiet) $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // matched charges that already exist in QuickBooks (they have an ERP link): memo, GL, class and receipt are added there by hand or in a Claude session
  function isQbTodo(r) { return !!(r._b && r['Integration ERP'] && ['Matched', 'Split charge', 'Combined charge'].includes(r.Status)); }

  const statusClass = s => ['Matched', 'Receipt only'].includes(s) ? 's-ok'
    : ['Split charge', 'Combined charge'].includes(s) ? 's-split'
    : ['Unmatched Brex', 'Unmatched Grab'].includes(s) ? 's-bad'
    : s.startsWith('Review') ? 's-warn' : 's-all';


  /* ---------- QuickBooks to-do table: for adding memo, GL, class and receipt in QuickBooks by hand (open the ERP link) ---------- */
  const QB_KEY = 'grab-receipts-qb-done';
  state.qbDone = (() => { try { return JSON.parse(localStorage.getItem(QB_KEY) || '{}'); } catch (e) { return {}; } })();
  const saveQbDone = () => { try { localStorage.setItem(QB_KEY, JSON.stringify(state.qbDone)); } catch (e) { /* storage blocked */ } };
  const qbNorm = v => String(v || '').replace(/^\[|\]$/g, '').trim().toLowerCase();

  const usdFmt = v => (v === '' || v === null || v === undefined || isNaN(+v)) ? '<span class="muted">n/a</span>' : '$' + Math.abs(+v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + (+v < 0 ? ' <span class="muted">(credit)</span>' : '');
  function renderQbTable() {
    const rows = state.rows.filter(isQbTodo);
    const done = rows.filter(r => state.qbDone[qbTxnId(r['Integration ERP'])]).length;
    const copy = v => v ? `<span class="cv">${esc(v)}<button class="copy" data-v="${esc(v)}" title="Copy" aria-label="Copy">${ICON_COPY}</button></span>` : '';
    const body = rows.map(r => {
      const id = qbTxnId(r['Integration ERP']), isDone = !!state.qbDone[id];
      const glNow = r['Current Brex GL'], clsNow = r._b.cls || '';
      const glFix = qbNorm(r['GL Account']) !== qbNorm(glNow), clsFix = qbNorm(r.Class) !== qbNorm(clsNow);
      const m = state.matches[r._m];
      return `<tr class="${isDone ? 'is-up' : ''}">
        <td class="up"><input type="checkbox" class="qbchk" data-k="${esc(id)}" ${isDone ? 'checked' : ''} title="Mark as done in QuickBooks"></td>
        <td><a href="${esc(r['Integration ERP'])}" target="_blank" rel="noopener">open in QuickBooks</a><br><span class="muted">#${esc(id)}</span></td>
        <td>${esc(r.Date)}</td>
        <td class="num">${esc(typeof r['Purchased Amount'] === 'number' ? r['Purchased Amount'].toLocaleString('en-US') : r['Purchased Amount'])} ${esc(r['Purchased Currency'])}</td>
        <td class="num">${usdFmt(r['USD Amount'])}</td>
        <td>${copy(r.Memo)}</td>
        <td>${copy(r['GL Account'])}${glFix ? `<br><span class="pill s-warn">now: ${esc(glNow || 'empty')}</span>` : '<br><span class="muted">already right</span>'}</td>
        <td>${copy(r.Class)}${clsFix ? `<br><span class="pill s-warn">now: ${esc(clsNow || 'empty')}</span>` : '<br><span class="muted">already right</span>'}</td>
        <td>${m.grabs.map((g, j) => `<button class="link-btn" data-m="${r._m}" data-g="${j}">PDF</button>`).join(' ')}</td>
        <td><a href="${esc(r['Brex Dashboard'])}" target="_blank" rel="noopener">Brex</a>${isUp(r) ? ' <span class="muted">✓ pushed</span>' : ''}</td>
      </tr>`;
    }).join('');
    $('table').innerHTML = `<caption style="text-align:left;padding:8px 12px"><strong>${done}</strong> of ${rows.length} done in QuickBooks. Open each link, paste the memo, set the GL and class if "now" differs, and attach the PDF. Tick the box when finished (saved in this browser).</caption>
      <thead><tr><th class="up">Done<br><input type="checkbox" id="qb-all" title="Select / clear all"></th><th>QuickBooks</th><th>Date</th><th>Amount</th><th>USD (as in QuickBooks)</th><th>Memo</th><th>GL account</th><th>Class</th><th>Receipt</th><th>Brex</th></tr></thead><tbody>${body}</tbody>`;
    $('table').querySelectorAll('.copy').forEach(btn => (btn.onclick = async () => {
      const ok = await copyText(btn.dataset.v);
      btn.innerHTML = ok ? ICON_CHECK : '!'; btn.classList.add(ok ? 'copied' : 'failed');
      setTimeout(() => { btn.innerHTML = ICON_COPY; btn.classList.remove('copied', 'failed'); }, 1200);
    }));
    const all = $('qb-all');
    const syncAll = () => {
      const n = rows.filter(r => state.qbDone[qbTxnId(r['Integration ERP'])]).length;
      $('table').querySelector('caption strong').textContent = n;
      all.checked = n === rows.length && n > 0; all.indeterminate = n > 0 && n < rows.length;
    };
    $('table').querySelectorAll('.qbchk').forEach(cb => (cb.onchange = () => {
      if (cb.checked) state.qbDone[cb.dataset.k] = new Date().toISOString(); else delete state.qbDone[cb.dataset.k];
      saveQbDone();
      cb.closest('tr').classList.toggle('is-up', cb.checked);
      syncAll();
    }));
    all.onchange = () => {
      const now = new Date().toISOString();
      $('table').querySelectorAll('.qbchk').forEach(cb => {
        cb.checked = all.checked;
        if (all.checked) state.qbDone[cb.dataset.k] = state.qbDone[cb.dataset.k] || now; else delete state.qbDone[cb.dataset.k];
        cb.closest('tr').classList.toggle('is-up', all.checked);
      });
      saveQbDone(); syncAll();
    };
    syncAll();
    $('table').querySelectorAll('.link-btn').forEach(b => (b.onclick = () => {
      const g = state.matches[+b.dataset.m].grabs[+b.dataset.g];
      R.buildReceipt(g).save(R.receiptFileName(g));
    }));
  }

  /* ---------- Upload tracking: a tick per Brex charge, remembered by expense ID across sessions ---------- */
  const UP_KEY = 'grab-receipts-uploaded';
  state.uploaded = (() => { try { return JSON.parse(localStorage.getItem(UP_KEY) || '{}'); } catch (e) { return {}; } })();
  const saveUploaded = () => { try { localStorage.setItem(UP_KEY, JSON.stringify(state.uploaded)); } catch (e) { /* storage blocked */ } };
  const rowKey = r => (r._b ? r._b.id || r._b.link : '');
  const isUp = r => !!(rowKey(r) && state.uploaded[rowKey(r)]);
  const trackable = () => state.rows.filter(r => rowKey(r));

  function renderProgress() {
    const all = trackable(), done = all.filter(isUp).length;
    $('up-progress').innerHTML = all.length
      ? `<strong>${done}</strong>&nbsp;of ${all.length} uploaded<span class="bar"><span style="width:${(done / all.length) * 100}%"></span></span>` : '';
  }

  function toggleUploaded(key, on) {
    if (on) state.uploaded[key] = new Date().toISOString(); else delete state.uploaded[key];
    saveUploaded();
    renderProgress();
    if (state.sessionId) autosave();
  }

  /* ---------- Results table ---------- */
  // Brex Dashboard link comes first; the QuickBooks (ERP) column only exists if at least one row has an ERP link
  const visibleCols = rows => {
    let cols = Object.keys(rows[0] || {});
    if (!rows.some(r => r['Integration ERP'])) cols = cols.filter(c => c !== 'Integration ERP');
    const i = cols.indexOf('Brex Dashboard'), j = cols.indexOf('Integration ERP');
    if (i > -1 && j > -1 && i > j) { cols.splice(i, 1); cols.splice(j, 0, 'Brex Dashboard'); }
    return cols;
  };
  function renderResults() {
    const counts = state.matches.reduce((c, m) => ((c[m.type] = (c[m.type] || 0) + 1), c), {}); // per match, not per row
    const qbTodo = state.rows.filter(isQbTodo).length;
    $('chips').innerHTML = [['All', state.matches.length], ...Object.entries(counts), ...(qbTodo ? [['QuickBooks to-do', qbTodo]] : [])].map(([k, v]) =>
      `<span class="chip ${k === 'All' ? 's-all' : k === 'QuickBooks to-do' ? 's-split' : statusClass(k)} ${state.filter === k ? 'active' : ''}" data-k="${esc(k)}">${esc(k)}: ${v}</span>`).join('');
    $('chips').querySelectorAll('.chip').forEach(c => (c.onclick = () => { state.filter = c.dataset.k; renderResults(); }));
    renderProgress();
    if (state.filter === 'QuickBooks to-do') return renderQbTable();

    const cols = visibleCols(state.rows);
    const linkCell = v => v ? v.split(' | ').map(u => `<a href="${esc(u)}" target="_blank" rel="noopener">open</a>`).join(' · ') : '';
    const rows = state.rows
      .filter(r => state.filter === 'All' || (state.filter === 'QuickBooks to-do' ? isQbTodo(r) : r.Status === state.filter))
      .filter(r => !(state.hideUploaded && isUp(r)));
    const body = rows.map(r => {
      const i = r._m, key = rowKey(r), up = isUp(r);
      const tick = key
        ? `<td class="up"><input type="checkbox" class="upchk" data-k="${esc(key)}" ${up ? 'checked' : ''} title="Mark as uploaded to Brex"></td>`
        : '<td class="up"><span class="muted" title="No Brex charge to upload">–</span></td>';
      return `<tr class="${up ? 'is-up' : ''}">` + tick + cols.map(c => {
        const v = r[c];
        if (c === 'Status') return `<td><span class="pill ${statusClass(v)}">${esc(v)}</span></td>`;
        if (LINK_COLS.has(c)) return `<td>${linkCell(v)}</td>`;
        if (COPY_COLS.has(c) && v) return `<td><span class="cv">${esc(v)}<button class="copy" data-v="${esc(v)}" title="Copy ${esc(c)}" aria-label="Copy ${esc(c)}">${ICON_COPY}</button></span></td>`;
        if (c === 'Receipt PDF') return `<td>${state.matches[i].grabs.map((g, j) => `<button class="link-btn" data-m="${i}" data-g="${j}">PDF</button>`).join(' ')}</td>`;
        return `<td class="${typeof v === 'number' ? 'num' : ''}">${esc(typeof v === 'number' ? v.toLocaleString('en-US') : v)}</td>`;
      }).join('') + '</tr>';
    }).join('');
    $('table').innerHTML = `<thead><tr><th class="up" title="Uploaded to Brex">Uploaded</th>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${body}</tbody>`;

    $('table').querySelectorAll('.copy').forEach(btn => (btn.onclick = async () => {
      const ok = await copyText(btn.dataset.v);
      btn.innerHTML = ok ? ICON_CHECK : '!'; btn.classList.add(ok ? 'copied' : 'failed');
      setTimeout(() => { btn.innerHTML = ICON_COPY; btn.classList.remove('copied', 'failed'); }, 1200);
    }));
    $('table').querySelectorAll('.upchk').forEach(cb => (cb.onchange = () => {
      toggleUploaded(cb.dataset.k, cb.checked);
      // a split charge has one row per Brex line, so only this row changes
      cb.closest('tr').classList.toggle('is-up', cb.checked);
      if (state.hideUploaded && cb.checked) setTimeout(renderResults, 250);
    }));
    $('table').querySelectorAll('.link-btn').forEach(b => (b.onclick = () => {
      const g = state.matches[+b.dataset.m].grabs[+b.dataset.g];
      R.buildReceipt(g).save(R.receiptFileName(g));
    }));
  }

  // clipboard API first (with a timeout: it can wait forever on a permission prompt), then the legacy fallback
  async function copyText(text) {
    try {
      await Promise.race([navigator.clipboard.writeText(text), new Promise((_, no) => setTimeout(() => no(new Error('timeout')), 1000))]);
      return true;
    } catch (e) {
      const t = Object.assign(document.createElement('textarea'), { value: text });
      t.style.position = 'fixed'; t.style.opacity = '0';
      document.body.appendChild(t); t.select();
      let ok = false;
      try { ok = document.execCommand('copy'); } catch (e2) { /* not allowed */ }
      t.remove(); return ok;
    }
  }

  /* ---------- Saved sessions (IndexedDB in this browser; export/import as a .json file) ---------- */
  // IndexedDB (room for many sessions); falls back to localStorage if it is unavailable or doesn't open in time
  // (private windows, some browsers). "?storage=local" forces the fallback.
  const db = (() => {
    const LS = 'grab-receipts-sessions';
    const read = () => { try { return JSON.parse(localStorage.getItem(LS) || '{}'); } catch (e) { return {}; } };
    const write = m => {
      try { localStorage.setItem(LS, JSON.stringify(m)); }
      catch (e) { throw new Error('browser storage is full. Export and delete old sessions'); }
    };
    const local = {
      all: async () => Object.values(read()),
      put: async v => { const m = read(); m[v.id] = v; write(m); },
      del: async id => { const m = read(); delete m[id]; write(m); },
    };
    let backend;
    const pick = () => backend || (backend = new Promise(res => {
      if (!window.indexedDB || new URLSearchParams(location.search).get('storage') === 'local') return res(local);
      const timer = setTimeout(() => res(local), 1500);
      try {
        const r = indexedDB.open('grab-receipts', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('sessions', { keyPath: 'id' });
        r.onsuccess = () => {
          clearTimeout(timer);
          const conn = r.result;
          const tx = (mode, fn) => new Promise((ok, no) => {
            const t = conn.transaction('sessions', mode), req = fn(t.objectStore('sessions'));
            t.oncomplete = () => ok(req && req.result);
            t.onerror = () => no(t.error);
          });
          res({ all: () => tx('readonly', st => st.getAll()), put: v => tx('readwrite', st => st.put(v)), del: id => tx('readwrite', st => st.delete(id)) });
        };
        r.onerror = () => { clearTimeout(timer); res(local); };
      } catch (e) { clearTimeout(timer); res(local); }
    }));
    return { all: async () => (await pick()).all(), put: async v => (await pick()).put(v), del: async id => (await pick()).del(id) };
  })();

  function defaultSessionName() {
    const first = state.grabs.map(g => g.date).filter(Boolean).sort((a, b) => a - b)[0];
    return `${first ? M.fmtDate(first, 8).slice(3) : 'Grab'} · Grab × Brex`;
  }

  function buildSession(id, name) {
    const keys = trackable().map(rowKey);
    const uploaded = Object.fromEntries(keys.filter(k => state.uploaded[k]).map(k => [k, state.uploaded[k]]));
    return {
      id, name, savedAt: new Date().toISOString(), version: 1,
      grabFiles: state.grabFiles.filter(f => !f.error).map(f => ({ name: f.name, matrix: f.matrix })),
      brex: { name: state.brexName, headers: state.brexHeaders, rows: state.brexRows.map(r => state.brexHeaders.map(h => r[h])) },
      settings: { map: state.map, card: state.card, win: $('win').value, tpl: $('link-tpl').value },
      uploaded, stats: { total: keys.length, uploaded: Object.keys(uploaded).length },
    };
  }

  async function saveSession(name) {
    const existing = (await db.all()).find(s => s.name === name);
    const id = state.sessionId || existing?.id || `s-${Date.now()}`;
    await db.put(buildSession(id, name));
    state.sessionId = id; state.sessionName = name;
    renderSaveState(`Saved ✓ ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`);
    renderSessions();
  }

  let autosaveTimer;
  function autosave() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => saveSession(state.sessionName).catch(e => renderSaveState('Could not save: ' + e.message)), 500);
  }

  function renderSaveState(text) {
    $('save-session').innerHTML = `${ICON_SAVE} ${state.sessionId ? 'Save' : 'Save session'}`;
    $('save-state').textContent = text || (state.sessionId ? `Session: ${state.sessionName} · ticks save automatically` : '');
  }
  const ICON_SAVE = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path d="M3 1.75h8.4l2.85 2.85v8.4c0 .7-.55 1.25-1.25 1.25H3c-.7 0-1.25-.55-1.25-1.25V3c0-.7.55-1.25 1.25-1.25z M5 1.75v3.5h5v-3.5 M4.5 14.25v-4.5h7v4.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>';

  async function renderSessions() {
    let list = [];
    try { list = (await db.all()).sort((a, b) => b.savedAt.localeCompare(a.savedAt)); } catch (e) { /* IndexedDB unavailable */ }
    $('session-list').innerHTML = list.length ? list.map(s => {
      const when = new Date(s.savedAt).toLocaleString([], { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
      const pct = s.stats.total ? Math.round((s.stats.uploaded / s.stats.total) * 100) : 0;
      return `<li class="${s.id === state.sessionId ? 'current' : ''}">
        <div class="s-main"><strong>${esc(s.name)}</strong>${s.id === state.sessionId ? ' <span class="pill s-ok">open</span>' : ''}
          <span class="muted">Saved ${esc(when)} · ${s.grabFiles.length} Grab file${s.grabFiles.length > 1 ? 's' : ''} + Brex</span></div>
        <div class="s-prog"><span class="bar"><span style="width:${pct}%"></span></span><span class="muted">${s.stats.uploaded}/${s.stats.total} uploaded</span></div>
        <div class="s-actions">
          <button class="secondary sm" data-open="${esc(s.id)}">Open</button>
          <button class="secondary sm" data-export="${esc(s.id)}" title="Download as a file to back up or share">Export</button>
          <button class="x" data-del="${esc(s.id)}" title="Delete">×</button>
        </div></li>`;
    }).join('') : '<li class="empty muted">No saved sessions yet. After matching, click “Save session”.</li>';
    const byId = id => list.find(s => s.id === id);
    $('session-list').querySelectorAll('[data-open]').forEach(b => (b.onclick = () => openSession(byId(b.dataset.open))));
    $('session-list').querySelectorAll('[data-export]').forEach(b => (b.onclick = () => exportSession(byId(b.dataset.export))));
    $('session-list').querySelectorAll('[data-del]').forEach(b => (b.onclick = async () => {
      const s = byId(b.dataset.del);
      if (!confirm(`Delete the saved session “${s.name}”? Your upload ticks stay remembered.`)) return;
      await db.del(s.id);
      if (state.sessionId === s.id) { state.sessionId = null; state.sessionName = ''; renderSaveState(); }
      renderSessions();
    }));
  }

  function openSession(s) {
    state.grabFiles = s.grabFiles.map(f => grabEntry(f.name, f.matrix));
    refreshGrab();
    setBrex(s.brex.name, s.brex.rows.map(a => Object.fromEntries(s.brex.headers.map((h, i) => [h, a[i]]))), s.settings.map);
    state.card = s.settings.card ?? '1521';
    renderMapping();
    $('win').value = s.settings.win || 2;
    $('link-tpl').value = s.settings.tpl || $('link-tpl').value;
    // ticks from the session file win only where this browser has none
    for (const [k, v] of Object.entries(s.uploaded || {})) if (!state.uploaded[k]) state.uploaded[k] = v;
    saveUploaded();
    state.sessionId = s.id; state.sessionName = s.name;
    run({ keepSession: true });
    renderSaveState(`Opened “${s.name}”`);
    renderSessions();
  }

  function exportSession(s) {
    const blob = new Blob([JSON.stringify(s)], { type: 'application/json' });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `${s.name.replace(/[\\/:*?"<>|]+/g, '-')}.grab-session.json` });
    a.click(); URL.revokeObjectURL(a.href);
  }

  async function importSession(file) {
    try {
      const s = JSON.parse(await file.text());
      if (!s.grabFiles || !s.brex || !s.name) throw new Error('not a Grab receipts session file');
      s.id = s.id || `s-${Date.now()}`;
      await db.put(s);
      renderSessions();
      openSession(s);
    } catch (e) {
      alert('Could not import: ' + e.message);
    }
  }

  /* ---------- Downloads ---------- */
  function downloadXlsx() {
    // first column records when each Brex charge was ticked as uploaded
    const local = iso => { const d = new Date(iso), p = n => String(n).padStart(2, '0');
      return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; };
    const rows = state.rows.map(r => ({ Uploaded: isUp(r) ? local(state.uploaded[rowKey(r)]) : '', ...r }));
    const cols = ['Uploaded', ...visibleCols(state.rows)];
    const ws = XLSX.utils.json_to_sheet(rows, { header: cols });
    // make single-URL link cells clickable
    cols.forEach((c, ci) => {
      if (!LINK_COLS.has(c)) return;
      rows.forEach((r, ri) => {
        const v = r[c]; if (!v || v.includes(' | ')) return;
        const ref = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
        ws[ref].l = { Target: v };
      });
    });
    ws['!cols'] = cols.map(c => ({ wch: LINK_COLS.has(c) ? 45 : ['Receipt PDF', 'Memo', 'Current Brex Memo', 'Note', 'Name Check'].includes(c) ? 36 : ['GL Account', 'Current Brex GL'].includes(c) ? 22 : 16 }));
    ws['!autofilter'] = { ref: ws['!ref'] };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Matched Receipts');
    XLSX.writeFile(wb, `Matched_Receipts_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // Charges that already exist in QuickBooks (they have an ERP link): a to-do list for adding memo, GL, class and receipt there by hand
  function qbTxnId(link) {
    const m = String(link || '').match(/txnId=(\d+)/i) || String(link || '').match(/(\d{3,})(?!.*\d{3,})/);
    return m ? m[1] : '';
  }
  function downloadQbList() {
    const norm = v => String(v || '').replace(/^\[|\]$/g, '').trim().toLowerCase();
    const rows = state.rows.filter(r => r._b && r['Integration ERP'] && PUSHABLE.has(r.Status)).map(r => {
      const glNow = r['Current Brex GL'], clsNow = r._b.cls || '';
      const glFix = norm(r['GL Account']) !== norm(glNow), clsFix = norm(r.Class) !== norm(clsNow);
      return {
        'QuickBooks ID': qbTxnId(r['Integration ERP']),
        'QuickBooks link': r['Integration ERP'],
        Date: r.Date,
        Amount: r['Purchased Amount'],
        Currency: r['Purchased Currency'],
        'USD Amount': r['USD Amount'],
        Memo: r.Memo,
        'Target GL': r['GL Account'],
        'Current GL': glNow,
        'Fix GL?': glFix ? 'yes' : 'no',
        'Target Class': r.Class,
        'Current Class': clsNow,
        'Fix Class?': clsFix ? 'yes' : 'no',
        'Receipt PDF': r['Receipt PDF'],
        'Brex Dashboard': r['Brex Dashboard'],
        'Pushed to Brex': isUp(r) ? 'yes' : 'no'
      };
    });
    if (!rows.length) return alert('No matched charges with a QuickBooks (ERP) link in these results.');
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'QuickBooks to do');
    XLSX.writeFile(wb, `QuickBooks_todo_${new Date().toISOString().slice(0, 10)}.csv`, { bookType: 'csv' });
  }

  async function downloadZip() {
    const btn = $('dl-zip'); btn.disabled = true;
    const zip = new JSZip(), used = new Set();
    const grabs = state.matches.flatMap(m => m.grabs);
    for (let i = 0; i < grabs.length; i++) {
      btn.textContent = `Building PDFs… ${i + 1}/${grabs.length}`;
      let name = R.receiptFileName(grabs[i]);
      while (used.has(name)) name = name.replace(/\.pdf$/, ' (2).pdf');
      used.add(name);
      zip.file(name, R.buildReceipt(grabs[i]).output('arraybuffer'));
      if (i % 20 === 0) await new Promise(r => setTimeout(r));
    }
    const blob = await zip.generateAsync({ type: 'blob' });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `Grab_Receipts_${new Date().toISOString().slice(0, 10)}.zip` });
    a.click(); URL.revokeObjectURL(a.href);
    btn.textContent = 'Download all receipt PDFs (.zip)'; btn.disabled = false;
  }

  function esc(s) { return String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

  function wireDrop(which, handler) {
    const drop = $('drop-' + which), input = $('file-' + which);
    const go = async files => { files = [...files]; if (!files.length) return; try { await handler(files); } catch (e) { setStatus(which, e.message, true); } };
    input.onchange = () => { go([...input.files]); input.value = ''; }; // allow re-picking the same file
    drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
    drop.ondragleave = () => drop.classList.remove('over');
    drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); go([...e.dataTransfer.files]); };
  }

  wireDrop('grab', loadGrab);
  wireDrop('brex', files => loadBrex(files[0]));
  // ask before replacing existing results
  $('run').onclick = () => {
    if (!state.rows.length) return run();
    $('rerun').showModal();
  };
  $('rerun').querySelectorAll('button').forEach(b => (b.onclick = () => {
    $('rerun').close();
    if (b.value === 'ok') run();
  }));

  /* ---------- Push to Brex (through the /api/brex Vercel function, which holds the token) ---------- */
  const PUSHABLE = new Set(['Matched', 'Split charge', 'Combined charge']); // Review rows are never pushed
  const needsMemo = r => !String(r._b.memo || '').trim();
  const needsReceipt = r => !String(r._b.receipt || '').trim() && r._b.amount > 0;
  // only charges that are missing a memo or a receipt; whatever Brex already has is never touched
  const pushable = () => state.rows.filter(r => r._b && r._b.id && PUSHABLE.has(r.Status) && r.Memo && !isUp(r) && (needsMemo(r) || needsReceipt(r)));
  const PUSH_BASE = 'https://belli.app.n8n.cloud/webhook/'; // n8n workflow "Grab Receipts: Push to Brex" holds the Brex token; the secret key you type is the last part of the address
  const KEY_STORE = 'grab-push-key';
  let pushKey = '';
  try { pushKey = localStorage.getItem(KEY_STORE) || ''; } catch (e) { /* storage blocked */ }
  const b64 = buf => { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };

  async function brexApi(body) {
    let r;
    try {
      r = await fetch(PUSH_BASE + encodeURIComponent(pushKey), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    } catch (e) { throw new Error('Could not reach n8n (check the workflow is published and the site origin is allowed)'); }
    const j = await r.json().catch(() => ({}));
    if (r.status === 404 || r.status === 403) throw Object.assign(new Error('Wrong push key (or the n8n workflow is not published).'), { status: 401 });
    if (!r.ok) throw Object.assign(new Error(j.error || j.message || `HTTP ${r.status}`), { status: r.status });
    return j;
  }

  async function pushOne(r) {
    const m = state.matches[r._m];
    const receipts = r._b.amount > 0 ? m.grabs.map(g => ({ name: R.receiptFileName(g), pdfBase64: b64(R.buildReceipt(g).output('arraybuffer')) })) : [];
    return brexApi({ action: 'push', expenseId: r._b.id, memo: r.Memo, receipts });
  }

  function openPush() {
    const dlg = $('push-dlg'), body = $('push-body'), actions = $('push-actions');
    const setActions = list => {
      actions.innerHTML = list.map(([label, id, cls]) => `<button type="button" id="${id}" class="${cls || ''}">${label}</button>`).join('');
      return id => $(id);
    };
    const close = () => { dlg.close(); };
    let running = false;

    const askPassword = (err) => {
      body.innerHTML = `<p class="muted">Enter the push key (the secret part of the n8n webhook address). Without it, the page cannot reach Brex.</p>
        <label>Push key<input id="push-pw" type="password" autocomplete="off"></label>
        <label class="inline"><input type="checkbox" id="push-rem"> Remember on this computer</label>${err ? `<p class="push-warn">${esc(err)}</p>` : ''}`;
      setActions([['Continue', 'push-go'], ['Cancel', 'push-cancel', 'secondary']]);
      $('push-cancel').onclick = close;
      const go = async () => {
        const k = $('push-pw').value.trim();
        if (!k) return $('push-pw').focus();
        pushKey = k;
        try {
          await brexApi({ action: 'ping' });
          try { if ($('push-rem').checked) localStorage.setItem(KEY_STORE, k); } catch (e) { /* storage blocked */ }
          preview();
        } catch (e) { pushKey = ''; askPassword(e.status === 401 ? e.message : e.message); }
      };
      $('push-go').onclick = go;
      $('push-pw').onkeydown = e => { if (e.key === 'Enter') go(); };
      $('push-pw').focus();
    };

    const preview = async () => {
      const rows = pushable();
      const skipped = state.rows.filter(r => r._b && !rows.includes(r));
      const reviewN = skipped.filter(r => !PUSHABLE.has(r.Status) && !isUp(r)).length;
      if (!rows.length) {
        body.innerHTML = '<p>Nothing to push: no matched Brex charges are left without a tick.</p>';
        setActions([['Close', 'push-cancel', 'secondary']]); $('push-cancel').onclick = close; return;
      }
      const first = rows[0];
      const qbN = rows.filter(r => r['Integration ERP']).length;
      body.innerHTML = `<p><strong>${rows.length}</strong> matched charge${rows.length === 1 ? '' : 's'} are missing a memo or a receipt and will get them from here${reviewN ? `. ${reviewN} Review/other rows are skipped.` : '.'}</p>
        <p class="hint">${rows.length - qbN} have no QuickBooks link yet, so Brex is all they need. ${qbN} already exist in QuickBooks: this button does not touch QuickBooks. Use \"QuickBooks to-do (.csv)\" to get their memo, GL, class and receipt list for adding there.</p>
        <p class="hint">Only missing things are written. A memo or receipt Brex already has is left alone, and nothing is deleted.</p>
        <p class="hint">Class and GL account are not part of the Brex API (they come from QuickBooks), so copy those by hand.</p>
        <div id="push-first" class="muted">Reading the first charge from Brex…</div>`;
      setActions([['Test with 1 charge', 'push-one'], [`Push all ${rows.length}`, 'push-all', 'secondary'], ['Cancel', 'push-cancel', 'secondary']]);
      $('push-cancel').onclick = close;
      $('push-one').onclick = () => run(rows.slice(0, 1));
      $('push-all').onclick = () => run(rows);
      try {
        const cur = await brexApi({ action: 'get', expenseId: first._b.id });
        $('push-first').innerHTML = `<div class="push-cmp"><b>Charge</b><span>${esc(first.Date)} · ${esc(first['Purchased Amount'])} ${esc(first['Purchased Currency'])}</span>
          <b>Memo now</b><span>${esc(cur.memo) || '<em>empty</em>'}</span><b>New memo</b><span>${cur.memo ? '<em>kept (already has one)</em>' : esc(first.Memo)}</span><b>Receipt</b><span>${cur.receipts ? '<em>already has one, skipped</em>' : esc(first['Receipt PDF'])}</span></div>`;
      } catch (e) {
        $('push-first').innerHTML = `<p class="push-warn">${esc(e.message)}</p>`;
      }
    };

    const run = async rows => {
      running = true;
      body.innerHTML = `<p id="push-st">Pushing 0 of ${rows.length}…</p><div class="push-bar"><span id="push-fill"></span></div><ul class="push-log" id="push-log"></ul>`;
      setActions([['Stop', 'push-stop', 'secondary']]);
      let stop = false, ok = 0, fail = 0;
      $('push-stop').onclick = () => { stop = true; $('push-stop').disabled = true; };
      const log = (cls, t) => { $('push-log').insertAdjacentHTML('afterbegin', `<li class="${cls}">${esc(t)}</li>`); };
      for (let i = 0; i < rows.length && !stop; i++) {
        const r = rows[i];
        try {
          const res = await pushOne(r);
          toggleUploaded(rowKey(r), true);
          ok++;
          log('ok', `✓ ${r.Date} ${r['Purchased Amount']} ${r['Purchased Currency']}: memo ${res.memo}, ${res.receiptSkipped ? 'receipt already there' : res.receipts + ' receipt' + (res.receipts === 1 ? '' : 's') + ' added'} · ${r['Integration ERP'] ? 'also needs QuickBooks update' : 'Brex only'}`);
        } catch (e) {
          fail++;
          log('err', `✗ ${r.Date} ${r['Purchased Amount']} ${r['Purchased Currency']}: ${e.message}`);
          if (e.status === 401 || e.status === 404 || e.status === 503) { stop = true; break; } // won't get better by retrying the next row
        }
        $('push-st').textContent = `Pushing ${ok + fail} of ${rows.length}…`;
        $('push-fill').style.width = `${((ok + fail) / rows.length) * 100}%`;
      }
      running = false;
      $('push-st').textContent = `Done: ${ok} pushed${fail ? `, ${fail} failed (left unticked)` : ''}${stop && ok + fail < rows.length ? ', stopped early' : ''}.`;
      setActions([['Close', 'push-cancel']]); $('push-cancel').onclick = () => { close(); renderResults(); };
    };

    dlg.onclose = () => { if (!running) renderResults(); };
    dlg.showModal();
    pushKey ? preview() : askPassword();
  }
  $('push-brex').onclick = openPush;

  $('dl-xlsx').onclick = downloadXlsx;
  $('dl-zip').onclick = downloadZip;
  $('dl-qb').onclick = downloadQbList;
  $('hide-up').onchange = e => { state.hideUploaded = e.target.checked; renderResults(); };

  // Save: first time asks for a name; after that it updates the same session
  $('save-session').onclick = () => {
    if (state.sessionId) return saveSession(state.sessionName);
    $('save-name').value = defaultSessionName();
    $('save-dlg').showModal();
    $('save-name').select();
  };
  $('save-dlg').querySelectorAll('button').forEach(b => (b.onclick = () => {
    const name = $('save-name').value.trim();
    if (b.value === 'ok' && !name) return $('save-name').focus();
    $('save-dlg').close();
    if (b.value === 'ok') saveSession(name).catch(e => renderSaveState('Could not save: ' + e.message));
  }));
  $('save-name').onkeydown = e => { if (e.key === 'Enter') $('save-dlg').querySelector('button[value=ok]').click(); };
  $('import-session').onchange = e => { if (e.target.files[0]) importSession(e.target.files[0]); e.target.value = ''; };
  renderSaveState();
  renderSessions();
})();
