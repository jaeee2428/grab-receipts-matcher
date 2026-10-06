/* Pure logic: parsing Grab for Business rows, normalising Brex rows, matching. No DOM. */
(function (root) {
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const DAY = 86400000, HOUR = 3600000;
  const BOOKING_RE = /A-[A-Z0-9]{8,}/;

  function cleanNum(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? 0 : n;
  }

  // "IDR 15,500" | "SGD 32.45" | 15500 -> {amount, currency}
  function parseMoney(v, fallbackCurrency) {
    if (typeof v === 'number') return { amount: v, currency: fallbackCurrency || '' };
    const s = String(v ?? '').trim();
    const m = s.match(/[A-Za-z]{3}/);
    return { amount: cleanNum(s), currency: m ? m[0].toUpperCase() : (fallbackCurrency || '') };
  }

  // Excel serial | "2026-09-01 10:23:45Z" | "09/01/2026" (US) | "01 Sep 2026 03:43:21 PM" -> Date.
  // Wall-clock values are stored as UTC fields; tzHours shifts them to real UTC (e.g. 8 for GMT+8).
  function parseDate(v, tzHours = 0) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return v;
    const shift = tzHours * HOUR;
    const n = Number(v);
    if (!isNaN(n) && n > 30000 && n < 80000) return new Date(Date.UTC(1899, 11, 30) + Math.round(n * DAY) - shift);
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) return new Date(Date.UTC(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) - shift);
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) return new Date(Date.UTC(+m[3], m[1] - 1, +m[2], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) - shift);
    m = s.match(/^(\d{1,2})[ \-]([A-Za-z]{3})[a-z]*[ \-,]+(\d{4})(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?/i);
    if (m) {
      let h = +(m[4] || 0);
      if (m[7]) h = (h % 12) + (/pm/i.test(m[7]) ? 12 : 0);
      const mon = MONTHS.findIndex(x => x.toLowerCase() === m[2].toLowerCase());
      return new Date(Date.UTC(+m[3], mon, +m[1], h, +(m[5] || 0), +(m[6] || 0)) - shift);
    }
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  const p2 = x => String(x).padStart(2, '0');
  // tzHours: show the date in that timezone (Grab receipts are GMT+8)
  function fmtDate(d, tzHours = 0) {
    if (!d) return '';
    d = new Date(+d + tzHours * HOUR);
    return `${p2(d.getUTCDate())} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
  }
  function fmtDateTime(d, tzHours = 0) {
    if (!d) return '-';
    d = new Date(+d + tzHours * HOUR);
    let h = d.getUTCHours(); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12;
    return `${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}, ${p2(h)}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())} ${ap}`;
  }

  function displayName(raw) {
    if (!raw) return 'Employee';
    const s = String(raw).trim();
    if (/thareq|abel/i.test(s)) return 'Abel';
    if (/januario/i.test(s)) return 'Ino';
    return s.split(/\s+/)[0];
  }

  /* ---------- Grab ---------- */
  // Grab for Business exports one file per category (Transport, Food, Express, Mart...).
  const CATEGORIES = [['Transport', /transport|ride/i], ['Food', /food/i], ['Express', /express|deliver/i], ['Mart', /mart/i]];
  function guessCategory(fileName, rows) {
    for (const [c, re] of CATEGORIES) if (re.test(fileName || '')) return c;
    const svc = rows.map(r => r['Service Type'] || r['Booking Type'] || '').join(' ');
    for (const [c, re] of CATEGORIES) if (re.test(svc)) return c;
    return 'Transport';
  }
  // first column whose header matches one of the patterns (exports differ per category)
  function pick(r, pats) {
    const keys = Object.keys(r);
    for (const p of pats) { const k = keys.find(k => p.test(k)); if (k && r[k] !== '' && r[k] !== '-') return r[k]; }
    return '';
  }

  // matrix: sheet as array of arrays. Finds the header row containing "Booking ID".
  function parseGrabMatrix(matrix, fileName) {
    const ID_HEAD = /^(booking( \/ order)?|order) id$/i; // Transport: "Booking ID", Food: "Booking / Order ID"
    const h = matrix.findIndex(r => r && r.some(c => ID_HEAD.test(String(c).trim())));
    if (h < 0) throw new Error('Could not find a "Booking ID" header row in the Grab file.');
    const head = matrix[h].map(c => String(c ?? '').trim());
    const rows = [];
    for (let i = h + 1; i < matrix.length; i++) {
      const row = matrix[i]; if (!row || !row.length) continue;
      const o = {}; head.forEach((k, j) => { if (k) o[k] = row[j]; });
      if (!pick(o, [ID_HEAD]) && !pick(o, [/^total/i])) continue;
      rows.push(o);
    }
    const category = guessCategory(fileName, rows);
    return rows.map((r, i) => normalizeGrab(r, i, category, fileName));
  }

  function normalizeGrab(r, i, category = 'Transport', file = '') {
    const cur = String(pick(r, [/^currency( code)?$/i])).trim().toUpperCase();
    // amount = what the card was charged; a refund afterwards shows up as a lower revised total
    const fare = parseMoney(pick(r, [/^total fare$/i, /^total$/i, /^total amount$/i, /^order total$/i, /^amount$/i]), cur);
    const revised = pick(r, [/^revised total/i]);
    const net = revised === '' ? fare.amount : cleanNum(revised);
    return {
      idx: i, category, file, bookingId: String(pick(r, [/^booking( \/ order)? id$/i, /^order id$/i])).trim(),
      date: parseDate(pick(r, [/^date & time/i, /^date/i, /time/i]), 8),
      amount: fare.amount, currency: fare.currency, refund: net < fare.amount ? +(fare.amount - net).toFixed(2) : 0,
      employee: r['Employee Name'] || '', name: displayName(r['Employee Name']),
      merchant: r['Merchant'] || '',
      pickup: pick(r, [/^pick-?up address$/i, /sender/i]),
      dropoff: pick(r, [/^drop-?off address$/i, /^deliver(y)? address$/i, /recipient address/i]),
      service: r['Service Type'] || '', city: r['City'] || '',
      payment: String(r['Payment Details'] || r['Payment Method'] || '').replace(/^\["?|"?\]$/g, ''),
      last4: (String(r['Payment Details'] || '').match(/(\d{4})\D*$/) || ['', ''])[1], raw: r,
    };
  }

  /* ---------- Brex ---------- */
  // First pattern that hits wins, so the most specific Brex export header comes first.
  const BREX_FIELDS = {
    id:         [/^expense id$/i, /^id$/i, /transaction id/i],
    time:       [/swipe time/i],
    date:       [/^transaction date$/i, /^(posted|purchase|initiated)?\s*date$/i],
    amount:     [/^purchased amount$/i, /^original amount$/i, /^amount$/i],
    currency:   [/^purchased currency$/i, /^original currency$/i, /^currency$/i],
    usd:        [/usd equivalent/i, /usd amount/i, /amount.*usd/i],
    descriptor: [/statement descriptor/i],
    merchant:   [/^merchant name$/i, /merchant/i],
    memo:       [/^memo$/i],
    user:       [/^integration location$/i, /^location$/i, /^user$/i],
    erp:        [/integration erp link/i, /erp link/i, /quickbooks|qbo/i],
    link:       [/created from url/i, /brex.*(link|url)/i, /dashboard/i],
    receipt:    [/^receipt url$/i],
    status:     [/^expense status$/i, /^status$/i],
    last4:      [/^last 4$/i, /last ?4|card last/i],
    gl:         [/^integration debit line category$/i],
    cls:        [/^integration class$/i],
  };

  function autoMapBrex(headers) {
    const map = {};
    for (const [field, pats] of Object.entries(BREX_FIELDS)) {
      for (const p of pats) {
        const h = headers.find(h => p.test(String(h).trim()) && !Object.values(map).includes(h));
        if (h) { map[field] = h; break; }
      }
    }
    return map;
  }

  function normalizeBrex(rows, map, linkTemplate) {
    return rows.map((r, i) => {
      const g = f => (map[f] ? r[map[f]] : undefined);
      const s = f => String(g(f) ?? '').trim();
      const m = parseMoney(g('amount'), s('currency').toUpperCase());
      const id = s('id');
      let link = s('link');
      if (!link && linkTemplate && id) link = linkTemplate.replace('{id}', encodeURIComponent(id));
      const desc = s('descriptor');
      return {
        idx: i, id, date: parseDate(g('time')) || parseDate(g('date')), amount: m.amount, currency: m.currency,
        usd: s('usd') !== '' ? cleanNum(g('usd')) : null,
        descriptor: desc, bookingRef: (desc.match(BOOKING_RE) || [''])[0], merchant: s('merchant'),
        memo: s('memo'), user: s('user'), erp: s('erp'), link, receipt: s('receipt'), status: s('status'),
        gl: s('gl'), cls: s('cls').replace(/^\[|\]$/g, ''),
        last4: s('last4').replace(/\D/g, '').slice(-4).padStart(s('last4') ? 4 : 0, '0'), raw: r,
      };
    }).filter(b => b.amount !== 0);
  }

  /* ---------- Matching ---------- */
  const eq = (a, b, tol = 0.011) => Math.abs(a - b) <= tol;
  const curOk = (a, b) => !a || !b || a === b;
  const dayDiff = (a, b) => (a && b ? Math.abs(a - b) / DAY : 3);
  const sum = (xs, f = x => x.amount) => +xs.reduce((s, x) => s + f(x), 0).toFixed(2);
  const mins = (a, b) => (a && b ? (b - a) / 60000 : Infinity); // signed: b after a
  const absMins = (a, b) => Math.abs(mins(a, b));

  // any token of the Grab employee name (or its display alias, e.g. Januario -> Ino) appears in Brex location/memo
  function nameHit(g, u) {
    const hay = ' ' + (u.user + ' ' + u.memo).toLowerCase().replace(/[^a-z0-9]+/g, ' ') + ' ';
    const toks = [g.name, ...String(g.employee || '').split(/\s+/)].map(t => String(t).toLowerCase().replace(/[^a-z0-9]/g, '')).filter(t => t.length > 2);
    return toks.some(t => hay.includes(' ' + t + ' ')) ? 1 : 0;
  }
  // Brex names someone else explicitly -> maybe not this employee's trip (or a memo typo)
  const nameClash = (g, u) => !nameHit(g, u) && u.lines.some(b => b.user || /^[A-Za-z]+,/.test(b.memo));
  const brexWho = u => [...new Set(u.lines.map(b => b.user || b.memo.split(',')[0]).filter(Boolean))].join(' / ');
  function gapText(g, u) {
    const m = Math.round(mins(g.date, u.date)), a = Math.abs(m);
    if (!isFinite(m)) return '';
    const t = a < 60 ? `${a} min` : a < 2880 ? `${Math.round(a / 60)} h` : `${Math.round(a / 1440)} days`;
    return `charged ${t} ${m >= 0 ? 'after' : 'before'} trip`;
  }

  // subset (minN..maxN) of items whose amounts sum to target; items pre-sorted by relevance
  function subsetSum(items, target, minN = 2, maxN = 4, tol = 0.011) {
    const pool = items.slice(0, 20);
    let best = null;
    (function dfs(start, picked, s) {
      if (best) return;
      if (picked.length >= minN && eq(s, target, tol)) { best = picked.slice(); return; }
      if (picked.length >= maxN || s > target + tol) return;
      for (let i = start; i < pool.length && !best; i++) {
        picked.push(pool[i]); dfs(i + 1, picked, s + pool[i].amount); picked.pop();
      }
    })(0, [], 0);
    return best;
  }

  // Brex lines that share a descriptor booking ID are one Grab charge (Grab splits a ride into
  // an upfront charge + top-ups). Lines without an ID are units of their own.
  function buildUnits(brexes) {
    const units = [], byRef = new Map();
    for (const b of [...brexes].sort((x, y) => (x.date || 0) - (y.date || 0))) {
      let u = b.bookingRef && byRef.get(b.bookingRef);
      if (!u) {
        u = { ref: b.bookingRef, lines: [], currency: b.currency, date: b.date };
        units.push(u); if (b.bookingRef) byRef.set(b.bookingRef, u);
      }
      u.lines.push(b);
    }
    for (const u of units) {
      u.amount = sum(u.lines);
      u.user = u.lines.map(b => b.user).join(' ');
      u.memo = u.lines.map(b => b.memo).join(' ');
      const tag = u.lines.map(b => b.merchant + ' ' + b.descriptor).join(' ');
      u.kind = u.amount < 0 ? 'refund' : /grab-?ec/i.test(tag) && !/rides/i.test(tag) ? 'ec' : /food/i.test(tag) ? 'food'
        : /grab/i.test(tag) ? 'ride' : 'other';
    }
    return units;
  }
  // Brex "Food Grab" merchant is unreliable (rides get it too), so only GRAB-EC is kept away from rides
  const kindOk = (g, u) => u.kind !== 'other' && u.kind !== 'refund' && (g.category !== 'Transport' || u.kind !== 'ec');

  /*
   * Returns [{type, method, note, grabs[], brex[]}].
   *  1. Booking ID: descriptor ID == trip ID. Short sums are topped up with ID-less ride lines charged within 3 h.
   *  2. Exact amount, charged within 30 min of the trip (any Brex unit, incl. IDs missing from the report).
   *  3. Anchored split: one unit within 30 min of the trip + up to 3 more ride units in the next 3 h = fare.
   *  4. Exact amount within the date window, same person (no swipe time / late posting).
   *  5. One charge = several trips of the same employee, same day.
   *  6. Amount within 3%, same day -> review.
   * A Brex memo naming someone else never blocks a time-anchored match; it flags "Review: name differs".
   */
  function matchAll(grabs, brexes, opts = {}) {
    const win = opts.dateWindowDays ?? 2;
    const card = String(opts.card || '').replace(/\D/g, '');
    const out = [];
    // card filter: Brex lines on other cards are dropped; Grab trips paid with another card are kept but not matched
    if (card) {
      brexes = brexes.filter(b => !b.last4 || b.last4 === card);
      grabs.filter(g => g.last4 && g.last4 !== card).forEach(g =>
        out.push({ type: 'Other card', method: '', note: `Paid with ${g.payment}`, grabs: [g], brex: [] }));
      grabs = grabs.filter(g => !g.last4 || g.last4 === card);
    }
    const units = buildUnits(brexes);
    const gLeft = new Set(grabs), uLeft = new Set(units);
    const take = (type, method, gs, us, note = '') => {
      gs.forEach(g => gLeft.delete(g)); us.forEach(u => uLeft.delete(u));
      out.push({ type, method, note, grabs: gs, brex: us.flatMap(u => u.lines) });
    };
    const lineCount = us => us.reduce((n, u) => n + u.lines.length, 0);
    // Booking-ID and minute-level time matches are certain: a different name in the Brex memo is reported in
    // the Name Check column (memo probably wrong). Weaker matches with a different name go to review.
    const result = (g, us, method, extra = '', certain = true) => {
      const clash = us.some(u => nameClash(g, u)) && !us.some(u => nameHit(g, u));
      const n = lineCount(us);
      const nameCheck = clash ? `Brex memo: ${us.map(brexWho).filter(Boolean).join(' / ')} · Grab: ${g.employee}` : 'OK';
      const note = [n > 1 ? `${n} Brex charges` : '', extra].filter(Boolean).join(' · ');
      const type = clash && !certain ? 'Review: name differs' : n > 1 ? 'Split charge' : 'Matched';
      take(type, method, [g], us, note);
      out[out.length - 1].nameCheck = nameCheck;
    };
    // ride units charged after `from` (within `upto` minutes), sorted by time, for topping up a split
    const followUps = (g, from, upto, exclude) => [...uLeft]
      .filter(u => !exclude.includes(u) && u.kind === 'ride' && !u.ref && curOk(g.currency, u.currency) && u.amount > 0)
      .filter(u => { const m = mins(from, u.date); return m >= -5 && m <= upto; })
      .sort((x, y) => x.date - y.date);

    // pass 1: booking ID
    const byId = new Map(grabs.filter(g => g.bookingId).map(g => [g.bookingId, g]));
    for (const u of units) {
      const g = u.ref && byId.get(u.ref); if (!g || !gLeft.has(g)) continue;
      if (!curOk(g.currency, u.currency)) { take('Review: amount differs', 'Booking ID', [g], [u], `Grab ${g.currency} vs Brex ${u.currency}`); continue; }
      if (eq(u.amount, g.amount)) { result(g, [u], 'Booking ID'); continue; }
      const top = u.amount < g.amount && subsetSum(followUps(g, u.date, 180, [u]), +(g.amount - u.amount).toFixed(2), 1, 3);
      if (top) result(g, [u, ...top], 'Booking ID + amount', 'top-up has no booking ID');
      else take('Review: amount differs', 'Booking ID', [g], [u], `Grab ${g.currency} ${g.amount} vs Brex ${u.currency} ${u.amount}`);
    }

    // pass 2: exact amount, charged within 30 min of the trip
    const pairs = [];
    for (const g of gLeft) for (const u of uLeft) {
      if (!kindOk(g, u) || !curOk(g.currency, u.currency) || !eq(g.amount, u.amount)) continue;
      const m = absMins(g.date, u.date);
      if (m <= 30) pairs.push({ g, u, s: m - nameHit(g, u) * 30 });
    }
    pairs.sort((x, y) => x.s - y.s);
    for (const { g, u } of pairs) if (gLeft.has(g) && uLeft.has(u)) result(g, [u], u.ref ? 'Amount + time (other ID)' : 'Amount + time', gapText(g, u), absMins(g.date, u.date) <= 10);

    // pass 3: anchored split
    for (const g of [...gLeft].sort((x, y) => (x.date || 0) - (y.date || 0))) {
      const anchors = [...uLeft]
        .filter(u => kindOk(g, u) && u.amount > 0 && u.amount < g.amount && curOk(g.currency, u.currency) && absMins(g.date, u.date) <= 30)
        .sort((x, y) => absMins(g.date, x.date) - absMins(g.date, y.date));
      for (const a of anchors) {
        const top = subsetSum(followUps(g, a.date, 180, [a]), +(g.amount - a.amount).toFixed(2), 1, 3);
        if (top) { result(g, [a, ...top], 'Amount + time (split)', gapText(g, a), absMins(g.date, a.date) <= 10); break; }
      }
    }

    // pass 4: exact amount within the date window, same person
    const loose = [];
    for (const g of gLeft) for (const u of uLeft) {
      if (!kindOk(g, u) || !curOk(g.currency, u.currency) || !eq(g.amount, u.amount)) continue;
      if (dayDiff(g.date, u.date) > win || nameClash(g, u)) continue;
      loose.push({ g, u, s: absMins(g.date, u.date) - nameHit(g, u) * 600 });
    }
    loose.sort((x, y) => x.s - y.s);
    for (const { g, u } of loose) if (gLeft.has(g) && uLeft.has(u)) result(g, [u], 'Amount + date', gapText(g, u), false);

    // pass 5: one charge covering several trips of the same employee (same day)
    for (const u of [...uLeft].filter(u => u.amount > 0)) {
      const cand = [...gLeft]
        .filter(g => kindOk(g, u) && curOk(g.currency, u.currency) && g.amount < u.amount && dayDiff(g.date, u.date) <= 1 && !nameClash(g, u))
        .sort((x, y) => (nameHit(y, u) - nameHit(x, u)) || (absMins(x.date, u.date) - absMins(y.date, u.date)));
      const hit = subsetSum(cand, u.amount);
      if (hit && new Set(hit.map(g => g.name)).size === 1) take('Combined charge', 'Amount sum', hit, [u], `${hit.length} Grab trips`);
    }

    // pass 6: amount within 3%, same day, same person -> review
    const near = [];
    for (const g of gLeft) for (const u of uLeft) {
      if (!kindOk(g, u) || u.amount <= 0 || !curOk(g.currency, u.currency) || nameClash(g, u) || dayDiff(g.date, u.date) > 1) continue;
      const diff = Math.abs(g.amount - u.amount) / g.amount;
      if (diff <= 0.03) near.push({ g, u, s: diff * 100 + absMins(g.date, u.date) / 60 - nameHit(g, u) * 20 });
    }
    near.sort((x, y) => x.s - y.s);
    for (const { g, u } of near) if (gLeft.has(g) && uLeft.has(u))
      take('Review: amount differs', 'Near amount', [g], [u], `Grab ${g.currency} ${g.amount} vs Brex ${u.currency} ${u.amount}`);

    // refunds: a Grab order with a lower revised total gets the matching negative Brex line attached
    for (const m of out) {
      const g = m.grabs.length === 1 && m.brex.length && m.grabs[0];
      if (!g || !g.refund) continue;
      const r = [...uLeft].filter(u => u.kind === 'refund' && curOk(g.currency, u.currency) && eq(-u.amount, g.refund) && dayDiff(g.date, u.date) <= 3)
        .sort((x, y) => absMins(g.date, x.date) - absMins(g.date, y.date))[0];
      if (r) { uLeft.delete(r); m.brex.push(...r.lines); }
      m.note = [m.note, `refund ${g.currency} ${g.refund}${r ? ' (Brex credit attached)' : ' (no Brex credit found)'} → net ${+(g.amount - g.refund).toFixed(2)}`].filter(Boolean).join(' · ');
    }

    // leftovers
    // Brex export range in whole days (UTC); trips outside it can't have a charge in this file
    const brexMax = Math.ceil(Math.max(0, ...brexes.map(b => +b.date || 0)) / DAY) * DAY;
    const brexMin = Math.floor(Math.min(Infinity, ...brexes.map(b => +b.date || Infinity)) / DAY) * DAY;
    for (const g of gLeft) {
      const outside = brexes.length && g.date && (+g.date >= brexMax || +g.date < brexMin);
      take(brexes.length ? (outside ? 'Outside Brex range' : 'Unmatched Grab') : 'Receipt only', '', [g], [],
        outside ? `Brex export covers ${fmtDate(new Date(brexMin))} to ${fmtDate(new Date(brexMax - DAY))}` : '');
    }
    const cats = new Set(grabs.map(g => g.category));
    for (const u of uLeft) {
      const notInReport = u.ref ? `Booking ${u.ref} not in the uploaded Grab report(s)` : '';
      const n = u.lines.length > 1 ? `${u.lines.length} Brex charges` : '';
      const note = [n, notInReport].filter(Boolean).join(' · ');
      if (u.kind === 'refund') take('Refund / credit', '', [], [u], note);
      else if (u.kind === 'other') take('Not Grab', '', [], [u], note);
      else if (u.kind === 'ec' && !cats.has('Food')) take('Grab-ec (food/mart, no ride)', '', [], [u], [note, 'upload the Grab Food report to match'].filter(Boolean).join(' · '));
      else if (u.kind === 'food' && !cats.has('Food')) take('Grab Food (no ride)', '', [], [u], [note, 'upload the Grab Food report to match'].filter(Boolean).join(' · '));
      else if (u.ref) take('Not in Grab report', '', [], [u], note + ' (personal profile?)');
      else take('Unmatched Brex', '', [], [u], note);
    }

    const key = m => +(m.grabs[0]?.date || m.brex[0]?.date || 0);
    out.sort((a, b) => key(a) - key(b));
    return out;
  }

  const MEMO_LABEL = { Transport: 'Transpo' };
  // QuickBooks GL account per Grab category (Brex "Integration Debit Line Category")
  const GL_ACCOUNT = { Transport: 'TR03 - Transportation', Food: 'TR04 - Meals' };
  const CLASS = 'travel';
  // "Name, Grab Transpo" / "Name, Grab Food": the name comes from the Grab for Business report
  const memoFor = gs => {
    const names = [...new Set(gs.map(g => g.name))].join(' & ');
    const cats = [...new Set(gs.map(g => MEMO_LABEL[g.category] || g.category))].join(' & ');
    return `${names}, Grab ${cats}`;
  };

  // Flatten to sheet rows (column order follows the Matched_Receipts_ERP sheet, plus Booking ID / Brex link).
  // A match with several Brex lines becomes one row per line, so every Brex expense gets its own memo:
  // split parts end "1/2", "2/2"; an attached refund line ends "refund".
  function toRows(matches, receiptName) {
    const rows = [];
    matches.forEach((m, mi) => {
      const gs = m.grabs, g0 = gs[0];
      const lines = m.brex.length ? [...m.brex].sort((a, b) => (a.amount < 0) - (b.amount < 0) || (a.date || 0) - (b.date || 0)) : [null];
      const parts = lines.filter(b => b && b.amount > 0).length;
      let part = 0;
      for (const b of lines) {
        let memo = b ? b.memo : '';
        if (gs.length) {
          memo = memoFor(gs);
          if (b && b.amount < 0) memo += ' refund';
          else if (parts > 1) memo += ` ${++part}/${parts}`;
        }
        const row = {
          Status: m.type,
          'Match Method': m.method,
          Date: g0 ? fmtDate(g0.date, 8) : fmtDate(b?.date),
          'Grab Category': [...new Set(gs.map(g => g.category))].join(', '),
          'Booking ID': gs.map(g => g.bookingId).join(', ') || (b?.bookingRef ?? ''),
          'Receipt PDF': gs.map(g => receiptName(g)).join(', '),
          'Integration ERP': b?.erp || '',
          'Brex Dashboard': b?.link || '',
          'Purchased Amount': b ? b.amount : (gs.length ? sum(gs) : ''),
          'Purchased Currency': b?.currency || g0?.currency || '',
          'Grab Amount': gs.length ? sum(gs) : '',
          Location: g0 ? [...new Set(gs.map(g => g.name))].join(' & ') : (b?.user || ''),
          Class: g0 ? CLASS : (b?.cls || ''),
          'GL Account': g0 ? (GL_ACCOUNT[g0.category] || '') : (b?.gl || ''),
          'Current Brex GL': b?.gl || '',
          'USD Amount': b && b.usd !== null ? b.usd : '',
          Memo: memo,
          'Current Brex Memo': b?.memo || '',
          'Name Check': m.nameCheck || '',
          Note: m.note,
        };
        Object.defineProperty(row, '_m', { value: mi }); // match index, not a column
        Object.defineProperty(row, '_b', { value: b });  // the Brex line this row is for
        rows.push(row);
      }
    });
    return rows;
  }

  const api = { cleanNum, parseMoney, parseDate, fmtDate, fmtDateTime, displayName, parseGrabMatrix, autoMapBrex,
                normalizeBrex, matchAll, toRows, BREX_FIELDS };
  if (typeof module !== 'undefined') module.exports = api; else root.GrabMatch = api;
})(typeof window !== 'undefined' ? window : globalThis);
