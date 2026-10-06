/* Draws a Grab receipt PDF with jsPDF, laid out like the n8n "Extract Passenger Info" HTML template. */
(function () {
  const M = window.GrabMatch;
  const C = { ink: [31, 36, 48], muted: [107, 114, 128], line: [231, 232, 236], head: [248, 250, 252], blue: [37, 99, 235], red: [220, 38, 38] };
  const n = v => M.cleanNum(v).toLocaleString('en-US');

  function receiptFileName(g) {
    const safe = s => String(s).replace(/[\\/:*?"<>|]+/g, '').trim();
    return safe(`${g.name} - ${g.currency} ${g.amount.toLocaleString('en-US')} - ${g.bookingId || 'receipt'}.pdf`);
  }

  function buildReceipt(g) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    const r = g.raw, X = 12, W = 186;
    const food = g.category !== 'Transport';
    const money = v => `${g.currency} ${M.cleanNum(v).toLocaleString('en-US')}`;
    const totalFare = money(g.amount);
    const text = (s, x, y, o = {}) => {
      doc.setFont('helvetica', o.bold ? 'bold' : 'normal'); doc.setFontSize(o.size || 9);
      doc.setTextColor(...(o.color || C.ink));
      doc.text(String(s ?? '-'), x, y, { align: o.align || 'left', maxWidth: o.max });
    };

    doc.setDrawColor(...C.line); doc.setLineWidth(0.3);

    // summary header bar
    doc.setFillColor(...C.head); doc.rect(X + 4, 16, W - 8, 16, 'F');
    const rawDate = r['Date & Time (GMT+8)'];
    const dateStr = typeof rawDate === 'string' && /[a-z]/i.test(rawDate) ? rawDate : M.fmtDateTime(g.date, 8);
    const cols = [['Date & Time (GMT+8)', dateStr], ['Employee Name', r['Employee Name']], ['Service Type', r['Service Type']],
                  [food ? 'Total' : 'Total Fare', totalFare], ['Payment Method', r['Payment Method']], ['Employee Group', r['Employee Group']]];
    const cw = (W - 8) / cols.length;
    cols.forEach(([k, v], i) => {
      text(k, X + 6 + i * cw, 21, { size: 6.5, color: C.muted });
      text(v || '-', X + 6 + i * cw, 27, { size: 7.5, bold: true, max: cw - 3 });
    });

    // left: fare breakdown box
    const bx = X + 4, by = 38, bw = 58;
    const num = k => M.cleanNum(r[k]);
    const spec = food
      ? { head: 'Food', base: ['Subtotal', num('Subtotal')], promo: num('Promo Amount'), fees: 'Fees',
          feeLines: [['Delivery fee', 'Delivery Fee'], ['Platform & partner fee', 'Platform & Partner Fee'], ['Small order fee', 'Small Order Fee'],
                     ['Parking fee', 'Parking Fee'], ['Other fees', 'Other Fees'], ['Tips', 'Tips']] }
      : { head: 'Ride fare', base: ['Base fare', num('Base Fare')], promo: num('Promo'), fees: 'Miscellaneous',
          feeLines: [['Booking fee', 'Booking Fee'], ['Tolls', 'Tolls'], ['Late fees', 'Late Fee'], ['Other fees', 'Other Fees']] };
    const promo = Math.abs(spec.promo);
    const fees = spec.feeLines.map(([, k]) => num(k));
    let y = by + 7;
    const line = (label, val, o = {}) => {
      text(label, bx + (o.indent ? 7 : 4), y, { size: o.size || 8.5, bold: o.bold, color: o.indent ? C.muted : C.ink });
      text(val, bx + bw - 4, y, { size: o.size || 8.5, bold: o.bold, align: 'right', color: o.indent ? C.muted : C.ink });
      y += o.gap || 5;
    };
    line(spec.head, n(spec.base[1] - promo), { bold: true });
    line(spec.base[0], n(spec.base[1]), { indent: true });
    line('Promo', promo > 0 ? '-' + n(promo) : '0', { indent: true, gap: 8 });
    line(spec.fees, n(fees.reduce((a, b) => a + b, 0)), { bold: true });
    spec.feeLines.forEach(([label], i) => line(label, n(fees[i]), { indent: true }));
    doc.setLineDashPattern([1, 1], 0); doc.line(bx + 4, y, bx + bw - 4, y); doc.setLineDashPattern([], 0);
    y += 7;
    line('Total', totalFare, { bold: true, size: 10 });
    if (g.refund) {
      line('Refund', '-' + money(g.refund), { indent: true });
      line('Net paid', money(g.amount - g.refund), { bold: true });
    }
    const boxBottom = y;

    // right: booking metadata
    const rx = bx + bw + 8, rw = W - bw - 16, mw = rw / 3;
    const meta = food
      ? [['Order ID', g.bookingId], ['Merchant', g.merchant], ['City', r['City']],
         ['Order / Cost Code', r['Order / Cost Code']], ['Order Source', r['Order Source']], ['Description', r['Order Description']]]
      : [['Booking ID', g.bookingId], ['Trip / Cost Code', r['Trip / Cost Code']], ['Distance', r['Distance'] ? `${r['Distance']} km` : '-'],
         ['City', r['City']], ['Booking Source', r['Booking Source']], ['Description', r['Description'] || r['Trip Description']]];
    meta.forEach(([k, v], i) => {
      const cx = rx + (i % 3) * mw, cy = by + 6 + Math.floor(i / 3) * 14;
      text(k, cx, cy, { size: 6.5, color: C.muted });
      text(v && v !== '-' ? v : '-', cx, cy + 5, { size: 8.5, bold: true, max: mw - 3 });
    });

    // route (ride) or merchant -> delivery address (food)
    doc.setDrawColor(...C.line); doc.line(rx, by + 34, rx + rw, by + 34);
    const from = food && g.merchant && !String(g.pickup).includes(g.merchant) ? `${g.merchant}, ${g.pickup}` : g.pickup;
    doc.setFillColor(...C.blue); doc.circle(rx + 1.5, by + 42.5, 1.2, 'F');
    text(from || '-', rx + 5, by + 43.5, { size: 8.5, max: rw - 6 });
    doc.setFillColor(...C.red); doc.circle(rx + 1.5, by + 56.5, 1.2, 'F');
    text(g.dropoff || '-', rx + 5, by + 57.5, { size: 8.5, max: rw - 6 });

    // food: ordered items
    let ry = by + 66;
    if (food) {
      let items = [];
      try { items = JSON.parse(r['Items'] || '[]'); } catch (e) { /* not JSON: show nothing */ }
      if (items.length) {
        doc.line(rx, ry, rx + rw, ry); ry += 6;
        text('Items', rx, ry, { size: 6.5, color: C.muted }); ry += 5;
        for (const it of items) {
          text(`${it.quantity || 1} × ${it.item}`, rx, ry, { size: 8, max: rw - 30 });
          text(n(it.price), rx + rw, ry, { size: 8, align: 'right' });
          ry += 4.5;
        }
      }
      if (r['Recipient Name'] && r['Recipient Name'] !== '-') { ry += 2; text(`Recipient: ${r['Recipient Name']}`, rx, ry, { size: 7.5, color: C.muted }); ry += 4; }
    }

    // card + breakdown box outlines, sized to the content
    const bottom = Math.max(by + 82, boxBottom + 2, ry + 2);
    doc.roundedRect(bx, by, bw, bottom - by, 2, 2);
    doc.roundedRect(X, 12, W, bottom - 12 + 4, 2, 2);

    return doc;
  }

  window.GrabReceipt = { buildReceipt, receiptFileName };
})();
