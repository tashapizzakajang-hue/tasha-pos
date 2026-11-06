/**
 * Akif Kitchen X Tasha Pizza Kajang: POS -> Google Sheets (segerak dua hala)
 *
 * Pemasangan: Extensions > Apps Script, tampal kod ini, Save, kemudian
 * Deploy > Manage deployments > pensil > New version > Deploy.
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * Kod akses (disyorkan): isi kod di bawah, dan isi kod yang SAMA dalam app (Tetapan).
 * Kalau dibiarkan kosong, sesiapa yang tahu URL boleh baca dan tulis data jualan.
 */
const KOD_AKSES = '';

/**
 * Tab yang dibuat automatik:
 *   Orders, Items        : data mentah setiap order
 *   Tempahan             : tempahan awal pelanggan
 *   Belanja              : modal masuk dan belanja
 *   Stok Log, Bahan      : pergerakan stok dan senarai bahan
 *   Baki Stok            : baki stok semasa
 *   Kiraan Tunai         : kiraan tunai harian
 *   Tetapan              : menu dan nama kedai yang dikongsi semua peranti
 *   Ringkasan Harian, Ringkasan Item : ringkasan jualan, belanja dan untung
 * Lajur "Data" (tersembunyi) dan "Rev" dipakai untuk menyegerakkan peranti. Jangan padam.
 */

const ORDER_HEADERS = ['ID Order','No','Tarikh','Masa','Saluran','Meja/Nama','Item','Jumlah Kecil (RM)','Diskaun (RM)','Jumlah (RM)','Bayaran','Tunai Diterima (RM)','Baki (RM)','Status','Cashier','Caj Hantar (RM)','Sebab Diskaun','Data','Rev'];
const ITEM_HEADERS  = ['ID Order','No','Tarikh','Masa','Saluran','Item','Harga (RM)','Kuantiti','Jumlah (RM)','Nota','Status'];
const BOOK_HEADERS  = ['ID','Dibuat','Nama','Telefon','Tarikh Ambil','Masa Ambil','Item','Jumlah (RM)','Saluran','Cashier','Nota','Status','ID Order','Data','Rev'];
const EXP_HEADERS   = ['ID','Tarikh','Masa','Jenis','Kategori','Butiran','Jumlah (RM)','Bayaran','Cashier','Status','Data','Rev'];
const LOG_HEADERS   = ['ID','Tarikh','Masa','Bahan','Jenis','Kuantiti','Unit','Kos (RM)','Nota','Cashier','Data','Rev'];
const BAHAN_HEADERS = ['ID','Nama','Unit','Minimum','Status','Data','Rev'];
const SET_HEADERS   = ['Kunci','Dikemas kini','Data','Rev'];
const STOCK_HEADERS = ['Bahan','Unit','Baki','Minimum','Status','Dikemas kini'];
const CASH_HEADERS  = ['Tarikh','Float Awal (RM)','Tunai Jualan (RM)','Modal Tunai (RM)','Belanja Tunai (RM)','Jangkaan (RM)','Dikira (RM)','Beza (RM)','Nota','Cashier','Data','Rev'];
const HARIAN  = 'Ringkasan Harian';
const ITEMSUM = 'Ringkasan Item';

// kunci -> [nama tab, tajuk, lajur teks]
const SPEC = {
  orders: ['Orders', ORDER_HEADERS, 'C:D'],
  items:  ['Items', ITEM_HEADERS, 'C:D'],
  book:   ['Tempahan', BOOK_HEADERS, 'E:F'],
  exp:    ['Belanja', EXP_HEADERS, 'B:C'],
  log:    ['Stok Log', LOG_HEADERS, 'B:C'],
  bahan:  ['Bahan', BAHAN_HEADERS, 'A:A'],
  set:    ['Tetapan', SET_HEADERS, 'A:A'],
  cash:   ['Kiraan Tunai', CASH_HEADERS, 'A:A'],
  stock:  ['Baki Stok', STOCK_HEADERS, 'F:F']
};

// ---------- titik masuk ----------

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'pull' || p.action === 'ping') {
    if (KOD_AKSES && p.key !== KOD_AKSES) return out_({ ok: false, error: 'Kod akses salah' });
    if (p.action === 'ping') return out_({ ok: true, msg: 'Sambungan berjaya', rev: curRev_(), epoch: epoch_() });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      return out_(pull_(SpreadsheetApp.getActiveSpreadsheet(), parseInt(p.since || '0', 10) || 0));
    } catch (err) {
      return out_({ ok: false, error: String(err) });
    } finally {
      lock.releaseLock();
    }
  }
  return out_({ ok: true, msg: 'Tasha POS sedia' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (KOD_AKSES && d.key !== KOD_AKSES) return out_({ ok: false, error: 'Kod akses salah' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (d.action === 'reset') {
      kosongkan_(ss);
      return out_({ ok: true, epoch: epoch_() });
    }

    if (d.action === 'stocksnap') {      // paparan sahaja, tidak disegerakkan
      const s = sh_(ss, 'stock');
      const last = s.getLastRow();
      if (last >= 2) s.getRange(2, 1, last - 1, STOCK_HEADERS.length).clearContent();
      if (d.items.length) {
        const rows = d.items.map(function (i) {
          const low = i.min > 0 && i.qty <= i.min;
          return [i.name, i.unit, i.qty, i.min, low ? (i.qty <= 0 ? 'HABIS' : 'HAMPIR HABIS') : 'OK', d.at];
        });
        s.getRange(2, 1, rows.length, STOCK_HEADERS.length).setValues(rows);
      }
      return out_({ ok: true });
    }

    const rev = nextRev_();
    const data = data_(d);
    const orderNo = (d.dev || '') + d.no;

    if (d.action === 'booking') {
      upsert_(sh_(ss, 'book'), d.id, [d.id, d.created, d.name, d.phone, d.date, d.time, itemText_(d.items), d.total,
        d.channel, d.cashier || '', d.note || '', d.status, d.orderId || '', data, rev]);
      return out_({ ok: true, rev: rev });
    }
    if (d.action === 'expense') {
      upsert_(sh_(ss, 'exp'), d.id, [d.id, d.date, d.time, d.type, d.cat, d.desc || '', d.amount, d.method,
        d.cashier || '', d.status || 'OK', data, rev]);
      try { ensureDayRow_(ss, ensureSummary_(ss), d.date); } catch (err) { /* abaikan */ }
      return out_({ ok: true, rev: rev });
    }
    if (d.action === 'stock') {
      const s = sh_(ss, 'log');
      const r = findRow_(s, d.id);
      const row = [d.id, d.date, d.time, d.item, d.type, d.qty, d.unit, d.cost || 0, d.note || '', d.cashier || '', data, rev];
      if (!r) s.appendRow(row);
      else if (!s.getRange(r, 11).getValue()) s.getRange(r, 1, 1, row.length).setValues([row]);   // baris lama tanpa Data
      return out_({ ok: true, rev: rev });
    }
    if (d.action === 'item') {
      upsert_(sh_(ss, 'bahan'), d.id, [d.id, d.name, d.unit, d.min, d.del ? 'PADAM' : 'OK', data, rev]);
      return out_({ ok: true, rev: rev });
    }
    if (d.action === 'settings') {
      upsert_(sh_(ss, 'set'), 'main', ['main', d.upd, data, rev]);
      return out_({ ok: true, rev: rev });
    }
    if (d.action === 'cashcount') {
      upsert_(sh_(ss, 'cash'), d.date, [d.date, d.float, d.cashSales, d.modalT, d.belT, d.expected, d.counted, d.diff,
        d.note || '', d.cashier || '', data, rev]);
      return out_({ ok: true, rev: rev });
    }

    // ---- order ----
    const so = sh_(ss, 'orders');
    const si = sh_(ss, 'items');

    if (d.action === 'void') {
      const r = findRow_(so, d.id);
      if (r) {
        so.getRange(r, 14).setValue('BATAL');
        so.getRange(r, 19).setValue(rev);
      }
      const last = si.getLastRow();
      if (last > 1) {
        const ids = si.getRange(2, 1, last - 1, 1).getValues();
        ids.forEach(function (row, i) {
          if (String(row[0]) === String(d.id)) si.getRange(i + 2, 11).setValue('BATAL');
        });
      }
      return out_({ ok: true, rev: rev });
    }

    const existing = findRow_(so, d.id);
    if (existing) {                       // elak duplikasi; isi Data jika baris ini dibuat versi lama
      if (!so.getRange(existing, 18).getValue()) {
        so.getRange(existing, 18, 1, 2).setValues([[data, rev]]);
      }
      return out_({ ok: true, dup: true, rev: rev });
    }

    so.appendRow([d.id, orderNo, d.date, d.time, d.channel, d.ref, itemText_(d.items),
      d.subtotal, d.discount, d.total, d.pay, d.tendered, d.change, d.status || 'OK',
      d.cashier || '', d.delivery || 0, d.discReason || '', data, rev]);

    const rows = d.items.map(function (i) {
      return [d.id, orderNo, d.date, d.time, d.channel, i.name, i.price, i.qty, i.price * i.qty, i.note, d.status || 'OK'];
    });
    if (rows.length) si.getRange(si.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);

    try { ensureDayRow_(ss, ensureSummary_(ss), d.date); } catch (err) { /* ringkasan tidak boleh gagalkan order */ }
    return out_({ ok: true, rev: rev });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- penyegerakan ----------

function props_() { return PropertiesService.getScriptProperties(); }
function curRev_() { return parseInt(props_().getProperty('REV') || '0', 10); }
function nextRev_() { const r = curRev_() + 1; props_().setProperty('REV', String(r)); return r; }
function epoch_() { return props_().getProperty('EPOCH') || 'e0'; }

// Pulangkan baris yang Rev-nya lebih besar daripada "since" (dibaca daripada hujung tab sahaja).
function changed_(ss, name, dataCol, revCol, since, limit) {
  const s = ss.getSheetByName(name);
  if (!s) return [];
  const last = s.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - limit + 1);
  const n = last - start + 1;
  const revs = s.getRange(start, revCol, n, 1).getValues();
  const idx = [];
  revs.forEach(function (r, i) { if (Number(r[0]) > since) idx.push(i); });
  if (!idx.length) return [];
  const lo = idx[0], hi = idx[idx.length - 1];
  const block = s.getRange(start + lo, 1, hi - lo + 1, revCol).getValues();
  const res = [];
  idx.forEach(function (i) {
    const row = block[i - lo];
    try { res.push({ o: JSON.parse(row[dataCol - 1]), row: row }); } catch (e) { /* baris lama tanpa Data */ }
  });
  return res;
}

function pull_(ss, since) {
  const cur = curRev_();
  const res = { ok: true, rev: cur, epoch: epoch_() };
  if (cur <= since) { res.same = true; return res; }
  res.orders = changed_(ss, 'Orders', 18, 19, since, 5000).map(function (x) { x.o.status = x.row[13]; return x.o; });
  res.bookings = changed_(ss, 'Tempahan', 14, 15, since, 2000).map(function (x) { return x.o; });
  res.expenses = changed_(ss, 'Belanja', 11, 12, since, 5000).map(function (x) { return x.o; });
  res.moves = changed_(ss, 'Stok Log', 11, 12, since, 10000).map(function (x) { return x.o; });
  res.items = changed_(ss, 'Bahan', 6, 7, since, 1000).map(function (x) { x.o.del = x.row[4] === 'PADAM'; return x.o; });
  res.cash = changed_(ss, 'Kiraan Tunai', 11, 12, since, 2000).map(function (x) { return x.o; });
  const st = changed_(ss, 'Tetapan', 3, 4, since, 10).map(function (x) { return x.o; });
  res.settings = st.length ? st[st.length - 1] : null;
  return res;
}

// Padam semua data jualan/belanja/stok (menu dan senarai bahan kekal). Semua peranti akan kosongkan salinan tempatan.
function kosongkan_(ss) {
  ['orders', 'items', 'book', 'exp', 'log', 'cash', 'stock'].forEach(function (k) {
    const s = ss.getSheetByName(SPEC[k][0]);
    if (!s) return;
    const last = s.getLastRow();
    if (last >= 2) s.getRange(2, 1, last - 1, s.getMaxColumns()).clearContent();
  });
  props_().setProperty('EPOCH', 'e' + new Date().getTime());
  buatSemulaRingkasan();
}

/** Jalankan daripada editor Apps Script untuk mengosongkan semua data ujian pada semua peranti. */
function kosongkanData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { kosongkan_(SpreadsheetApp.getActiveSpreadsheet()); } finally { lock.releaseLock(); }
}

/** Jalankan SEKALI daripada editor Apps Script untuk bina semua tab dan ringkasan. */
function buatRingkasan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SPEC).forEach(function (k) { sh_(ss, k); });
  ensureSummary_(ss);
}

/** Padam dan bina semula kedua-dua tab ringkasan (data lain tidak disentuh). */
function buatSemulaRingkasan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [HARIAN, ITEMSUM].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s) ss.deleteSheet(s);
  });
  buatRingkasan();
}

// ---------- pembantu ----------

function data_(d) {
  const o = {};
  Object.keys(d).forEach(function (k) { if (k !== 'action' && k !== 'key') o[k] = d[k]; });
  return JSON.stringify(o);
}

function itemText_(items) {
  return items.map(function (i) {
    return i.qty + 'x ' + i.name + (i.note ? ' (' + i.note + ')' : '');
  }).join('; ');
}

function sheet_(ss, name, headers, textCols) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(textCols || 'C:D').setNumberFormat('@');   // tarikh dan masa sebagai teks
    s.appendRow(headers);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    const di = headers.indexOf('Data');
    if (di >= 0) s.hideColumns(di + 1);
  }
  return s;
}

// Tambah lajur baru pada tab yang dibuat dengan versi lama.
function fixHeaders_(s, headers) {
  if (s.getRange(1, headers.length).getValue() === headers[headers.length - 1]) return;
  s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  const di = headers.indexOf('Data');
  if (di >= 0) s.hideColumns(di + 1);
}

function sh_(ss, key) {
  const sp = SPEC[key];
  const s = sheet_(ss, sp[0], sp[1], sp[2]);
  fixHeaders_(s, sp[1]);
  return s;
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

// Kemas kini baris jika ID wujud, jika tidak tambah baris baru (kunci dalam lajur A).
function upsert_(sheet, id, row) {
  const r = findRow_(sheet, id);
  if (r) sheet.getRange(r, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}

function norm_(v, ss) {
  if (v instanceof Date) return Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

// Pastikan lajur Tarikh/Masa ialah teks supaya formula ringkasan sentiasa padan.
function fixDates_(ss) {
  ['Orders', 'Items'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (!s) return;
    const last = s.getLastRow();
    let dates = null, times = null;
    if (last >= 2) {
      dates = s.getRange(2, 3, last - 1, 1).getValues();
      times = s.getRange(2, 4, last - 1, 1).getDisplayValues();
    }
    s.getRange('C:D').setNumberFormat('@');
    if (last >= 2) {
      s.getRange(2, 3, last - 1, 1).setValues(dates.map(function (x) { return [norm_(x[0], ss)]; }));
      s.getRange(2, 4, last - 1, 1).setValues(times.map(function (x) {
        const m = String(x[0]).match(/^(\d{1,2}):(\d{2})/);
        return [m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(x[0])];
      }));
    }
  });
}

function ensureSummary_(ss) {
  let h = ss.getSheetByName(HARIAN);
  if (h) { ensureItemSheet_(ss); return h; }

  fixDates_(ss);
  sh_(ss, 'exp');   // diperlukan oleh formula belanja
  h = ss.insertSheet(HARIAN, 0);
  const hd = ['Tarikh', 'Bil. Order', 'Jualan (RM)', 'Tunai (RM)', 'QR / DuitNow (RM)', 'Kad (RM)', 'Platform (RM)', 'Diskaun (RM)', 'Purata/Order (RM)', 'Caj Hantar (RM)', 'Belanja (RM)', 'Untung (RM)'];
  h.getRange('A:A').setNumberFormat('@');
  h.getRange(1, 1, 1, hd.length).setValues([hd]).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  h.setFrozenRows(1);
  h.setColumnWidth(1, 110);
  h.setColumnWidths(2, 11, 125);

  ensureItemSheet_(ss);

  // masukkan semua tarikh yang sudah ada dalam Orders dan Belanja
  const seen = {};
  [['Orders', 3], ['Belanja', 2]].forEach(function (p) {
    const s = ss.getSheetByName(p[0]);
    if (s && s.getLastRow() >= 2) {
      s.getRange(2, p[1], s.getLastRow() - 1, 1).getValues().forEach(function (x) {
        const d = norm_(x[0], ss);
        if (d) seen[d] = 1;
      });
    }
  });
  Object.keys(seen).sort().forEach(function (d) { ensureDayRow_(ss, h, d); });
  return h;
}

function ensureDayRow_(ss, h, date) {
  const last = h.getLastRow();
  if (last >= 2) {
    const dates = h.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (norm_(dates[i][0], ss) === String(date)) return;
    }
  }
  h.insertRowBefore(2);
  h.getRange(2, 1, 1, 12).clearFormat();
  h.getRange(2, 1).setNumberFormat('@').setValue(String(date));

  const base = '(Orders!$C$2:$C$20000=$A2)*(Orders!$N$2:$N$20000<>"BATAL")';
  const f = function (extra) { return '=SUMPRODUCT(' + base + extra + ')'; };
  const J = '*Orders!$J$2:$J$20000';
  const pay = function (p) { return '*(Orders!$K$2:$K$20000="' + p + '")' + J; };
  const belanja = '=SUMPRODUCT((Belanja!$B$2:$B$20000=$A2)*(Belanja!$D$2:$D$20000="Belanja")*(Belanja!$J$2:$J$20000<>"DIPADAM")*Belanja!$G$2:$G$20000)';
  h.getRange(2, 2, 1, 11).setFormulas([[
    f(''),
    f(J),
    f(pay('Tunai')),
    f(pay('QR / DuitNow')),
    f(pay('Kad')),
    f(pay('Platform')),
    f('*Orders!$I$2:$I$20000'),
    '=IFERROR(C2/B2,0)',
    f('*Orders!$P$2:$P$20000'),
    belanja,
    '=C2-K2'
  ]]);
  h.getRange(2, 3, 1, 10).setNumberFormat('#,##0.00');
}

function ensureItemSheet_(ss) {
  let s = ss.getSheetByName(ITEMSUM);
  if (s) return s;
  s = ss.insertSheet(ITEMSUM, 1);

  s.getRange('A1').setValue('Jualan ikut item (semua hari)').setFontWeight('bold');
  s.getRange('A2').setFormula(`=IFERROR(QUERY(Items!A:K,"select F, sum(H), sum(I) where K<>'BATAL' and F<>'' group by F order by sum(H) desc label F 'Item', sum(H) 'Kuantiti', sum(I) 'Jumlah (RM)'",1),"")`);

  s.getRange('E1').setValue('Jualan item ikut tarikh').setFontWeight('bold');
  s.getRange('E2').setValue('Tarikh (yyyy-mm-dd):');
  s.getRange('F2').setFormula('=TEXT(TODAY(),"yyyy-mm-dd")');
  s.getRange('E4').setFormula(`=IFERROR(QUERY(Items!A:K,"select F, sum(H), sum(I) where K<>'BATAL' and C='"&TEXT($F$2,"yyyy-mm-dd")&"' group by F order by sum(H) desc label F 'Item', sum(H) 'Kuantiti', sum(I) 'Jumlah (RM)'",1),"")`);

  s.getRange('I1').setValue('Jualan ikut saluran (semua hari)').setFontWeight('bold');
  s.getRange('I2').setFormula(`=IFERROR(QUERY(Orders!A:N,"select E, count(A), sum(J) where N<>'BATAL' and A<>'' group by E order by sum(J) desc label E 'Saluran', count(A) 'Bil. Order', sum(J) 'Jumlah (RM)'",1),"")`);

  s.getRange('M1').setValue('Jualan ikut cashier (semua hari)').setFontWeight('bold');
  s.getRange('M2').setFormula(`=IFERROR(QUERY(Orders!A:Q,"select O, count(A), sum(J) where N<>'BATAL' and A<>'' and O<>'' group by O order by sum(J) desc label O 'Cashier', count(A) 'Bil. Order', sum(J) 'Jumlah (RM)'",1),"")`);

  s.setColumnWidth(1, 160); s.setColumnWidth(5, 160); s.setColumnWidth(9, 160); s.setColumnWidth(13, 160);
  s.setColumnWidths(2, 2, 100); s.setColumnWidths(6, 2, 100); s.setColumnWidths(10, 2, 100); s.setColumnWidths(14, 2, 100);
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
