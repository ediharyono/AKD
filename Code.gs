/** Pengumpulan Soal UTS - UII
 *  Database : spreadsheet "Soal UTS Ganjil 26-26", sheet gid 318214032
 *  Link soal: kolom S ("Link Soal") pada sheet tersebut
 */
const CFG = {
  SS_ID: '1f7BIiHQhEMv2cdF7BgbO4rlAXpe6nbDO8sd697bTEKU',
  GID: 318214032,
  FOLDER_ID: '17slqoroK4qsIth_e707ArnnUX6Hv56Xg',
  DOMAIN: 'uii.ac.id',
  LINK_HEADER: 'Link Soal',   // kolom S
  TIME_HEADER: 'Waktu Kumpul', // dibuat otomatis di kolom setelah kolom terakhir jika belum ada
  TIME_FMT: 'dd/MM/yyyy HH:mm',
  ATT_HEADER: 'Lampiran Soal', // link lampiran (opsional), dibuat otomatis jika belum ada
  // Kolom sifat ujian (dibuat otomatis di sheet jadwal jika belum ada)
  PROP_COLS: ['Soal Dikembalikan', 'Sifat Ujian', 'Cetak Soal Take Home', 'Pengumpulan Jawaban',
    'Catatan Take Home', 'Kertas Folio', 'Kertas Buram', 'Catatan'],
  SIFAT: ['Close Book', 'Open Book', 'Take Home'],                       // sesuaikan pilihan
  KUMPUL: ['Dikumpulkan online (LMS/email)', 'Dikumpulkan langsung'],   // sesuaikan pilihan
  ADMIN_GID: 187256259,       // sheet "admin" (kolom A: EMAIL, kolom B: NAMA)
  LOG: 'log_upload',
  USER_SHEET: 'user_dosen',   // hasil generate user dosen
  MAX_MB: 15,
  MIME: {
    pdf: 'application/pdf',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pengumpulan Soal UTS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Jalankan SEKALI dari editor: membuat sheet admin & log, memberi izin akses. */
function setup() {
  jadwal_();
  adminSheet_();
  logSheet_();
  DriveApp.getFolderById(CFG.FOLDER_ID).getName();
}

// ---------- helper ----------
const ss_ = () => SpreadsheetApp.openById(CFG.SS_ID);

function read_(sh) {
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0].map(s => String(s).trim());
  const rows = v.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => o[k] = r[j]);
    return o;
  });
  return { sh, h, rows };
}

function jadwal_() {
  const sh = ss_().getSheets().find(s => s.getSheetId() === CFG.GID);
  if (!sh) throw new Error('Sheet dengan gid ' + CFG.GID + ' tidak ditemukan.');
  let t = read_(sh);
  if (t.h.indexOf(CFG.LINK_HEADER) < 0) throw new Error('Kolom "' + CFG.LINK_HEADER + '" tidak ditemukan di baris 1.');
  let baru = false;
  [CFG.TIME_HEADER, CFG.ATT_HEADER].concat(CFG.PROP_COLS).forEach(k => {
    if (t.h.indexOf(k) >= 0) return;
    const c = sh.getLastColumn() + 1;
    if (c > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
    sh.getRange(1, c).setValue(k).setFontWeight('bold');
    if (k === CFG.TIME_HEADER) sh.getRange(2, c, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat(CFG.TIME_FMT);
    baru = true;
  });
  if (baru) t = read_(sh);
  t.c = k => t.h.indexOf(k) + 1;
  t.linkCol = t.h.indexOf(CFG.LINK_HEADER) + 1;
  t.timeCol = t.h.indexOf(CFG.TIME_HEADER) + 1;
  t.flagH = t.h.find(k => /keterangan/i.test(k)) || '';   // kolom N "Keterangan flag kelas"
  return t;
}

function adminSheet_() {
  const sh = ss_().getSheets().find(s => s.getSheetId() === CFG.ADMIN_GID);
  if (!sh) throw new Error('Sheet admin (gid ' + CFG.ADMIN_GID + ') tidak ditemukan.');
  const t = read_(sh);
  if (t.h.indexOf('EMAIL') < 0) throw new Error('Sheet admin harus punya header EMAIL di baris 1.');
  t.rows = t.rows.filter(r => String(r.EMAIL).trim());
  return t;
}

function logSheet_() {
  let sh = ss_().getSheetByName(CFG.LOG);
  if (!sh) {
    sh = ss_().insertSheet(CFG.LOG);
    sh.getRange(1, 1, 1, 7).setValues([['WAKTU', 'EMAIL', 'AKSI', 'NO_KELAS', 'NAMA_FILE', 'LINK', 'UKURAN_KB']]).setFontWeight('bold');
  }
  return sh;
}

function me_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const m = email.match(/^([^@]+)@(.+)$/);
  const admin = !!email && adminSheet_().rows.some(r => String(r.EMAIL).trim().toLowerCase() === email);
  return { email, kd: m && m[2] === CFG.DOMAIN ? m[1] : '', admin };
}

const isOpen_ = () => PropertiesService.getScriptProperties().getProperty('OPEN') !== '0';
function needAdmin_() { if (!me_().admin) throw new Error('Hanya admin yang boleh melakukan ini.'); }

function idFromLink_(url) {
  const m = String(url || '').match(/\/d\/([-\w]{20,})|[?&]id=([-\w]{20,})/);
  return m ? (m[1] || m[2]) : '';
}

/** Hapus file hanya jika ada di folder soal & tidak lagi dipakai baris mana pun. */
function trashUnused_(ids) {
  const t = jadwal_();
  const used = t.rows.map(r => idFromLink_(r[CFG.LINK_HEADER])).concat(t.rows.map(r => idFromLink_(r[CFG.ATT_HEADER])));
  ids.filter(id => id && used.indexOf(id) < 0).forEach(id => {
    try {
      const f = DriveApp.getFileById(id), ps = f.getParents();
      while (ps.hasNext()) if (ps.next().getId() === CFG.FOLDER_ID) { f.setTrashed(true); break; }
    } catch (e) { }
  });
}

/** 'I' = Kelas Gabungan - Induk, 'P' = Kelas Gabungan - Peserta, '' = kelas biasa. */
function flag_(r, t) {
  const v = String(r[t.flagH] || '').toLowerCase();
  if (v.indexOf('gabungan') < 0) return '';
  return v.indexOf('induk') >= 0 ? 'I' : v.indexOf('peserta') >= 0 ? 'P' : '';
}

/**
 * Satu entri per baris tabel. Kelas gabungan (Induk + Peserta) dengan KD_DOSEN, hari, sesi, dan ruang yang sama
 * digabung menjadi satu entri: prim = kelas induk, sub = kelas lain dalam gabungan, rows = semua baris sheet.
 * Gabungan tanpa induk ditampilkan sebagai kelas biasa agar tetap bisa diunggah.
 */
function groups_(t) {
  const out = [], map = {};
  t.rows.filter(r => r.NO).forEach(r => {
    if (!flag_(r, t)) { out.push({ prim: r, rows: [r], sub: [] }); return; }
    const k = [r.KD_DOSEN, r['HARI UJIAN'], r['SESI UJIAN'], r['RUANG UJIAN']].join('|');
    (map[k] = map[k] || []).push(r);
  });
  Object.keys(map).forEach(k => {
    const rows = map[k], ind = rows.filter(r => flag_(r, t) === 'I');
    if (!ind.length) { rows.forEach(r => out.push({ prim: r, rows: [r], sub: [] })); return; }
    out.push({ prim: ind[0], rows: rows, sub: rows.filter(r => r !== ind[0]) });
  });
  return out.sort((a, b) => Number(a.prim.NO) - Number(b.prim.NO));
}

/** Daftar user dosen dari data mata kuliah: email = KD_DOSEN (NIK) @ domain. */
function users_() {
  const u = {};
  groups_(jadwal_()).map(g => g.prim).filter(r => String(r.KD_DOSEN).trim()).forEach(r => {
    const k = String(r.KD_DOSEN).trim();
    const x = u[k] || (u[k] = { nik: k, nama: r.NAMA_DOSEN, email: k + '@' + CFG.DOMAIN, kelas: 0, terkumpul: 0 });
    x.kelas++;
    if (r[CFG.LINK_HEADER]) x.terkumpul++;
  });
  return Object.keys(u).map(k => u[k]).sort((a, b) => a.nama.localeCompare(b.nama));
}

/** Tulis daftar user dosen ke sheet "user_dosen" (admin). */
function generateUsers() {
  needAdmin_();
  const us = users_();
  const sh = ss_().getSheetByName(CFG.USER_SHEET) || ss_().insertSheet(CFG.USER_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['NIK', 'NAMA_DOSEN', 'EMAIL', 'JUMLAH_KELAS']]).setFontWeight('bold');
  if (us.length) {
    sh.getRange(2, 1, us.length, 3).setNumberFormat('@');
    sh.getRange(2, 1, us.length, 4).setValues(us.map(x => [x.nik, x.nama, x.email, x.kelas]));
  }
  return getData();
}

// ---------- API untuk client ----------
function getData() {
  const me = me_();
  if (!me.email) throw new Error('Email tidak terbaca. Login dengan akun @' + CFG.DOMAIN + '.');
  if (!me.admin && !me.kd) return { me, error: 'Akun ' + me.email + ' bukan akun @' + CFG.DOMAIN + '.' };
  const us = users_();
  if (!me.admin && !us.some(u => u.nik === me.kd))
    return { me, error: 'Email ' + me.email + ' belum terdaftar sebagai dosen pengampu. Format email harus NIK@' + CFG.DOMAIN + ' sesuai KD_DOSEN pada jadwal.' };
  const t = jadwal_();
  const items = groups_(t)
    .filter(g => me.admin || g.prim.KD_DOSEN === me.kd)
    .map(g => {
      const r = g.prim;
      return {
        no: r.NO, mk: r.KD_MATAKULIAH, nama: r.NAMA_MATAKULIAH, kelas: r.KELAS,
        hari: r['HARI UJIAN'], sesi: r['SESI UJIAN'], ruang: r['RUANG UJIAN'],
        kd: r.KD_DOSEN, dosen: r.NAMA_DOSEN, prodi: r.Prodi, link: r[CFG.LINK_HEADER], waktu: r[CFG.TIME_HEADER], att: r[CFG.ATT_HEADER],
        sub: g.sub.map(x => ({ mk: x.KD_MATAKULIAH, nama: x.NAMA_MATAKULIAH, kelas: x.KELAS, mhs: x.MHS, flag: flag_(x, t) })),
        prop: CFG.PROP_COLS.reduce((o, k) => (o[k] = r[k] || '', o), {})
      };
    });
  const out = { me, open: isOpen_(), items, max: CFG.MAX_MB, opt: { sifat: CFG.SIFAT, kumpul: CFG.KUMPUL } };
  if (me.admin) out.users = us;
  if (me.admin) out.admins = adminSheet_().rows.map(r => ({ email: r.EMAIL, nama: r.NAMA }));
  return out;
}

function checkProps_(p) {
  const YN = ['Ya', 'Tidak'], P = {};
  const must = (k, list) => {
    const v = String(p[k] || '').trim();
    if (list.indexOf(v) < 0) throw new Error('Lengkapi isian: ' + k + '.');
    P[k] = v;
  };
  must('Soal Dikembalikan', YN);
  must('Sifat Ujian', CFG.SIFAT);
  if (P['Sifat Ujian'] === 'Take Home') {
    must('Cetak Soal Take Home', YN);
    must('Pengumpulan Jawaban', CFG.KUMPUL);
    P['Catatan Take Home'] = String(p['Catatan Take Home'] || '').trim();
  } else {
    P['Cetak Soal Take Home'] = P['Pengumpulan Jawaban'] = P['Catatan Take Home'] = '';
  }
  must('Kertas Folio', YN);
  must('Kertas Buram', YN);
  P['Catatan'] = String(p['Catatan'] || '').trim();
  return P;
}

/** Simpan berkas ke folder Drive; mengembalikan {file, bytes, fname}. */
function saveFile_(name, b64, jd, now, tag) {
  const f0 = jd[0];
  const ext = (String(name).split('.').pop() || '').toLowerCase();
  if (!CFG.MIME[ext]) throw new Error('Format harus PDF, Word (doc/docx), atau Excel (xls/xlsx).');
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length > CFG.MAX_MB * 1048576) throw new Error('Ukuran maksimal ' + CFG.MAX_MB + ' MB.');
  const stamp = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd-HHmm');
  const fname = [f0.KD_MATAKULIAH, f0.NAMA_MATAKULIAH.replace(/[^\w]+/g, '_'),
    'Kls' + jd.map(r => r.KELAS).join(''), f0.KD_DOSEN, tag, stamp].filter(String).join('_') + '.' + ext;
  const file = DriveApp.getFolderById(CFG.FOLDER_ID).createFile(Utilities.newBlob(bytes, CFG.MIME[ext], fname));
  jd.map(r => r.KD_DOSEN).filter((v, i, x) => v && x.indexOf(v) === i).forEach(k => {
    try { file.addViewer(k + '@' + CFG.DOMAIN); } catch (e) { }
  });
  return { file, bytes, fname };
}

/** p = {nos:[], prop:{...}, name?, b64?}. Berkas boleh kosong bila semua kelas sudah punya soal (hanya ubah sifat ujian). */
function upload(p) {
  const me = me_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!me.admin && !isOpen_()) throw new Error('Pengumpulan soal sedang ditutup.');
    const P = checkProps_(p.prop || {});
    const nos = p.nos.map(String), t = jadwal_(), G = groups_(t);
    const sel = G.filter(g => g.rows.some(r => nos.indexOf(r.NO) >= 0));
    if (!sel.length || nos.some(n => !G.some(g => g.rows.some(r => r.NO === n)))) throw new Error('Data kelas tidak ditemukan.');
    const jd = sel.map(g => g.prim);                                   // kelas induk / mandiri
    const semua = [].concat.apply([], sel.map(g => g.rows));           // termasuk peserta gabungan
    jd.forEach(r => {
      if (!me.admin && r.KD_DOSEN !== me.kd) throw new Error('Anda tidak berhak mengunggah untuk kelas ' + r.KELAS + '.');
    });

    if (jd.some(r => r.KD_DOSEN !== jd[0].KD_DOSEN)) throw new Error('Kelas yang dipilih harus milik dosen pengampu yang sama.');
    const now = new Date();
    if (!p.b64 && jd.some(r => !r[CFG.LINK_HEADER])) throw new Error('Pilih berkas soal untuk kelas yang belum memiliki soal.');
    const main = p.b64 ? saveFile_(p.name, p.b64, jd, now, '') : null;
    const att = p.attB64 ? saveFile_(p.attName, p.attB64, jd, now, 'Lampiran') : null;
    const file = main && main.file, bytes = main && main.bytes, fname = main ? main.fname : '';

    const oldIds = [];
    semua.forEach(r => {
      if (file) {
        oldIds.push(idFromLink_(r[CFG.LINK_HEADER]));
        t.sh.getRange(r._row, t.linkCol).setValue(file.getUrl());
        t.sh.getRange(r._row, t.timeCol).setNumberFormat(CFG.TIME_FMT).setValue(now);
      }
      if (att || p.hapusAtt) {
        oldIds.push(idFromLink_(r[CFG.ATT_HEADER]));
        t.sh.getRange(r._row, t.c(CFG.ATT_HEADER)).setValue(att ? att.file.getUrl() : '');
      }
      CFG.PROP_COLS.forEach(k => t.sh.getRange(r._row, t.c(k)).setValue(P[k]));
    });
    SpreadsheetApp.flush();
    trashUnused_(oldIds);
    logSheet_().appendRow([Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'), me.email,
      file ? 'UPLOAD' : 'UBAH_SIFAT', semua.map(r => r.NO).join(','), fname + (att ? ' + lampiran: ' + att.fname : ''), file ? file.getUrl() : '', bytes ? Math.round(bytes.length / 1024) : '']);
    return getData();
  } finally { lock.releaseLock(); }
}

function hapus(no) {
  const me = me_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!me.admin && !isOpen_()) throw new Error('Pengumpulan soal sedang ditutup.');
    const t = jadwal_(), g = groups_(t).find(x => x.rows.some(r => r.NO === String(no))), r = g && g.prim;
    if (!r || (!me.admin && r.KD_DOSEN !== me.kd)) throw new Error('Tidak berhak menghapus.');
    g.rows.forEach(x => {
      t.sh.getRange(x._row, t.linkCol).clearContent();
      t.sh.getRange(x._row, t.timeCol).clearContent();
      t.sh.getRange(x._row, t.c(CFG.ATT_HEADER)).clearContent();
    });
    SpreadsheetApp.flush();
    trashUnused_([idFromLink_(r[CFG.LINK_HEADER]), idFromLink_(r[CFG.ATT_HEADER])]);
    logSheet_().appendRow([Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'), me.email,
      'HAPUS', no, '', r[CFG.LINK_HEADER], '']);
    return getData();
  } finally { lock.releaseLock(); }
}

// ---------- admin ----------
function setOpen(open) {
  needAdmin_();
  PropertiesService.getScriptProperties().setProperty('OPEN', open ? '1' : '0');
  return getData();
}

function addAdmin(email, nama) {
  needAdmin_();
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Format email tidak valid.');
  const t = adminSheet_();
  if (!t.rows.some(r => String(r.EMAIL).trim().toLowerCase() === email)) t.sh.appendRow([email, nama || '']);
  return getData();
}

function removeAdmin(email) {
  needAdmin_();
  email = String(email).toLowerCase();
  if (email === me_().email) throw new Error('Anda tidak bisa menghapus akun Anda sendiri dari daftar admin.');
  const t = adminSheet_(), r = t.rows.find(x => String(x.EMAIL).trim().toLowerCase() === email);
  if (r) t.sh.deleteRow(r._row);
  return getData();
}
