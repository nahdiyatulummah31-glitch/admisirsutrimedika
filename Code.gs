/** RSU Trimedika Ketapang — Pendaftaran rencana Rawat Jalan dan Rawat Inap. */
const APP = {
  SHEET: 'DATA_RENCANA_PASIEN',
  TZ: 'Asia/Jakarta',
  POLI: ['Poli Penyakit Dalam','Poli Bedah Umum','Poli Bedah Mulut','Poli Obgyn/ Kebidanan','Poli Anak','Poli Paru','Poli Saraf/ Neuro'],
  DOKTER: [
    'dr. Sisilia Dewanti, Sp.PD',
    'dr. Jo Jonathan Jose Johan, Sp.PD',
    'dr. Eldaa Putik Bunga Melati, Sp.P',
    'dr. Deddy Setiawan Widjaja, Sp.B',
    'dr. Kartika Istiningtyas, Sp.A',
    'dr. Adia Triyarintana, Sp.OG',
    'dr. Arif Eko Wibowo, Sp.N',
    'drg. Aris Setyawan, MPH, M.Ked.Klin Sp.BMM TM-TMJ (K) SH'
  ],
  KELAS: ['VIP','Kelas 1','Kelas 2','Kelas 3'],
  BANGSAL: ['Cemara','Cendana','Gaharu','HCU','Isolasi','ICU'],
  HEADERS: [
    'ID Pendaftaran','Waktu Input','Jenis Layanan','No. Rekam Medis','Nama Pasien',
    'Poli Tujuan','Dokter','Kelas Rawat Inap','Bangsal Rawat Inap','Asuransi Pembiayaan','Nama Asuransi Lain','No. HP',
    'Tanggal Rencana','Diagnosis','Catatan'
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Rumah Sakit Umum Trimedika Ketapang | Pendaftaran Rencana Layanan')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

/** Jalankan satu kali setelah memasang kode. Tidak menghapus data yang sudah ada. */
function setupAplikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Hubungkan Apps Script ke Google Spreadsheet terlebih dahulu.');
  let sh = ss.getSheetByName(APP.SHEET);
  if (!sh) sh = ss.insertSheet(APP.SHEET);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,APP.HEADERS.length).setValues([APP.HEADERS]);
  else {
    let headers = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getDisplayValues()[0];
    const aliases = {
      'No. Rekam Medis':['No Rekam Medis'],
      'Asuransi Pembiayaan':['Status Asuransi'],
      'Tanggal Rencana':['Tanggal Kontrol','Tanggal Rencana Rawat Inap'],
      'Catatan':['Keterangan']
    };
    APP.HEADERS.forEach(h => {
      if (headers.indexOf(h) >= 0) return;
      const oldName=(aliases[h]||[]).find(a=>headers.indexOf(a)>=0);
      if(oldName){const col=headers.indexOf(oldName)+1;sh.getRange(1,col).setValue(h);headers[col-1]=h;return;}
      const col=sh.getLastColumn()+1;
      sh.getRange(1,col).setValue(h);headers.push(h);
    });
  }
  const headersNow = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const phoneCol = headersNow.indexOf('No. HP') + 1;
  if (phoneCol > 0 && sh.getMaxRows() > 1) sh.getRange(2,phoneCol,sh.getMaxRows()-1,1).setNumberFormat('@');
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#dbeafe').setWrap(true);
  return {ok:true,url:ss.getUrl(),message:'Sheet pendaftaran siap digunakan.'};
}

function getAppData(inputDate) {
  const sh = getSheet_();
  const date = clean_(inputDate,10) || Utilities.formatDate(new Date(),APP.TZ,'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal input pendaftaran tidak valid.');
  const registrations = readSheet_(sh).filter(r => normalizeInputDate_(r['Waktu Input']) === date);
  return {registrations:registrations, spreadsheetUrl:SpreadsheetApp.getActiveSpreadsheet().getUrl()};
}

/** Dashboard unit hanya membaca data yang dibuat dari menu Pendaftaran/Admin. */
function getDashboardData(unit, dateText) {
  if (!['Rawat Jalan','Rawat Inap'].includes(unit)) throw new Error('Unit dashboard tidak valid.');
  const date = clean_(dateText,10) || Utilities.formatDate(new Date(),APP.TZ,'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal dashboard tidak valid.');
  const rows = readSheet_(getSheet_()).filter(r =>
    r['Jenis Layanan'] === unit && normalizeDate_(r['Tanggal Rencana']) === date
  );
  if (unit === 'Rawat Jalan') {
    const doctorOrder = new Map(APP.DOKTER.map((name,index) => [name,index]));
    rows.sort((a,b) => (doctorOrder.get(a['Dokter']) ?? APP.DOKTER.length) - (doctorOrder.get(b['Dokter']) ?? APP.DOKTER.length));
  }
  return {unit:unit,date:date,total:rows.length,rows:rows.map(r => ({
    id:r['ID Pendaftaran'], noRekamMedis:r['No. Rekam Medis']||'', nama:r['Nama Pasien']||'',
    poli:r['Poli Tujuan']||'', dokter:r['Dokter']||'', kelas:r['Kelas Rawat Inap']||'', bangsal:r['Bangsal Rawat Inap']||'', asuransi:r['Asuransi Pembiayaan']||r['Status Asuransi']||'',
    asuransiLain:r['Nama Asuransi Lain']||'', noHp:displayPhone_(r['No. HP']),
    tanggal:normalizeDate_(r['Tanggal Rencana']), diagnosis:r['Diagnosis']||'',
    catatan:r['Catatan']||r['Keterangan']||''
  }))};
}

function saveAdmisi(data) {
  const d = data || {};
  const kind = clean_(d.jenisLayanan,30);
  const mrn = clean_(d.noRekamMedis,50);
  const name = clean_(d.namaPasien,120);
  const phone = cleanPhone_(d.noHp);
  const insurance = clean_(d.asuransi,50);
  const date = clean_(d.tanggal,20);
  const diagnosis = clean_(d.diagnosis,500);
  const note = clean_(d.catatan,1000);
  const poli = kind === 'Rawat Jalan' ? clean_(d.poliTujuan,100) : '';
  const dokter = clean_(d.dokter,150);
  const kelas = kind === 'Rawat Inap' ? clean_(d.kelasRawat,30) : '';
  const bangsal = kind === 'Rawat Inap' ? clean_(d.bangsalRawat,30) : '';
  if (!['Rawat Jalan','Rawat Inap'].includes(kind)) throw new Error('Pilih rencana Rawat Jalan atau Rawat Inap.');
  if (!mrn || !name || !phone || !insurance || !date || !diagnosis) throw new Error('Lengkapi nomor rekam medis, nama, nomor HP, asuransi, tanggal rencana, dan diagnosis.');
  if (kind === 'Rawat Jalan' && !APP.POLI.includes(poli)) throw new Error('Pilih poli tujuan dari daftar.');
  if (!APP.DOKTER.includes(dokter)) throw new Error('Pilih dokter dari daftar.');
  if (kind === 'Rawat Inap' && !APP.KELAS.includes(kelas)) throw new Error('Pilih kelas rawat inap.');
  if (kind === 'Rawat Inap' && !APP.BANGSAL.includes(bangsal)) throw new Error('Pilih bangsal rawat inap.');
  if (!['BPJS Kesehatan','Umum','Asuransi Lain'].includes(insurance)) throw new Error('Pilih asuransi pembiayaan yang tersedia.');
  const other = insurance === 'Asuransi Lain' ? clean_(d.namaAsuransiLain,100) : '';
  if (insurance === 'Asuransi Lain' && !other) throw new Error('Isi nama asuransi lain.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal rencana tidak valid.');

  const sh = getSheet_();
  const id = clean_(d.editId,100);
  const stamp = Utilities.formatDate(new Date(),APP.TZ,'yyyy-MM-dd HH:mm:ss');
  const rowObject = {
    'ID Pendaftaran':id || Utilities.getUuid(), 'Waktu Input':stamp, 'Jenis Layanan':kind,
    'No. Rekam Medis':mrn, 'Nama Pasien':name, 'Poli Tujuan':poli, 'Dokter':dokter, 'Kelas Rawat Inap':kelas, 'Bangsal Rawat Inap':bangsal,
    'Asuransi Pembiayaan':insurance, 'Nama Asuransi Lain':other, 'No. HP':phone,
    'Tanggal Rencana':date, 'Diagnosis':diagnosis, 'Catatan':note
  };
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  let targetRow;
  if (id) {
    const ids = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().flat() : [];
    const idx = ids.indexOf(id);
    if (idx < 0) throw new Error('Data pendaftaran yang akan diedit tidak ditemukan.');
    targetRow = idx + 2;
    const old = readRowObject_(sh,targetRow,headers);
    rowObject['Waktu Input'] = old['Waktu Input'] || stamp;
  } else targetRow = sh.getLastRow()+1;
  const row = headers.map(h => rowObject[h] !== undefined ? rowObject[h] : '');
  const phoneIndex = headers.indexOf('No. HP');
  if (phoneIndex >= 0) sh.getRange(targetRow,phoneIndex+1).setNumberFormat('@');
  sh.getRange(targetRow,1,1,headers.length).setNumberFormat('@').setValues([row]);
  return {ok:true,id:rowObject['ID Pendaftaran'],message:id?'Data pendaftaran berhasil diperbarui.':'Rencana berhasil didaftarkan.'};
}

function deleteAdmisi(id) {
  const sh = getSheet_();
  const count = Math.max(sh.getLastRow()-1,0);
  if (!count) throw new Error('Tidak ada data untuk dihapus.');
  const ids = sh.getRange(2,1,count,1).getDisplayValues().flat();
  const idx = ids.indexOf(String(id));
  if (idx < 0) throw new Error('Data tidak ditemukan.');
  sh.deleteRow(idx+2);
  return {ok:true};
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan. Buka Apps Script melalui Extensions > Apps Script dari Google Sheets.');
  const sh = ss.getSheetByName(APP.SHEET);
  if (!sh) throw new Error('Sheet DATA_RENCANA_PASIEN belum ada. Jalankan setupAplikasi() terlebih dahulu.');
  return sh;
}
function readSheet_(sh) {
  if (!sh || sh.getLastRow()<2) return [];
  const values=sh.getDataRange().getDisplayValues();
  const headers=values.shift();
  return values.filter(r=>r[0]).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});
}
function readRowObject_(sh,row,headers) {
  const values=sh.getRange(row,1,1,headers.length).getDisplayValues()[0];
  const o={};headers.forEach((h,i)=>o[h]=values[i]);return o;
}
function normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value,APP.TZ,'yyyy-MM-dd');
  const s=String(value||'').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  return m ? `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}` : s;
}
function normalizeInputDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value,APP.TZ,'yyyy-MM-dd');
  const s=String(value||'').trim();
  const iso=s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  const local=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  return local ? `${local[3]}-${String(local[2]).padStart(2,'0')}-${String(local[1]).padStart(2,'0')}` : '';
}
function clean_(v,max){return String(v==null?'':v).trim().replace(/[\u0000-\u001f]/g,' ').slice(0,max);}
function cleanPhone_(v){return String(v==null?'':v).trim().replace(/[^0-9+()\-\s]/g,'').slice(0,30);}
function displayPhone_(v){
  const s=String(v==null?'':v).trim();
  // Nomor lama yang pernah tersimpan sebagai angka kehilangan nol awal; tampilkan kembali sebagai nomor lokal.
  if (/^\d{9,14}$/.test(s) && s[0]!=='0' && !s.startsWith('62')) return '0'+s;
  return s;
}
