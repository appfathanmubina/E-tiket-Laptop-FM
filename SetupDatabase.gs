/** Jalankan setupDatabase() satu kali dari Apps Script. */
const APP_CONFIG = {
  SPREADSHEET_ID: '',
  SHEETS: {
    Pengaturan:['Kunci','Nilai','Keterangan'],
    Santri:['ID_Santri','NIS','Nama','Kelas','Asrama','Status','No_HP_Wali','Foto_File_ID','Catatan','Dibuat_Pada','Diubah_Pada'],
    Laptop:['ID_Laptop','Kode_Aset','Merek_Model','Nomor_Seri','Kelengkapan','Kondisi','Status','Catatan','Dibuat_Pada','Diubah_Pada'],
    Penitipan:['ID_Transaksi','Tanggal_Masuk','ID_Santri','Nama_Santri','ID_Laptop','Kode_Aset','Kelengkapan','Kondisi_Masuk','Petugas','Status','Tanggal_Keluar','Catatan','Dibuat_Pada'],
    Pemakaian:['ID_Pemakaian','Tanggal_Pinjam','Batas_Kembali','Tanggal_Kembali','ID_Santri','Nama_Santri','ID_Laptop','Kode_Aset','Tujuan','Pemberi_Izin','Petugas_Keluar','Petugas_Kembali','Kondisi_Keluar','Kondisi_Kembali','Status','Foto_File_ID','Catatan','Dibuat_Pada'],
    Pelanggaran:['ID_Pelanggaran','Tanggal','ID_Santri','Nama_Santri','Kategori','Tingkat','Lokasi','Uraian','Bukti_URL','Foto_Bukti_File_ID','Tindak_Lanjut','Status_Penanganan','Pencatat','Dibuat_Pada','Diubah_Pada'],
    Pengguna:['ID_Pengguna','Nama','Username','Password_Hash','Peran','Status','Dibuat_Pada','Login_Terakhir'],
    LogAktivitas:['ID_Log','Waktu','Username','Aksi','Modul','ID_Referensi','Detail']
  }
};
function setupDatabase(){
  const props=PropertiesService.getScriptProperties(); let ss;
  if(APP_CONFIG.SPREADSHEET_ID) ss=SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID);
  else if(props.getProperty('ETIKET_SPREADSHEET_ID')) ss=SpreadsheetApp.openById(props.getProperty('ETIKET_SPREADSHEET_ID'));
  else {ss=SpreadsheetApp.create('Database E-Tiket Laptop FM');props.setProperty('ETIKET_SPREADSHEET_ID',ss.getId());}
  Object.keys(APP_CONFIG.SHEETS).forEach(n=>{
    let sh=ss.getSheetByName(n);if(!sh)sh=ss.insertSheet(n);
    const h=APP_CONFIG.SHEETS[n], old=sh.getLastColumn()?sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]:[];
    if(!old.length||old.every(x=>!x))sh.getRange(1,1,1,h.length).setValues([h]);
    else {const missing=h.filter(x=>!old.includes(x));if(missing.length)sh.getRange(1,old.length+1,1,missing.length).setValues([missing]);}
    sh.setFrozenRows(1);sh.getRange(1,1,1,sh.getLastColumn()).setBackground('#5b21b6').setFontColor('#ffffff').setFontWeight('bold');
  });
  const cfg=ss.getSheetByName('Pengaturan');const present=cfg.getLastRow()>1?cfg.getRange(2,1,cfg.getLastRow()-1,1).getDisplayValues().flat():[];
  [['Nama_Aplikasi','E-Tiket Laptop FM','Nama aplikasi'],['Nama_Pesantren','Pesantren Fathan Mubina','Nama lembaga'],['Zona_Waktu','Asia/Jakarta','Zona waktu']].forEach(r=>{if(!present.includes(r[0]))cfg.appendRow(r);});
  const users=ss.getSheetByName('Pengguna');
  if(users.getLastRow()<2){
    const pwd='123456';
    users.appendRow(['USR-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Administrator','admin',hashPassword_(pwd),'Administrator','Aktif',new Date(),'']);
    props.setProperty('ETIKET_INITIAL_USERNAME','admin');props.setProperty('ETIKET_INITIAL_PASSWORD',pwd);
  }
  props.setProperty('ETIKET_SPREADSHEET_ID',ss.getId());
  const result={success:true,spreadsheetUrl:ss.getUrl(),spreadsheetId:ss.getId(),initialUsername:props.getProperty('ETIKET_INITIAL_USERNAME')||'',initialPassword:props.getProperty('ETIKET_INITIAL_PASSWORD')||'',message:'Simpan kredensial awal dengan aman.'};
  Logger.log('HASIL SETUP: '+JSON.stringify(result));
  return result;
}
function getDatabase_(){const id=PropertiesService.getScriptProperties().getProperty('ETIKET_SPREADSHEET_ID')||APP_CONFIG.SPREADSHEET_ID;if(!id)throw new Error('Jalankan setupDatabase() terlebih dahulu.');return SpreadsheetApp.openById(id);}
function hashPassword_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}