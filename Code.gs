/** E-Tiket Laptop FM — backend Google Apps Script */
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('E-Tiket Laptop FM').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function apiLogin(p){return api_(login_,p)}
function apiGetAppData(p){return api_(getAppData_,p)}
function apiSaveSantri(p){return api_(x=>saveSimple_('Santri',x,['NIS','Nama']),p)}
function apiGetFotoSantri(p){return api_(getFotoSantri_,p)}
function apiGetFotoBukti(p){return api_(getFotoBukti_,p)}
function apiSaveLaptop(p){return api_(saveLaptopAndDeposit_,p)}
function apiSavePenitipan(p){return api_(saveDeposit_,p)}
function apiCheckoutLaptop(p){return api_(checkout_,p)}
function apiReturnLaptop(p){return api_(returnLaptop_,p)}
function apiSavePelanggaran(p){return api_(saveViolation_,p)}
function apiGetReport(p){return api_(report_,p)}
function api_(fn,p){try{return {success:true,data:fn(p||{})}}catch(e){return {success:false,message:e.message||String(e)}}}
function login_(p){
 const u=String(p.username||'').trim(),pwd=String(p.password||'');if(!u||!pwd)throw Error('Masukkan username dan kata sandi.');
 const sh=getDatabase_().getSheetByName('Pengguna'),rows=objects_(sh),i=rows.findIndex(r=>String(r.Username).toLowerCase()===u.toLowerCase()&&r.Status==='Aktif'),user=rows[i];
 if(!user||user.Password_Hash!==hashPassword_(pwd))throw Error('Username atau kata sandi tidak sesuai.');
 const token=Utilities.getUuid()+Utilities.getUuid(),session={id:user.ID_Pengguna,username:user.Username,name:user.Nama,role:user.Peran};
 CacheService.getScriptCache().put('eticket_'+token,JSON.stringify(session),21600);sh.getRange(i+2,8).setValue(new Date());log_(u,'LOGIN','Sistem',user.ID_Pengguna,'Login berhasil');
 return {token:token,user:{name:user.Nama,username:user.Username,role:user.Peran}};
}
function session_(p){const s=p.token&&CacheService.getScriptCache().get('eticket_'+p.token);if(!s)throw Error('Sesi berakhir. Silakan masuk kembali.');return JSON.parse(s)}
function getAppData_(p){
 session_(p);const ss=getDatabase_(),santri=objects_(ss.getSheetByName('Santri')),laptop=objects_(ss.getSheetByName('Laptop')),penitipan=objects_(ss.getSheetByName('Penitipan')),pemakaian=objects_(ss.getSheetByName('Pemakaian')),pelanggaran=objects_(ss.getSheetByName('Pelanggaran')),now=new Date();
 return {user:session_(p),santri:santri.reverse(),laptop:laptop.reverse(),penitipan:penitipan.reverse(),pemakaian:pemakaian.reverse(),pelanggaran:pelanggaran.reverse(),recent:objects_(ss.getSheetByName('LogAktivitas')).slice(-10).reverse(),stats:{santri:santri.filter(x=>x.Status!=='Nonaktif').length,laptop:laptop.length,tersimpan:penitipan.filter(x=>x.Status==='Dititipkan').length,dipakai:pemakaian.filter(x=>x.Status==='Dipinjam').length,terlambat:pemakaian.filter(x=>x.Status==='Dipinjam'&&x.Batas_Kembali&&new Date(x.Batas_Kembali)<now).length,pelanggaran:pelanggaran.filter(x=>x.Tanggal&&new Date(x.Tanggal).getMonth()===now.getMonth()&&new Date(x.Tanggal).getFullYear()===now.getFullYear()).length}};
}
function saveSimple_(name,p,required){
 const u=session_(p);required.forEach(k=>{if(!String(p[k]||'').trim())throw Error('Kolom wajib belum diisi: '+k)});
 const sh=getDatabase_().getSheetByName(name),headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],all=objects_(sh),key=headers[0],id=p[key]||({Santri:'SNT',Laptop:'LPT'}[name]+'-'+Utilities.getUuid().slice(0,8).toUpperCase());
 if(name==='Santri'&&all.some(x=>String(x.NIS).trim()===String(p.NIS).trim()))throw Error('NIS sudah terdaftar.');
 if(name==='Laptop'&&all.some(x=>String(x.Kode_Aset).trim()===String(p.Kode_Aset).trim()))throw Error('Kode aset sudah digunakan.');
 const o=Object.assign({},p);delete o.Foto_DataUrl;
 if(name==='Santri'&&p.Foto_DataUrl){
  const match=String(p.Foto_DataUrl).match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/);
  if(!match)throw Error('Format foto tidak valid. Gunakan JPG, PNG, atau WebP.');
  if(match[2].length>1000000)throw Error('Ukuran foto terlalu besar. Pilih foto yang lebih kecil.');
  let folderId=PropertiesService.getScriptProperties().getProperty('ETIKET_FOTO_FOLDER_ID'),folder;
  if(folderId){try{folder=DriveApp.getFolderById(folderId)}catch(e){folder=null}}
  if(!folder){folder=DriveApp.createFolder('E-Tiket Laptop FM - Foto Santri');PropertiesService.getScriptProperties().setProperty('ETIKET_FOTO_FOLDER_ID',folder.getId())}
  const mime='image/'+(match[1]==='jpeg'?'jpeg':match[1]),ext=match[1]==='jpeg'?'jpg':match[1];
  const blob=Utilities.newBlob(Utilities.base64Decode(match[2]),mime,'Foto-Santri-'+String(p.NIS).replace(/[^a-zA-Z0-9_-]/g,'_')+'.'+ext);
  const file=folder.createFile(blob);o.Foto_File_ID=file.getId();
 }
 o[key]=id;o.Status=p.Status||(name==='Laptop'?'Tersedia':'Aktif');o.Dibuat_Pada=new Date();o.Diubah_Pada=new Date();
 sh.appendRow(headers.map(h=>o[h]===undefined?'':o[h]));log_(u.username,'TAMBAH',name,id,p.Nama||p.Kode_Aset||'');return {message:'Data berhasil disimpan.'};
}
function savePrivatePhoto_(dataUrl, prefix, reference){
 if(!dataUrl)return '';
 const match=String(dataUrl).match(/^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/);
 if(!match)throw Error('Format foto tidak valid. Gunakan JPG, PNG, atau WebP.');
 if(match[2].length>1000000)throw Error('Ukuran foto terlalu besar. Pilih foto yang lebih kecil.');
 let folderId=PropertiesService.getScriptProperties().getProperty('ETIKET_FOTO_FOLDER_ID'),folder;
 if(folderId){try{folder=DriveApp.getFolderById(folderId)}catch(e){folder=null}}
 if(!folder){folder=DriveApp.createFolder('E-Tiket Laptop FM - Foto Privat');PropertiesService.getScriptProperties().setProperty('ETIKET_FOTO_FOLDER_ID',folder.getId())}
 const mime='image/'+match[1],ext=match[1]==='jpeg'?'jpg':match[1];
 const safe=String(reference||Utilities.getUuid()).replace(/[^a-zA-Z0-9_-]/g,'_');
 const blob=Utilities.newBlob(Utilities.base64Decode(match[2]),mime,prefix+'-'+safe+'.'+ext);
 return folder.createFile(blob).getId();
}
function getFotoSantri_(p){
 session_(p);
 const id=String(p.ID_Santri||'').trim();if(!id)throw Error('ID santri tidak ditemukan.');
 const student=objects_(getDatabase_().getSheetByName('Santri')).find(x=>String(x.ID_Santri)===id);
 if(!student||!student.Foto_File_ID)throw Error('Foto santri belum tersedia.');
 const file=DriveApp.getFileById(String(student.Foto_File_ID)),blob=file.getBlob(),mime=blob.getContentType()||'image/jpeg';
 if(!/^image\/(jpeg|png|webp)$/.test(mime))throw Error('File foto tidak valid.');
 return {dataUrl:'data:'+mime+';base64,'+Utilities.base64Encode(blob.getBytes())};
}
function getFotoBukti_(p){
 session_(p);
 const type=String(p.type||''),id=String(p.id||'').trim();
 if(!id||!['pemakaian','pelanggaran'].includes(type))throw Error('Data foto tidak valid.');
 const sheetName=type==='pemakaian'?'Pemakaian':'Pelanggaran';
 const key=type==='pemakaian'?'ID_Pemakaian':'ID_Pelanggaran';
 const photoKey=type==='pemakaian'?'Foto_File_ID':'Foto_Bukti_File_ID';
 const row=objects_(getDatabase_().getSheetByName(sheetName)).find(x=>String(x[key])===id);
 if(!row||!row[photoKey])throw Error('Foto belum tersedia.');
 const blob=DriveApp.getFileById(String(row[photoKey])).getBlob(),mime=blob.getContentType()||'image/jpeg';
 if(!/^image\/(jpeg|png|webp)$/.test(mime))throw Error('File foto tidak valid.');
 return {dataUrl:'data:'+mime+';base64,'+Utilities.base64Encode(blob.getBytes())};
}
function saveLaptopAndDeposit_(p){
 const u=session_(p),ss=getDatabase_(),allSantri=objects_(ss.getSheetByName('Santri')),activeDeposits=objects_(ss.getSheetByName('Penitipan')).filter(x=>String(x.Status||'').trim().toLowerCase()==='dititipkan'),hasActiveDeposit=activeDeposits.some(x=>String(x.ID_Santri)===String(p.ID_Santri)),s=allSantri.find(x=>String(x.ID_Santri)===String(p.ID_Santri)&&(String(x.Status||'').trim().toLowerCase()!=='nonaktif'||hasActiveDeposit));
 if(!s)throw Error('Pilih pemilik dari data santri aktif. Santri nonaktif hanya dapat dipilih jika masih memiliki penitipan aktif.');
 if(!String(p.Kode_Aset||'').trim())throw Error('Kode aset wajib diisi.');
 const laptops=objects_(ss.getSheetByName('Laptop'));
 if(laptops.some(x=>String(x.Kode_Aset).trim().toLowerCase()===String(p.Kode_Aset).trim().toLowerCase()))throw Error('Kode aset sudah digunakan.');
 const now=new Date(),laptopId='LPT-'+Utilities.getUuid().slice(0,8).toUpperCase(),depositId='TTP-'+Utilities.getUuid().slice(0,8).toUpperCase();
 const laptop={ID_Laptop:laptopId,Kode_Aset:String(p.Kode_Aset).trim(),Merek_Model:p.Merek_Model||'',Nomor_Seri:p.Nomor_Seri||'',Kelengkapan:p.Kelengkapan||'',Kondisi:p.Kondisi||'Baik',Status:'Dititipkan',Catatan:p.Catatan||'',Dibuat_Pada:now,Diubah_Pada:now};
 append_('Laptop',laptop);
 append_('Penitipan',{ID_Transaksi:depositId,Tanggal_Masuk:now,ID_Santri:s.ID_Santri,Nama_Santri:s.Nama,ID_Laptop:laptopId,Kode_Aset:laptop.Kode_Aset,Kelengkapan:laptop.Kelengkapan,Kondisi_Masuk:laptop.Kondisi,Petugas:u.name,Status:'Dititipkan',Catatan:p.Catatan||'',Dibuat_Pada:now});
 log_(u.username,'TAMBAH','Laptop',laptopId,laptop.Kode_Aset+' / pemilik: '+s.Nama);
 log_(u.username,'PENITIPAN','Penitipan',depositId,s.Nama+' / '+laptop.Kode_Aset);
 return {message:'Laptop berhasil ditambahkan dan otomatis dititipkan atas nama '+s.Nama+'.'};
}
function saveDeposit_(p){
 const u=session_(p),ss=getDatabase_(),s=objects_(ss.getSheetByName('Santri')).find(x=>x.ID_Santri===p.ID_Santri),l=objects_(ss.getSheetByName('Laptop')).find(x=>x.ID_Laptop===p.ID_Laptop);
 if(!s||!l)throw Error('Pilih santri dan laptop yang valid.');
 if(objects_(ss.getSheetByName('Penitipan')).some(x=>x.ID_Laptop===l.ID_Laptop&&x.Status==='Dititipkan'))throw Error('Laptop masih tercatat dititipkan.');
 const id='TTP-'+Utilities.getUuid().slice(0,8).toUpperCase(),now=new Date();
 append_('Penitipan',{ID_Transaksi:id,Tanggal_Masuk:now,ID_Santri:s.ID_Santri,Nama_Santri:s.Nama,ID_Laptop:l.ID_Laptop,Kode_Aset:l.Kode_Aset,Kelengkapan:p.Kelengkapan||l.Kelengkapan,Kondisi_Masuk:p.Kondisi||l.Kondisi||'Baik',Petugas:u.name,Status:'Dititipkan',Catatan:p.Catatan||'',Dibuat_Pada:now});
 update_('Laptop','ID_Laptop',l.ID_Laptop,{Status:'Dititipkan',Diubah_Pada:now});log_(u.username,'PENITIPAN','Penitipan',id,s.Nama+' / '+l.Kode_Aset);return {message:'Penitipan berhasil dicatat.'};
}
function checkout_(p){
 const u=session_(p),ss=getDatabase_(),s=objects_(ss.getSheetByName('Santri')).find(x=>String(x.ID_Santri)===String(p.ID_Santri)&&x.Status!=='Nonaktif');
 if(!s||!p.Batas_Kembali)throw Error('Santri aktif dan batas kembali wajib diisi.');
 const activeDeposits=objects_(ss.getSheetByName('Penitipan')).filter(x=>String(x.ID_Santri)===String(s.ID_Santri)&&x.Status==='Dititipkan');
 if(activeDeposits.length===0)throw Error('Santri ini tidak memiliki laptop yang sedang dititipkan. Pemakaian ditolak.');
 if(activeDeposits.length>1)throw Error('Santri ini memiliki lebih dari satu laptop dengan status penitipan aktif. Rapikan data penitipan terlebih dahulu.');
 const deposit=activeDeposits[0],l=objects_(ss.getSheetByName('Laptop')).find(x=>String(x.ID_Laptop)===String(deposit.ID_Laptop));
 if(!l)throw Error('Data laptop pada penitipan tidak ditemukan. Periksa data penitipan.');
 if(objects_(ss.getSheetByName('Penitipan')).some(x=>String(x.ID_Laptop)===String(l.ID_Laptop)&&x.Status==='Dititipkan'&&String(x.ID_Santri)!==String(s.ID_Santri)))throw Error('Laptop ini memiliki konflik data pemilik penitipan. Pemakaian ditolak.');
 if(objects_(ss.getSheetByName('Pemakaian')).some(x=>x.ID_Laptop===l.ID_Laptop&&x.Status==='Dipinjam'))throw Error('Laptop sedang dipinjam.');
 if(new Date(p.Batas_Kembali)<=new Date())throw Error('Batas kembali harus di masa mendatang.');
 const id='PMK-'+Utilities.getUuid().slice(0,8).toUpperCase(),now=new Date(),fotoId=savePrivatePhoto_(p.Foto_DataUrl,'Foto-Pemakaian',id);
 append_('Pemakaian',{ID_Pemakaian:id,Tanggal_Pinjam:now,Batas_Kembali:new Date(p.Batas_Kembali),ID_Santri:s.ID_Santri,Nama_Santri:s.Nama,ID_Laptop:l.ID_Laptop,Kode_Aset:l.Kode_Aset,Tujuan:p.Tujuan||'',Pemberi_Izin:p.Pemberi_Izin||u.name,Petugas_Keluar:u.name,Kondisi_Keluar:p.Kondisi||'Baik',Status:'Dipinjam',Foto_File_ID:fotoId,Catatan:p.Catatan||'',Dibuat_Pada:now});
 update_('Laptop','ID_Laptop',l.ID_Laptop,{Status:'Dipinjam',Diubah_Pada:now});log_(u.username,'PEMINJAMAN','Pemakaian',id,s.Nama+' / '+l.Kode_Aset);return {message:'Pemakaian laptop berhasil dicatat.'};
}
function returnLaptop_(p){
 const u=session_(p),ss=getDatabase_(),rows=objects_(ss.getSheetByName('Pemakaian')),x=rows.find(r=>r.ID_Pemakaian===p.ID_Pemakaian&&r.Status==='Dipinjam');if(!x)throw Error('Peminjaman aktif tidak ditemukan.');
 update_('Pemakaian','ID_Pemakaian',x.ID_Pemakaian,{Tanggal_Kembali:new Date(),Petugas_Kembali:u.name,Kondisi_Kembali:p.Kondisi||'Baik',Status:'Dikembalikan'});
 const activeDeposit=objects_(ss.getSheetByName('Penitipan')).some(r=>r.ID_Laptop===x.ID_Laptop&&r.Status==='Dititipkan');
 update_('Laptop','ID_Laptop',x.ID_Laptop,{Status:activeDeposit?'Dititipkan':'Tersedia',Kondisi:p.Kondisi||'Baik',Diubah_Pada:new Date()});log_(u.username,'PENGEMBALIAN','Pemakaian',x.ID_Pemakaian,x.Kode_Aset);return {message:'Pengembalian berhasil dicatat.'};
}
function saveViolation_(p){
 const u=session_(p),s=objects_(getDatabase_().getSheetByName('Santri')).find(x=>x.ID_Santri===p.ID_Santri);if(!s||!p.Kategori||!p.Uraian)throw Error('Santri, kategori, dan uraian wajib diisi.');
 const id='PLG-'+Utilities.getUuid().slice(0,8).toUpperCase(),now=new Date(),fotoId=savePrivatePhoto_(p.Foto_Bukti_DataUrl,'Foto-Bukti-Pelanggaran',id);append_('Pelanggaran',{ID_Pelanggaran:id,Tanggal:p.Tanggal?new Date(p.Tanggal):now,ID_Santri:s.ID_Santri,Nama_Santri:s.Nama,Kategori:p.Kategori,Tingkat:p.Tingkat||'Ringan',Lokasi:p.Lokasi||'',Uraian:p.Uraian,Bukti_URL:p.Bukti_URL||'',Foto_Bukti_File_ID:fotoId,Tindak_Lanjut:p.Tindak_Lanjut||'',Status_Penanganan:'Belum Ditindaklanjuti',Pencatat:u.name,Dibuat_Pada:now,Diubah_Pada:now});log_(u.username,'CATAT PELANGGARAN','Pelanggaran',id,s.Nama+' / '+p.Kategori);return {message:'Pelanggaran berhasil dicatat.'};
}
function report_(p){session_(p);const ss=getDatabase_(),start=p.start?new Date(p.start):new Date(new Date().getFullYear(),new Date().getMonth(),1),end=p.end?new Date(p.end+'T23:59:59'):new Date(),within=(v)=>v&&new Date(v)>=start&&new Date(v)<=end;const data={penitipan:objects_(ss.getSheetByName('Penitipan')).filter(x=>within(x.Tanggal_Masuk)),pemakaian:objects_(ss.getSheetByName('Pemakaian')).filter(x=>within(x.Tanggal_Pinjam)),pelanggaran:objects_(ss.getSheetByName('Pelanggaran')).filter(x=>within(x.Tanggal))};return {data:data,totals:{penitipan:data.penitipan.length,pemakaian:data.pemakaian.length,pelanggaran:data.pelanggaran.length}}}
function objects_(sh){if(!sh||sh.getLastRow()<2)return [];const a=sh.getDataRange().getValues(),h=a.shift();return a.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i] instanceof Date?r[i].toISOString():r[i]);return o})}
function append_(name,o){const sh=getDatabase_().getSheetByName(name),h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];sh.appendRow(h.map(k=>o[k]===undefined?'':o[k]))}
function update_(name,key,id,changes){const sh=getDatabase_().getSheetByName(name),h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],r=objects_(sh),i=r.findIndex(x=>String(x[key])===String(id));if(i<0)throw Error('Data tidak ditemukan.');const old=sh.getRange(i+2,1,1,h.length).getValues()[0],o={};h.forEach((k,j)=>o[k]=old[j]);Object.assign(o,changes);sh.getRange(i+2,1,1,h.length).setValues([h.map(k=>o[k]===undefined?'':o[k])])}
function log_(u,a,m,id,d){try{append_('LogAktivitas',{ID_Log:'LOG-'+Utilities.getUuid().slice(0,8),Waktu:new Date(),Username:u,Aksi:a,Modul:m,ID_Referensi:id,Detail:d||''})}catch(e){}}

/**
 * Reset password akun admin yang sudah ada ke password awal 123456.
 * Jalankan manual hanya bila diperlukan; tidak mengubah akun lain.
 */
function resetPasswordAdminAwal() {
  const sh = getDatabase_().getSheetByName('Pengguna');
  if (!sh) throw new Error('Sheet Pengguna tidak ditemukan.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('Data pengguna belum tersedia.');

  const headers = values[0].map(h => String(h).trim());
  const usernameCol = headers.indexOf('Username');
  const hashCol = headers.indexOf('Password_Hash');
  const statusCol = headers.indexOf('Status');

  if (usernameCol < 0 || hashCol < 0 || statusCol < 0) {
    throw new Error('Kolom Username, Password_Hash, atau Status tidak ditemukan.');
  }

  const rowIndex = values.findIndex((row, i) =>
    i > 0 &&
    String(row[usernameCol]).trim().toLowerCase() === 'admin' &&
    String(row[statusCol]).trim() === 'Aktif'
  );
  if (rowIndex < 1) throw new Error('Akun admin aktif tidak ditemukan.');

  sh.getRange(rowIndex + 1, hashCol + 1).setValue(hashPassword_('123456'));
  log_('admin', 'RESET_PASSWORD', 'Pengguna', '', 'Password admin diperbarui oleh fungsi reset.');
  return 'Password akun admin berhasil diperbarui ke 123456.';
}

/**
 * Buat/perbarui dua akun petugas khusus.
 * Jalankan manual satu kali dari editor Apps Script.
 * Tidak mengubah akun admin utama.
 */
function setupAkunPetugas() {
  const sh = getDatabase_().getSheetByName('Pengguna');
  if (!sh) throw new Error('Sheet Pengguna tidak ditemukan. Jalankan setupDatabase() terlebih dahulu.');

  const values = sh.getDataRange().getValues();
  if (!values.length) throw new Error('Header sheet Pengguna tidak ditemukan.');
  const headers = values[0].map(h => String(h).trim());
  const col = {};
  ['ID_Pengguna','Nama','Username','Password_Hash','Peran','Status','Dibuat_Pada','Login_Terakhir'].forEach(k => {
    col[k] = headers.indexOf(k);
    if (col[k] < 0) throw new Error('Kolom ' + k + ' tidak ditemukan pada sheet Pengguna.');
  });

  const accounts = [
    { username: 'adminpakai', name: 'Admin Pemakaian', role: 'Admin Pemakaian' },
    { username: 'admindisiplin', name: 'Admin Pelanggaran', role: 'Admin Pelanggaran' }
  ];
  const now = new Date();
  accounts.forEach(account => {
    const rowIndex = values.findIndex((row, i) =>
      i > 0 && String(row[col.Username]).trim().toLowerCase() === account.username
    );
    if (rowIndex > 0) {
      const rowNumber = rowIndex + 1;
      sh.getRange(rowNumber, col.Nama + 1).setValue(account.name);
      sh.getRange(rowNumber, col.Password_Hash + 1).setValue(hashPassword_('123456'));
      sh.getRange(rowNumber, col.Peran + 1).setValue(account.role);
      sh.getRange(rowNumber, col.Status + 1).setValue('Aktif');
    } else {
      const row = [];
      headers.forEach(() => row.push(''));
      row[col.ID_Pengguna] = 'USR-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      row[col.Nama] = account.name;
      row[col.Username] = account.username;
      row[col.Password_Hash] = hashPassword_('123456');
      row[col.Peran] = account.role;
      row[col.Status] = 'Aktif';
      row[col.Dibuat_Pada] = now;
      sh.appendRow(row);
      values.push(row);
    }
  });

  log_('admin', 'SETUP_AKUN_PETUGAS', 'Pengguna', '', 'Akun adminpakai dan admindisiplin dibuat/diperbarui.');
  return 'Akun adminpakai dan admindisiplin berhasil dibuat/diperbarui dengan password awal 123456.';
}

