// ==========================================
// 1. SETUP & INISIALISASI DATABASE (JALANKAN INI PERTAMA KALI)
// ==========================================
function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupDatabase() {
  const db = getDb();
  
  // Skema Database beserta Data Dummy
  const schema = {
    'Sistem': {
      headers: ['key', 'value'],
      dummy: [
        ['nama_aplikasi', 'AI VS YOU'],
        ['nama_sekolah', 'SMAN 1 Bululawang'],
        ['tagline', 'Don’t Just Ask AI. Challenge It.'],
        ['kepala_sekolah', 'Dr. H. Kepala Sekolah, M.Pd'],
        ['nip_kepsek', '197001012000121001']
      ]
    },
    'Kelas': {
      headers: ['id_kelas', 'nama_kelas'],
      dummy: [
        ['KLS_001', 'X MIPA 1'],
        ['KLS_002', 'XI IPS 1']
      ]
    },
    'Mapel': {
      headers: ['id_mapel', 'nama_mapel'],
      dummy: [
        ['MPL_001', 'Informatika'],
        ['MPL_002', 'Bahasa Indonesia']
      ]
    },
    'Users': {
      headers: ['id_user', 'username', 'password', 'role', 'nama_lengkap', 'nip'],
      dummy: [
        ['USR_ADMIN', 'admin', 'admin123', 'admin', 'Administrator Sistem', '1234567890'],
        ['USR_GURU1', 'guru', 'guru123', 'guru', 'Budi Santoso, S.Pd', '198001012010011002']
      ]
    },
    'Siswa': {
      headers: ['id_siswa', 'username', 'password', 'nama_siswa', 'nisn', 'id_kelas'],
      dummy: [
        ['SSW_001', 'siswa1', 'siswa123', 'Andi Dharmawan', '0051234567', 'KLS_001'],
        ['SSW_002', 'siswa2', 'siswa123', 'Budi Haryanto', '0057654321', 'KLS_001']
      ]
    },
    'Aktivitas': {
      headers: ['id_aktivitas', 'id_guru', 'id_kelas', 'id_mapel', 'tipe_tugas', 'detail_json', 'status', 'created_at'],
      dummy: [
        [
          'AKT_001', 'USR_GURU1', 'KLS_001', 'MPL_001', 'cek_ai', 
          JSON.stringify({judul: "Cek Fakta Sejarah Komputer", deskripsi: "Gunakan AI untuk mencari penemu komputer, lalu verifikasi hasilnya.", ai_tools: "ChatGPT"}), 
          'aktif', new Date().toISOString()
        ]
      ]
    },
    'Submisi': {
      headers: ['id_submisi', 'id_aktivitas', 'id_siswa', 'hasil_json', 'nilai', 'feedback_guru', 'waktu_kumpul'],
      dummy: [] // Dibiarkan kosong agar bisa dites kumpul tugas oleh siswa
    }
  };

  for (const sheetName in schema) {
    let sheet = db.getSheetByName(sheetName);
    const tableInfo = schema[sheetName];

    // Jika sheet belum ada, buat baru
    if (!sheet) {
      sheet = db.insertSheet(sheetName);
      
      // Set Header di baris 1
      sheet.getRange(1, 1, 1, tableInfo.headers.length).setValues([tableInfo.headers]);
      sheet.getRange(1, 1, 1, tableInfo.headers.length).setFontWeight("bold").setBackground("#d9ead3"); // Desain header hijau tipis
      
      // Masukkan Data Dummy jika ada (mulai baris 2)
      if (tableInfo.dummy.length > 0) {
        sheet.getRange(2, 1, tableInfo.dummy.length, tableInfo.headers.length).setValues(tableInfo.dummy);
      }
    }
  }

  // Hapus "Sheet1" bawaan default jika masih ada
  const defaultSheet = db.getSheetByName('Sheet1');
  if (defaultSheet && db.getSheets().length > 1) {
    db.deleteSheet(defaultSheet);
  }
}

// ==========================================
// 2. ROUTING WEB
// ==========================================
function doGet(e) {
  let html = HtmlService.createTemplateFromFile('index');
  return html.evaluate()
      .setTitle('AI VS YOU - SMAN 1 Bululawang')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 3. UTILITY FUNCTIONS
// ==========================================
function generateId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

function getSheetData(sheetName) {
  const sheet = getDb().getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

// ==========================================
// 4. AUTENTIKASI
// ==========================================
function processLogin(username, password) {
  const users = getSheetData('Users');
  let user = users.find(u => u.username === username && String(u.password) === String(password));
  
  if (user) {
    return { success: true, role: user.role, data: user };
  }
  
  const siswa = getSheetData('Siswa');
  let dataSiswa = siswa.find(s => s.username === username && String(s.password) === String(password));
  
  if (dataSiswa) {
    const kelasData = getSheetData('Kelas').find(k => k.id_kelas === dataSiswa.id_kelas);
    dataSiswa.nama_kelas = kelasData ? kelasData.nama_kelas : '-';
    return { success: true, role: 'siswa', data: dataSiswa };
  }
  
  return { success: false, message: 'Username atau password salah!' };
}

// ==========================================
// 5. CRUD ADMIN (Menambah Data)
// ==========================================
function addKelas(namaKelas) {
  const sheet = getDb().getSheetByName('Kelas');
  sheet.appendRow([generateId('KLS'), namaKelas]);
  return { success: true, message: 'Kelas berhasil ditambahkan' };
}

function addMapel(namaMapel) {
  const sheet = getDb().getSheetByName('Mapel');
  sheet.appendRow([generateId('MPL'), namaMapel]);
  return { success: true, message: 'Mata pelajaran berhasil ditambahkan' };
}

function addUser(data) {
  const sheet = getDb().getSheetByName('Users');
  sheet.appendRow([generateId('USR'), data.username, data.password, data.role, data.nama_lengkap, data.nip]);
  return { success: true, message: 'User berhasil ditambahkan' };
}

function addSiswa(data) {
  const sheet = getDb().getSheetByName('Siswa');
  sheet.appendRow([generateId('SSW'), data.username, data.password, data.nama_siswa, data.nisn, data.id_kelas]);
  return { success: true, message: 'Siswa berhasil ditambahkan' };
}

// ==========================================
// 6. MANAJEMEN GURU & SISWA
// ==========================================
function buatAktivitas(dataTugas) {
  const sheet = getDb().getSheetByName('Aktivitas');
  const timestamp = new Date().toISOString();
  const detailJson = JSON.stringify({
    judul: dataTugas.judul, deskripsi: dataTugas.deskripsi, ai_tools: dataTugas.ai_tools
  });
  
  sheet.appendRow([
    generateId('AKT'), dataTugas.id_guru, dataTugas.id_kelas, dataTugas.id_mapel, 
    dataTugas.tipe_tugas, detailJson, 'aktif', timestamp
  ]);
  return { success: true, message: 'Aktivitas berhasil dibuat' };
}

function kumpulTugas(id_aktivitas, id_siswa, payload_jawaban) {
  const sheet = getDb().getSheetByName('Submisi');
  const timestamp = new Date().toISOString();
  const hasilJson = JSON.stringify(payload_jawaban);
  
  sheet.appendRow([generateId('SUB'), id_aktivitas, id_siswa, hasilJson, '', '', timestamp]);
  return { success: true, message: 'Tugas berhasil dikirim.' };
}

function simpanNilai(id_submisi, nilai, feedback) {
  const sheet = getDb().getSheetByName('Submisi');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id_submisi) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 5).setValue(nilai); 
      sheet.getRange(rowIndex, 6).setValue(feedback);
      return { success: true, message: 'Nilai dan Feedback berhasil disimpan' };
    }
  }
  return { success: false, message: 'Data submisi tidak ditemukan' };
}