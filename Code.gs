/**
 * ============================================================================
 * SIMAGIS-PRO-BACKEND (Google Apps Script)
 * Instansi : SMA Negeri 10 Bone
 * Fitur    : RBAC (admin, guru_bk, wali_kelas, guru_mapel), LockService,
 *            Incremental Pull Sync, Timestamp-based Upsert, & Drive JSON Backup
 * ============================================================================
 */

const CONFIG = {
  APP_NAME: 'SIMAGIS PRO - SMA Negeri 10 Bone',
  BACKUP_FOLDER_NAME: 'SIMAGIS_PRO_BACKUP_SMAN10BONE',
  LOCK_TIMEOUT_MS: 15000,
  SHEETS: {
    Users: ['id', 'nip', 'nama', 'role', 'pin_hash', 'kelas_binaan', 'mapel_ampu', 'status_aktif', 'updated_at'],
    Jurnal_Mengajar: ['id_jurnal', 'tanggal', 'kelas', 'mapel', 'materi', 'kode_guru', 'device_id', 'updated_at', 'deleted_at'],
    Absensi_Siswa: ['id_absensi', 'tanggal', 'nisn', 'nama_siswa', 'kelas', 'status', 'keterangan', 'kode_guru', 'device_id', 'updated_at', 'deleted_at'],
    Nilai_Asesmen: ['id_nilai', 'nisn', 'nama_siswa', 'kelas', 'mapel', 'jenis_asesmen', 'skor', 'kode_guru', 'device_id', 'updated_at', 'deleted_at'],
    Layanan_BK: ['id_layanan', 'tanggal', 'nisn', 'nama_siswa', 'kelas', 'jenis_layanan', 'ringkasan_kasus', 'tindak_lanjut', 'kode_bk', 'device_id', 'updated_at', 'deleted_at'],
    Log_Backup: ['timestamp', 'nama_file', 'drive_file_id', 'sha256_hash', 'ukuran_byte', 'pemicu', 'oleh_nip']
  }
};

// Matriks Hak Akses per Peran
const ROLE_PERMISSIONS = {
  admin: {
    read: ['Users', 'Jurnal_Mengajar', 'Absensi_Siswa', 'Nilai_Asesmen', 'Layanan_BK', 'Log_Backup'],
    write: ['Users', 'Jurnal_Mengajar', 'Absensi_Siswa', 'Nilai_Asesmen', 'Layanan_BK', 'Log_Backup']
  },
  guru_bk: {
    read: ['Layanan_BK', 'Absensi_Siswa'],
    write: ['Layanan_BK']
  },
  wali_kelas: {
    read: ['Absensi_Siswa', 'Nilai_Asesmen'],
    write: ['Absensi_Siswa']
  },
  guru_mapel: {
    read: ['Jurnal_Mengajar', 'Absensi_Siswa', 'Nilai_Asesmen'],
    write: ['Jurnal_Mengajar', 'Absensi_Siswa', 'Nilai_Asesmen']
  }
};

/**
 * Jalankan fungsi ini SEKALI dari editor Apps Script untuk membuat
 * 6 sheet beserta header dan 1 akun admin awal.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CONFIG.SHEETS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const headers = CONFIG.SHEETS[sheetName];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e2e8f0');
      sheet.setFrozenRows(1);
    }
  });

  // Buat akun admin default jika sheet Users masih kosong
  const userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() === 1) {
    const defaultPinHash = computeSha256('123456'); // Segera ganti setelah login pertama
    userSheet.appendRow([
      Utilities.getUuid(),
      '197409222000011001',
      'Administrator SIMAGIS',
      'admin',
      defaultPinHash,
      '-',
      '-',
      'AKTIF',
      new Date().toISOString()
    ]);
  }

  return 'Inisialisasi 6 sheet SIMAGIS-PRO-BACKEND berhasil.';
}

/**
 * ============================================================================
 * HANDLER HTTP GET: Portal Web & Incremental Pull Sync
 * ============================================================================
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || '';

    // Jika diakses tanpa parameter action, tampilkan UI Web App (Index.html)
    if (!action) {
      return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle(CONFIG.APP_NAME)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (action === 'ping') {
      return jsonResponse({
        ok: true,
        app: CONFIG.APP_NAME,
        server_time: new Date().toISOString()
      });
    }

    // Autentikasi pengguna untuk rute API
    const authUser = authenticateUser(params.nip, params.pin_hash);
    if (!authUser) {
      return jsonResponse({ ok: false, error: 'UNAUTHORIZED: NIP atau PIN tidak valid / akun nonaktif.' }, 401);
    }

    if (action === 'pull') {
      const requestedTables = params.tables
        ? params.tables.split(',').map(t => t.trim())
        : ROLE_PERMISSIONS[authUser.role].read;
      const since = params.since || ''; // ISO timestamp sinkronisasi terakhir klien

      const payload = {};
      const allowedRead = ROLE_PERMISSIONS[authUser.role].read;

      requestedTables.forEach(tableName => {
        if (allowedRead.includes(tableName) && CONFIG.SHEETS[tableName]) {
          payload[tableName] = fetchTableData(tableName, authUser, since);
        }
      });

      return jsonResponse({
        ok: true,
        role: authUser.role,
        server_time: new Date().toISOString(),
        data: payload
      });
    }

    return jsonResponse({ ok: false, error: 'UNKNOWN_ACTION: Parameter action tidak dikenali.' }, 400);

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, 500);
  }
}

/**
 * ============================================================================
 * HANDLER HTTP POST: Batch Upsert Sync, Auth Login, & Drive JSON Backup
 * ============================================================================
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'EMPTY_PAYLOAD: Body POST kosong.' }, 400);
    }

    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';

    // Verifikasi hash integritas payload opsional dari klien
    if (body.payload_sha256 && body.data) {
      const calculated = computeSha256(JSON.stringify(body.data));
      if (calculated !== body.payload_sha256) {
        return jsonResponse({ ok: false, error: 'INTEGRITY_ERROR: Hash SHA-256 payload tidak cocok.' }, 422);
      }
    }

    // Autentikasi pengguna
    const authUser = authenticateUser(body.nip, body.pin_hash);
    if (!authUser) {
      return jsonResponse({ ok: false, error: 'UNAUTHORIZED: Kredensial tidak valid atau akun nonaktif.' }, 401);
    }

    // 1. Rute Login / Verifikasi Peran
    if (action === 'auth_login') {
      return jsonResponse({
        ok: true,
        user: {
          id: authUser.id,
          nip: authUser.nip,
          nama: authUser.nama,
          role: authUser.role,
          kelas_binaan: authUser.kelas_binaan,
          mapel_ampu: authUser.mapel_ampu
        },
        permissions: ROLE_PERMISSIONS[authUser.role],
        server_time: new Date().toISOString()
      });
    }

    // 2. Rute Push Batch (Upsert Berbasis Timestamp)
    if (action === 'push_batch') {
      const batches = body.data || {}; // Format: { Jurnal_Mengajar: [...], Absensi_Siswa: [...] }
      const allowedWrite = ROLE_PERMISSIONS[authUser.role].write;
      const syncResults = {};

      Object.keys(batches).forEach(tableName => {
        if (!CONFIG.SHEETS[tableName]) {
          syncResults[tableName] = { error: 'TABEL_TIDAK_VALID' };
          return;
        }
        if (!allowedWrite.includes(tableName)) {
          syncResults[tableName] = { error: 'FORBIDDEN_ROLE: Peran ' + authUser.role + ' tidak diizinkan menulis ke ' + tableName };
          return;
        }

        const records = Array.isArray(batches[tableName]) ? batches[tableName] : [];
        syncResults[tableName] = upsertRecords(tableName, records, authUser);
      });

      return jsonResponse({
        ok: true,
        server_time: new Date().toISOString(),
        results: syncResults
      });
    }

    // 3. Rute Backup Snapshot JSON ke Google Drive
    if (action === 'backup_drive') {
      if (authUser.role !== 'admin') {
        return jsonResponse({ ok: false, error: 'FORBIDDEN: Hanya admin yang dapat memicu backup cloud.' }, 403);
      }
      const backupMeta = createJsonBackupToDrive(body.snapshot || null, body.trigger || 'manual', authUser.nip);
      return jsonResponse({
        ok: true,
        backup: backupMeta
      });
    }

    return jsonResponse({ ok: false, error: 'UNKNOWN_POST_ACTION: Action tidak didukung.' }, 400);

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, 500);
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * LOGIKA UPSERT BERBASIS TIMESTAMP (LAST-WRITE-WINS)
 * ============================================================================
 */
function upsertRecords(sheetName, records, authUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = CONFIG.SHEETS[sheetName];
  const pkField = headers[0]; // Kolom pertama selalu primary key (id_jurnal, id_absensi, dll.)
  const updatedAtIdx = headers.indexOf('updated_at');

  const lastRow = sheet.getLastRow();
  const existingData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];

  // Indeks baris berdasarkan Primary Key untuk pencarian cepat O(1)
  const rowIndexMap = {};
  existingData.forEach((row, idx) => {
    const pkVal = String(row[0]).trim();
    if (pkVal) {
      rowIndexMap[pkVal] = {
        sheetRow: idx + 2,
        updatedAt: updatedAtIdx >= 0 ? String(row[updatedAtIdx]) : ''
      };
    }
  });

  const itemStatuses = [];
  const rowsToAppend = [];

  records.forEach(item => {
    const pkVal = String(item[pkField] || '').trim();
    if (!pkVal) {
      itemStatuses.push({ id: null, status: 'rejected', reason: 'Primary key kosong' });
      return;
    }

    // Amankan penulisan data sensitif
    if (sheetName === 'Users' && item.pin_plain) {
      item.pin_hash = computeSha256(String(item.pin_plain));
    }

    const incomingUpdatedAt = item.updated_at || new Date().toISOString();
    item.updated_at = incomingUpdatedAt;

    const rowArray = headers.map(col => (item[col] !== undefined && item[col] !== null) ? item[col] : '');
    const existing = rowIndexMap[pkVal];

    if (!existing) {
      // INSERT: Data baru dari klien offline
      rowsToAppend.push(rowArray);
      itemStatuses.push({ id: pkVal, status: 'inserted', updated_at: incomingUpdatedAt });
    } else {
      // UPDATE: Bandingkan timestamp (incoming >= existing)
      const existingTime = existing.updatedAt ? new Date(existing.updatedAt).getTime() : 0;
      const incomingTime = new Date(incomingUpdatedAt).getTime();

      if (incomingTime >= existingTime) {
        sheet.getRange(existing.sheetRow, 1, 1, headers.length).setValues([rowArray]);
        itemStatuses.push({ id: pkVal, status: 'updated', updated_at: incomingUpdatedAt });
      } else {
        // Server memiliki versi yang lebih baru -> kembalikan info konflik
        itemStatuses.push({
          id: pkVal,
          status: 'skipped_stale',
          server_updated_at: existing.updatedAt
        });
      }
    }
  });

  // Batch append untuk efisiensi eksekusi
  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }

  return itemStatuses;
}

/**
 * ============================================================================
 * PENARIKAN DATA TERFILTER PERAN & INCREMENTAL SYNC (SINCE)
 * ============================================================================
 */
function fetchTableData(sheetName, authUser, sinceIso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = CONFIG.SHEETS[sheetName];
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const sinceTime = sinceIso ? new Date(sinceIso).getTime() : 0;
  const results = [];

  values.forEach(row => {
    const obj = {};
    headers.forEach((h, i) => {
      // Jangan pernah kirim pin_hash ke klien saat pull tabel Users
      if (sheetName === 'Users' && h === 'pin_hash') return;
      obj[h] = row[i];
    });

    // Filter incremental sync berdasarkan updated_at
    if (sinceTime > 0 && obj.updated_at) {
      const rowTime = new Date(obj.updated_at).getTime();
      if (rowTime <= sinceTime) return;
    }

    // Filter pembatasan lingkup data per peran (Row-Level Security)
    if (authUser.role === 'guru_mapel') {
      if (obj.kode_guru && String(obj.kode_guru) !== String(authUser.nip)) return;
    } else if (authUser.role === 'wali_kelas') {
      if (authUser.kelas_binaan && authUser.kelas_binaan !== '-' && obj.kelas) {
        if (String(obj.kelas).toUpperCase() !== String(authUser.kelas_binaan).toUpperCase()) return;
      }
    }

    results.push(obj);
  });

  return results;
}

/**
 * ============================================================================
 * BACKUP SNAPSHOT JSON TERVERIFIKASI SHA-256 KE GOOGLE DRIVE
 * ============================================================================
 */
function createJsonBackupToDrive(clientSnapshot, triggerSource, nipPelaku) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snapshotData = clientSnapshot || {};

  // Jika klien tidak mengirim snapshot, ambil seluruh isi sheet dari server
  if (!clientSnapshot) {
    Object.keys(CONFIG.SHEETS).forEach(tableName => {
      if (tableName === 'Log_Backup') return;
      const sheet = ss.getSheetByName(tableName);
      const headers = CONFIG.SHEETS[tableName];
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
        snapshotData[tableName] = rows.map(r => {
          const item = {};
          headers.forEach((h, idx) => { item[h] = r[idx]; });
          return item;
        });
      } else {
        snapshotData[tableName] = [];
      }
    });
  }

  const jsonString = JSON.stringify({
    app: CONFIG.APP_NAME,
    exported_at: new Date().toISOString(),
    data: snapshotData
  }, null, 2);

  const sha256 = computeSha256(jsonString);
  const timestampStr = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyyMMdd_HHmmss');
  const fileName = `SIMAGIS_PRO_SMAN10Bone_${timestampStr}.json`;

  const folder = getOrCreateBackupFolder();
  const file = folder.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);

  const logSheet = ss.getSheetByName('Log_Backup');
  logSheet.appendRow([
    new Date().toISOString(),
    fileName,
    file.getId(),
    sha256,
    file.getSize(),
    triggerSource,
    nipPelaku
  ]);

  return {
    file_name: fileName,
    sha256: sha256,
    size_bytes: file.getSize(),
    url: file.getUrl()
  };
}

/**
 * ============================================================================
 * FUNGSI UTILITAS (AUTENTIKASI, SHA-256, FOLDER BACKUP, RESPONSE JSON)
 * ============================================================================
 */
function authenticateUser(nip, pinHash) {
  if (!nip || !pinHash) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return null;

  const headers = CONFIG.SHEETS.Users;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowNip = String(r[headers.indexOf('nip')]).trim();
    const rowPinHash = String(r[headers.indexOf('pin_hash')]).trim();
    const rowStatus = String(r[headers.indexOf('status_aktif')]).trim().toUpperCase();

    if (rowNip === String(nip).trim() && rowPinHash === String(pinHash).trim() && rowStatus === 'AKTIF') {
      return {
        id: r[headers.indexOf('id')],
        nip: rowNip,
        nama: r[headers.indexOf('nama')],
        role: String(r[headers.indexOf('role')]).trim().toLowerCase(),
        kelas_binaan: r[headers.indexOf('kelas_binaan')],
        mapel_ampu: r[headers.indexOf('mapel_ampu')]
      };
    }
  }
  return null;
}

function computeSha256(text) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return rawHash.map(byte => {
    const v = (byte < 0) ? byte + 256 : byte;
    return v.toString(16).padStart(2, '0');
  }).join('');
}

function getOrCreateBackupFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(CONFIG.BACKUP_FOLDER_NAME);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
