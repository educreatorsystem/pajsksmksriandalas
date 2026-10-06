const SHEET_ID = '1lf9FIF17DZG2M9AndpA1dmMh-Cnrp_EBbFmi5SjNe_M';
const FOLDER_ID = '1KLolK_Y-pIMhrLgCF9OEOFC6ZHQH32Sm';

const LAPORAN_SHEET = 'LAPORAN';
const JAWATAN_SHEET = 'JAWATAN';

const LAPORAN_FIELDS = {
  id: { canonical: 'ID', aliases: ['ID'] },
  startDate: { canonical: 'Tarikh Mula', aliases: ['Tarikh Mula', 'Tarikh', 'Tarikh Program'] },
  endDate: { canonical: 'Tarikh Akhir', aliases: ['Tarikh Akhir'] },
  name: { canonical: 'Nama Program', aliases: ['Nama Program'] },
  level: { canonical: 'Peringkat', aliases: ['Peringkat'] },
  unit: { canonical: 'Unit', aliases: ['Unit'] },
  category: { canonical: 'Kategori', aliases: ['Kategori'] },
  achievement: { canonical: 'Pencapaian', aliases: ['Pencapaian'] },
  obj: { canonical: 'Objektif', aliases: ['Objektif', 'Objektif Program'] },
  swot: { canonical: 'SWOT', aliases: ['SWOT', 'Analisis SWOT'] },
  students: { canonical: 'Murid (JSON)', aliases: ['Murid (JSON)', 'Murid List (JSON)', 'Senarai Murid (JSON)'] },
  teachers: { canonical: 'Guru (JSON)', aliases: ['Guru (JSON)', 'Guru List (JSON)', 'Senarai Guru (JSON)'] },
  images: { canonical: 'Gambar (JSON)', aliases: ['Gambar (JSON)', 'Gambar URLs (JSON)', 'Gambar URL (JSON)'] },
  preparedBy: { canonical: 'Nama Guru Sediakan Laporan', aliases: ['Nama Guru Sediakan Laporan', 'Disediakan Oleh'] },
  preparedPosition: { canonical: 'Jawatan Guru', aliases: ['Jawatan Guru', 'Jawatan Penyedia'] },
  createdAt: { canonical: 'Tarikh Dicipta', aliases: ['Tarikh Dicipta', 'Timestamp'] }
};

const JAWATAN_HEADERS = [
  'Tarikh Rekod Dicipta',
  'Nama Murid',
  'Kelas',
  'Unit',
  'Pecahan Unit',
  'Jawatan',
  'Tahun / Sesi'
];

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase();
}

function getHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(v => String(v || '').trim());
}

function findHeaderIndex(headers, aliases) {
  const normalized = headers.map(normalizeHeader);
  for (let i = 0; i < aliases.length; i++) {
    const idx = normalized.indexOf(normalizeHeader(aliases[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function getFieldIndex(headers, fieldName) {
  const field = LAPORAN_FIELDS[fieldName];
  return findHeaderIndex(headers, field.aliases);
}

function getFieldValue(row, headers, fieldName) {
  const index = getFieldIndex(headers, fieldName);
  if (index === -1) return '';
  const value = row[index];
  return value === null || value === undefined ? '' : value;
}

function ensureLaporanSheet(ss) {
  let sheet = ss.getSheetByName(LAPORAN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LAPORAN_SHEET);
    const headers = Object.keys(LAPORAN_FIELDS).map(key => LAPORAN_FIELDS[key].canonical);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  let headers = getHeaders(sheet);

  // Tambah hanya field yang benar-benar tiada. Alias lama seperti "Tarikh" atau
  // "Murid List (JSON)" dikekalkan dan terus digunakan, jadi data lama tidak rosak.
  Object.keys(LAPORAN_FIELDS).forEach(fieldName => {
    const field = LAPORAN_FIELDS[fieldName];
    if (findHeaderIndex(headers, field.aliases) === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(field.canonical);
      headers.push(field.canonical);
    }
  });

  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold')
    .setBackground('#dbeafe');
  sheet.setFrozenRows(1);

  return sheet;
}

function ensureJawatanSheet(ss) {
  let sheet = ss.getSheetByName(JAWATAN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(JAWATAN_SHEET);
    sheet.getRange(1, 1, 1, JAWATAN_HEADERS.length).setValues([JAWATAN_HEADERS]);
  }

  let headers = getHeaders(sheet);
  JAWATAN_HEADERS.forEach(header => {
    if (findHeaderIndex(headers, [header]) === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(header);
      headers.push(header);
    }
  });

  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold')
    .setBackground('#dbeafe');
  sheet.setFrozenRows(1);
  return sheet;
}

function safeJsonArray(value) {
  if (Array.isArray(value)) return value;
  if (value === null || value === undefined || value === '') return [];

  if (typeof value === 'object') {
    return Array.isArray(value) ? value : [];
  }

  const text = String(value).trim();
  if (!text) return [];

  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    // Cuba pulihkan JSON yang disimpan dengan tanda petik berganda.
    try {
      const unwrapped = text
        .replace(/^"|"$/g, '')
        .replace(/""/g, '"');
      const parsed = JSON.parse(unwrapped);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err2) {
      return [];
    }
  }
}

function extractDriveFileId(value) {
  if (!value) return '';
  const text = String(value).trim();

  // Jika memang hanya ID Drive.
  if (/^[a-zA-Z0-9_-]{20,}$/.test(text)) return text;

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]{20,})/,
    /[?&]id=([a-zA-Z0-9_-]{20,})/,
    /\/file\/d\/([a-zA-Z0-9_-]{20,})/,
    /googleusercontent\.com\/d\/([a-zA-Z0-9_-]{20,})/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match) return match[1];
  }

  return '';
}

function normalizeDriveImageUrl(value) {
  if (!value) return '';

  if (typeof value === 'object') {
    if (value.url) return normalizeDriveImageUrl(value.url);
    if (value.fileId) return 'https://lh3.googleusercontent.com/d/' + value.fileId;
    if (value.id) return 'https://lh3.googleusercontent.com/d/' + value.id;
    return '';
  }

  const text = String(value).trim();
  if (!text) return '';

  const fileId = extractDriveFileId(text);
  if (fileId) return 'https://lh3.googleusercontent.com/d/' + fileId;

  return text;
}

function normalizeImageArray(value) {
  let arr = safeJsonArray(value);

  // Jika sel lama hanya mengandungi satu URL atau senarai URL dipisahkan koma.
  if (arr.length === 0 && typeof value === 'string' && value.trim()) {
    const text = value.trim();
    if (/^https?:\/\//i.test(text)) {
      arr = text.split(',').map(v => v.trim()).filter(Boolean);
    }
  }

  return arr.map(normalizeDriveImageUrl).filter(Boolean);
}

function formatDateForClient(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

function appendLaporanByHeaders(sheet, data) {
  const headers = getHeaders(sheet);
  const row = new Array(headers.length).fill('');

  function setField(fieldName, value) {
    const idx = getFieldIndex(headers, fieldName);
    if (idx !== -1) row[idx] = value;
  }

  setField('id', data.id || 'LAP-' + Date.now());
  setField('startDate', data.startDate || '');
  setField('endDate', data.endDate || '');
  setField('name', data.name || '');
  setField('level', data.level || '');
  setField('unit', data.unit || '');
  setField('category', data.category || '');
  setField('achievement', data.achievement || '');
  setField('obj', data.obj || '');
  setField('swot', data.swot || '');
  setField('students', JSON.stringify(data.students || []));
  setField('teachers', JSON.stringify(data.teachers || []));
  setField('images', JSON.stringify(data.images || []));
  setField('preparedBy', data.preparedBy || '');
  setField('preparedPosition', data.preparedPosition || '');
  setField('createdAt', new Date());

  sheet.appendRow(row);
}

function getDriveFolder() {
  try {
    return DriveApp.getFolderById(FOLDER_ID);
  } catch (error) {
    throw new Error(
      'Akses Google Drive belum dibenarkan untuk Web App ini. ' +
      'Buka Apps Script, jalankan fungsi authorizeDriveAccess() sekali, benarkan akses Drive, ' +
      'kemudian Deploy semula sebagai New version dengan Execute as: Me. ' +
      'Butiran asal: ' + (error && error.message ? error.message : String(error))
    );
  }
}

/**
 * Jalankan fungsi ini SEKALI secara manual dari editor Apps Script.
 * Ia memaksa Google memaparkan skrin kebenaran OAuth untuk Google Drive.
 */
function authorizeDriveAccess() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const folderName = folder.getName();
  console.log('Akses Drive berjaya: ' + folderName);
  return 'Akses Drive berjaya: ' + folderName;
}

/**
 * Ujian bacaan + tulis ke folder. Fail ujian dipadam terus selepas berjaya.
 */
function testDriveWriteAccess() {
  const folder = getDriveFolder();
  const testFile = folder.createFile(
    Utilities.newBlob('Ujian akses Drive ' + new Date().toISOString(), 'text/plain', 'TEST_DRIVE_ACCESS.txt')
  );
  const result = {
    status: 'success',
    folderName: folder.getName(),
    folderId: FOLDER_ID,
    fileId: testFile.getId()
  };
  testFile.setTrashed(true);
  console.log(JSON.stringify(result));
  return result;
}

function saveImagesToDrive(newImages, existingImages) {
  const imageUrls = (existingImages || []).map(normalizeDriveImageUrl).filter(Boolean);
  if (!newImages || newImages.length === 0) return imageUrls;

  const folder = getDriveFolder();

  newImages.forEach(function(img, index) {
    if (!img || !img.base64) return;

    const bytes = Utilities.base64Decode(img.base64);
    const safeName = img.name || ('gambar-program-' + Date.now() + '-' + (index + 1) + '.png');
    const blob = Utilities.newBlob(
      bytes,
      img.mime || 'image/png',
      safeName
    );

    const file = folder.createFile(blob);

    // Tidak perlu setSharing() setiap kali. Fail dicipta di dalam folder yang telah
    // dikongsi dan ini mengelakkan ralat Access denied pada sesetengah akaun Workspace/KPM.
    imageUrls.push('https://lh3.googleusercontent.com/d/' + file.getId());
  });

  return imageUrls;
}

function readLaporanData(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const headers = getHeaders(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return values.map(function(row, index) {
    const students = safeJsonArray(getFieldValue(row, headers, 'students'));
    const teachers = safeJsonArray(getFieldValue(row, headers, 'teachers'));
    const images = normalizeImageArray(getFieldValue(row, headers, 'images'));

    const startDate = formatDateForClient(getFieldValue(row, headers, 'startDate'));
    const endDate = formatDateForClient(getFieldValue(row, headers, 'endDate'));
    const createdAt = formatDateForClient(getFieldValue(row, headers, 'createdAt'));

    const preparedBy = String(getFieldValue(row, headers, 'preparedBy') || '').trim();
    const preparedPosition = String(getFieldValue(row, headers, 'preparedPosition') || '').trim();

    return {
      id: String(getFieldValue(row, headers, 'id') || ('LAP-' + (index + 1))).trim(),
      startDate: startDate,
      endDate: endDate,
      dateStart: startDate,
      dateEnd: endDate,
      tarikh: startDate,
      name: String(getFieldValue(row, headers, 'name') || '').trim(),
      namaProgram: String(getFieldValue(row, headers, 'name') || '').trim(),
      level: String(getFieldValue(row, headers, 'level') || '').trim(),
      peringkat: String(getFieldValue(row, headers, 'level') || '').trim(),
      unit: String(getFieldValue(row, headers, 'unit') || '').trim(),
      category: String(getFieldValue(row, headers, 'category') || '').trim(),
      kategori: String(getFieldValue(row, headers, 'category') || '').trim(),
      achievement: String(getFieldValue(row, headers, 'achievement') || '').trim(),
      pencapaian: String(getFieldValue(row, headers, 'achievement') || '').trim(),
      obj: String(getFieldValue(row, headers, 'obj') || '').trim(),
      objective: String(getFieldValue(row, headers, 'obj') || '').trim(),
      swot: String(getFieldValue(row, headers, 'swot') || '').trim(),
      students: students,
      murid: students,
      teachers: teachers,
      guru: teachers,
      images: images,
      gambar: images,
      preparedBy: preparedBy,
      preparedPosition: preparedPosition,
      namaGuruSediakanLaporan: preparedBy,
      jawatanGuru: preparedPosition,
      timestamp: createdAt
    };
  }).filter(function(item) {
    return item.name || item.startDate || item.students.length || item.teachers.length;
  });
}

function readJawatanData(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const headers = getHeaders(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  function value(row, header) {
    const idx = findHeaderIndex(headers, [header]);
    return idx === -1 ? '' : row[idx];
  }

  return rows.map(function(row) {
    const created = value(row, 'Tarikh Rekod Dicipta');
    return {
      name: String(value(row, 'Nama Murid') || '').trim(),
      class: String(value(row, 'Kelas') || '').trim(),
      unit: String(value(row, 'Unit') || '').trim(),
      pecahanUnit: String(value(row, 'Pecahan Unit') || '').trim(),
      jawatan: String(value(row, 'Jawatan') || '').trim(),
      sesi: String(value(row, 'Tahun / Sesi') || '').trim(),
      timestamp: Object.prototype.toString.call(created) === '[object Date]' && !isNaN(created.getTime())
        ? created.getTime()
        : 0
    };
  }).filter(item => item.name || item.unit || item.jawatan);
}

function doPost(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();
    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (action === 'saveLaporan') {
      const sheet = ensureLaporanSheet(ss);

      const students = safeJsonArray(e.parameter.students);
      const teachers = safeJsonArray(e.parameter.teachers);
      const newImages = safeJsonArray(e.parameter.newImages);
      const existingImages = safeJsonArray(e.parameter.existingImages);
      const imageUrls = saveImagesToDrive(newImages, existingImages);

      const preparedBy =
        e.parameter.preparedBy ||
        e.parameter.namaGuruSediakanLaporan ||
        e.parameter.namaGuru ||
        '';

      const preparedPosition =
        e.parameter.preparedPosition ||
        e.parameter.jawatanGuru ||
        e.parameter.jawatanGuruSediakanLaporan ||
        '';

      appendLaporanByHeaders(sheet, {
        id: e.parameter.id || '',
        startDate: e.parameter.startDate || e.parameter.tarikh || '',
        endDate: e.parameter.endDate || '',
        name: e.parameter.name || e.parameter.namaProgram || '',
        level: e.parameter.level || e.parameter.peringkat || '',
        unit: e.parameter.unit || '',
        category: e.parameter.category || e.parameter.kategori || '',
        achievement: e.parameter.achievement || e.parameter.pencapaian || '',
        obj: e.parameter.obj || e.parameter.objective || '',
        swot: e.parameter.swot || '',
        students: students,
        teachers: teachers,
        images: imageUrls,
        preparedBy: preparedBy,
        preparedPosition: preparedPosition
      });

      SpreadsheetApp.flush();

      return jsonOutput({
        status: 'success',
        message: 'Laporan berjaya disimpan.',
        images: imageUrls
      });
    }

    if (action === 'saveJawatan') {
      const sheet = ensureJawatanSheet(ss);
      const headers = getHeaders(sheet);
      const records = safeJsonArray(e.parameter.records);

      records.forEach(function(rec) {
        const row = new Array(headers.length).fill('');
        const fieldMap = {
          'Tarikh Rekod Dicipta': new Date(),
          'Nama Murid': rec.name || rec.studentName || '',
          'Kelas': rec.class || rec.studentClass || '',
          'Unit': rec.unit || '',
          'Pecahan Unit': rec.pecahanUnit || '',
          'Jawatan': rec.jawatan || '',
          'Tahun / Sesi': rec.sesi || ''
        };

        Object.keys(fieldMap).forEach(function(header) {
          const idx = findHeaderIndex(headers, [header]);
          if (idx !== -1) row[idx] = fieldMap[header];
        });

        sheet.appendRow(row);
      });

      SpreadsheetApp.flush();
      return jsonOutput({ status: 'success', message: 'Rekod jawatan berjaya disimpan.' });
    }

    return jsonOutput({ status: 'error', message: 'Unknown action: ' + action });

  } catch (error) {
    console.error(error);
    return jsonOutput({
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  }
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();
    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (action === 'getData' || action === 'getLaporan' || action === '') {
      const sheetLaporan = ensureLaporanSheet(ss);
      const laporan = readLaporanData(sheetLaporan);

      let jawatan = [];
      const sheetJawatan = ss.getSheetByName(JAWATAN_SHEET);
      if (sheetJawatan) {
        ensureJawatanSheet(ss);
        jawatan = readJawatanData(sheetJawatan);
      }

      return jsonOutput({
        status: 'success',
        data: {
          laporan: laporan,
          jawatan: jawatan
        },
        laporan: laporan,
        jawatan: jawatan,
        meta: {
          laporanCount: laporan.length,
          jawatanCount: jawatan.length,
          source: 'Google Apps Script -> Google Sheets / Google Drive',
          generatedAt: new Date().toISOString()
        }
      });
    }

    if (action === 'ping') {
      return jsonOutput({
        status: 'success',
        message: 'Apps Script aktif.',
        sheetId: SHEET_ID,
        generatedAt: new Date().toISOString()
      });
    }

    return jsonOutput({ status: 'error', message: 'Unknown action: ' + action });

  } catch (error) {
    console.error(error);
    return jsonOutput({
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  }
}
