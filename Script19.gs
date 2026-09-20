/****************************************************
 * WEB INCOMING
 * HALAL + NON HALAL
 ****************************************************/

const CONFIG = {
  HALAL_ID: "1nTEOigckVl1Y8bksW8gm44PQt5WN-EGO1oMuz3nTJyI",
  NON_HALAL_ID: "17I5dw1umu3jzO_H7WbVo8VfB2updHeHxe3MXWqlj1nc",

  VENDOR_SHEET: "VENDOR",
  MASTER_SHEET: "MASTER",

  MAX_PRODUCTS: 5,

  BLOCKS: [
    { start: 25, end: 53 },
    { start: 80, end: 108 },
    { start: 135, end: 163 },
    { start: 190, end: 218 },
    { start: 245, end: 273 }
  ],

  DATA_COLUMNS: 13
};

/****************************************************
 * GET SPREADSHEET
 ****************************************************/
function getSpreadsheet(category, tanggal, createIfMissing) {
  category = normalizeCategory(category);
  if (!category) throw new Error("Kategori harus HALAL atau NON HALAL.");

  // Tanpa tanggal = file MASTER/template.
  // Dengan tanggal = file transaksi bulanan sesuai bulan/tahun.
  if (!tanggal) {
    if (category === "HALAL") return SpreadsheetApp.openById(CONFIG.HALAL_ID);
    if (category === "NON HALAL") return SpreadsheetApp.openById(CONFIG.NON_HALAL_ID);
  }

  const date = parseDate(tanggal);
  return getMonthlySpreadsheet(category, date, createIfMissing !== false);
}

/****************************************************
 * NORMALIZE CATEGORY
 ****************************************************/
function normalizeCategory(category) {
  if (!category) return "";
  const value = String(category).trim().toUpperCase();
  if (value === "HALAL") return "HALAL";
  if (value === "NON HALAL" || value === "NON-HALAL" || value === "NONHALAL") return "NON HALAL";
  return "";
}

/****************************************************
 * MONTHLY SPREADSHEET
 *
 * File CONFIG.HALAL_ID / NON_HALAL_ID menjadi MASTER.
 * File bulanan dibuat otomatis dengan copy seluruh file,
 * sehingga semua sheet seperti 01-31, I01-I31, VENDOR,
 * MASTER, dan sheet tambahan lainnya ikut tercopy.
 ****************************************************/

/****************************************************
 * ISI BULAN + TAHUN PADA TEMPLATE
 *
 * Hanya mengisi B22, B77, B132, B187, B242
 * pada setiap sheet tanggal 01-31.
 ****************************************************/
function setTemplateMonthLabels_(ss, date) {
  const months = [
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
  ];

  const monthLabel =
    months[date.getMonth()] + " " + date.getFullYear();

  const cells = ["B22", "B77", "B132", "B187", "B242"];

  for (let day = 1; day <= 31; day++) {
    const sheetName = String(day).padStart(2, "0");
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    cells.forEach(function(a1) {
      sheet.getRange(a1).setValue(': ' + monthLabel);
    });
  }
}

function getMonthlySpreadsheet(category, date, createIfMissing) {
  category = normalizeCategory(category);
  if (!category) throw new Error("Kategori tidak valid.");

  date = parseDate(date);

  const tz = Session.getScriptTimeZone();
  const monthName = [
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
  ][date.getMonth()];
  const year = date.getFullYear();

  const suffix = category === "HALAL" ? "HALAL" : "NON HALAL";
  const fileName = "FORM INCOMING BSD " + suffix + " " + monthName + " " + year;

  const props = PropertiesService.getScriptProperties();
  const key = "INCOMING_MONTHLY_" + category.replace(/\s+/g, "_") + "_" + year + "_" + String(date.getMonth() + 1).padStart(2, "0");
  const savedId = props.getProperty(key);

  if (savedId) {
    try {
      const ss = SpreadsheetApp.openById(savedId);
      setTemplateMonthLabels_(ss, date);
      return ss;
    } catch (err) {
      props.deleteProperty(key);
    }
  }

  const existing = DriveApp.getFilesByName(fileName);
  if (existing.hasNext()) {
    const file = existing.next();
    props.setProperty(key, file.getId());
    const ss = SpreadsheetApp.openById(file.getId());
    setTemplateMonthLabels_(ss, date);
    return ss;
  }

  // Jika hanya membaca/menampilkan data, JANGAN membuat file baru.
  // File baru hanya boleh dibuat ketika proses memang membutuhkan
  // spreadsheet transaksi, misalnya saat submit Incoming.
  if (createIfMissing === false) {
    return null;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // Cek lagi setelah lock agar dua request bersamaan tidak membuat
    // dua file bulan yang sama.
    const again = DriveApp.getFilesByName(fileName);
    if (again.hasNext()) {
      const file = again.next();
      props.setProperty(key, file.getId());
      const ss = SpreadsheetApp.openById(file.getId());
    setTemplateMonthLabels_(ss, date);
    return ss;
    }

    const masterId = category === "HALAL"
      ? CONFIG.HALAL_ID
      : CONFIG.NON_HALAL_ID;

    const masterFile = DriveApp.getFileById(masterId);
    const newFile = masterFile.makeCopy(fileName);

    const ss = SpreadsheetApp.openById(newFile.getId());

    // File hasil copy menjadi file transaksi bulan baru.
    // Bersihkan hanya area data transaksi pada sheet 01-31.
    // Sheet I01-I31, VENDOR, MASTER, dan sheet lainnya tidak disentuh.
    clearMonthlyTransactionData(ss);

    // Index transaksi bulan baru harus kosong.
    resetMonthlyIndex(ss);

    props.setProperty(key, newFile.getId());

    setTemplateMonthLabels_(ss, date);

    return ss;
  } finally {
    lock.releaseLock();
  }
}

/****************************************************
 * CLEAR DATA TRANSAKSI BULAN BARU
 ****************************************************/
function clearMonthlyTransactionData(ss) {
  for (let day = 1; day <= 31; day++) {
    const sheetName = String(day).padStart(2, "0");
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    CONFIG.BLOCKS.forEach(function(block) {
      sheet.getRange(
        block.start,
        1,
        block.end - block.start + 1,
        CONFIG.DATA_COLUMNS
      ).clearContent();
    });
  }
}

/****************************************************
 * RESET INDEX BULAN BARU
 ****************************************************/
function resetMonthlyIndex(ss) {
  const index = ss.getSheetByName("_INCOMING_INDEX");
  if (!index) return;

  const lastRow = index.getLastRow();
  if (lastRow >= 2) {
    index.getRange(2, 1, lastRow - 1, 13).clearContent();
  }
}

/****************************************************
 * DO GET
 ****************************************************/
function doGet(e) {
  try {
    const action = e?.parameter?.action || "ping";

    if (action === "ping") {
      return json({
        success: true,
        message: "WEB INCOMING AKTIF",
        time: new Date().toISOString()
      });
    }

    if (action === "vendor") {
      return json({
        success: true,
        data: getVendor(e.parameter.category)
      });
    }

    if (action === "master") {
      return json({
        success: true,
        data: getMaster(e.parameter.category)
      });
    }

    if (action === "history") {
      return json(getHistory(e.parameter.category, e.parameter.tanggal));
    }

    // TAMBAHAN: daftar file Spreadsheet bulanan yang benar-benar masih ada di Drive.
    if (action === "listSpreadsheets") {
      return json({
        success: true,
        data: listMonthlySpreadsheets(e.parameter.category)
      });
    }

    return json({
      success: false,
      message: "Action tidak dikenal."
    });
  } catch (err) {
    return json({
      success: false,
      message: err && err.message ? err.message : String(err)
    });
  }
}

/****************************************************
 * DO POST
 ****************************************************/
/****************************************************
 * LIST SPREADSHEET BULANAN
 * Hanya membaca file yang masih ada di Drive.
 ****************************************************/
function listMonthlySpreadsheets(category) {
  category = normalizeCategory(category);
  if (!category) throw new Error("Kategori harus HALAL atau NON HALAL.");

  const suffix = category === "HALAL" ? "HALAL" : "NON HALAL";
  const prefix = "FORM INCOMING BSD " + suffix + " ";
  const files = DriveApp.getFiles();
  const result = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!name.startsWith(prefix)) continue;

    const rest = name.slice(prefix.length).trim();
    const parts = rest.split(/\s+/);
    if (parts.length < 2) continue;

    const year = Number(parts[parts.length - 1]);
    const monthName = parts.slice(0, -1).join(" ").trim().toUpperCase();
    const monthIndex = getMonthIndex(monthName);
    if (monthIndex < 0 || !/^\d{4}$/.test(String(year))) continue;

    result.push({
      id: file.getId(),
      name: name,
      url: file.getUrl(),
      category: category,
      month: monthIndex + 1,
      monthName: monthName,
      year: year,
      label: monthName + " " + year
    });
  }

  result.sort(function(a, b) {
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });

  return result;
}

function getMonthIndex(monthName) {
  const months = [
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
  ];
  return months.indexOf(String(monthName || "").trim().toUpperCase());
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("POST data kosong.");
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action || "save";

    if (action === "save") return json(saveIncoming(data));
    if (action === "update") return json(updateIncoming(data));
    if (action === "delete") return json(deleteIncoming(data));

    throw new Error("Action tidak dikenal.");
  } catch (err) {
    return json({
      success: false,
      message: err && err.message ? err.message : String(err)
    });
  }
}

/****************************************************
 * SAVE INCOMING
 ****************************************************/
function saveIncoming(data) {
  validateIncoming(data);

  const category = normalizeCategory(data.category);
  data.noPo = String(data.noPo ?? data.noPO ?? "").trim();
  const date = parseDate(data.tanggal);
  const ss = getSpreadsheet(category, date);

  const sheetName = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd");
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) throw new Error("Sheet tanggal " + sheetName + " tidak ditemukan.");

  const products = normalizeProducts(data.products);
  if (!products.length) throw new Error("Minimal 1 produk.");
  if (products.length > CONFIG.MAX_PRODUCTS) throw new Error("Maksimal 5 produk.");

  const requiredRows = products.length + 1;
  const location = findAvailableBlock(sheet, requiredRows);

  if (!location) {
    throw new Error(
      "Semua lembar pada tanggal " + sheetName +
      " sudah tidak memiliki ruang yang cukup."
    );
  }

  const id = createTransactionId(date, category);

  products.forEach(function(product, index) {
    writeRow(sheet, location.start + index, {
      tanggal: date,
      noPo: data.noPo,
      vendor: data.vendor,
      product: product,
      suhuMobil: data.suhuMobil,
      kebersihanMobil: data.kebersihanMobil,
      firstRow: index === 0
    });
  });

  const separatorRow = location.start + products.length;
  clearRow(sheet, separatorRow);

  saveIndex(ss, {
    id: id,
    category: category,
    tanggal: formatDate(date),
    sheet: sheetName,
    startRow: location.start,
    endRow: location.start + products.length - 1,
    separatorRow: separatorRow,
    noPo: data.noPo,
    vendor: data.vendor,
    products: products,
    suhuMobil: data.suhuMobil,
    kebersihanMobil: data.kebersihanMobil,
    kualitasProduk: getTransactionQuality(data, products),
    tindakanKoreksi: getTransactionCorrection(data, products),
    keputusan: getTransactionDecision(data, products),
    pic: data.pic || ""
  });

  return {
    success: true,
    message: "Incoming berhasil disimpan.",
    id: id,
    category: category,
    sheet: sheetName,
    block: location.block,
    startRow: location.start,
    endRow: location.start + products.length - 1
  };
}

/****************************************************
 * CARI BLOK YANG CUKUP
 *
 * Setiap PO membutuhkan:
 * produk 1..N + 1 baris separator kosong.
 ****************************************************/
function findAvailableBlock(sheet, requiredRows) {
  /*
   * Separator hanya dianggap RESERVED jika transaksi yang
   * terkait masih benar-benar ada di sheet.
   *
   * Jika PO dihapus manual langsung dari Spreadsheet, data
   * pada sheet menjadi kosong walaupun record lama di
   * _INCOMING_INDEX masih ACTIVE. Dalam kondisi itu separator
   * lama TIDAK lagi diblokir, sehingga slot kosong dari awal
   * bisa dipakai kembali.
   */
  const reservedSeparators = new Set();

  // Baca blok sekali saja agar pengecekan transaksi lama
  // tidak membuat banyak getRange() berulang.
  const blockData = CONFIG.BLOCKS.map(function(block) {
    return {
      start: block.start,
      end: block.end,
      values: sheet.getRange(
        block.start, 1, block.end - block.start + 1, 11
      ).getDisplayValues()
    };
  });

  function transactionStillExists(startRow, endRow) {
    if (!startRow || !endRow || endRow < startRow) return false;

    for (let b = 0; b < blockData.length; b++) {
      const block = blockData[b];
      const from = Math.max(startRow, block.start);
      const to = Math.min(endRow, block.end);

      if (from > to) continue;

      for (let rowNum = from; rowNum <= to; rowNum++) {
        const row = block.values[rowNum - block.start];
        if (row && row.some(function(cell) {
          return String(cell).trim() !== "";
        })) {
          return true;
        }
      }
    }

    return false;
  }

  try {
    const ss = sheet.getParent();
    const index = ss.getSheetByName("_INCOMING_INDEX");

    if (index && index.getLastRow() >= 2) {
      const rows = index.getRange(
        2, 1, index.getLastRow() - 1, 13
      ).getDisplayValues();

      rows.forEach(function(row) {
        const status = String(row[12] || "").trim().toUpperCase();
        const sheetName = String(row[3] || "").trim();
        const startRow = Number(row[4]) || 0;
        const endRow = Number(row[5]) || 0;
        const separatorRow = Number(row[6]) || 0;

        if (
          status === "ACTIVE" &&
          sheetName === sheet.getName() &&
          separatorRow > 0 &&
          transactionStillExists(startRow, endRow)
        ) {
          reservedSeparators.add(separatorRow);
        }
      });
    }
  } catch (err) {
    // Jika index belum ada / gagal dibaca, lanjutkan
    // dengan pengecekan normal berdasarkan isi sheet.
  }

  for (let b = 0; b < CONFIG.BLOCKS.length; b++) {
    const block = CONFIG.BLOCKS[b];
    const capacity = block.end - block.start + 1;

    if (requiredRows > capacity) continue;

    const values = blockData[b].values;

    for (let i = 0; i <= capacity - requiredRows; i++) {
      const candidateStart = block.start + i;

      // Jangan pernah memakai separator PO yang masih aktif
      // sebagai awal PO baru.
      if (reservedSeparators.has(candidateStart)) continue;

      // Pastikan selalu ada 1 baris kosong sebelum PO baru
      // jika PO tersebut bukan PO pertama dalam blok.
      if (candidateStart > block.start) {
        const previousRow = values[i - 1];
        const previousUsed = previousRow.some(function(cell) {
          return String(cell).trim() !== "";
        });

        if (previousUsed) continue;
      }

      let available = true;

      for (let j = 0; j < requiredRows; j++) {
        const currentRow = candidateStart + j;

        // Separator PO lama yang masih aktif tidak boleh ditimpa.
        if (reservedSeparators.has(currentRow)) {
          available = false;
          break;
        }

        const row = values[i + j];
        const used = row.some(function(cell) {
          return String(cell).trim() !== "";
        });

        if (used) {
          available = false;
          break;
        }
      }

      if (available) {
        return {
          block: b + 1,
          start: candidateStart,
          end: block.end
        };
      }
    }
  }

  return null;
}

/****************************************************
 * WRITE ROW
 ****************************************************/
function writeRow(sheet, row, data) {
  const p = data.product;
  const firstRow = data.firstRow !== false;

  const values = [[
    firstRow ? data.tanggal : "",
    firstRow ? (data.noPo || "") : "",
    firstRow ? (data.vendor || "") : "",
    p.name || "",
    p.qty || "",
    firstRow ? (data.suhuMobil || "") : "",
    firstRow ? (data.kebersihanMobil || "") : "",
    p.suhuProduk || "",
    firstRow ? (p.kualitasProduk || "") : "",
    firstRow ? (p.tindakanKoreksi || "") : "",
    firstRow ? (p.keputusan || "") : "",
    "",
    ""
  ]];

  sheet.getRange(row, 1, 1, CONFIG.DATA_COLUMNS).setValues(values);
}

/****************************************************
 * CLEAR ROW
 ****************************************************/
function clearRow(sheet, row) {
  sheet.getRange(row, 1, 1, CONFIG.DATA_COLUMNS).clearContent();
}

/****************************************************
 * INDEX SHEET
 ****************************************************/
function getIndexSheet(ss) {
  const name = "_INCOMING_INDEX";
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 13).setValues([[
      "ID",
      "CATEGORY",
      "TANGGAL",
      "SHEET",
      "START_ROW",
      "END_ROW",
      "SEPARATOR_ROW",
      "NO_PO",
      "VENDOR",
      "PRODUCTS",
      "CREATED",
      "UPDATED",
      "STATUS"
    ]]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/****************************************************
 * SAVE INDEX
 *
 * Kolom PRODUCTS menyimpan JSON yang juga memuat
 * metadata transaksi. Ini menjaga struktur index
 * lama tetap 13 kolom.
 ****************************************************/
function saveIndex(ss, data) {
  const sheet = getIndexSheet(ss);

  const payload = buildHistoryPayload(data);

  sheet.appendRow([
    data.id,
    data.category,
    data.tanggal,
    data.sheet,
    data.startRow,
    data.endRow,
    data.separatorRow,
    data.noPo || "",
    data.vendor || "",
    JSON.stringify(payload),
    new Date(),
    new Date(),
    "ACTIVE"
  ]);
}

/****************************************************
 * BUILD HISTORY PAYLOAD
 ****************************************************/
function buildHistoryPayload(data) {
  return {
    products: Array.isArray(data.products) ? data.products : [],
    suhuMobil: data.suhuMobil || "",
    kebersihanMobil: data.kebersihanMobil || "",
    kualitasProduk: data.kualitasProduk || "",
    tindakanKoreksi: data.tindakanKoreksi || "",
    keputusan: data.keputusan || "",
    pic: data.pic || ""
  };
}

/****************************************************
 * READ HISTORY PAYLOAD
 *
 * Mendukung index lama yang menyimpan PRODUCTS
 * sebagai array langsung.
 ****************************************************/
function parseHistoryPayload(raw) {
  try {
    const parsed = JSON.parse(raw || "[]");

    if (Array.isArray(parsed)) {
      return {
        products: parsed,
        suhuMobil: "",
        kebersihanMobil: "",
        kualitasProduk: "",
        tindakanKoreksi: "",
        keputusan: "",
        pic: ""
      };
    }

    return {
      products: Array.isArray(parsed.products) ? parsed.products : [],
      suhuMobil: parsed.suhuMobil || "",
      kebersihanMobil: parsed.kebersihanMobil || "",
      kualitasProduk: parsed.kualitasProduk || "",
      tindakanKoreksi: parsed.tindakanKoreksi || "",
      keputusan: parsed.keputusan || "",
      pic: parsed.pic || ""
    };
  } catch (e) {
    return {
      products: [],
      suhuMobil: "",
      kebersihanMobil: "",
      kualitasProduk: "",
      tindakanKoreksi: "",
      keputusan: "",
      pic: ""
    };
  }
}

/****************************************************
 * HISTORY
 ****************************************************/
function getHistory(category, tanggal) {
  category = normalizeCategory(category);
  if (!category) throw new Error("Kategori wajib dipilih.");

  const filterDate = tanggal ? formatDate(parseDate(tanggal)) : formatDate(new Date());
  const ss = getSpreadsheet(category, parseDate(filterDate), false);

  if (!ss) {
    return { success: true, category: category, data: [] };
  }

  const index = getIndexSheet(ss);
  const indexMap = {};
  const indexRowMap = {};
  const lastIndexRow = index.getLastRow();

  if (lastIndexRow >= 2) {
    const indexValues = index.getRange(2, 1, lastIndexRow - 1, 13).getDisplayValues();

    indexValues.forEach(function(row) {
      const id = String(row[0] || '').trim();
      const status = String(row[12] || '').trim().toUpperCase();
      if (!id || status === 'DELETED') return;
      if (filterDate && String(row[2] || '').trim() !== filterDate) return;

      const item = {
        id: id,
        category: row[1],
        tanggal: row[2],
        sheet: String(row[3] || '').trim(),
        block: getBlockNumber(Number(row[4]) || 0),
        startRow: Number(row[4]) || 0,
        endRow: Number(row[5]) || 0,
        separatorRow: Number(row[6]) || 0,
        noPo: row[7],
        vendor: row[8],
        meta: parseHistoryPayload(row[9]),
        created: row[10],
        updated: row[11]
      };

      // Primary key: sheet + startRow + PO (format lama/baru).
      const key = item.sheet + '|' + item.startRow + '|' + String(item.noPo || '').trim();
      indexMap[key] = item;

      // Secondary key: sheet + startRow.
      // Ini penting bila format PO di sheet berubah (mis. leading zero/apostrophe)
      // sehingga key lama tidak persis sama, tetapi posisi transaksi tetap sama.
      const rowKey = item.sheet + '|' + item.startRow;
      indexRowMap[rowKey] = item;
    });
  }

  const sheetNames = filterDate
    ? [Utilities.formatDate(parseDate(filterDate), Session.getScriptTimeZone(), 'dd')]
    : CONFIG.BLOCKS.map(function(_, i) { return String(i + 1).padStart(2, '0'); });

  const result = [];

  sheetNames.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const groups = readPOGroupsFromSheet(sheet);

    groups.forEach(function(group) {
      if (!group.noPo && !group.products.length) return;

      const normalizedSheet = String(sheetName).trim();
      const normalizedStart = String(group.startRow).trim();
      const normalizedPo = String(group.noPo || '').trim();

      const key = normalizedSheet + '|' + normalizedStart + '|' + normalizedPo;

      // Pertama cari dengan key lengkap. Jika tidak ketemu, cari berdasarkan
      // posisi sheet + startRow agar transaksi web lama tidak berubah menjadi MANUAL.
      const indexed = indexMap[key] || indexRowMap[normalizedSheet + '|' + normalizedStart];

      if (indexed) {
        const meta = indexed.meta || {};
        const sheetProducts = Array.isArray(group.products) ? group.products : [];
        const metaProducts = Array.isArray(meta.products) ? meta.products : [];

        // Produk utama tetap mengikuti sheet agar perubahan qty/suhu terlihat.
        // Jika sheet tidak memiliki produk lengkap, gunakan data index sebagai fallback.
        const products = sheetProducts.length ? sheetProducts : metaProducts;

        result.push({
          id: indexed.id,
          category: category,
          tanggal: indexed.tanggal || group.tanggal || filterDate,
          sheet: indexed.sheet || normalizedSheet,
          block: getBlockNumber(group.startRow || indexed.startRow),
          startRow: indexed.startRow || group.startRow,
          endRow: indexed.endRow || group.endRow,
          separatorRow: indexed.separatorRow || group.separatorRow,
          noPo: group.noPo || indexed.noPo,
          vendor: group.vendor || indexed.vendor,
          products: products,
          suhuMobil: group.suhuMobil || meta.suhuMobil || '',
          kebersihanMobil: group.kebersihanMobil || meta.kebersihanMobil || '',
          kualitasProduk: group.kualitasProduk || meta.kualitasProduk || '',
          tindakanKoreksi: group.tindakanKoreksi || meta.tindakanKoreksi || '',
          keputusan: group.keputusan || meta.keputusan || '',
          pic: meta.pic || '',
          created: indexed.created,
          updated: indexed.updated
        });

        // Hapus kedua kemungkinan key agar tidak dipakai lagi sebagai unmatched index.
        const fullKey = indexed.sheet + '|' + indexed.startRow + '|' + String(indexed.noPo || '').trim();
        const rowKey = indexed.sheet + '|' + indexed.startRow;
        delete indexMap[fullKey];
        delete indexRowMap[rowKey];
        return;
      }

      // Benar-benar tidak ada di index: ini memang data manual dari Spreadsheet.
      const manualId =
        'MANUAL-' +
        category.replace(/\s+/g, '') + '-' +
        normalizedSheet + '-' +
        group.startRow + '-' +
        String(group.noPo || 'NOPo').replace(/\W/g, '');

      result.push({
        id: manualId,
        category: category,
        tanggal: group.tanggal,
        sheet: normalizedSheet,
        block: getBlockNumber(group.startRow),
        startRow: group.startRow,
        endRow: group.endRow,
        separatorRow: group.separatorRow,
        noPo: group.noPo,
        vendor: group.vendor,
        products: group.products,
        suhuMobil: group.suhuMobil,
        kebersihanMobil: group.kebersihanMobil,
        kualitasProduk: group.kualitasProduk,
        tindakanKoreksi: group.tindakanKoreksi,
        keputusan: group.keputusan,
        pic: '',
        created: '',
        updated: ''
      });
    });
  });

  result.sort(function(a, b) {
    if (a.sheet !== b.sheet) return String(a.sheet).localeCompare(String(b.sheet));
    return Number(a.startRow || 0) - Number(b.startRow || 0);
  });

  result.reverse();

  return { success: true, category: category, data: result };
}

/****************************************************
 * BACA GROUP PO LANGSUNG DARI SHEET
 ****************************************************/
function readPOGroupsFromSheet(sheet) {
  const groups = [];

  CONFIG.BLOCKS.forEach(function(block, blockIndex) {
    const values = sheet.getRange(
      block.start,
      1,
      block.end - block.start + 1,
      11
    ).getDisplayValues();

    let current = null;

    values.forEach(function(row, offset) {
      const rowNumber = block.start + offset;

      const cells = row.map(function(v) {
        return String(v || "").trim();
      });

      const hasAny = cells.some(function(v) {
        return v !== "";
      });

      const hasProduct = cells[3] !== "";
      const hasPO = cells[1] !== "";

      // Baris kosong = separator.
      if (!hasAny) {
        if (current) {
          current.endRow = rowNumber - 1;
          current.separatorRow = rowNumber;
          groups.push(current);
          current = null;
        }
        return;
      }

      // Jika ada PO baru, tutup group sebelumnya terlebih dahulu.
      if (hasPO) {
        if (current) {
          current.endRow = rowNumber - 1;
          current.separatorRow = rowNumber;
          groups.push(current);
          current = null;
        }

        current = {
          block: blockIndex + 1,
          startRow: rowNumber,
          endRow: rowNumber,
          separatorRow: 0,
          tanggal: cells[0],
          noPo: cells[1],
          vendor: cells[2],
          suhuMobil: cells[5],
          kebersihanMobil: cells[6],
          kualitasProduk: cells[8],
          tindakanKoreksi: cells[9],
          keputusan: cells[10],
          products: []
        };
      }

      // Jika belum ada header PO, abaikan baris.
      if (!current) return;

      // Hanya baris dengan nama produk yang masuk sebagai product.
      if (hasProduct) {
        current.products.push({
          code: "",
          name: cells[3],
          qty: cells[4],
          suhuProduk: cells[7],
          kualitasProduk: cells[8] || current.kualitasProduk || "",
          tindakanKoreksi: cells[9] || current.tindakanKoreksi || "",
          keputusan: cells[10] || current.keputusan || ""
        });

        current.endRow = rowNumber;
      }
    });

    // Tutup group jika block berakhir tanpa separator.
    if (current) {
      current.endRow = block.end;
      current.separatorRow = block.end + 1;
      groups.push(current);
    }
  });

  return groups;
}

/****************************************************
 * BLOCK NUMBER
 ****************************************************/
function getBlockNumber(row) {
  row = Number(row) || 0;

  for (let i = 0; i < CONFIG.BLOCKS.length; i++) {
    if (row >= CONFIG.BLOCKS[i].start && row <= CONFIG.BLOCKS[i].end) {
      return i + 1;
    }
  }

  return 0;
}

/****************************************************
 * UPDATE
 ****************************************************/
function updateIncoming(data) {
  if (!data.id) throw new Error("ID transaksi tidak ditemukan.");

  data.noPo = String(data.noPo ?? data.noPO ?? "").trim();
  validateIncoming(data);

  const category = normalizeCategory(data.category);
  if (!category) throw new Error("Kategori tidak valid.");

  const date = parseDate(data.tanggal);
  const ss = getSpreadsheet(category, date);
  const index = getIndexSheet(ss);
  const indexRow = findIndexRow(index, data.id);

  if (!indexRow) throw new Error("Transaksi tidak ditemukan.");

  const storedCategory = index.getRange(indexRow, 2).getDisplayValue();
  if (normalizeCategory(storedCategory) !== category) {
    throw new Error("Kategori transaksi tidak cocok.");
  }

  const sheetName = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd");
  const newSheet = ss.getSheetByName(sheetName);

  if (!newSheet) throw new Error("Sheet tanggal " + sheetName + " tidak ditemukan.");

  const products = normalizeProducts(data.products);
  if (!products.length) throw new Error("Minimal 1 produk.");
  if (products.length > CONFIG.MAX_PRODUCTS) throw new Error("Maksimal 5 produk.");

  const oldSheetName = index.getRange(indexRow, 4).getDisplayValue();
  const oldStart = Number(index.getRange(indexRow, 5).getValue()) || 0;
  const oldEnd = Number(index.getRange(indexRow, 6).getValue()) || 0;
  const oldSeparator = Number(index.getRange(indexRow, 7).getValue()) || 0;

  let location = null;

  // Pakai lokasi lama hanya jika jumlah produk baru tetap muat.
  if (
    oldSheetName === sheetName &&
    oldStart >= 1 &&
    oldEnd >= oldStart &&
    products.length <= (oldEnd - oldStart + 1)
  ) {
    location = {
      start: oldStart,
      block: getBlockNumber(oldStart),
      end: oldEnd
    };
  }

  // Jika lokasi lama tidak cukup, cari lokasi baru.
  if (!location) {
    location = findAvailableBlock(newSheet, products.length + 1);

    if (!location) {
      throw new Error("Tidak ada ruang untuk transaksi.");
    }
  }

  const oldSheet = ss.getSheetByName(oldSheetName);

  if (oldSheet) {
    const sameLocation =
      oldSheetName === sheetName &&
      location.start === oldStart;

    if (!sameLocation) {
      for (let r = oldStart; r <= oldEnd; r++) clearRow(oldSheet, r);
      if (oldSeparator >= oldStart) clearRow(oldSheet, oldSeparator);
    } else {
      // Bersihkan seluruh area lama agar pengurangan jumlah produk
      // tidak meninggalkan baris produk lama.
      const oldClearEnd = Math.max(oldEnd, oldSeparator);
      for (let r = oldStart; r <= oldClearEnd; r++) clearRow(oldSheet, r);
    }
  }

  // Bersihkan area tujuan.
  const targetClearEnd = Math.min(
    location.start + products.length,
    CONFIG.BLOCKS[CONFIG.BLOCKS.length - 1].end
  );
  for (let r = location.start; r <= targetClearEnd; r++) clearRow(newSheet, r);

  products.forEach(function(product, i) {
    writeRow(newSheet, location.start + i, {
      tanggal: date,
      noPo: data.noPo,
      vendor: data.vendor,
      suhuMobil: data.suhuMobil,
      kebersihanMobil: data.kebersihanMobil,
      product: product,
      firstRow: i === 0
    });
  });

  const separator = location.start + products.length;
  clearRow(newSheet, separator);

  const createdValue = index.getRange(indexRow, 11).getValue();

  const payload = buildHistoryPayload({
    products: products,
    suhuMobil: data.suhuMobil,
    kebersihanMobil: data.kebersihanMobil,
    kualitasProduk: getTransactionQuality(data, products),
    tindakanKoreksi: getTransactionCorrection(data, products),
    keputusan: getTransactionDecision(data, products),
    pic: data.pic || ""
  });

  index.getRange(indexRow, 3, 1, 11).setValues([[
    formatDate(date),
    sheetName,
    location.start,
    location.start + products.length - 1,
    separator,
    data.noPo || "",
    data.vendor || "",
    JSON.stringify(payload),
    createdValue,
    new Date(),
    "ACTIVE"
  ]]);

  return {
    success: true,
    message: "Data berhasil diperbarui.",
    id: data.id,
    category: category,
    sheet: sheetName,
    block: location.block,
    startRow: location.start,
    endRow: location.start + products.length - 1
  };
}

/****************************************************
 * DELETE
 ****************************************************/
function deleteIncoming(data) {
  if (!data.id) throw new Error("ID transaksi tidak ditemukan.");

  const category = normalizeCategory(data.category);
  if (!category) throw new Error("Kategori wajib dipilih.");

  const idText = String(data.id).trim();
  const idMatch = idText.match(/^INC-(?:H|NH)-(\d{8})-/i);
  if (!idMatch) throw new Error("ID transaksi tidak valid.");

  const idDateText = idMatch[1];
  const idDate = parseDate(
    idDateText.substring(0, 4) + "-" +
    idDateText.substring(4, 6) + "-" +
    idDateText.substring(6, 8)
  );

  const ss = getSpreadsheet(category, idDate);
  const index = getIndexSheet(ss);
  const row = findIndexRow(index, data.id);

  if (!row) throw new Error("Transaksi tidak ditemukan.");

  const storedCategory = index.getRange(row, 2).getDisplayValue();
  if (normalizeCategory(storedCategory) !== category) {
    throw new Error("Kategori transaksi tidak cocok.");
  }

  const sheetName = index.getRange(row, 4).getDisplayValue();
  const start = Number(index.getRange(row, 5).getValue()) || 0;
  const end = Number(index.getRange(row, 6).getValue()) || 0;
  const separator = Number(index.getRange(row, 7).getValue()) || 0;

  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    for (let r = start; r <= end; r++) clearRow(sheet, r);
    if (separator >= start) clearRow(sheet, separator);
  }

  index.getRange(row, 13).setValue("DELETED");
  index.getRange(row, 12).setValue(new Date());

  return {
    success: true,
    message: "Data berhasil dihapus.",
    id: data.id,
    category: category
  };
}

/****************************************************
 * FIND INDEX ROW
 ****************************************************/
function findIndexRow(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(id).trim()) return i + 2;
  }

  return null;
}

/****************************************************
 * VENDOR
 * A = Kode Vendor
 * B = Nama Supplier
 ****************************************************/
function getVendor(category) {
  const ss = getSpreadsheet(category);
  const sheet = ss.getSheetByName(CONFIG.VENDOR_SHEET);

  if (!sheet) throw new Error("Sheet VENDOR tidak ditemukan.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();

  return values
    .filter(function(row) {
      return String(row[0]).trim() !== "" || String(row[1]).trim() !== "";
    })
    .map(function(row) {
      return [
        String(row[0] || "").trim(),
        String(row[1] || "").trim()
      ];
    });
}

/****************************************************
 * MASTER
 * A = Barcode
 * B = Article
 * C = Nama Produk
 ****************************************************/
function getMaster(category) {
  const ss = getSpreadsheet(category);
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);

  if (!sheet) throw new Error("Sheet MASTER tidak ditemukan.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();

  return values
    .filter(function(row) {
      return (
        String(row[0]).trim() !== "" ||
        String(row[1]).trim() !== "" ||
        String(row[2]).trim() !== ""
      );
    })
    .map(function(row) {
      return [
        String(row[0] || "").trim(),
        String(row[1] || "").trim(),
        String(row[2] || "").trim()
      ];
    });
}

/****************************************************
 * NORMALIZE PRODUCTS
 ****************************************************/
function normalizeProducts(products) {
  if (!Array.isArray(products)) return [];

  return products
    .filter(function(p) {
      return p && (p.name || p.nama || p.namaProduk || p.product);
    })
    .slice(0, CONFIG.MAX_PRODUCTS)
    .map(function(p) {
      return {
        code: p.code || p.kode || p.barcode || p.article || "",
        name: p.name || p.nama || p.namaProduk || p.product || "",
        qty: p.qty ?? p.quantity ?? "",
        suhuProduk: p.suhuProduk ?? p.suhu ?? "",
        kualitasProduk: p.kualitasProduk || p.kualitas || "",
        tindakanKoreksi: p.tindakanKoreksi || "",
        keputusan: p.keputusan || ""
      };
    });
}

/****************************************************
 * TRANSACTION-LEVEL HELPERS
 ****************************************************/
function getTransactionQuality(data, products) {
  if (data.kualitasProduk) return data.kualitasProduk;
  return products.length ? (products[0].kualitasProduk || "") : "";
}

function getTransactionCorrection(data, products) {
  if (data.tindakanKoreksi) return data.tindakanKoreksi;
  return products.length ? (products[0].tindakanKoreksi || "") : "";
}

function getTransactionDecision(data, products) {
  if (data.keputusan) return data.keputusan;
  return products.length ? (products[0].keputusan || "") : "";
}

/****************************************************
 * VALIDATION
 ****************************************************/
function validateIncoming(data) {
  if (!data) throw new Error("Data kosong.");
  if (!normalizeCategory(data.category)) throw new Error("Kategori wajib dipilih.");
  if (!data.tanggal) throw new Error("Tanggal wajib diisi.");
  const noPo = String(data.noPo ?? data.noPO ?? "").trim();
  if (!noPo) throw new Error("Nomor PO wajib diisi.");
  if (!/^\d+$/.test(noPo)) throw new Error("Nomor PO harus berupa angka saja.");
  data.noPo = noPo;
  if (!data.vendor) throw new Error("Vendor wajib diisi.");

  if (
    data.suhuMobil === undefined ||
    data.suhuMobil === null ||
    String(data.suhuMobil).trim() === ""
  ) {
    throw new Error("Suhu mobil wajib diisi.");
  }

  if (!data.kebersihanMobil) throw new Error("Kebersihan mobil wajib dipilih.");

  if (!Array.isArray(data.products) || !data.products.length) {
    throw new Error("Minimal 1 produk wajib diisi.");
  }

  if (data.products.length > CONFIG.MAX_PRODUCTS) {
    throw new Error("Maksimal 5 produk.");
  }
}

/****************************************************
 * CREATE ID
 ****************************************************/
function createTransactionId(date, category) {
  const prefix = category === "HALAL" ? "H" : "NH";
  const datePart = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  const random = Utilities.getUuid()
    .replace(/-/g, "")
    .substring(0, 8)
    .toUpperCase();

  return "INC-" + prefix + "-" + datePart + "-" + random;
}

/****************************************************
 * DATE
 ****************************************************/
function parseDate(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") return value;

  const text = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const p = text.split("-");
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    const p = text.split("/");
    return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  }

  const date = new Date(text);
  if (isNaN(date.getTime())) throw new Error("Tanggal tidak valid.");

  return date;
}

/****************************************************
 * FORMAT DATE
 ****************************************************/
function formatDate(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

/****************************************************
 * JSON
 ****************************************************/
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/****************************************************
 * TEST SPREADSHEET
 ****************************************************/
function testHalal() {
  const ss = getSpreadsheet("HALAL");
  Logger.log(ss.getName());
}

function testNonHalal() {
  const ss = getSpreadsheet("NON HALAL");
  Logger.log(ss.getName());
}

/****************************************************
 * TEST VENDOR
 ****************************************************/
function testVendorHalal() {
  const data = getVendor("HALAL");
  Logger.log("Jumlah vendor: " + data.length);
  if (data.length) Logger.log(JSON.stringify(data.slice(0, 10), null, 2));
}

function testVendorNonHalal() {
  const data = getVendor("NON HALAL");
  Logger.log("Jumlah vendor: " + data.length);
  if (data.length) Logger.log(JSON.stringify(data.slice(0, 10), null, 2));
}

/****************************************************
 * TEST MASTER
 ****************************************************/
function testMasterHalal() {
  const data = getMaster("HALAL");
  Logger.log("Jumlah master: " + data.length);
  if (data.length) Logger.log(JSON.stringify(data.slice(0, 10), null, 2));
}

function testMasterNonHalal() {
  const data = getMaster("NON HALAL");
  Logger.log("Jumlah master: " + data.length);
  if (data.length) Logger.log(JSON.stringify(data.slice(0, 10), null, 2));
}

/****************************************************
 * TEST PENEMPATAN PO
 * TIDAK MENULIS / TIDAK MENGUBAH DATA
 ****************************************************/
function testPlacement() {
  const testSheet = "01";
  const categories = ["HALAL", "NON HALAL"];
  const jumlahProduk = 5;
  const requiredRows = jumlahProduk + 1;

  Logger.log("====================================");
  Logger.log("TEST PENEMPATAN PO");
  Logger.log("Jumlah produk: " + jumlahProduk);
  Logger.log("Kebutuhan baris: " + requiredRows);
  Logger.log("Sheet tanggal: " + testSheet);
  Logger.log("====================================");

  categories.forEach(function(category) {
    Logger.log("------------------------------------");
    Logger.log("CATEGORY: " + category);

    const ss = getSpreadsheet(category);
    Logger.log("Spreadsheet: " + ss.getName());

    const sheet = ss.getSheetByName(testSheet);

    if (!sheet) {
      Logger.log("❌ Sheet " + testSheet + " TIDAK DITEMUKAN");
      return;
    }

    Logger.log("✅ Sheet " + testSheet + " ditemukan");

    CONFIG.BLOCKS.forEach(function(block, index) {
      const capacity = block.end - block.start + 1;
      Logger.log(
        "Lembar " + (index + 1) +
        ": baris " + block.start + "-" + block.end +
        " | kapasitas " + capacity
      );
    });

    const location = findAvailableBlock(sheet, requiredRows);

    if (!location) {
      Logger.log("❌ TIDAK ADA RUANG YANG CUKUP");
    } else {
      Logger.log("✅ RUANG DITEMUKAN");
      Logger.log("Lembar: " + location.block);
      Logger.log("Mulai baris: " + location.start);
      Logger.log("Baris produk terakhir: " + (location.start + jumlahProduk - 1));
      Logger.log("Baris pemisah: " + (location.start + jumlahProduk));
    }
  });

  Logger.log("====================================");
  Logger.log("TEST SELESAI");
  Logger.log("TIDAK ADA DATA YANG DITULIS.");
  Logger.log("====================================");
}

/****************************************************
 * TEST SAVE INCOMING
 * MENULIS DATA TEST KE HALAL
 ****************************************************/
function testSaveIncoming() {
  const testData = {
    category: "HALAL",
    tanggal: "2026-09-19",
    noPO: "TEST-INCOMING-001",
    vendor: "TEST VENDOR",
    suhuMobil: "0,0",
    kebersihanMobil: "OK",
    kualitasProduk: "OKE",
    tindakanKoreksi: "-",
    keputusan: "DITERIMA",
    pic: "TEST",
    products: [{
      code: "TEST001",
      name: "TEST PRODUK",
      qty: 1,
      suhuProduk: "0,0",
      kualitasProduk: "OKE",
      tindakanKoreksi: "-",
      keputusan: "DITERIMA"
    }]
  };

  const result = saveIncoming(testData);
  Logger.log(JSON.stringify(result, null, 2));
}

/****************************************************
 * HAPUS DATA TEST
 ****************************************************/
function deleteTestIncoming() {
  const ss = getSpreadsheet("HALAL");
  const index = getIndexSheet(ss);
  const lastRow = index.getLastRow();

  if (lastRow < 2) {
    Logger.log("Tidak ada data test.");
    return;
  }

  const values = index.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
  let found = false;

  for (let i = 0; i < values.length; i++) {
    const noPO = values[i][7];

    if (noPO === "TEST-INCOMING-001") {
      const row = i + 2;
      const sheetName = values[i][3];
      const start = Number(values[i][4]);
      const end = Number(values[i][5]);
      const separator = Number(values[i][6]);
      const sheet = ss.getSheetByName(sheetName);

      if (sheet) {
        for (let r = start; r <= end; r++) clearRow(sheet, r);
        if (separator >= start) clearRow(sheet, separator);
      }

      index.getRange(row, 13).setValue("DELETED");
      index.getRange(row, 12).setValue(new Date());

      Logger.log("Data TEST berhasil dihapus.");
      Logger.log("PO: TEST-INCOMING-001");
      found = true;
      break;
    }
  }

  if (!found) Logger.log("Data TEST-INCOMING-001 tidak ditemukan.");
}
