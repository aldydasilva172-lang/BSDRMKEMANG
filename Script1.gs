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

function getSpreadsheet(category) {
  category = normalizeCategory(category);

  if (category === "HALAL") {
    return SpreadsheetApp.openById(CONFIG.HALAL_ID);
  }

  if (category === "NON HALAL") {
    return SpreadsheetApp.openById(CONFIG.NON_HALAL_ID);
  }

  throw new Error("Kategori harus HALAL atau NON HALAL.");
}


/****************************************************
 * NORMALIZE CATEGORY
 ****************************************************/

function normalizeCategory(category) {
  if (!category) return "";

  const value = String(category).trim().toUpperCase();

  if (value === "HALAL") return "HALAL";

  if (
    value === "NON HALAL" ||
    value === "NON-HALAL" ||
    value === "NONHALAL"
  ) {
    return "NON HALAL";
  }

  return "";
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
      return json(getHistory(e.parameter.category));
    }

    return json({
      success: false,
      message: "Action tidak dikenal."
    });

  } catch (err) {
    return json({
      success: false,
      message: err.message
    });
  }
}


/****************************************************
 * DO POST
 ****************************************************/

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("POST data kosong.");
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action || "save";

    if (action === "save") {
      return json(saveIncoming(data));
    }

    if (action === "update") {
      return json(updateIncoming(data));
    }

    if (action === "delete") {
      return json(deleteIncoming(data));
    }

    throw new Error("Action tidak dikenal.");

  } catch (err) {
    return json({
      success: false,
      message: err.message
    });
  }
}


/****************************************************
 * SAVE INCOMING
 ****************************************************/

function saveIncoming(data) {
  validateIncoming(data);

  const category = normalizeCategory(data.category);
  const ss = getSpreadsheet(category);
  const date = parseDate(data.tanggal);

  const sheetName = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "dd"
  );

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Sheet tanggal " + sheetName + " tidak ditemukan.");
  }

  const products = normalizeProducts(data.products);

  if (!products.length) {
    throw new Error("Minimal 1 produk.");
  }

  if (products.length > CONFIG.MAX_PRODUCTS) {
    throw new Error("Maksimal 5 produk.");
  }

  const requiredRows = products.length + 1;

  const location = findAvailableBlock(sheet, requiredRows);

  if (!location) {
    throw new Error(
      "Semua lembar pada tanggal " +
      sheetName +
      " sudah tidak memiliki ruang yang cukup."
    );
  }

  const id = createTransactionId(date, category);

  products.forEach(function(product, index) {
    writeRow(sheet, location.start + index, {
      id: id,
      tanggal: date,
      noPO: data.noPO,
      vendor: data.vendor,
      product: product,
      suhuMobil: data.suhuMobil,
      kebersihanMobil: data.kebersihanMobil,
      category: category
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
    noPO: data.noPO,
    vendor: data.vendor,
    products: products
  });

  return {
    success: true,
    message: "Incoming berhasil disimpan.",
    id: id,
    category: category,
    sheet: sheetName,
    startRow: location.start,
    endRow: location.start + products.length - 1
  };
}


/****************************************************
 * CARI BLOK YANG CUKUP
 ****************************************************/

function findAvailableBlock(sheet, requiredRows) {
  for (let b = 0; b < CONFIG.BLOCKS.length; b++) {
    const block = CONFIG.BLOCKS[b];

    const capacity = block.end - block.start + 1;

    if (requiredRows > capacity) {
      continue;
    }

    const values = sheet
      .getRange(block.start, 1, capacity, 11)
      .getDisplayValues();

    for (let i = 0; i <= capacity - requiredRows; i++) {
      let available = true;

      for (let j = 0; j < requiredRows; j++) {
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
          start: block.start + i,
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

  const values = [[
    data.tanggal,
    data.noPO || "",
    data.vendor || "",
    p.name || "",
    p.qty || "",
    data.suhuMobil || "",
    data.kebersihanMobil || "",
    p.suhuProduk || "",
    p.kualitasProduk || "",
    p.tindakanKoreksi || "",
    p.keputusan || "",
    "",
    ""
  ]];

  sheet
    .getRange(row, 1, 1, CONFIG.DATA_COLUMNS)
    .setValues(values);
}


/****************************************************
 * CLEAR ROW
 ****************************************************/

function clearRow(sheet, row) {
  sheet
    .getRange(row, 1, 1, CONFIG.DATA_COLUMNS)
    .clearContent();
}


/****************************************************
 * INDEX SHEET
 ****************************************************/

function getIndexSheet(ss) {
  const name = "_INCOMING_INDEX";
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);

    sheet
      .getRange(1, 1, 1, 13)
      .setValues([[
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
 ****************************************************/

function saveIndex(ss, data) {
  const sheet = getIndexSheet(ss);

  sheet.appendRow([
    data.id,
    data.category,
    data.tanggal,
    data.sheet,
    data.startRow,
    data.endRow,
    data.separatorRow,
    data.noPO || "",
    data.vendor || "",
    JSON.stringify(data.products),
    new Date(),
    new Date(),
    "ACTIVE"
  ]);
}


/****************************************************
 * HISTORY
 ****************************************************/

function getHistory(category) {
  category = normalizeCategory(category);

  if (!category) {
    throw new Error("Kategori wajib dipilih.");
  }

  const ss = getSpreadsheet(category);
  const sheet = getIndexSheet(ss);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: true,
      category: category,
      data: []
    };
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 13)
    .getDisplayValues();

  const result = [];

  values.forEach(function(row) {
    if (!row[0] || row[12] === "DELETED") {
      return;
    }

    let products = [];

    try {
      products = JSON.parse(row[9] || "[]");
    } catch (e) {
      products = [];
    }

    result.push({
      id: row[0],
      category: row[1],
      tanggal: row[2],
      sheet: row[3],
      startRow: Number(row[4]),
      endRow: Number(row[5]),
      noPO: row[7],
      vendor: row[8],
      products: products,
      created: row[10],
      updated: row[11]
    });
  });

  result.reverse();

  return {
    success: true,
    category: category,
    data: result
  };
}


/****************************************************
 * UPDATE
 ****************************************************/

function updateIncoming(data) {
  if (!data.id) {
    throw new Error("ID transaksi tidak ditemukan.");
  }

  validateIncoming(data);

  const category = normalizeCategory(data.category);

  if (!category) {
    throw new Error("Kategori tidak valid.");
  }

  const ss = getSpreadsheet(category);
  const index = getIndexSheet(ss);

  const indexRow = findIndexRow(index, data.id);

  if (!indexRow) {
    throw new Error("Transaksi tidak ditemukan.");
  }

  const storedCategory = index
    .getRange(indexRow, 2)
    .getDisplayValue();

  if (normalizeCategory(storedCategory) !== category) {
    throw new Error("Kategori transaksi tidak cocok.");
  }

  const date = parseDate(data.tanggal);

  const sheetName = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "dd"
  );

  const newSheet = ss.getSheetByName(sheetName);

  if (!newSheet) {
    throw new Error("Sheet tanggal " + sheetName + " tidak ditemukan.");
  }

  const products = normalizeProducts(data.products);

  if (!products.length) {
    throw new Error("Minimal 1 produk.");
  }

  if (products.length > CONFIG.MAX_PRODUCTS) {
    throw new Error("Maksimal 5 produk.");
  }

  /*
   * Ambil lokasi lama.
   * Lokasi lama boleh dipakai kembali untuk update.
   */
  const oldSheetName = index
    .getRange(indexRow, 4)
    .getDisplayValue();

  const oldStart = Number(
    index.getRange(indexRow, 5).getValue()
  );

  const oldEnd = Number(
    index.getRange(indexRow, 6).getValue()
  );

  const oldSeparator = Number(
    index.getRange(indexRow, 7).getValue()
  );

  let location = null;

  /*
   * Jika tanggal tetap sama dan jumlah produk
   * masih muat di lokasi lama, gunakan lokasi lama.
   */
  if (
    oldSheetName === sheetName &&
    oldStart >= 1 &&
    oldEnd >= oldStart &&
    products.length <= (oldEnd - oldStart + 1)
  ) {
    location = {
      start: oldStart,
      block: null,
      end: oldEnd
    };
  }

  /*
   * Kalau lokasi lama tidak bisa dipakai,
   * cari lokasi kosong lain.
   *
   * Data lama belum dihapus sebelum lokasi baru
   * ditemukan agar update tidak menghilangkan data
   * jika ruang baru tidak tersedia.
   */
  if (!location) {
    location = findAvailableBlock(
      newSheet,
      products.length + 1
    );

    if (!location) {
      throw new Error("Tidak ada ruang untuk transaksi.");
    }
  }

  /*
   * Hapus lokasi lama jika berbeda.
   */
  const oldSheet = ss.getSheetByName(oldSheetName);

  if (oldSheet) {
    const sameLocation =
      oldSheetName === sheetName &&
      location.start === oldStart;

    if (!sameLocation) {
      for (let r = oldStart; r <= oldEnd; r++) {
        clearRow(oldSheet, r);
      }

      if (oldSeparator >= oldStart) {
        clearRow(oldSheet, oldSeparator);
      }
    }
  }

  /*
   * Bersihkan lokasi tujuan terlebih dahulu
   * supaya perubahan jumlah produk tidak meninggalkan
   * baris lama.
   */
  for (
    let r = location.start;
    r <= location.start + products.length;
    r++
  ) {
    clearRow(newSheet, r);
  }

  products.forEach(function(product, i) {
    writeRow(newSheet, location.start + i, {
      tanggal: date,
      noPO: data.noPO,
      vendor: data.vendor,
      suhuMobil: data.suhuMobil,
      kebersihanMobil: data.kebersihanMobil,
      product: product
    });
  });

  const separator = location.start + products.length;
  clearRow(newSheet, separator);

  const createdValue = index
    .getRange(indexRow, 11)
    .getValue();

  index
    .getRange(indexRow, 3, 1, 11)
    .setValues([[
      formatDate(date),
      sheetName,
      location.start,
      location.start + products.length - 1,
      separator,
      data.noPO || "",
      data.vendor || "",
      JSON.stringify(products),
      createdValue,
      new Date(),
      "ACTIVE"
    ]]);

  return {
    success: true,
    message: "Data berhasil diperbarui.",
    id: data.id,
    category: category
  };
}


/****************************************************
 * DELETE
 ****************************************************/

function deleteIncoming(data) {
  if (!data.id) {
    throw new Error("ID transaksi tidak ditemukan.");
  }

  const category = normalizeCategory(data.category);

  if (!category) {
    throw new Error("Kategori wajib dipilih.");
  }

  const ss = getSpreadsheet(category);
  const index = getIndexSheet(ss);
  const row = findIndexRow(index, data.id);

  if (!row) {
    throw new Error("Transaksi tidak ditemukan.");
  }

  const storedCategory = index
    .getRange(row, 2)
    .getDisplayValue();

  if (normalizeCategory(storedCategory) !== category) {
    throw new Error("Kategori transaksi tidak cocok.");
  }

  const sheetName = index
    .getRange(row, 4)
    .getDisplayValue();

  const start = Number(
    index.getRange(row, 5).getValue()
  );

  const end = Number(
    index.getRange(row, 6).getValue()
  );

  const separator = Number(
    index.getRange(row, 7).getValue()
  );

  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    for (let r = start; r <= end; r++) {
      clearRow(sheet, r);
    }

    if (separator >= start) {
      clearRow(sheet, separator);
    }
  }

  index
    .getRange(row, 13)
    .setValue("DELETED");

  index
    .getRange(row, 12)
    .setValue(new Date());

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

  if (lastRow < 2) {
    return null;
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0]).trim() ===
      String(id).trim()
    ) {
      return i + 2;
    }
  }

  return null;
}


/****************************************************
 * VENDOR
 *
 * FORMAT SHEET VENDOR:
 * A = Kode Vendor
 * B = Nama Supplier
 *
 * RESPONSE:
 * [kode, nama]
 *
 * Web bisa search berdasarkan:
 * - kode
 * - nama supplier
 *
 * Nilai yang dipilih web:
 * nama supplier
 ****************************************************/

function getVendor(category) {
  const ss = getSpreadsheet(category);

  const sheet = ss.getSheetByName(CONFIG.VENDOR_SHEET);

  if (!sheet) {
    throw new Error("Sheet VENDOR tidak ditemukan.");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues();

  return values
    .filter(function(row) {
      return (
        String(row[0]).trim() !== "" ||
        String(row[1]).trim() !== ""
      );
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
 *
 * FORMAT SHEET MASTER:
 * A = Barcode
 * B = Article
 * C = Nama Produk
 *
 * RESPONSE:
 * [barcode, article, namaProduk]
 *
 * Web bisa search berdasarkan:
 * - barcode
 * - article
 * - nama produk
 *
 * Jika barcode/article cocok,
 * web dapat langsung mengambil nama produk C.
 ****************************************************/

function getMaster(category) {
  const ss = getSpreadsheet(category);

  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);

  if (!sheet) {
    throw new Error("Sheet MASTER tidak ditemukan.");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 3)
    .getDisplayValues();

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
  if (!Array.isArray(products)) {
    return [];
  }

  return products
    .filter(function(p) {
      return (
        p &&
        (
          p.name ||
          p.nama ||
          p.namaProduk ||
          p.product
        )
      );
    })
    .slice(0, CONFIG.MAX_PRODUCTS)
    .map(function(p) {
      return {
        code:
          p.code ||
          p.kode ||
          p.barcode ||
          p.article ||
          "",

        name:
          p.name ||
          p.nama ||
          p.namaProduk ||
          p.product ||
          "",

        qty:
          p.qty ||
          p.quantity ||
          "",

        suhuProduk:
          p.suhuProduk ||
          p.suhu ||
          "",

        kualitasProduk:
          p.kualitasProduk ||
          p.kualitas ||
          "",

        tindakanKoreksi:
          p.tindakanKoreksi ||
          "",

        keputusan:
          p.keputusan ||
          ""
      };
    });
}


/****************************************************
 * VALIDATION
 ****************************************************/

function validateIncoming(data) {
  if (!data) {
    throw new Error("Data kosong.");
  }

  if (!normalizeCategory(data.category)) {
    throw new Error("Kategori wajib dipilih.");
  }

  if (!data.tanggal) {
    throw new Error("Tanggal wajib diisi.");
  }

  if (!data.noPO) {
    throw new Error("Nomor PO wajib diisi.");
  }

  if (!data.vendor) {
    throw new Error("Vendor wajib diisi.");
  }

  if (
    data.suhuMobil === undefined ||
    data.suhuMobil === null ||
    String(data.suhuMobil).trim() === ""
  ) {
    throw new Error("Suhu mobil wajib diisi.");
  }

  if (!data.kebersihanMobil) {
    throw new Error("Kebersihan mobil wajib dipilih.");
  }

  if (
    !Array.isArray(data.products) ||
    !data.products.length
  ) {
    throw new Error("Minimal 1 produk wajib diisi.");
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
  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {
    return value;
  }

  const text = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const p = text.split("-");

    return new Date(
      Number(p[0]),
      Number(p[1]) - 1,
      Number(p[2])
    );
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    const p = text.split("/");

    return new Date(
      Number(p[2]),
      Number(p[1]) - 1,
      Number(p[0])
    );
  }

  const date = new Date(text);

  if (isNaN(date.getTime())) {
    throw new Error("Tanggal tidak valid.");
  }

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

  Logger.log(
    "Jumlah vendor: " + data.length
  );

  if (data.length) {
    Logger.log(
      JSON.stringify(data.slice(0, 10), null, 2)
    );
  }
}

function testVendorNonHalal() {
  const data = getVendor("NON HALAL");

  Logger.log(
    "Jumlah vendor: " + data.length
  );

  if (data.length) {
    Logger.log(
      JSON.stringify(data.slice(0, 10), null, 2)
    );
  }
}


/****************************************************
 * TEST MASTER
 ****************************************************/

function testMasterHalal() {
  const data = getMaster("HALAL");

  Logger.log(
    "Jumlah master: " + data.length
  );

  if (data.length) {
    Logger.log(
      JSON.stringify(data.slice(0, 10), null, 2)
    );
  }
}

function testMasterNonHalal() {
  const data = getMaster("NON HALAL");

  Logger.log(
    "Jumlah master: " + data.length
  );

  if (data.length) {
    Logger.log(
      JSON.stringify(data.slice(0, 10), null, 2)
    );
  }
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
      Logger.log(
        "❌ Sheet " +
        testSheet +
        " TIDAK DITEMUKAN"
      );
      return;
    }

    Logger.log(
      "✅ Sheet " +
      testSheet +
      " ditemukan"
    );

    CONFIG.BLOCKS.forEach(function(block, index) {
      const capacity =
        block.end - block.start + 1;

      Logger.log(
        "Lembar " +
        (index + 1) +
        ": baris " +
        block.start +
        "-" +
        block.end +
        " | kapasitas " +
        capacity
      );
    });

    const location =
      findAvailableBlock(
        sheet,
        requiredRows
      );

    if (!location) {
      Logger.log(
        "❌ TIDAK ADA RUANG YANG CUKUP"
      );
    } else {
      Logger.log("✅ RUANG DITEMUKAN");
      Logger.log("Lembar: " + location.block);
      Logger.log("Mulai baris: " + location.start);
      Logger.log(
        "Baris produk terakhir: " +
        (location.start + jumlahProduk - 1)
      );
      Logger.log(
        "Baris pemisah: " +
        (location.start + jumlahProduk)
      );
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

  Logger.log(
    JSON.stringify(result, null, 2)
  );
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

  const values = index
    .getRange(2, 1, lastRow - 1, 13)
    .getDisplayValues();

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
        for (let r = start; r <= end; r++) {
          clearRow(sheet, r);
        }

        if (separator >= start) {
          clearRow(sheet, separator);
        }
      }

      index
        .getRange(row, 13)
        .setValue("DELETED");

      index
        .getRange(row, 12)
        .setValue(new Date());

      Logger.log("Data TEST berhasil dihapus.");
      Logger.log("PO: TEST-INCOMING-001");

      found = true;
      break;
    }
  }

  if (!found) {
    Logger.log(
      "Data TEST-INCOMING-001 tidak ditemukan."
    );
  }
}
