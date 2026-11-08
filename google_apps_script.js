/**
 * Zabezpečený Google Apps Script pro archivní verifikační centrum.
 * 
 * Podporuje standardizovaný dynamický katalog kategorií:
 * Všechna známá pole jsou v tabulce vždy přítomna pro každou osobu
 * (pokud není vyplněno, zůstává prázdné), plus dynamické vlastní kategorie.
 */

var AUTH_TOKEN = "e87f8b9e4a83b2719fbc490d12e68c94a53e";

// Standardní pevné pořadí sloupců
var STANDARD_COLUMNS = [
  "Zdrojový soubor (Snímek)",
  "Folio / Strana",
  "Lokalita / Ves",
  "Křestní jméno",
  "Příjmení / Přízvisko",
  "Celé jméno",
  "Role / Povolání",
  "Společenský stav",
  "Držba půdy / Grunt",
  "Úrok sv. Jiří",
  "Úrok sv. Havel",
  "Celkový peněžitý úrok",
  "Robota",
  "Slepice (ks)",
  "Vejce (ks)",
  "Oves / obilí",
  "Ostatní naturálie",
  "Datum narození / křtu",
  "Datum sňatku",
  "Datum úmrtí / pohřbu",
  "Místo / datum pobytu",
  "Poznámky",
  "Vlastní pole (JSON)",
  "Stav ověření",
  "Datum uložení"
];

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // List 1: Záznamy
  var sheetRecords = ss.getSheetByName("Urbářské záznamy");
  if (!sheetRecords) {
    sheetRecords = ss.insertSheet("Urbářské záznamy");
    sheetRecords.appendRow(STANDARD_COLUMNS);
    sheetRecords.getRange(1, 1, 1, STANDARD_COLUMNS.length).setBackground("#2B4C7E").setFontColor("#FFFFFF").setFontWeight("bold");
    sheetRecords.setFrozenRows(1);
  } else {
    // Pokud má starou strukturu záhlaví (např. 12 sloupců), aktualizujeme ho na standardizované
    var headerRow = sheetRecords.getRange(1, 1, 1, sheetRecords.getLastColumn() || 1).getValues()[0];
    if (headerRow.length < STANDARD_COLUMNS.length || headerRow[3] !== "Křestní jméno") {
      sheetRecords.getRange(1, 1, 1, STANDARD_COLUMNS.length).setValues([STANDARD_COLUMNS]);
      sheetRecords.getRange(1, 1, 1, STANDARD_COLUMNS.length).setBackground("#2B4C7E").setFontColor("#FFFFFF").setFontWeight("bold");
      sheetRecords.setFrozenRows(1);
    }
  }

  // List 2: Stav ověření snímků
  var sheetStatus = ss.getSheetByName("Stav ověření snímků");
  if (!sheetStatus) {
    sheetStatus = ss.insertSheet("Stav ověření snímků");
    sheetStatus.appendRow(["Snímek", "Zkontrolováno (Ano/Ne)", "Počet osob", "Poznámka", "Datum a čas"]);
    sheetStatus.getRange(1, 1, 1, 5).setBackground("#2B4C7E").setFontColor("#FFFFFF").setFontWeight("bold");
    sheetStatus.setFrozenRows(1);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Empty request" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var contents = JSON.parse(e.postData.contents);

    // KONTROLA BEZPEČNOSTNÍHO TOKENU
    if (contents.auth_token !== AUTH_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "403 Forbidden: Neplatný autorizační token"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    setupSheet();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRecords = ss.getSheetByName("Urbářské záznamy");
    var sheetStatus = ss.getSheetByName("Stav ověření snímků");

    var photoPath = contents.photo_path || "";
    var records = contents.records || [];
    var isVerified = contents.is_verified ? "Ano" : "Ne";
    var notes = contents.notes || "";
    var timestamp = new Date();

    // 1. Odstranění předchozích verzí řádků pro daný snímek (prevence duplicit)
    var data = sheetRecords.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === photoPath) {
        sheetRecords.deleteRow(i + 1);
      }
    }

    // 2. Vložení aktualizovaných řádků se všemi standardizovanými poli
    for (var r = 0; r < records.length; r++) {
      var rec = records[r];
      var customFieldsStr = "";
      if (rec.custom_fields && typeof rec.custom_fields === 'object') {
        customFieldsStr = JSON.stringify(rec.custom_fields);
      }

      sheetRecords.appendRow([
        photoPath,
        rec.source_page_folio || "",
        rec.settlement_orig || "",
        rec.first_name || "",
        rec.last_name || "",
        rec.full_name || (rec.first_name + " " + rec.last_name).trim(),
        rec.raw_role || "",
        rec.social_status || "",
        rec.land_tenure_raw || "",
        rec.rent_george || "",
        rec.rent_gallus || "",
        rec.money_rent_urok || "",
        rec.corvee_robota || "",
        rec.dues_chickens || "",
        rec.dues_eggs || "",
        rec.dues_grain || "",
        rec.dues_other || rec.natural_dues || "",
        rec.birth_date || "",
        rec.marriage_date || "",
        rec.death_date || "",
        rec.residence_info || "",
        rec.notes || rec.estate_notes || "",
        customFieldsStr,
        isVerified,
        timestamp
      ]);
    }

    // 3. Aktualizace stavu ověření snímku
    var statusData = sheetStatus.getDataRange().getValues();
    var statusRowIndex = -1;
    for (var s = 1; s < statusData.length; s++) {
      if (statusData[s][0] === photoPath) {
        statusRowIndex = s + 1;
        break;
      }
    }

    if (statusRowIndex > 0) {
      sheetStatus.getRange(statusRowIndex, 2, 1, 4).setValues([[isVerified, records.length, notes, timestamp]]);
    } else {
      sheetStatus.appendRow([photoPath, isVerified, records.length, notes, timestamp]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      records_saved: records.length,
      photo: photoPath
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var token = (e && e.parameter) ? e.parameter.auth_token : "";
  if (token !== AUTH_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: "403 Forbidden: Neplatný autorizační token"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  setupSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRecords = ss.getSheetByName("Urbářské záznamy");
  var data = sheetRecords.getDataRange().getValues();
  
  var photoQuery = e.parameter.photo || "";
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (!photoQuery || data[i][0] === photoQuery) {
      var row = data[i];
      var customObj = {};
      if (row[22]) {
        try { customObj = JSON.parse(row[22]); } catch(ex) {}
      }

      result.push({
        source_file: row[0],
        source_page_folio: row[1],
        settlement_orig: row[2],
        first_name: row[3],
        last_name: row[4],
        full_name: row[5],
        raw_role: row[6],
        social_status: row[7],
        land_tenure_raw: row[8],
        rent_george: row[9],
        rent_gallus: row[10],
        money_rent_urok: row[11],
        corvee_robota: row[12],
        dues_chickens: row[13],
        dues_eggs: row[14],
        dues_grain: row[15],
        dues_other: row[16],
        birth_date: row[17],
        marriage_date: row[18],
        death_date: row[19],
        residence_info: row[20],
        notes: row[21],
        custom_fields: customObj,
        is_verified: row[23] === "Ano"
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
