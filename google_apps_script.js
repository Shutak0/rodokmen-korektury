/**
 * Zabezpečený Google Apps Script pro archivní verifikační centrum.
 * 
 * BEZPEČNOSTNÍ PRVKY:
 * 1. Každý příchozí požadavek (POST i GET) je striktně ověřován tajným autorizačním tokenem AUTH_TOKEN.
 * 2. Cizí osoby bez znalosti tokenu nemohou do tabulky nic zapsat ani z ní číst (vrací se 403 Forbidden).
 * 3. Nastavení "Anyone" (Kdokoli) v Google nasazení slouží pouze k tomu, aby se specialista nemusel
 *    přihlašovat Google účtem, ale přístup je chráněn naším API klíčem.
 */

// Tajný bezpečnostní token (musí se shodovat s tokenem v index.html)
var AUTH_TOKEN = "e87f8b9e4a83b2719fbc490d12e68c94a53e";

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // List 1: Záznamy
  var sheetRecords = ss.getSheetByName("Urbářské záznamy");
  if (!sheetRecords) {
    sheetRecords = ss.insertSheet("Urbářské záznamy");
    sheetRecords.appendRow([
      "Zdrojový soubor (Snímek)",
      "Folio / Strana",
      "Lokalita / Ves",
      "Jméno a Příjmení",
      "Role / Povolání",
      "Držba půdy / Grunt",
      "Peněžitý úrok (sv. Jiří / sv. Havel)",
      "Robota",
      "Naturální dávky",
      "Poznámky k usedlosti",
      "Stav ověření",
      "Datum uložení"
    ]);
    sheetRecords.getRange(1, 1, 1, 12).setBackground("#2B4C7E").setFontColor("#FFFFFF").setFontWeight("bold");
    sheetRecords.setFrozenRows(1);
  }

  // List 2: Stav ověření
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

    // 2. Vložení aktualizovaných řádků
    for (var r = 0; r < records.length; r++) {
      var rec = records[r];
      sheetRecords.appendRow([
        photoPath,
        rec.source_page_folio || "",
        rec.settlement_orig || "",
        rec.full_name || "",
        rec.raw_role || "",
        rec.land_tenure_raw || "",
        rec.money_rent_urok || "",
        rec.corvee_robota || "",
        rec.natural_dues || "",
        rec.estate_notes || "",
        isVerified,
        timestamp
      ]);
    }

    // 3. Aktualizace stavu ověření
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
  // Ověření tokenu i pro čtení
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
      result.push({
        source_file: data[i][0],
        source_page_folio: data[i][1],
        settlement_orig: data[i][2],
        full_name: data[i][3],
        raw_role: data[i][4],
        land_tenure_raw: data[i][5],
        money_rent_urok: data[i][6],
        corvee_robota: data[i][7],
        natural_dues: data[i][8],
        estate_notes: data[i][9],
        is_verified: data[i][10] === "Ano",
        updated_at: data[i][11]
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
