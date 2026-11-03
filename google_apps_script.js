/**
 * Google Apps Script Backend for Archival OCR Verification Station.
 * 
 * INSTRUCTIONS:
 * 1. Open your Google Drive and create a new Google Sheet (e.g., "Archivní korektury 1632-1651").
 * 2. In the Google Sheet menu, click: Extensions -> Apps Script (Rozšíření -> Apps Script).
 * 3. Delete any default code in Code.gs and paste this ENTIRE file into it.
 * 4. Click "Deploy" (Nasadit) -> "New deployment" (Nové nasazení).
 * 5. Click the gear icon next to "Select type" and choose "Web app" (Webová aplikace).
 * 6. Set:
 *    - Description: "Archivum OCR API"
 *    - Execute as: "Me" (Já)
 *    - Who has access: "Anyone" (Kdokoli)  <-- VERY IMPORTANT for GitHub Pages to be able to send data!
 * 7. Click "Deploy" and authorize the script.
 * 8. COPY the "Web app URL" (looks like https://script.google.com/macros/s/.../exec)
 *    and paste it into the settings of your GitHub Pages site!
 */

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet 1: Records
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

  // Sheet 2: Verification Status
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
    setupSheet();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRecords = ss.getSheetByName("Urbářské záznamy");
    var sheetStatus = ss.getSheetByName("Stav ověření snímků");

    var contents = JSON.parse(e.postData.contents);
    var photoPath = contents.photo_path || "";
    var records = contents.records || [];
    var isVerified = contents.is_verified ? "Ano" : "Ne";
    var notes = contents.notes || "";
    var timestamp = new Date();

    // 1. Remove previous entries for this photo to avoid duplicates
    var data = sheetRecords.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === photoPath) {
        sheetRecords.deleteRow(i + 1);
      }
    }

    // 2. Append updated records
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

    // 3. Update Verification Status sheet
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
