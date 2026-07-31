/**
 * ============================================================================
 * WAA Player Evaluation System
 * AccessCodes.gs
 * Version 1.1
 * ============================================================================
 */

function generateCoachAccessCodes() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Coaches");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Coaches sheet not found.");
    return;
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("No coaches found.");
    return;
  }

  const headers = values[0];

  const COL = {
    coachName: headers.indexOf("Coach Name"),
    teamName: headers.indexOf("Team Name"),
    accessCode: headers.indexOf("Access Code"),
    active: headers.indexOf("Active")
  };

  for (const key in COL) {
    if (COL[key] === -1) {
      throw new Error("Missing required Coaches header: " + key);
    }
  }

  const existingCodes = new Set();

  values.slice(1).forEach(row => {
    if (row[COL.accessCode]) {
      existingCodes.add(row[COL.accessCode]);
    }
  });

  let created = 0;
  let skipped = 0;
  let inactive = 0;

  for (let r = 1; r < values.length; r++) {

    const row = values[r];

    const active =
      String(row[COL.active]).toUpperCase() === "TRUE" ||
      String(row[COL.active]).toUpperCase() === "YES";

    if (!active) {
      inactive++;
      continue;
    }

    if (row[COL.accessCode]) {
      skipped++;
      continue;
    }

    let code;

    do {
      code = createAccessCode(row[COL.teamName]);
    } while (existingCodes.has(code));

    sheet
      .getRange(r + 1, COL.accessCode + 1)
      .setValue(code);

    existingCodes.add(code);

    created++;
  }

  SpreadsheetApp.getUi().alert(
      "Access Code Generation Complete\n\n" +
      "Created: " + created +
      "\nSkipped: " + skipped +
      "\nInactive: " + inactive
  );

}


function createAccessCode(teamName) {

  const prefix = teamName
    .replace(/[^A-Za-z]/g,"")
    .substring(0,3)
    .toUpperCase();

  const number =
      Math.floor(1000 + Math.random()*9000);

  return prefix + "-" + number;

}


function clearCoachAccessCodes(){

  const sheet =
      SpreadsheetApp
      .getActive()
      .getSheetByName("Coaches");

  if(!sheet){
    SpreadsheetApp.getUi().alert("Coaches sheet not found.");
    return;
  }

  const lastRow = sheet.getLastRow();

  if(lastRow <=1) return;

  const headers =
      sheet
      .getRange(1,1,1,sheet.getLastColumn())
      .getValues()[0];

  const col =
      headers.indexOf("Access Code")+1;

  if(col===0){
      throw new Error("Access Code column not found.");
  }

  sheet
      .getRange(2,col,lastRow-1)
      .clearContent();

  SpreadsheetApp
      .getUi()
      .alert("All coach access codes cleared.");

}
