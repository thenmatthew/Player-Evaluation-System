/**
 * ============================================================================
 * Player.gs
 * WAA Player Evaluation System
 * Version 1.1
 * ============================================================================
 */

function getTeamPlayers(teamId) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Players");

  if (!sheet) {
    throw new Error("Players sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0];

  const COL = {
    playerId: headers.indexOf("Player ID"),
    teamId: headers.indexOf("Team ID"),
    firstName: headers.indexOf("First Name"),
    lastName: headers.indexOf("Last Name"),
    grade: headers.indexOf("Grade")
  };

  for (const key in COL) {
    if (COL[key] === -1) {
      throw new Error("Missing Players column: " + key);
    }
  }

  return values
    .slice(1)
    .filter(row => String(row[COL.teamId]).trim() === String(teamId).trim())
    .map(row => ({
      playerId: row[COL.playerId],
      firstName: row[COL.firstName],
      lastName: row[COL.lastName],
      grade: row[COL.grade]
    }));

}