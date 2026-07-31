/*
==========================================================
League Player Evaluation System (LPES)

Module:
Evaluations.gs

Purpose:
Creates, updates, and retrieves player evaluations.

Version:
1.1.0

Status:
Development

Last Updated:
2026-07-31
==========================================================
*/


/**
 * ---------------------------------------------------------
 * saveEvaluation()
 *
 * Creates a new player evaluation or updates an existing one.
 * ---------------------------------------------------------
 */
function saveEvaluation(evaluation) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Evaluations");

  if (!sheet) {
    throw new Error("Evaluations sheet not found.");
  }

  // Find an existing evaluation for this player
  const data = sheet.getDataRange().getValues();
  let rowToUpdate = null;

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][5]) === String(evaluation.playerId)
    ) {
      rowToUpdate = i + 1;
      break;
    }

  }

  const season =
    new Date().getFullYear();

  const row = [

    new Date(),                    // Timestamp
    season,                        // Season
    evaluation.coachName,          // Coach
    evaluation.teamId,             // Team ID
    evaluation.teamName,           // Team Name
    evaluation.playerId,           // Player ID
    evaluation.playerName,         // Player Name
    evaluation.fielding,           // Fielding
    evaluation.hitting,            // Hitting
    evaluation.throwing,           // Throwing
    evaluation.running,            // Running
    evaluation.gameConcepts || "", // Game Concepts
    evaluation.attitude || "",     // Attitude
    evaluation.overall,            // Overall
    evaluation.comments,           // Comments
    evaluation.primaryPosition,    // Primary Position
    evaluation.pitchingAbility,    // Pitching Ability
    evaluation.catchingAbility     // Catching Ability

  ];

  if (rowToUpdate) {

    sheet
      .getRange(rowToUpdate, 1, 1, row.length)
      .setValues([row]);

    return {
      success: true,
      message: "Evaluation updated successfully."
    };

  }

  sheet.appendRow(row);

  return {
    success: true,
    message: "Evaluation saved successfully."
  };

}


/**
 * ---------------------------------------------------------
 * getCoachSummary(teamId)
 *
 * Returns every player on a team along with their
 * evaluation information, if available.
 * ---------------------------------------------------------
 */
function getCoachSummary(teamId) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const playersSheet =
    ss.getSheetByName("Players");

  const evaluationsSheet =
    ss.getSheetByName("Evaluations");

  if (!playersSheet || !evaluationsSheet) {

    throw new Error(
      "Missing Players or Evaluations sheet."
    );

  }

  const players =
    playersSheet
      .getDataRange()
      .getValues();

  const evaluations =
    evaluationsSheet
      .getDataRange()
      .getValues();

  //--------------------------------------------------
  // Build Evaluation Lookup
  //--------------------------------------------------

  const evaluationLookup = {};

  for (let i = 1; i < evaluations.length; i++) {

    const playerId =
      String(evaluations[i][5]).trim();

    if (!playerId) {
      continue;
    }

    evaluationLookup[playerId] = {

      fielding:
        evaluations[i][7],

      hitting:
        evaluations[i][8],

      throwing:
        evaluations[i][9],

      running:
        evaluations[i][10],

      gameConcepts:
        evaluations[i][11],

      attitude:
        evaluations[i][12],

      overall:
        evaluations[i][13],

      comments:
        evaluations[i][14],

      primaryPosition:
        evaluations[i][15],

      pitchingAbility:
        evaluations[i][16],

      catchingAbility:
        evaluations[i][17]

    };

  }

  //--------------------------------------------------
  // Build Team Summary
  //--------------------------------------------------

  const requestedTeam =
    String(teamId).trim();

  const summaryPlayers = [];

  let evaluatedPlayers = 0;

  for (let i = 1; i < players.length; i++) {

    const playerTeam =
      String(players[i][1]).trim();

    if (playerTeam !== requestedTeam) {
      continue;
    }

    const playerId =
      String(players[i][0]).trim();

    const evaluation =
      evaluationLookup[playerId];

    if (evaluation) {
      evaluatedPlayers++;
    }

    summaryPlayers.push({

      playerId:
        playerId,

      playerName:
        String(players[i][2]) +
        " " +
        String(players[i][3]),

      grade:
        players[i][4],

      evaluated:
        !!evaluation,

      fielding:
        evaluation ? evaluation.fielding : "",

      hitting:
        evaluation ? evaluation.hitting : "",

      throwing:
        evaluation ? evaluation.throwing : "",

      running:
        evaluation ? evaluation.running : "",

      gameConcepts:
        evaluation ? evaluation.gameConcepts : "",

      attitude:
        evaluation ? evaluation.attitude : "",

      overall:
        evaluation ? evaluation.overall : "",

      comments:
        evaluation ? evaluation.comments : "",

      primaryPosition:
        evaluation ? evaluation.primaryPosition : "",

      pitchingAbility:
        evaluation ? evaluation.pitchingAbility : "",

      catchingAbility:
        evaluation ? evaluation.catchingAbility : ""

    });

  }

  //--------------------------------------------------
  // Return Structured Summary
  //--------------------------------------------------

  const totalPlayers =
    summaryPlayers.length;

  return {

    teamId:
      requestedTeam,

    totalPlayers:
      totalPlayers,

    evaluatedPlayers:
      evaluatedPlayers,

    percentComplete:
      totalPlayers === 0
        ? 0
        : Math.round(
            (evaluatedPlayers / totalPlayers) * 100
          ),

    players:
      summaryPlayers

  };

}