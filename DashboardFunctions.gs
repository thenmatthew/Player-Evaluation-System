function openDashboard() {

  const html =
    HtmlService
      .createHtmlOutputFromFile("DashboardView")
      .setTitle("League Dashboard");

  SpreadsheetApp
    .getUi()
    .showSidebar(html);

}


function getDashboardData() {


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const coachesSheet =
    ss.getSheetByName("Coaches");

  const playersSheet =
    ss.getSheetByName("Players");

  const evaluationsSheet =
    ss.getSheetByName("Evaluations");


  if (
    !coachesSheet ||
    !playersSheet ||
    !evaluationsSheet
  ) {

    throw new Error(
      "Missing required sheet. " +
      "Required: Coaches, Players, Evaluations."
    );

  }


  const coaches =
    coachesSheet
      .getDataRange()
      .getValues();

  const players =
    playersSheet
      .getDataRange()
      .getValues();

  const evaluations =
    evaluationsSheet
      .getDataRange()
      .getValues();


const teams = {};


// ------------------------------------------
// BUILD COACH LOOKUP
// Coaches:
// A = Coach Code
// B = Coach Name
// C = Team ID
// D = Team Name
// ------------------------------------------

const coachMap = {};

for (let c = 1; c < coaches.length; c++) {

  const coachTeamId =
    String(coaches[c][2])
      .trim()
      .toUpperCase();

  if (!coachTeamId) {
    continue;
  }

  coachMap[coachTeamId] = {
    coachName:
      String(coaches[c][1]).trim(),

    teamName:
      String(coaches[c][3]).trim()
  };

}


// ------------------------------------------
// BUILD TEAMS FROM PLAYERS
// ------------------------------------------

for (let i = 1; i < players.length; i++) {

  const teamId =
    String(players[i][1]).trim();

  if (!teamId) {
    continue;
  }


  if (!teams[teamId]) {

    const playerTeamId =
      teamId.toUpperCase();

    const coach =
      coachMap[playerTeamId];


    teams[teamId] = {

      teamId:
        teamId,

      teamName:
        coach
          ? coach.teamName
          : "Unknown",

      coachName:
        coach
          ? coach.coachName
          : "Not Assigned",

      players:
        0,

      evaluated:
        0

    };

  }


  teams[teamId].players++;

}


  // ------------------------------------------
  // COUNT EVALUATIONS
  // ------------------------------------------

  for (
    let i = 1;
    i < evaluations.length;
    i++
  ) {

    const evaluationTeamId =
      String(evaluations[i][3])
        .trim();


    if (
      teams[evaluationTeamId]
    ) {

      teams[evaluationTeamId]
        .evaluated++;

    }

  }


  // ------------------------------------------
  // BUILD FINAL TEAM LIST
  // ------------------------------------------

  const teamList = [];


  Object.keys(teams)
    .forEach(function(id) {

      const team =
        teams[id];


      team.percent =
        team.players === 0
          ? 0
          : Math.round(
              (
                team.evaluated /
                team.players
              ) * 100
            );


      teamList.push(team);

    });


  // Sort alphabetically by team name

 teamList.sort(function(a, b) {

  return a.teamName
    .localeCompare(b.teamName);

});


// ------------------------------------------
// BUILD LEAGUE ANALYTICS
// ------------------------------------------

// ------------------------------------------
// CALCULATE LEAGUE AVERAGE
// ------------------------------------------

let totalOverall = 0;
let ratedPlayers = 0;

for (let i = 1; i < evaluations.length; i++) {

  const overall =
    Number(evaluations[i][13]);

  if (!isNaN(overall) && overall > 0) {

    totalOverall += overall;
    ratedPlayers++;

  }

}

const leagueAverage =
  ratedPlayers === 0
    ? 0
    : Number(
        (totalOverall / ratedPlayers)
          .toFixed(2)
      );
const totalPlayers =
  teamList.reduce(function(total, team) {

    return total + team.players;

  }, 0);


const evaluatedPlayers =
  teamList.reduce(function(total, team) {

    return total + team.evaluated;

  }, 0);


const completionPercent =
  totalPlayers === 0
    ? 0
    : Number(
        (
          evaluatedPlayers /
          totalPlayers *
          100
        ).toFixed(1)
      );


return {

  analytics: {

    teams:
      teamList.length,

    totalPlayers:
      totalPlayers,

    evaluatedPlayers:
      evaluatedPlayers,

    completionPercent:
      completionPercent,

leagueAverage:
      leagueAverage

  },

  teams:
    teamList

};

}



function getTeamPlayers(teamId) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const playersSheet =
    ss.getSheetByName("Players");

  const evaluationsSheet =
    ss.getSheetByName("Evaluations");


  if (
    !playersSheet ||
    !evaluationsSheet
  ) {

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


  const evaluatedPlayers = {};


  // ------------------------------------------
  // BUILD EVALUATED PLAYER LOOKUP
  // ------------------------------------------

  for (
    let i = 1;
    i < evaluations.length;
    i++
  ) {

    const playerId =
      String(evaluations[i][5])
        .trim();


    if (playerId) {

      evaluatedPlayers[playerId] = {

        playerName:
          evaluations[i][6],

        position:
          evaluations[i][15],

        pitching:
          evaluations[i][16],

        catching:
          evaluations[i][17]

      };

    }

  }


  const teamPlayers = [];


  // ------------------------------------------
  // BUILD PLAYER LIST FOR TEAM
  // ------------------------------------------

  const requestedTeamId =
    String(teamId)
      .trim();


  for (
    let i = 1;
    i < players.length;
    i++
  ) {

    const playerTeamId =
      String(players[i][1])
        .trim();


    if (
      playerTeamId === requestedTeamId
    ) {

      const playerId =
        String(players[i][0])
          .trim();


      teamPlayers.push({

        playerId:
          playerId,

        name:
          String(players[i][2]) +
          " " +
          String(players[i][3]),

        grade:
          players[i][4],

        evaluated:
          !!evaluatedPlayers[playerId],

        position:
          evaluatedPlayers[playerId]
            ? evaluatedPlayers[playerId]
                .position
            : "Not Evaluated"

      });

    }

  }


  return teamPlayers;

}
function testCoachMatch() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const coachesSheet =
    ss.getSheetByName("Coaches");

  const playersSheet =
    ss.getSheetByName("Players");

  const coaches =
    coachesSheet
      .getDataRange()
      .getValues();

  const players =
    playersSheet
      .getDataRange()
      .getValues();


  // Find the first player using 34MW-1

  let playerTeamId = "";

  for (
    let i = 1;
    i < players.length;
    i++
  ) {

    if (
      String(players[i][1])
        .trim()
        .toUpperCase()
        === "34MW-1"
    ) {

      playerTeamId =
        String(players[i][1])
          .trim()
          .toUpperCase();

      break;

    }

  }


  // Find the coach row using 34MW-1

  let coachResult = "NOT FOUND";

  for (
    let i = 1;
    i < coaches.length;
    i++
  ) {

    const coachTeamId =
      String(coaches[i][2])
        .trim()
        .toUpperCase();

    if (
      coachTeamId === "34MW-1"
    ) {

      coachResult = {
        row: i + 1,
        coachCode: coaches[i][0],
        coachName: coaches[i][1],
        teamId: coaches[i][2],
        teamName: coaches[i][3],
        rawRow: coaches[i]
      };

      break;

    }

  }


  throw new Error(
    JSON.stringify({
      playerTeamId: playerTeamId,
      coachResult: coachResult
    }, null, 2)
  );

}
