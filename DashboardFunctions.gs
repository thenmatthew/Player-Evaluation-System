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

const DIVISION_ORDER = [
  "CP M/W",
  "CP T/R",
  "MP M/W",
  "MP T/R",
  "3rd/4th M/W",
  "5th/6th",
  "7th/8th/9th",
  "Senior High"
];

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

    division:
      String(coaches[c][0]).trim(),

    coachName:
      String(coaches[c][1]).trim(),

    teamName:
      String(coaches[c][3]).trim()

};

}

// ------------------------------------------
// PLAYER LOOKUP
// ------------------------------------------

const playerLookup = {};


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

      division:
        coach
          ? coach.division
          : "Unknown",

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

// ==========================================
// ANALYTICS ENGINE
//
// All dashboard metrics are calculated here.
// The UI only displays the returned values.
// ==========================================

// ------------------------------------------
// LEAGUE ANALYTICS
// ------------------------------------------

// ------------------------------------------
// DIVISION ANALYTICS
// ------------------------------------------

const divisionAnalytics = {};

// ------------------------------------------
// TEAM ANALYTICS
// ------------------------------------------

// Story 5

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

const overall =
  Number(evaluations[i][13]);

if (
  teams[evaluationTeamId]
) {

  const team =
    teams[evaluationTeamId];

  team.evaluated++;

  const division =
    team.division;

  if (!divisionAnalytics[division]) {

    divisionAnalytics[division] = {
      total: 0,
      count: 0
    };

  }

  if (!isNaN(overall) && overall > 0) {

    divisionAnalytics[division]
      .total += overall;

    divisionAnalytics[division]
      .count++;

  }

}

}   // <-- closes the for loop

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


teamList.sort(function(a, b) {

  const divisionA =
    DIVISION_ORDER.indexOf(a.division);

  const divisionB =
    DIVISION_ORDER.indexOf(b.division);

  // Unknown divisions sort to the end
  const orderA =
    divisionA === -1
      ? Number.MAX_SAFE_INTEGER
      : divisionA;

  const orderB =
    divisionB === -1
      ? Number.MAX_SAFE_INTEGER
      : divisionB;

  if (orderA !== orderB) {
    return orderA - orderB;
  }

  return a.teamName.localeCompare(b.teamName);

});



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



// ------------------------------------------
// BUILD DIVISION ANALYTICS
// ------------------------------------------

const divisionMap = {};

teamList.forEach(function(team) {

if (!divisionMap[team.division]) {

  divisionMap[team.division] = {
    name: team.division,
    teamCount: 0,
    totalPlayers: 0,
    evaluatedPlayers: 0,
    completionPercent: 0,

    // Internal accumulators
    totalOverall: 0,
    ratedPlayers: 0
  };

}

  const division = divisionMap[team.division];

  division.teamCount++;
  division.totalPlayers += team.players;
  division.evaluatedPlayers += team.evaluated;

});

const divisions = [];

DIVISION_ORDER.forEach(function(name) {

  const division = divisionMap[name];

  const analytics =
  divisionAnalytics[name];

  if (!division) {
    return;
  }

  division.completionPercent =
    division.totalPlayers === 0
      ? 0
      : Number(
          (
            division.evaluatedPlayers /
            division.totalPlayers *
            100
          ).toFixed(1)
        );

division.averageRating =
  analytics && analytics.count > 0
    ? Number(
        (
          analytics.total /
          analytics.count
        ).toFixed(2)
      )
    : 0;
  divisions.push(division);

});

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

  divisions:
    divisions,

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

  overall:
    Number(evaluations[i][13]),

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
