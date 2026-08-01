function validateCoachCode(code) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Coaches");


  if (!sheet) {

    throw new Error(
      "Coaches sheet not found."
    );

  }


  const data = sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {


    const accessCode = data[i][4];
    const active = data[i][5];


    if (
      accessCode === code &&
      active !== "No"
    ) {


      // Update Last Used column
      sheet
        .getRange(i + 1, 7)
        .setValue(new Date());


      return {

        success: true,

        coachName: data[i][1],

        teamId: data[i][2],

        teamName: data[i][3],

        players: getTeamPlayers(data[i][2])

      };

    }

  }


  return {

    success: false,

    message: "Invalid access code."

  };

}
