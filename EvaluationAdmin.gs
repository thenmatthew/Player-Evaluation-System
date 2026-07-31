function resetEvaluation() {

  const ui =
    SpreadsheetApp.getUi();


  const response =
    ui.prompt(
      "Reset Evaluation",
      "Enter Player ID:",
      ui.ButtonSet.OK_CANCEL
    );


  if (
    response.getSelectedButton()
    !== ui.Button.OK
  ) {

    return;

  }


  const playerId =
    response.getResponseText()
    .trim();



  if (!playerId) {

    ui.alert(
      "No Player ID entered."
    );

    return;

  }



  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Evaluations");



  const data =
    sheet.getDataRange()
    .getValues();



  let deleted = false;



  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {


    // Player ID is column F (index 5)

    if (
      data[i][5]
      .toString()
      === playerId
    ) {


      const confirm =
        ui.alert(
          "Confirm Reset",
          "Delete evaluation for "
          + data[i][6]
          + "?",
          ui.ButtonSet.YES_NO
        );


      if (
        confirm === ui.Button.YES
      ) {

        sheet.deleteRow(i + 1);

        deleted = true;

      }


      break;

    }

  }



  if (deleted) {

    ui.alert(
      "Evaluation removed successfully."
    );

  } else {

    ui.alert(
      "No evaluation found for that Player ID."
    );

  }

}
