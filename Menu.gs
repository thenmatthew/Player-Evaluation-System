function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("League Tools")

    .addItem(
      "Setup League Workbook",
      "setupLeagueWorkbook"
    )

    .addItem(
      "Generate Coach Access Codes",
      "generateCoachAccessCodes"
    )

    .addItem(
      "Open Dashboard",
      "openDashboard"
    )

    .addItem(
      "Reset Evaluation",
      "resetEvaluation"
    )

    .addToUi();

}