function setupLeagueWorkbook() {

 const ss = SpreadsheetApp.getActiveSpreadsheet();


 createSheet(
   ss,
   "Settings",
   [
     "Setting",
     "Value"
   ]
 );


 createSheet(
   ss,
   "Teams",
   [
     "Team ID",
     "Team Name",
     "Division",
     "Active"
   ]
 );


 createSheet(
   ss,
   "Coaches",
   [
     "Coach ID",
     "Coach Name",
     "Team ID",
     "Team Name",
     "Access Code",
     "Active"
   ]
 );


 createSheet(
   ss,
   "Players",
   [
     "Player ID",
     "Team ID",
     "First Name",
     "Last Name",
     "Grade"
   ]
 );


 createSheet(
   ss,
   "Evaluations",
   [
     "Timestamp",
     "Season",
     "Coach",
     "Team ID",
     "Team Name",
     "Player ID",
     "Player Name",
     "Fielding",
     "Hitting",
     "Throwing",
     "Running",
     "Game Concepts",
     "Attitude",
     "Pitcher",
     "Catcher",
     "Overall",
     "Comments"
   ]
 );


 const settings =
   ss.getSheetByName("Settings");


 settings
   .getRange("A2:B3")
   .setValues([
     [
       "Season",
       "2026"
     ],
     [
       "League Name",
       "Your League Name"
     ]
   ]);


 SpreadsheetApp
   .getUi()
   .alert(
     "League database setup complete!"
   );

}



function createSheet(ss, name, headers) {

 let sheet =
   ss.getSheetByName(name);


 if (!sheet) {

   sheet =
     ss.insertSheet(name);

 }


 if (sheet.getLastRow() === 0) {

   sheet
     .getRange(
       1,
       1,
       1,
       headers.length
     )
     .setValues([
       headers
     ]);

 }


 sheet
   .getRange(
     1,
     1,
     1,
     headers.length
   )
   .setFontWeight("bold");


 sheet.autoResizeColumns(
   1,
   headers.length
 );

}

