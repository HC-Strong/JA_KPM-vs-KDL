// Script written by Hannah Strong <stronghannahc@gmail.com> for James Atkins, July 2018
// Last edited: July 25, 2018


/** @OnlyCurrentDoc */


function onOpen() {
  
    var compareMenu = [{name:"Compare to KPM", functionName:"CompareToKPM"}];
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Comparison', compareMenu);
}


function CompareToKPM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var settingsSheet = ss.getSheetByName('Settings');
  
  var kpmID = settingsSheet.getRange("A2").getValue();
  
  Logger.log(kpmID);
}
