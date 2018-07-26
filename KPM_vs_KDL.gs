// Script written by Hannah Strong <stronghannahc@gmail.com> for James Atkins, July 2018
// Last edited: July 25, 2018


/** @OnlyCurrentDoc */


function onOpen() {
  
    var compareMenu = [{name:"Compare to KPM", functionName:"CompareToKPM"}];
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Comparison', compareMenu);
}

// Gets 
function CompareToKPM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var settingsSheetName = 'Settings';
  
  var settingsSheet = ss.getSheetByName(settingsSheetName);
  var importSheet = ss.getSheetByName('KPM Data');
  
  var kpmInfo = settingsSheet.getRange("B1:B5").getValues();
  var kpmIdCell = "B3";
  
  var kpmNamesRange = "L:L";
  var kpmAmountsRange = "X:X";
  
  var sheetNameCell = "Q1";
  
  var importHeadersRange = "B1:BZ1";
  
  
  ////// All Parameters above this ///////
  
  var kpmURL = kpmInfo[0][0];
  var kpmIsURL = kpmInfo[1][0];
  var kpmID = kpmInfo[2][0];
  var kpmIsLinked = kpmInfo[4][0];
  Logger.log("URL is " + kpmURL  + " and ID is " + kpmID + ".   It is " + kpmIsURL + " that the URL is a URL and " + kpmIsLinked + " that the KPM spreadsheet is linked");

  
  if(!kpmIsURL)
  {
    Browser.msgBox("No URL added in cell B1 of the Settings tab. Please add the URL for the KPM sheet, connect the sheets (see the Settings tab for more info) then re-run this script.");
  } else if (!kpmIsLinked)
  {
    Browser.msgBox("The KPM sheet isn't connectd. Please connect the sheets (see the Settings tab for more info) then re-run this script.");
  } else {
    var curSheetName = sheet.getName();
    sheet.getRange(sheetNameCell).setValue(curSheetName);
    
    var importHeaders = importSheet.getRange(importHeadersRange).getValues();
    
    var emptyCol = FindFirstEmpty(importHeaders[0]);
    var importRangeFormula = '=arrayformula(regexreplace(sort({importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmNamesRange + '"), importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmAmountsRange + '")}), " \\(([0-9]+)\\)", ""))';

    importSheet.getRange(1, emptyCol).setValue('=' + curSheetName + '!' + sheetNameCell);
    importSheet.getRange(1, emptyCol+1).setValue('=' + curSheetName + '!' + sheetNameCell);
    importSheet.getRange(2, emptyCol).setValue(importRangeFormula);
    
    Browser.msgBox("Script completed. Status if patients updated but patients only in the KPM spreadsheet not yet added.");
  }
}


function FindFirstEmpty(array) {
  for (var i = 0; i < array.length; i++) {
    if(array[i].length < 1) {
      return i+2;
    }
  }
}
