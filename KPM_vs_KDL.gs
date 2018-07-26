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
    Browser.msgBox("not URL, and you'll need to connect it");
  } else if (!kpmIsLinked)
  {
    Browser.msgBox("you'll need to connect it");
  } else {
    var curSheetName = sheet.getName();
    sheet.getRange(sheetNameCell).setValue(curSheetName);
    
    var importHeaders = importSheet.getRange(importHeadersRange).getValues();
    
    var emptyCol = FindFirstEmpty(importHeaders[0]);
    var importRangeFormula = '=arrayformula(regexreplace(sort({importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmNamesRange + '"), importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmAmountsRange + '")}), " \\(([0-9]+)\\)", ""))';
    //var importRangeFormula = '=sort({importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmNamesRange + '"), importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmAmountsRange + '")})';
    //var importRangeFormula = '=sort(importRange("' + kpmID + '", ' + '"July19!L:L"))';
    
    //Browser.msgBox(importRangeFormula);
    
    importSheet.getRange(1, emptyCol).setValue('=' + curSheetName + '!' + sheetNameCell);
    importSheet.getRange(2, emptyCol).setValue(importRangeFormula);
    
    Browser.msgBox("placeholder for adding a) finding where to put importRange formulas and b) adding them");
    Browser.msgBox("done");
  }
}


function FindFirstEmpty(array) {
  for (var i = 0; i < array.length; i++) {
    if(array[i].length < 1) {
      return i+2;
    }
  }
}
