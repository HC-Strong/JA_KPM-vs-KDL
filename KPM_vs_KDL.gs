// Script written by Hannah Strong <stronghannahc@gmail.com> for James Atkins, July 2018
// Last edited: July 26, 2018


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
  
  var kpmPatientsRange = "L:L";
  var kpmAmountsRange = "X:X";
  
  var sheetPatientsRange = "B:B";
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
    var importRangeFormula = '=arrayformula(regexreplace(sort({importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmPatientsRange + '"), importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmAmountsRange + '")}), " \\(([0-9]+)\\)", ""))';

    importSheet.getRange(1, emptyCol).setValue('=' + curSheetName + '!' + sheetNameCell);
    importSheet.getRange(1, emptyCol+1).setValue('=' + curSheetName + '!' + sheetNameCell);
    importSheet.getRange(2, emptyCol).setValue(importRangeFormula);
    
    var notFoundInKDL = FindKpmExclusives(importSheet.getRange(1, emptyCol,500, 1).getValues(), sheet.getRange("M:M").getValues());
    
    
   // FindFirstEmpty(
    
    importSheet.getRange(12, 1, notFoundInKDL.length, 1).setValues(notFoundInKDL);
    
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




function oldFindKpmExclusives(importSheet, importCol, kdlSheet, kdlRange){ // gets arrays of KPM (imported) and KDL names and checks each KPM entry to see if it's in KDL. If it's not, it's added to the output array
  
  var notFound = [[]];
  var kpmPatients = importSheet.getRange(1, importCol,500, 1).getValues();
  var kdlPatients = kdlSheet.getRange(kdlRange).getValues();
  
  for (var i = 1; i < 200; i++) {
   var curKpmPatient = kpmPatients[i][0];
    //Logger.log(curKpmPatient);
    
    for (var j = 1; j < 200; j++) {
      if(curKpmPatient == kdlPatients[j][0]) {
        break;
      }
    }
    notFound[0].push(curKpmPatient);
  }
  Logger.log(notFound);
}


function FindKpmExclusives(kpmPatients, kdlPatients){ // gets arrays of KPM (imported) and KDL names and checks each KPM entry to see if it's in KDL. If it's not, it's added to the output array
  
  var notFound = [];
  //var kpmPatients = importSheet.getRange(1, importCol,500, 1).getValues();
  //var kdlPatients = kdlSheet.getRange(kdlRange).getValues();
  
  
  for (var i = 1; i < 5; i++) {
    var curKpmPatient = kpmPatients[i][0];
    Logger.log(curKpmPatient);
    
    notFound.push([curKpmPatient]);  //add to array to delete later if found
    Logger.log(notFound + " < after temp add");
    
    if(curKpmPatient.length > 1) {
      for (var j = 1; j < 5; j++) {
        if(curKpmPatient.toUpperCase() == kdlPatients[j][0].toUpperCase()) {
          notFound.pop();
          Logger.log("Found in KDL");
          Logger.log(notFound);
          break;
        }
      }
    }
  }
      Logger.log("Final result of not found names: " + notFound);
  return notFound;
}