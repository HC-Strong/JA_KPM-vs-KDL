// Script written by Hannah Strong <stronghannahc@gmail.com> for James Atkins, July 2018
// Last edited: July 27, 2018


/** @OnlyCurrentDoc */


function onOpen() {
  
    var compareMenu = [{name:"Compare to KPM", functionName:"CompareToKPM"}];
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Comparison', compareMenu);
}






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
  Logger.log('URL is ' + kpmURL  + ' and ID is ' + kpmID + '.   It is ' + kpmIsURL + ' that the URL is a URL and ' + kpmIsLinked + ' that the KPM spreadsheet is linked');

  
  if(!kpmIsURL) {
    Browser.msgBox('No URL added in cell B1 of the Settings tab.   Please add the URL for the KPM sheet, connect the sheets (see the Settings tab for more info) then re-run this script.');
  } else if (!kpmIsLinked)
  {
    Browser.msgBox("The KPM sheet isn't connected.   Please connect the sheets (see the Settings tab for more info) then re-run this script.");
  } else {
    var curSheetName = sheet.getName();
    Logger.log("curSheetName is " + curSheetName);
    sheet.getRange(sheetNameCell).setValue(curSheetName);
    
    var importHeaders = importSheet.getRange(importHeadersRange).getValues();
    
    var headerExists = false;
    var blanksAllowed = 5;
    var blanks = 0;
    
    for (var i = 0; i < importHeaders[0].length; i++) {
      Logger.log("Checking column header: " + importHeaders[0][i]);
      if(importHeaders[0][i] == curSheetName) {
        headerExists = true;
        var emptyCol = i+2;
        Logger.log("emptyCol is " + emptyCol);
        break;
      } else if (!importHeaders[0][i]){
        blanks++;
        //Logger.log("This is #" + blanks + " of a max allowed " + blanksAllowed + " blank cells before the look exits");
        if (blanks >= blanksAllowed) {
          Logger.log("exiting after checking " + blanks + " blank column headers");
          break;
        }
      }
    }
      
    if (!headerExists) {
    
      emptyCol = FindFirstEmpty(importHeaders[0]);
      var importRangeFormula = '=arrayformula(regexreplace(sort({importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmPatientsRange + '"), importRange(' + settingsSheetName + '!' +  kpmIdCell + ', ' + '"' + curSheetName + '!' + kpmAmountsRange + '")}), " \\(([0-9]+)\\)", ""))';

      importSheet.getRange(1, emptyCol).setValue('=' + curSheetName + '!' + sheetNameCell);
      importSheet.getRange(1, emptyCol+1).setValue('=' + curSheetName + '!' + sheetNameCell);
      importSheet.getRange(2, emptyCol).setValue(importRangeFormula);
    } else {
     Logger.log("Header already exists, skipping it and only checking for any items not found in KDL"); 
    }
    
      
      
    var kdlNameValues = sheet.getRange("M:M").getValues();
    
    var notFoundInKDL = FindKpmExclusives(importSheet.getRange(1, emptyCol,500, 1).getValues(), kdlNameValues);
    
    
   var colEnd = FindFirstEmpty2(kdlNameValues);
    Logger.log("Last row is " + colEnd);
    
    Logger.log(notFoundInKDL.length);
    sheet.getRange(colEnd, 2, notFoundInKDL.length, 1).setValues(notFoundInKDL);
    Logger.log("1");
    if(sheet.getRange(colEnd, 2, 1, 1).getValue() == "#REF!") {
       Browser.msgBox('Comparison Failed.   No tab named "' + curSheetName + '" found in KPM spreadsheet. Please ensure the tabs for each day have identical names in the KPM and KDL spreadsheets. Then re-run this script.'); 
    } else {
      Browser.msgBox('Comparison complete.   The names of any patients in KPM but not found in this daily log have been added at the bottom of the Patient column. No other data has been added for these patients.');
    }
  }
}






function FindFirstEmpty(array) {
  for (var i = 0; i < array.length; i++) {
    if(array[i].length < 1) {
      return i+2;
    }
  }
}


function FindFirstEmpty2(array) {

  for (var i = 0; i < array.length; i++) {
      //Logger.log("it is: " + array[i][0].length);
    if(array[i][0].length < 1) {
      return i+1;
    }
  }
}






// gets arrays of KPM (imported) and KDL names and checks each KPM entry to see if it's in KDL. If it's not, it's added to the output array
function FindKpmExclusives(kpmPatients, kdlPatients){ 
  
  var notFoundRaw = [];
  var notFound = [];
  var checkMax = 200;   // set max number of rows for patients per day. James said 150 max so using 200 to be safe
  
  for (var i = 1; i < checkMax; i++) {
    var curKpmPatient = kpmPatients[i][0];
    
    notFoundRaw.push([curKpmPatient]);  //add to array to delete later if found
    
    if(curKpmPatient.length > 1) {
      for (var j = 1; j < checkMax; j++) {
        if(curKpmPatient.toUpperCase() == kdlPatients[j][0].toUpperCase()) {
          notFoundRaw.pop();
          //Logger.log("Found in KDL");
          break;
        }
      }
    }
  }
  Logger.log("Total # of patient names not found in KDL: " + notFoundRaw.length);
  Logger.log("The final result of not found names, before modifying to match KDL format, is: " + notFoundRaw);
  
    var blanksAllowed = 5;
    var blanks = 0;
  
    for (var i = 0; i < notFoundRaw.length; i++) {
      if(notFoundRaw[i][0].length > 0) {
        Logger.log("yes, it is " + notFoundRaw[i][0] + " and the length is " + notFoundRaw[i][0].length);
        var pieces = notFoundRaw[i][0].split(", ");
        var formatted = pieces[1] + " " + pieces[0];
        notFound.push([formatted]);
      } else {
        blanks++;
        Logger.log("no, it's " + notFoundRaw[i][0] + " and the length is " + notFoundRaw[i][0].length + ", and blanks are " + blanks);
        if (blanks >= blanksAllowed) {
          Logger.log("exiting after checking " + blanks + " blank rows");
          break;
        }
 
      }
      //notFoundRaw[i] = notFoundRaw[i] + "test";
    }
  
  Logger.log("And the final, final result of not found names is: " + notFound);
  
  return notFound;
}