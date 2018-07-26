/** @OnlyCurrentDoc */

function FormatDailyLog() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:B').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('C:C').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('D:E').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('E:E').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getActiveSheet().setColumnWidth(4, 559);
  spreadsheet.getActiveSheet().setColumnWidth(11, 177);
  spreadsheet.getRange('E2:J66').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('"$"#,##0.00');
  spreadsheet.getRange('J3').activate();
  spreadsheet.getCurrentCell().setFormula('=SUM(e3:i3)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('J3:J62'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('J3:J62').activate();
  spreadsheet.getActiveSheet().setColumnWidth(2, 175);
  spreadsheet.getRange('2:2').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};