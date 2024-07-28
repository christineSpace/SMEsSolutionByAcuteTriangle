function onOpenexpenses() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Expenses Menu')
    .addItem('Open Dashboard', 'openDashboardLink')
    .addItem('Summarize Expenses', 'summarizeExpenses')
    .addToUi();
}

function openDashboardLink() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center;">
      <h2>Open Looker Studio Dashboard</h2>
      <p>Pls use personal account to open it</p>
      <p><a href="https://lookerstudio.google.com/reporting/08da9c78-21b1-48bf-bd1c-c8650a7990a2" target="_blank" 
            style="display: inline-block; padding: 10px 20px; margin: 10px; background-color: #4285F4; color: white; text-decoration: none; border-radius: 5px;">
            Open Looker Studio Dashboard</a></p>
      <p><button onclick="google.script.host.close()" style="padding: 10px 20px; background-color: #f44336; color: white; border: none; border-radius: 5px;">Close</button></p>
    </div>
  `)
  .setWidth(400)
  .setHeight(250);
  
  ui.showModalDialog(html, 'Dashboard Link');
}

function summarizeExpenses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses');
  if (!sheet) {
    Logger.log('No sheet named "Expenses" found.');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('No data found in the "Expenses" sheet.');
    return;
  }
  
  Logger.log('Data read from sheet:');
  Logger.log(data);
  
  var summary = {};
  
  for (var i = 1; i < data.length; i++) {
    var category = data[i][1];
    var amount = data[i][3]; // Amount is already a number
    
    if (!summary[category]) {
      summary[category] = 0;
    }
    summary[category] += amount;
  }
  
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary');
  if (!summarySheet) {
    summarySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Summary');
  }
  summarySheet.clear();
  summarySheet.appendRow(['Category', 'Total Amount']);
  
  for (var category in summary) {
    summarySheet.appendRow([category, 'RM' + summary[category].toFixed(2)]);
  }
  
  Logger.log('Summary created:');
  Logger.log(summary);
}
