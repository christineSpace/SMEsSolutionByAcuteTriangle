function showNavigationDialog() {
  var html = HtmlService.createHtmlOutputFromFile('HRnavigationDialog')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Navigation Dialog');
}

function openUrl(url) {
  var html = HtmlService.createHtmlOutput('<script>window.close();window.open("' + url + '");</script>')
      .setWidth(100)
      .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Open URL');
}

function onOpenhr() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HR')
      .addItem('Open HR navigation', 'showNavigationDialog')
      .addToUi();
}
