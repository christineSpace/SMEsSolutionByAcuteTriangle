function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Leave Requests')
    .addItem('Open Dialog', 'openDialog')
    .addToUi();
}

function openDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Leave Request Action');
}

function handleFormSubmission(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = form.row;
  var action = form.action;

  if (action === 'approve') {
    sheet.getRange(row, 9).setValue('Approved');
    var email = sheet.getRange(row, 2).getValue();
    sendEmailNotification(email, 'Approved');
    addWholeDayEventToCalendar(sheet, row);
  } else if (action === 'reject') {
    sheet.getRange(row, 9).setValue('Rejected');
    var email = sheet.getRange(row, 2).getValue();
    sendEmailNotification(email, 'Rejected');
  }

  return 'Success';
}

function sendEmailNotification(emailAddress, status) {
  var subject = 'Leave Request Status Update';
  var message = 'Dear Employee,\n\n' +
                'We would like to inform you that your leave request has been ' + status.toLowerCase() + '.\n\n' +
                'If you have any questions or require further information, please do not hesitate to contact HR.\n\n' +
                'Thank you,\n' +
                'HR Department';

  GmailApp.sendEmail(emailAddress, subject, message);
  Logger.log("Email sent to: " + emailAddress);
}

function addWholeDayEventToCalendar(sheet, row) {
  var name = sheet.getRange(row, 3).getValue();
  var startDate = sheet.getRange(row, 6).getValue();
  var numberOfDays = sheet.getRange(row, 7).getValue();
  var comments = sheet.getRange(row, 10).getValue();

  var calendarId = 'jodiebeh0822@1utar.my'; // Replace with your specific calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);
  var start = new Date(startDate);
  var end = new Date(start);
  end.setDate(end.getDate() + numberOfDays);

  var description = 'Leave approved for ' + name;
  if (comments) {
    description += '\nRemarks: ' + comments;
  }

  var event = calendar.createAllDayEvent(name + ' - Leave', start, end, {
    description: description
  });

  Logger.log('Event created: ' + event.getId() + ' for ' + name + ' from ' + start + ' to ' + end);
}