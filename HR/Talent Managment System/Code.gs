function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Applicant Tracking')
    .addItem('Open Dialog', 'openDialog')
    .addToUi();
}

function openDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Applicant Tracking');
}

function handleFormSubmission(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = form.row;
  var action = form.action;

  if (action === 'initialReview') {
    sheet.getRange(row, 15).setValue("Yes");
    scheduleInterview(sheet, row);
  } else if (action === 'offerMade') {
    sheet.getRange(row, 17).setValue("Yes");
    sendOfferEmail(sheet, row);
  } else if (action === 'offerAccepted') {
    sheet.getRange(row, 18).setValue("Yes");
    sendOnboardingEmail(sheet, row);
  }

  return 'Success';
}

function scheduleInterview(sheet, row) {
  var applicantName = sheet.getRange(row, 3).getValue();
  var applicantEmail = sheet.getRange(row, 2).getValue();

  var interviewDate = new Date();
  interviewDate.setDate(interviewDate.getDate() + 2);
  interviewDate.setHours(14, 0, 0, 0);
  var endTime = new Date(interviewDate.getTime() + (60 * 60 * 1000));

  var timeZone = Session.getScriptTimeZone();
  var calendar = CalendarApp.getDefaultCalendar();
  var event = calendar.createEvent('Interview with ' + applicantName,
                                   interviewDate,
                                   endTime,
                                   {
                                     description: 'Interview with ' + applicantName,
                                     location: 'Google Meet Link: meet.google.com/efw-fore-edk',
                                     timeZone: timeZone
                                   });

  sheet.getRange(row, 16).setValue(interviewDate);
  var subject = "Interview Scheduled - Acute Triangle Minimarket";
  var message = "Dear " + applicantName + ",\n\n" +
                "Your interview has been scheduled for " + interviewDate.toDateString() + " at 2:00 PM.\n" +
                "Google Meet Link: meet.google.com/efw-fore-edk\n\n" +
                "Best regards,\n" +
                "Acute Triangle Minimarket Recruitment Team";

  GmailApp.sendEmail(applicantEmail, subject, message);
}

function sendOfferEmail(sheet, row) {
  var applicantName = sheet.getRange(row, 3).getValue();
  var applicantEmail = sheet.getRange(row, 2).getValue();
  var position = sheet.getRange(row, 5).getValue();

  var currentDate = new Date();
  var startDate = new Date(currentDate);
  startDate.setDate(startDate.getDate() + 21);
  var deadlineDate = new Date(currentDate);
  deadlineDate.setDate(deadlineDate.getDate() + 14);

  var formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM dd, yyyy");
  var formattedDeadlineDate = Utilities.formatDate(deadlineDate, Session.getScriptTimeZone(), "MMMM dd, yyyy");

  var subject = "Offer Letter from Acute Triangle Minimarket";
  var message = "Dear " + applicantName + ",\n\n" +
                "We are delighted to extend an offer for the position of " + position + " at Acute Triangle Minimarket. " +
                "Your start date will be " + formattedStartDate + ". Please review the attached offer letter and confirm by " + formattedDeadlineDate + ".\n\n" +
                "Best regards,\n" +
                "Acute Triangle Minimarket";

  GmailApp.sendEmail(applicantEmail, subject, message);
}

function sendOnboardingEmail(sheet, row) {
  var applicantName = sheet.getRange(row, 3).getValue();
  var applicantEmail = sheet.getRange(row, 2).getValue();

  var htmlContent = "<html><body>" +
                    "<h1>Welcome to Acute Triangle Minimarket!</h1>" +
                    "<p>We are excited to have you join our team. This document will guide you through the onboarding process and provide important information you need to get started.</p>" +
                    "<h2>Step 1: Employee Detail</h2>" +
                    "<p>Please fill in the google form link: <a href='https://forms.gle/2iYYC2u7URUcZSnSA'>Google Form</a></p>" +
                    "<h2>Step 2: Company Policies</h2>" +
                    "<p><strong>Punctuality:</strong> Arrive on time for all scheduled shifts.<br>" +
                    "<strong>Notification:</strong> Inform your manager in advance if you are unable to attend work.<br>" +
                    "<strong>Uniform:</strong> Wear the provided company uniform.</p>" +
                    "<h2>Step 3: Employee Benefits</h2>" +
                    "<p><strong>Health Insurance</strong><br>" +
                    "<strong>Paid Time Off</strong><br>" +
                    "14 days of paid vacation annually.<br>" +
                    "7 days of paid sick leave.<br>" +
                    "Public holidays off.<br>" +
                    "<strong>Employee Discounts</strong><br>" +
                    "20% discount on all store products.</p>" +
                    "<h2>Step 4: Important Contacts</h2>" +
                    "<p>Branch Manager: Jane Doe, jane.doe@trianglemart.com, +60123456789</p>" +
                    "<h2>Step 5: Acknowledgment</h2>" +
                    "<p>I acknowledge that I have received and reviewed the onboarding document.</p>" +
                    "<p>Employee Signature: _____________________________<br>" +
                    "Date: _____________________________</p>" +
                    "</body></html>";

  var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('Onboarding Document.pdf');

  var subject = "Welcome to Acute Triangle Minimarket!";
  var message = "Dear " + applicantName + ",\n\n" +
                "Congratulations on accepting our offer! We are excited to have you join our team. " +
                "Please find attached the onboarding document which will guide you through the onboarding process.\n\n" +
                "Best regards,\n" +
                "Acute Triangle Minimarket";

  GmailApp.sendEmail(applicantEmail, subject, message, {
    attachments: [blob]
  });
}