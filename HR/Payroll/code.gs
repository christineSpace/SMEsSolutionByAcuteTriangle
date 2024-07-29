function formatPayroll() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();

  let headers = sheet.getRange('A1:P1');
  let table = sheet.getDataRange();

  headers.setFontWeight('bold');
  headers.setFontColor('white');
  headers.setBackground('#52489C');

  table.setFontFamily('Roboto');
  table.setHorizontalAlignment('center');
  table.setBorder(true,true,true,true,false,true,'#52489C',SpreadsheetApp.BorderStyle.SOLID);

  //table.createFilter();

}

function createAndSendSalarySlip(){
  var empId = "";
  var empName = "";
  var empEmail = "";
  var month="";

  var noOfDayWorked = 0;
  var noOfHolidayWorked = 0;
  var noHourOT = 0;

  var basicSalary = 0;
  var holidayAllowance = 0;
  var otAllowance = 0;
  var totalIncome = 0;

  var epf = 0;
  var socso = 0;
  var totalDeduct = 0;
  var netSalary =0;

  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var salSheet = spSheet.getSheetByName('Payroll_Details');

  var salaryDetailFolder = DriveApp.getFolderById('1S3aIHmVLPrKjF1wYYN7rE4-dpUFOtEyo');
  var salaryTemplate =  DriveApp.getFileById('1Qekesszz6i5_CbCQOP2DOTEKihWKBD83tZ1Y6qc23sM');

  var totalRows = salSheet.getLastRow();

  for(var rowNo=2; rowNo<=totalRows; rowNo++){
    empId = salSheet.getRange('A'+rowNo).getDisplayValue();
    empName = salSheet.getRange('B'+rowNo).getDisplayValue();
    empEmail = salSheet.getRange('C'+rowNo).getDisplayValue();
    month = salSheet.getRange('D'+rowNo).getDisplayValue();
    noOfDayWorked = salSheet.getRange('F'+rowNo).getDisplayValue();
    noOfHolidayWorked = salSheet.getRange('H'+rowNo).getDisplayValue();
    noHourOT = salSheet.getRange('J'+rowNo).getDisplayValue();
    basicSalary = salSheet.getRange('G'+rowNo).getDisplayValue();
    holidayAllowance = salSheet.getRange('I'+rowNo).getDisplayValue();
    otAllowance = salSheet.getRange('K'+rowNo).getDisplayValue();
    totalIncome = salSheet.getRange('L'+rowNo).getDisplayValue();
    epf = salSheet.getRange('M'+rowNo).getDisplayValue();
    socso = salSheet.getRange('N'+rowNo).getDisplayValue();
    totalDeduct = salSheet.getRange('O'+rowNo).getDisplayValue();
    netSalary = salSheet.getRange('P'+rowNo).getDisplayValue();

    var rawSalFile = salaryTemplate.makeCopy(salaryDetailFolder);
    var rawFile = DocumentApp.openById(rawSalFile.getId());
    var rawFileContent = rawFile.getBody();

    rawFileContent.replaceText("EMP_ID_XXXX", empId);
    rawFileContent.replaceText("EMP_NAME_XXXX", empName);
    rawFileContent.replaceText("EMP_EMAIL_XXXX", empEmail);
    rawFileContent.replaceText("MONTH_XXXX", month);
    rawFileContent.replaceText("DAY_WORKED_XXXX", noOfDayWorked);
    rawFileContent.replaceText("HOLI_WORKED_XXXX", noOfHolidayWorked);
    rawFileContent.replaceText("OT_HOURS_XXXX", noHourOT);
    rawFileContent.replaceText("BASIC_SAL_XXXX", basicSalary);
    rawFileContent.replaceText("HOLIDAY_XXXX", holidayAllowance);
    rawFileContent.replaceText("OT_XXXX", otAllowance);
    rawFileContent.replaceText("EPF_XXXX", epf);
    rawFileContent.replaceText("SOCSO_XXXX", socso);
    rawFileContent.replaceText("TOTAL_INCOME_XXXX", totalIncome);
    rawFileContent.replaceText("TOTAL_DEDUCT_XXXX", totalDeduct);
    rawFileContent.replaceText("NET_SALARY_XXXX", netSalary);

    rawFile.saveAndClose();
    var salSlip = rawFile.getAs(MimeType.PDF);
    var salPDF = salaryDetailFolder.createFile(salSlip).setName("Salary_" + empId +"_"+ month);

    DriveApp.getFolderById(salaryDetailFolder.getId()).removeFile(rawSalFile);
    
    var mailSubject = "Salary Slip";
    var mailBody = "Please find the salary slip for the month of " + month + " attached.";
    MailApp.sendEmail(empEmail, mailSubject, mailBody, {attachments: [salPDF.getAs(MimeType.PDF)]});
  }
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SalarySlip')
    .addItem('Create & Send Salary Slip','createAndSendSalarySlip')
    .addToUi();
}
