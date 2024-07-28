function importContacts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Contact Information (Responses)');
  if (!sheet) {
    Logger.log('Sheet not found.');
    return;
  }

  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var email = row[1];
    var fullName = row[2].split(' ');
    var firstName = fullName[0];
    var lastName = fullName.slice(1).join(' ');
    var homeAddress = row[3];
    var phoneNumber = row[4];
    var dateOfBirth = row[5];
    var jobTitle = row[6];
    var employeeId = row[7];
    var startDate = row[8];
    var emergencyContactName = row[9];
    var relationship = row[10];
    var emergencyContactPhone = row[11];
    var emergencyContactEmail = row[12];

    // Search for existing contact by email
    var contacts = ContactsApp.getContactsByEmailAddress(email);
    var contact;

    if (contacts.length > 0) {
      contact = contacts[0];
      Logger.log('Updating existing contact: ' + firstName + ' ' + lastName);
    } else {
      contact = ContactsApp.createContact(firstName, lastName, email);
      Logger.log('Created new contact: ' + firstName + ' ' + lastName);
    }

    // Add or update additional fields
    if (homeAddress) {
      if (contact.getAddresses(ContactsApp.Field.HOME_ADDRESS).length > 0) {
        contact.getAddresses(ContactsApp.Field.HOME_ADDRESS)[0].setAddress(homeAddress);
      } else {
        contact.addAddress(ContactsApp.Field.HOME_ADDRESS, homeAddress);
      }
      Logger.log('Set home address: ' + homeAddress);
    }
    if (phoneNumber) {
      if (contact.getPhones(ContactsApp.Field.MOBILE_PHONE).length > 0) {
        contact.getPhones(ContactsApp.Field.MOBILE_PHONE)[0].setPhoneNumber(phoneNumber);
      } else {
        contact.addPhone(ContactsApp.Field.MOBILE_PHONE, phoneNumber);
      }
      Logger.log('Set phone number: ' + phoneNumber);
    }

    var notes = [];
    if (dateOfBirth) notes.push('Date of Birth: ' + dateOfBirth);
    if (jobTitle) notes.push('Job Title: ' + jobTitle);
    if (employeeId) notes.push('Employee ID: ' + employeeId);
    if (startDate) notes.push('Start Date of Employment: ' + new Date(startDate).toDateString());
    if (emergencyContactName) notes.push('Emergency Contact Name: ' + emergencyContactName);
    if (relationship) notes.push('Relationship: ' + relationship);
    if (emergencyContactPhone) notes.push('Emergency Contact Phone: ' + emergencyContactPhone);
    if (emergencyContactEmail) notes.push('Emergency Contact Email: ' + emergencyContactEmail);

    if (notes.length > 0) {
      var existingNotes = contact.getNotes();
      contact.setNotes((existingNotes ? existingNotes + '\n' : '') + notes.join('\n'));
      Logger.log('Set notes: ' + notes.join(', '));
    }
  }
  
  Logger.log('Contacts Imported Successfully');
}

function showNavigationDialog() {
  var html = HtmlService.createHtmlOutputFromFile('navigationDialog')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Navigation Dialog');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Navigation Menu')
      .addItem('Open Navigation Dialog', 'showNavigationDialog')
      .addToUi();
}
