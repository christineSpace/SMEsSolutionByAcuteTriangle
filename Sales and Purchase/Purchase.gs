function onOpenpurchase() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Purchase Menus')
    .addSubMenu(ui.createMenu('Quotation Menu')
      .addItem('Check Quotations', 'checkStockAndNotify')
      .addItem('Send Quotations', 'sendConsolidatedQuotations')
      .addItem('Manual Quotation', 'showQuotationDialog'))
    .addSubMenu(ui.createMenu('Analyze Quotation')
      .addItem('Analyze and Send Report', 'analyzeAndSendReport')
      .addItem('Analyze and Update PO', 'analyzeAndUpdatePO'))
    .addSubMenu(ui.createMenu('PO Menu')
      .addItem('Send Purchase Orders', 'openDialog'))
    .addItem('Update Product Prices', 'updateProductPrices')
    .addItem('Update Receiving Goods', 'openLink')
    .addToUi();
}

//---------------------------------------Quotation----------------------------------------------------------
//got two option for send quotation= first is check the current quantity of stock, if low stock, it will automate send quotation to supplier
//or manual quotation
function getProducts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product'); // Replace 'Products' with your sheet name
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); // Adjust the range according to your sheet structure
  var products = {};
  
  data.forEach(function(row) {
    products[row[0]] = row[1];
  });

  return products;
}
function checkStockAndNotify() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productSheet = spreadsheet.getSheetByName("Product");
  var currentUserEmail = Session.getActiveUser().getEmail(); // Get the email of the current user

  var productData = productSheet.getDataRange().getValues();
  
  var subjectOwner = "Stock Reorder Notification";
  var bodyOwner = "The following products are below the minimum stock level and need to be reordered:\n\n";
  
  var productsToReorder = [];
  
  for (var i = 1; i < productData.length; i++) {
    var productId = productData[i][0];
    var productName = productData[i][1];
    var currentStock = productData[i][4];
    var minStockLevel = productData[i][5];
    var reorderQuantity = productData[i][6];
    
    if (currentStock < minStockLevel) {
      bodyOwner += "Product ID: " + productId + "\n" +
                   "Product Name: " + productName + "\n" +
                   "Current Stock: " + currentStock + "\n" +
                   "Minimum Stock Level: " + minStockLevel + "\n" +
                   "Reorder Quantity: " + reorderQuantity + "\n\n";
      productsToReorder.push([productId, productName, reorderQuantity]);
    }
  }
  
  if (productsToReorder.length > 0) {
    try {
      MailApp.sendEmail(currentUserEmail, subjectOwner, bodyOwner);
      Logger.log('Notification sent to owner.');
    } catch (e) {
      Logger.log('Failed to send notification to owner. Error: ' + e.message);
    }
  }
}

function sendConsolidatedQuotations() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productSheet = spreadsheet.getSheetByName("Product");
  var supplierSheet = spreadsheet.getSheetByName("Suppliers");

  var productData = productSheet.getDataRange().getValues();
  var supplierData = supplierSheet.getDataRange().getValues();
  
  var productsToReorder = [];
  
  for (var i = 1; i < productData.length; i++) {
    var productId = productData[i][0];
    var productName = productData[i][1];
    var currentStock = productData[i][4];
    var minStockLevel = productData[i][5];
    var reorderQuantity = productData[i][6];
    
    if (currentStock < minStockLevel) {
      productsToReorder.push([productId, productName, reorderQuantity]);
    }
  }
  
  if (productsToReorder.length > 0) {
    var googleFormLink = "https://docs.google.com/forms/d/e/1FAIpQLSfl_Ie0F8kNTlNYmbDgnwwVFurnUcuernEIkpSuv_rSO6lYjg/viewform?usp=sf_link";
    var subject = "Request for Quotation for Multiple Products";
    
    var body = "Dear Supplier,\n\n" +
               "We would like to request a quotation for the following products:\n\n" +
               "<table border='1' style='border-collapse: collapse; width: 100%;'>" +
               "<tr>" +
               "<th>Product ID</th><th>Product Name</th><th>Quantity Needed</th>" +
               "</tr>";
    
    for (var i = 0; i < productsToReorder.length; i++) {
      var productId = productsToReorder[i][0];
      var productName = productsToReorder[i][1];
      var reorderQuantity = productsToReorder[i][2];
      
      body += "<tr>" +
              "<td>" + productId + "</td>" +
              "<td>" + productName + "</td>" +
              "<td>" + reorderQuantity + "</td>" +
              "</tr>";
    }
    
    body += "</table><br><br>" +
            "Please submit your quotation using the following Google Form:<br>" +
            "<a href='" + googleFormLink + "'>" + googleFormLink + "</a><br><br>" +
            "Thank you,<br>" +
            "Acute Triangle Mini Market";
    
    for (var j = 1; j < supplierData.length; j++) {
      var email = supplierData[j][4];
      if (email) {
        try {
          MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: body
          });
          Logger.log('Email sent to: ' + email);
        } catch (e) {
          Logger.log('Failed to send email to: ' + email + ' Error: ' + e.message);
        }
      } else {
        Logger.log('Empty email field for supplier: ' + supplierData[j][1]);
      }
    }
  }
}

function showQuotationDialog() {
  var html = HtmlService.createHtmlOutputFromFile('QuotationForm')
      .setWidth(600)
      .setHeight(400);
  SpreadsheetApp.getUi().showDialog(html);
}


function processQuotationForm(products) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var supplierSheet = spreadsheet.getSheetByName("Suppliers");
  var googleFormLink = "https://docs.google.com/forms/d/e/1FAIpQLSfl_Ie0F8kNTlNYmbDgnwwVFurnUcuernEIkpSuv_rSO6lYjg/viewform?usp=sf_link";
  
  var supplierData = supplierSheet.getDataRange().getValues();
  
  var subject = "Request for Quotation";
  var body = "Dear Supplier,\n\n" +
             "We would like to request a quotation for the following products:\n\n" +
             "<table border='1' style='border-collapse: collapse; width: 100%;'>" +
             "<tr>" +
             "<th>Product ID</th><th>Product Name</th><th>Quantity Needed</th>" +
             "</tr>";
  
  products.forEach(function(product) {
    body += "<tr>" +
            "<td>" + product.productId + "</td>" +
            "<td>" + product.productName + "</td>" +
            "<td>" + product.quantity + "</td>" +
            "</tr>";
  });
  
  body += "</table><br><br>" +
          "Please use the following link to submit your quotation:<br>" +
          "<a href='" + googleFormLink + "'>" + googleFormLink + "</a><br><br>" +
          "Thank you,<br>" +
          "Acute Triangle Mini Market";
  
  for (var j = 1; j < supplierData.length; j++) {
    var email = supplierData[j][4];
    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: body
        });
        Logger.log('Email sent to: ' + email);
      } catch (e) {
        Logger.log('Failed to send email to: ' + email + ' Error: ' + e.message);
      }
    } else {
      Logger.log('Empty email field for supplier: ' + supplierData[j][1]);
    }
  }
}
//--------------------------------------analyze the supplier and then update po database----------------------

function analyzeAndSendReport() {
  var analysis = analyzeSuppliers();
  var chosenSuppliers = chooseBestSuppliers(analysis);
  sendAnalysisReportToOwner(analysis, chosenSuppliers);
}

function analyzeAndUpdatePO() {
  var analysis = analyzeSuppliers();
  var chosenSuppliers = chooseBestSuppliers(analysis);
  updatePOWithChosenSuppliers(chosenSuppliers);
}

function analyzeSuppliers() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = spreadsheet.getSheetByName("Quotation");
  var data = tempSheet.getDataRange().getValues();
  
  Logger.log("Data from Quotation Sheet: " + JSON.stringify(data));

  var analysis = {};
  
  for (var i = 1; i < data.length; i++) {
    var productCode = data[i][6];
    var supplierName = data[i][1];
    var unitPrice = parseFloat(data[i][8]);
    var deliveryTimeframe = data[i][13];

    if (!analysis[productCode]) {
      analysis[productCode] = [];
    }
    analysis[productCode].push({ 
      supplierName: supplierName, 
      unitPrice: unitPrice, 
      deliveryTimeframe: deliveryTimeframe,
      response: data[i]
    });
  }

  Logger.log("Analysis Data: " + JSON.stringify(analysis));
  return analysis;
}

function chooseBestSuppliers(analysis) {
  var chosenSuppliers = {};
  for (var productCode in analysis) {
    // Convert delivery timeframes to numeric values for sorting
    analysis[productCode].sort(function(a, b) {
      var aTime = convertDeliveryTimeframeToNumber(a.deliveryTimeframe);
      var bTime = convertDeliveryTimeframeToNumber(b.deliveryTimeframe);
      return a.unitPrice - b.unitPrice || aTime - bTime;
    });
    
    // Ensure that the chosen supplier is correctly selected
    chosenSuppliers[productCode] = analysis[productCode][0].response;
  }

  Logger.log("Chosen Suppliers Data: " + JSON.stringify(chosenSuppliers));
  return chosenSuppliers;
}

function convertDeliveryTimeframeToNumber(timeframe) {
  var daysMap = {
    "Within 1 week": 7,
    "1-2 weeks": 14,
    "2-4 weeks": 28,
    "More than 4 weeks": 30
  };
  return daysMap[timeframe] || 30;
}

function updatePOWithChosenSuppliers(chosenSuppliers) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var poSheet = spreadsheet.getSheetByName("PurchaseOrder");
  var poiSheet = spreadsheet.getSheetByName("PurchaseOrderItem");
  var poData = poSheet.getDataRange().getValues();
  var poiData = poiSheet.getDataRange().getValues();

  // Function to get the last PO ID
  function getLastPOID() {
    if (poData.length > 1) {
      var lastPOID = poData[poData.length - 1][0];
      return parseInt(lastPOID.replace("PO", ""), 10);
    }
    return 0; // If no POs exist yet
  }

  // Function to get the last POI ID
  function getLastPOIID() {
    if (poiData.length > 1) {
      var lastPOIID = poiData[poiData.length - 1][0];
      return parseInt(lastPOIID.replace("POI", ""), 10);
    }
    return 0; // If no POIs exist yet
  }

  // Initialize the last PO and POI IDs
  var lastPOID = getLastPOID();
  var lastPOIID = getLastPOIID();

  // Function to generate a new unique PO ID
  function generateUniquePOID() {
    lastPOID++;
    return "PO" + lastPOID.toString().padStart(3, '0');
  }

  // Function to generate a new unique POI ID
  function generateUniquePOIID() {
    lastPOIID++;
    return "POI" + lastPOIID.toString().padStart(3, '0');
  }

  // Function to get Supplier ID from Supplier Name
  function getSupplierID(supplierName) {
    var supplierSheet = spreadsheet.getSheetByName("Suppliers"); // Assuming you have a "Suppliers" sheet
    var data = supplierSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === supplierName) { // Assuming supplier name is in column 2 (index 1)
        return data[i][0]; // Assuming supplier ID is in column 1 (index 0)
      }
    }
    
    Logger.log("Supplier ID for " + supplierName + " not found.");
    return null;
  }

  for (var productCode in chosenSuppliers) {
    var response = chosenSuppliers[productCode];
    
    var supplierName = response[1];
    var supplierID = getSupplierID(supplierName); // Get Supplier ID from Supplier Name
    var productName = response[5];
    var quantity = parseInt(response[7]);
    var unitPrice = parseFloat(response[8]);
    var totalPrice = parseFloat(response[9]);
    var paymentTerms = response[12];
    var deliveryTimeframe = response[13];
    var deliveryCharges = parseFloat(response[14]);
    var shippingMethod = response[15];
    var comments = response[18];

    // Auto-generate new PO ID for each product
    var poID = generateUniquePOID();
    var newPO = [
      poID,
      supplierID, // Use supplierID instead of supplierName
      new Date().toLocaleDateString(),
      totalPrice,
      "SST (6%)",
      totalPrice * 0.06,
      deliveryCharges,
      totalPrice + (totalPrice * 0.06) + deliveryCharges,
      paymentTerms,
      shippingMethod,
      comments
    ];
    poSheet.appendRow(newPO);

    // Auto-generate new POI ID for each product
    var poiID = generateUniquePOIID();
    var newPOI = [
      poiID,
      poID, // Reference to the newly created PO
      productCode,
      quantity,
      0, // Quantity Received (if applicable)
      unitPrice,
      totalPrice
    ];
    poiSheet.appendRow(newPOI);
  }
}

function sendAnalysisReportToOwner(analysis, chosenSuppliers) {
   var currentUserEmail = Session.getActiveUser().getEmail(); // Get the email of the current user
  var subject = "Supplier Analysis and Chosen Suppliers Report";
  
  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  var htmlBody = '<html><body>';
  htmlBody += "<p>Dear Owner,</p>";
  htmlBody += "<p>This report was generated on " + timestamp + ".</p>";
  htmlBody += "<p>Here is the analysis of suppliers for the recent quotation requests and the chosen suppliers for each product:</p>";

  // Iterate through each product code
  for (var productCode in analysis) {
    htmlBody += "<h3>Product Code: " + productCode + "</h3>";
    htmlBody += "<table border='1' cellpadding='5' cellspacing='0'>";
    htmlBody += "<tr><th>Supplier</th><th>Unit Price</th><th>Delivery Timeframe</th><th>Chosen</th></tr>";
    
    // Iterate through each supplier for the current product
    analysis[productCode].forEach(function(supplier) {
      var chosen = (JSON.stringify(chosenSuppliers[productCode]) === JSON.stringify(supplier.response)) ? "Yes" : "No";
      htmlBody += "<tr>";
      htmlBody += "<td>" + supplier.supplierName + "</td>";
      htmlBody += "<td>" + supplier.unitPrice + "</td>";
      htmlBody += "<td>" + supplier.deliveryTimeframe + "</td>";
      htmlBody += "<td>" + chosen + "</td>";
      htmlBody += "</tr>";
    });

    htmlBody += "</table>";
  }
  
  htmlBody += "<p>Best regards,<br/>Your Automated System</p>";
  htmlBody += "</body></html>";

  // Log the HTML body for debugging
  Logger.log("Email HTML Body: " + htmlBody);

  // Send the email with HTML content
  MailApp.sendEmail({
    to: currentUserEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}


function getSupplierID(supplierName) {
  var supplierSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suppliers"); // Assuming you have a "Suppliers" sheet
  var data = supplierSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === supplierName) { // Assuming supplier name is in column 2 (index 1)
      return data[i][0]; // Assuming supplier ID is in column 1 (index 0)
    }
  }
  
  Logger.log("Supplier ID for " + supplierName + " not found.");
  return null;
}

//-----------------------------send po-------------------------------------------------------------------
function openDialog() {
  const html = HtmlService.createHtmlOutputFromFile('po')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Purchase Orders');
}

function getPOList() {
  var ss = SpreadsheetApp.openById('1RWL-8SIeGhJ3G7JDMTnltguj0Tvin5f4vBHDRo06qDw');
  var poSheet = ss.getSheetByName('PurchaseOrder');
  var poData = poSheet.getRange(3, 1, poSheet.getLastRow() - 2, poSheet.getLastColumn()).getValues();
  
  var poList = poData.map(row => {
    return {
      poID: row[0],
      orderDate: row[2], // Assuming order date is in 3rd column
      supplier: row[1],  // Assuming supplier is in 2nd column
      totalAmount: row[7] // Assuming total amount is in 8th column
    };
  });
  
  return JSON.stringify(poList);
}

function sendPurchaseOrders(selectedIDs) {
  var ss = SpreadsheetApp.openById('1RWL-8SIeGhJ3G7JDMTnltguj0Tvin5f4vBHDRo06qDw');
  const poSheet = ss.getSheetByName('PurchaseOrder');
  const itemsSheet = ss.getSheetByName('PurchaseOrderItem');
  const productsSheet = ss.getSheetByName('Product');
  const suppliersSheet = ss.getSheetByName('Suppliers');

  const poData = poSheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();
  const productsData = productsSheet.getDataRange().getValues();
  const suppliersData = suppliersSheet.getDataRange().getValues();

  selectedIDs.forEach(poID => {
    const poRow = poData.find(row => row[0] === poID);
    if (!poRow) {
      Logger.log(`PO ID ${poID} not found in PurchaseOrder sheet.`);
      return; // Skip if PO ID not found in PurchaseOrder
    }

    const items = itemsData.filter(item => item[1] === poID);
    const services = items.map(item => {
      const product = productsData.find(prod => prod[0] === item[2]);
      const total = item && typeof item[6] === 'string'
        ? parseFloat(item[6].replace('RM', '').replace(',', ''))
        : parseFloat(item[6] || 0); // Fallback to 0 if conversion fails

      return {
        productID: item[2],
        listed: product ? product[1] : 'Unknown Product',
        UnitP: item[5],
        quantity: item[3],
        total: total
      };
    });

    let subtotal = services.reduce((acc, curr) => acc + curr.total, 0);
    let taxRate = poRow[4];
    let taxTotal = parseFloat(poRow[5]) || 0;
    let shippingFee = parseFloat(poRow[6]) || 0;
    let totalDue = parseFloat(poRow[7]) || 0;
    let paymentTerms = poRow[8];
    let shippingMethod = poRow[9];

    const supplier = suppliersData.find(supplier => supplier[0] === poRow[1]);
    if (!supplier) {
      Logger.log(`Supplier ID ${poRow[1]} not found in Suppliers sheet for PO ID ${poID}.`);
      return; // Skip if supplier not found
    }

    const supplierEmail = supplier[4];

    const doc = DocumentApp.create(`PurchaseOrder-${poID}`);
    const body = doc.getBody();
    body.setMarginTop(72); // 1 inch
    body.setMarginBottom(72);
    body.setMarginLeft(72);
    body.setMarginRight(72);

    // Document Header
    body.appendParagraph('Acute Triangle Mini Market')
        .setFontSize(16)
        .setBold(true)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('1234, Jln Maju, 31900, Kampar, Perak')
        .setFontSize(10)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('Phone: 012-345 6789')
        .setFontSize(10)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("");

    // Purchase Order Details
    body.appendParagraph(`Purchase Order #: ${poID}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Order Date: ${new Date(poRow[2]).toLocaleDateString()}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Vendor: ${supplier[1]}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    body.appendParagraph(`Email: ${supplierEmail}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    body.appendParagraph(`Phone: ${supplier[3]}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    body.appendParagraph(`Payment Terms: ${paymentTerms}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    body.appendParagraph(`Shipping Method: ${shippingMethod}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    body.appendParagraph("");

    // Products Table
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell('Product ID').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Product Name').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Quantity').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Unit Price (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Total (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    services.forEach(service => {
      const row = table.appendTableRow();
      row.appendTableCell(service.productID).setFontSize(10);
      row.appendTableCell(service.listed).setFontSize(10);
      row.appendTableCell(service.quantity.toString()).setFontSize(10);
      row.appendTableCell(`RM ${parseFloat(service.UnitP).toFixed(2)}`).setFontSize(10);
      row.appendTableCell(`RM ${service.total.toFixed(2)}`).setFontSize(10);
    });

    // Financial Summary
    body.appendParagraph(`Subtotal: RM ${subtotal.toFixed(2)}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Tax Rate: ${taxRate}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Tax Total: RM ${taxTotal.toFixed(2)}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Shipping: RM ${shippingFee.toFixed(2)}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Total Due: RM ${totalDue.toFixed(2)}`).setFontSize(10).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph("");

    // PDF Generation and Sharing
    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');
    const folder = DriveApp.getFoldersByName("PurchaseOrders").hasNext() ? DriveApp.getFoldersByName("PurchaseOrders").next() : DriveApp.createFolder("PurchaseOrders");
    let version = 1;
    let pdfFileName = `PurchaseOrder-${poID}_V${String(version).padStart(2, '0')}.pdf`;
    while (folder.getFilesByName(pdfFileName).hasNext()) {
      version++;
      pdfFileName = `PurchaseOrder-${poID}_V${String(version).padStart(2, '0')}.pdf`;
    }
    const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const pdfUrl = pdfFile.getUrl();

    // Send Email with PDF Attachment
    GmailApp.sendEmail(supplierEmail, `Purchase Order: ${poID}`, `Please find attached Purchase Order ${poID}.`, {
      attachments: [pdfBlob]
    });

    Logger.log(`Purchase Order PDF generated successfully for PO ID ${poID}. Version: ${version}. Link: ${pdfUrl}`);
  });

  return `Purchase Orders generated successfully for PO IDs: ${selectedIDs.join(', ')}.`;
}

//-----------------------------receiving goods--------------------------------------------------
//update the inventory when the receiving the goods through google form
function onFormSubmit(e) {
  var ss = SpreadsheetApp.openById('1RWL-8SIeGhJ3G7JDMTnltguj0Tvin5f4vBHDRo06qDw');
  var deliveryNotesSheet  = ss.getSheetByName('DeliveryNotes');
  var productsSheet  = ss.getSheetByName('Product');
  var poItemsSheet=ss.getSheetByName('PurchaseOrderItem')


  // Get the last form response
  var lastRow = deliveryNotesSheet.getLastRow();
  var response = deliveryNotesSheet.getRange(lastRow, 1, 1, deliveryNotesSheet.getLastColumn()).getValues()[0];

  // Extract form response values
  var purchaseOrderId = response[1];
  var productId = response[2];
  var quantityReceived = response[4];

  Logger.log("Purchase Order ID: " + purchaseOrderId);
  Logger.log("Product ID: " + productId);
  Logger.log("Quantity Received: " + quantityReceived);

  // Update the Product Database
  var productRange = productsSheet.getRange('A2:A' + productsSheet.getLastRow()).createTextFinder(productId).findNext();
  if (productRange) {
    var productRow = productRange.getRow();
    var currentQuantity = productsSheet.getRange(productRow, 5).getValue();
    productsSheet.getRange(productRow, 5).setValue(currentQuantity + quantityReceived);
    Logger.log("Updated Product Quantity: " + (currentQuantity + quantityReceived));
  } else {
    Logger.log("Product ID not found in Products sheet.");
  }

  // Update the PO Items Database
  var poItemRange = poItemsSheet.getRange('B2:B' + poItemsSheet.getLastRow()).createTextFinder(purchaseOrderId).findAll();
  for (var i = 0; i < poItemRange.length; i++) {
    var poItemRow = poItemRange[i].getRow();
    var poProductId = poItemsSheet.getRange(poItemRow, 3).getValue();
    if (poProductId === productId) {
      var currentReceived = poItemsSheet.getRange(poItemRow, 5).getValue();
      poItemsSheet.getRange(poItemRow, 5).setValue(currentReceived + quantityReceived);
      Logger.log("Updated PO Item Quantity Received: " + (currentReceived + quantityReceived));
      
      // Check if quantity received matches quantity ordered
      var quantityOrdered = poItemsSheet.getRange(poItemRow, 4).getValue();
      if (currentReceived + quantityReceived != quantityOrdered) {
        // Handle discrepancy (e.g., log the discrepancy, notify staff, etc.)
        Logger.log('Quantity received (' + (currentReceived + quantityReceived) + ') does not match quantity ordered (' + quantityOrdered + ') for Product ID ' + productId);
        
      // Get the email address of the sheet owner
        var ownerEmail = Session.getActiveUser().getEmail();
        Logger.log('Email Address of Sheet Owner: ' + ownerEmail);

        // Send email notification
        MailApp.sendEmail({
          to: ownerEmail, // Send to the sheet owner's email address
          subject: 'Discrepancy in Received Goods',
          body: 'Quantity received (' + (currentReceived + quantityReceived) + ') does not match quantity ordered (' + quantityOrdered + ') for Product ID ' + productId + ' in Purchase Order ' + purchaseOrderId
        });
      }
    }
  }
}

// Set the trigger for form submission
function createOnFormSubmitTrigger() {
  var form = FormApp.openById('1LvkPlvStzdLwaXd-nlENZkM66f2TpucWHkrZL37sVWo');
  ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();
}

function openLink() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://docs.google.com/forms/d/e/1FAIpQLSeVZ4wN5lDvemK32V_t8ZKu-tMsbiYnOvWz-GthY0NGwNTdkw/viewform?usp=sf_link" target="_blank">Click here to open the link</a>')
    .setWidth(250)
    .setHeight(100);
  ui.showModalDialog(htmlOutput, 'Open Link');
}



//user manually key in the purchase invoice into google sheet
//---------------send email to remind owner for the due date-------------------------------------------
function sendEmailAndAddCalendarEvent() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PurchaseInvoice'); // Change 'Sheet1' to your sheet name
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  
  var today = new Date();
  var ownerEmail = Session.getActiveUser().getEmail(); // Get the owner's email address
  
  for (var i = 1; i < data.length; i++) {
    var poID = data[i][0]; // Purchase Invoice ID
    var supplierID = data[i][1]; // Supplier ID
    var purchaseDate = new Date(data[i][2]); // Purchase Invoice Date
    var dueDate = new Date(data[i][3]); // Due Date
    var totalPurchase = data[i][4]; // Total Purchase
    var paymentStatus = data[i][5]; // Payment Status
    var paymentMethod = data[i][6]; // Payment Method
    
    // Check if due date is within 3 days
    var diffDays = Math.floor((dueDate - today) / (1000 * 60 * 60 * 24));
    
    Logger.log("Processing PO ID: " + poID);
    Logger.log("Due Date: " + dueDate);
    Logger.log("Days until due: " + diffDays);
    
    if (diffDays <= 3 && paymentStatus !== "Paid") {
      var subject = "Purchase Order Due Date Reminder";
      var body = "Reminder: Purchase Order " + poID + " is due on " + dueDate.toDateString() + ".\n" +
                 "Total Purchase: " + totalPurchase + "\n" +
                 "Payment Status: " + paymentStatus + "\n" +
                 "Please take necessary action.";
      
      // Send email
      MailApp.sendEmail(ownerEmail, subject, body);
      
      Logger.log("Email sent to: " + ownerEmail);
      
      // Add event to Google Calendar
      var calendar = CalendarApp.getDefaultCalendar();
      var eventTitle = "PO Due: " + poID;
      var eventOptions = {
        description: body,
        guests: ownerEmail,
        sendInvites: true
      };
      calendar.createEvent(eventTitle, dueDate, new Date(dueDate.getTime() + (1 * 60 * 60 * 1000)), eventOptions); // Event duration is 1 hour
      
      Logger.log("Event created for PO ID: " + poID);
    }
  }
}


//-------------------update product price from purchase invoice--------------------------------------------
function updateProductPrices() {
  var ss = SpreadsheetApp.openById('1RWL-8SIeGhJ3G7JDMTnltguj0Tvin5f4vBHDRo06qDw');
  var productsSheet = ss.getSheetByName("Product");
  var purchaseInvoiceItemsSheet = ss.getSheetByName("PurchaseInvoiceItem");

  var productsData = productsSheet.getDataRange().getValues();
  var purchaseInvoiceItemsData = purchaseInvoiceItemsSheet.getDataRange().getValues();

  var productPricesMap = {};
  var priceChanges = [];

  // Log initialization
  Logger.log('Starting price calculation');

  for (var i = 1; i < purchaseInvoiceItemsData.length; i++) {
    var productId = purchaseInvoiceItemsData[i][3]; // Product ID in the 4th column
    var purchaseQuantity = purchaseInvoiceItemsData[i][4]; // Purchase Quantity
    var purchasePriceRaw = purchaseInvoiceItemsData[i][5]; // Purchase Per Unit Price

    // Ensure purchaseQuantity is a number
    if (typeof purchaseQuantity === 'string') {
      purchaseQuantity = parseFloat(purchaseQuantity) || 0;
    }

    if (purchasePriceRaw) {
      var purchasePrice;
      if (typeof purchasePriceRaw === 'string') {
        purchasePrice = parseFloat(purchasePriceRaw.replace('RM', '').trim()) || 0;
      } else if (typeof purchasePriceRaw === 'number') {
        purchasePrice = purchasePriceRaw;
      } else {
        Logger.log('Unexpected data type for purchasePriceRaw: ' + typeof purchasePriceRaw);
        continue;
      }

      if (purchasePrice > 0) {
        if (productPricesMap[productId]) {
          productPricesMap[productId].totalQuantity += purchaseQuantity;
          productPricesMap[productId].totalPrice += (purchaseQuantity * purchasePrice);
        } else {
          productPricesMap[productId] = {
            totalQuantity: purchaseQuantity,
            totalPrice: (purchaseQuantity * purchasePrice)
          };
        }
      }
    } else {
      Logger.log('Missing or invalid purchasePriceRaw at row ' + (i + 1));
    }
  }

  for (var i = 1; i < productsData.length; i++) {
    var productId = productsData[i][0]; // Product ID in the 1st column
    if (productPricesMap[productId]) {
      var averagePrice = productPricesMap[productId].totalPrice / productPricesMap[productId].totalQuantity;
      var oldPrice = productsData[i][7]; // Price in the 8th column
      var newPrice = averagePrice.toFixed(2); // Use a numeric value instead of a string

      // Log the price change
      Logger.log('Product ID: ' + productId);
      Logger.log('Old Price: ' + oldPrice);
      Logger.log('New Price: ' + newPrice);

      productsSheet.getRange(i + 1, 8).setValue(parseFloat(newPrice)); // Update the price in the 8th column with a numeric value
      priceChanges.push({
        productId: productId,
        oldPrice: oldPrice,
        newPrice: newPrice
      });
    }
  }

  // Log completion
  Logger.log('Price update completed');
  showPriceChanges(priceChanges);
}

function showPriceChanges(priceChanges) {
  var template = HtmlService.createTemplateFromFile('PriceChanges');
  template.priceChanges = priceChanges;
  var htmlOutput = template.evaluate()
      .setWidth(600)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Price Changes');
}

