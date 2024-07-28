// Function to handle HTTP GET requests
function doGet() {
  // Returns the HTML output from 'index' file
  return HtmlService.createHtmlOutputFromFile('index');
}

// Function to display the Sales Invoice HTML UI
function openInvoiceGenerator() {
  // Creates HTML output from 'SalesInvoice' file and sets its dimensions
  const html = HtmlService.createHtmlOutputFromFile('SalesInvoice')
      .setWidth(600)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  // Displays the HTML as a modal dialog in the Google Sheets UI
  SpreadsheetApp.getUi().showModalDialog(html, 'Sales Invoice');
}

// Function to display the Sales Report HTML UI
function showSalesReportUI() {
  // Creates HTML output from 'SalesReportUI' file and sets its dimensions
  const html = HtmlService.createHtmlOutputFromFile('SalesReportUI')
      .setWidth(600)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  // Displays the HTML as a modal dialog in the Google Sheets UI
  SpreadsheetApp.getUi().showModalDialog(html, 'Sales Report');
}

// Function to display the Index HTML UI
function showIndexUI() {
  // Creates HTML output from 'index' file and sets its dimensions
  const html = HtmlService.createHtmlOutputFromFile('SalesIndex')
      .setWidth(400)
      .setHeight(400);
  // Displays the HTML as a modal dialog in the Google Sheets UI
  SpreadsheetApp.getUi().showModalDialog(html, 'Main Menu');
}

// Custom menu on opening the spreadsheet
function onOpen() {
  // Accesses the Google Sheets UI
  const ui = SpreadsheetApp.getUi();
  // Creates a custom menu with an item that triggers 'showIndexUI'
  ui.createMenu('Sales')
    .addItem('Sales Main Menu', 'showIndexUI')
    .addToUi();
}

// Function to generate invoices based on selected IDs
function generateInvoices(selectedIDs) {
  // Accesses the relevant sheets from the active spreadsheet
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoice');
  const itemsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoiceItem');
  const productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product');

  // Retrieves all data from the sheets
  const salesData = salesSheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();
  const productsData = productsSheet.getDataRange().getValues();

  // Iterates over each selected Sales ID to generate an invoice
  selectedIDs.forEach(salesID => {
    // Finds the corresponding sales row by Sales ID
    const salesRow = salesData.find(row => row[0] === salesID);
    if (!salesRow) return; // Skip if Sales ID not found in SalesInvoice

    // Filters items for the given Sales ID
    const items = itemsData.filter(item => item[1] === salesID);
    const services = items.map(item => {
      // Finds the product details for each item
      const product = productsData.find(prod => prod[0] === item[2]);
      // Converts the price and amount to numbers
      const fee = product && typeof product[3] === 'string' 
      ? parseFloat(product[3].replace('RM', '').replace(',', ''))
      : parseFloat(product[3] || 0); // Fallback to 0 if conversion fails

      const total = item && typeof item[8] === 'string'
      ? parseFloat(item[8].replace('RM', '').replace(',', ''))
      : parseFloat(item[8] || 0); // Fallback to 0 if conversion fails

      return {
        listed: product ? product[1] : 'Unknown Product',
        fee: fee,
        quantity: item[3],
        total: total
      };
    });

    // Calculates subtotal and total due including SST
    let subtotal = services.reduce((acc, curr) => acc + curr.total, 0);
    let totalDue = subtotal * 1.06; // Including SST 6%

    // Creates a new Google Document for the invoice
    const doc = DocumentApp.create(`Invoice-${salesID}`);
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
    body.appendParagraph("");

    // Invoice Details
    body.appendParagraph(`Invoice #: ${salesID}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Invoice Date: ${new Date().toLocaleDateString()}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Sales Date: ${salesRow[1]}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Payment Method: ${salesRow[3]}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph("");

    // Services Table
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell('Item').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Quantity').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Unit Price (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    headerRow.appendTableCell('Amount (RM)').setBackgroundColor('#f3f3f3').setBold(true).setFontSize(10);
    services.forEach(service => {
      const row = table.appendTableRow();
      row.appendTableCell(service.listed).setFontSize(10);
      row.appendTableCell(`${service.quantity}`).setFontSize(10);
      row.appendTableCell(`RM ${service.fee.toFixed(2)}`).setFontSize(10);
      row.appendTableCell(`RM ${service.total.toFixed(2)}`).setFontSize(10);
    });

    // Financial Summary
    body.appendParagraph(`Subtotal: RM ${subtotal.toFixed(2)}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`SST (6%): RM ${(subtotal * 0.06).toFixed(2)}`).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph(`Total Due: RM ${totalDue.toFixed(2)}`).setFontSize(10).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph("");

    // PDF Generation and Sharing
    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');
    // Creates or retrieves 'Invoices' folder in Google Drive
    const folder = DriveApp.getFoldersByName("Invoices").hasNext() ? DriveApp.getFoldersByName("Invoices").next() : DriveApp.createFolder("Invoices");
    let version = 1;
    let pdfFileName = `Invoice-${salesID}_V${String(version).padStart(2, '0')}.pdf`;
    // Checks for existing files with the same name and increments the version number if necessary
    while (folder.getFilesByName(pdfFileName).hasNext()) {
      version++;
      pdfFileName = `Invoice-${salesID}_V${String(version).padStart(2, '0')}.pdf`;
    }
    const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const pdfUrl = pdfFile.getUrl();

    Logger.log(`Invoice PDF generated successfully. Version: ${version}. Link: ${pdfUrl}`);
  });

  return `Invoices generated successfully for Sales IDs: ${selectedIDs.join(', ')}.`;
}

// Function to retrieve sales data from 'SalesInvoice' sheet
function getSalesData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoice'); 
  const data = sheet.getDataRange().getValues();
  
  // Skip the header row and map data to objects
  const salesData = data.slice(1).map(row => ({
    salesId: row[0],
    salesDate: row[1],
    totalSales: row[2],
    paymentMethod: row[3]
  }));
  Logger.log(salesData);
  return JSON.stringify(salesData);
}

// Function triggered on cell edit
function onEdit(e) {
  try {
    // Ensure the edited sheet is 'SalesInvoiceItem'
    const sheet = e.source.getSheetByName('SalesInvoiceItem');
    if (!sheet || e.range.getSheet().getName() !== 'SalesInvoiceItem') {
      return;
    }

    // Check if the edit is in the relevant columns (Product ID or Sales Quantity)
    const editedRange = e.range;
    if (editedRange.getColumn() === 3 || editedRange.getColumn() === 4) { // Product ID (Column C) or Sales Quantity (Column D)
      handleRowInsertion(sheet);
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error.message);
  }
}

// Function to handle new row insertion in 'SalesInvoiceItem' sheet
function handleRowInsertion(sheet) {
  const lastRow = sheet.getLastRow();
  const newRowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const salesID = newRowData[1]; // Sales Invoice ID (Column B)
  const productId = newRowData[2]; // Product ID (Column C)
  const quantitySold = parseFloat(newRowData[3]); // Sales Quantity (Column D)
  const pricePerUnit = newRowData[4]; // Price Per Unit (Column E)

  if (isNaN(quantitySold)) {
    Logger.log('Invalid quantity sold: ' + newRowData[3]);
    return;
  }

  const productSheet = sheet.getParent().getSheetByName('Product');
  const salesSheet = sheet.getParent().getSheetByName('SalesInvoice');

  if (productSheet) {
    updateProductStock(productSheet, productId, quantitySold);
  } else {
    Logger.log('Product sheet not found.');
  }

  if (salesSheet) {
    handleSalesInvoice(salesSheet, salesID);
  } else {
    Logger.log('SalesInvoice sheet not found.');
  }
}

// Function to update the stock of a product in 'Product' sheet
function updateProductStock(productSheet, productId, quantitySold) {
  const data = productSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // Skip header row
    if (data[i][0] === productId) { // Product ID (Column A)
      let currentStock = parseFloat(data[i][4]); // Current Stock (Column E)
      let newStock = currentStock - quantitySold;

      productSheet.getRange(i + 1, 5).setValue(newStock); // Update Current Stock (Column E in 1-based index)
      Logger.log('Updated stock for Product ID: ' + productId);
      break;
    }
  }
}

// Function to handle updating or inserting a sales invoice
function handleSalesInvoice(salesSheet, salesID) {
  const salesData = salesSheet.getDataRange().getValues();
  const itemsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoiceItem');
  const itemsData = itemsSheet.getDataRange().getValues();

  // Check if the Sales Invoice ID exists
  const salesRowIndex = salesData.findIndex(row => row[0] === salesID) + 1; // 1-based index

  if (salesRowIndex === 0) {
    // Insert new Sales Invoice ID
    const newRow = [salesID, new Date().toLocaleDateString(), 'RM 0.00', ''];
    salesSheet.appendRow(newRow);
    Logger.log('Added new Sales Invoice ID: ' + salesID);
  }

  // Update total sales for the invoice
  updateSalesInvoice(salesSheet, salesID);
}

// Function to update total sales in 'SalesInvoice' sheet
function updateSalesInvoice(salesSheet, salesID) {
  const salesData = salesSheet.getDataRange().getValues();
  const itemsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoiceItem');
  const itemsData = itemsSheet.getDataRange().getValues();

  // Find the sales invoice row
  const salesRowIndex = salesData.findIndex(row => row[0] === salesID) + 1; // 1-based index
  if (salesRowIndex === 0) {
    Logger.log('Sales Invoice ID not found: ' + salesID);
    return;
  }

  // Calculate total sales
  const items = itemsData.filter(item => item[1] === salesID);

  let totalSales = 0;
  items.forEach(item => {
    const salesAmount = item[8]; // Sales Amount (Column I)

    // Log the type and value of salesAmount
    Logger.log('Sales Amount Value: ' + salesAmount + ', Type: ' + typeof salesAmount);

    if (typeof salesAmount === 'number') {
      totalSales += salesAmount;
    } else if (typeof salesAmount === 'string') {
      // Attempt to parse string as number
      const parsedAmount = parseFloat(salesAmount.replace('RM', '').replace(',', ''));
      if (!isNaN(parsedAmount)) {
        totalSales += parsedAmount;
      } else {
        Logger.log('Failed to parse salesAmount: ' + salesAmount);
      }
    } else {
      Logger.log('Unexpected data type for salesAmount: ' + salesAmount);
    }
  });

  // Update total sales in SalesInvoice sheet
  salesSheet.getRange(salesRowIndex, 3).setValue(`RM ${totalSales.toFixed(2)}`); // Column C: Total Sales
  Logger.log('Updated total sales for Sales Invoice ID: ' + salesID);
}

// Test function for row insertion handling
function testHandleRowInsertion() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoiceItem');
  handleRowInsertion(sheet);
}


//-----------------------------------------SalesReports----------------------------------------------

// Function to generate a Sales Summary Report
function generateSalesSummaryReport() {
  // Retrieve data from SalesInvoiceItem and SalesInvoice sheets
  const itemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoiceItem');
  const invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoice');
  
  const itemData = itemSheet.getDataRange().getValues();
  const invoiceData = invoiceSheet.getDataRange().getValues();

  const salesData = [];

  // Create a mapping of Sales Invoice IDs to Sales Dates
  const invoiceDateMap = {};
  invoiceData.slice(1).forEach(row => {
    const salesId = row[0];
    const salesDate = new Date(row[1]);
    if (salesDate instanceof Date && !isNaN(salesDate.getTime())) {
      invoiceDateMap[salesId] = salesDate;
    }
  });

  // Process each row of SalesInvoiceItem
  itemData.slice(1).forEach(row => {
    try {
      const salesId = row[1];
      let salesAmount = row[8];
      const salesDate = invoiceDateMap[salesId];

      // Skip if no date found for the Sales Invoice ID
      if (!salesDate) {
        Logger.log(`No date found for Sales Invoice ID: ${salesId}`);
        return;
      }

      // Convert salesAmount to a number
      if (salesAmount && typeof salesAmount === 'string') {
        salesAmount = salesAmount.replace('RM', '').replace(',', '').trim();
      } else if (salesAmount && typeof salesAmount === 'number') {
        salesAmount = salesAmount.toString();
      } else {
        salesAmount = '0';
      }

      const formattedAmount = parseFloat(salesAmount);

      // Skip if salesAmount is not a valid number
      if (isNaN(formattedAmount)) {
        Logger.log(`Invalid amount format: ${salesAmount}`);
        return;
      }

      salesData.push({
        salesDate: salesDate,
        salesAmount: formattedAmount
      });

    } catch (error) {
      Logger.log(`Error processing row: ${error.message}`);
    }
  });

  // Create and save the Sales Summary Report PDF
  createPDF(salesData);
}

// Function to create a PDF with the Sales Summary Report
function createPDF(salesData) {
  const doc = DocumentApp.create('Sales Summary Report');
  const body = doc.getBody();
  
  body.appendParagraph('Sales Summary Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  // Generate summaries for daily, monthly, and yearly sales
  const dailySummary = aggregateData(salesData, 'daily');
  const monthlySummary = aggregateData(salesData, 'monthly');
  const yearlySummary = aggregateData(salesData, 'yearly');

  // Append summary tables to the document
  appendSummaryTable(body, 'Daily Summary', dailySummary);
  appendSummaryTable(body, 'Monthly Summary', monthlySummary);
  appendSummaryTable(body, 'Yearly Summary', yearlySummary);

  doc.saveAndClose();
  
  // Convert the document to PDF and save it to Google Drive
  const docId = doc.getId();
  const pdf = DriveApp.getFileById(docId).getAs('application/pdf');
  
  const folder = DriveApp.getFolderById('1buPIxFFVPCJ4pJtxzYZpL3OmhZBT-frE');
  folder.createFile(pdf);
  
  Logger.log('PDF created and saved to Google Drive.');
}

// Function to aggregate sales data by period (daily, monthly, yearly)
function aggregateData(salesData, period) {
  const summary = {};

  salesData.forEach(item => {
    let key;
    switch (period) {
      case 'daily':
        key = item.salesDate.toISOString().split('T')[0]; // YYYY-MM-DD
        break;
      case 'monthly':
        key = item.salesDate.getFullYear() + '-' + (item.salesDate.getMonth() + 1); // YYYY-MM
        break;
      case 'yearly':
        key = item.salesDate.getFullYear().toString(); // YYYY
        break;
    }

    if (!summary[key]) {
      summary[key] = 0;
    }

    summary[key] += item.salesAmount;
  });

  return summary;
}

// Function to append a summary table to the document
function appendSummaryTable(body, title, summary) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  const table = [['Period', 'Total Sales']];
  
  for (const key in summary) {
    table.push([
      key,
      `RM ${summary[key].toFixed(2)}`
    ]);
  }
  
  body.appendTable(table);
}

//----------------------------------------SalesReportByProduct----------------------------

// Function to generate a Sales by Product Report and save it as a PDF
function generateSalesByProductReport() {
  const itemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SalesInvoiceItem');
  const productSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product');
  
  const itemData = itemSheet.getDataRange().getValues();
  const productData = productSheet.getDataRange().getValues();
  
  const report = {};
  
  // Process each row of SalesInvoiceItem
  itemData.slice(1).forEach(row => {
    const productId = row[2];
    let salesAmount = row[8];
    const quantitySold = parseFloat(row[3]);
    
    // Ensure salesAmount is a number
    if (typeof salesAmount === 'string') {
      salesAmount = parseFloat(salesAmount.replace('RM', '').replace(',', ''));
    } else if (typeof salesAmount === 'number') {
      salesAmount = parseFloat(salesAmount);
    } else {
      Logger.log('Invalid sales amount data: ' + salesAmount);
      return; // Skip this row
    }
    
    if (!report[productId]) {
      report[productId] = { totalSales: 0, totalQuantity: 0, productName: '' };
    }
    
    report[productId].totalSales += salesAmount;
    report[productId].totalQuantity += quantitySold;
    
    // Retrieve product name from Product sheet
    const productRow = productData.find(p => p[0] === productId);
    if (productRow) {
      report[productId].productName = productRow[1];
    }
  });
  
  // Convert the report object to an array and sort by totalQuantity in descending order
  const reportArray = Object.keys(report).map(productId => ({
    productId: productId,
    productName: report[productId].productName,
    totalSales: report[productId].totalSales,
    totalQuantity: report[productId].totalQuantity
  }));
  
  reportArray.sort((a, b) => b.totalQuantity - a.totalQuantity);
  
  // Select the top 5 products
  const topProducts = reportArray.slice(0, 5);

  const reportDoc = DocumentApp.create('Top 5 Sales by Product Report');
  const docBody = reportDoc.getBody();
  
  docBody.appendParagraph('Top 5 Sales by Product Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  // Create and style the table
  const table = [['Product ID', 'Product Name', 'Total Sales', 'Total Quantity Sold']];
  
  topProducts.forEach(item => {
    table.push([
      item.productId,
      item.productName,
      `RM ${item.totalSales.toFixed(2)}`,
      item.totalQuantity
    ]);
  });

  const pdfTable = docBody.appendTable(table);
  pdfTable.setBorderWidth(1);
  pdfTable.setBorderColor('#000000');

  // Style header row
  const headerRow = pdfTable.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    const cell = headerRow.getCell(i);
    cell.getChild(0).asText().setBackgroundColor('#4F81BD'); // Dark blue background
    cell.getChild(0).asText().setForegroundColor('#FFFFFF'); // White text
    cell.getChild(0).asText().setBold(true);
  }

  // Style other rows
  for (let i = 1; i < pdfTable.getNumRows(); i++) {
    const row = pdfTable.getRow(i);
    for (let j = 0; j < row.getNumCells(); j++) {
      const cell = row.getCell(j);
      cell.getChild(0).asText().setBackgroundColor(i % 2 === 0 ? '#F2F2F2' : '#FFFFFF'); // Alternate row colors
    }
  }
  
  reportDoc.saveAndClose();
  
  const docId = reportDoc.getId();
  const pdf = DriveApp.getFileById(docId).getAs('application/pdf');
  
  const folder = DriveApp.getFolderById('1buPIxFFVPCJ4pJtxzYZpL3OmhZBT-frE'); // Replace with your folder ID
  folder.createFile(pdf);
  
  Logger.log('PDF created and saved to Google Drive.');
  
  return 'Top 5 Sales by Product Report generated and saved as PDF successfully!';
}