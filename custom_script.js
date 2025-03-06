function createAndSendInvoices() {
  const sheetId = "your_datasheet_id";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Billing');
  const docTemplateId = "billing_template_doc_id";
  const folderId = "saving_path_folder_id";
  const folder = DriveApp.getFolderById(folderId);

  const data = sheet.getDataRange().getValues();
  const lastRow = data.length;
  Logger.log(`Total number of rows with data: ${lastRow}`);

  for (let i = 1; i < lastRow; i++) {
    const row = data[i];
    Logger.log(`Processing row ${i + 1}: ${row}`);

    const invoiceNo = row[0];  
    const contactName = row[1]; 
    const invoiceDate = row[2]; 
    const clientCompanyName = row[3]; 
    const dueDate = row[4];  
    const address = row[5];  
    const phone = row[6];  
    const email = row[7];  

    if (!email || email.trim() === '') {
      Logger.log(`Skipping row ${i + 1}: No email provided for ${contactName}`);
      continue;
    }

    const description = row[8];  
    const hsnCode = row[9];  
    const lastReading = row[10];  
    const currentReading = row[11];  
    const qty = row[12];  
    const rate = row[13];  
    const freight = row[14];  

    const totalBeforeTax = qty * rate; 
    const cgst = 0.09 * totalBeforeTax; 
    const sgst = 0.09 * totalBeforeTax; 
    const totalTax = cgst + sgst; 
    const grandTotal = totalBeforeTax + totalTax + parseFloat(freight);

    const roundedGrandTotal = Math.round(grandTotal);
    const roundOffValue = (roundedGrandTotal - grandTotal).toFixed(2);
    const amountInWords = convertNumberToWords(roundedGrandTotal);

    const fileName = `Invoice_${clientCompanyName}_${invoiceNo}.pdf`;
    if (checkIfFileExistsInFolder(fileName, folder)) {
      Logger.log(`Invoice already exists: ${fileName}`);
      continue;
    }

    // copy of our template
    const docTemplate = DriveApp.getFileById(docTemplateId);
    const copy = docTemplate.makeCopy(`Invoice_${clientCompanyName}_${invoiceNo}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    
    // placeholders stuff
    body.replaceText('{{InvoiceNo}}', invoiceNo);
    body.replaceText('{{ContactName}}', contactName);
    body.replaceText('{{InvoiceDate}}', invoiceDate);
    body.replaceText('{{ClientCompanyName}}', clientCompanyName);
    body.replaceText('{{DueDate}}', dueDate);
    body.replaceText('{{Address}}', address);
    body.replaceText('{{Phone}}', phone);
    body.replaceText('{{Email}}', email);
    body.replaceText('{{Description}}', description);
    body.replaceText('{{CurrentReading}}', currentReading);
    body.replaceText('{{LastReading}}', lastReading);
    body.replaceText('{{HSNCode}}', hsnCode);
    body.replaceText('{{QTY}}', qty.toString());
    body.replaceText('{{Rate}}', rate.toString());
    body.replaceText('{{Total}}', totalBeforeTax.toFixed(2)); 
    body.replaceText('{{Freight}}', freight);
    body.replaceText('{{CGST}}', cgst.toFixed(2)); 
    body.replaceText('{{SGST}}', sgst.toFixed(2)); 
    body.replaceText('{{TotalBeforeTax}}', totalBeforeTax.toFixed(2)); 
    body.replaceText('{{TotalTax}}', totalTax.toFixed(2)); 
    body.replaceText('{{TotalTaxRate}}', '18'); 
    body.replaceText('{{AmountInWords}}', amountInWords);
    body.replaceText('{{GrandTotal}}', roundedGrandTotal.toFixed(2)); 
    body.replaceText('{{RoundOff}}', roundOffValue.toString()); 
    
    doc.saveAndClose();
    
    const pdfFile = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
    const pdf = folder.createFile(pdfFile);
    
    MailApp.sendEmail({
      to: email,
      subject: `Experiment_Bill_Format Invoice ${invoiceNo}`,
      body: `Dear ${contactName},\n\nPlease find attached your invoice.\n\nRegards,\nHSK Enterprises`,
      attachments: [pdf]
    });
    
    Logger.log(`Email sent to ${email} for Invoice No: ${invoiceNo}`);
    copy.setTrashed(true);
  }
}

//existing file or not
function checkIfFileExistsInFolder(fileName, folder) {
  const files = folder.getFilesByName(fileName);
  
  Logger.log(`Checking for file: ${fileName} in folder: ${folder.getName()}`);
  
  if (files.hasNext()) {
    Logger.log(`File found: ${fileName}`);
    return true; // exists
  } else {
    Logger.log(`File not found: ${fileName}`);
    return false; // does not exist
  }
}




function convertNumberToWords(number) {
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  
  if (number === 0) return 'Zero Rupees';

  let word = '';
  const crore = Math.floor(number / 10000000); // Extract crore
  number %= 10000000;
  const lakh = Math.floor(number / 100000); // Extract lakh
  number %= 100000;
  const thousand = Math.floor(number / 1000); // Extract thousand
  number %= 1000;
  const hundred = Math.floor(number / 100); // Extract hundred
  number %= 100;
  const tensAndOnes = number; // Remaining tens and ones

  if (crore > 0) {
    word += convertLessThanThousand(crore) + ' Crore ';
  }
  if (lakh > 0) {
    word += convertLessThanThousand(lakh) + ' Lakh ';
  }
  if (thousand > 0) {
    word += convertLessThanThousand(thousand) + ' Thousand ';
  }
  if (hundred > 0) {
    word += ones[hundred] + ' Hundred ';
  }
  if (tensAndOnes > 0) {
    if (tensAndOnes < 20) {
      word += ones[tensAndOnes] + ' ';
    } else {
      word += tens[Math.floor(tensAndOnes / 10)] + ' ' + ones[tensAndOnes % 10] + ' ';
    }
  }

  return word.trim() + ' Rupees';
}

function convertLessThanThousand(number) {
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

  let word = '';
  const hundred = Math.floor(number / 100);
  number %= 100;

  if (hundred > 0) {
    word += ones[hundred] + ' Hundred ';
  }
  if (number > 0) {
    if (number < 20) {
      word += ones[number] + ' ';
    } else {
      word += tens[Math.floor(number / 10)] + ' ' + ones[number % 10] + ' ';
    }
  }

  return word.trim();
}
