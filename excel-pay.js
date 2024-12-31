// This function reads data from the specified columns
function excelPay() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues(); // Get all data in the sheet
  
    var messages = []; // Array to store messages
  
    // Loop through each row of data
    for (var i = 1; i < data.length; i++) { // Start from 1 to skip header
      var date = data[i][0]; // Assuming date is in column A (index 0)
      var recipient = data[i][1]; // Assuming recipient is in column B (index 1)
      var amount = data[i][2]; // Assuming amount is in column C (index 2)
      var payer = data[i][3]; // Assuming payer is in column D (index 3)
      var purpose = data[i][4]; // Assuming purpose is in column E (index 4)
      // Check if recipient, amount, payer, and purpose are valid
      if (recipient && !isNaN(amount) && payer && purpose) {
        // Format the message to resemble a payment notification
        var message = `💰 <strong>匯款成功</strong><br>` +
                      `日期 : <strong>${date}</strong><br>` +
                      `匯款給 : <strong>${recipient}</strong><br>` +
                      `金額 : <strong>${amount}</strong><br>` +
                      `匯款人 : <strong>${payer}</strong><br>` +
                      `用途 : <strong>${purpose}</strong><br>` +
                      `<br>----- 分隔線！ -----<br>`
        messages.push(message);
        Logger.log(`${message}`);
        Logger.log(`${messages}`);
      } else {
          Logger.log(`Row ${i} is invalid.`);
      }
    }
    messages.push(`感謝您的使用！<br>`)
    messages.push(`如有任何問題，請聯繫 <a href="https://line.me/R/ti/p/@504uaynd" target="_blank">Line 客服</a>。`);
  
    // Create HTML output
    var htmlOutput = HtmlService.createHtmlOutput('<h2>Excel Pay 付款通知</h2>' + messages.join("<br><br>"))
        .setWidth(800)
        .setHeight(600);
    
    // Add logo image using Base64 or URL
    htmlOutput.append('<img src="' + base64Image + '" alt="Logo" style="width:20%; height:20%;">');
  
    // Show the HTML dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Excel Pay');
  }
  
  // This function creates a custom menu in Google Sheets
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Excel Pay')
        .addItem('Excel Pay', 'excelPay')
        .addToUi();
  }