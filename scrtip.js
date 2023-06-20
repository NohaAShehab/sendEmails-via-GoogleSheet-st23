
let EMAIL_SENT = 'Sent';
var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
Logger.log("Remaining email quota: " + emailQuotaRemaining);

function sendNonDuplicateEmails() {
  try{
    // Get the active sheet in spreadsheet
    const sheet = SpreadsheetApp.getActiveSheet();
    let startRow = 5; // First row of data to process
    let numRows = 35; // Number of rows to process
    // Fetch the range of cells 
    const dataRange = sheet.getRange(startRow, 2, numRows, 20);
    // Fetch values for each row in the Range.
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; ++i) {
      const row = data[i];
      const emailAddress = row[3]; // Fourth Column in our selection
      console.log(row)
      var email_body = 
      `<pre>Greetings,${row[0]}
Congratulations you have been acceppted in ITI Mansoura Summer Training, 
      <h3 style='text-align:center;'> Track: ${row[7]} </h3>
and use this link to join the track group on WhatsApp: ${row[6]}</pre>
      <br/>
      <br/>
Best of luck ^^ <br/>
ITI Mansoura <br/>`

    console.log(email_body);
      const emailSent = row[12]; //
      if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
        let subject = 'ITI-Mansoura Summer Training';
        // Send emails to emailAddresses which are presents in Fourth column
        MailApp.sendEmail( {to:emailAddress, name:'ITI Mansoura', subject:subject, body:email_body,htmlBody:email_body});
        sheet.getRange(startRow + i,15).setValue(EMAIL_SENT);
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      }
    }
  }
  catch(err){
    Logger.log(err)
  }
}
