function myFunction() {
  // Auto email script for HRL permament key changes
  // Mitch Bath 2024

  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");

  var startRow = 2;  // First row of data to process - 2 exempts my header row
  var numRows = responseSheet.getLastRow();   // Number of rows to process
  var numColumns = responseSheet.getLastColumn();
 
  var responseDataRange = responseSheet.getRange(startRow, 1, numRows-1, numColumns);
  var responseData = responseDataRange.getValues();
  var complete = "Done";

  // villages, LKO and REC
  var buildinglist = [//REDACTED//]; //Do not edit this
  
  // To update: replace email with the current building's REC email
  var reclist = [// REDACTED //
  ];
  // To update: replace email with the current building's LKO email
  var lkolist = [
  // REDACTED//
  ];

  // run per row
  for (var i = 0; i < responseData.length; ++i) {

    // collect data
    var row = responseData[i];

    var timestamp = row[0];
    var emailaddress = row[1];
    var staffname = row[2];
    var studentname = row[3];
    var studentID = row[4];
    var studentroomnumber = row[5];
    var studentbuilding = row[6];

    var currentrec = "";
    var currentlko = "";

    for (var k = 0; k < buildinglist.length; k++) {

      if (studentbuilding.localeCompare(buildinglist[k]) == 0) {
        currentrec = reclist[k];
        currentlko = lkolist[k];
        break;
      }

    }

    var reason = row[7];
    var brokenkeycode = "N/A";
    if (row[8] != null && row[8] != "") {
      brokenkeycode = row[8];
    }
    var newkeycode = row[9];
    var completestatus = row[10];

    // new submissions only
    if (completestatus != complete) {

      // write the email
      var message = "<b>Name: </b>" + studentname.toString() +"<br><b>49er ID: </b>" + studentID.toString() + "<br><b>Room Assignment: </b>" + studentbuilding.toString() + " " + studentroomnumber.toString() +"<br><b>Issue Description: </b>" + reason.toString() + "<br><b>Broken Key Code: </b>" + brokenkeycode.toString() + "<br><b>New Perm. Key Code: </b>" + newkeycode.toString() + "<br><br>Thank you!<br><br>" + staffname.toString() + "<br>" + emailaddress.toString();

      var subject = studentbuilding.toString() + " " + studentroomnumber.toString() + ": Permanent Key Change - New Loan Key Needed";
      var sendto = "// REDACTED// ;
      var tocc = currentrec.toString() + ", " + currentlko.toString() + ", " + emailaddress.toString();

      if (staffname.localeCompare("TEST")==0) {
        sendto = emailaddress;
        tocc = "";
        subject = "[TEST]" + subject;
      }

      // email compile and send
      MailApp.sendEmail({
        to: sendto,
        subject: subject,
        htmlBody: message,
        cc: tocc
      }
        );

      // mark the final column as Done
      responseSheet.getRange(startRow + i, numColumns).setValue(complete);

    }

  }

}
