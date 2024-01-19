function sendEmails() {
  var sheetName = "Your Sheet Name"; // Replace with your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Function to get HTML content from a Google Document
  function getHtmlContent(docId) {
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + docId + "&exportFormat=html";
    var param = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    };
    return UrlFetchApp.fetch(url, param).getContentText();
  }

  // Get HTML content for each document (Replace with your document IDs)
  var bodyAcceptedDocId = 'your-document-id-for-accepted'; 
  var bodyRejectedDocId = 'your-document-id-for-rejected';
  var bodyLocationDocId = 'your-document-id-for-location';

  var bodyAccepted = getHtmlContent(bodyAcceptedDocId);
  var bodyRejected = getHtmlContent(bodyRejectedDocId);
  var bodyLocation = getHtmlContent(bodyLocationDocId);

  // Function to send HTML email
  function sendHtmlEmail(to, subject, htmlBody, options) {
    var emailOptions = {
      htmlBody: htmlBody,
      name: options.name,
      replyTo: options.replyTo
      // attachments: options.attachments // Uncomment and use if needed
    };
    GmailApp.sendEmail(to, subject, "", emailOptions);
  }

  // Email sending logic
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var name = row[0]; // Adjust the index based on your sheet structure
    var email = row[1]; // Adjust the index based on your sheet structure
    var applicationStatus = row[2]; // Adjust the index based on your sheet structure
    var applicationEmailStatus = row[3]; // Adjust the index based on your sheet structure

    // Define html placeholders for each email type to replace specific text (e.g. [Student's name])
    var placeholderAccepted = 'placeholder-for-accepted'; // Adjust this
    var placeholderRejected = 'placeholder-for-rejected'; // Adjust this
    var placeholderLocation = 'placeholder-for-location'; // Adjust this

    if (applicationStatus == "Accepted" && applicationEmailStatus == "") {
      var personalizedBodyAccepted = bodyAccepted.replace(placeholderAccepted, name);
      sendHtmlEmail(email, "Your Acceptance Subject", personalizedBodyAccepted, {name: "Your Name", replyTo: "your-email@example.com"});
      sheet.getRange(i + 1, 4).setValue("SentAccepted"); // Adjust the index based on your sheet structure
    } else if (applicationStatus == "Rejected" && applicationEmailStatus == "") {
      var personalizedBodyRejected = bodyRejected.replace(placeholderRejected, name);
      sendHtmlEmail(email, "Your Rejection Subject", personalizedBodyRejected, {name: "Your Name", replyTo: "your-email@example.com"});
      sheet.getRange(i + 1, 4).setValue("SentRejected"); // Adjust the index based on your sheet structure
    } else if (applicationStatus == "Location" && applicationEmailStatus == "") {
      var personalizedBodyLocation = bodyLocation.replace(placeholderLocation, name);
      sendHtmlEmail(email, "Your Location Subject", personalizedBodyLocation, {name: "Your Name", replyTo: "your-email@example.com"});
      sheet.getRange(i + 1, 4).setValue("SentReferral"); // Adjust the index based on your sheet structure
    }
  }
}
