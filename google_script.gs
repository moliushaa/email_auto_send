function sendInvitationEmails() {
  var snap = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1yMjZKd65ZK62Eqr9f8kYyGvNycTpL701WGNYRtJHo2w/edit#gid=709961477");
  var sheet = snap.getSheetByName("达人搜索-美国-娱乐");
  var startRow = 1; // Change this to the row number where data starts
  var numRows = sheet.getLastRow(); // Number of rows with data (excluding header row)
  var dataRange = sheet.getRange(startRow, 1, numRows, 13); // Columns A to M have 13 columns
  var data = dataRange.getRichTextValues();

  for (var i = 0; i < numRows; i++) {
    var row = data[i];
    var name = row[1].getText();
    var cell = dataRange.getCell(i+1, 12+1);
    var richTextValue = cell.getRichTextValue();
    if (richTextValue) {
      var email = richTextValue.getLinkUrl();
      if (email && email.startsWith('mailto:')) {
        email = email.substring(7);
      }
    }

    var msg_tmplt_url = snap.getSheetByName("MessagesTemplate").getRange("!A1:B2").getValues();
    
    if (email !== "") {
      var subject = "✨Invitation for L.A. Starlight Gala by GloW x N2M✨";
      var message_url = msg_tmplt_url[0][1];
      var message_id = DocumentApp.openByUrl(message_url).getId();
      var message = parseMessage(getHtmlByDocId(message_id), name);
  
      try {
        // Send the email
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: message
        });
        Utilities.sleep(1000);
      } catch (error) {
        // Log the error and continue to the next row
        Logger.log("row number: " + i + " Error sending email to " + email + ": " + error);
        continue;
      }
    } else {
      Logger.log("Skipping row " + (startRow + i) + " because email address is empty or invalid.");
    }
  }
}
function getHtmlByDocId(id) {
  // fetch html content from google doc
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id+"&exportFormat=html";
  var param = 
  {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
  };
  var html = UrlFetchApp.fetch(url,param).getContentText();
  
  // beautify html a bit
  html = html.replaceAll(/<p class=\"[c0-9\ ]+\"><span class=\"[c0-9\ ]+\"><\/span>/g, "<br>");
  var start = html.search("<body");
  var end = html.search("</body>");
  html = html.substr(start, (end-start+7));
  
  return html;
}
function parseMessage(msg, name) {
  msg_keywords = {
    'name': name,
  };
  match = [...msg.matchAll(/\${([a-zA-Z_]+)}/g)];
  for (m of match) {
    if (m[1] in msg_keywords) {
      replacement = msg_keywords[m[1]];
      msg = msg.replace(m[0], replacement);
    }
  }
  return msg;
}
