function extractEmails() {
  var threads = GmailApp.search('is:inbox'); // Puoi modificare la ricerca, ad esempio 'from:client@dominio.com'
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var emails = [];
  
  // Estrai gli indirizzi da tutte le email trovate
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var email = messages[j].getFrom();
      emails.push([email]);
    }
  }
  
  // Scrivi gli indirizzi email nel foglio
  sheet.getRange(1, 1, emails.length, 1).setValues(emails);
}
