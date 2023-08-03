function compareEmailsAndShowAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("Deduper");
  var sheet2 = ss.getSheetByName("trevorExport");
  
  // Get the data from both sheets
  var data1 = sheet1.getRange("A:A").getValues();
  var data2 = sheet2.getRange("A:Q").getValues(); // Entire possible range
  
  // Create a map of email addresses and their corresponding information from the second sheet
  var emailInfoMap = {};
  var externalIDToEmailsMap = {}; // Map to store all emails with the same externalID
  for (var i = 1; i < data2.length; i++) {
    var email = data2[i][0].toString().toLowerCase();
    var externalID = data2[i][4]; // Column E contains externalCustomerID
    var status = data2[i][16]; // Column Q contains status
    var customerType = data2[i][5]; // Column F contains customerType
    
    if (email) { // Ensure emails are non-empty
      emailInfoMap[email] = {
        externalID: externalID,
        status: status,
        customerType: customerType
      };
      
      // Add emails to the externalIDToEmailsMap for grouping by externalID
      if (externalID && externalID !== "") {
        if (!externalIDToEmailsMap[externalID]) {
          externalIDToEmailsMap[externalID] = [email];
        } else {
          externalIDToEmailsMap[externalID].push(email);
        }
      }
    }
  }
  
  var foundEmails = [];
  var notFoundEmails = [];
  var revokeEmails = []; // Store emails with the same externalID as foundEmails but not present in foundEmails
  var duplicateEmails = {};
  var notEmails = [];
  var emailOccurrences = {}; // Track occurrences of each email in Sheet1
  
  // Loop through the first sheet to compare emails and check for duplicates
  for (var i = 0; i < data1.length; i++) {
    var email = data1[i][0].toString().trim().toLowerCase();
    
    // Validate email with regex pattern
    var emailRegex = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    var isValidEmail = emailRegex.test(email);
    
    if (isValidEmail) {
      // Check if the email is found in the list from the second sheet
      if (email in emailInfoMap) {
        // Add to the result array only if it's not a duplicate
        if (!duplicateEmails[email]) {
          foundEmails.push({ email: email, externalID: emailInfoMap[email].externalID, status: emailInfoMap[email].status, customerType: emailInfoMap[email].customerType });
          duplicateEmails[email] = true; // Mark as duplicate
        }
      } else {
        // Add to the result array only if it's not a duplicate
        if (!duplicateEmails[email]) {
          notFoundEmails.push(email);
          duplicateEmails[email] = true; // Mark as duplicate
        }
      }
      
      // Track occurrences of each email in Sheet1
      if (email in emailOccurrences) {
        emailOccurrences[email]++;
      } else {
        emailOccurrences[email] = 1;
      }
    } else {
      // Check for input that fails the email validation regex
      if (email.trim() !== "" && !notEmails.includes(email)) {
        notEmails.push(email);
      }
    }
  }
  
  // Remove empty entries from the notFoundEmails array
  notFoundEmails = notFoundEmails.filter(email => email.trim() !== "");
  
  // Find emails with the same externalID as foundEmails but not in foundEmails
  foundEmails.forEach(function (foundEmailInfo) {
    var externalID = foundEmailInfo.externalID;
    if (externalID && externalIDToEmailsMap[externalID]) {
      externalIDToEmailsMap[externalID].forEach(function (email) {
        // Check if the email has status "Active" before adding to revokeEmails
        if (!duplicateEmails[email] && emailInfoMap[email].status === "Active") {
          revokeEmails.push(email);
          duplicateEmails[email] = true; // Mark as duplicate
        }
      });
    }
  });
  
  // Create the alert message
  var message = "";
  
  if (foundEmails.length > 0) {
    message += "Existing Users (" + foundEmails.length + "):\n";
    
    for (var j = 0; j < foundEmails.length; j++) {
      var emailInfo = foundEmails[j];
      message += emailInfo.email + " (externalCustomerID: " + emailInfo.externalID + ", status: " + emailInfo.status + ", customerType: " + emailInfo.customerType + ")\n";
    }
  }
  
  if (notFoundEmails.length > 0) {
    message += "\nNew Users (" + notFoundEmails.length + "):\n";
    
    for (var k = 0; k < notFoundEmails.length; k++) {
      message += notFoundEmails[k] + "\n";
    }
  }

  if (notFoundEmails.length > 0) {
  message += "\nNew Users (CSV FORMAT) (" + notFoundEmails.length + "):\n";
  
  // Create a comma-separated string of notFoundEmails
  var notFoundEmailsCSV = notFoundEmails.join(",");
  
  // Append the CSV string to the message
  message += notFoundEmailsCSV + "\n";
}
  
  if (revokeEmails.length > 0) {
    message += "\nRevoke (" + revokeEmails.length + "):\n";
    
    for (var l = 0; l < revokeEmails.length; l++) {
      message += revokeEmails[l] + "\n";
    }
  }
  
  // Display duplicate email submissions
  var duplicateEmailsList = Object.keys(emailOccurrences).filter(email => emailOccurrences[email] > 1);
  if (duplicateEmailsList.length > 0) {
    message += "\nDuplicate Submissions (" + duplicateEmailsList.length + "):\n";
    for (var m = 0; m < duplicateEmailsList.length; m++) {
      message += duplicateEmailsList[m] + " (Occurrences: " + emailOccurrences[duplicateEmailsList[m]] + ")\n";
    }
  }
  
  // Display input that fails to be validated as an email
  if (notEmails.length > 0) {
    // Remove empty entries from the notEmails array
    notEmails = notEmails.filter(email => email.trim() !== "");
    
    message += "\nNot Valid Email (check for Typos) (" + notEmails.length + "):\n";
    
    for (var n = 0; n < notEmails.length; n++) {
      message += notEmails[n] + "\n";
    }
  }
  
  if (message === "") {
    message = "No emails were found.";
  }
  
  // Show the alert
  SpreadsheetApp.getUi().alert(message);
}
