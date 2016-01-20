//edit password was up here

function getCurRentData() {
  var billsSpread = SpreadsheetApp.openById("1-3UdWZzfnIZYVAtiM7khCvUoW-wPf-ewxGL6_zBlAMk");
  SpreadsheetApp.setActiveSpreadsheet(billsSpread);
  var rentSheet = billsSpread.getSheetByName("Rent");
  var data = rentSheet.getDataRange();
  
  return data;
}

function getCurCharterData() {
  var billsSpread = SpreadsheetApp.openById("1-3UdWZzfnIZYVAtiM7khCvUoW-wPf-ewxGL6_zBlAMk");
  SpreadsheetApp.setActiveSpreadsheet(billsSpread);
  var rentSheet = billsSpread.getSheetByName("Rent");
  var data = rentSheet.getDataRange();
  
  return data;
}

function onOpen(e) {
   updateScriptProperties();
   //var rowDateInfo =[["September 2015"]];
   //var today = new Date();
   //var sendDate = new Date(today.getFullYear(), today.getMonth()+1, 0);
   //var dueDate = new Date(today.getFullYear(), today.getMonth()+1, 5);
   //rowDateInfo[0].push(sendDate.toString());
   //rowDateInfo[0].push(dueDate.toString());
   //setScriptProperty("monthRowData",JSON.stringify(rowDateInfo))
   
   //SpreadsheetApp.getUi().alert(test.toString() + " " + test2.toString());
  
   //custom string format to date
   //var test = getDateFromMN_Y("October 2015");
   //Date.toString() back to date
   //var test2 = new Date(Date.parse(sendDate.toString()));
}

function onChange(e) {
  //runs after new row is inserted. active range seems to be an empty row
  var props = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger requires authorization that has not
  // been granted yet; if so, warn the user via email. This check is required
  // when using triggers with add-ons to maintain functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to re-authorize; the normal trigger action is not
    // conducted, since it requires authorization first. Send at most one
    // "Authorization Required" email per day to avoid spamming users.
    var lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    //if (lastAuthEmailDate != today) {
      //if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('AuthorizationEmail');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = addonTitle;
        var message = html.evaluate();
        MailApp.sendEmail("clayjacobs245@gmail.com",
            'Authorization Required',
            message.getContent(), {
                name: addonTitle,
                htmlBody: message.getContent()
            }
        );
      //}
      props.setProperty('lastAuthEmailDate', today);
    //}
  } else {
    // Authorization has been granted, so continue to respond to the trigger.
    // Main trigger logic here.
    //SpreadsheetApp.getUi().alert("test " + e.changeType);
    //Logger.log(typeOe.)
    //var rowRange = SpreadsheetApp.getActiveRange();
    //var sheetName = rowRange.getSheet().getName();
    if (sheetName == "Rent") {
       if (e.changeType == "INSERT_ROW") {
      
          //var billsSpread = SpreadsheetApp.openById("1-3UdWZzfnIZYVAtiM7khCvUoW-wPf-ewxGL6_zBlAMk");
          var done = false;
          var title = "Operation on row requires password";
          var message = "You should't perform this action in the current spreadsheet because it will be performed automatically." +
            "To run this action anyway, enter the password below.";
          if (addPasswordLock("4289&64", title, message)) {
             updateScriptProperties();
          }
          else {
             restoreLastRentData();
          }
          rowRange.getSheet().deleteRow(rowRange.getRowIndex());
       }
       else if (e.changeType == "REMOVE_ROW") {
          var response = SpreadsheetApp.getUi().alert("Removing rows", "Unfortunately, you can't remove just any row, but " +
            "you can remove rows for completed months. Would you like to remove these rows?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
          if (response == SpreadsheetApp.getUi().Button.YES) {
             
             removeCompletedRows();
          }
       }
    }  
  }
}



function updateScriptProperties() {
  //rent sheet properties
  //var rentValidations = JSON.stringify(getCurRentData().getDataValidations());
  var rentDataString = JSON.stringify(getCurRentData().getValues());
  setScriptProperty("rentData", rentDataString);
  var rentData = getCurRentData().getValues();
  var monthRows = [];
  for (var i=1; i<rentData.length; i++) {
     var monthString = rentData[i][0];
     var monthStartDate = getDateFromMN_Y(monthString);
     var sendDate = new Date(monthStartDate.getFullYear(), monthStartDate.getMonth()+1, 0);
     var dueDate = new Date(monthStartDate.getFullYear(), monthStartDate.getMonth()+1, 5);
     var monthInfo = [monthString, sendDate.toString(), dueDate.toString()];
     monthRows.push(monthInfo);
  }
  setScriptProperty("monthRowData", JSON.stringify(monthRows));
}

function restoreLastRentData() {
   var lastRentData = JSON.parse(getScriptProperty("rentData"));
   //var lastRentValidation = JSON.parse(getScriptProperty("rentValidations"));
   var rentSheet = getCurRentData().getSheet();
   var rentData = getCurRentData().getValues();
  
   //check same num rows 
   if (lastRentData.length != rentData.length) {
      var numRows = lastRentData.length - rentData.length;
      
      if (numRows > 0) {
         //rows removed, need to add rows
         rentSheet.insertRowsAfter(getCurRentData().getLastRow(), numRows);
      }
      else {
         //rows added, need to remove rows
         numRows = numRows*-1;
         rentSheet.deleteRows(getCurRentData().getLastRow()-numRows+1, numRows);
      }
      rentSheet = getCurRentData().getSheet();
      rentData = getCurRentData().getValues();
   }
  
   //check same num cols 
   if (lastRentData[0].length != rentData[0].length) {
      var numCols = lastRentData[0].length - rentData[0].length;
      
     
      if (numCols > 0) {
         //cols removed, need to add cols
         rentSheet.insertColumnsAfter(getCurRentData().getLastColumn(), numCols);
      }
      else {
         //cols added, need to remove cols
         numCols = numCols*-1;
         rentSheet.deleteColumns(getCurRentData().getLastColumn()-numCols+1, numCols);
      }
      rentSheet = getCurRentData().getSheet();
      rentData = getCurRentData().getValues();
   }
  
   var workflowValues = ["To Do","Writing Check","Check Written","Gave Check to Clay","Check Mailed"];
   var flowRule = SpreadsheetApp.newDataValidation().requireValueInList(workflowValues, true)
      .setAllowInvalid(false).build(); 
  //possibly missing .withCriteria(SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST, workflowValues)
  
   for (var row = 0; row < lastRentData.length; row++) {
      for (var col = 0; col < lastRentData[row].length; col++) {
         var cell = rentSheet.getRange(row+1, col+1, 1, 1);
         
         cell.clearDataValidations();
         
         if (row == 0) {
            cell.setFontWeight("bold"); 
         }
         else {
            cell.setFontWeight("normal");
            if (col >= 1 && col <= 3) {
               cell.setDataValidation(flowRule);
            }
         }
        
         cell.setValue(lastRentData[row][col]);
         //cell.setDataValidation(lastRentValidation[row][col]);
         cell.setHorizontalAlignment("center");
         

      }
   }
}

function removeCompletedRows() {
   var completedRows = [];
   updateScriptProperties();
   var rentData = getCurRentData().getValues();
   var monthRowDates = JSON.parse(getScriptProperty("monthRowData"));
   
   for (var i=0; i+1<rentData.length; i++) {
   
   }
}

//returns true if correct password was entered, false otherwise
function addPasswordLock(password,title,message) {
   var passwordCorrect = false;
   var isDone = false;
   while (!isDone) {
      var passResponse = SpreadsheetApp.getUi().prompt(title, message, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
      if (passResponse.getSelectedButton() == SpreadsheetApp.getUi().Button.CANCEL || 
          passResponse.getSelectedButton() == SpreadsheetApp.getUi().Button.CLOSE) {
         isDone = true;
      }
      else if(passResponse.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
         if (passResponse.getResponseText() == password) {
            passwordCorrect = true;
            isDone = true;
            SpreadsheetApp.getUi().alert("Success!", "You entered the correct password. This action will now be performed", 
             SpreadsheetApp.getUi().ButtonSet.OK);
         }
         else {
            var badPassResponse = SpreadsheetApp.getUi().alert("Incorrect Password!", "The password you entered was incorrect." +
              "Do you want to make another attempt?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
            isDone = (badPassResponse == SpreadsheetApp.getUi().Button.NO);
         } 
      }
   }
   return passwordCorrect;
}

function setScriptProperty(name,value) {
   var scriptProperties = PropertiesService.getScriptProperties();
   scriptProperties.setProperty(name, value);
}

function getScriptProperty(name) {
   var scriptProperties = PropertiesService.getScriptProperties();
   var dataAsText = scriptProperties.getProperty(name);
   var data = JSON.parse(dataAsText);
   return data;
}


function readStatus() {
  Logger.log(SpreadsheetApp.getActiveSpreadsheet().getName());
  var billsSpread = SpreadsheetApp.openById("spreadsheet id here");
  SpreadsheetApp.setActiveSpreadsheet(billsSpread);
  
   //var billsSpread = SpreadsheetApp.getActiveSpreadsheet();
   var rentSheet = billsSpread.getSheetByName("Rent");

   var data = rentSheet.getDataRange().getValues();
   

   //test to be sure cells contain expected values
   for (var row = 0; row < data.length; row++) {
      for (var col = 0; col < data[row].length; col++) {
         Logger.log('Row ' + row + ', Col ' + col + ': "' + data[row][col] + '"');
         //if(rentSheet.getDataRange().getCell(row+1, col+1).getDataValidation() != null) {
         //Logger.log(rentSheet.getDataRange().getCell(row+1, col+1).getDataValidation().getCriteriaValues());
         //rentSheet.getDataRange().getCell(row+1, col+2).setDataValidation(rentSheet.getDataRange().getCell(row+1, col+1).getDataValidation().copy());
         //}
      }
   }
   
   //Logger.log(getMonthLength());
   
    /*email flags 0=weekly update, 1=monthly reminder, 2=completed*/
   var emailFlags = [0, 0, 0];
   var readyToMail = []; 
   
   var curMonthRow =-1;
   var rowsToUpdate = [];
   for (var row = 1; row < data.length; row++) {
       //Logger.log('Product name: ' + data[row][0]);
       //Logger.log('Product number: ' + data[row][1]);
       if (data[row][0] == getTodayMN_Y()) {
          curMonthRow = row;
       }
       if(data[row][4] != "Completed") {
          rowsToUpdate.push(row);
       }
   }

   
   if (curMonthRow == -1) {
      
      //add new row
      rentSheet.insertRows(2);
      rentSheet.getDataRange().getCell(2, 1).setValue(getTodayMN_Y());
      rentSheet.getDataRange().getCell(2, 2).setValue("To Do");
      rentSheet.getDataRange().getCell(2, 3).setValue("To Do");
      rentSheet.getDataRange().getCell(2, 4).setValue("To Do");
      rentSheet.getDataRange().getCell(2, 5).setValue("Not Started");
      //rentSheet.getDataRange().getCell(2, 6).setValue("0 (None incurred)")
      data = rentSheet.getDataRange().getValues();
      
     
      //update rowsToUpdate
      for(var i=0; i<rowsToUpdate.length; i++) {
         rowsToUpdate[i] = rowsToUpdate[i]+1;
      }
      rowsToUpdate.push(1);
      curMonthRow = 1;
      
      //TODO:Maybe look into row removal logic too.
   }
   
   //update row statuses
   for(var i=0; i < rowsToUpdate.length; i++) {
      var rowNum = rowsToUpdate[i];
      if(data[rowNum][1] == data[rowNum][2] && data[rowNum][1] == data[rowNum][3]) {

         //parse cell text if needed 

         if (data[rowNum][1] == "To Do") {
            rentSheet.getDataRange().getCell(rowNum+1, 5).setValue("Not Started");
            data = rentSheet.getDataRange().getValues();
         }
         else if (data[rowNum][1] == "Check Mailed") {
            rentSheet.getDataRange().getCell(rowNum+1, 5).setValue("Completed");
            data = rentSheet.getDataRange().getValues();
            emailFlags[2] = 1; 
         }
         else {
            rentSheet.getDataRange().getCell(rowNum+1, 5).setValue("In Progress");
            data = rentSheet.getDataRange().getValues();
            if (data[rowNum][1] == "Gave Check to Clay") {
               readyToMail.push(rowNum);
            }
         }
      }
      else {
         rentSheet.getDataRange().getCell(rowNum+1, 5).setValue("In Progress");
         data = rentSheet.getDataRange().getValues();
      }
   }
   updateScriptProperties();   


   //send email logic
   
   //ready to mail message just to Clay
   if(isWithinHour(10) && readyToMail.length > 0) {
      sendReadyEmail(data, readyToMail);
   }
   
   //monthly reminder. X days from the end of the month and current month isn't completed 
   if(isWithinHour(10) && isNumDaysUntilMonthEnd(10) && data[1][4] != "Completed") {
      emailFlags[1] = 1;
   }
   
   //weekly reminder. If there are rows to update that aren't the current row.
   if (isWithinHour(10) && getTodayWeekday() == "Monday" && rowsToUpdate.length > 1) {
       emailFlags[0] = 1;
   }

   if(emailFlags.indexOf(1) != -1) {
      sendEmail(data, emailFlags, rowsToUpdate);
      Logger.log("Sent email");
   }
   else {
      Logger.log("Didn't send email");
   }

}

function sendReadyEmail(data, rows) {
   
   var body = "Checks are ready to mail for the month(s) of ";
   
   for(var i=0; i<rows.length; i++) {
      if (i == 0) {
         body = body + data[rows[i]][0]; 
      }
      else if (i == 1 && rows.length == 2) {
         body = body + " and " + data[rows[i]][0] + ".";
      }
      else {
         body = body + ", ";
         if(i != completed.length -1) {
            body = body + data[rows[i]][0];
         }
         else {
            body = body + "and " + data[rows[i]][0] + ".";
         }
      }
   }
   
   MailApp.sendEmail('e', "Rent Checks Ready", body, {
      name: "Automated message from Bills Spreadsheet"
   });
}

function sendEmail(data, flags, rows) {
   var completed = [];
   var holdHeaders = ["Clay's Hold(s)", "Michael's Hold(s)", "Mytch's Hold(s)", "Nick's Hold(s)"];
   var curMonthHolds = [[], "", "", ""]; //0 = clay, 1 = michael, 2 = mytch, 3 = nick 
   var curMonthCompleted = [0, 0]; // 0 = now, 1 = before
   var michaelsHolds = [];
   var mytchsHolds = [];
   var nicksHolds = [];
   var claysHolds = [];
   
   
   if(data[1][4] == "Completed") {
      curMonthCompleted[1] = 1;
   }
   
   //set up email subject
   var subject = "";
   if(flags[0] == 1) {
      subject = subject + "Weekly Reminder";
   }
   if(flags[1] == 1) {
      if(subject.length > 0) {
         subject = subject + " & ";
      }
      if(isNumDaysUntilMonthEnd(10)) {
         subject = subject + "Monthly 10 Day Reminder";
      }
      else {
         subject = subject + "Monthly Reminder";
      }
   }
   if(flags[2] == 1) {
      if(subject.length > 0) {
         subject = subject + " & ";
      }
      subject = subject + "Completed Payments";
   }
   subject = "Bills Update: " + subject;
   
   
   //parse data into completed and holds arrays
   for(var i=0; i < rows.length; i++) {
      var rowNum = rows[i];
      if(data[rowNum][4] == "Completed") {
         if (rowNum == 1) {
            curMonthCompleted[0] = 1;
         }
         else {
            completed.push(rowNum);
         }
         rows.splice(i, 1);
         i = i-1;
      }
      else {
         for(var j=1; j<4; j++) {
            var curHold = "";
            if(data[0][j] == "Michael's Check") {
               if(data[rowNum][j] == "Gave Check to Clay") {
                  curHold = "Hasn't mailed Michael's check for " + data[rowNum][0];
                  if(rowNum == 1) {
                     curMonthHolds[0].push(curHold);
                  }
                  else {
                     claysHolds.push(curHold);
                  }
               }
               else if (data[rowNum][j] != "Check Mailed") {
                  curHold = "Status for " + data[rowNum][0] + " is still " + data[rowNum][j];
                  if(rowNum == 1) {
                     curMonthHolds[1] = curHold;
                  }
                  else {
                     michaelsHolds.push(curHold);
                  }
               }
            }
            else if(data[0][j] == "Mytch's Check") {
               
               if(data[rowNum][j] == "Gave Check to Clay") {
                  curHold = "Hasn't mailed Mytch's check for " + data[rowNum][0];
                  if(rowNum == 1) {
                     curMonthHolds[0].push(curHold);
                  }
                  else {
                     claysHolds.push(curHold);
                  }
               }
               else if (data[rowNum][j] != "Check Mailed") {
                  curHold = "Status for " + data[rowNum][0] + " is still " + data[rowNum][j];
                  
                  if(rowNum == 1) {
                     curMonthHolds[2] = curHold;
                  }
                  else {
                     mytchsHolds.push(curHold);
                  }
               }
            }
            else if(data[0][j] == "Nick's Check") {
               if(data[rowNum][j] == "Gave Check to Clay") {
                  curHold = "Hasn't mailed Nick's check for " + data[rowNum][0];
                  if(rowNum == 1) {
                     curMonthHolds[0].push(curHold);
                  }
                  else {
                     claysHolds.push(curHold);
                  }
               }
               else if (data[rowNum][j] != "Check Mailed") {
                  curHold = "Status for " + data[rowNum][0] + " is still " + data[rowNum][j];
                  if(rowNum == 1) {
                     curMonthHolds[3] = curHold;
                  }
                  else {
                     nicksHolds.push(curHold);
                  }
               }
            }
         }
      }
   }
   
   
   //setup message body
   var body = "";
   if(flags[2] == 1) {
      if(curMonthCompleted[0] == 1) {
         body = body + "Rent for the current month of "+ data[1][0] + " has been completed\n";
      }
      body = body + "Rent for the month(s) of ";
         
      for(var i=0; i<completed.length; i++) {
         if (i == 0) {
            body = body + data[completed[i]][0]; 
         }
         else if (i == 1 && completed.length == 2) {
            body = body + " and " + data[completed[i]][0];
         }
         else {
            body = body + ", ";
            if(i != completed.length -1) {
               body = body + data[completed[i]][0];
            }
            else {
               body = body + "and " + data[completed[i]][0];
            }
         }
      }
         
      body = body + " has been completed\n";
      
   }
   
   if(flags[1] == 1 || flags[0] == 1) {
   
      if(body.length > 0) {
         body = body + "\n\n";
      }
   
      body = body + "===Current Month's Status===\n\n" + data[1][4].toUpperCase() + "\n\n";
      
      if(curMonthCompleted[1] == 0) {
         for(var i=0; i<curMonthHolds.length; i++) {
            body = body + "   " + holdHeaders[i] + "\n";
            
            if(curMonthHolds[i].length > 0) {
               if(i == 0) {
                  for(var j=0; j<curMonthHolds[i].length; j++) {
                     body = body + "      " + curMonthHolds[i][j] + "\n";
                  }
               }
               else {
                  body = body + "      " + curMonthHolds[i] + "\n";
               }
            }
            else {
               body = body + "      None\n";
            }
         }
      }
      else {
         body = body + "No holds\n";
      }
      
      body = body + "\n\n===Overdue Months===\n\n";
      if (rows.length != 0) {
         body = body + "   " + holdHeaders[0] + "\n";
         if(claysHolds.length > 0) {
            for(var i=0; i<claysHolds.length; i++) {
               body = body + "      "+ claysHolds[i]+"\n";
            }
         }
         else {
            body = body + "      None\n";
         }
         
         
         body = body + "   " + holdHeaders[1] + "\n";
         if(michaelsHolds.length > 0) {
            for(var i=0; i<michaelsHolds.length; i++) {
               body = body + "      "+ michaelsHolds[i]+"\n";
            }
         }
         else {
            body = body + "      None\n";
         }
         
         body = body + "   " + holdHeaders[2] + "\n";
         if(mytchsHolds.length > 0) {
            for(var i=0; i<mytchsHolds.length; i++) {
               body = body + "      "+ mytchsHolds[i]+"\n";
            }
         }
         else {
            body = body + "      None\n";
         }
         
         body = body + "   " + holdHeaders[3] + "\n";
         if(nicksHolds.length > 0) {
            for(var i=0; i<nicksHolds.length; i++) {
               body = body + "      "+ nicksHolds[i]+"\n";
            }
         }
         else {
            body = body + "      None\n";
         }
      }
      else {
         body = body + "None overdue\n";
      }
   }
   
   
   MailApp.sendEmail('Mailing list goes here', subject, body, {
      name: "Automated message from Bills Spreadsheet"
   });
}

Date.prototype.getMonthName = function(lang) {
    lang = lang && (lang in Date.locale) ? lang : 'en';
    return Date.locale[lang].month_names[this.getMonth()];
};

Date.prototype.getMonthNameShort = function(lang) {
    lang = lang && (lang in Date.locale) ? lang : 'en';
    return Date.locale[lang].month_names_short[this.getMonth()];
};

Date.locale = {
    en: {
       month_names: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
       month_names_short: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    }
};

function getTodayMN_Y() {
   var today = new Date();
   return today.getMonthName() + " " + today.getFullYear().toString();
}

function getAnyDayMN_Y(date) {
   return date.getMonthName() + " " + date.getFullYear().toString();
}

function getDateFromMN_Y(format) {
   var dateParts = format.split(" ");
  
   var lang = lang && (lang in Date.locale) ? lang : 'en';
   if (Date.locale["en"].month_names.indexOf(dateParts[0]) != -1) {
      var monthNum = Date.locale["en"].month_names.indexOf(dateParts[0]);
      var yearNum = dateParts[1].toString();
      var converted = new Date(yearNum, monthNum, 1);
      return converted;
   }
   else {
      return null; 
   } 
}

function getTodayWeekday() {
   var today = new Date();
   var weekdays = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
   return weekdays[today.getDay()];
}

function getTimeFormat() {
   var today = new Date();
   if (today.getHours()+1 > 12) {
      var hours = today.getHours()-11;
      return hours.toString() + ":" + today.getMinutes.toString() + " PM";
   }
   else {
      var hours = today.getHours()+1;
      return hours.toString() + ":" + today.getMinutes.toString() + " AM";
   }
}

function isWithinHour(hour) {
   var today = new Date();
   if(today.getHours()+1 == hour) {
     return true;
   }
   return false;
}

function isNumDaysUntilMonthEnd(days) {
  var today = new Date();
  var lastMonthDayDate = new Date(today.getFullYear(), today.getMonth()+1, 0);
  //var test = new Date(today.getFullYear(), 12, 0);
  var monthLength = lastMonthDayDate.getDate();
  
  if(monthLength-today.getDate() == days) {
     return true;
  }
  else {
     return false;
  }
  
}