function onOpen() {
  //Adds the feedback menu to the translation editors schedule sheet
  SpreadsheetApp.getUi()
      .createMenu('Wallace')
      .addItem('Translation Editing Feedback', 'teFeedback')
      .addItem('View Feedback', 'viewFeedback')
      .addToUi();  
}

function setPermissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var name = ss.getName().split(" ")[0];
  if (!(name == "Chris")) {
    var email_list_sheet = SpreadsheetApp.openById("1n3PTIw3sO1oxZSNdE_vGHdE7kd8bYeedZOr2kT96dXE")
                                   .getSheetByName("Editor List & Complaints Email");
    var email_list = email_list_sheet.getRange(2, 1, email_list_sheet.getLastRow(), 2).getValues();
    email_list = email_list.join().split(",");
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (protections.length == 0) {
        var protection = sheets[i].protect();
        protection.removeEditor(email_list[email_list.indexOf(name) + 1])
      }
    }
  }
}

function teFeedback() {
  //Opens the feedback sidebar
  
  var months = {"JAN": "01",
                "FEB": "02",
                "MAR": "03",
                "APR": "04",
                "MAY": "05", 
                "JUN": "06",
                "JUL": "07",
                "AUG": "08",
                "SEP": "09",
                "OCT": "10",
                "NOV": "11",
                "DEC": "12"}
  
  var user = Session.getActiveUser().getEmail().split("@")[0].substr(0,1).toUpperCase() + Session.getActiveUser().getEmail().split("@")[0].substr(1)
  var cases = getCases(user);
  if (cases == false) {
    SpreadsheetApp.getUi().alert("You have no cases assigned for this month.", SpreadsheetApp.getUi().ButtonSet.OK)
  }
  else {
    var html = HtmlService.createTemplateFromFile('Index');
    var sheet_name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    html.cases = cases;
    html.user = user;
    html.month = "-" + months[sheet_name.split(" ")[0].slice(0, 3).toUpperCase()] + "-" + sheet_name.split(" ")[1].slice(2, 4)
    html = html.evaluate().setTitle("Translation Editing Feedback");
    SpreadsheetApp.getUi().showSidebar(html);
  }
}


function viewFeedback() {
  var html = HtmlService.createTemplateFromFile("View_Feedback_Index");
  html = html.evaluate()
  .setTitle("View Feedback")
  .setHeight(450)
  .setWidth(750)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  
  SpreadsheetApp.getUi().showModalDialog(html, "Submitted Feedback")
}


function getCases(user) {
  //Searches the case ID columns of the active sheet for TE entries (with a green background).
  //Does not return future cases
  //Returns an array of the case IDs.
  
  var months = {0: "January", 
                1: "February",
                2: "March",
                3: "April",
                4: "May",
                5: "June",
                6: "July",
                7: "August",
                8: "September",
                9: "October",
                10: "November",
                11: "December"}
  
  var color_dic = {"Chris": "#00ff00",
                   "Mark": "#76a5af",
                   "Yankz": "#f1c232",
                   "Hannah": "#93c47d",
                   "Ben": "#b4a7d6",
                   "Richard": "#cc4125", 
                   "Kimberlee": "",
                   "Samuel": "#00ffff"}
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month_sheet = ss.getActiveSheet();
  var cases = [];
  var colors = []
  
  //Find the start row for each week in the schedule sheet
  var week_indexes = []
  var column_A_values = month_sheet.getRange("A:A").getValues();
  var sheet_height = column_A_values.length;
  for (var i = 0; i < sheet_height; i++) {
    if (column_A_values[i][0] == "Date") {
      week_indexes.push(i);
    }
  }
  week_indexes.push(sheet_height)
  
  
  //Extract all data from the case ID columns and the background colors for these cells.
  for (var i = 0; i < week_indexes.length - 1; i++) {
    for (var j = 0; j < 7; j++) {
      cases = cases.concat(month_sheet.getRange(week_indexes[i] + 1, (3*j+2), week_indexes[i+1] - week_indexes[i], 1).getValues());
      colors = colors.concat(month_sheet.getRange(week_indexes[i] + 1, (3*j+2), week_indexes[i+1] - week_indexes[i], 1).getBackgrounds());
    }
  }
  
  //Find the index of tomorrows date if the current month is selected (add 16 hours to convert to Taipei time)
  var date = new Date()
  if (Session.getScriptTimeZone() == "America/Los_Angeles") {
    date.addHours(16);
  }
  var month = date.getMonth();
  date = date.getDate() + 1;
  var date_index = 0;
  if (month_sheet.getName().split(" ")[0] == months[month]) {
    var found = false;
    while (!found && date_index < cases.length) {
      if (cases[date_index][0] == date) {
        found = true;
      }
      else {
        date_index += 1;
      }
    }
  }
  else {
    date_index = colors.length
  }

  //Add the case IDs to the assigned_te_cases array if the case is TE (green background).
  var assigned_te_cases = [];
  for (var i = 0; i < date_index; i++) {
    if (colors[i][0] == color_dic[user]) {
      assigned_te_cases.push(cases[i][0]);
    }
  }

  return assigned_te_cases;
}


function submitFeedback(values) {
  
  //Delete existing feedback with the same case ID
  var feedback_sheet = SpreadsheetApp.openById("13QDsOkVGVPMsbqg0_Qyet8cg3y9ySzR3XV9IlDLmEBs").getSheetByName("TE Feedback");
  var case_ID = values[2];
  var case_ID_column = feedback_sheet.getRange(2, 3, feedback_sheet.getLastRow() - 1, 1).getValues();
  var row_index = -1;
  var i = 0;
  while (row_index == -1 && i < case_ID_column.length) {
    if (case_ID_column[i][0] == case_ID) {
      row_index = i;
    }
    i += 1;
  }
  if (row_index != -1) {
    feedback_sheet.deleteRow(row_index + 2);
  }
  
  //Add the new feedback
  feedback_sheet.getRange(feedback_sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
}



function getOutstanding(cases) {
  //Get list of cases with incomplete feedback for the selected month
  
  var months = {"JAN": " January", 
                "FEB": " February",
                "MAR": " March",
                "APR": " April",
                "MAY": " May",
                "JUN": " June",
                "JUL": " July",
                "AUG": " August",
                "SEP": " September",
                "OCT": " October",
                "NOV": " November",
                "DEC": " December"}
  
  var TPR_Feedback_sheet = SpreadsheetApp.openById("13QDsOkVGVPMsbqg0_Qyet8cg3y9ySzR3XV9IlDLmEBs").getSheetByName("TE Feedback");  
  var complete = TPR_Feedback_sheet.getRange(2, 3, TPR_Feedback_sheet.getLastRow() - 1, 1).getValues();
  var completed_cases = [];
  for (i = 0; i < complete.length; i++) {
    completed_cases.push(complete[i][0].slice(0, 7));
  }
  
  var all_cases = [];
  for (var i = 0; i < Object.keys(cases).length; i++) {
    var day = Object.keys(cases)[i];
    all_cases = all_cases.concat(cases[day]);
  }
      
  
  var incomplete_cases = []
  for (i = 0; i < all_cases.length; i++) {
    if (completed_cases.indexOf(all_cases[i]) == -1) {
      incomplete_cases.push(all_cases[i]);
    }
  }
  
  return incomplete_cases
}


function getFeedback() {
  //Extracts the submitted feedback for the active user
  
  var user = Session.getActiveUser().getEmail().split("@")[0].substr(0,1).toUpperCase() + Session.getActiveUser().getEmail().split("@")[0].substr(1);
  var TE_Feedback_sheet = SpreadsheetApp.openById("13QDsOkVGVPMsbqg0_Qyet8cg3y9ySzR3XV9IlDLmEBs").getSheetByName("TE Feedback");
  var feedback = []
  for (var i = 2; i < TE_Feedback_sheet.getLastRow() + 1; i++) {
    var entry = TE_Feedback_sheet.getRange(i, 2, 1, 7).getValues()[0];
    if (entry[0] == user) {
      entry[6] = entry[6].toString();
      feedback.push(entry.slice(1));
    }
  }
  
  return feedback.reverse();
}

function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}



Date.prototype.addHours = function(h){
    this.setHours(this.getHours()+h);
    return this;
}