/**
 * @fileoverview A collection of Apps Script functions I've had to write for professional projects.
 * These are collected here for display in a portfolio of my work.
 */


/**
 * Copies a plain text file (likely a csv or tab-delimited txt) into a spreadsheet
 * 
 * @param {string} sourceFolder ID of folder containing source file in Google Drive
 * @param {string} sourceFile Name of file whose text is being copied
 * @param {string} targetFile ID of the Google sheets file where copied text will be pasted
 * @param {string} targetTab Name of sheet/tab within the target file
 */
function copyText(sourceFolder, sourceFile, targetFile, targetTab){
  //Open source text file and copy contents
  var file = DriveApp.getFolderById(sourceFolder).getFilesByName(sourceFile).next();
  var text = file.getBlob().getDataAsString();
  
  //Break up text file into cells by delimiters. Delimiter is determined by file format
  var extension = sourceFile.split(".")[1];
  if (extension == "csv") {
    var result = Utilities.parseCsv(text);
  } else if (extension == "txt") {
    var result = Utilities.parseCsv(text, '\t');
  }

  //Write resulting text to spreadsheet
  var targetSheet = SpreadsheetApp.openById(targetFile).getSheetByName(targetTab);
  targetSheet.getRange(1,1,result.length,result[0].length).setValues(result);
}


/**
 * Looks for names that have previously been marked in an attendance form according to a list in attached spreadsheet 
 * and flags them
 */
function updateNames() {
  //Replace with IDs of whatever form and spreadsheet
  const formID = "";
  const sheetID = "";
  //Whatever note your names will be flagged with (e.g. "Absent")
  const flagMessage = "";
  
  //Open form and save as variable
  const attendanceForm = FormApp.openById(formID);

  //Open sheet of previous responses and save as array
  var absenceList = SpreadsheetApp.openById(sheetID)
    .getSheetByName('Attendance List').getDataRange().getValues(); //Rename spreadsheet tab here if using other name
  absenceList = absenceList.filter(x => x.filter(Boolean).length > 0);
  absenceList = absenceList.map(x => x[0]);

  //Get IDs for each form item
  const formItems = attendanceForm.getItems();
  formItems.forEach(function(item) {
    console.log(item.getTitle() + ': ' + item.getId());
  });

  const ids = ["", ""]; //Replace with applicable IDs for form questions after running above
  
  //Iterate through each ID listed above (i.e. every question with names)
  ids.forEach(function(section) {
    //Get list of applicant names for the given section
    var checkboxItem = attendanceForm.getItemById(section).asCheckboxItem();
    var nameOptions = checkboxItem.getChoices();
    var names = nameOptions.map(x => x.getValue());

    //Checks every name against the spreadsheet list
    //If their name appears and wasn't already marked, it's appended to reflect this in the form going forward
    var updatedNames = [];
    var updatedNamesDouble = [];
    names.forEach(function(row) {
      if (absenceList.includes(row)) {
        if (!row.endsWith(']')) {
          row = row + " [" + flagMessage + "]";
        }  
      }
      updatedNames.push(row);
      updatedNamesDouble.push([row]);
    });

    //Replace form choices with the updated name list
    checkboxItem.setChoiceValues(updatedNames);
  });
}


/**
 * Sends an email listing present/absent students given a spreadsheet with these lists
 */
function sendAttendanceEmail() {
  //Extract existing list of present students from spreadsheet and remove any blank values
  var presentList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Present").getDataRange().getValues();
  presentList = presentList.filter(function (row) {
    return row[0].length > 0
  })
  //Same process but for absent students
  var absentList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Absent").getDataRange().getValues();
  absentList = absentList.filter(function (row) {
    return row[0].length > 0
  })

  //Enter names into html tables to be added to email
  const presentTable = tableHTML(presentList);
  const absentTable = tableHTML(absentList);

  //HTML used to write the email itself
  const htmlForEmail = 
    `<html>
      <head>
        <base target="_top">
      </head>
      <body>
        <p>Hi team,</p>
        <br><p>The following students were present today:</p>
        ` + presentTable + `
        <br><p>Tthe following students were absent:</p>
        ` + absentTable + `
        <br><p>Have a great rest of your day!</p>
        <p>Aaron</p>
      </body>
    </html>`

  //Add the recipient and subject line as variables
  const recipient = ""; //Add address of recipient here (separate by commas if more than one)
  const date = new Date().toLocaleDateString("en-US");
  const subject = "Attendance Report - " + date;

  //Actually send the email
  GmailApp.sendEmail(recipient, subject, "", { htmlBody: htmlForEmail });
}


/**
 * Helper function to assemble a 2D array into an HTML table string
 * @param {array} array An n x 2 array of names, each to become a row in the table
 * @return {string} table A string containing the full HTML needed to create a table with the values
 */
function tableHTML(array) {
  var table = "<table>\n"
  array.forEach(function (row) {
    table = table +
    `<tr>
      <td>`+ row[0] +`</td>
      <td>`+ row[1] +`</td>
    </tr>`
  });
  table = table + "</table>"
  return table;
}
