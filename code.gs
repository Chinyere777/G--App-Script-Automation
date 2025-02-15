function processFormSubmissions() {
  console.log("Starting processFormSubmissions");
  var lock = LockService.getScriptLock();
  var lockAcquired = lock.tryLock(300000); // Wait for up to 5 minutes for other processes to finish
  console.log("Lock acquired: " + lockAcquired);

  if (!lockAcquired) {
    console.log("Failed to acquire lock. Exiting.");
    return;
  }

  try {
    var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form responses"); // Replace with your actual sheet name
    var lastRow = formSheet.getLastRow();

    console.log("Last row in form sheet: " + lastRow);

    if (lastRow <= 1) {
      console.log("No form responses to process.");
      return; // Exit the function if there are no form responses
    }

    var targetSSId1 = "Google_Spreadsheet_ID"; // Backend Spreadsheet
    var targetSSId2 = "Google_Spreadsheet_ID"; // Rejection Validation Spreadsheet

    var formValues = formSheet.getRange(2, 2, lastRow - 1, 4).getValues(); // Get all form responses at once
    console.log("Number of form responses: " + formValues.length);

    var targetSS1 = SpreadsheetApp.openById(targetSSId1);
    var targetSS2 = SpreadsheetApp.openById(targetSSId2);

    var sheetUpdates = {};

    formValues.forEach(function(row, index) {
      console.log("Processing row " + (index + 1));
      var firstName = row[0];
      var slrn = row[1];
      var slrnID = row[2];
      var reasonForReject = row[3];

      console.log("Finding sheet for: " + firstName);
      var teamMemberSheetName = findTeamMemberSheetName(firstName, targetSSId1);
      console.log("Sheet found: " + teamMemberSheetName);

      if (teamMemberSheetName) {
        if (!sheetUpdates[teamMemberSheetName]) {
          sheetUpdates[teamMemberSheetName] = {
            targetSheet1: targetSS1.getSheetByName(teamMemberSheetName),
            targetSheet2: targetSS2.getSheetByName(teamMemberSheetName),
            updates: []
          };
        }

        sheetUpdates[teamMemberSheetName].updates.push({
          slrn: slrn,
          slrnID: slrnID,
          reasonForReject: reasonForReject
        });
      } else {
        console.log("No matching sheet found for: " + firstName);
      }
    });

    Object.keys(sheetUpdates).forEach(function(sheetName) {
      console.log("Updating sheet: " + sheetName);
      var update = sheetUpdates[sheetName];
      var targetSheet1 = update.targetSheet1;
      var targetSheet2 = update.targetSheet2;

      // Get the header row to find column indices
      var headers1 = targetSheet1.getRange(1, 1, 1, targetSheet1.getLastColumn()).getValues()[0];
      var headers2 = targetSheet2.getRange(1, 1, 1, targetSheet2.getLastColumn()).getValues()[0];

      var columnIndices1 = getColumnIndices(headers1);
      var columnIndices2 = getColumnIndices(headers2);

      if (validateColumnIndices(columnIndices1) && validateColumnIndices(columnIndices2)) {
        var lastRow1 = getLastRowWithData(targetSheet1);
        var lastRow2 = getLastRowWithData(targetSheet2);

        var today = new Date();
        today.setHours(0, 0, 0, 0);  // Clear the time part of the date

        var dateAlreadyExists1 = false;
        if (lastRow1 > 1) {
          var dateValues1 = targetSheet1.getRange(2, columnIndices1.dateColumn, lastRow1 - 1).getValues().flat();
          dateAlreadyExists1 = dateValues1.some(function(cellDate) {
            cellDate = new Date(cellDate);
            cellDate.setHours(0, 0, 0, 0);
            return cellDate.getTime() === today.getTime();
          });
        }

        var formattedDate = formatDate(today);
        var newRows1 = [];
        var newRows2 = [];

        if (!dateAlreadyExists1 && lastRow1 > 2) {
          // Insert a new row to demarcate the date change
          targetSheet1.insertRowBefore(lastRow1 + 1);
          targetSheet2.insertRowBefore(lastRow2 + 1);

          // Merge the new row to create visual separation
          var demarcationRange1 = targetSheet1.getRange(lastRow1 + 1, 1, 1, targetSheet1.getLastColumn());
          var demarcationRange2 = targetSheet2.getRange(lastRow2 + 1, 1, 1, targetSheet2.getLastColumn());

          demarcationRange1.merge();
          demarcationRange2.merge();
          demarcationRange1.setValue('');
          demarcationRange2.setValue('');

          // Move the new data down by one row
          lastRow1 += 1;
          lastRow2 += 1;
        }

        update.updates.forEach(function(item, index) {
          var slrnPrefix = item.slrn.substring(0, 5).toUpperCase(); // Get the first 5 characters of SLRN
          var baseUrl = "https://caims.azurewebsites.net/";
          var segment = "";

          switch(slrnPrefix) {
            case "ECGBD":
              segment = "customers/";
              break;
            case "ECGDT":
            case "IEDCDT":  // New prefix added here
              segment = "transformers/";
              break;
            case "ECGLT":
              segment = "poles/";
              break;
            default:
              console.log("Unknown SLRN prefix: " + slrnPrefix);
              segment = "customers/"; // Default to customers if unknown prefix
          }

          var slrnLink1 = baseUrl + segment + "data-quality/" + item.slrnID;
          var slrnLink2 = baseUrl + segment + "edit/" + item.slrnID;

          // Only add the date to the first row of the day
          var dateForThisRow = (index === 0 && !dateAlreadyExists1) ? formattedDate : '';

          var newRow1 = createNewRow(columnIndices1, dateForThisRow, slrnLink1, item.slrn, item.reasonForReject);
          var newRow2 = createNewRow(columnIndices2, dateForThisRow, slrnLink2, item.slrn, item.reasonForReject);

          newRows1.push(newRow1);
          newRows2.push(newRow2);
        });

        // Add all new rows at once
        if (newRows1.length > 0) {
          console.log("Adding " + newRows1.length + " new rows to " + sheetName);
          
          // Check if the sheet is empty (only has headers)
          if (lastRow1 === 1) {
            lastRow1 = 1; // Start adding from the second row (right after headers)
            lastRow2 = 1;
          }
          
          targetSheet1.getRange(lastRow1 + 1, 1, newRows1.length, newRows1[0].length).setValues(newRows1);
          targetSheet2.getRange(lastRow2 + 1, 1, newRows2.length, newRows2[0].length).setValues(newRows2);

          // Set SLRN hyperlinks
          setSlrnHyperlinks(targetSheet1, lastRow1 + 1, columnIndices1.slrnColumn, newRows1);
          setSlrnHyperlinks(targetSheet2, lastRow2 + 1, columnIndices2.slrnColumn, newRows2);


          /*// Number the rows
          numberRows(targetSheet1, columnIndices1.snColumn);
          numberRows(targetSheet2, columnIndices2.snColumn);*/



        
        } else {
          console.log("No new rows to add to " + sheetName);
        }
      } else {
        console.log("Invalid column indices for " + sheetName);
      }
    });

    // Remove the processed entries from the form responses sheet
    console.log("Deleting processed rows from form responses sheet");
    formSheet.deleteRows(2, lastRow - 1);
  } catch (error) {
    console.error("Error occurred: " + error.message);
    console.error("Stack trace: " + error.stack);
  } finally {
    lock.releaseLock();
    console.log("Lock released");
  }
}

function getColumnIndices(headers) {
  return {

    //snColumn: headers.indexOf("S/N") + 1,
    
    dateColumn: headers.indexOf("Date") + 1,
    slrnColumn: headers.indexOf("SLRN") + 1,
    reasonForRejectColumn: headers.indexOf("Reason For Rejection") + 1,
    statusColumn: headers.indexOf("Status") + 1
  };
}

function validateColumnIndices(indices) {
  return Object.values(indices).every(index => index > 0);
}

function createNewRow(indices, formattedDate, slrnLink, slrn, reasonForReject) {
  var newRow = new Array(Math.max(...Object.values(indices)));
  newRow[indices.dateColumn - 1] = formattedDate;
  newRow[indices.slrnColumn - 1] = [slrnLink, slrn]; // Store both the link and SLRN
  newRow[indices.reasonForRejectColumn - 1] = reasonForReject;
  newRow[indices.statusColumn - 1] = "Pending";
  return newRow;
}

//Pervious SLRN cloumn glitch function
/*function setSlrnHyperlinks(sheet, startRow, slrnColumn, newRows) {
  newRows.forEach((row, index) => {
    var cell = sheet.getRange(startRow + index, slrnColumn);
    var [fullLink, slrn] = row[slrnColumn - 1];
    cell.setFormula('=HYPERLINK("' + fullLink + '","' + slrn + '")');
    cell.setFontStyle('normal').setFontLine('none');
  });
}*/

function setSlrnHyperlinks(sheet, startRow, slrnColumn, newRows) {
  newRows.forEach((row, index) => {
    var cell = sheet.getRange(startRow + index, slrnColumn);
    var [fullLink, slrn] = row[slrnColumn - 1];
    
    // Error checking and fallback
    if (!fullLink) {
      console.error(`Missing link for row ${startRow + index}`);
      return; // Skip this iteration
    }
    
    if (!slrn) {
      console.warn(`Missing SLRN for row ${startRow + index}, extracting from URL`);
      // Extract SLRN from the URL
      var urlParts = fullLink.split('/');
      slrn = urlParts[urlParts.length - 1]; // Assume SLRN is the last part of the URL
    }
    
    cell.setFormula('=HYPERLINK("' + fullLink + '","' + slrn + '")');
    cell.setFontStyle('normal').setFontLine('none');
  });
}

function formatDate(date) {
  var daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  var dayOfWeek = daysOfWeek[date.getDay()];
  var month = months[date.getMonth()];
  var dayOfMonth = date.getDate();
  var year = date.getFullYear();
  
  return dayOfWeek + ' ' + month + ' ' + dayOfMonth + ', ' + year;
}

function findTeamMemberSheetName(name, targetSSId) {
  console.log("Searching for sheet name for: " + name);
  var targetSS = SpreadsheetApp.openById(targetSSId);
  var sheets = targetSS.getSheets();
  var lowerCaseName = name.toLowerCase();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName().toLowerCase();
    var [firstName, lastName] = sheetName.split(" ");
    
    if (lowerCaseName === sheetName || lowerCaseName === firstName || lowerCaseName === lastName) {
      console.log("Found matching sheet: " + sheets[i].getName());
      return sheets[i].getName();
    }
  }
  
  console.log("No matching sheet found for: " + name);
  return null;
}



/*function numberRows(sheet, snColumn) {
  var lastRow = sheet.getLastRow();
  var snValues = [];
  var rowCount = 1; // Start numbering from 1
  for (var i = 2; i <= lastRow; i++) {
    var cell = sheet.getRange(i, snColumn);
    if (!cell.getMergedRanges().length) {
      snValues.push([rowCount]);
      rowCount++;
    } else {
      snValues.push(['']);
    }
  }
  sheet.getRange(2, snColumn, snValues.length, 1).setValues(snValues);
}*/



function updateStatusColumns() {
  console.log("Starting updateStatusColumns");
  var targetSSId1 = "Google_Spreadsheet_ID"; // Backend Spreadsheet
  var targetSSId2 = "Google_Spreadsheet_ID"; // Rejection Validation Spreadsheet

  var targetSS1 = SpreadsheetApp.openById(targetSSId1);
  var targetSS2 = SpreadsheetApp.openById(targetSSId2);

  var sheets1 = targetSS1.getSheets();
  var sheets2 = targetSS2.getSheets();

  for (var i = 0; i < sheets1.length; i++) {
    var sheetName = sheets1[i].getName();
    console.log("Processing sheet: " + sheetName);
    var targetSheet1 = sheets1[i];
    var targetSheet2 = sheets2.filter(function(sheet) {
      return sheet.getName() === sheetName;
    })[0];

    if (targetSheet2) {
      var headers1 = targetSheet1.getRange(1, 1, 1, targetSheet1.getLastColumn()).getValues()[0];
      var statusColumn1 = headers1.indexOf("Status") + 1;

      var headers2 = targetSheet2.getRange(1, 1, 1, targetSheet2.getLastColumn()).getValues()[0];
      var statusColumn2 = headers2.indexOf("Status") + 1;

    
    } else {
      console.log("Matching sheet not found in second spreadsheet for " + sheetName);
    }
  }
  console.log("Finished updateStatusColumns");
}

function getLastRowWithData(sheet) {
  var column = 4; // You can change this if you want to check a specific column
  var values = sheet.getRange(1, column, sheet.getLastRow()).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1;
    }
  }
  return 1; // Return 1 if no data found (only headers)
}

