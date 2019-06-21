/*
 * script to export data in all sheets in the current spreadsheet as individual text files
 * files will be named according to the name of the sheet
 * author: Mike Han
*/

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileMenuEntries = [{name: "Pipe-delimited files", functionName: "saveAsPipeFile"}];
  ss.addMenu("Export Flat File", fileMenuEntries);
};

function saveAsFile(sep, ext) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var fileTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'YYYYMMDD hh.mm.ss');
  var folderDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'YYYYMMDD');
  var folder = DriveApp.getFoldersByName("Flat File Exports - " + folderDate);
  if (folder.hasNext()) {
    folder = folder.next();
  }
  else {
    // create a folder from the name of the spreadsheet
    folder = DriveApp.createFolder("Flat File Exports - " + folderDate);
  }
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append extension to the sheet name
    var fileName = sheet.getName() + " " + fileTime + ext;
    // convert all available sheet data to txt format
    var flatFile = convertRangeToFile_(fileName, sheet, sep);
    // create a file in the Docs List with the given name and the data
    folder.createFile(fileName, flatFile);
  }
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

function convertRangeToFile_(fileName, sheet, sep) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var flatFile = undefined;

    // loop through the data in the range and build a string with the custom separator data
    if (data.length > 1) {
      var content = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(sep) != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length) {
          content += data[row].join(sep) + "\r\n";
        }
        else {
          content += data[row];
        }
      }
      flatFile = content;
    }
    return flatFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

function saveAsPipeFile() {
  saveAsFile(sep="|", ext=".txt")
}
