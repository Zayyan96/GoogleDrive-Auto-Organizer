function moveNewFilesToFolder() {
  var sheetId = "12Cma0qxmsz89EsaK2U6wALvPud2QLwZHGx0UzmS2fIM"; // Replace with the ID of your Google Sheet file
  var sheet = SpreadsheetApp.openById(sheetId);
  if (!sheet) {
    Logger.log("Sheet named 'Assignment' not found in the specified Google Sheet file.");
    return;
  }
  var data = sheet.getDataRange().getValues();
  var lastProcessedRow = getLastProcessedRow();

  // Create the "MovedFiles" sheet if it doesn't exist
  var movedFilesSheet = sheet.getSheetByName("MovedFiles");
  if (!movedFilesSheet) {
    sheet.insertSheet("MovedFiles");
    movedFilesSheet = sheet.getSheetByName("MovedFiles");
    movedFilesSheet.appendRow(["Timestamp", "Fullname", "Student ID", "Student Grade", "Assignment Subject", "Assignment Name", "New File Location"]);
  }

  // Loop through the rows in the sheet starting from the next row after the last processed row
  for (var rowIndex = lastProcessedRow + 1; rowIndex < data.length; rowIndex++) {
    var fileURL = data[rowIndex][6] || data[rowIndex][8] || data[rowIndex][10] || data[rowIndex][12] || data[rowIndex][15] || data[rowIndex][17] || data[rowIndex][19] || data[rowIndex][21] || data[rowIndex][24] || data[rowIndex][26] || data[rowIndex][28] || data[rowIndex][30] || data[rowIndex][33] || data[rowIndex][35] || data[rowIndex][37] || data[rowIndex][39] || data[rowIndex][41] || data[rowIndex][43];

    var grandfatherFolderLink = "https://drive.google.com/drive/folders/13ifbzhmcsqhB5XM92Nyvqmb6-I3vVAde";
    var parentFolderName = data[rowIndex][3];
    var childFolderName = data[rowIndex][4] || data[rowIndex][13] || data[rowIndex][22] || data[rowIndex][31];
    var subChildFolderName = data[rowIndex][5] || data[rowIndex][7] || data[rowIndex][9] || data[rowIndex][11] || data[rowIndex][14] || data[rowIndex][16] || data[rowIndex][18] || data[rowIndex][20] || data[rowIndex][23] || data[rowIndex][25] || data[rowIndex][27] || data[rowIndex][29] || data[rowIndex][32] || data[rowIndex][34] || data[rowIndex][36] || data[rowIndex][38] || data[rowIndex][40] || data[rowIndex][42];
    var grandfatherFolder = DriveApp.getFolderById(getFolderIdFromLink(grandfatherFolderLink));
    var parentFolder = getOrCreateFolder(parentFolderName, grandfatherFolder);
    var childFolder = getOrCreateFolder(childFolderName, parentFolder);
    var subChildFolder = getOrCreateFolder(subChildFolderName, childFolder);

    var newFileName = data[rowIndex][1] + " - " + data[rowIndex][2] + ".pdf";

    // Check if the file with the same name exists in the subChildFolder
    var fileExists = subChildFolder.getFilesByName(newFileName).hasNext();

    // If the file already exists, rename the new file
    if (fileExists) {
      var counter = 1;
      var newFileNameWithCounter = newFileName.replace(".pdf", " (" + counter + ").pdf");
      while (subChildFolder.getFilesByName(newFileNameWithCounter).hasNext()) {
        counter++;
        newFileNameWithCounter = newFileName.replace(".pdf", " (" + counter + ").pdf");
      }
      newFileName = newFileNameWithCounter;
    }

    try {
      var fileBlob = DriveApp.getFileById(getFileIdFromURL(fileURL)).getBlob();
      var newFile = subChildFolder.createFile(fileBlob).setName(newFileName);
      DriveApp.getFileById(getFileIdFromURL(fileURL)).setTrashed(true);

      var timestamp = new Date();
      var fullName = data[rowIndex][1];
      var studentId = data[rowIndex][2];
      var studentGrade = data[rowIndex][3];
      var assignmentSubject = data[rowIndex][4] || data[rowIndex][13] || data[rowIndex][22] || data[rowIndex][31];
      var assignmentName = data[rowIndex][5] || data[rowIndex][7] || data[rowIndex][9] || data[rowIndex][11] || data[rowIndex][14] || data[rowIndex][16] || data[rowIndex][18] || data[rowIndex][20] || data[rowIndex][23] || data[rowIndex][25] || data[rowIndex][27] || data[rowIndex][29] || data[rowIndex][32] || data[rowIndex][34] || data[rowIndex][36] || data[rowIndex][38];

      // Update the "MovedFiles" sheet with the new file location
      movedFilesSheet.appendRow([timestamp, fullName, studentId, studentGrade, assignmentSubject, assignmentName, newFile.getUrl()]);
    } catch (error) {
      Logger.log("Error processing file with URL: " + fileURL);
      Logger.log(error.message);
      var subject = "Failed to move files for " + data[rowIndex][1];
      var htmlBodyContent =
        "<p>Dear " + data[rowIndex][1] + ",</p>" +
        "<p>Failed to move files for " + data[rowIndex][1] + "</p>" +
        "<ul>" +
        "<li><strong>Student ID:</strong> " + data[rowIndex][2] + "</li>" +
        "<li><strong>Name:</strong> " + data[rowIndex][1] + "</li>" +
        "<li><strong>Grade:</strong> " + (data[rowIndex][3] || "Not specified") + "</li>" +
        "<li><strong>Assignment's Subject:</strong> " + (data[rowIndex][4] || data[rowIndex][13] || data[rowIndex][22] || data[rowIndex][31]) + "</li>" +
        "<li><strong>Assignment Name:</strong> " + (data[rowIndex][5] || data[rowIndex][7] || data[rowIndex][9] || data[rowIndex][11] || data[rowIndex][14] || data[rowIndex][16] || data[rowIndex][18] || data[rowIndex][20] || data[rowIndex][23] || data[rowIndex][25] || data[rowIndex][27] || data[rowIndex][29] || data[rowIndex][32] || data[rowIndex][34] || data[rowIndex][36] || data[rowIndex][38]) + "</li>" +
        "<li><strong>Assignment Link:</strong> " + (data[rowIndex][6] || data[rowIndex][8] || data[rowIndex][10] || data[rowIndex][12] || data[rowIndex][15] || data[rowIndex][17] || data[rowIndex][19] || data[rowIndex][21] || data[rowIndex][24] || data[rowIndex][26] || data[rowIndex][28] || data[rowIndex][30] || data[rowIndex][33] || data[rowIndex][35] || data[rowIndex][37] || data[rowIndex][39] || data[rowIndex][41] || data[rowIndex][43]) + "</li>" +
        "</ul>";
      GmailApp.sendEmail("onlineteaching.159@gmail.com", subject, null, {
        htmlBody: htmlBodyContent,
        from: "info@online-tuitions.com"
      });
      GmailApp.sendEmail("hamzatariq017@gmail.com", subject, null, {
        htmlBody: htmlBodyContent,
        from: "info@online-tuitions.com"
      });
      GmailApp.sendEmail("danish5001danish@gmail.com", subject, null, {
        htmlBody: htmlBodyContent,
        from: "info@online-tuitions.com"
      });
    }
  }

  updateLastProcessedRow(data.length - 1);
}

// Helper function to retrieve the last processed row
function getLastProcessedRow() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastProcessedRow = scriptProperties.getProperty('lastProcessedRow');
  if (lastProcessedRow) {
    return parseInt(lastProcessedRow);
  } else {
    return 0;
  }
}

// Helper function to update the last processed row
function updateLastProcessedRow(row) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('lastProcessedRow', row.toString());
}

// Helper function to retrieve a folder by name within a given parent folder or create it if it doesn't exist
function getOrCreateFolder(folderName, parentFolder) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

// Helper function to extract the file ID from a Google Drive URL
function getFileIdFromURL(url) {
  var regex = /[-\w]{25,}/;
  var match = url.match(regex);
  if (match && match[0]) {
    return match[0];
  } else {
    throw new Error("File ID not found in the URL: " + url);
  }
}

// Helper function to extract the folder ID from a Google Drive link
function getFolderIdFromLink(link) {
  var regex = /[-\w]{25,}/;
  var folderId = link.match(regex)[0];
  return folderId;
}
