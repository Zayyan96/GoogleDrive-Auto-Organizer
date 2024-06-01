# Google Drive File Organizer

This Google Apps Script automates the process of moving new files to specific folders in Google Drive based on data from a Google Sheet.

## Features

- Retrieves file information from a Google Sheet
- Moves files to organized folders in Google Drive
- Renames files if a file with the same name already exists in the target folder
- Logs file movements in a separate sheet within the Google Sheet
- Sends email notifications in case of errors

## Setup

1. Open your Google Sheet and go to `Extensions > Apps Script`.
2. Replace the content with the provided script code.
3. Update the script with your Google Sheet ID and desired folder IDs.
4. Save and run the script.

## How to Use

1. Populate the Google Sheet with the necessary file information.
2. Run the script manually or set up a trigger to run it periodically.
3. Check the "MovedFiles" sheet for logs of moved files.

## Configuration

- Update the `sheetId` variable with your Google Sheet ID.
- Update the `grandfatherFolderLink` with your Google Drive folder link.

## Functions

### moveNewFilesToFolder()

Main function to move new files to the designated folders.

### getLastProcessedRow()

Helper function to retrieve the last processed row from script properties.

### updateLastProcessedRow(row)

Helper function to update the last processed row in script properties.

### getOrCreateFolder(folderName, parentFolder)

Helper function to retrieve or create a folder by name within a given parent folder.

### getFileIdFromURL(url)

Helper function to extract the file ID from a Google Drive URL.

### getFolderIdFromLink(link)

Helper function to extract the folder ID from a Google Drive link.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
