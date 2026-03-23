// -----------------------------------------------------------------------------
// UPDATED APP SCRIPT CODE V2 (WITH SERIAL NUMBER SUPPORT)
// -----------------------------------------------------------------------------
// INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Replace your ENTIRE existing code in Code.gs with the code below.
// 4. Click 'Deploy' > 'New deployment'
// 5. Ensure "Execute as" is set to "Me" and "Who has access" is "Anyone".
// 6. Click 'Deploy' and copy the new URL if it changes.
// -----------------------------------------------------------------------------

const SPREADSHEET_ID = "1M1yQCmHlROwpZE2ZBv_dvSbCyEB5dyTRNNDAcZFyaa0";
const SCRIPT_VERSION = "v3.0.0"; // Increment this to verify deployment

// Helper function to create JSON response
function createJsonResponse(data, statusCode = 200) {
    // Add version to every response for debugging
    if (typeof data === 'object' && data !== null) {
        data._version = SCRIPT_VERSION;
    }
    const output = ContentService.createTextOutput(JSON.stringify(data));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
}

// Function to log to "Subscription Approval" sheet
function logSubscriptionApproval(approvalData) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName("Subscription Approval");

        // Create the sheet if it doesn't exist
        if (!sheet) {
            sheet = ss.insertSheet("Subscription Approval");

            // Set up headers
            const headers = [
                "Timestamp",
                "Approval No",
                "Subscription No",
                "Approved By",
                "Approval Status",
                "Note"
            ];

            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

            // Format header row
            const headerRange = sheet.getRange(1, 1, 1, headers.length);
            headerRange.setBackground("#4f46e5");
            headerRange.setFontColor("white");
            headerRange.setFontWeight("bold");

            // Auto-resize columns
            for (let i = 1; i <= headers.length; i++) {
                sheet.autoResizeColumn(i);
            }
        }

        // Prepare row data
        const timestamp = new Date();
        const rowData = [
            timestamp,
            approvalData.approvalNo || "",
            approvalData.subscriptionNo || "",
            approvalData.approvedBy || "",
            approvalData.approvalStatus || "",
            approvalData.note || ""
        ];

        // Append the row
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);

        // Format the timestamp column
        const timestampCell = sheet.getRange(lastRow + 1, 1);
        timestampCell.setNumberFormat("yyyy-mm-dd hh:mm:ss");

        console.log("Logged subscription approval:", approvalData);
        return true;

    } catch (error) {
        console.error("Error logging subscription approval:", error);
        return false;
    }
}

// Function to log sharing activity to "Shared Documents" sheet
function logSharingActivity(shareData) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName("Shared Documents");

        // Create the sheet if it doesn't exist
        if (!sheet) {
            sheet = ss.insertSheet("Shared Documents");

            // Set up headers
            const headers = [
                "Timestamp",
                "Email",
                "Name",
                "Document Name",
                "Document Type",
                "Category",
                "Serial No",
                "Image",
                "Source Sheet",
                "Share Method",
                "Number"
            ];

            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

            // Format header row
            const headerRange = sheet.getRange(1, 1, 1, headers.length);
            headerRange.setBackground("#4f46e5");
            headerRange.setFontColor("white");
            headerRange.setFontWeight("bold");

            // Auto-resize columns
            for (let i = 1; i <= headers.length; i++) {
                sheet.autoResizeColumn(i);
            }
        }

        // Prepare row data
        const timestamp = new Date();
        const rowData = [
            timestamp,
            shareData.email || "",
            shareData.recipientName || "",
            shareData.documentName || "",
            shareData.documentType || "",
            shareData.category || "",
            shareData.serialNo || "",
            shareData.image || shareData.fileContent || "",
            shareData.sourceSheet || "Documents",
            shareData.shareMethod || "Email",
            shareData.number || ""
        ];

        // Append the row
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);

        // Format the timestamp column
        const timestampCell = sheet.getRange(lastRow + 1, 1);
        timestampCell.setNumberFormat("yyyy-mm-dd hh:mm:ss");

        // Auto-resize columns again
        for (let i = 1; i <= headers.length; i++) {
            sheet.autoResizeColumn(i);
        }

        console.log("Logged sharing activity:", shareData);
        return true;

    } catch (error) {
        console.error("Error logging sharing activity:", error);
        return false;
    }
}

// Helper function to remove duplicates and empty rows
// MODIFIED: Now appends the ORIGINAL row index to the end of each row
function removeDuplicatesAndEmptyRows(data) {
    if (data.length === 0) return [];

    const headers = [...data[0], "originalRowIndex"]; // Add header for debug/clarity
    const rows = data.slice(1); // Rows starting from index 1 (Sheet Row 2)

    const seen = new Set();
    const uniqueRows = [];

    rows.forEach((row, index) => {
        // Calculate 1-based sheet row index.
        // data[0] is Row 1. data.slice(1) starts at Row 2.
        // So current row's sheet index = index (in sliced array) + 2.
        const originalRowIndex = index + 2;

        // Skip completely empty rows
        if (!row || row.length === 0 || !row.some(cell => cell && cell.toString().trim() !== '')) {
            // console.log(`Skipping empty row at index ${originalRowIndex}`);
            return;
        }

        // Create a unique key using SN (col B), Document Name (col C), and Company (col F)
        const sn = row[1] || '';
        const docName = row[2] || '';
        const company = row[5] || '';

        // For Loan sheet specifically, we might want weaker deduplication or rely on SN
        // But keeping original logic for safety, just appending index.
        const key = `${sn}_${docName}_${company}`.toLowerCase().trim();

        if (!seen.has(key)) {
            seen.add(key);
            // PUSH THE ROW WITH ITS ORIGINAL INDEX APPENDED
            uniqueRows.push([...row, originalRowIndex]);
        } else {
            console.log(`Removing duplicate at row ${originalRowIndex}: ${key}`);
        }
    });

    // Return headers + unique rows
    return [headers, ...uniqueRows];
}




// Extract file ID from Google Drive URL
function extractFileIdFromUrl(url) {
    var fileId = null;
    if (url && url.includes('drive.google.com')) {
        // Format: https://drive.google.com/file/d/FILE_ID/view
        var viewMatch = url.match(/\/file\/d\/([^/]+)/);
        if (viewMatch) {
            fileId = viewMatch[1];
        } else {
            // Format: https://drive.google.com/open?id=FILE_ID
            var openMatch = url.match(/[?&]id=([^&]+)/);
            if (openMatch) {
                fileId = openMatch[1];
            }
        }
    }
    return fileId;
}

// Get file information
function getFileInfo(fileId) {
    try {
        var file = DriveApp.getFileById(fileId);
        return {
            id: fileId,
            name: file.getName(),
            url: file.getUrl(),
            downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId,
            size: file.getSize(),
            mimeType: file.getMimeType()
        };
    } catch (e) {
        console.error("Error in getFileInfo:", e);
        return null;
    }
}

// Upload file to Google Drive
function uploadFileToDrive(base64Data, fileName, mimeType, folderId) {
    try {
        console.log('uploadFileToDrive called with:', {
            fileName: fileName,
            mimeType: mimeType,
            folderId: folderId,
            base64DataLength: base64Data ? base64Data.length : 0
        });

        let fileData = base64Data;
        if (base64Data.indexOf('base64,') !== -1) {
            fileData = base64Data.split('base64,')[1];
            console.log('Removed data URL prefix');
        }

        const decoded = Utilities.base64Decode(fileData);
        console.log('Base64 decoded, length:', decoded.length);

        const blob = Utilities.newBlob(decoded, mimeType, fileName);
        console.log('Blob created');

        const folder = DriveApp.getFolderById(folderId);
        console.log('Folder retrieved:', folder.getName());

       const file = folder.createFile(blob);

file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

// ✅ use view link instead of uc?export=view
const fileUrl = "https://drive.google.com/open?id=" + file.getId();

return fileUrl;

    } catch (error) {
        console.error("Error in uploadFileToDrive:", error);
        console.error("Error details:", error.toString());
        return null;
    }
}

// Handle file upload
function handleFileUpload(e) {
    try {
        console.log('handleFileUpload called');
        var params = e.parameter;

        console.log('File upload params:', {
            hasBase64Data: !!params.base64Data,
            fileName: params.fileName,
            mimeType: params.mimeType,
            folderId: params.folderId
        });

        if (!params.base64Data || !params.fileName || !params.mimeType || !params.folderId) {
            throw new Error("Missing required parameters for file upload. Required: base64Data, fileName, mimeType, folderId");
        }

        var fileUrl = uploadFileToDrive(params.base64Data, params.fileName, params.mimeType, params.folderId);
        if (!fileUrl) {
            throw new Error("Failed to upload file to Google Drive");
        }

        console.log('File uploaded successfully:', fileUrl);
        return createJsonResponse({
            success: true,
            fileUrl: fileUrl,
            message: "File uploaded successfully"
        });

    } catch (error) {
        console.error("Error in handleFileUpload:", error);
        return createJsonResponse({
            success: false,
            error: error.toString()
        });
    }
}

function sendBatchEmail(to, subject, body, isHtml, attachmentUrls, shareData) {
    try {

        var emailOptions = {
            to: to,
            subject: subject
        };

        // Only body send karega, attachment nahi
        if (isHtml) {
            emailOptions.htmlBody = body;
        } else {
            emailOptions.body = body;
        }

        // ❌ attachment logic completely removed

        // Send email
        MailApp.sendEmail(emailOptions);

        // Log sharing activity
        if (shareData) {
            logSharingActivity(shareData);
        }

        return true;

    } catch (error) {
        console.error("Error sending email:", error);
        throw error;
    }
}

// Helper function to update specific cells in a row (by index)
function updateSpecificCells(sheet, rowIndex, updates) {
    try {
        console.log('Updating specific cells at row', rowIndex, 'with:', updates);

        // updates should be an array of objects: {column: number, value: any}
        for (var i = 0; i < updates.length; i++) {
            var update = updates[i];
            var column = update.column;
            var value = update.value;

            if (column < 1) {
                throw new Error("Invalid column index: " + column);
            }

            // Update the specific cell
            sheet.getRange(rowIndex, column).setValue(value);
            console.log('Updated cell at row', rowIndex, 'column', column, 'with:', value);
        }

        return true;
    } catch (error) {
        console.error("Error updating specific cells:", error);
        throw error;
    }
}

// NEW FUNCTION: Update specific cells by Serial No (Column B)
function updateSpecificCellsBySn(sheet, sn, updates) {
    try {
        console.log('Searching for row with SN:', sn);

        // Get all data from Column B (Serial No)
        // We get range from B2 to B[LastRow]
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) throw new Error("Sheet is empty");

        // Get values from Column B (index 2)
        const snValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

        let targetRowIndex = -1;

        // Find the row index (adding 2 because we started at row 2)
        for (let i = 0; i < snValues.length; i++) {
            if (snValues[i][0] && snValues[i][0].toString().trim() === sn.trim()) {
                targetRowIndex = i + 2;
                break;
            }
        }

        if (targetRowIndex === -1) {
            throw new Error(`Serial No '${sn}' not found in sheet`);
        }

        console.log(`Found SN '${sn}' at row ${targetRowIndex}`);

        // Reuse existing update logic
        return updateSpecificCells(sheet, targetRowIndex, updates);

    } catch (error) {
        console.error("Error updating specific cells by SN:", error);
        throw error;
    }
}


// Main doGet function
function doGet(e) {
    try {
        // Get parameters or use defaults
        const params = e ? e.parameter : {};
        const sheetName = params.sheet || "Documents";

        console.log(`Request for sheet: ${sheetName}`);

        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            return createJsonResponse({
                success: false,
                error: `Sheet '${sheetName}' not found`
            });
        }

        // Get all data
        const data = sheet.getDataRange().getValues();
        console.log(`Fetched ${data.length} rows from ${sheetName}`);

        // For "Shared Documents" sheet, only remove empty rows but keep ALL records (no deduplication)
        // For other sheets, apply deduplication as before
        let filteredData;
        if (sheetName === "Shared Documents") {
            // Only remove empty rows, keep all share records
            filteredData = data.filter((row, index) => {
                // Keep header row (index 0)
                if (index === 0) return true;
                // Remove completely empty rows
                return row && row.length > 0 && row.some(cell => cell && cell.toString().trim() !== '');
            });
        } else {
            // Apply deduplication for other sheets
            filteredData = removeDuplicatesAndEmptyRows(data);
        }

        const result = {
            success: true,
            sheet: sheetName,
            totalRows: data.length,
            filteredRows: filteredData.length,
            removedDuplicates: sheetName === "Shared Documents" ? 0 : data.length - filteredData.length,
            data: filteredData,
            timestamp: new Date().toISOString()
        };

        // Return the response
        return createJsonResponse(result);

    } catch (err) {
        console.error("Error in doGet:", err);

        return createJsonResponse({
            success: false,
            error: err.toString(),
            timestamp: new Date().toISOString()
        });
    }
}

// For preflight requests
function doOptions() {
    return createJsonResponse({});
}


// Helper to generate next Serial Number safely
function getNextSerialNumber(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "SN-001";

    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let maxSn = 0;

    for (let i = 0; i < values.length; i++) {
        const sn = values[i][0];
        if (sn && typeof sn === 'string' && sn.trim().startsWith("SN-")) {
            const part = sn.trim().replace("SN-", "");
            const numPart = parseInt(part, 10);
            if (!isNaN(numPart) && numPart > maxSn) {
                maxSn = numPart;
            }
        }
    }

    const nextVal = maxSn + 1;
    return "SN-" + nextVal.toString().padStart(3, "0");
}

// Helper to generate next Subscription Number safely (SUB-XXX)
function getNextSubscriptionNumber(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "SUB-001";

    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let maxSn = 0;

    for (let i = 0; i < values.length; i++) {
        const sn = values[i][0];
        if (sn && typeof sn === 'string') {
            const trimmedSn = sn.trim().toUpperCase();
            if (trimmedSn.startsWith("SUB-")) {
                const part = trimmedSn.replace("SUB-", "");
                const numPart = parseInt(part, 10);
                if (!isNaN(numPart) && numPart > maxSn) {
                    maxSn = numPart;
                }
            }
        }
    }

    const nextVal = maxSn + 1;
    return "SUB-" + nextVal.toString().padStart(3, "0");
}

// Helper to generate next Loan Number safely (SN-XXX) for Loan sheet
// Assuming Loan sheet uses "SN-" prefix as per user request
function getNextLoanNumber(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "SN-001";

    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let maxSn = 0;

    for (let i = 0; i < values.length; i++) {
        const sn = values[i][0];
        if (sn && typeof sn === 'string') {
            const trimmedSn = sn.trim().toUpperCase();
            if (trimmedSn.startsWith("SN-")) {
                const part = trimmedSn.replace("SN-", "");
                const numPart = parseInt(part, 10);
                if (!isNaN(numPart) && numPart > maxSn) {
                    maxSn = numPart;
                }
            }
        }
    }

    const nextVal = maxSn + 1;
    return "SN-" + nextVal.toString().padStart(3, "0");
}

// Main doPost function
function doPost(e) {
    try {
        var params = e.parameter;
        var action = params.action || 'insert';
        console.log('Action received:', action);

        // Handle different actions
        if (action === 'uploadFile') {
            return handleFileUpload(e);
        }

        if (action === 'getFileInfo') {
            var fileId = params.fileId;

            if (!fileId) {
                throw new Error("fileId parameter is required");
            }

            var fileInfo = getFileInfo(fileId);

            if (!fileInfo) {
                throw new Error("File not found or access denied");
            }

            return createJsonResponse({
                success: true,
                fileInfo: fileInfo
            });
        }

        if (action === 'sendEmail') {
            // Extract email parameters
            var to = params.to;
            var subject = params.subject;
            var body = params.body;
            var isHtml = params.isHtml === 'true';

            // Attachments are disabled as per user request (sharing via links only)
            // var attachmentArray = [];

            // Extract sharing data for logging
            var shareData = {
                email: to,
                recipientName: params.recipientName || "",
                documentName: params.documentName || subject || "Document",
                documentType: params.documentType || "",
                category: params.category || "",
                serialNo: params.serialNo || "",
                fileContent: attachmentArray[0] || "",
                sourceSheet: "Documents",
                shareMethod: "Email",
                number: ""
            };

            if (!to || !subject || !body) {
                throw new Error("Missing required email parameters: to, subject, body");
            }

            // Send the email with multiple attachments
            sendBatchEmail(to, subject, body, isHtml, attachmentArray, shareData);

            return createJsonResponse({
                success: true,
                message: "Email sent successfully",
                // attachmentsCount: attachmentArray.length
            });
        }

        // Action to log sharing activity without sending email (for WhatsApp)
        if (action === 'logSharing') {
            var shareData = {
                email: params.email || "",
                recipientName: params.recipientName || "",
                documentName: params.documentName || "",
                documentType: params.documentType || "",
                category: params.category || "",
                serialNo: params.serialNo || "",
                fileContent: params.fileContent || "",
                sourceSheet: "Documents",
                shareMethod: params.shareMethod || "WhatsApp",
                number: params.number || ""
            };

            // Log the sharing activity
            var logged = logSharingActivity(shareData);

            if (logged) {
                return createJsonResponse({
                    success: true,
                    message: "Sharing activity logged successfully"
                });
            } else {
                throw new Error("Failed to log sharing activity");
            }
        }

        // Handle sheet operations
        var sheetName = params.sheetName || "Documents";
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            throw new Error(`Sheet '${sheetName}' not found`);
        }

        // Action: Update specific cells by row index
        if (action === 'updateCells') {
            var rowIndex = parseInt(params.rowIndex);
            var cellUpdates = JSON.parse(params.cellUpdates); // Array of {column: number, value: any}

            if (isNaN(rowIndex) || rowIndex < 2) {
                throw new Error("Invalid row index for update");
            }

            if (!Array.isArray(cellUpdates)) {
                throw new Error("cellUpdates must be an array");
            }

            updateSpecificCells(sheet, rowIndex, cellUpdates);

            return createJsonResponse({
                success: true,
                message: "Cells updated successfully"
            });
        }

        // NEW ACTION: Update cells by Serial No (Column B)
        if (action === 'updateCellsBySn') {
            var sn = params.sn;
            var cellUpdates = JSON.parse(params.cellUpdates);

            if (!sn) {
                throw new Error("Serial No (sn) is required");
            }

            if (!Array.isArray(cellUpdates)) {
                throw new Error("cellUpdates must be an array");
            }

            updateSpecificCellsBySn(sheet, sn, cellUpdates);

            return createJsonResponse({
                success: true,
                message: "Cells updated successfully by SN"
            });
        }

        if (action === 'insert') {
            // Use LockService to prevent race conditions
            var lock = LockService.getScriptLock();
            // Wait for up to 30 seconds for other processes to finish.
            lock.waitLock(30000);

            try {
                var rowData = JSON.parse(params.rowData);

                // Only generate SN for "Documents" sheet (Column B check)
                if (sheetName === "Documents") {
                    var newSN = getNextSerialNumber(sheet);
                    // Overwrite the SN in index 1 (Column B)
                    rowData[1] = newSN;
                }
                // Handle "Subscription" sheet SN generation (SUB-XXX)
                else if (sheetName === "Subscription") {
                    var newSN = getNextSubscriptionNumber(sheet);
                    rowData[1] = newSN;
                }
                // Handle "Loan" sheet SN generation (SN-XXX)
                // Case-insensitive check to be safe
                else if (sheetName.trim().toLowerCase() === "loan") {
                    console.log("Generating SN for Loan sheet");
                    var newSN = getNextLoanNumber(sheet);
                    console.log("Generated Loan SN:", newSN);
                    rowData[1] = newSN;
                }

                const lastRow = sheet.getLastRow();
                sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);

                return createJsonResponse({
                    success: true,
                    message: "Data inserted successfully",
                    // Return the generated SN so the client knows what happened
                    serialNo: (sheetName === "Documents" || sheetName === "Subscription" || sheetName.trim().toLowerCase() === "loan") ? rowData[1] : undefined,
                    _version: SCRIPT_VERSION
                });
            } finally {
                // Must release the lock
                lock.releaseLock();
            }
        }
        else if (action === 'update') {
            // Update an existing row
            var rowIndex = parseInt(params.rowIndex);
            var rowData = JSON.parse(params.rowData);
            if (isNaN(rowIndex) || rowIndex < 2) {
                throw new Error("Invalid row index for update");
            }
            for (var i = 0; i < rowData.length; i++) {
                // Skip empty cells to preserve original data if not changed
                if (rowData[i] !== '') {
                    sheet.getRange(rowIndex, i + 1).setValue(rowData[i]);
                }
            }
            return createJsonResponse({
                success: true,
                message: "Data updated successfully"
            });
        }
        else if (action === 'updateCell') {
            var rowIndex = parseInt(params.rowIndex);
            var columnIndex = parseInt(params.columnIndex);
            var value = params.value;
            if (isNaN(rowIndex) || rowIndex < 1 || isNaN(columnIndex) || columnIndex < 1) {
                throw new Error("Invalid row or column index for update");
            }
            sheet.getRange(rowIndex, columnIndex).setValue(value);
            return createJsonResponse({
                success: true,
                message: "Cell updated successfully"
            });
        }
        else if (action === 'delete') {
            var rowIndex = parseInt(params.rowIndex);

            if (isNaN(rowIndex) || rowIndex < 2) {
                throw new Error("Invalid row index for delete");
            }
            sheet.deleteRow(rowIndex);
            return createJsonResponse({
                success: true,
                message: "Row deleted successfully"
            });
        }
        else if (action === 'markDeleted') {
            var rowIndex = parseInt(params.rowIndex);
            var columnIndex = parseInt(params.columnIndex);
            var value = params.value || 'Yes';

            if (isNaN(rowIndex) || rowIndex < 2) {
                throw new Error("Invalid row index for marking as deleted");
            }
            if (isNaN(columnIndex) || columnIndex < 1) {
                throw new Error("Invalid column index for marking as deleted");
            }
            sheet.getRange(rowIndex, columnIndex).setValue(value);

            return createJsonResponse({
                success: true,
                message: "Row marked as deleted successfully"
            });
        }
        else {
            throw new Error("Unknown action: " + action);
        }
    } catch (error) {
        console.error("Error in doPost:", error);
        return createJsonResponse({
            success: false,
            error: error.toString()
        });
    }
}
