/**
 * When the spreadsheet is opened, add a custom menu item to open the sidebar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Shipment Tracker')
    .addItem('Show Tracker Sidebar', 'showTrackingSidebar')
    .addToUi();
}

/**
 * Opens the persistent sidebar.
 */
function showTrackingSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('TrackingSidebar')
      .setTitle('Shipment Tracker');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Main function to process tracking.
 *
 * Parameters:
 *  - trackingColumn: column where Tracking IDs reside (e.g., "A")
 *  - courierColumn: column where Courier names are (e.g., "B")
 *  - statusColumn: output column for Status (e.g., "C")
 *  - statusTimeColumn: output column for Status_time (e.g., "D")
 *  - executionTimeColumn: output column for Execution_time (e.g., "E")
 *
 * For each row (starting at row 2), if the courier is "DELHIVERY" and the Status cell is not "DELIVERED",
 * the script calls the API with the tracking ID (as the 'wbn' parameter) and required header,
 * then writes the delivery status, status timestamp, and current execution time into the specified output columns.
 * A 2-second delay is added after each API call.
 *
 * @param {string} trackingColumn
 * @param {string} courierColumn
 * @param {string} statusColumn
 * @param {string} statusTimeColumn
 * @param {string} executionTimeColumn
 * @return {string} A summary message including any error details.
 */
function processTracking(trackingColumn, courierColumn, statusColumn, statusTimeColumn, executionTimeColumn) {
  // Convert column letters to 1-indexed numbers.
  var trackingColIndex = columnToIndex(trackingColumn);
  var courierColIndex = columnToIndex(courierColumn);
  var statusColIndex = columnToIndex(statusColumn);
  var statusTimeColIndex = columnToIndex(statusTimeColumn);
  var executionTimeColIndex = columnToIndex(executionTimeColumn);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return "No data to process (only header row found).";
  }
  
  // Optionally update header row for the output columns if not already set.
  if (!sheet.getRange(1, statusColIndex).getValue()) {
    sheet.getRange(1, statusColIndex).setValue("Status");
  }
  if (!sheet.getRange(1, statusTimeColIndex).getValue()) {
    sheet.getRange(1, statusTimeColIndex).setValue("Status_time");
  }
  if (!sheet.getRange(1, executionTimeColIndex).getValue()) {
    sheet.getRange(1, executionTimeColIndex).setValue("Execution_time");
  }
  
  var errors = [];
  var updateCount = 0;
  var processedCount = 0;
  var currentTime = new Date();
  
  // Loop over each row starting from row 2 (assuming row 1 is the header)
  for (var i = 2; i <= lastRow; i++) {
    // Track whether an API call was attempted on this row.
    var apiCalled = false;
    
    try {
      // Read the courier name from the specified column.
      var courierName = sheet.getRange(i, courierColIndex).getValue();
      if (!courierName) continue; // Skip if courier name is empty.
      
      // Only process rows where courier is "DELHIVERY" (case-insensitive).
      if (courierName.toString().trim().toUpperCase() !== "DELHIVERY") continue;
      
      // Check if the Status cell already contains "DELIVERED". If so, skip this row.
      var currentStatus = sheet.getRange(i, statusColIndex).getValue();
      if (currentStatus && currentStatus.toString().trim().toUpperCase() === "DELIVERED") continue;
      
      // Get the tracking ID from the tracking column.
      var trackingId = sheet.getRange(i, trackingColIndex).getValue();
      if (!trackingId) {
        errors.push("Row " + i + ": Tracking ID is empty.");
        continue;
      }
      
      processedCount++;
      
      // Construct the API URL with the tracking ID as the 'wbn' parameter.
      var apiUrl = "https://dlv-api.delhivery.com/v3/unified-tracking?wbn=" + encodeURIComponent(trackingId);
      
      // Set the options with the required Origin header.
      var options = {
        "method": "get",
        "headers": {
          "Origin": "https://www.delhivery.com"
        },
        "muteHttpExceptions": true
      };
      
      // Make the API call.
      var response = UrlFetchApp.fetch(apiUrl, options);
      var responseCode = response.getResponseCode();
      
      // Mark that an API call was made.
      apiCalled = true;
      
      if (responseCode !== 200) {
        errors.push("Row " + i + ": API call failed with status code " + responseCode);
        sheet.getRange(i, statusColIndex).setValue("Error: " + responseCode);
        continue;
      }
      
      var jsonResponse = JSON.parse(response.getContentText());
      
      if (!jsonResponse.data || jsonResponse.data.length === 0) {
        errors.push("Row " + i + ": No data found in API response.");
        sheet.getRange(i, statusColIndex).setValue("No data");
        continue;
      }
      
      // Extract shipment details.
      var shipmentData = jsonResponse.data[0];
      var deliveryStatus = (shipmentData.status && shipmentData.status.status) ? shipmentData.status.status : "";
      var statusDateTime = (shipmentData.status && shipmentData.status.statusDateTime) ? shipmentData.status.statusDateTime : "";
      
      // Write the values into the output columns.
      sheet.getRange(i, statusColIndex).setValue(deliveryStatus);
      sheet.getRange(i, statusTimeColIndex).setValue(statusDateTime);
      sheet.getRange(i, executionTimeColIndex).setValue(currentTime);
      
      updateCount++;
      
    } catch (e) {
      errors.push("Row " + i + ": Exception - " + e.toString());
      sheet.getRange(i, statusColIndex).setValue("Error: " + e.toString());
    } finally {
      // If an API call was attempted on this row, wait for 2 seconds before continuing.
      if (apiCalled) {
        Utilities.sleep(2000); // 2000 milliseconds = 2 seconds
      }
    }
  }
  
  var summary = "Processed " + (lastRow - 1) + " rows.\nAPI called on " + processedCount + " rows.\nUpdated " + updateCount + " rows.";
  if (errors.length > 0) {
    summary += "\n\nErrors:\n" + errors.join("\n");
  }
  
  return summary;
}

/**
 * Helper function: converts a column letter (or letters) to its corresponding 1-indexed number.
 * For example, "A" → 1, "B" → 2, …, "AA" → 27, etc.
 *
 * @param {string} column - The column letter(s).
 * @return {number} The 1-indexed column number.
 */
function columnToIndex(column) {
  var letters = column.toUpperCase().trim();
  var sum = 0;
  for (var i = 0; i < letters.length; i++) {
    sum *= 26;
    sum += letters.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return sum;
}
