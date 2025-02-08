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
 *   trackingColumn: column where Tracking IDs reside (e.g., "A")
 *   courierColumn: column where Courier names are (e.g., "B")
 *   statusColumn: output column for Status (e.g., "C")
 *   statusTimeColumn: output column for Status_time (e.g., "D")
 *   executionTimeColumn: output column for Execution_time (e.g., "E")
 *
 * For each row (starting at row 2), if the courier is either "DELHIVERY" or "DTDC" and the Status cell
 * is not already "DELIVERED", the appropriate API is called to fetch shipment details.
 *
 * A 2-second delay is applied only after an API call is made.
 *
 * The code also checks a script property ("trackingExecutionState") for pause/resume/stop commands.
 */
function processTracking(trackingColumn, courierColumn, statusColumn, statusTimeColumn, executionTimeColumn) {
  // Set initial execution state to "running".
  PropertiesService.getScriptProperties().setProperty("trackingExecutionState", "running");

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
  
  // Update header row for the output columns if not already set.
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
  
  // Process each row starting from row 2.
  for (var i = 2; i <= lastRow; i++) {
    
    // Check the execution state before processing each row.
    var execState = PropertiesService.getScriptProperties().getProperty("trackingExecutionState");
    if (execState === "stopped") {
      errors.push("Execution stopped by user at row " + i);
      break;
    }
    // If paused, wait until the state becomes "running" or "stopped".
    while (PropertiesService.getScriptProperties().getProperty("trackingExecutionState") === "paused") {
      Utilities.sleep(1000); // Wait 1 second.
      if (PropertiesService.getScriptProperties().getProperty("trackingExecutionState") === "stopped") {
        errors.push("Execution stopped by user during pause at row " + i);
        break;
      }
    }
    
    var apiCalled = false;
    try {
      // Read the courier name.
      var courierName = sheet.getRange(i, courierColIndex).getValue();
      if (!courierName) continue; // Skip if empty.
      courierName = courierName.toString().trim().toUpperCase();
      
      // Only process rows where courier is either DELHIVERY or DTDC.
      if (courierName !== "DELHIVERY" && courierName !== "DTDC") continue;
      
      // Skip row if the Status cell already contains "DELIVERED" (case insensitive).
      var currentStatus = sheet.getRange(i, statusColIndex).getValue();
      if (currentStatus && currentStatus.toString().trim().toUpperCase() === "DELIVERED") continue;
      
      // Get the tracking ID.
      var trackingId = sheet.getRange(i, trackingColIndex).getValue();
      if (!trackingId) {
        errors.push("Row " + i + ": Tracking ID is empty.");
        continue;
      }
      
      processedCount++;
      var deliveryStatus = "";
      var statusDateTime = "";
      
      if (courierName === "DELHIVERY") {
        // -----------------------
        // DELHIVERY API CALL (GET)
        // -----------------------
        var apiUrl = "https://dlv-api.delhivery.com/v3/unified-tracking?wbn=" + encodeURIComponent(trackingId);
        var options = {
          "method": "get",
          "headers": {
            "Origin": "https://www.delhivery.com"
          },
          "muteHttpExceptions": true
        };
        var response = UrlFetchApp.fetch(apiUrl, options);
        var responseCode = response.getResponseCode();
        apiCalled = true;
        
        if (responseCode !== 200) {
          errors.push("Row " + i + ": Delhivery API call failed with status code " + responseCode);
          sheet.getRange(i, statusColIndex).setValue("Error: " + responseCode);
          continue;
        }
        
        var jsonResponse = JSON.parse(response.getContentText());
        if (!jsonResponse.data || jsonResponse.data.length === 0) {
          errors.push("Row " + i + ": No data found in Delhivery API response.");
          sheet.getRange(i, statusColIndex).setValue("No data");
          continue;
        }
        
        var shipmentData = jsonResponse.data[0];
        deliveryStatus = shipmentData.status && shipmentData.status.status ? shipmentData.status.status : "";
        statusDateTime = shipmentData.status && shipmentData.status.statusDateTime ? shipmentData.status.statusDateTime : "";
        
      } else if (courierName === "DTDC") {
        // -----------------------
        // DTDC API CALL (POST)
        // -----------------------
        var apiUrl = "https://trackcourier.io/api/v1/get_checkpoints_table/5874d65fdac8775f005f43355e368693/dtdc/" + encodeURIComponent(trackingId);
        var options = {
          "method": "post",
          "muteHttpExceptions": true
        };
        var response = UrlFetchApp.fetch(apiUrl, options);
        var responseCode = response.getResponseCode();
        apiCalled = true;
        
        if (responseCode !== 200) {
          errors.push("Row " + i + ": DTDC API call failed with status code " + responseCode);
          sheet.getRange(i, statusColIndex).setValue("Error: " + responseCode);
          continue;
        }
        
        var jsonResponse = JSON.parse(response.getContentText());
        if (jsonResponse.Result !== "success") {
          errors.push("Row " + i + ": DTDC API response not successful");
          sheet.getRange(i, statusColIndex).setValue("Error: API not successful");
          continue;
        }
        
        // Use MostRecentStatus as the delivery status.
        deliveryStatus = jsonResponse.MostRecentStatus || "";
        
        // Merge Date and Time from the first entry in the Checkpoints array.
        if (jsonResponse.Checkpoints && jsonResponse.Checkpoints.length > 0) {
          var firstCheckpoint = jsonResponse.Checkpoints[0];
          statusDateTime = firstCheckpoint.Date + " " + firstCheckpoint.Time;
        } else {
          errors.push("Row " + i + ": DTDC API returned no checkpoints.");
          sheet.getRange(i, statusColIndex).setValue("Error: No checkpoints");
          continue;
        }
      }
      
      // Write the fetched values into the output columns.
      sheet.getRange(i, statusColIndex).setValue(deliveryStatus);
      sheet.getRange(i, statusTimeColIndex).setValue(statusDateTime);
      sheet.getRange(i, executionTimeColIndex).setValue(new Date());
      updateCount++;
      
    } catch (e) {
      errors.push("Row " + i + ": Exception - " + e.toString());
      sheet.getRange(i, statusColIndex).setValue("Error: " + e.toString());
    } finally {
      // Apply 2-second delay only if an API call was made.
      if (apiCalled) {
        Utilities.sleep(2000);
      }
    }
  }
  
  // Clear the execution state at the end.
  PropertiesService.getScriptProperties().deleteProperty("trackingExecutionState");
  
  var summary = "Processed " + (lastRow - 1) + " rows.\nAPI called on " + processedCount + " rows.\nUpdated " + updateCount + " rows.";
  if (errors.length > 0) {
    summary += "\n\nErrors:\n" + errors.join("\n");
  }
  return summary;
}

/**
 * Toggle the pause/resume state.
 * If currently running, sets state to "paused".
 * If currently paused, sets state to "running".
 * If execution is already stopped, returns a message.
 *
 * @return {string} A message indicating the new state.
 */
function pauseResumeExecution() {
  var prop = PropertiesService.getScriptProperties();
  var currentState = prop.getProperty("trackingExecutionState");
  if (currentState === "running") {
    prop.setProperty("trackingExecutionState", "paused");
    return "Paused";
  } else if (currentState === "paused") {
    prop.setProperty("trackingExecutionState", "running");
    return "Resumed";
  } else {
    return "Cannot toggle: Execution has been stopped.";
  }
}

/**
 * Set the execution state to "stopped".
 *
 * @return {string} A message indicating execution has been stopped.
 */
function stopExecution() {
  PropertiesService.getScriptProperties().setProperty("trackingExecutionState", "stopped");
  return "Execution stopped.";
}

/**
 * Helper function that converts a column letter (or letters) to its corresponding 1-indexed number.
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
