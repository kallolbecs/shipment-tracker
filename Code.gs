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
 * is not already "DELIVERED" (case insensitive), the appropriate API is called to fetch shipment details.
 *
 * For DELHIVERY rows, a GET call is made; for DTDC rows, the code attempts the primary DTDC API call
 * and retries up to 4 attempts in total.
 *
 * A 2-second delay is applied only after a DELHIVERY API call.
 *
 * The code also checks a script property ("trackingExecutionState") for pause/resume/stop commands.
 * Additionally, it updates script properties ("currentRow", "currentApi", and "currentResponseCode")
 * with the row number, API short name, and response code currently being used (only for rows where an API call is made)
 * so that the sidebar UI can display live progress.
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
  
  // Ensure header cells exist.
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
    // Check the execution state before processing this row.
    var execState = PropertiesService.getScriptProperties().getProperty("trackingExecutionState");
    if (execState === "stopped") {
      errors.push("Execution stopped by user at row " + i);
      break;
    }
    while (PropertiesService.getScriptProperties().getProperty("trackingExecutionState") === "paused") {
      Utilities.sleep(1000); // Wait 1 second.
      if (PropertiesService.getScriptProperties().getProperty("trackingExecutionState") === "stopped") {
        errors.push("Execution stopped by user during pause at row " + i);
        break;
      }
    }
    
    var apiCalled = false;
    var currentCourier = "";
    try {
      // Read the courier name.
      var courierName = sheet.getRange(i, courierColIndex).getValue();
      if (!courierName) {
        // Clear current row if skipping.
        PropertiesService.getScriptProperties().setProperty("currentRow", "");
        continue;
      }
      courierName = courierName.toString().trim().toUpperCase();
      currentCourier = courierName;
      
      // Only process rows for DELHIVERY or DTDC.
      if (courierName !== "DELHIVERY" && courierName !== "DTDC") {
        PropertiesService.getScriptProperties().setProperty("currentRow", "");
        continue;
      }
      
      // Skip row if the Status cell already contains "DELIVERED" (case insensitive).
      var currentStatus = sheet.getRange(i, statusColIndex).getValue();
      if (currentStatus && currentStatus.toString().trim().toUpperCase() === "DELIVERED") {
        PropertiesService.getScriptProperties().setProperty("currentRow", "");
        continue;
      }
      
      // Get the tracking ID.
      var trackingId = sheet.getRange(i, trackingColIndex).getValue();
      if (!trackingId) {
        errors.push("Row " + i + ": Tracking ID is empty.");
        PropertiesService.getScriptProperties().setProperty("currentRow", "");
        continue;
      }
      
      processedCount++;
      var deliveryStatus = "";
      var statusDateTime = "";
      
      // Update current row property.
      PropertiesService.getScriptProperties().setProperty("currentRow", i.toString());
      
      if (courierName === "DELHIVERY") {
        // -----------------------
        // DELHIVERY API CALL (GET)
        // -----------------------
        PropertiesService.getScriptProperties().setProperty("currentApi", "DEL");
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
        // Update response code property.
        PropertiesService.getScriptProperties().setProperty("currentResponseCode", responseCode.toString());
        apiCalled = true;
        // Check if response code is in 20X range.
        if (!(responseCode >= 200 && responseCode < 300)) {
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
        deliveryStatus = (shipmentData.status && shipmentData.status.status) ? shipmentData.status.status : "";
        statusDateTime = (shipmentData.status && shipmentData.status.statusDateTime) ? shipmentData.status.statusDateTime : "";
      
      } else if (courierName === "DTDC") {
        // -----------------------
        // DTDC API CALLS WITH RETRY (MAX 4 ATTEMPTS)
        // -----------------------
        var success = false;
        var attempts = 0;
        while (!success && attempts < 4) {
          attempts++;
          // Check execution state inside the retry loop.
          var state = PropertiesService.getScriptProperties().getProperty("trackingExecutionState");
          if (state === "stopped") {
            errors.push("Row " + i + ": Execution stopped during DTDC retry.");
            break;
          }
          while (PropertiesService.getScriptProperties().getProperty("trackingExecutionState") === "paused") {
            Utilities.sleep(1000);
            if (PropertiesService.getScriptProperties().getProperty("trackingExecutionState") === "stopped") {
              errors.push("Row " + i + ": Execution stopped during pause in DTDC retry.");
              break;
            }
          }
          
          try {
            PropertiesService.getScriptProperties().setProperty("currentApi", "DTDC-P");
            var primaryUrl = "https://trackcourier.io/api/v1/get_checkpoints_table/5874d65fdac8775f005f43355e368693/dtdc/" + encodeURIComponent(trackingId);
            var primaryOptions = {
              "method": "post",
              "muteHttpExceptions": true
            };
            var primaryResponse = UrlFetchApp.fetch(primaryUrl, primaryOptions);
            var primaryCode = primaryResponse.getResponseCode();
            PropertiesService.getScriptProperties().setProperty("currentResponseCode", primaryCode.toString());
            if (primaryCode >= 200 && primaryCode < 300) {
              var primaryJson = JSON.parse(primaryResponse.getContentText());
              if (primaryJson.Result === "success" && primaryJson.Checkpoints && primaryJson.Checkpoints.length > 0) {
                deliveryStatus = primaryJson.MostRecentStatus || "";
                var firstCheckpoint = primaryJson.Checkpoints[0];
                statusDateTime = firstCheckpoint.Date + " " + firstCheckpoint.Time;
                success = true;
                break;
              }
            }
          } catch (e) {
            // Ignore exception.
          }
          
          if (!success && attempts < 4) {
            errors.push("Row " + i + ": DTDC primary API failed on attempt " + attempts + ", retrying...");
            Utilities.sleep(2000); // Wait 2 seconds before retrying.
          }
        } // End of DTDC retry loop.
        if (!success) {
          errors.push("Row " + i + ": DTDC API call failed after " + attempts + " attempts.");
          sheet.getRange(i, statusColIndex).setValue("Error: DTDC API call failed after " + attempts + " attempts.");
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
      // For DELHIVERY calls, apply a 2-second delay if an API call was made.
      if (currentCourier === "DELHIVERY" && apiCalled) {
        Utilities.sleep(2000);
      }
    }
  }
  
  // Clear progress and execution state.
  PropertiesService.getScriptProperties().deleteProperty("currentRow");
  PropertiesService.getScriptProperties().deleteProperty("currentApi");
  PropertiesService.getScriptProperties().deleteProperty("currentResponseCode");
  PropertiesService.getScriptProperties().deleteProperty("trackingExecutionState");
  
  var summary = "Processed " + (lastRow - 1) + " rows.\nAPI called on " + processedCount + " rows.\nUpdated " + updateCount + " rows.";
  if (errors.length > 0) {
    summary += "\n\nErrors:\n" + errors.join("\n");
  }
  return summary;
}

/**
 * Toggle the pause/resume state.
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
 */
function stopExecution() {
  PropertiesService.getScriptProperties().setProperty("trackingExecutionState", "stopped");
  return "Execution stopped.";
}

/**
 * Returns the current row number being processed.
 */
function getCurrentRow() {
  return PropertiesService.getScriptProperties().getProperty("currentRow") || "";
}

/**
 * Returns the current API being used (short form).
 */
function getCurrentApi() {
  return PropertiesService.getScriptProperties().getProperty("currentApi") || "";
}

/**
 * Returns the current API response code.
 */
function getCurrentResponseCode() {
  return PropertiesService.getScriptProperties().getProperty("currentResponseCode") || "";
}

/**
 * Helper function: converts a column letter to its corresponding 1-indexed number.
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
