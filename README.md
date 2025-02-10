# Shipment Tracker for Google Sheets

This project contains a Google Apps Script that tracks shipments using two different couriers—**DELHIVERY** and **DTDC**—directly from a Google Sheet. The script retrieves shipment status via API calls and updates the sheet with the latest tracking information. A persistent sidebar provides real-time progress, including details such as the current row, API in use, and response codes. In the DTDC branch, only the primary API is used, with up to 4 retry attempts if needed.

## Features

- **DELHIVERY Tracking:** Uses the DELHIVERY unified tracking API via a GET request.
- **DTDC Tracking:** Uses the primary DTDC API via a POST request and retries up to 4 times if unsuccessful.
- **Real-Time Sidebar:** Displays the current row being processed, the API being used, and the API response code.
- **Pause/Resume/Stop Functionality:** Allows you to pause, resume, or stop the script execution at any time.
- **Automatic Column Conversion:** Converts column letters (e.g., A, B, C) into column indices automatically.
- **Execution Summary:** Provides a summary of the rows processed, updates made, and errors encountered.

## Files

- **Code.gs:** Contains the main script logic, including the API calls, retry logic for DTDC, and execution control (pause/resume/stop).
- **TrackingSidebar.html:** Provides the HTML and JavaScript for the sidebar that shows live progress and control buttons.

## Setup Instructions

1. **Create or Open a Google Sheet:**
   - Open Google Sheets and create a new spreadsheet (or use an existing one).
   - Ensure that your sheet has appropriate headers (e.g., Tracking ID, Courier Name, etc.) in the first row.

2. **Access the Script Editor:**
   - In the Google Sheet, click on `Extensions` → `Apps Script` to open the script editor.

3. **Add the Code Files:**
   - Create a new script file named `Code.gs` and paste in the contents from the [Code.gs](./Code.gs) file.
   - Create a new HTML file named `TrackingSidebar.html` and paste in the contents from the [TrackingSidebar.html](./TrackingSidebar.html) file.

4. **Save and Authorize:**
   - Save your project (e.g., name it "Shipment Tracker").
   - Run the `onOpen` function (or any function) to trigger the authorization flow and grant the necessary permissions.

5. **Reload and Launch the Sidebar:**
   - Reload your Google Sheet.
   - You should see a new custom menu named **Shipment Tracker**.  
   - Click **Shipment Tracker → Show Tracker Sidebar** to open the sidebar.

## How to Use

1. **Enter Column Details:**
   - In the sidebar, input the column letters for:
     - **Tracking ID Column** (e.g., `A`)
     - **Courier Name Column** (e.g., `B`)
     - **Output: Status Column** (e.g., `C`)
     - **Output: Status_time Column** (e.g., `D`)
     - **Output: Execution_time Column** (e.g., `E`)

2. **Start Tracking:**
   - Click the **Start Tracking** button.
   - The script will begin processing rows from row 2 onward.
   - The sidebar displays the current row, the API being used, and the response code.

3. **Control Execution:**
   - **Pause/Resume:** Use the **Pause** button to temporarily halt execution and **Resume** to continue.
   - **Stop:** Click the **Stop Execution** button to completely halt the script.

4. **Review the Results:**
   - As the script runs, the sheet’s designated output columns will update with shipment status, status time, and the execution timestamp.
   - A summary of the execution (including any errors) is displayed once the process completes.

## Customization

- **API Endpoints:** If needed, you can update the API endpoints for DELHIVERY or DTDC directly in the `Code.gs` file.
- **Retry Logic:** The DTDC API call is retried up to 4 times. Modify the number of attempts or delay (currently 2 seconds between retries) if necessary.
- **Sidebar UI:** Feel free to modify `TrackingSidebar.html` to customize the user interface.

## Troubleshooting

- **Permissions:** Ensure you grant the necessary permissions for the script to run and access external services.
- **API Limits:** Be aware of any API rate limits. Adjust the delays if you encounter rate limiting issues.
- **Error Messages:** Errors are logged in the sidebar output. Use these messages to help troubleshoot issues.

## License

This project is licensed under the [MIT License](LICENSE).

## Contributing

Contributions are welcome! If you have suggestions or improvements, please fork the repository and submit a pull request.

---

*Happy tracking!*
