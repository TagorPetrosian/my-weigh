# MyWeigh - Google Sheet Program Parser

MyWeigh is a web-based tool that takes a link to a Google Sheet, downloads it, and parses it into a structured JSON file suitable for analysis by other applications, including LLMs.

> **Note:** The command-line tool from the initial development stage is available on the `cli-tool` branch.

## Google Credentials Setup

To allow the application to access your private Google Sheets, you need to create a Service Account and share the sheet with it.

1.  **Create a Google Cloud Project:** Go to the Google Cloud Console and create a new project (or select an existing one).

2.  **Enable APIs:** In your project, go to the "APIs & Services" > "Library" section. Search for and enable the following API:
    *   Google Drive API

3.  **Create a Service Account:**
    *   Go to "APIs & Services" > "Credentials".
    *   Click "+ CREATE CREDENTIALS" and select "Service account".
    *   Give your service account a name (e.g., "sheets-parser") and click "Create and Continue".
    *   You can skip granting roles for now and click "Done".

4.  **Generate a JSON Key:**
    *   Find your newly created service account in the credentials list and click on it.
    *   Go to the "KEYS" tab.
    *   Click "ADD KEY" > "Create new key".
    *   Select "JSON" as the key type and click "CREATE". A JSON file will be downloaded to your computer.

5.  **Configure the Project:**
    *   Rename the downloaded JSON file to `credentials.json`.
    *   Place this `credentials.json` file in the root directory of this project.

6.  **Share Your Google Sheet:**
    *   Open the `credentials.json` file and find the `client_email` address (it looks like `...iam.gserviceaccount.com`).
    *   Open your Google Sheet, click the "Share" button, and paste this email address into the sharing dialog.
    *   Give it at least "Viewer" permissions and click "Share".

## How to Run

1.  **Install Dependencies**:
    ```bash
    npm install
    ```
2.  **Run the Server**:
    ```bash
    npm start
    ```
3.  **Open the App**: Open your web browser and navigate to `http://localhost:3000`.