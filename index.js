const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const { google } = require('googleapis');
const path = require('path');

const app = express();
const PORT = 3000;

// The core transformation logic remains the same.
function transformRawDataToProgram(rawData, sheetName) {
  const sessions = [];
  // Define the columns that represent a session. This can be made dynamic if needed.
  const sessionColumns = ['אימון 1', 'אימון 2', 'אימון 3'];

  for (const sessionName of sessionColumns) {
    // Check if this session column exists in any of the data rows.
    if (!rawData.some(row => row.hasOwnProperty(sessionName))) {
      continue; // Skip to the next session if this column is empty for the whole sheet.
    }

    const session = {
      session_number: sessionName,
      exercises: []
    };
    let currentExercise = null;

    for (const row of rawData) {
      // Process only if the row has a value for the current session column.
      if (row.hasOwnProperty(sessionName)) {
        const cellValue = String(row[sessionName]).trim();

        // Heuristic: A cell is a "set" if it contains '%' or starts with a digit.
        // Otherwise, it's an exercise "title".
        const isSet = cellValue.includes('%') || /^\d/.test(cellValue);

        if (!isSet) {
          // This is a title for a new exercise.
          currentExercise = {
            title: cellValue,
            sets: []
          };
          session.exercises.push(currentExercise);
        } else if (currentExercise) {
          // This is a set for the current exercise.
          currentExercise.sets.push(cellValue);
        }
      }
    }

    // Only add the session to our list if it actually contains exercises.
    if (session.exercises.length > 0) {
      sessions.push(session);
    }
  }

  // The sheet name becomes the week_date, and we wrap everything in the final structure.
  const week = {
    week_date: sheetName,
    sessions: sessions
  };

  return { weeks: [week] };
}

// --- Server Setup ---

// Route to serve the main landing page.
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});
// Middleware to parse URL-encoded bodies (as sent by HTML forms)
app.use(express.urlencoded({ extended: true }));

// Route to handle the form submission
app.post('/process', async (req, res) => {
  const { sheetUrl } = req.body;
  if (!sheetUrl) {
    return res.status(400).send('Google Sheet URL is required.');
  }

  // Extract the Sheet ID from the URL
  const match = sheetUrl.match(/\/d\/(.+?)\//);
  if (!match || !match[1]) {
    return res.status(400).send('Invalid Google Sheet URL format.');
  }
  const spreadsheetId = match[1];

  try {
    // Authenticate with Google using a service account
    const auth = new google.auth.GoogleAuth({
      keyFile: 'credentials.json', // The path to your service account key file
      scopes: ['https://www.googleapis.com/auth/drive.readonly'],
    });

    // Use the Google Drive API to export the sheet as an XLSX file buffer
    const drive = google.drive({ version: 'v3', auth });
    const fileResponse = await drive.files.export({
      fileId: spreadsheetId,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }, { responseType: 'arraybuffer' });

    // Parse the downloaded buffer
    const workbook = xlsx.read(fileResponse.data, { type: 'buffer' });

    // Process the workbook data just like in the original script
    const allWeeks = [];
    for (const sheetName of workbook.SheetNames) {
      console.log(`- Processing sheet: '${sheetName}'`);
      const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
      const programForSheet = transformRawDataToProgram(sheetData, sheetName);
      if (programForSheet.weeks && programForSheet.weeks.length > 0) {
        allWeeks.push(programForSheet.weeks[0]);
      }
    }

    const finalProgram = { weeks: allWeeks };
    const outputFileName = `parsed_program.json`;
    const jsonContent = JSON.stringify(finalProgram, null, 2);
    fs.writeFileSync(outputFileName, jsonContent, 'utf8');

    console.log(`\nSuccessfully created combined JSON file: ${outputFileName}`);
    // Send a success response to the browser
    res.send(`
      <h1>Success!</h1>
      <p>File processed and <code>${outputFileName}</code> has been created.</p>
      <a href="/download" download="${outputFileName}">Download JSON File</a>
      <br><br>
      <a href="/">Parse another</a>
    `);

  } catch (error) {
    console.error('Error processing Google Sheet:', error.message);
    res.status(500).send(`Error processing Google Sheet: ${error.message}. <br><br><strong>Common issues:</strong><br>1. Is the Google Drive API enabled in your Google Cloud project?<br>2. Have you shared the sheet with the service account email?<br>3. Is the <code>credentials.json</code> file in the correct location?`);
  }
});

app.get('/download', (req, res) => {
  const outputFileName = 'parsed_program.json';
  const filePath = path.join(__dirname, outputFileName);

  res.download(filePath, outputFileName, (err) => {
    if (err) {
      console.error(`Error sending file ${outputFileName}:`, err.message);
    }
  });
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
  console.log('Open your browser and navigate to the address to use the application.');
});
