const xlsx = require('xlsx');
const fs = require('fs');

/**
 * Parses an XLSX file and returns its content as a JSON object.
 * Each sheet in the XLSX file becomes a key in the returned object,
 * and its value is an array of objects representing the rows.
 * @param {string} filePath The path to the XLSX file.
 * @returns {Object} An object where keys are sheet names and values are sheet data.
 */
function parseXLSX(filePath) {
  try {
    // Read the file from the provided path.
    const workbook = xlsx.readFile(filePath);
    const sheetData = {};

    // Loop through each sheet name in the workbook.
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      // Convert the sheet to a JSON array of objects.
      // Each object represents a row, with headers as keys.
      const jsonData = xlsx.utils.sheet_to_json(sheet);
      sheetData[sheetName] = jsonData;
    });

    return sheetData;
  } catch (err) {
    // If there's an error reading or parsing the file, log it and return null.
    console.error(`Error processing XLSX file: ${err.message}`);
    return null;
  }
}

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



function main() {
  // The file path is expected as the third command-line argument.
  const filePath = process.argv[2];

  if (!filePath) {
    console.log('Please provide the path to your XLSX file.');
    console.log('Usage: node index.js <path_to_file.xlsx>');
    return;
  }

  const data = parseXLSX(filePath);

  if (data && Object.keys(data).length > 0) {
    const allWeeks = [];
    console.log('Processing all sheets from the workbook...');

    // Iterate over each sheet, treating each one as a week.
    for (const sheetName in data) {
      console.log(`- Processing sheet: '${sheetName}'`);
      const sheetData = data[sheetName];
      const programForSheet = transformRawDataToProgram(sheetData, sheetName);
      // The transform function returns { weeks: [week] }, so we extract the week object.
      if (programForSheet.weeks && programForSheet.weeks.length > 0) {
        allWeeks.push(programForSheet.weeks[0]);
      }
    }

    // Assemble the final program with all the processed weeks.
    const finalProgram = { weeks: allWeeks };

    // Define the output file name and prepare the JSON content.
    const outputFileName = `prased_pogram.json`;
    const jsonContent = JSON.stringify(finalProgram, null, 2);

    // Write the structured JSON to a new file in the main directory.
    fs.writeFileSync(outputFileName, jsonContent, 'utf8');
    console.log(`\nSuccessfully created combined JSON file: ${outputFileName}`);
  } else {
    console.log('No data found in the XLSX file or the file is empty.');
  }
}

main();
