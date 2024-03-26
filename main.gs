function vacantDutyList() {
  // Get the active sheet
  const sheet = SpreadsheetApp.getActiveSheet();

  // Set the range of cells to search for empty cells
  const startRow = 15; // Change this to the first row of your range
  const startCol = 4; // Change this to the first column of your range
  const endRow = 45; // Change this to the last row of your range
  const endCol = 6; // Change this to the last column of your range
  const range = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
  
  // Get the values of the range
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  
  // Create an array of empty cells
  const emptyCells = [];
  values.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const backgroundColor = backgrounds[rowIndex][colIndex];
      if (cell === "" && backgroundColor === "#ffffff") {
        const cellAddress = sheet.getRange(startRow + rowIndex, startCol + colIndex).getA1Notation();
        emptyCells.push(cellAddress);
      }
    });
  });

  // Log the list of empty cells to the console
  console.log("Empty cells: " + emptyCells.join(", "));
}


function getTutorsAvailability() {
  const ss = SpreadsheetApp.getActiveSheet();
  const availabilitySheet = ss;
  const startRow = 6;  // Starting row for the data
  const startCol = 2;  // Starting column (column B)
  const numRows = 4;   // Number of rows to read (B6:B9)
  const numCols = 11;  // Number of columns to read (B:L)

  // Get the range of cells that contain the availability data
  const dataRange = availabilitySheet.getRange(startRow, startCol, numRows, numCols);
  const dataValues = dataRange.getValues();

  // Create an object to hold the availability of each tutor
  let tutorsAvailability = {};

  // Loop over each row to read each tutor's availability
  for (let i = 0; i < dataValues.length; i++) {
    let tutorName = String(dataValues[i][0]); // Ensure the tutor's name is a string
    tutorsAvailability[tutorName] = {};

    // Loop over the columns in pairs (date and duty type)
    for (let j = 1; j < dataValues[i].length; j += 2) {
      let blockingDay = dataValues[i][j];
      let dutyType = dataValues[i][j + 1];

      // Convert all values to strings to avoid errors
      blockingDay = String(blockingDay);
      dutyType = String(dutyType);

      if (!blockingDay || blockingDay === "") continue; // Skip if blocking day is empty

      // Check if the blocking day is a specific date or a day of the week
      let isDate = blockingDay.indexOf(' ') > -1; // If there's a space, assume it's a date (e.g. "Apr 17")
      let key = isDate ? 'dates' : 'days';

      // Initialize arrays to hold dates or days if they don't exist
      if (!tutorsAvailability[tutorName][key]) {
        tutorsAvailability[tutorName][key] = [];
      }

      // Add the date or day to the relevant array
      if (isDate) {
        // If it's a date range (e.g. "17 April - 24 April"), split and process
        if (blockingDay.includes('-')) {
          let dateRange = blockingDay.split('-').map(d => d.trim());
          tutorsAvailability[tutorName][key].push({ 'start': dateRange[0], 'end': dateRange[1], 'type': dutyType.trim() });
        } else {
          // Single date
          tutorsAvailability[tutorName][key].push({ 'date': blockingDay.trim(), 'type': dutyType.trim() });
        }
      } else {
        // Day of the week
        tutorsAvailability[tutorName][key].push({ 'day': blockingDay.trim(), 'type': dutyType.trim() });
      }
    }
  }

  // Log the tutor's availability for verification
  console.log(tutorsAvailability);
  console.log(JSON.stringify(tutorsAvailability, null, 2));

  
  return tutorsAvailability;
}

function assignDuties() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dutySheet = ss.getSheetByName('Duty Roster'); // Replace with your actual duty roster sheet name
  const tutorsAvailability = getTutorsAvailability(); // Get the availability object from the previous function

  // Get the vacant duty list (you would need to define how to retrieve this)
  const vacantDuties = vacantDutyList(); // This should return an array of objects with cell coordinates and duty types

  // Create a rotation index to ensure fair distribution
  const tutorNames = Object.keys(tutorsAvailability);
  let rotationIndex = 0;

  // Iterate through each vacant duty
  vacantDuties.forEach(vacantDuty => {
    let assigned = false;

    // Try to assign duty to a tutor in a round-robin manner
    while (!assigned) {
      const tutorName = tutorNames[rotationIndex % tutorNames.length];
      rotationIndex++;

      // Check tutor availability for this duty
      if (isTutorAvailable(tutorName, vacantDuty, tutorsAvailability)) {
        // Assign duty to tutor
        dutySheet.getRange(vacantDuty.cell).setValue(tutorName);
        assigned = true;
      }
    }
  });
}

function isTutorAvailable(tutorName, vacantDuty, tutorsAvailability) {
  // Extract the date and type of duty from vacantDuty
  const dutyDate = vacantDuty.date; // need to parse the cell to find the date
  const dutyType = vacantDuty.type; // need to parse the cell or have this info in vacantDuty

  // Check if tutor is unavailable on the specific dates
  if (tutorsAvailability[tutorName].dates) {
    for (const unavailability of tutorsAvailability[tutorName].dates) {
      if (unavailability.date === dutyDate && unavailability.type.includes(dutyType)) {
        return false;
      }
      // Check for date ranges and days of the week as well
    }
  }

  // Check if tutor is unavailable on the specific days of the week
  if (tutorsAvailability[tutorName].days) {
    for (const unavailability of tutorsAvailability[tutorName].days) {
      if (matchesDayOfWeek(unavailability.day, dutyDate) && unavailability.type.includes(dutyType)) {
        return false;
      }
    }
  }

  return true;
}

function matchesDayOfWeek(day, date) {
  // Convert the duty date to a day of the week and compare
  const dutyDayOfWeek = getDayOfWeekFromDate(date);
  return day === dutyDayOfWeek;
}

function getDayOfWeekFromDate(date) {
  // This is a placeholder function
  return "Monday"; // Replace with actual logic
}

// The vacantDutyList() function needs to return an array with details of each vacant duty, including the cell reference, date, and duty type

