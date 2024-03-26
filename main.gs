function vacantDutyList() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dutyCols = { 'D': 'Morning', 'E': 'Prep', 'F': 'Night' };
  const startRow = 15;
  const endRow = sheet.getLastRow();
  const filledColors = ['#4a86e8', '#93c47d', '#e06666']; // Replace with actual color codes of filled cells
  const vacantDuties = [];

  for (const col in dutyCols) {
    for (let row = startRow; row <= endRow; row++) {
      const cell = sheet.getRange(col + row);
      const cellColor = cell.getBackground();

      // Check if the cell's color is not one of the filled colors
      if (!filledColors.includes(cellColor)) {
        const cellDate = convertCellToDate('C' + row); // Assuming the date is in column C
        vacantDuties.push({
          cell: col + row,
          type: dutyCols[col],
          date: cellDate
        });
      }
    }
  }

  return vacantDuties;
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
  // console.log(tutorsAvailability);
  // console.log(JSON.stringify(tutorsAvailability, null, 2));

  
  return tutorsAvailability;
}


// main driver
function assignDuties() {
  const ss = SpreadsheetApp.getActiveSheet();
  const dutySheet = ss;
  const tutorsAvailability = getTutorsAvailability(); // Get the availability object from the previous function
  const vacantDuties = vacantDutyList(); // Assume this returns an array of objects with info about each vacant duty

  // Create a rotation index for each duty type to ensure fair distribution
  let rotationIndices = {
    'Morning': 0,
    'Prep': 0,
    'Night': 0
  };

  // Sort tutors by name to have a consistent order for rotation
  const tutorNames = Object.keys(tutorsAvailability).sort();

  // Iterate through each vacant duty
  vacantDuties.forEach(vacantDuty => {
    let dutyAssigned = false;

    // Track how many times we've looped through tutors to prevent infinite loops
    let loopCount = 0;

    // Keep trying to assign duty until successful
    while (!dutyAssigned && loopCount < tutorNames.length) {
      const tutorName = tutorNames[rotationIndices[vacantDuty.type] % tutorNames.length];
      rotationIndices[vacantDuty.type]++; // Move to the next tutor for this duty type

      if (isTutorAvailable(tutorName, vacantDuty, tutorsAvailability)) {
        // Assign duty to tutor
        dutySheet.getRange(vacantDuty.cell).setValue(tutorName);
        dutyAssigned = true;
      }

      loopCount++;
    }

    if (!dutyAssigned) {
      // Handle the case where no tutor was available for this duty
      Logger.log('No tutor available for duty on ' + vacantDuty.date + ' for ' + vacantDuty.type);
    }
  });
}


// Utils
function isTutorAvailable(tutorName, vacantDuty, tutorsAvailability) {
  const dutyDate = new Date(formatDate(vacantDuty.date)); // Convert the date string to a Date object
  const dutyType = vacantDuty.type;

  // Helper function to check if a duty type matches the tutor's unavailability
  function dutyMatchesUnavailableType(duty, unavailableType) {
    if (unavailableType === 'All') return true;
    if (unavailableType === duty) return true;
    if (unavailableType === 'Prep + Night' && (duty === 'Prep' || duty === 'Night')) return true;
    return false;
  }

  // Check specific dates unavailability
  if (tutorsAvailability[tutorName].dates) {
    for (const unavailability of tutorsAvailability[tutorName].dates) {
      if (dutyMatchesUnavailableType(dutyType, unavailability.type)) {
        // Single date unavailability
        if (unavailability.date) {
          const unavailableDate = new Date(formatDate(unavailability.date));
          if (isSameDate(dutyDate, unavailableDate)) {
            return false;
          }
        }
        // Date range unavailability
        if (unavailability.start && unavailability.end) {
          const startDate = new Date(formatDate(unavailability.start));
          const endDate = new Date(formatDate(unavailability.end));
          if (isDateInRange(dutyDate, startDate, endDate)) {
            return false;
          }
        }
      }
    }
  }

  // Check days of the week unavailability
  if (tutorsAvailability[tutorName].days) {
    for (const unavailability of tutorsAvailability[tutorName].days) {
      if (dutyMatchesUnavailableType(dutyType, unavailability.type)) {
        const unavailableDayIndex = convertDayStringToIndex(unavailability.day);
        if (dutyDate.getDay() === unavailableDayIndex) {
          return false;
        }
      }
    }
  }

  return true;
}

// Helper functions
function isSameDate(date1, date2) {
  return date1.getDate() === date2.getDate() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getFullYear() === date2.getFullYear();
}

function isDateInRange(date, start, end) {
  return start <= date && date <= end;
}

function convertDayStringToIndex(dayString) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days.indexOf(dayString);
}

function formatDate(date) {
  // Assuming the date is in "MMM dd" format, for example, "May 08"
  const dateParts = date.split(" ");
  const month = dateParts[0];
  const day = parseInt(dateParts[1], 10);
  const year = (new Date()).getFullYear();  // Assuming the year is the current year
  return new Date(`${month} ${day}, ${year}`).toISOString().substring(0, 10);  // Returns date in "YYYY-MM-DD" format
}

function matchesDayOfWeek(day, date) {
  // Convert the duty date to a day of the week and compare
  const dutyDayOfWeek = getDayOfWeekFromDate(date);
  return day === dutyDayOfWeek;
}

function convertCellToDate(cellRef) {
  // Get the active sheet
  const sheet = SpreadsheetApp.getActiveSheet();
  // Extract the row number from the cell reference
  const rowNum = cellRef.match(/\d+/)[0];
  // Get the date from column C of the same row
  const dateCell = 'C' + rowNum;
  const dateValue = sheet.getRange(dateCell).getDisplayValue();
  
  // Assuming the year is the current year
  const year = new Date().getFullYear();
  // Construct the full date string
  const fullDateString = `${dateValue}, ${year}`;
  
  // Return the full date string
  return fullDateString;
}

function getDayOfWeekFromDate(dateString) {
  // Parse the date string into a Date object
  const date = new Date(dateString);
  // Get the day of the week as a string (e.g., "Monday")
  const dayOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()];
  return dayOfWeek;
}



