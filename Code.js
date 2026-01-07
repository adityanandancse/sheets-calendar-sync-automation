/**
 * FULL CALENDAR SCRIPT - Google Sheets to Calendar Sync
 * * FEATURES:
 * 1. Forces Event Time to 11:00 AM - 12:00 PM.
 * 2. Sets TWO Reminders: 1 Week (7 Days) before + 3 Days before.
 * 3. Safe Write Mode: Saves "Status" and "ID" immediately after each row.
 */

const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('Reminder Manager')
    .addItem('Set Events (1 Week & 3 Day Reminders)', 'createTargetedEvents')
    .addToUi();
};

// --- CONFIGURATION ---
const createTargetedEvents = () => {
  // We pass an array of days for reminders: [7, 3]
  // 7 = 1 Week before, 3 = 3 Days before
  processEvents([7, 3]); 
};

// --- MAIN LOGIC ---
const processEvents = (daysArray) => {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Calendar Events Data'); 
  
  // 1. Check if sheet exists
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error: Sheet named 'Calendar Events Data' not found.");
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length <= 1) {
    SpreadsheetApp.getUi().alert("Error: No data found in the sheet.");
    return;
  }

  const [head, ...data] = values;

  // 2. Helper to find columns safely
  const getCol = (name) => {
    const idx = head.indexOf(name);
    if (idx === -1) throw new Error(`Column '${name}' is missing. Please check exact spelling.`);
    return idx;
  };

  try {
    // 3. Map Columns (Ensure these match your Sheet Headers exactly)
    const titleCol = getCol('Client Name');
    const outfitCol = getCol('Outfit Type');
    const assignDateCol = getCol('Date of Assigning');
    const deliveryDateCol = getCol('Date of Delivery');
    const jobCardCol = getCol('Job-Card No.');
    const statusCol = getCol('Status');
    const idCol = getCol('Event ID');
    
    // Optional Table No check
    const tableNoCol = head.indexOf('Table No.'); 

    const calendar = CalendarApp.getDefaultCalendar(); 
    
    // 4. LOOP THROUGH EVERY ROW
    data.forEach((row, i) => {
      const rowIndex = i + 2; // Calculate actual Excel row number
      
      const title = row[titleCol];
      let deliveryDate = row[deliveryDateCol]; 
      const existingEventId = row[idCol]; 
      
      // --- TIME LOGIC: FORCE 11:00 AM ---
      let startTime, endTime;
      if (deliveryDate instanceof Date) {
        startTime = new Date(deliveryDate);
        startTime.setHours(11, 0, 0); // Set start to 11:00 AM
        
        endTime = new Date(startTime);
        endTime.setHours(12, 0, 0);   // Set end to 12:00 PM
      }

      // --- DESCRIPTION BUILDER ---
      let description = `Outfit: ${row[outfitCol]}\nJob-Card No.: ${row[jobCardCol]}`;
      
      if(row[assignDateCol] instanceof Date) {
        description += `\nDate Assigned: ${row[assignDateCol].toDateString()}`;
      }
      
      if (tableNoCol !== -1 && row[tableNoCol]) {
        description += `\nTable No.: ${row[tableNoCol]}`;
      }

      let statusToSet = "";
      let idToSet = existingEventId;
      let event = null;

      try {
        // A. CHECK EXISTING
        if (existingEventId) {
          try {
            event = calendar.getEventById(existingEventId);
          } catch (e) {
            event = null; // ID existed in sheet, but event was deleted from Calendar
          }
        }

        // B. CREATE OR UPDATE
        if (event) {
          // UPDATE EXISTING EVENT
          event.setTitle(title);
          event.setDescription(description);
          if (startTime) event.setTime(startTime, endTime);
          
          applyReminders(event, daysArray);
          
          statusToSet = "Updated (1wk & 3day)";
          idToSet = event.getId(); 
        } else {
          // CREATE NEW EVENT
          if (title && startTime) {
            const options = { description: description };
            event = calendar.createEvent(title, startTime, endTime, options);
            
            applyReminders(event, daysArray);
            
            idToSet = event.getId();
            statusToSet = "Created (1wk & 3day)";
          } else {
            if (!title && !startTime) statusToSet = ""; 
            else statusToSet = "Error: Missing Name or Date";
          }
        }

      } catch (error) {
        statusToSet = `Error: ${error.message}`;
        console.error(error);
      }

      // --- CRITICAL: WRITE TO SHEET IMMEDIATELY ---
      if (statusToSet !== "") {
        sheet.getRange(rowIndex, statusCol + 1).setValue(statusToSet);
        if (idToSet) {
          sheet.getRange(rowIndex, idCol + 1).setValue(idToSet);
        }
        SpreadsheetApp.flush(); // Forces the sheet to update visually NOW
      }
    });

    SpreadsheetApp.getUi().alert("Success! All events updated with 1 Week & 3 Day reminders.");

  } catch (error) {
    SpreadsheetApp.getUi().alert(`Global Error: ${error.message}`);
  }
};

// --- HELPER: APPLY REMINDERS ---
function applyReminders(event, daysArray) {
  if (!event) return;
  
  try {
    event.removeAllReminders(); // Clear old ones
    
    daysArray.forEach(days => {
      const minutes = Math.floor(days * 24 * 60); 
      event.addPopupReminder(minutes);
    });
  } catch (e) {
    console.error("Error setting reminders: " + e.message);
  }
}
