function saveMeetingSlots(meetingSlots) {
  Logger.log('Saving meeting slots to Google Sheet: ' + JSON.stringify(meetingSlots));
  
  try {
    if (!Array.isArray(meetingSlots)) {
      throw new Error('Meeting slots must be provided as an array');
    }
    
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Validate each meeting slot
    const validSlots = meetingSlots.map(slot => {
      if (!slot.dayOfWeek || !slot.time || !slot.duration) {
        throw new Error('Each meeting slot must have dayOfWeek, time, and duration');
      }
      return {
        dayOfWeek: String(slot.dayOfWeek).trim(),
        time: String(slot.time).trim(),
        duration: String(slot.duration).trim()
      };
    });
    
    // Clear existing data and write new slots
    const lastRow = meetingSlotsTab.getLastRow();
    if (lastRow > 1) {
      meetingSlotsTab.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    
    if (validSlots.length > 0) {
      // Prepare data for writing
      const currentDate = new Date();
      const dataToWrite = validSlots.map(slot => [
        slot.dayOfWeek,    // Column A: Day of Week
        slot.time,         // Column B: Time
        slot.duration,     // Column C: Duration (mins)
        currentDate,       // Column D: Date Added
        'Active'           // Column E: Status
      ]);
      
      // Write all data to sheet
      const range = meetingSlotsTab.getRange(2, 1, dataToWrite.length, 5);
      range.setValues(dataToWrite);
      
      // Format the date column
      const dateRange = meetingSlotsTab.getRange(2, 4, dataToWrite.length, 1);
      dateRange.setNumberFormat('MM/dd/yyyy HH:mm:ss');
      
      // Auto-resize columns
      meetingSlotsTab.autoResizeColumns(1, 5);
    }
    
    Logger.log('Successfully saved ' + validSlots.length + ' meeting slots to Google Sheet');
    return {
      success: true,
      count: validSlots.length,
      slots: validSlots
    };
    
  } catch (error) {
    Logger.log('Error in saveMeetingSlots: ' + error.toString());
    throw new Error('Failed to save meeting slots: ' + error.message);
  }
}

function getMeetingSlots() {
  Logger.log('Getting meeting slots from Google Sheet');
  
  try {
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Get all data from the Meeting Slots tab
    const lastRow = meetingSlotsTab.getLastRow();
    const lastCol = meetingSlotsTab.getLastColumn();
    
    Logger.log(`Meeting Slots tab has ${lastRow} rows and ${lastCol} columns`);
    
    if (lastRow <= 1) {
      Logger.log('No meeting slots found in sheet - only header row or empty');
      return [];
    }
    
    // Get data from columns A through C (starting from row 2 to skip headers)
    const dataRange = meetingSlotsTab.getRange(2, 1, lastRow - 1, 3);
    const allData = dataRange.getValues();
    
    Logger.log('Raw data from Meeting Slots tab:');
    for (let i = 0; i < allData.length; i++) {
      Logger.log(`Row ${i + 2}: [${allData[i][0]}] [${allData[i][1]}] [${allData[i][2]}]`);
    }
    
    // Process the data into the expected format with enhanced debugging
    const slots = allData
      .filter((row, index) => {
        // Enhanced checking with detailed logging
        const dayValue = row[0];
        const timeValue = row[1];
        const durationValue = row[2];
        
        Logger.log(`Filtering row ${index + 2}: Day type=${typeof dayValue}, Time type=${typeof timeValue}, Duration type=${typeof durationValue}`);
        
        // Check for day of week (allow string values)
        const hasDay = dayValue !== null && dayValue !== undefined && dayValue !== '' && String(dayValue).trim().length > 0;
        
        // Check for time (allow Date objects or string values)
        const hasTime = timeValue !== null && timeValue !== undefined && timeValue !== '' && 
                       (timeValue instanceof Date || String(timeValue).trim().length > 0);
        
        // Check for duration (allow numbers or string numbers)
        const hasDuration = durationValue !== null && durationValue !== undefined && durationValue !== '' && 
                           (typeof durationValue === 'number' || String(durationValue).trim().length > 0);
        
        Logger.log(`Row ${index + 2} checks: hasDay=${hasDay}, hasTime=${hasTime}, hasDuration=${hasDuration}`);
        Logger.log(`Row ${index + 2} values: Day="${dayValue}", Time="${timeValue}", Duration="${durationValue}"`);
        
        if (!hasDay || !hasTime || !hasDuration) {
          Logger.log(`SKIPPING row ${index + 2} due to missing data: Day=[${dayValue}] Time=[${timeValue}] Duration=[${durationValue}]`);
          return false;
        }
        
        Logger.log(`KEEPING row ${index + 2}: valid meeting slot data found`);
        return true;
      })
      .map(row => {
        // Format time properly if it's a Date object from Google Sheets
        let timeStr = row[1];
        if (row[1] instanceof Date) {
          // Extract just the time portion and format as 24-hour time to preserve accuracy
          const hours = row[1].getHours();
          const minutes = row[1].getMinutes();
          timeStr = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
          Logger.log(`Converted Date object to 24-hour format: ${timeStr}`);
        } else {
          timeStr = String(row[1]).trim();
          Logger.log(`Using string time value: ${timeStr}`);
        }
        
        const slot = {
          dayOfWeek: String(row[0]).trim(),
          time: timeStr,
          duration: String(row[2]).trim()
        };
        Logger.log(`Processed slot: ${JSON.stringify(slot)}`);
        return slot;
      });
    
    Logger.log('Retrieved ' + slots.length + ' valid meeting slots from Google Sheet');
    if (slots.length === 0) {
      Logger.log('WARNING: No valid meeting slots found. Check that the Meeting Slots tab has data in columns A, B, and C.');
    }
    
    return slots;
    
  } catch (error) {
    Logger.log('Error in getMeetingSlots: ' + error.toString());
    throw new Error('Failed to retrieve meeting slots: ' + error.message);
  }
}

function clearMeetingSlots() {
  Logger.log('Clearing meeting slots from Google Sheet');
  
  try {
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Clear all data except headers
    const lastRow = meetingSlotsTab.getLastRow();
    if (lastRow > 1) {
      meetingSlotsTab.getRange(2, 1, lastRow - 1, 5).clearContent();
      Logger.log('Cleared ' + (lastRow - 1) + ' rows of meeting slots from Google Sheet');
    } else {
      Logger.log('No meeting slots data to clear from Google Sheet');
    }
    
    Logger.log('Successfully cleared meeting slots from Google Sheet');
    return { success: true };
    
  } catch (error) {
    Logger.log('Error in clearMeetingSlots: ' + error.toString());
    throw new Error('Failed to clear meeting slots: ' + error.message);
  }
}

function getMeetingSlotsWithDates() {
  Logger.log('Getting meeting slots with future dates and availability based on recurring interval');
  
  try {
    // Get meeting slots configuration
    const meetingSlots = getMeetingSlots();
    if (!meetingSlots || meetingSlots.length === 0) {
      Logger.log('No meeting slots configured');
      return [];
    }
    
    // Get recurring interval (number of weeks to show)
    const weeksToShow = getRecurringInterval();
    Logger.log(`Showing meeting slots for next ${weeksToShow} weeks`);
    
    const today = new Date();
    const slotsWithDates = [];
    
    // For each meeting slot, find the next occurrence and subsequent weeks
    meetingSlots.forEach(slot => {
      const nextDate = getNextDateForDayOfWeek(today, slot.dayOfWeek);
      
      if (nextDate) {
        // Add occurrences for the specified number of weeks
        for (let weekOffset = 0; weekOffset < weeksToShow; weekOffset++) {
          const futureDate = new Date(nextDate);
          futureDate.setDate(nextDate.getDate() + (weekOffset * 7));
          
          // Parse the time and create start/end datetimes for availability check
          const slotDateTime = parseSlotTime(futureDate, slot.time, slot.duration);
          
          if (slotDateTime) {
            const isAvailable = checkSlotAvailability(slotDateTime.start, slotDateTime.end);
            
            slotsWithDates.push({
              dayOfWeek: slot.dayOfWeek,
              time: slot.time,
              duration: slot.duration,
              date: futureDate.toDateString(),
              weekOffset: weekOffset,
              available: isAvailable,
              startDateTime: slotDateTime.start,
              endDateTime: slotDateTime.end
            });
          }
        }
      }
    });
    
    // Sort by date
    slotsWithDates.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    Logger.log(`Generated ${slotsWithDates.length} meeting slot occurrences with availability`);
    return slotsWithDates;
    
  } catch (error) {
    Logger.log('Error in getMeetingSlotsWithDates: ' + error.toString());
    throw new Error('Failed to get meeting slots with dates: ' + error.message);
  }
}