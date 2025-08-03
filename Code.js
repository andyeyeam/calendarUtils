function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Andy\'s Calendar Utilities')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateStateSheet() {
  Logger.log('Getting or creating CalendarUtilities State Sheet');
  
  try {
    const sheetName = 'CalendarUtilities State Sheet';
    const tabName = 'Names';
    
    // Try to find existing sheet first
    const files = DriveApp.getFilesByName(sheetName);
    let sheet = null;
    
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('Found existing sheet with ID: ' + file.getId());
      sheet = SpreadsheetApp.openById(file.getId());
    } else {
      Logger.log('Creating new CalendarUtilities State Sheet');
      sheet = SpreadsheetApp.create(sheetName);
      Logger.log('Created new sheet with ID: ' + sheet.getId());
    }
    
    // Get or create the Names tab
    let namesTab = sheet.getSheetByName(tabName);
    if (!namesTab) {
      Logger.log('Creating Names tab');
      namesTab = sheet.insertSheet(tabName);
      
      // Set up headers
      namesTab.getRange(1, 1).setValue('Name');
      namesTab.getRange(1, 2).setValue('Date Added'); 
      namesTab.getRange(1, 3).setValue('Status');
      namesTab.getRange(1, 4).setValue('Calendar Event ID');
      namesTab.getRange(1, 5).setValue('Event Title');
      namesTab.getRange(1, 6).setValue('Calendar Link');
      namesTab.getRange(1, 7).setValue('Next Occurrence');
      
      // Format headers
      const headerRange = namesTab.getRange(1, 1, 1, 7);
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // Auto-resize columns
      namesTab.autoResizeColumns(1, 7);
      
      Logger.log('Names tab created and formatted');
    } else {
      Logger.log('Names tab already exists');
    }
    
    // Get or create the Meeting Slots tab
    const meetingSlotsTabName = 'Meeting Slots';
    let meetingSlotsTab = sheet.getSheetByName(meetingSlotsTabName);
    if (!meetingSlotsTab) {
      Logger.log('Creating Meeting Slots tab');
      meetingSlotsTab = sheet.insertSheet(meetingSlotsTabName);
      
      // Set up headers for Meeting Slots
      meetingSlotsTab.getRange(1, 1).setValue('Day of Week');
      meetingSlotsTab.getRange(1, 2).setValue('Time');
      meetingSlotsTab.getRange(1, 3).setValue('Duration (mins)');
      meetingSlotsTab.getRange(1, 4).setValue('Date Added');
      meetingSlotsTab.getRange(1, 5).setValue('Status');
      
      // Format headers
      const meetingHeaderRange = meetingSlotsTab.getRange(1, 1, 1, 5);
      meetingHeaderRange.setBackground('#cd853f');
      meetingHeaderRange.setFontColor('white');
      meetingHeaderRange.setFontWeight('bold');
      
      // Auto-resize columns
      meetingSlotsTab.autoResizeColumns(1, 5);
      
      Logger.log('Meeting Slots tab created and formatted');
    } else {
      Logger.log('Meeting Slots tab already exists');
    }
    
    // Get or create the Properties tab
    const propertiesTabName = 'Properties';
    let propertiesTab = sheet.getSheetByName(propertiesTabName);
    if (!propertiesTab) {
      Logger.log('Creating Properties tab');
      propertiesTab = sheet.insertSheet(propertiesTabName);
      
      // Set up headers for Properties
      propertiesTab.getRange(1, 1).setValue('Property Name');
      propertiesTab.getRange(1, 2).setValue('Value');
      propertiesTab.getRange(1, 3).setValue('Description');
      propertiesTab.getRange(1, 4).setValue('Last Updated');
      
      // Format headers
      const propertiesHeaderRange = propertiesTab.getRange(1, 1, 1, 4);
      propertiesHeaderRange.setBackground('#28a745');
      propertiesHeaderRange.setFontColor('white');
      propertiesHeaderRange.setFontWeight('bold');
      
      // Add default Recurring Interval property
      propertiesTab.getRange(2, 1).setValue('Recurring Interval');
      propertiesTab.getRange(2, 2).setValue(8); // Default value
      propertiesTab.getRange(2, 3).setValue('Meeting recurrence interval in weeks (1-26)');
      propertiesTab.getRange(2, 4).setValue(new Date());
      
      // Format the date column
      const dateRange = propertiesTab.getRange(2, 4, 1, 1);
      dateRange.setNumberFormat('MM/dd/yyyy HH:mm:ss');
      
      // Auto-resize columns
      propertiesTab.autoResizeColumns(1, 4);
      
      Logger.log('Properties tab created and formatted with default Recurring Interval');
    } else {
      Logger.log('Properties tab already exists');
    }
    
    return { sheet: sheet, namesTab: namesTab, meetingSlotsTab: meetingSlotsTab, propertiesTab: propertiesTab };
    
  } catch (error) {
    Logger.log('Error in getOrCreateStateSheet: ' + error.toString());
    throw new Error('Failed to get or create state sheet: ' + error.message);
  }
}

function searchCalendarEventForName(name) {
  Logger.log('Searching calendar events for name: ' + name);
  
  try {
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Search in a 6-month range to find recurring events
    const today = new Date();
    const sixMonthsAhead = new Date();
    sixMonthsAhead.setMonth(today.getMonth() + 6);
    
    // Get events in date range
    const allEvents = calendar.getEvents(today, sixMonthsAhead);
    Logger.log('Found ' + allEvents.length + ' events in 6-month range');
    
    // Look for events with "Skip Level:" and the specific name
    for (const event of allEvents) {
      const title = event.getTitle();
      
      if (title && title.includes('Skip Level:') && title.includes(name)) {
        try {
          const eventSeries = event.getEventSeries();
          if (eventSeries) {
            // This is a recurring event
            const eventId = eventSeries.getId();
            const eventStartTime = event.getStartTime();
            // Format date for Google Calendar day view URL: YYYY/M/D
            const year = eventStartTime.getFullYear();
            const month = eventStartTime.getMonth() + 1; // JavaScript months are 0-based
            const day = eventStartTime.getDate();
            const calendarLink = `https://calendar.google.com/calendar/u/0/r/day/${year}/${month}/${day}`;
            // Format date/time without timezone abbreviation
            const dateStr = eventStartTime.toLocaleDateString();
            const timeStr = eventStartTime.toLocaleTimeString().replace(/\s*\([^)]*\)/, '');
            const nextOccurrence = dateStr + ' ' + timeStr;
            
            Logger.log('Found calendar event for ' + name + ': ' + title);
            return {
              found: true,
              eventId: eventId,
              eventTitle: title,
              calendarLink: calendarLink,
              nextOccurrence: nextOccurrence,
              startTime: eventStartTime
            };
          }
        } catch (error) {
          Logger.log('Error processing event for ' + name + ': ' + error.toString());
          // Continue searching other events
        }
      }
    }
    
    Logger.log('No calendar event found for name: ' + name);
    return {
      found: false,
      eventId: null,
      eventTitle: null,
      calendarLink: null,
      nextOccurrence: null,
      startTime: null
    };
    
  } catch (error) {
    Logger.log('Error in searchCalendarEventForName: ' + error.toString());
    return {
      found: false,
      eventId: null,
      eventTitle: null,
      calendarLink: null,
      nextOccurrence: null,
      startTime: null
    };
  }
}


function getIndexPage() {
  return HtmlService.createHtmlOutputFromFile('index').getContent();
}




function debugCalendarEvents() {
  Logger.log('Debug: Getting recent calendar events for debugging');
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const today = new Date();
    const nextWeek = new Date();
    nextWeek.setDate(today.getDate() + 7);
    
    const recentEvents = calendar.getEvents(today, nextWeek);
    Logger.log('Found ' + recentEvents.length + ' events in next 7 days');
    
    const eventTitles = recentEvents.map((event, index) => {
      const title = event.getTitle();
      Logger.log('Event ' + (index + 1) + ': "' + title + '"');
      return title;
    });
    
    return {
      count: recentEvents.length,
      titles: eventTitles
    };
    
  } catch (error) {
    Logger.log('Error in debugCalendarEvents: ' + error.toString());
    throw new Error('Failed to debug calendar events: ' + error.message);
  }
}


function saveSkipLevelNames(names) {
  Logger.log('Saving skip level names to Google Sheet: ' + JSON.stringify(names));
  
  try {
    if (!Array.isArray(names)) {
      throw new Error('Names must be provided as an array');
    }
    
    // Clean and validate names
    const cleanNames = names
      .map(name => String(name).trim())
      .filter(name => name.length > 0);
    
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Allow empty arrays (when all names are removed)
    if (cleanNames.length === 0) {
      Logger.log('Clearing all names from sheet');
      
      // Clear all data except headers
      const lastRow = namesTab.getLastRow();
      if (lastRow > 1) {
        namesTab.getRange(2, 1, lastRow - 1, 7).clearContent();
      }
      
      return {
        success: true,
        count: 0,
        names: [],
        duplicatesRemoved: 0
      };
    }
    
    // Check for duplicates and remove them
    const uniqueNames = [];
    const seenNames = new Set();
    const duplicates = [];
    
    cleanNames.forEach(name => {
      const lowerName = name.toLowerCase();
      if (seenNames.has(lowerName)) {
        duplicates.push(name);
      } else {
        seenNames.add(lowerName);
        uniqueNames.push(name);
      }
    });
    
    if (duplicates.length > 0) {
      Logger.log('Removed duplicates: ' + JSON.stringify(duplicates));
    }
    
    // Clear existing data and write new names
    const lastRow = namesTab.getLastRow();
    if (lastRow > 1) {
      namesTab.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    
    // Search calendar events for each name and prepare data
    const currentDate = new Date();
    const dataToWrite = [];
    
    Logger.log('Searching calendar events for ' + uniqueNames.length + ' names...');
    
    for (const name of uniqueNames) {
      const calendarData = searchCalendarEventForName(name);
      
      dataToWrite.push([
        name,                                    // Column A: Name
        currentDate,                            // Column B: Date Added
        'Active',                               // Column C: Status
        calendarData.eventId || '',             // Column D: Calendar Event ID
        calendarData.eventTitle || '',          // Column E: Event Title
        calendarData.calendarLink || '',        // Column F: Calendar Link
        calendarData.nextOccurrence || ''       // Column G: Next Occurrence
      ]);
    }
    
    if (dataToWrite.length > 0) {
      const range = namesTab.getRange(2, 1, dataToWrite.length, 7);
      range.setValues(dataToWrite);
      
      // Format the date column
      const dateRange = namesTab.getRange(2, 2, dataToWrite.length, 1);
      dateRange.setNumberFormat('MM/dd/yyyy HH:mm:ss');
      
      // Auto-resize columns
      namesTab.autoResizeColumns(1, 7);
    }
    
    Logger.log('Successfully saved ' + uniqueNames.length + ' unique names to Google Sheet');
    return {
      success: true,
      count: uniqueNames.length,
      names: uniqueNames,
      duplicatesRemoved: duplicates.length
    };
    
  } catch (error) {
    Logger.log('Error in saveSkipLevelNames: ' + error.toString());
    throw new Error('Failed to save skip level names: ' + error.message);
  }
}

function getSkipLevelNames() {
  Logger.log('Getting skip level names from Google Sheet');
  
  try {
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Get all data from the Names tab
    const lastRow = namesTab.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('No names found in sheet');
      return [];
    }
    
    // Get names from column A (starting from row 2 to skip headers)
    const namesRange = namesTab.getRange(2, 1, lastRow - 1, 1);
    const namesData = namesRange.getValues();
    
    // Extract names and filter out empty cells
    const names = namesData
      .map(row => row[0])
      .filter(name => name && String(name).trim().length > 0)
      .map(name => String(name).trim());
    
    Logger.log('Retrieved ' + names.length + ' names from Google Sheet: ' + JSON.stringify(names));
    return names;
    
  } catch (error) {
    Logger.log('Error in getSkipLevelNames: ' + error.toString());
    throw new Error('Failed to retrieve skip level names: ' + error.message);
  }
}

function clearSkipLevelNames() {
  Logger.log('Clearing skip level names from Google Sheet');
  
  try {
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Clear all data except headers
    const lastRow = namesTab.getLastRow();
    if (lastRow > 1) {
      namesTab.getRange(2, 1, lastRow - 1, 7).clearContent();
      Logger.log('Cleared ' + (lastRow - 1) + ' rows of data from Google Sheet');
    } else {
      Logger.log('No data to clear from Google Sheet');
    }
    
    Logger.log('Successfully cleared skip level names from Google Sheet');
    return { 
      success: true, 
      count: 0,
      names: []
    };
    
  } catch (error) {
    Logger.log('Error in clearSkipLevelNames: ' + error.toString());
    throw new Error('Failed to clear skip level names: ' + error.message);
  }
}

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
    
    if (lastRow <= 1) {
      Logger.log('No meeting slots found in sheet');
      return [];
    }
    
    // Get data from columns A through C (starting from row 2 to skip headers)
    const dataRange = meetingSlotsTab.getRange(2, 1, lastRow - 1, 3);
    const allData = dataRange.getValues();
    
    // Process the data into the expected format
    const slots = allData
      .filter(row => row[0] && row[1] && row[2]) // Filter out rows with missing data
      .map(row => ({
        dayOfWeek: String(row[0]).trim(),
        time: String(row[1]).trim(),
        duration: String(row[2]).trim()
      }));
    
    Logger.log('Retrieved ' + slots.length + ' meeting slots from Google Sheet: ' + JSON.stringify(slots));
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

function removeSkipLevel(nameToRemove) {
  Logger.log('Removing skip level for name: ' + nameToRemove);
  
  try {
    if (!nameToRemove || typeof nameToRemove !== 'string') {
      throw new Error('Name to remove must be provided as a string');
    }
    
    const trimmedName = nameToRemove.trim();
    if (trimmedName.length === 0) {
      throw new Error('Name cannot be empty');
    }
    
    // Get current names
    const currentNames = getSkipLevelNames();
    if (!currentNames.includes(trimmedName)) {
      throw new Error('Name "' + trimmedName + '" not found in the stored names list');
    }
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Calculate date range - today to one year ahead to find all instances
    const today = new Date();
    const oneYearAhead = new Date();
    oneYearAhead.setFullYear(today.getFullYear() + 1);
    
    // Get all events in the date range
    const allEvents = calendar.getEvents(today, oneYearAhead);
    Logger.log('Total events found for removal search: ' + allEvents.length);
    
    // Find recurring events that contain both "Skip Level:" and the specific name
    let foundEventSeries = null;
    let deletedCount = 0;
    
    for (const event of allEvents) {
      const title = event.getTitle();
      
      // Check if this event matches our criteria
      if (title && title.includes('Skip Level:') && title.includes(trimmedName)) {
        try {
          const eventSeries = event.getEventSeries();
          if (eventSeries) {
            // This is a recurring event - we want to delete the entire series
            if (!foundEventSeries) {
              foundEventSeries = eventSeries;
              Logger.log('Found recurring event series to delete: ' + title);
              
              // Delete the entire event series (all future occurrences)
              eventSeries.deleteEventSeries();
              deletedCount = 1; // We deleted the series, count as 1
              Logger.log('Deleted recurring event series for: ' + trimmedName);
              break; // We found and deleted the series, no need to continue
            }
          } else {
            // This is a single event, delete it
            event.deleteEvent();
            deletedCount++;
            Logger.log('Deleted single event: ' + title);
          }
        } catch (error) {
          Logger.log('Error processing event "' + title + '": ' + error.toString());
          // Continue with other events even if one fails
        }
      }
    }
    
    if (deletedCount === 0) {
      Logger.log('No calendar events found for name: ' + trimmedName);
    }
    
    // Remove the name from the stored names list
    const updatedNames = currentNames.filter(name => name !== trimmedName);
    
    // Save the updated names list to Google Sheet
    saveSkipLevelNames(updatedNames);
    
    Logger.log('Successfully removed "' + trimmedName + '" from names list. Deleted ' + deletedCount + ' calendar event(s)');
    
    return {
      success: true,
      removedName: trimmedName,
      deletedEventCount: deletedCount,
      remainingNames: updatedNames
    };
    
  } catch (error) {
    Logger.log('Error in removeSkipLevel: ' + error.toString());
    throw new Error('Failed to remove skip level: ' + error.message);
  }
}



function createRecurringMeeting(nameForMeeting) {
  Logger.log('Creating recurring meeting for: ' + nameForMeeting);
  
  try {
    if (!nameForMeeting || typeof nameForMeeting !== 'string') {
      throw new Error('Name for meeting must be provided as a string');
    }
    
    const trimmedName = nameForMeeting.trim();
    if (trimmedName.length === 0) {
      throw new Error('Name cannot be empty');
    }
    
    // Get current names and meeting slots
    const allNames = getSkipLevelNames();
    const meetingSlots = getMeetingSlots();
    
    Logger.log('Total names: ' + allNames.length);
    Logger.log('Total meeting slots: ' + meetingSlots.length);
    Logger.log('Meeting slots data: ' + JSON.stringify(meetingSlots));
    
    if (meetingSlots.length === 0) {
      throw new Error('No meeting slots configured. Please configure meeting slots first.');
    }
    
    // Get the recurring interval from the Properties tab
    const recurringInterval = getRecurringInterval();
    const weeksToSearch = recurringInterval;
    
    Logger.log('Using recurring interval from Properties: ' + recurringInterval + ' weeks');
    Logger.log('Weeks to search for available slot: ' + weeksToSearch);
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Calculate search date range
    const now = new Date();
    const endDate = new Date();
    endDate.setDate(now.getDate() + (weeksToSearch * 7));
    
    Logger.log('Search range: ' + now.toString() + ' to ' + endDate.toString());
    Logger.log('Will search for available slots in the next ' + weeksToSearch + ' weeks');
    
    // Get all events in the search period to check for conflicts
    const existingEvents = calendar.getEvents(now, endDate);
    Logger.log('Found ' + existingEvents.length + ' existing events in search period');
    
    // Create array of all possible future slots to check chronologically
    const allPossibleSlots = [];
    
    for (let weekOffset = 0; weekOffset < weeksToSearch; weekOffset++) {
      for (const slot of meetingSlots) {
        // Get the day of week index (0 = Sunday, 1 = Monday, etc.)
        const dayMap = {
          'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, 
          'Thursday': 4, 'Friday': 5, 'Saturday': 6
        };
        const targetDayOfWeek = dayMap[slot.dayOfWeek];
        
        if (targetDayOfWeek === undefined) {
          Logger.log('Invalid day of week: ' + slot.dayOfWeek);
          continue;
        }
        
        // Parse the time with validation - handle different time formats
        let timeString = slot.time;
        
        if (!timeString) {
          Logger.log('Missing time value in slot: ' + JSON.stringify(slot));
          continue;
        }
        
        // Convert to string if it's not already
        timeString = String(timeString).trim();
        
        // If it's a Date object from Google Sheets, extract time portion
        if (timeString.includes('T') || timeString.includes('GMT') || timeString.length > 10) {
          try {
            const dateObj = new Date(timeString);
            if (!isNaN(dateObj.getTime())) {
              const hours = dateObj.getHours().toString().padStart(2, '0');
              const minutes = dateObj.getMinutes().toString().padStart(2, '0');
              timeString = hours + ':' + minutes;
              Logger.log('Converted date to time format: ' + timeString);
            }
          } catch (e) {
            Logger.log('Failed to parse as date: ' + timeString);
          }
        }
        
        const timeParts = timeString.split(':');
        if (timeParts.length < 2) {
          Logger.log('Time format must be HH:MM, got: ' + timeString + ' from original: ' + slot.time);
          continue;
        }
        
        const hour = parseInt(timeParts[0]);
        const minute = parseInt(timeParts[1]);
        const duration = parseInt(slot.duration);
        
        // Validate parsed values
        if (isNaN(hour) || isNaN(minute) || isNaN(duration)) {
          Logger.log('Failed to parse time/duration - Hour: ' + hour + ', Minute: ' + minute + ', Duration: ' + duration + ' from slot: ' + JSON.stringify(slot));
          continue;
        }
        
        if (hour < 0 || hour > 23 || minute < 0 || minute > 59 || duration <= 0) {
          Logger.log('Invalid time values - Hour: ' + hour + ', Minute: ' + minute + ', Duration: ' + duration);
          continue;
        }
        
        // Calculate the exact date for this slot in this week
        const weekStartDate = new Date(now);
        weekStartDate.setDate(now.getDate() + (weekOffset * 7));
        
        // Move to the beginning of this week (Sunday)
        const daysFromSunday = weekStartDate.getDay();
        weekStartDate.setDate(weekStartDate.getDate() - daysFromSunday);
        
        // Move to the target day of the week
        const slotDate = new Date(weekStartDate);
        slotDate.setDate(weekStartDate.getDate() + targetDayOfWeek);
        
        // Set the time
        const slotStart = new Date(slotDate);
        slotStart.setHours(hour, minute, 0, 0);
        
        // Validate the slot start time
        if (isNaN(slotStart.getTime())) {
          Logger.log('Invalid slot start time created');
          continue;
        }
        
        const slotEnd = new Date(slotStart);
        slotEnd.setMinutes(slotEnd.getMinutes() + duration);
        
        // Validate the slot end time
        if (isNaN(slotEnd.getTime())) {
          Logger.log('Invalid slot end time created');
          continue;
        }
        
        // Only include future slots
        if (slotStart <= now) {
          continue;
        }
        
        // Only include slots within our search range
        if (slotStart > endDate) {
          continue;
        }
        
        allPossibleSlots.push({
          slotStart,
          slotEnd,
          slot,
          weekOffset: weekOffset + 1
        });
      }
    }
    
    // Sort slots chronologically
    allPossibleSlots.sort((a, b) => a.slotStart.getTime() - b.slotStart.getTime());
    
    Logger.log('Found ' + allPossibleSlots.length + ' possible future slots to check');
    
    // Check each slot for availability
    for (const possibleSlot of allPossibleSlots) {
      const { slotStart, slotEnd, slot, weekOffset } = possibleSlot;
      
      Logger.log('Checking availability for: ' + slot.dayOfWeek + ' ' + slotStart.toDateString() + ' at ' + slot.time + ' (Week ' + weekOffset + ')');
      
      // Check for time overlap with existing events
      let hasConflict = false;
      for (const event of existingEvents) {
        const eventStart = event.getStartTime();
        const eventEnd = event.getEndTime();
        
        // Check for time overlap
        if ((slotStart < eventEnd && slotEnd > eventStart)) {
          hasConflict = true;
          Logger.log('Conflict found with event: ' + event.getTitle() + ' at ' + eventStart.toString());
          break;
        }
      }
      
      // If no conflict, create the recurring event
      if (!hasConflict) {
        Logger.log('Available slot found: ' + slot.dayOfWeek + ' ' + slotStart.toDateString() + ' at ' + slot.time + ' (Week ' + weekOffset + ' of ' + weeksToSearch + ')');
        
        const eventTitle = 'Skip Level: ' + trimmedName;
        
        Logger.log('Creating recurring event: ' + eventTitle);
        Logger.log('Start time: ' + slotStart.toString());
        Logger.log('End time: ' + slotEnd.toString());
        Logger.log('Recurrence: Every ' + recurringInterval + ' weeks');
        
        // Create the recurring event
        const recurringEvent = calendar.createEventSeries(
          eventTitle,
          slotStart,
          slotEnd,
          CalendarApp.newRecurrence().addWeeklyRule().interval(recurringInterval)
        );
        
        Logger.log('Successfully created recurring event series with ID: ' + recurringEvent.getId());
        
        // Update the Google Sheet Names tab with the calendar event details
        try {
          const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
          
          // Find the row with this name
          const lastRow = namesTab.getLastRow();
          if (lastRow > 1) {
            const nameRange = namesTab.getRange(2, 1, lastRow - 1, 7); // Get all data starting from row 2
            const allData = nameRange.getValues();
            
            // Find the row index for this name
            let targetRowIndex = -1;
            for (let i = 0; i < allData.length; i++) {
              if (allData[i][0] && String(allData[i][0]).trim().toLowerCase() === trimmedName.toLowerCase()) {
                targetRowIndex = i + 2; // +2 because we started from row 2 and need 1-based index
                break;
              }
            }
            
            if (targetRowIndex > 0) {
              Logger.log('Found name "' + trimmedName + '" at row ' + targetRowIndex + ', updating calendar details');
              
              // Generate calendar link to the day view instead of specific event
              // Format date for Google Calendar day view URL: YYYY/M/D
              const year = slotStart.getFullYear();
              const month = slotStart.getMonth() + 1; // JavaScript months are 0-based
              const day = slotStart.getDate();
              const calendarLink = `https://calendar.google.com/calendar/u/0/r/day/${year}/${month}/${day}`;
              
              // Calculate next occurrence date
              const nextOccurrenceDate = slotStart.toLocaleDateString() + ' ' + slotStart.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
              
              // Update the calendar-related columns
              namesTab.getRange(targetRowIndex, 3).setValue('Active'); // Status
              namesTab.getRange(targetRowIndex, 4).setValue(recurringEvent.getId()); // Calendar Event ID
              namesTab.getRange(targetRowIndex, 5).setValue(eventTitle); // Event Title
              namesTab.getRange(targetRowIndex, 6).setValue(calendarLink); // Calendar Link
              namesTab.getRange(targetRowIndex, 7).setValue(nextOccurrenceDate); // Next Occurrence
              
              Logger.log('Updated Google Sheet Names tab for "' + trimmedName + '" with calendar event details');
            } else {
              Logger.log('Warning: Could not find name "' + trimmedName + '" in Names tab to update calendar details');
            }
          }
        } catch (updateError) {
          Logger.log('Error updating Google Sheet Names tab: ' + updateError.toString());
          // Don't fail the entire operation if sheet update fails
        }
        
        return {
          success: true,
          meetingCreated: true,
          weeksToSearch: weeksToSearch,
          recurringInterval: recurringInterval,
          meetingDetails: {
            dayOfWeek: slot.dayOfWeek,
            time: slot.time,
            duration: slot.duration,
            startDateTime: slotStart.toISOString(),
            endDateTime: slotEnd.toISOString(),
            title: eventTitle,
            weekFound: weekOffset
          },
          recurringEventId: recurringEvent.getId()
        };
      } else {
        Logger.log('Conflict found for ' + slot.dayOfWeek + ' at ' + slot.time + ' in week ' + weekOffset + ', trying next available slot...');
      }
    }
    
    // No available slots found
    Logger.log('No available slots found in any meeting slot configuration');
    
    return {
      success: true,
      meetingCreated: false,
      weeksToSearch: weeksToSearch,
      recurringInterval: recurringInterval,
      message: 'No available slots found in the next ' + weeksToSearch + ' weeks'
    };
    
  } catch (error) {
    Logger.log('Error in createRecurringMeeting: ' + error.toString());
    throw new Error('Failed to create recurring meeting: ' + error.message);
  }
}

function createMeetingsForSpecificNames(namesToProcess) {
  Logger.log('Creating recurring meetings for specific names: ' + JSON.stringify(namesToProcess));
  
  try {
    if (!Array.isArray(namesToProcess) || namesToProcess.length === 0) {
      Logger.log('No names provided to process');
      return {
        success: true,
        namesProcessed: 0,
        meetingsCreated: 0,
        namesWithoutSlots: 0,
        weeksCalculated: 0,
        errors: []
      };
    }
    
    // Get all loaded names and meeting slots
    const allNames = getSkipLevelNames();
    const meetingSlots = getMeetingSlots();
    
    if (meetingSlots.length === 0) {
      throw new Error('No meeting slots configured. Please configure meeting slots first.');
    }
    
    Logger.log('Processing ' + namesToProcess.length + ' specific names with ' + meetingSlots.length + ' meeting slots');
    
    // Calculate X weeks: max of 8 or (total names / total meeting slots) rounded up
    // Use total names count for calculation, not just the names being processed
    const calculatedWeeks = Math.ceil(allNames.length / meetingSlots.length);
    const weeksToSearch = Math.max(8, calculatedWeeks);
    
    Logger.log('Calculated weeks: ' + calculatedWeeks + ', Final weeks to search: ' + weeksToSearch);
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Calculate search date range
    const today = new Date();
    const endDate = new Date();
    endDate.setDate(today.getDate() + (weeksToSearch * 7));
    
    Logger.log('Search range: ' + today.toString() + ' to ' + endDate.toString());
    
    // Get all events in the search period to check for conflicts
    const existingEvents = calendar.getEvents(today, endDate);
    Logger.log('Found ' + existingEvents.length + ' existing events in search period');
    
    let meetingsCreated = 0;
    let namesWithoutSlots = 0;
    const errors = [];
    const createdMeetings = [];
    
    // Track allocated slots to avoid reuse within this batch
    const allocatedSlots = new Set(); // Set of "weekIndex-slotIndex" strings
    
    // Create a function to generate slot key
    function getSlotKey(weekIndex, slotIndex) {
      return weekIndex + '-' + slotIndex;
    }
    
    // Process each specified name
    for (const name of namesToProcess) {
      Logger.log('=== Processing name: ' + name + ' ===');
      
      let meetingCreatedForName = false;
      
      // Search for next available slot that hasn't been allocated yet
      for (let weekCount = 0; weekCount < weeksToSearch && !meetingCreatedForName; weekCount++) {
        Logger.log('Checking Week ' + (weekCount + 1) + ' of ' + weeksToSearch + ' for ' + name);
        
        // Try all meeting slots in this week in order
        for (let slotIndex = 0; slotIndex < meetingSlots.length; slotIndex++) {
          const slot = meetingSlots[slotIndex];
          const slotKey = getSlotKey(weekCount, slotIndex);
          
          // Skip if this slot has already been allocated to another name in this batch
          if (allocatedSlots.has(slotKey)) {
            Logger.log('Slot already allocated to another name in this batch: Week ' + (weekCount + 1) + ', Slot ' + (slotIndex + 1) + ' (' + slot.dayOfWeek + ' ' + slot.time + ')');
            continue;
          }
          
          Logger.log('Checking slot for ' + name + ' in week ' + (weekCount + 1) + ', slot ' + (slotIndex + 1) + ': ' + JSON.stringify(slot));
          
          // Get the day of week index (0 = Sunday, 1 = Monday, etc.)
          const dayMap = {
            'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, 
            'Thursday': 4, 'Friday': 5, 'Saturday': 6
          };
          const targetDayOfWeek = dayMap[slot.dayOfWeek];
          
          if (targetDayOfWeek === undefined) {
            Logger.log('Invalid day of week: ' + slot.dayOfWeek);
            continue;
          }
          
          // Parse the time
          const timeParts = slot.time.split(':');
          const hour = parseInt(timeParts[0]);
          const minute = parseInt(timeParts[1]);
          const duration = parseInt(slot.duration);
          
          // Find the occurrence of this day in this specific week
          let currentDate = new Date(today);
          currentDate.setDate(today.getDate() + (weekCount * 7)); // Move to the target week
          
          // Move to the target day of the week within this week
          while (currentDate.getDay() !== targetDayOfWeek) {
            currentDate.setDate(currentDate.getDate() + 1);
          }
          
          // Create the start and end times for this potential slot
          const slotStart = new Date(currentDate);
          slotStart.setHours(hour, minute, 0, 0);
          
          const slotEnd = new Date(slotStart);
          slotEnd.setMinutes(slotEnd.getMinutes() + duration);
          
          // Skip if the slot is in the past
          if (slotStart <= new Date()) {
            Logger.log('Skipping past slot for ' + name + ': ' + slotStart.toString());
            continue;
          }
          
          // Check if this slot is beyond our search range
          if (slotStart > endDate) {
            Logger.log('Slot is beyond search range for ' + name + ': ' + slotStart.toString());
            continue;
          }
          
          Logger.log('Checking availability for ' + name + ': ' + slot.dayOfWeek + ' ' + slotStart.toDateString() + ' at ' + slot.time + ' (Week ' + (weekCount + 1) + ', Slot ' + (slotIndex + 1) + ')');
          
          // Check for time overlap with existing events (including previously created meetings in this batch)
          let hasConflict = false;
          for (const event of existingEvents) {
            const eventStart = event.getStartTime();
            const eventEnd = event.getEndTime();
            
            // Check for time overlap
            if ((slotStart < eventEnd && slotEnd > eventStart)) {
              hasConflict = true;
              Logger.log('Conflict found for ' + name + ' with existing event: ' + event.getTitle() + ' at ' + eventStart.toString());
              break;
            }
          }
          
          // If no conflict, create the recurring event and mark slot as allocated
          if (!hasConflict) {
            try {
              Logger.log('Available slot found for ' + name + ': ' + slot.dayOfWeek + ' ' + currentDate.toDateString() + ' at ' + slot.time + ' (Week ' + (weekCount + 1) + ', Slot ' + (slotIndex + 1) + ')');
              
              const eventTitle = 'Skip Level: ' + name;
              
              Logger.log('Creating recurring event for ' + name + ': ' + eventTitle);
              Logger.log('Start time: ' + slotStart.toString());
              Logger.log('End time: ' + slotEnd.toString());
              Logger.log('Recurrence: Every ' + weeksToSearch + ' weeks');
              
              // Create the recurring event
              const recurringEvent = calendar.createEventSeries(
                eventTitle,
                slotStart,
                slotEnd,
                CalendarApp.newRecurrence().addWeeklyRule().interval(weeksToSearch)
              );
              
              Logger.log('Successfully created recurring event for ' + name + ' with ID: ' + recurringEvent.getId());
              
              // Mark this slot as allocated so no other name in this batch can use it
              allocatedSlots.add(slotKey);
              Logger.log('Marked slot as allocated: ' + slotKey + ' (' + slot.dayOfWeek + ' ' + slot.time + ')');
              
              meetingsCreated++;
              meetingCreatedForName = true;
              
              createdMeetings.push({
                name: name,
                dayOfWeek: slot.dayOfWeek,
                time: slot.time,
                duration: slot.duration,
                weekFound: weekCount + 1,
                slotIndex: slotIndex + 1
              });
              
              break; // Move to next name
              
            } catch (error) {
              const errorMsg = 'Failed to create meeting for "' + name + '": ' + error.message;
              Logger.log(errorMsg);
              errors.push(errorMsg);
            }
          } else {
            Logger.log('Conflict found for ' + name + ' in ' + slot.dayOfWeek + ' at ' + slot.time + ' in week ' + (weekCount + 1) + ', slot ' + (slotIndex + 1));
          }
        }
      }
      
      if (!meetingCreatedForName) {
        Logger.log('No available slots found for ' + name + ' in ' + weeksToSearch + ' weeks');
        namesWithoutSlots++;
      }
    }
    
    Logger.log('Specific names meeting creation completed. Processed: ' + namesToProcess.length + ', Created: ' + meetingsCreated + ', Without slots: ' + namesWithoutSlots + ', Errors: ' + errors.length);
    
    return {
      success: true,
      namesProcessed: namesToProcess.length,
      meetingsCreated: meetingsCreated,
      namesWithoutSlots: namesWithoutSlots,
      weeksCalculated: weeksToSearch,
      createdMeetings: createdMeetings,
      errors: errors
    };
    
  } catch (error) {
    Logger.log('Error in createMeetingsForSpecificNames: ' + error.toString());
    throw new Error('Failed to create meetings for specific names: ' + error.message);
  }
}

function createAllRecurringMeetings() {
  Logger.log('Creating recurring meetings for all names');
  
  try {
    // Get all loaded names and meeting slots
    const allNames = getSkipLevelNames();
    const meetingSlots = getMeetingSlots();
    
    if (allNames.length === 0) {
      Logger.log('No names loaded');
      return {
        success: true,
        totalNames: 0,
        meetingsCreated: 0,
        namesAlreadyWithMeetings: 0,
        namesWithoutSlots: 0,
        weeksCalculated: 0,
        errors: []
      };
    }
    
    if (meetingSlots.length === 0) {
      throw new Error('No meeting slots configured. Please configure meeting slots first.');
    }
    
    Logger.log('Processing ' + allNames.length + ' names with ' + meetingSlots.length + ' meeting slots');
    
    // Calculate X weeks: max of 8 or (total names / total meeting slots) rounded up
    const calculatedWeeks = Math.ceil(allNames.length / meetingSlots.length);
    const weeksToSearch = Math.max(8, calculatedWeeks);
    
    Logger.log('Calculated weeks: ' + calculatedWeeks + ', Final weeks to search: ' + weeksToSearch);
    
    // Get existing calendar events to check which names already have meetings
    const namesWithEvents = getNamesWithCalendarEvents();
    const namesAlreadyWithMeetings = namesWithEvents.filter(nameData => nameData.found);
    const namesToProcess = namesWithEvents.filter(nameData => !nameData.found);
    
    Logger.log('Names already with meetings: ' + namesAlreadyWithMeetings.length);
    Logger.log('Names to process: ' + namesToProcess.length);
    
    if (namesToProcess.length === 0) {
      Logger.log('All names already have meetings');
      return {
        success: true,
        totalNames: allNames.length,
        meetingsCreated: 0,
        namesAlreadyWithMeetings: namesAlreadyWithMeetings.length,
        namesWithoutSlots: 0,
        weeksCalculated: weeksToSearch,
        errors: []
      };
    }
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Calculate search date range
    const today = new Date();
    const endDate = new Date();
    endDate.setDate(today.getDate() + (weeksToSearch * 7));
    
    Logger.log('Search range: ' + today.toString() + ' to ' + endDate.toString());
    
    // Get all events in the search period to check for conflicts
    const existingEvents = calendar.getEvents(today, endDate);
    Logger.log('Found ' + existingEvents.length + ' existing events in search period');
    
    let meetingsCreated = 0;
    let namesWithoutSlots = 0;
    const errors = [];
    const createdMeetings = [];
    
    // Process each name that doesn't have a meeting
    for (const nameData of namesToProcess) {
      const name = nameData.name;
      Logger.log('=== Processing name: ' + name + ' ===');
      
      let meetingCreatedForName = false;
      
      // Search for first available slot by checking all meeting slots in each week
      for (let weekCount = 0; weekCount < weeksToSearch && !meetingCreatedForName; weekCount++) {
        Logger.log('Checking Week ' + (weekCount + 1) + ' of ' + weeksToSearch + ' for ' + name);
        
        // Try all meeting slots in this week
        for (const slot of meetingSlots) {
          Logger.log('Checking slot for ' + name + ' in week ' + (weekCount + 1) + ': ' + JSON.stringify(slot));
          
          // Get the day of week index (0 = Sunday, 1 = Monday, etc.)
          const dayMap = {
            'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, 
            'Thursday': 4, 'Friday': 5, 'Saturday': 6
          };
          const targetDayOfWeek = dayMap[slot.dayOfWeek];
          
          if (targetDayOfWeek === undefined) {
            Logger.log('Invalid day of week: ' + slot.dayOfWeek);
            continue;
          }
          
          // Parse the time
          const timeParts = slot.time.split(':');
          const hour = parseInt(timeParts[0]);
          const minute = parseInt(timeParts[1]);
          const duration = parseInt(slot.duration);
          
          // Find the occurrence of this day in this specific week
          let currentDate = new Date(today);
          currentDate.setDate(today.getDate() + (weekCount * 7)); // Move to the target week
          
          // Move to the target day of the week within this week
          while (currentDate.getDay() !== targetDayOfWeek) {
            currentDate.setDate(currentDate.getDate() + 1);
          }
          
          // Create the start and end times for this potential slot
          const slotStart = new Date(currentDate);
          slotStart.setHours(hour, minute, 0, 0);
          
          const slotEnd = new Date(slotStart);
          slotEnd.setMinutes(slotEnd.getMinutes() + duration);
          
          // Skip if the slot is in the past
          if (slotStart <= new Date()) {
            Logger.log('Skipping past slot for ' + name + ': ' + slotStart.toString());
            continue;
          }
          
          // Check if this slot is beyond our search range
          if (slotStart > endDate) {
            Logger.log('Slot is beyond search range for ' + name + ': ' + slotStart.toString());
            continue;
          }
          
          Logger.log('Checking availability for ' + name + ': ' + slot.dayOfWeek + ' ' + slotStart.toDateString() + ' at ' + slot.time);
          
          // Check for time overlap with existing events
          let hasConflict = false;
          for (const event of existingEvents) {
            const eventStart = event.getStartTime();
            const eventEnd = event.getEndTime();
            
            // Check for time overlap
            if ((slotStart < eventEnd && slotEnd > eventStart)) {
              hasConflict = true;
              Logger.log('Conflict found for ' + name + ' with event: ' + event.getTitle() + ' at ' + eventStart.toString());
              break;
            }
          }
          
          // If no conflict, create the recurring event
          if (!hasConflict) {
            try {
              Logger.log('Available slot found for ' + name + ': ' + slot.dayOfWeek + ' ' + currentDate.toDateString() + ' at ' + slot.time);
              
              const eventTitle = 'Skip Level: ' + name;
              
              Logger.log('Creating recurring event for ' + name + ': ' + eventTitle);
              Logger.log('Start time: ' + slotStart.toString());
              Logger.log('End time: ' + slotEnd.toString());
              Logger.log('Recurrence: Every ' + weeksToSearch + ' weeks');
              
              // Create the recurring event
              const recurringEvent = calendar.createEventSeries(
                eventTitle,
                slotStart,
                slotEnd,
                CalendarApp.newRecurrence().addWeeklyRule().interval(weeksToSearch)
              );
              
              Logger.log('Successfully created recurring event for ' + name + ' with ID: ' + recurringEvent.getId());
              
              meetingsCreated++;
              meetingCreatedForName = true;
              
              createdMeetings.push({
                name: name,
                dayOfWeek: slot.dayOfWeek,
                time: slot.time,
                duration: slot.duration,
                weekFound: weekCount + 1
              });
              
              break; // Move to next name
              
            } catch (error) {
              const errorMsg = 'Failed to create meeting for "' + name + '": ' + error.message;
              Logger.log(errorMsg);
              errors.push(errorMsg);
            }
          } else {
            Logger.log('Conflict found for ' + name + ' in ' + slot.dayOfWeek + ' at ' + slot.time + ' in week ' + (weekCount + 1));
          }
        }
      }
      
      if (!meetingCreatedForName) {
        Logger.log('No available slots found for ' + name + ' in ' + weeksToSearch + ' weeks');
        namesWithoutSlots++;
      }
    }
    
    Logger.log('Bulk meeting creation completed. Created: ' + meetingsCreated + ', Without slots: ' + namesWithoutSlots + ', Errors: ' + errors.length);
    
    return {
      success: true,
      totalNames: allNames.length,
      meetingsCreated: meetingsCreated,
      namesAlreadyWithMeetings: namesAlreadyWithMeetings.length,
      namesWithoutSlots: namesWithoutSlots,
      weeksCalculated: weeksToSearch,
      createdMeetings: createdMeetings,
      errors: errors
    };
    
  } catch (error) {
    Logger.log('Error in createAllRecurringMeetings: ' + error.toString());
    throw new Error('Failed to create all recurring meetings: ' + error.message);
  }
}

function deleteAllRecurringMeetings() {
  Logger.log('Deleting all recurring meetings');
  
  try {
    // Get all loaded names
    const allNames = getSkipLevelNames();
    if (allNames.length === 0) {
      Logger.log('No names loaded');
      return {
        success: true,
        totalNames: 0,
        meetingsDeleted: 0,
        namesWithoutMeetings: 0,
        errors: []
      };
    }
    
    Logger.log('Processing ' + allNames.length + ' names for meeting deletion');
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Use 1 year range to find all recurring events
    const today = new Date();
    const oneYearAhead = new Date();
    oneYearAhead.setFullYear(today.getFullYear() + 1);
    
    // Get all events in the date range
    const allEvents = calendar.getEvents(today, oneYearAhead);
    Logger.log('Total events found for deletion search: ' + allEvents.length);
    
    // Find all Skip Level events
    const skipLevelEvents = allEvents.filter(event => {
      const title = event.getTitle();
      return title && title.includes('Skip Level:');
    });
    Logger.log('Events with "Skip Level:" in title: ' + skipLevelEvents.length);
    
    // Group events by series ID to avoid deleting the same series multiple times
    const eventSeriesMap = new Map(); // seriesId -> {series, title, names}
    const singleEventsToDelete = []; // Non-recurring events
    
    skipLevelEvents.forEach(event => {
      const title = event.getTitle();
      Logger.log('Processing event: "' + title + '"');
      
      try {
        const eventSeries = event.getEventSeries();
        if (eventSeries) {
          // This is a recurring event
          const seriesId = eventSeries.getId();
          if (!eventSeriesMap.has(seriesId)) {
            // Find which names this series matches
            const matchingNames = allNames.filter(name => title.includes(name));
            eventSeriesMap.set(seriesId, {
              series: eventSeries,
              title: title,
              names: matchingNames
            });
            Logger.log('Added recurring series for deletion: ' + title + ' (matches: ' + matchingNames.join(', ') + ')');
          }
        } else {
          // This is a single event
          singleEventsToDelete.push({
            event: event,
            title: title
          });
          Logger.log('Added single event for deletion: ' + title);
        }
      } catch (error) {
        Logger.log('Error processing event "' + title + '": ' + error.toString());
      }
    });
    
    Logger.log('Found ' + eventSeriesMap.size + ' recurring series and ' + singleEventsToDelete.length + ' single events to delete');
    
    let meetingsDeleted = 0;
    const errors = [];
    
    // Delete recurring event series
    for (const [seriesId, eventData] of eventSeriesMap) {
      try {
        Logger.log('Deleting recurring series: ' + eventData.title);
        eventData.series.deleteEventSeries();
        meetingsDeleted++;
        Logger.log('Successfully deleted recurring series: ' + eventData.title);
      } catch (error) {
        const errorMsg = 'Failed to delete recurring series "' + eventData.title + '": ' + error.message;
        Logger.log(errorMsg);
        errors.push(errorMsg);
      }
    }
    
    // Delete single events
    singleEventsToDelete.forEach(eventData => {
      try {
        Logger.log('Deleting single event: ' + eventData.title);
        eventData.event.deleteEvent();
        meetingsDeleted++;
        Logger.log('Successfully deleted single event: ' + eventData.title);
      } catch (error) {
        const errorMsg = 'Failed to delete single event "' + eventData.title + '": ' + error.message;
        Logger.log(errorMsg);
        errors.push(errorMsg);
      }
    });
    
    const namesWithoutMeetings = allNames.length - (eventSeriesMap.size + singleEventsToDelete.length);
    
    Logger.log('Deletion completed. Meetings deleted: ' + meetingsDeleted + ', Names without meetings: ' + namesWithoutMeetings + ', Errors: ' + errors.length);
    
    return {
      success: true,
      totalNames: allNames.length,
      meetingsDeleted: meetingsDeleted,
      namesWithoutMeetings: namesWithoutMeetings,
      errors: errors
    };
    
  } catch (error) {
    Logger.log('Error in deleteAllRecurringMeetings: ' + error.toString());
    throw new Error('Failed to delete all recurring meetings: ' + error.message);
  }
}

function getNamesWithCalendarEvents() {
  Logger.log('Getting names with calendar events from Google Sheet');
  
  try {
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Get all data from the Names tab
    const lastRow = namesTab.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('No names found in sheet');
      return [];
    }
    
    // Get all data from columns A through G (starting from row 2 to skip headers)
    const dataRange = namesTab.getRange(2, 1, lastRow - 1, 7);
    const allData = dataRange.getValues();
    
    // Process the data into the expected format
    const results = allData
      .filter(row => row[0] && String(row[0]).trim().length > 0) // Filter out empty names
      .map(row => {
        const name = String(row[0]).trim();
        const eventId = row[3] ? String(row[3]).trim() : '';
        const eventTitle = row[4] ? String(row[4]).trim() : '';
        const calendarLink = row[5] ? String(row[5]).trim() : '';
        const nextOccurrence = row[6] ? String(row[6]).trim() : '';
        
        return {
          name: name,
          found: !!(eventId || eventTitle), // Has calendar event if either ID or title exists
          nextOccurrence: nextOccurrence || null,
          calendarLink: calendarLink || null
        };
      });
    
    Logger.log('Retrieved ' + results.length + ' names with calendar data from Google Sheet');
    return results;
    
  } catch (error) {
    Logger.log('Error in getNamesWithCalendarEvents: ' + error.toString());
    throw new Error('Failed to get names with calendar events: ' + error.message);
  }
}

function saveProperty(propertyName, value, description = '') {
  Logger.log(`Saving property: ${propertyName} = ${value}`);
  
  try {
    if (!propertyName || typeof propertyName !== 'string') {
      throw new Error('Property name must be a non-empty string');
    }
    
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Find if property already exists
    const data = propertiesTab.getDataRange().getValues();
    let propertyRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === propertyName) {
        propertyRow = i + 1; // Convert to 1-based row number
        break;
      }
    }
    
    if (propertyRow > 0) {
      // Update existing property
      propertiesTab.getRange(propertyRow, 2).setValue(value);
      propertiesTab.getRange(propertyRow, 4).setValue(new Date());
    } else {
      // Add new property
      const nextRow = propertiesTab.getLastRow() + 1;
      propertiesTab.getRange(nextRow, 1).setValue(propertyName);
      propertiesTab.getRange(nextRow, 2).setValue(value);
      propertiesTab.getRange(nextRow, 3).setValue(description);
      propertiesTab.getRange(nextRow, 4).setValue(new Date());
      
      // Format the date column
      propertiesTab.getRange(nextRow, 4).setNumberFormat('MM/dd/yyyy HH:mm:ss');
    }
    
    Logger.log(`Property ${propertyName} saved successfully`);
    return {
      success: true,
      propertyName: propertyName,
      value: value
    };
    
  } catch (error) {
    Logger.log('Error in saveProperty: ' + error.toString());
    throw new Error('Failed to save property: ' + error.message);
  }
}

function getProperty(propertyName) {
  Logger.log(`Getting property: ${propertyName}`);
  
  try {
    if (!propertyName || typeof propertyName !== 'string') {
      throw new Error('Property name must be a non-empty string');
    }
    
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Check if Properties tab has any data
    const lastRow = propertiesTab.getLastRow();
    const lastCol = propertiesTab.getLastColumn();
    
    Logger.log(`Properties tab has ${lastRow} rows and ${lastCol} columns`);
    
    if (lastRow < 2) {
      Logger.log('Properties tab has no data rows, property not found');
      return {
        found: false,
        value: null,
        description: null,
        lastUpdated: null
      };
    }
    
    // Get all data and find the property
    const data = propertiesTab.getRange(1, 1, lastRow, Math.max(lastCol, 4)).getValues();
    Logger.log(`Retrieved data from Properties tab: ${data.length} rows`);
    
    // Log all properties for debugging
    for (let i = 1; i < data.length; i++) {
      Logger.log(`Row ${i + 1}: Property="${data[i][0]}", Value="${data[i][1]}"`);
      
      if (data[i][0] === propertyName) {
        const value = data[i][1];
        const description = data[i][2];
        const lastUpdated = data[i][3];
        
        Logger.log(`Found property ${propertyName} with value: ${value} (type: ${typeof value})`);
        return {
          found: true,
          value: value,
          description: description,
          lastUpdated: lastUpdated
        };
      }
    }
    
    Logger.log(`Property ${propertyName} not found in Properties tab`);
    return {
      found: false,
      value: null,
      description: null,
      lastUpdated: null
    };
    
  } catch (error) {
    Logger.log('Error in getProperty: ' + error.toString());
    throw new Error('Failed to get property: ' + error.message);
  }
}

function getRecurringInterval() {
  Logger.log('Getting recurring interval property');
  
  try {
    const result = getProperty('Recurring Interval');
    
    if (result.found) {
      const value = parseInt(result.value);
      
      // Validate range
      if (isNaN(value) || value < 1 || value > 26) {
        Logger.log(`Invalid recurring interval value: ${result.value}, using default 8`);
        return 8;
      }
      
      Logger.log(`Retrieved recurring interval: ${value}`);
      return value;
    } else {
      Logger.log('Recurring interval not found, using default 8');
      return 8;
    }
    
  } catch (error) {
    Logger.log('Error in getRecurringInterval: ' + error.toString());
    return 8; // Return default value on error
  }
}

function setRecurringInterval(weeks) {
  Logger.log(`Setting recurring interval to: ${weeks}`);
  
  try {
    // Validate input
    const value = parseInt(weeks);
    
    if (isNaN(value)) {
      throw new Error('Recurring interval must be a number');
    }
    
    if (value < 1 || value > 26) {
      throw new Error('Recurring interval must be between 1 and 26 weeks');
    }
    
    // Save the property
    const result = saveProperty(
      'Recurring Interval', 
      value, 
      'Meeting recurrence interval in weeks (1-26)'
    );
    
    Logger.log(`Recurring interval set to ${value} weeks`);
    return {
      success: true,
      value: value,
      message: `Recurring interval set to ${value} weeks`
    };
    
  } catch (error) {
    Logger.log('Error in setRecurringInterval: ' + error.toString());
    throw new Error('Failed to set recurring interval: ' + error.message);
  }
}


function calculateOptimalRecurringInterval() {
  Logger.log('Calculating optimal recurring interval based on names and meeting slots');
  
  try {
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Count total names (excluding header row)
    const namesData = namesTab.getDataRange().getValues();
    const totalNames = Math.max(0, namesData.length - 1); // Subtract 1 for header row
    
    // Count total meeting slots (excluding header row)
    const slotsData = meetingSlotsTab.getDataRange().getValues();
    const totalSlots = Math.max(0, slotsData.length - 1); // Subtract 1 for header row
    
    Logger.log(`Found ${totalNames} total names and ${totalSlots} total meeting slots`);
    
    // Calculate the optimal interval
    let calculatedValue;
    
    if (totalSlots === 0) {
      // If no meeting slots, default to 8 weeks
      calculatedValue = 8;
      Logger.log('No meeting slots found, defaulting to 8 weeks');
    } else {
      // Calculate: total names / total slots, rounded up
      calculatedValue = Math.ceil(totalNames / totalSlots);
      
      // Ensure the value is within the valid range (1-26)
      calculatedValue = Math.max(1, Math.min(26, calculatedValue));
      
      Logger.log(`Calculated value: Math.ceil(${totalNames} / ${totalSlots}) = ${calculatedValue}`);
    }
    
    return {
      success: true,
      calculatedValue: calculatedValue,
      totalNames: totalNames,
      totalSlots: totalSlots,
      formula: `Math.ceil(${totalNames} / ${totalSlots}) = ${calculatedValue}`,
      message: `Auto-calculated recurring interval: ${calculatedValue} weeks (${totalNames} names ÷ ${totalSlots} slots, rounded up)`
    };
    
  } catch (error) {
    Logger.log('Error in calculateOptimalRecurringInterval: ' + error.toString());
    throw new Error('Failed to calculate optimal recurring interval: ' + error.message);
  }
}

function removeRecurringMeetingOnly(nameToRemove) {
  Logger.log('Removing recurring meeting only for: ' + nameToRemove);
  
  try {
    if (!nameToRemove || typeof nameToRemove !== 'string') {
      throw new Error('Name must be provided as a string');
    }
    
    const trimmedName = nameToRemove.trim();
    if (trimmedName.length === 0) {
      throw new Error('Name cannot be empty');
    }
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Use 1 year range to find all recurring events
    const today = new Date();
    const oneYearAhead = new Date();
    oneYearAhead.setFullYear(today.getFullYear() + 1);
    
    // Get all events in the date range
    const allEvents = calendar.getEvents(today, oneYearAhead);
    Logger.log('Total events found for deletion search: ' + allEvents.length);
    
    // Find all Skip Level events for this specific name
    const skipLevelEvents = allEvents.filter(event => {
      const title = event.getTitle();
      return title && title.includes('Skip Level:') && title.includes(trimmedName);
    });
    Logger.log('Events with "Skip Level:" and name "' + trimmedName + '": ' + skipLevelEvents.length);
    
    let deletedSeries = 0;
    let deletedSingleEvents = 0;
    const eventSeriesMap = new Map(); // seriesId -> series object
    
    skipLevelEvents.forEach(event => {
      const title = event.getTitle();
      Logger.log('Processing event: "' + title + '"');
      
      try {
        const eventSeries = event.getEventSeries();
        if (eventSeries) {
          // This is a recurring event
          const seriesId = eventSeries.getId();
          if (!eventSeriesMap.has(seriesId)) {
            eventSeriesMap.set(seriesId, eventSeries);
          }
        } else {
          // This is a single event
          Logger.log('Deleting single event: ' + title);
          event.deleteEvent();
          deletedSingleEvents++;
        }
      } catch (error) {
        Logger.log('Error processing event "' + title + '": ' + error.toString());
      }
    });
    
    // Delete recurring event series
    eventSeriesMap.forEach((series, seriesId) => {
      try {
        Logger.log('Deleting recurring event series: ' + series.getTitle());
        series.deleteEventSeries();
        deletedSeries++;
      } catch (error) {
        Logger.log('Error deleting event series: ' + error.toString());
      }
    });
    
    Logger.log('Calendar deletion complete. Series deleted: ' + deletedSeries + ', Single events deleted: ' + deletedSingleEvents);
    
    // Now update the Google Sheet to clear calendar attributes for this name
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Find the row with this name and clear calendar attributes
    const lastRow = namesTab.getLastRow();
    if (lastRow > 1) {
      const nameRange = namesTab.getRange(2, 1, lastRow - 1, 7); // Get all data starting from row 2
      const allData = nameRange.getValues();
      
      // Find the row index for this name
      let targetRowIndex = -1;
      for (let i = 0; i < allData.length; i++) {
        if (allData[i][0] && String(allData[i][0]).trim().toLowerCase() === trimmedName.toLowerCase()) {
          targetRowIndex = i + 2; // +2 because we started from row 2 and need 1-based index
          break;
        }
      }
      
      if (targetRowIndex > 0) {
        Logger.log('Found name "' + trimmedName + '" at row ' + targetRowIndex + ', clearing calendar attributes');
        
        // Clear the calendar-related columns but keep the name
        namesTab.getRange(targetRowIndex, 3).setValue(''); // Status
        namesTab.getRange(targetRowIndex, 4).setValue(''); // Calendar Event ID
        namesTab.getRange(targetRowIndex, 5).setValue(''); // Event Title
        namesTab.getRange(targetRowIndex, 6).setValue(''); // Calendar Link
        namesTab.getRange(targetRowIndex, 7).setValue(''); // Next Occurrence
        
        Logger.log('Cleared calendar attributes for "' + trimmedName + '" in Google Sheet Names tab');
      } else {
        Logger.log('Warning: Could not find name "' + trimmedName + '" in Names tab to clear calendar attributes');
      }
    }
    
    return {
      success: true,
      nameKept: true,
      meetingsDeleted: deletedSeries + deletedSingleEvents,
      recurringSeriesDeleted: deletedSeries,
      singleEventsDeleted: deletedSingleEvents,
      message: `Calendar meetings for "${trimmedName}" have been removed. The name remains in the list.`
    };
    
  } catch (error) {
    Logger.log('Error in removeRecurringMeetingOnly: ' + error.toString());
    throw new Error('Failed to remove recurring meeting: ' + error.message);
  }
}
