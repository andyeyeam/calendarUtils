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

function clearAllNamesAndMeetings() {
  Logger.log('Clearing all names and their calendar meetings');
  
  try {
    // Get all names with their calendar event data
    const namesWithEvents = getNamesWithCalendarEvents();
    Logger.log('Found ' + namesWithEvents.length + ' names to process');
    
    let meetingsDeleted = 0;
    let errors = [];
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Delete calendar events for each name that has an event ID
    for (const nameData of namesWithEvents) {
      if (nameData.found && nameData.name) {
        Logger.log('Processing deletion for: ' + nameData.name);
        
        try {
          // Get the event ID from the Google Sheet data
          const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
          const lastRow = namesTab.getLastRow();
          
          if (lastRow > 1) {
            const dataRange = namesTab.getRange(2, 1, lastRow - 1, 7);
            const allData = dataRange.getValues();
            
            // Find the row with this name and get the event ID
            for (let i = 0; i < allData.length; i++) {
              const rowName = allData[i][0] ? String(allData[i][0]).trim() : '';
              if (rowName.toLowerCase() === nameData.name.toLowerCase()) {
                const eventId = allData[i][3] ? String(allData[i][3]).trim() : '';
                
                if (eventId) {
                  Logger.log('Found event ID for ' + nameData.name + ': ' + eventId);
                  
                  try {
                    // Try to get the event by ID and delete it
                    const event = calendar.getEventById(eventId);
                    if (event) {
                      // Check if it's a recurring event
                      const eventSeries = event.getEventSeries();
                      if (eventSeries) {
                        Logger.log('Deleting recurring event series for: ' + nameData.name);
                        eventSeries.deleteEventSeries();
                      } else {
                        Logger.log('Deleting single event for: ' + nameData.name);
                        event.deleteEvent();
                      }
                      meetingsDeleted++;
                      Logger.log('Successfully deleted calendar event for: ' + nameData.name);
                    } else {
                      Logger.log('Event not found by ID for ' + nameData.name + ', event may have been already deleted');
                    }
                  } catch (eventError) {
                    const errorMsg = 'Failed to delete calendar event for "' + nameData.name + '": ' + eventError.message;
                    Logger.log(errorMsg);
                    errors.push(errorMsg);
                  }
                } else {
                  Logger.log('No event ID found for: ' + nameData.name);
                }
                break;
              }
            }
          }
        } catch (deleteError) {
          const errorMsg = 'Error processing deletion for "' + nameData.name + '": ' + deleteError.message;
          Logger.log(errorMsg);
          errors.push(errorMsg);
        }
      }
    }
    
    // Now clear all data from the Google Sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    const lastRow = namesTab.getLastRow();
    if (lastRow > 1) {
      namesTab.getRange(2, 1, lastRow - 1, 7).clearContent();
      Logger.log('Cleared ' + (lastRow - 1) + ' rows of data from Google Sheet');
    } else {
      Logger.log('No data to clear from Google Sheet');
    }
    
    Logger.log('Successfully cleared all names and meetings. Meetings deleted: ' + meetingsDeleted + ', Errors: ' + errors.length);
    
    return { 
      success: true, 
      count: 0,
      names: [],
      meetingsDeleted: meetingsDeleted,
      errors: errors
    };
    
  } catch (error) {
    Logger.log('Error in clearAllNamesAndMeetings: ' + error.toString());
    throw new Error('Failed to clear all names and meetings: ' + error.message);
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

function updateNameWithCalendarDetails(name, eventId, eventTitle, calendarLink, nextOccurrence) {
  Logger.log('Updating calendar details for name: ' + name);
  
  try {
    if (!name || typeof name !== 'string') {
      throw new Error('Name must be provided as a string');
    }
    
    const trimmedName = name.trim();
    if (trimmedName.length === 0) {
      throw new Error('Name cannot be empty');
    }
    
    // Get or create the state sheet
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Find the row with this name
    const lastRow = namesTab.getLastRow();
    if (lastRow <= 1) {
      throw new Error('No names found in the Names tab');
    }
    
    // Get all data from the Names tab
    const dataRange = namesTab.getRange(2, 1, lastRow - 1, 7); // Get all data starting from row 2
    const allData = dataRange.getValues();
    
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
      
      // Update the calendar-related columns
      namesTab.getRange(targetRowIndex, 3).setValue('Active'); // Status
      namesTab.getRange(targetRowIndex, 4).setValue(eventId || ''); // Calendar Event ID
      namesTab.getRange(targetRowIndex, 5).setValue(eventTitle || ''); // Event Title
      namesTab.getRange(targetRowIndex, 6).setValue(calendarLink || ''); // Calendar Link
      namesTab.getRange(targetRowIndex, 7).setValue(nextOccurrence || ''); // Next Occurrence
      
      // Auto-resize columns to fit content
      namesTab.autoResizeColumns(1, 7);
      
      Logger.log('Updated calendar details for "' + trimmedName + '" in Google Sheet Names tab');
      
      return {
        success: true,
        name: trimmedName,
        rowUpdated: targetRowIndex,
        eventId: eventId,
        eventTitle: eventTitle,
        calendarLink: calendarLink,
        nextOccurrence: nextOccurrence
      };
    } else {
      throw new Error('Name "' + trimmedName + '" not found in Names tab');
    }
    
  } catch (error) {
    Logger.log('Error in updateNameWithCalendarDetails: ' + error.toString());
    throw new Error('Failed to update name with calendar details: ' + error.message);
  }
}