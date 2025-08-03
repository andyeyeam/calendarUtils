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
    
    // Get the recurring interval from Properties tab
    const recurringIntervalResult = getProperty('Recurring Interval');
    if (!recurringIntervalResult.found) {
      const errorMessage = 'Recurring interval has not been defined in the Properties tab of the CalendarUtilities State Sheet. Please set the "Recurring Interval" property first.';
      Logger.log(errorMessage);
      throw new Error(errorMessage);
    }
    
    const recurringInterval = parseInt(recurringIntervalResult.value);
    if (isNaN(recurringInterval) || recurringInterval < 1 || recurringInterval > 26) {
      const errorMessage = 'Invalid recurring interval value: ' + recurringIntervalResult.value + '. Must be between 1 and 26 weeks.';
      Logger.log(errorMessage);
      throw new Error(errorMessage);
    }
    
    Logger.log('Using recurring interval from Properties tab: ' + recurringInterval + ' weeks');
    
    // Get all loaded names and meeting slots
    const allNames = getSkipLevelNames();
    const meetingSlots = getMeetingSlots();
    
    if (meetingSlots.length === 0) {
      throw new Error('No meeting slots configured. Please configure meeting slots first.');
    }
    
    Logger.log('Processing ' + namesToProcess.length + ' specific names with ' + meetingSlots.length + ' meeting slots');
    
    // Calculate X weeks: max of recurring interval or (total names / total meeting slots) rounded up
    // Use total names count for calculation, not just the names being processed
    const calculatedWeeks = Math.ceil(allNames.length / meetingSlots.length);
    const weeksToSearch = Math.max(recurringInterval, calculatedWeeks);
    
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
              Logger.log('Recurrence: Every ' + recurringInterval + ' weeks');
              
              // Create the recurring event
              const recurringEvent = calendar.createEventSeries(
                eventTitle,
                slotStart,
                slotEnd,
                CalendarApp.newRecurrence().addWeeklyRule().interval(recurringInterval)
              );
              
              Logger.log('Successfully created recurring event for ' + name + ' with ID: ' + recurringEvent.getId());
              
              // Get calendar event details for updating the Google Sheet
              const eventId = recurringEvent.getId();
              
              // Format date for Google Calendar day view URL: YYYY/M/D
              const year = slotStart.getFullYear();
              const month = slotStart.getMonth() + 1; // JavaScript months are 0-based
              const day = slotStart.getDate();
              const calendarLink = `https://calendar.google.com/calendar/u/0/r/day/${year}/${month}/${day}`;
              
              // Format date/time for next occurrence
              const dateStr = slotStart.toLocaleDateString();
              const timeStr = slotStart.toLocaleTimeString().replace(/\s*\([^)]*\)/, '');
              const nextOccurrence = dateStr + ' ' + timeStr;
              
              // Update the Google Sheet with calendar event details
              try {
                updateNameWithCalendarDetails(name, eventId, eventTitle, calendarLink, nextOccurrence);
                Logger.log('Updated Google Sheet with calendar details for ' + name);
              } catch (updateError) {
                Logger.log('Error updating Google Sheet for ' + name + ': ' + updateError.toString());
                // Continue with meeting creation even if sheet update fails
              }
              
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
                slotIndex: slotIndex + 1,
                eventId: eventId,
                eventTitle: eventTitle,
                calendarLink: calendarLink,
                nextOccurrence: nextOccurrence
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
      weeksCalculated: recurringInterval,
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
    // Get the recurring interval from Properties tab
    const recurringIntervalResult = getProperty('Recurring Interval');
    if (!recurringIntervalResult.found) {
      const errorMessage = 'Recurring interval has not been defined in the Properties tab of the CalendarUtilities State Sheet. Please set the "Recurring Interval" property first.';
      Logger.log(errorMessage);
      throw new Error(errorMessage);
    }
    
    const recurringInterval = parseInt(recurringIntervalResult.value);
    if (isNaN(recurringInterval) || recurringInterval < 1 || recurringInterval > 26) {
      const errorMessage = 'Invalid recurring interval value: ' + recurringIntervalResult.value + '. Must be between 1 and 26 weeks.';
      Logger.log(errorMessage);
      throw new Error(errorMessage);
    }
    
    Logger.log('Using recurring interval from Properties tab: ' + recurringInterval + ' weeks');
    
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
    
    // Calculate X weeks: max of recurring interval or (total names / total meeting slots) rounded up
    const calculatedWeeks = Math.ceil(allNames.length / meetingSlots.length);
    const weeksToSearch = Math.max(recurringInterval, calculatedWeeks);
    
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
        weeksCalculated: recurringInterval,
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
              Logger.log('Recurrence: Every ' + recurringInterval + ' weeks');
              
              // Create the recurring event
              const recurringEvent = calendar.createEventSeries(
                eventTitle,
                slotStart,
                slotEnd,
                CalendarApp.newRecurrence().addWeeklyRule().interval(recurringInterval)
              );
              
              Logger.log('Successfully created recurring event for ' + name + ' with ID: ' + recurringEvent.getId());
              
              // Get calendar event details for updating the Google Sheet
              const eventId = recurringEvent.getId();
              
              // Format date for Google Calendar day view URL: YYYY/M/D
              const year = slotStart.getFullYear();
              const month = slotStart.getMonth() + 1; // JavaScript months are 0-based
              const day = slotStart.getDate();
              const calendarLink = `https://calendar.google.com/calendar/u/0/r/day/${year}/${month}/${day}`;
              
              // Format date/time for next occurrence
              const dateStr = slotStart.toLocaleDateString();
              const timeStr = slotStart.toLocaleTimeString().replace(/\s*\([^)]*\)/, '');
              const nextOccurrence = dateStr + ' ' + timeStr;
              
              // Update the Google Sheet with calendar event details
              try {
                updateNameWithCalendarDetails(name, eventId, eventTitle, calendarLink, nextOccurrence);
                Logger.log('Updated Google Sheet with calendar details for ' + name);
              } catch (updateError) {
                Logger.log('Error updating Google Sheet for ' + name + ': ' + updateError.toString());
                // Continue with meeting creation even if sheet update fails
              }
              
              meetingsCreated++;
              meetingCreatedForName = true;
              
              createdMeetings.push({
                name: name,
                dayOfWeek: slot.dayOfWeek,
                time: slot.time,
                duration: slot.duration,
                weekFound: weekCount + 1,
                eventId: eventId,
                eventTitle: eventTitle,
                calendarLink: calendarLink,
                nextOccurrence: nextOccurrence
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
      weeksCalculated: recurringInterval,
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