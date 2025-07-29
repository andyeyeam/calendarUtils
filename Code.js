function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Andy\'s Calendar Utilities')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  Logger.log('Saving skip level names: ' + JSON.stringify(names));
  
  try {
    if (!Array.isArray(names)) {
      throw new Error('Names must be provided as an array');
    }
    
    // Clean and validate names
    const cleanNames = names
      .map(name => String(name).trim())
      .filter(name => name.length > 0);
    
    // Allow empty arrays (when all names are removed)
    if (cleanNames.length === 0) {
      Logger.log('Saving empty names list');
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('SKIP_LEVEL_NAMES', JSON.stringify([]));
      
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
    
    // Save to PropertiesService for persistence
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('SKIP_LEVEL_NAMES', JSON.stringify(uniqueNames));
    
    Logger.log('Successfully saved ' + uniqueNames.length + ' unique names');
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
  Logger.log('Getting skip level names');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    const savedNames = properties.getProperty('SKIP_LEVEL_NAMES');
    
    if (!savedNames) {
      Logger.log('No saved names found');
      return [];
    }
    
    const names = JSON.parse(savedNames);
    Logger.log('Retrieved ' + names.length + ' names: ' + JSON.stringify(names));
    return names;
    
  } catch (error) {
    Logger.log('Error in getSkipLevelNames: ' + error.toString());
    throw new Error('Failed to retrieve skip level names: ' + error.message);
  }
}

function clearSkipLevelNames() {
  Logger.log('Clearing skip level names');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    // Save an empty array instead of deleting the property
    properties.setProperty('SKIP_LEVEL_NAMES', JSON.stringify([]));
    
    Logger.log('Successfully cleared skip level names - saved empty array');
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
  Logger.log('Saving meeting slots: ' + JSON.stringify(meetingSlots));
  
  try {
    if (!Array.isArray(meetingSlots)) {
      throw new Error('Meeting slots must be provided as an array');
    }
    
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
    
    // Save to PropertiesService for persistence
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('MEETING_SLOTS', JSON.stringify(validSlots));
    
    Logger.log('Successfully saved ' + validSlots.length + ' meeting slots');
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
  Logger.log('Getting meeting slots');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    const savedSlots = properties.getProperty('MEETING_SLOTS');
    
    if (!savedSlots) {
      Logger.log('No saved meeting slots found');
      return [];
    }
    
    const slots = JSON.parse(savedSlots);
    Logger.log('Retrieved ' + slots.length + ' meeting slots: ' + JSON.stringify(slots));
    return slots;
    
  } catch (error) {
    Logger.log('Error in getMeetingSlots: ' + error.toString());
    throw new Error('Failed to retrieve meeting slots: ' + error.message);
  }
}

function clearMeetingSlots() {
  Logger.log('Clearing meeting slots');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('MEETING_SLOTS');
    
    Logger.log('Successfully cleared meeting slots');
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
    
    // Save the updated names list
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('SKIP_LEVEL_NAMES', JSON.stringify(updatedNames));
    
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
    
    if (meetingSlots.length === 0) {
      throw new Error('No meeting slots configured. Please configure meeting slots first.');
    }
    
    // Calculate X weeks: max of 8 or (total names / total meeting slots) rounded up
    const calculatedWeeks = Math.ceil(allNames.length / meetingSlots.length);
    const weeksToSearch = Math.max(8, calculatedWeeks);
    
    Logger.log('Calculated weeks based on names/slots: ' + calculatedWeeks);
    Logger.log('Final weeks to search (max of 8 or calculated): ' + weeksToSearch);
    
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
    
    // Search for first available slot by checking all meeting slots in each week
    for (let weekCount = 0; weekCount < weeksToSearch; weekCount++) {
      Logger.log('=== Checking Week ' + (weekCount + 1) + ' of ' + weeksToSearch + ' ===');
      
      // Try all meeting slots in this week
      for (const slot of meetingSlots) {
        Logger.log('Checking slot in week ' + (weekCount + 1) + ': ' + JSON.stringify(slot));
        
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
          Logger.log('Skipping past slot: ' + slotStart.toString());
          continue;
        }
        
        // Check if this slot is beyond our search range
        if (slotStart > endDate) {
          Logger.log('Slot is beyond search range: ' + slotStart.toString());
          continue;
        }
        
        Logger.log('Checking availability for: ' + slot.dayOfWeek + ' ' + slotStart.toDateString() + ' at ' + slot.time + ' (Week ' + (weekCount + 1) + ')');
        
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
          Logger.log('Available slot found: ' + slot.dayOfWeek + ' ' + currentDate.toDateString() + ' at ' + slot.time + ' (Week ' + (weekCount + 1) + ' of ' + weeksToSearch + ')');
          
          const eventTitle = 'Skip Level: ' + trimmedName;
          
          Logger.log('Creating recurring event: ' + eventTitle);
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
          
          Logger.log('Successfully created recurring event series with ID: ' + recurringEvent.getId());
          
          return {
            success: true,
            meetingCreated: true,
            weeksCalculated: weeksToSearch,
            meetingDetails: {
              dayOfWeek: slot.dayOfWeek,
              time: slot.time,
              duration: slot.duration,
              startDateTime: slotStart.toISOString(),
              endDateTime: slotEnd.toISOString(),
              title: eventTitle,
              weekFound: weekCount + 1
            },
            recurringEventId: recurringEvent.getId()
          };
        } else {
          Logger.log('Conflict found for ' + slot.dayOfWeek + ' at ' + slot.time + ' in week ' + (weekCount + 1) + ', trying next slot in this week...');
        }
      }
      
      Logger.log('No available slots found in week ' + (weekCount + 1) + ', moving to next week...');
    }
    
    // No available slots found
    Logger.log('No available slots found in any meeting slot configuration');
    
    return {
      success: true,
      meetingCreated: false,
      weeksCalculated: weeksToSearch,
      message: 'No available slots found in the next ' + weeksToSearch + ' weeks'
    };
    
  } catch (error) {
    Logger.log('Error in createRecurringMeeting: ' + error.toString());
    throw new Error('Failed to create recurring meeting: ' + error.message);
  }
}

function getNamesWithCalendarEvents() {
  Logger.log('Getting names with calendar events');
  
  try {
    // Get loaded names
    const names = getSkipLevelNames();
    if (names.length === 0) {
      Logger.log('No names loaded, returning empty array');
      return [];
    }
    
    Logger.log('Checking calendar events for ' + names.length + ' names: ' + JSON.stringify(names));
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // For recurring events, we need a longer range to ensure we find the series
    // Use 6 months to balance performance with event discovery
    const today = new Date();
    const sixMonthsAhead = new Date();
    sixMonthsAhead.setMonth(today.getMonth() + 6);
    
    // Get events in date range
    const allEvents = calendar.getEvents(today, sixMonthsAhead);
    Logger.log('Total events found in 6-month range: ' + allEvents.length);
    
    // First, let's see all event titles to debug
    const skipLevelEvents = allEvents.filter(event => {
      const title = event.getTitle();
      return title && title.includes('Skip Level:');
    });
    Logger.log('Events with "Skip Level:" in title: ' + skipLevelEvents.length);
    
    skipLevelEvents.forEach((event, index) => {
      Logger.log('Skip Level Event ' + (index + 1) + ': "' + event.getTitle() + '"');
    });
    
    // Pre-process events: group recurring events by series ID and find next occurrence
    const recurringEventSeries = new Map(); // seriesId -> {title, nextOccurrence, calendarLink}
    const currentTime = new Date();
    
    allEvents.forEach(event => {
      const title = event.getTitle();
      
      // Only process events with "Skip Level:" in title
      if (!title || !title.includes('Skip Level:')) {
        return;
      }
      
      Logger.log('Processing Skip Level event: "' + title + '"');
      
      try {
        const eventSeries = event.getEventSeries();
        if (eventSeries) {
          // This is a recurring event
          const seriesId = eventSeries.getId();
          const eventStartTime = event.getStartTime();
          
          // Only consider future events for "next occurrence"
          if (eventStartTime >= currentTime) {
            if (!recurringEventSeries.has(seriesId)) {
              // First time seeing this series
              recurringEventSeries.set(seriesId, {
                title: title,
                nextOccurrence: eventStartTime.toLocaleDateString() + ' ' + eventStartTime.toLocaleTimeString(),
                startTime: eventStartTime,
                calendarLink: 'https://calendar.google.com/calendar/u/0/r/day/' + eventStartTime.getFullYear() + '/' + (eventStartTime.getMonth() + 1) + '/' + eventStartTime.getDate()
              });
            } else {
              // Check if this is a sooner next occurrence
              const existingEvent = recurringEventSeries.get(seriesId);
              if (eventStartTime < existingEvent.startTime) {
                recurringEventSeries.set(seriesId, {
                  title: title,
                  nextOccurrence: eventStartTime.toLocaleDateString() + ' ' + eventStartTime.toLocaleTimeString(),
                  startTime: eventStartTime,
                  calendarLink: 'https://calendar.google.com/calendar/u/0/r/day/' + eventStartTime.getFullYear() + '/' + (eventStartTime.getMonth() + 1) + '/' + eventStartTime.getDate()
                });
              }
            }
          }
        }
      } catch (error) {
        // Skip if error getting event series
      }
    });
    
    Logger.log('Found ' + recurringEventSeries.size + ' distinct recurring skip level event series');
    
    // Now check each name against the pre-processed recurring events
    const results = [];
    
    names.forEach(name => {
      Logger.log('Checking for name: "' + name + '"');
      
      let found = false;
      let nextOccurrence = null;
      let calendarLink = null;
      
      // Search through the pre-processed recurring events
      for (const [seriesId, eventData] of recurringEventSeries) {
        Logger.log('Comparing name "' + name + '" with event title: "' + eventData.title + '"');
        
        // Check if the event title contains "Skip Level:" and the name
        if (eventData.title.includes('Skip Level:') && eventData.title.includes(name)) {
          found = true;
          nextOccurrence = eventData.nextOccurrence;
          calendarLink = eventData.calendarLink;
          Logger.log('✓ Found match for "' + name + '" in series: ' + eventData.title);
          break; // We only need to find one match
        } else {
          Logger.log('✗ No match for "' + name + '" in title: ' + eventData.title);
        }
      }
      
      Logger.log('Name "' + name + '" - found: ' + found);
      
      results.push({
        name: name,
        found: found,
        nextOccurrence: nextOccurrence,
        calendarLink: calendarLink
      });
    });
    
    Logger.log('Calendar check completed for ' + names.length + ' names');
    return results;
    
  } catch (error) {
    Logger.log('Error in getNamesWithCalendarEvents: ' + error.toString());
    throw new Error('Failed to get names with calendar events: ' + error.message);
  }
}
