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

function getSkipLevel1to1List() {
  Logger.log('Getting Skip Level 1:1 List');
  
  try {
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log('Calendar obtained: ' + calendar.getName());
    
    // Calculate date range - today to one year ahead
    const today = new Date();
    const oneYearAhead = new Date();
    oneYearAhead.setFullYear(today.getFullYear() + 1);
    
    Logger.log('Date range: ' + today.toString() + ' to ' + oneYearAhead.toString());
    
    // Get all events in the date range
    const allEvents = calendar.getEvents(today, oneYearAhead);
    Logger.log('Total events found: ' + allEvents.length);
    
    // Filter events that contain "Skip Level:" in the title (case-sensitive)
    const skipLevelEvents = allEvents.filter(event => {
      const title = event.getTitle();
      Logger.log('Checking event title: "' + title + '"');
      const hasSkipLevel = title && title.includes('Skip Level:');
      Logger.log('Contains "Skip Level:": ' + hasSkipLevel);
      return hasSkipLevel;
    });
    
    Logger.log('Events with "Skip Level:" in title: ' + skipLevelEvents.length);
    
    // Log all found events for debugging
    skipLevelEvents.forEach((event, index) => {
      Logger.log('Skip Level Event ' + (index + 1) + ': "' + event.getTitle() + '"');
    });
    
    // Get distinct recurring events only by grouping by event series ID and finding next occurrence
    const seenEventSeries = new Map(); // Use Map to store next occurrence for each series
    const currentTime = new Date();
    
    skipLevelEvents.forEach(event => {
      try {
        // Check if this is a recurring event
        const eventSeries = event.getEventSeries();
        if (eventSeries) {
          // This is a recurring event
          const seriesId = eventSeries.getId();
          const eventStartTime = event.getStartTime();
          
          // Only consider future events for "next occurrence"
          if (eventStartTime >= currentTime) {
            if (!seenEventSeries.has(seriesId)) {
              // First time seeing this series, add it
              seenEventSeries.set(seriesId, {
                title: event.getTitle(),
                isRecurring: true,
                seriesId: seriesId,
                nextOccurrence: eventStartTime.toLocaleDateString() + ' ' + eventStartTime.toLocaleTimeString(),
                description: (event.getDescription() || '').substring(0, 500),
                startTime: eventStartTime,
                id: event.getId(),
                htmlLink: 'https://calendar.google.com/calendar/u/0/r/day/' + eventStartTime.getFullYear() + '/' + (eventStartTime.getMonth() + 1) + '/' + eventStartTime.getDate()
              });
            } else {
              // We've seen this series before, check if this is a sooner next occurrence
              const existingEvent = seenEventSeries.get(seriesId);
              if (eventStartTime < existingEvent.startTime) {
                // This is a sooner next occurrence, update the record
                seenEventSeries.set(seriesId, {
                  title: event.getTitle(),
                  isRecurring: true,
                  seriesId: seriesId,
                  nextOccurrence: eventStartTime.toLocaleDateString() + ' ' + eventStartTime.toLocaleTimeString(),
                  description: (event.getDescription() || '').substring(0, 500),
                  startTime: eventStartTime,
                  id: event.getId(),
                  htmlLink: 'https://calendar.google.com/calendar/u/0/r/day/' + eventStartTime.getFullYear() + '/' + (eventStartTime.getMonth() + 1) + '/' + eventStartTime.getDate()
                });
              }
            }
          }
        }
        // Skip single events - we only want recurring event series
      } catch (error) {
        // If there's an error getting event series, skip this event
        Logger.log('Error processing event: ' + error.toString());
      }
    });
    
    // Convert recurring events to array and remove the startTime property
    const distinctEvents = [];
    seenEventSeries.forEach(eventData => {
      // Remove the startTime property before adding to results (it was just for comparison)
      delete eventData.startTime;
      distinctEvents.push(eventData);
    });
    
    Logger.log('Distinct events found: ' + distinctEvents.length);
    Logger.log('Returning distinct events: ' + JSON.stringify(distinctEvents));
    return distinctEvents;
    
  } catch (error) {
    Logger.log('Error in getSkipLevel1to1List: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    throw new Error('Failed to retrieve skip level 1:1 events: ' + error.message);
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
    
    if (cleanNames.length === 0) {
      throw new Error('No valid names provided');
    }
    
    // Save to PropertiesService for persistence
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('SKIP_LEVEL_NAMES', JSON.stringify(cleanNames));
    
    Logger.log('Successfully saved ' + cleanNames.length + ' names');
    return {
      success: true,
      count: cleanNames.length,
      names: cleanNames
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
    properties.deleteProperty('SKIP_LEVEL_NAMES');
    
    Logger.log('Successfully cleared skip level names');
    return { success: true };
    
  } catch (error) {
    Logger.log('Error in clearSkipLevelNames: ' + error.toString());
    throw new Error('Failed to clear skip level names: ' + error.message);
  }
}

function checkMissingSkipLevels() {
  Logger.log('Checking missing skip levels');
  
  try {
    // Get loaded names
    const names = getSkipLevelNames();
    if (names.length === 0) {
      throw new Error('No names have been loaded. Please use "Load Skip Level Names" first.');
    }
    
    Logger.log('Checking skip levels for ' + names.length + ' names: ' + JSON.stringify(names));
    
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Calculate date range - today to one year ahead
    const today = new Date();
    const oneYearAhead = new Date();
    oneYearAhead.setFullYear(today.getFullYear() + 1);
    
    // Get all events in the date range
    const allEvents = calendar.getEvents(today, oneYearAhead);
    Logger.log('Total events found: ' + allEvents.length);
    
    // Filter events that contain "Skip Level:" in the title
    const skipLevelEvents = allEvents.filter(event => {
      const title = event.getTitle();
      return title && title.includes('Skip Level:');
    });
    
    Logger.log('Events with "Skip Level:" in title: ' + skipLevelEvents.length);
    
    // For each name, check if there's a recurring calendar event
    const results = [];
    
    names.forEach(name => {
      Logger.log('Checking for name: "' + name + '"');
      
      // Find recurring events that contain both "Skip Level:" and the name
      const matchingEvents = skipLevelEvents.filter(event => {
        const title = event.getTitle();
        const hasName = title.includes(name);
        let isRecurring = false;
        
        try {
          const eventSeries = event.getEventSeries();
          isRecurring = eventSeries !== null;
        } catch (error) {
          // If error getting event series, it's likely a single event
          isRecurring = false;
        }
        
        Logger.log('Event "' + title + '" - has name: ' + hasName + ', is recurring: ' + isRecurring);
        return hasName && isRecurring;
      });
      
      const found = matchingEvents.length > 0;
      Logger.log('Name "' + name + '" - found: ' + found + ' (' + matchingEvents.length + ' matching events)');
      
      results.push({
        name: name,
        found: found,
        eventCount: matchingEvents.length,
        events: matchingEvents.map(event => ({
          title: event.getTitle(),
          nextOccurrence: event.getStartTime().toLocaleDateString() + ' ' + event.getStartTime().toLocaleTimeString()
        }))
      });
    });
    
    Logger.log('Check results: ' + JSON.stringify(results));
    return results;
    
  } catch (error) {
    Logger.log('Error in checkMissingSkipLevels: ' + error.toString());
    throw new Error('Failed to check missing skip levels: ' + error.message);
  }
}
