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
    
    // Pre-process events: group recurring events by series ID and find next occurrence
    const recurringEventSeries = new Map(); // seriesId -> {title, nextOccurrence}
    const currentTime = new Date();
    
    allEvents.forEach(event => {
      const title = event.getTitle();
      
      // Only process events with "Skip Level:" in title
      if (!title || !title.includes('Skip Level:')) {
        return;
      }
      
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
                startTime: eventStartTime
              });
            } else {
              // Check if this is a sooner next occurrence
              const existingEvent = recurringEventSeries.get(seriesId);
              if (eventStartTime < existingEvent.startTime) {
                recurringEventSeries.set(seriesId, {
                  title: title,
                  nextOccurrence: eventStartTime.toLocaleDateString() + ' ' + eventStartTime.toLocaleTimeString(),
                  startTime: eventStartTime
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
      
      // Search through the pre-processed recurring events
      for (const [seriesId, eventData] of recurringEventSeries) {
        if (eventData.title.includes(name)) {
          found = true;
          nextOccurrence = eventData.nextOccurrence;
          Logger.log('Found match for "' + name + '" in series: ' + eventData.title);
          break; // We only need to find one match
        }
      }
      
      Logger.log('Name "' + name + '" - found: ' + found);
      
      results.push({
        name: name,
        found: found,
        eventCount: found ? 1 : 0,
        events: found ? [{
          title: 'Recurring calendar events found',
          nextOccurrence: nextOccurrence
        }] : []
      });
    });
    
    Logger.log('Check completed for ' + names.length + ' names');
    return results;
    
  } catch (error) {
    Logger.log('Error in checkMissingSkipLevels: ' + error.toString());
    throw new Error('Failed to check missing skip levels: ' + error.message);
  }
}
