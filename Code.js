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
                description: event.getDescription() || '',
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
                  description: event.getDescription() || '',
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
