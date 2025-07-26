function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Andy\'s Calendar Utilities')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAvailabilityFinderPage() {
  return HtmlService.createHtmlOutputFromFile('availability-finder').getContent();
}

function getIndexPage() {
  return HtmlService.createHtmlOutputFromFile('index').getContent();
}

function getCalendarEvents(startDate, endDate) {
  Logger.log('Getting calendar events from ' + startDate + ' to ' + endDate);
  
  try {
    // Get the default calendar
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log('Calendar obtained: ' + calendar.getName());
    
    // Parse the date strings more reliably
    const start = new Date(startDate + 'T00:00:00');
    const end = new Date(endDate + 'T23:59:59');
    
    Logger.log('Parsed start date: ' + start.toString());
    Logger.log('Parsed end date: ' + end.toString());
    
    // Get events in the date range
    const events = calendar.getEvents(start, end);
    Logger.log('Raw events count: ' + events.length);
    
    // Format events for display
    const eventList = events.map(event => {
      const eventData = {
        title: event.getTitle() || 'Untitled Event',
        start: event.getStartTime().toLocaleDateString() + ' ' + event.getStartTime().toLocaleTimeString(),
        end: event.getEndTime().toLocaleDateString() + ' ' + event.getEndTime().toLocaleTimeString(),
        description: event.getDescription() || ''
      };
      Logger.log('Event: ' + eventData.title + ' at ' + eventData.start);
      return eventData;
    });
    
    Logger.log('Found ' + eventList.length + ' formatted events');
    Logger.log('Returning event list: ' + JSON.stringify(eventList));
    return eventList;
    
  } catch (error) {
    Logger.log('Error in getCalendarEvents: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    throw new Error('Failed to retrieve calendar events: ' + error.message);
  }
}

function onSkipLevelManagerClick() {
  Logger.log('Skip Level Manager button clicked');
  // Add your skip level manager logic here
  return 'Skip Level Manager functionality will be implemented here';
}

function testConnection() {
  Logger.log('Test connection function called');
  return 'Connection working!';
}
