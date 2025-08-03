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

function checkSlotAvailability(startTime, endTime) {
  try {
    Logger.log(`Checking availability for ${startTime.toString()} - ${endTime.toString()}`);
    
    // Get user's primary calendar
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Check for existing events in this time slot
    const existingEvents = calendar.getEvents(startTime, endTime);
    
    const isAvailable = existingEvents.length === 0;
    Logger.log(`Slot availability check: ${isAvailable ? 'AVAILABLE' : 'BUSY'} (${existingEvents.length} conflicting events)`);
    
    return isAvailable;
    
  } catch (error) {
    Logger.log(`Error checking slot availability: ${error.toString()}`);
    return false; // Default to not available if there's an error
  }
}

function getNextDateForDayOfWeek(startDate, dayOfWeek) {
  const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const targetDayIndex = daysOfWeek.indexOf(dayOfWeek);
  
  if (targetDayIndex === -1) {
    Logger.log(`Invalid day of week: ${dayOfWeek}`);
    return null;
  }
  
  const currentDayIndex = startDate.getDay();
  let daysToAdd = targetDayIndex - currentDayIndex;
  
  if (daysToAdd < 0) {
    daysToAdd += 7; // Next week
  }
  
  const targetDate = new Date(startDate);
  targetDate.setDate(startDate.getDate() + daysToAdd);
  
  return targetDate;
}

function parseTimeString(timeStr) {
  try {
    Logger.log(`Parsing time string: ${timeStr}`);
    
    if (!timeStr || !timeStr.includes(':')) {
      Logger.log(`Invalid time format: ${timeStr}`);
      return null;
    }
    
    const timeParts = timeStr.split(':');
    let hours = parseInt(timeParts[0]);
    const minutesPart = timeParts[1].toLowerCase();
    
    // Extract minutes (remove any non-digit characters like AM/PM)
    const minutes = parseInt(minutesPart.replace(/[^\d]/g, ''));
    
    // Handle 12-hour format (with AM/PM)
    if (minutesPart.includes('pm') && hours !== 12) {
      hours += 12;
    } else if (minutesPart.includes('am') && hours === 12) {
      hours = 0;
    }
    // For 24-hour format (no AM/PM), hours stay as-is
    
    // Validate hours and minutes
    if (isNaN(hours) || isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
      Logger.log(`Invalid time values: hours=${hours}, minutes=${minutes}`);
      return null;
    }
    
    Logger.log(`Parsed time: ${hours}:${minutes.toString().padStart(2, '0')} (24-hour format)`);
    return { hours: hours, minutes: minutes };
    
  } catch (error) {
    Logger.log(`Error parsing time string: ${error.toString()}`);
    return null;
  }
}

function parseSlotTime(date, timeStr, duration) {
  try {
    Logger.log(`Parsing slot time: ${timeStr} on ${date.toDateString()} for ${duration} minutes`);
    
    const timeResult = parseTimeString(timeStr);
    if (!timeResult) {
      return null;
    }
    
    // Create start datetime
    const startDateTime = new Date(date);
    startDateTime.setHours(timeResult.hours, timeResult.minutes, 0, 0);
    
    // Parse duration and create end datetime
    const durationMinutes = parseInt(duration);
    if (isNaN(durationMinutes)) {
      Logger.log(`Invalid duration: ${duration}`);
      return null;
    }
    
    const endDateTime = new Date(startDateTime.getTime() + (durationMinutes * 60 * 1000));
    
    Logger.log(`Parsed slot: ${startDateTime.toString()} to ${endDateTime.toString()}`);
    
    return {
      start: startDateTime,
      end: endDateTime
    };
    
  } catch (error) {
    Logger.log(`Error parsing slot time: ${error.toString()}`);
    return null;
  }
}