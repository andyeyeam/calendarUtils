// All backend service functions have been moved to separate service files:
// - SheetService.js: Sheet management functions (getOrCreateStateSheet)
// - CalendarService.js: Calendar operations (searchCalendarEventForName, debugCalendarEvents, checkSlotAvailability, getNextDateForDayOfWeek, parseSlotTime, parseTimeString)
// - NamesService.js: Names management (saveSkipLevelNames, getSkipLevelNames, clearSkipLevelNames, removeSkipLevel, getNamesWithCalendarEvents, removeRecurringMeetingOnly, updateNameWithCalendarDetails)
// - MeetingSlotsService.js: Meeting slots management (saveMeetingSlots, getMeetingSlots, clearMeetingSlots, getMeetingSlotsWithDates)
// - PropertiesService.js: Properties management (saveProperty, getProperty, getRecurringInterval, setRecurringInterval, calculateOptimalRecurringInterval)
// - MeetingCreationService.js: Meeting creation and deletion (createMeetingsForSpecificNames, createAllRecurringMeetings, deleteAllRecurringMeetings)
// - HtmlService.js: HTML service functions (doGet, getIndexPage)

// This main Code.js file now serves as the entry point that delegates to service functions
// All functions are available globally in Google Apps Script environment

// Note: In Google Apps Script, all .js files in the project are automatically included
// and their functions are available globally. No explicit imports are needed.