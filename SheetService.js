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