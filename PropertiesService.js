function saveProperty(propertyName, value, description = '') {
  Logger.log(`Saving property: ${propertyName} = ${value}`);
  
  try {
    if (!propertyName || typeof propertyName !== 'string') {
      throw new Error('Property name must be a non-empty string');
    }
    
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Find if property already exists
    const data = propertiesTab.getDataRange().getValues();
    let propertyRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === propertyName) {
        propertyRow = i + 1; // Convert to 1-based row number
        break;
      }
    }
    
    if (propertyRow > 0) {
      // Update existing property
      propertiesTab.getRange(propertyRow, 2).setValue(value);
      propertiesTab.getRange(propertyRow, 4).setValue(new Date());
    } else {
      // Add new property
      const nextRow = propertiesTab.getLastRow() + 1;
      propertiesTab.getRange(nextRow, 1).setValue(propertyName);
      propertiesTab.getRange(nextRow, 2).setValue(value);
      propertiesTab.getRange(nextRow, 3).setValue(description);
      propertiesTab.getRange(nextRow, 4).setValue(new Date());
      
      // Format the date column
      propertiesTab.getRange(nextRow, 4).setNumberFormat('MM/dd/yyyy HH:mm:ss');
    }
    
    Logger.log(`Property ${propertyName} saved successfully`);
    return {
      success: true,
      propertyName: propertyName,
      value: value
    };
    
  } catch (error) {
    Logger.log('Error in saveProperty: ' + error.toString());
    throw new Error('Failed to save property: ' + error.message);
  }
}

function getProperty(propertyName) {
  Logger.log(`Getting property: ${propertyName}`);
  
  try {
    if (!propertyName || typeof propertyName !== 'string') {
      throw new Error('Property name must be a non-empty string');
    }
    
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Check if Properties tab has any data
    const lastRow = propertiesTab.getLastRow();
    const lastCol = propertiesTab.getLastColumn();
    
    Logger.log(`Properties tab has ${lastRow} rows and ${lastCol} columns`);
    
    if (lastRow < 2) {
      Logger.log('Properties tab has no data rows, property not found');
      return {
        found: false,
        value: null,
        description: null,
        lastUpdated: null
      };
    }
    
    // Get all data and find the property
    const data = propertiesTab.getRange(1, 1, lastRow, Math.max(lastCol, 4)).getValues();
    Logger.log(`Retrieved data from Properties tab: ${data.length} rows`);
    
    // Log all properties for debugging
    for (let i = 1; i < data.length; i++) {
      Logger.log(`Row ${i + 1}: Property="${data[i][0]}", Value="${data[i][1]}"`);
      
      if (data[i][0] === propertyName) {
        const value = data[i][1];
        const description = data[i][2];
        const lastUpdated = data[i][3];
        
        Logger.log(`Found property ${propertyName} with value: ${value} (type: ${typeof value})`);
        return {
          found: true,
          value: value,
          description: description,
          lastUpdated: lastUpdated
        };
      }
    }
    
    Logger.log(`Property ${propertyName} not found in Properties tab`);
    return {
      found: false,
      value: null,
      description: null,
      lastUpdated: null
    };
    
  } catch (error) {
    Logger.log('Error in getProperty: ' + error.toString());
    throw new Error('Failed to get property: ' + error.message);
  }
}

function getRecurringInterval() {
  Logger.log('Getting recurring interval property');
  
  try {
    const result = getProperty('Recurring Interval');
    
    if (result.found) {
      const value = parseInt(result.value);
      
      // Validate range
      if (isNaN(value) || value < 1 || value > 26) {
        Logger.log(`Invalid recurring interval value: ${result.value}, using default 8`);
        return 8;
      }
      
      Logger.log(`Retrieved recurring interval: ${value}`);
      return value;
    } else {
      Logger.log('Recurring interval not found, using default 8');
      return 8;
    }
    
  } catch (error) {
    Logger.log('Error in getRecurringInterval: ' + error.toString());
    return 8; // Return default value on error
  }
}

function setRecurringInterval(weeks) {
  Logger.log(`Setting recurring interval to: ${weeks}`);
  
  try {
    // Validate input
    const value = parseInt(weeks);
    
    if (isNaN(value)) {
      throw new Error('Recurring interval must be a number');
    }
    
    if (value < 1 || value > 26) {
      throw new Error('Recurring interval must be between 1 and 26 weeks');
    }
    
    // Save the property
    const result = saveProperty(
      'Recurring Interval', 
      value, 
      'Meeting recurrence interval in weeks (1-26)'
    );
    
    Logger.log(`Recurring interval set to ${value} weeks`);
    return {
      success: true,
      value: value,
      message: `Recurring interval set to ${value} weeks`
    };
    
  } catch (error) {
    Logger.log('Error in setRecurringInterval: ' + error.toString());
    throw new Error('Failed to set recurring interval: ' + error.message);
  }
}

function calculateOptimalRecurringInterval() {
  Logger.log('Calculating optimal recurring interval based on names and meeting slots');
  
  try {
    const { sheet, namesTab, meetingSlotsTab, propertiesTab } = getOrCreateStateSheet();
    
    // Count total names (excluding header row)
    const namesData = namesTab.getDataRange().getValues();
    const totalNames = Math.max(0, namesData.length - 1); // Subtract 1 for header row
    
    // Count total meeting slots (excluding header row)
    const slotsData = meetingSlotsTab.getDataRange().getValues();
    const totalSlots = Math.max(0, slotsData.length - 1); // Subtract 1 for header row
    
    Logger.log(`Found ${totalNames} total names and ${totalSlots} total meeting slots`);
    
    // Calculate the optimal interval
    let calculatedValue;
    
    if (totalSlots === 0) {
      // If no meeting slots, default to 8 weeks
      calculatedValue = 8;
      Logger.log('No meeting slots found, defaulting to 8 weeks');
    } else {
      // Calculate: total names / total slots, rounded up
      calculatedValue = Math.ceil(totalNames / totalSlots);
      
      // Ensure the value is within the valid range (1-26)
      calculatedValue = Math.max(1, Math.min(26, calculatedValue));
      
      Logger.log(`Calculated value: Math.ceil(${totalNames} / ${totalSlots}) = ${calculatedValue}`);
    }
    
    return {
      success: true,
      calculatedValue: calculatedValue,
      totalNames: totalNames,
      totalSlots: totalSlots,
      formula: `Math.ceil(${totalNames} / ${totalSlots}) = ${calculatedValue}`,
      message: `Auto-calculated recurring interval: ${calculatedValue} weeks (${totalNames} names ÷ ${totalSlots} slots, rounded up)`
    };
    
  } catch (error) {
    Logger.log('Error in calculateOptimalRecurringInterval: ' + error.toString());
    throw new Error('Failed to calculate optimal recurring interval: ' + error.message);
  }
}