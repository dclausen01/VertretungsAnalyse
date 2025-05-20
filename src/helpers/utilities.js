/*
 * Utility functions for the Vertretungsanalyse add-in
 */

/**
 * Formats a date string to include the weekday abbreviation
 * @param {string} dateString - The date string to format (e.g., "12.05.2025")
 * @returns {string} - The formatted date string with weekday abbreviation (e.g., "12.05.2025 (Mo)")
 */
export function formatDateWithWeekday(dateString) {
  try {
    // Parse the date string (assuming German format DD.MM.YYYY)
    const parts = dateString.split('.');
    if (parts.length !== 3) {
      return dateString; // Return original if not in expected format
    }
    
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // JavaScript months are 0-indexed
    const year = parseInt(parts[2], 10);
    
    const date = new Date(year, month, day);
    
    // Validate that the date is valid
    if (isNaN(date.getTime())) {
      console.warn(`Invalid date: ${dateString}`);
      return dateString; // Return original if date is invalid
    }
    
    // Get the weekday abbreviation in German
    const weekdays = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
    const weekday = weekdays[date.getDay()];
    
    // Return the formatted date
    return `${dateString} (${weekday})`;
  } catch (error) {
    console.error('Error formatting date:', error);
    return dateString; // Return original on error
  }
}

/**
 * Formats a date range to include weekday abbreviations
 * @param {string} dateRange - The date range string (e.g., "12.05.2025-14.05.2025")
 * @returns {string} - The formatted date range with weekday abbreviations
 */
export function formatDateRangeWithWeekdays(dateRange) {
  try {
    // Check if it's a date range (contains a hyphen)
    if (dateRange.includes('-')) {
      // Validate the format
      const parts = dateRange.split('-');
      if (parts.length !== 2) {
        console.warn(`Invalid date range format: ${dateRange}`);
        return dateRange; // Return original if format is invalid
      }
      
      const [startDate, endDate] = parts;
      // Trim any whitespace that might be present
      const trimmedStartDate = startDate.trim();
      const trimmedEndDate = endDate.trim();
      
      // Format both dates
      return `${formatDateWithWeekday(trimmedStartDate)}-${formatDateWithWeekday(trimmedEndDate)}`;
    }
    
    // If it's not a range, just format the single date
    return formatDateWithWeekday(dateRange);
  } catch (error) {
    console.error('Error formatting date range:', error);
    return dateRange; // Return original on error
  }
}

/**
 * Validates an OpenAI API key format
 * @param {string} apiKey - The API key to validate
 * @returns {boolean} - Whether the API key format is valid
 */
export function isValidApiKeyFormat(apiKey) {
  // Check if apiKey is a string
  if (typeof apiKey !== 'string') {
    return false;
  }
  
  // Remove any whitespace
  const trimmedKey = apiKey.trim();
  
  // OpenAI API keys typically start with "sk-" and are 51 characters long
  // They contain alphanumeric characters and hyphens
  const openAIKeyRegex = /^sk-[A-Za-z0-9]{48}$/;
  
  // Check if it's a valid OpenAI API key format
  if (trimmedKey.startsWith('sk-') && trimmedKey.length >= 48) {
    return true;
  }
  
  return false;
}

/**
 * Safely stores a value in Office.context.roamingSettings
 * @param {string} key - The key to store
 * @param {any} value - The value to store
 * @returns {Promise<boolean>} - Whether the operation was successful
 */
export function storeInRoamingSettings(key, value) {
  return new Promise((resolve) => {
    try {
      if (!Office.context || !Office.context.roamingSettings) {
        console.warn('RoamingSettings not available');
        resolve(false);
        return;
      }
      
      // Set the value
      Office.context.roamingSettings.set(key, value);
      
      // Create a timeout to handle cases where saveAsync might hang
      const timeoutId = setTimeout(() => {
        console.warn('RoamingSettings saveAsync timed out');
        resolve(false);
      }, 5000); // 5 second timeout
      
      // Save the settings
      Office.context.roamingSettings.saveAsync((result) => {
        clearTimeout(timeoutId);
        
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(true);
        } else {
          console.error('Error saving to roaming settings:', result.error.message);
          resolve(false);
        }
      });
    } catch (error) {
      console.error('Error storing in roaming settings:', error);
      resolve(false);
    }
  });
}

/**
 * Safely retrieves a value from Office.context.roamingSettings
 * @param {string} key - The key to retrieve
 * @returns {any} - The retrieved value, or null if not found
 */
export function getFromRoamingSettings(key) {
  try {
    // Check if roamingSettings is available
    if (!Office.context || !Office.context.roamingSettings) {
      console.warn('RoamingSettings not available');
      return null;
    }
    
    // Get the value
    const value = Office.context.roamingSettings.get(key);
    
    // Check if the value exists
    if (value === undefined) {
      return null;
    }
    
    return value;
  } catch (error) {
    console.error('Error retrieving from roaming settings:', error);
    return null;
  }
}
