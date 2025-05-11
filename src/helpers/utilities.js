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
      const [startDate, endDate] = dateRange.split('-');
      return `${formatDateWithWeekday(startDate)}-${formatDateWithWeekday(endDate)}`;
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
  // OpenAI API keys typically start with "sk-" and are 51 characters long
  return typeof apiKey === 'string' && apiKey.startsWith('sk-') && apiKey.length > 40;
}

/**
 * Safely stores a value in Office.context.roamingSettings
 * @param {string} key - The key to store
 * @param {any} value - The value to store
 */
export function storeInRoamingSettings(key, value) {
  try {
    if (Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.set(key, value);
      Office.context.roamingSettings.saveAsync();
    }
  } catch (error) {
    console.error('Error storing in roaming settings:', error);
  }
}

/**
 * Safely retrieves a value from Office.context.roamingSettings
 * @param {string} key - The key to retrieve
 * @returns {any} - The retrieved value, or null if not found
 */
export function getFromRoamingSettings(key) {
  try {
    if (Office.context && Office.context.roamingSettings) {
      return Office.context.roamingSettings.get(key);
    }
    return null;
  } catch (error) {
    console.error('Error retrieving from roaming settings:', error);
    return null;
  }
}
