/**
 * Server-side utility functions for the Interactive Training Projects Web App
 * Used by Google Apps Script server code
 */

if (typeof LOG_LEVELS === 'undefined') {
  const LOG_LEVELS = {
    DEBUG: 0,
    INFO: 1,
    WARN: 2,
    ERROR: 3
  };
}

// Current log level (adjust as needed)
const CURRENT_LOG_LEVEL = LOG_LEVELS.INFO;

/**
 * Generates a UUID (version 4) for unique identifiers
 * 
 * @return {string} A UUID string
 */
function generateUUID() {
  return Utilities.getUuid();
}

/**
 * Formats a date object to a human-readable string
 * 
 * @param {Date|number} date - Date object or timestamp
 * @param {string} format - Optional format string (default: MM/dd/yyyy HH:mm:ss)
 * @return {string} Formatted date string
 */
function formatDate(date, format) {
  if (!date) return "";
  
  const dateObj = typeof date === "number" ? new Date(date) : date;
  format = format || "MM/dd/yyyy HH:mm:ss";
  
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), format);
}

/**
 * Converts a date to a timestamp (milliseconds since epoch)
 * 
 * @param {Date|string} date - Date object or string
 * @return {number} Timestamp in milliseconds
 */
function toTimestamp(date) {
  if (!date) return null;
  
  if (typeof date === "string") {
    date = new Date(date);
  }
  
  return date.getTime();
}

/**
 * Validates if a string is a valid UUID
 * 
 * @param {string} uuid - String to validate
 * @return {boolean} True if string is a valid UUID
 */
function isValidUUID(uuid) {
  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  return uuidRegex.test(uuid);
}

/**
 * Sanitizes a string for use as a sheet name
 * Removes/replaces invalid characters and ensures length constraints
 * 
 * @param {string} name - String to sanitize
 * @return {string} Sanitized string
 */
function sanitizeSheetName(name) {
  if (!name) return "";
  
  // Replace invalid characters with underscores
  let sanitized = name.replace(/[\\\/\*\[\]\?\:]/g, "_");
  
  // Enforce 31 character limit for sheet names
  if (sanitized.length > 31) {
    sanitized = sanitized.substring(0, 28) + "...";
  }
  
  return sanitized;
}

/**
 * Clones an object deeply
 * 
 * @param {Object} obj - Object to clone
 * @return {Object} Cloned object
 */
function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/**
 * Merges two objects, with values from the second object overriding the first
 * 
 * @param {Object} target - Target object
 * @param {Object} source - Source object whose properties will override target
 * @return {Object} Merged object
 */
function mergeObjects(target, source) {
  const result = deepClone(target);
  
  for (const key in source) {
    if (source.hasOwnProperty(key)) {
      if (typeof source[key] === 'object' && source[key] !== null && 
          typeof result[key] === 'object' && result[key] !== null) {
        result[key] = mergeObjects(result[key], source[key]);
      } else {
        result[key] = source[key];
      }
    }
  }
  
  return result;
}

/**
 * Converts column letter to column index (0-based)
 * 
 * @param {string} columnLetter - Column letter (A-Z, AA-ZZ, etc.)
 * @return {number} Column index (0-based)
 */
function columnToIndex(columnLetter) {
  let column = 0;
  const length = columnLetter.length;
  
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  
  return column - 1; // Convert to 0-based index
}

/**
 * Converts column index (0-based) to column letter
 * 
 * @param {number} columnIndex - Column index (0-based)
 * @return {string} Column letter (A-Z, AA-ZZ, etc.)
 */
function indexToColumn(columnIndex) {
  let temp = columnIndex + 1; // Convert to 1-based index
  let columnLetter = "";
  
  while (temp > 0) {
    const remainder = temp % 26 || 26;
    columnLetter = String.fromCharCode(64 + remainder) + columnLetter;
    temp = Math.floor((temp - 1) / 26);
  }
  
  return columnLetter;
}

/**
 * Validates an email address
 * 
 * @param {string} email - Email address to validate
 * @return {boolean} True if email is valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Checks if a value is empty (null, undefined, empty string, or empty array/object)
 * 
 * @param {*} value - Value to check
 * @return {boolean} True if value is empty
 */
function isEmpty(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === "string" && value.trim() === "") return true;
  if (Array.isArray(value) && value.length === 0) return true;
  if (typeof value === "object" && Object.keys(value).length === 0) return true;
  
  return false;
}

/**
 * Truncates a string to a specified length and adds ellipsis if truncated
 * 
 * @param {string} str - String to truncate
 * @param {number} maxLength - Maximum length
 * @return {string} Truncated string
 */
function truncateString(str, maxLength) {
  if (!str || str.length <= maxLength) return str;
  return str.substring(0, maxLength - 3) + "...";
}

/**
 * Extracts YouTube video ID from various YouTube URL formats
 * 
 * @param {string} url - YouTube URL
 * @return {string|null} YouTube video ID or null if invalid
 */
function extractYouTubeVideoId(url) {
  if (!url) return null;
  
  // Regular expressions for different YouTube URL formats
  const patterns = [
    /(?:youtube\.com\/(?:[^\/]+\/.+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^"&?\/\s]{11})/i,
    /^([^"&?\/\s]{11})$/i
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return null;
}

/**
 * Logs a message with timestamp and appropriate level
 * 
 * @param {string} message - Message to log
 * @param {number} level - Log level (from LOG_LEVELS)
 */
function log(message, level = LOG_LEVELS.INFO) {
  if (level < CURRENT_LOG_LEVEL) return;
  
  const timestamp = formatDate(new Date(), "yyyy-MM-dd HH:mm:ss");
  let levelStr = "";
  
  switch(level) {
    case LOG_LEVELS.DEBUG:
      levelStr = "DEBUG";
      break;
    case LOG_LEVELS.INFO:
      levelStr = "INFO";
      break;
    case LOG_LEVELS.WARN:
      levelStr = "WARN";
      break;
    case LOG_LEVELS.ERROR:
      levelStr = "ERROR";
      break;
    default:
      levelStr = "INFO";
  }
  
  console.log(`[${timestamp}] [${levelStr}] ${message}`);
}

/**
 * Convenience method for debug log
 * 
 * @param {string} message - Message to log
 */
function logDebug(message) {
  log(message, LOG_LEVELS.DEBUG);
}

/**
 * Convenience method for info log
 * 
 * @param {string} message - Message to log
 */
function logInfo(message) {
  log(message, LOG_LEVELS.INFO);
}

/**
 * Convenience method for warning log
 * 
 * @param {string} message - Message to log
 */
function logWarning(message) {
  log(message, LOG_LEVELS.WARN);
}

/**
 * Convenience method for error log
 * 
 * @param {string} message - Message to log
 */
function logError(message) {
  log(message, LOG_LEVELS.ERROR);
}

/**
 * Executes a function with error handling and logging
 * 
 * @param {Function} func - Function to execute
 * @param {string} errorMessage - Error message prefix
 * @param {*} defaultValue - Default value to return on error
 * @return {*} Function result or default value on error
 */
function safeExecute(func, errorMessage, defaultValue = null) {
  try {
    return func();
  } catch (error) {
    logError(`${errorMessage}: ${error.message}\n${error.stack}`);
    return defaultValue;
  }
}

/**
 * Validates a hexadecimal color code
 * 
 * @param {string} color - Color code to validate
 * @return {boolean} True if color code is valid
 */
function isValidHexColor(color) {
  return /^#([0-9A-F]{3}){1,2}$/i.test(color);
}

/**
 * Converts RGB values to a hex color code
 * 
 * @param {number} r - Red (0-255)
 * @param {number} g - Green (0-255)
 * @param {number} b - Blue (0-255)
 * @return {string} Hex color code
 */
function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

/**
 * Converts a hex color code to RGB values
 * 
 * @param {string} hex - Hex color code
 * @return {Object} Object with r, g, b properties (0-255)
 */
function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}

/**
 * Safely serializes objects to ensure they can be properly transmitted 
 * between server and client in Google Apps Script
 * 
 * @param {*} obj - Object to serialize
 * @return {*} Serialization-safe object
 */
function safeSerialize(obj) {
  if (obj === null || obj === undefined) {
    return obj;
  }
  
  // Handle primitive types directly
  if (typeof obj !== 'object' && typeof obj !== 'function') {
    return obj;
  }
  
  // Handle Date objects
  if (obj instanceof Date) {
    return {
      __type: 'Date',
      iso: obj.toISOString(),
      timestamp: obj.getTime()
    };
  }
  
  // Handle arrays
  if (Array.isArray(obj)) {
    return obj.map(item => safeSerialize(item));
  }
  
  // Handle objects
  const result = {};
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      // Skip function properties
      if (typeof obj[key] === 'function') {
        continue;
      }
      
      try {
        // Try to serialize the property
        result[key] = safeSerialize(obj[key]);
      } catch (e) {
        // If serialization fails, use a placeholder
        console.warn(`Could not serialize property ${key}:`, e);
        result[key] = `[Unserializable: ${typeof obj[key]}]`;
      }
    }
  }
  
  return result;
}