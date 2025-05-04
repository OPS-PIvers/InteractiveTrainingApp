/**
 * Client-side utility functions for the Interactive Training Projects Web App
 * Browser-compatible implementations for canvas editor functionality
 */

/**
 * Generates a UUID (version 4) for unique identifiers
 * 
 * @return {string} A UUID string
 */
function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }
  
  /**
   * Formats a date object to a human-readable string
   * 
   * @param {Date|number} date - Date object or timestamp
   * @param {string} format - Optional format string (default: local date/time)
   * @return {string} Formatted date string
   */
  function formatDate(date, format) {
    if (!date) return "";
    
    const dateObj = typeof date === "number" ? new Date(date) : date;
    
    // Simple format option handling, or default to locale string
    if (format === "short") {
      return dateObj.toLocaleDateString();
    } else if (format === "time") {
      return dateObj.toLocaleTimeString();
    } else if (format === "full") {
      return dateObj.toLocaleString();
    }
    
    return dateObj.toLocaleString();
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
   * Creates a debounced function that delays invoking func until after wait milliseconds
   * have elapsed since the last time the debounced function was invoked
   * 
   * @param {Function} func - Function to debounce
   * @param {number} wait - Milliseconds to wait
   * @return {Function} Debounced function
   */
  function debounce(func, wait) {
    let timeout;
    return function() {
      const context = this;
      const args = arguments;
      clearTimeout(timeout);
      timeout = setTimeout(() => func.apply(context, args), wait);
    };
  }
  
  /**
   * Capitalizes the first letter of a string
   * 
   * @param {string} string - String to capitalize
   * @return {string} Capitalized string
   */
  function capitalizeFirstLetter(string) {
    if (!string) return '';
    return string.charAt(0).toUpperCase() + string.slice(1);
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