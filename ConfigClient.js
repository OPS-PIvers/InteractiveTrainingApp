/**
 * ConfigClient.js - Client-side configuration constants for the Interactive Training Projects Web App
 * This file provides a client-side namespace for configuration to avoid conflicts with server-side Config.js
 */

// Create a namespace for client-side configuration to avoid conflicts
const ClientConfig = {
    // Element types and configurations
    ELEMENT_TYPES: {
      RECTANGLE: {
        ID: "rectangle",
        LABEL: "Rectangle",
        DEFAULT_WIDTH: 100,
        DEFAULT_HEIGHT: 60
      },
      ROUNDED_RECTANGLE: {
        ID: "roundedRectangle",
        LABEL: "Rounded Rectangle",
        DEFAULT_WIDTH: 100,
        DEFAULT_HEIGHT: 60,
        DEFAULT_RADIUS: 10
      },
      CIRCLE: {
        ID: "circle",
        LABEL: "Circle",
        DEFAULT_RADIUS: 50
      },
      ARROW: {
        ID: "arrow",
        LABEL: "Arrow",
        DEFAULT_LENGTH: 100,
        DEFAULT_WIDTH: 2
      },
      TEXT: {
        ID: "text",
        LABEL: "Text",
        DEFAULT_WIDTH: 200,
        DEFAULT_HEIGHT: 50,
        DEFAULT_FONT: "Roboto",
        DEFAULT_FONT_SIZE: 14
      },
      HOTSPOT: {
        ID: "hotspot",
        LABEL: "Hotspot",
        DEFAULT_WIDTH: 100,
        DEFAULT_HEIGHT: 100,
        DEFAULT_OPACITY: 0.2
      }
    },
    
    // File type options for Slide background
    FILE_TYPES: {
      IMAGE: "Image",
      VIDEO: "YouTube Video",
      AUDIO: "Audio"
    },
    
    // Font options
    FONTS: [
      "Roboto",
      "Arial",
      "Times New Roman",
      "Verdana",
      "Georgia"
    ],
    
    // Trigger options
    TRIGGER_TYPES: {
      HOVER: "Hover",
      CLICK: "Click",
      BOTH: "Both"
    },
    
    // Interaction types
    INTERACTION_TYPES: {
      REVEAL: "Reveal",
      SPOTLIGHT: "Spotlight",
      PAN_ZOOM: "Pan/Zoom",
      CENTER: "Center",
      QUIZ: "Quiz"
    },
    
    // Animation types
    ANIMATION_TYPES: {
      WIGGLE: "Wiggle",
      FLOAT: "Float",
      HINT: "Hint",
      GROW_SHRINK: "Grow/Shrink",
      NONE: "None"
    },
    
    // Animation speed options
    ANIMATION_SPEEDS: {
      SLOW: "Slow",
      MEDIUM: "Medium",
      FAST: "Fast"
    },
    
    // Animation entrance types
    ANIMATION_IN_TYPES: {
      FADE_IN: "Fade In",
      SLIDE_IN: "Slide In",
      POP: "Pop",
      NONE: "None"
    },
    
    // Animation exit types
    ANIMATION_OUT_TYPES: {
      FADE_OUT: "Fade Out",
      SLIDE_OUT: "Slide Out",
      SHRINK: "Shrink",
      NONE: "None"
    },
    
    // Question types for quizzes
    QUESTION_TYPES: {
      MULTIPLE_CHOICE: "Multiple choice",
      TRUE_FALSE: "True/False",
      FILL_BLANK: "Fill in the blank",
      MATCHING: "Matching",
      ORDERING: "Ordering",
      HOTSPOT: "Hotspot"
    },
    
    // Default colors
    DEFAULT_COLORS: {
      BACKGROUND: "#FFFFFF",
      TEXT: "#000000",
      ELEMENT: "#4285F4",
      OUTLINE: "#000000",
      SECTION_HEADER_BG: "#E0E0E0",
      COMPLETED: "#4CAF50",
      ERROR: "#F44336",
      WARNING: "#FFC107"
    },
    
    // UI settings
    UI_SETTINGS: {
      CANVAS: {
        DEFAULT_WIDTH: 800,
        DEFAULT_HEIGHT: 600,
        MIN_WIDTH: 400,
        MIN_HEIGHT: 300,
        MAX_WIDTH: 1920,
        MAX_HEIGHT: 1080
      },
      TIMELINE: {
        TRACK_HEIGHT: 40,
        KEYFRAME_WIDTH: 12,
        TICK_INTERVAL: 1 // in seconds
      }
    }
  };