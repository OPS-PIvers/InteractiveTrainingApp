/**
 * Configuration constants for the Interactive Training Projects Web App
 */

// Application metadata
const APP_NAME = "Interactive Training Projects";
const APP_VERSION = "1.0.0";
const APP_AUTHOR = "Your Organization";

// Spreadsheet structure constants
const SHEET_STRUCTURE = {
  // ProjectIndex tab structure
  PROJECT_INDEX: {
    TAB_NAME: "ProjectIndex",
    COLUMNS: {
      PROJECT_ID: 0,
      TITLE: 1,
      CREATED_AT: 2,
      MODIFIED_AT: 3,
      LAST_ACCESSED: 4,
      PROJECT_ADMINS: 5
    },
    HEADERS: ["Project ID", "Title", "Created At", "Modified At", "Last Accessed"]
  },
  
  // Template and project tab structure
  PROJECT_TAB: {
    // PROJECT INFO section (Rows 1-7)
    PROJECT_INFO: {
      SECTION_START_ROW: 1,
      SECTION_END_ROW: 7,
      HEADER: "PROJECT INFO",
      FIELDS: {
        PROJECT_ID: { ROW: 2, LABEL_COL: "A", VALUE_COL: "B" },
        PROJECT_WEB_APP_URL: { ROW: 3, LABEL_COL: "A", VALUE_COL: "B" },
        TITLE: { ROW: 4, LABEL_COL: "A", VALUE_COL: "B" },
        CREATED_AT: { ROW: 5, LABEL_COL: "A", VALUE_COL: "B" },
        MODIFIED_AT: { ROW: 6, LABEL_COL: "A", VALUE_COL: "B" },
        PROJECT_FOLDER_ID: { ROW: 7, LABEL_COL: "A", VALUE_COL: "B" }
      }
    },
    
    // SLIDE INFO section (Rows 8-15)
    SLIDE_INFO: {
      SECTION_START_ROW: 8,
      SECTION_END_ROW: 15,
      HEADER: "SLIDE 1 INFO (Required)",
      FIELDS: {
        SLIDE_ID: { ROW: 9, LABEL_COL: "A", VALUE_COL: "B" },
        SLIDE_TITLE: { ROW: 10, LABEL_COL: "A", VALUE_COL: "B" },
        BACKGROUND_COLOR: { ROW: 11, LABEL_COL: "A", VALUE_COL: "B" },
        FILE_TYPE: { ROW: 12, LABEL_COL: "A", VALUE_COL: "B" },
        FILE_URL: { ROW: 13, LABEL_COL: "A", VALUE_COL: "B" },
        SLIDE_NUMBER: { ROW: 14, LABEL_COL: "A", VALUE_COL: "B" },
        SHOW_CONTROLS: { ROW: 15, LABEL_COL: "A", VALUE_COL: "B" }
      }
    },
    
    // ELEMENT INFO section (Columns D-G, Rows 1-28)
    ELEMENT_INFO: {
      SECTION_START_ROW: 1,
      SECTION_END_ROW: 28,
      SECTION_START_COL: "D",
      SECTION_END_COL: "G",
      HEADER: "ELEMENT INFO",
      ELEMENT_HEADERS_ROW: 1,
      FIELDS: {
        ELEMENT_ID: { ROW: 2, LABEL_COL: "D" },
        NICKNAME: { ROW: 3, LABEL_COL: "D" },
        SLIDE_ID: { ROW: 4, LABEL_COL: "D" },
        SEQUENCE: { ROW: 5, LABEL_COL: "D" },
        TYPE: { ROW: 6, LABEL_COL: "D" },
        LEFT: { ROW: 7, LABEL_COL: "D" },
        TOP: { ROW: 8, LABEL_COL: "D" },
        WIDTH: { ROW: 9, LABEL_COL: "D" },
        HEIGHT: { ROW: 10, LABEL_COL: "D" },
        ANGLE: { ROW: 11, LABEL_COL: "D" },
        INITIALLY_HIDDEN: { ROW: 12, LABEL_COL: "D" },
        OPACITY: { ROW: 13, LABEL_COL: "D" },
        COLOR: { ROW: 14, LABEL_COL: "D" },
        OUTLINE: { ROW: 15, LABEL_COL: "D" },
        OUTLINE_WIDTH: { ROW: 16, LABEL_COL: "D" },
        OUTLINE_COLOR: { ROW: 17, LABEL_COL: "D" },
        SHADOW: { ROW: 18, LABEL_COL: "D" },
        TEXT: { ROW: 19, LABEL_COL: "D" },
        FONT: { ROW: 20, LABEL_COL: "D" },
        FONT_COLOR: { ROW: 21, LABEL_COL: "D" },
        FONT_SIZE: { ROW: 22, LABEL_COL: "D" },
        TRIGGERS: { ROW: 23, LABEL_COL: "D" },
        INTERACTION_TYPE: { ROW: 24, LABEL_COL: "D" },
        TEXT_MODAL: { ROW: 25, LABEL_COL: "D" },
        TEXT_MODAL_MESSAGE: { ROW: 26, LABEL_COL: "D" },
        ANIMATION_TYPE: { ROW: 27, LABEL_COL: "D" },
        ANIMATION_SPEED: { ROW: 28, LABEL_COL: "D" }
      }
    },
    
    // TIMELINE section (Columns D-G, Rows 29-36)
    TIMELINE: {
      SECTION_START_ROW: 29,
      SECTION_END_ROW: 36,
      SECTION_START_COL: "D",
      SECTION_END_COL: "G",
      HEADER: "TIMELINE",
      FIELDS: {
        ELEMENT_ID: { ROW: 30, LABEL_COL: "D" },
        START_TIME: { ROW: 31, LABEL_COL: "D" },
        END_TIME: { ROW: 32, LABEL_COL: "D" },
        PAUSE_AT: { ROW: 33, LABEL_COL: "D" },
        SHOW_FOR_DURATION: { ROW: 34, LABEL_COL: "D" },
        ANIMATION_IN: { ROW: 35, LABEL_COL: "D" },
        ANIMATION_OUT: { ROW: 36, LABEL_COL: "D" }
      }
    },
    
    // QUIZ section (Columns D-G, Rows 37-49)
    QUIZ: {
      SECTION_START_ROW: 37,
      SECTION_END_ROW: 49,
      SECTION_START_COL: "D",
      SECTION_END_COL: "G",
      HEADER: "QUIZ",
      FIELDS: {
        QUESTION_TYPE: { ROW: 38, LABEL_COL: "D" },
        QUESTION_TEXT: { ROW: 39, LABEL_COL: "D" },
        CORRECT_ANSWER: { ROW: 40, LABEL_COL: "D" },
        INCORRECT_ANSWER_1: { ROW: 41, LABEL_COL: "D" },
        INCORRECT_ANSWER_2: { ROW: 42, LABEL_COL: "D" },
        INCORRECT_ANSWER_3: { ROW: 43, LABEL_COL: "D" },
        INCLUDE_FEEDBACK: { ROW: 44, LABEL_COL: "D" },
        CORRECT_FEEDBACK: { ROW: 45, LABEL_COL: "D" },
        INCORRECT_FEEDBACK: { ROW: 46, LABEL_COL: "D" },
        POINTS: { ROW: 47, LABEL_COL: "D" },
        ATTEMPTS: { ROW: 48, LABEL_COL: "D" },
        PROVIDE_CORRECT_ANSWER: { ROW: 49, LABEL_COL: "D" }
      }
    },
    
    // USER TRACKING section (Columns D-G, Rows 50-55)
    USER_TRACKING: {
      SECTION_START_ROW: 50,
      SECTION_END_ROW: 55,
      SECTION_START_COL: "D",
      SECTION_END_COL: "G",
      HEADER: "USER TRACKING",
      FIELDS: {
        TRACK_COMPLETION: { ROW: 51, LABEL_COL: "D" },
        REQUIRE_QUIZ_COMPLETION: { ROW: 52, LABEL_COL: "D" },
        PASSING_SCORE: { ROW: 53, LABEL_COL: "D" },
        SEND_COMPLETION_EMAIL: { ROW: 54, LABEL_COL: "D" },
        INSTRUCTOR_EMAIL: { ROW: 55, LABEL_COL: "D" }
      }
    }
  }
};

// Element types and configurations
const ELEMENT_TYPES = {
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
};

// File type options for Slide background
const FILE_TYPES = {
  IMAGE: "Image",
  VIDEO: "YouTube Video",
  AUDIO: "Audio"
};

// Font options
const FONTS = [
  "Roboto",
  "Arial",
  "Times New Roman",
  "Verdana",
  "Georgia"
];

// Trigger options
const TRIGGER_TYPES = {
  HOVER: "Hover",
  CLICK: "Click",
  BOTH: "Both"
};

// Interaction types
const INTERACTION_TYPES = {
  REVEAL: "Reveal",
  SPOTLIGHT: "Spotlight",
  PAN_ZOOM: "Pan/Zoom",
  CENTER: "Center",
  QUIZ: "Quiz"
};

// Animation types
const ANIMATION_TYPES = {
  WIGGLE: "Wiggle",
  FLOAT: "Float",
  HINT: "Hint",
  GROW_SHRINK: "Grow/Shrink",
  NONE: "None"
};

// Animation speed options
const ANIMATION_SPEEDS = {
  SLOW: "Slow",
  MEDIUM: "Medium",
  FAST: "Fast"
};

// Animation entrance types
const ANIMATION_IN_TYPES = {
  FADE_IN: "Fade In",
  SLIDE_IN: "Slide In",
  POP: "Pop",
  NONE: "None"
};

// Animation exit types
const ANIMATION_OUT_TYPES = {
  FADE_OUT: "Fade Out",
  SLIDE_OUT: "Slide Out",
  SHRINK: "Shrink",
  NONE: "None"
};

// Question types for quizzes
const QUESTION_TYPES = {
  MULTIPLE_CHOICE: "Multiple choice",
  TRUE_FALSE: "True/False",
  FILL_BLANK: "Fill in the blank",
  MATCHING: "Matching",
  ORDERING: "Ordering",
  HOTSPOT: "Hotspot"
};

// Default colors
const DEFAULT_COLORS = {
  BACKGROUND: "#FFFFFF",
  TEXT: "#000000",
  ELEMENT: "#4285F4",
  OUTLINE: "#000000",
  SECTION_HEADER_BG: "#E0E0E0",
  COMPLETED: "#4CAF50",
  ERROR: "#F44336",
  WARNING: "#FFC107"
};

// Google Drive settings
const DRIVE_SETTINGS = {
  ROOT_FOLDER_NAME: "Interactive Training Projects",
  MEDIA_FOLDER_NAME: "Media Assets",
  THUMBNAIL_FOLDER_NAME: "Thumbnails",
  ALLOWED_MIME_TYPES: {
    IMAGE: ["image/jpeg", "image/png", "image/gif", "image/svg+xml"],
    AUDIO: ["audio/mpeg", "audio/wav", "audio/ogg"]
  }
};

// UI settings
const UI_SETTINGS = {
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
};

// Logging levels
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  NONE: 4
};

// Current logging level
const CURRENT_LOG_LEVEL = LOG_LEVELS.INFO;

// Export as needed