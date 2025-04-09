/**
 * @OnlyCurrentDoc
 */


// --- Add-on Menu & Sidebar ---


/**
 * Creates the add-on menu in Google Slides when the presentation is opened.
 */
function onOpen() {
  SlidesApp.getUi()
      .createMenu('Onboarding Tools')
      .addItem('Show Builder Panel', 'showSidebar')
      .addItem('Get Shareable URL', 'showShareableUrl')
      .addSeparator()
      // .addItem('Reset Step Counter', 'resetStepCounter') // Removed as Step Numbers section is gone
      .addItem('Set Web App Deployment ID', 'setDeploymentId') // Exposed via menu
      .addToUi();
}


/**
 * Shows the sidebar panel by loading the Sidebar.html file.
 */
function showSidebar() {
  // Creates an HTML output object from the Sidebar.html file.
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Training Builder'); // Sets the title displayed at the top of the sidebar.
  // Displays the HTML output as a sidebar in the Google Slides UI.
  SlidesApp.getUi().showSidebar(htmlOutput);
}


// --- Sidebar Communication Functions ---


/**
 * Gets information about the currently selected element(s) on the slide.
 * Returns details for a single selected page element, or an error message.
 * Includes flags indicating if navigation to previous/next elements is possible.
 * Called by the sidebar (refreshSelectedElement) to update its display.
 * @returns {object} An object containing element details (id, type, position, size, description, canGoPrev, canGoNext) or an error message.
 */
function getSelectedElementInfo() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();


  // Check if a single page element (shape, text box, image, etc.) is selected.
  if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
    const pageElements = selection.getPageElementRange().getPageElements();
    if (pageElements.length === 1) {
      // If exactly one element is selected, gather its properties.
      const element = pageElements[0];
      const elementId = element.getObjectId();
      const slide = element.getParentPage(); // Get the slide the element is on


      if (!slide) {
        return { error: "Could not determine the element's slide." };
      }


      const allSlideElements = slide.getPageElements();
      let currentIndex = -1;
      for (let i = 0; i < allSlideElements.length; i++) {
        if (allSlideElements[i].getObjectId() === elementId) {
          currentIndex = i;
          break;
        }
      }


      const canGoPrev = currentIndex > 0;
      const canGoNext = currentIndex >= 0 && currentIndex < allSlideElements.length - 1;


      return {
        id: elementId,
        type: element.getPageElementType().toString(),
        left: element.getLeft(),
        top: element.getTop(),
        width: element.getWidth(),
        height: element.getHeight(),
        description: element.getDescription(), // The description holds the JSON interaction data.
        canGoPrev: canGoPrev, // Flag for previous element navigation
        canGoNext: canGoNext  // Flag for next element navigation
      };
    } else if (pageElements.length > 1) {
      // Error if multiple elements are selected.
      return { error: "Select only one element." };
    }
  } else if (selectionType === SlidesApp.SelectionType.CURRENT_PAGE) {
    // Error if the entire slide is selected instead of a specific element.
    return { error: "Select a specific shape or text box, not just the slide." };
  }
  // Default error if nothing suitable is selected.
  // Also return nav flags as false in this case.
  return { error: "Select a single shape or text box on the slide.", canGoPrev: false, canGoNext: false };
}


/**
 * Selects the previous element on the current slide relative to the current selection.
 * @returns {object} The result of getSelectedElementInfo() for the newly selected element, or an error.
 */
function selectPreviousObjectOnSlide() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();


  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    return { error: "Select an element first.", canGoPrev: false, canGoNext: false };
  }


  const pageElements = selection.getPageElementRange().getPageElements();
  if (pageElements.length !== 1) {
    return { error: "Select only one element.", canGoPrev: false, canGoNext: false };
  }


  const currentElement = pageElements[0];
  const elementId = currentElement.getObjectId();
  const slide = currentElement.getParentPage();


  if (!slide) {
    return { error: "Could not determine the element's slide.", canGoPrev: false, canGoNext: false };
  }


  const allSlideElements = slide.getPageElements();
  let currentIndex = -1;
  for (let i = 0; i < allSlideElements.length; i++) {
    if (allSlideElements[i].getObjectId() === elementId) {
      currentIndex = i;
      break;
    }
  }


  if (currentIndex > 0) {
    const prevElement = allSlideElements[currentIndex - 1];
    prevElement.select();
    Utilities.sleep(50); // Short pause might help UI update before getting info
    return getSelectedElementInfo(); // Return info for the newly selected element
  } else {
    // Already at the first element, cannot go previous
    return getSelectedElementInfo(); // Return current info, canGoPrev will be false
  }
}


/**
 * Selects the next element on the current slide relative to the current selection.
 * @returns {object} The result of getSelectedElementInfo() for the newly selected element, or an error.
 */
function selectNextObjectOnSlide() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();


  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    return { error: "Select an element first.", canGoPrev: false, canGoNext: false };
  }


  const pageElements = selection.getPageElementRange().getPageElements();
  if (pageElements.length !== 1) {
    return { error: "Select only one element.", canGoPrev: false, canGoNext: false };
  }


  const currentElement = pageElements[0];
  const elementId = currentElement.getObjectId();
  const slide = currentElement.getParentPage();


  if (!slide) {
    return { error: "Could not determine the element's slide.", canGoPrev: false, canGoNext: false };
  }


  const allSlideElements = slide.getPageElements();
  let currentIndex = -1;
  for (let i = 0; i < allSlideElements.length; i++) {
    if (allSlideElements[i].getObjectId() === elementId) {
      currentIndex = i;
      break;
    }
  }


  if (currentIndex >= 0 && currentIndex < allSlideElements.length - 1) {
    const nextElement = allSlideElements[currentIndex + 1];
    nextElement.select();
    Utilities.sleep(50); // Short pause might help UI update before getting info
    return getSelectedElementInfo(); // Return info for the newly selected element
  } else {
    // Already at the last element, cannot go next
    return getSelectedElementInfo(); // Return current info, canGoNext will be false
  }
}




/**
 * Sets the description of a specific element, storing interaction data as a JSON string.
 * Called by the sidebar's save functions (saveInteractionData, saveAnimationData).
 * @param {string} elementId The ID of the element to modify.
 * @param {string} interactionDataJson The interaction data formatted as a JSON string. An empty string clears the interaction.
 * @returns {object} An object indicating success or failure, with a message or error.
 */
function setElementInteractionData(elementId, interactionDataJson) {
   const presentation = SlidesApp.getActivePresentation();
   const selection = presentation.getSelection();
   // Ensure there's a selection context to find the slide.
   if (!selection || selection.getSelectionType() === SlidesApp.SelectionType.NONE) {
       // Try getting the first slide if no specific selection exists
       const firstSlide = presentation.getSlides()[0];
       if (!firstSlide) {
            return { success: false, error: "No slide found in the presentation." };
       }
       const elementOnFirstSlide = firstSlide.getPageElementById(elementId);
       if(elementOnFirstSlide) {
           return saveDescription(elementOnFirstSlide, interactionDataJson);
       } else {
            return { success: false, error: "No slide or element context found, and element not on first slide." };
       }
   }


   // Get the current slide from the selection.
   const slide = selection.getCurrentPage();
   if (!slide) {
       return { success: false, error: "Could not determine the active slide." };
   }


   // Find the element on the current slide using its ID.
   const element = slide.getPageElementById(elementId);
   if (element) {
       // Call helper function to save the description
       return saveDescription(element, interactionDataJson);
   } else {
       // Element not found on the currently selected slide.
       // Let's search all slides as a fallback
       console.log(`Element ${elementId} not found on current slide, searching all slides...`);
       const allSlides = presentation.getSlides();
       for (let s of allSlides) {
            const foundElement = s.getPageElementById(elementId);
            if (foundElement) {
                console.log(`Found element ${elementId} on slide ${s.getObjectId()}`);
                return saveDescription(foundElement, interactionDataJson);
            }
       }
       // If still not found after searching all slides
       return { success: false, error: "Element not found in the presentation." };
   }
}


/**
 * Helper function to validate JSON (if not empty) and save to element description.
 * @param {PageElement} element The element to modify.
 * @param {string} jsonString The JSON data string or an empty string.
 * @returns {object} An object indicating success or failure.
 */
function saveDescription(element, jsonString) {
    try {
        // If the string is not empty, try parsing it to ensure it's valid JSON.
        if (jsonString && jsonString.trim() !== "") {
            JSON.parse(jsonString); // This will throw an error if invalid JSON.
        }
        // Set the element's description (stores the JSON string or clears it).
        element.setDescription(jsonString);
        console.log(`Saved description for ${element.getObjectId()}: ${jsonString.substring(0, 100)}...`);
        return { success: true, message: "Data saved successfully." };
    } catch (e) {
        // Log the error if parsing or setting fails.
        console.error(`Error saving description for element ${element.getObjectId()}: ${e.message}. Data: ${jsonString}`);
        return { success: false, error: `Error saving data: Invalid format. ${e.message}` };
    }
}


/**
 * Merges new interaction or animation data with existing data for a specific element.
 * Handles 'overlayText', 'showOverlayText', 'useCustomOpacity', 'customOpacity', and 'overlayStyle' within the interaction object.
 *
 * @param {string} elementId The ID of the element to modify.
 * @param {object} newData The new data to merge (e.g., { type: 'showText', text: '...', overlayStyle: {...} }).
 * @param {string} dataType Either 'interaction' or 'animation' to indicate which part to update.
 * @returns {object} An object indicating success or failure, with message or error.
 */
function mergeElementData(elementId, newData, dataType) {
  try {
    console.log(`mergeElementData called for ${elementId}, type: ${dataType}`);
    console.log(`New data: ${JSON.stringify(newData)}`);


    // Get the current element description to check for existing data
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();


    // Find the element (search current slide first, then all slides if needed)
    let element = null;


    // First try the current slide from selection
    if (selection && selection.getSelectionType() !== SlidesApp.SelectionType.NONE) {
      const slide = selection.getCurrentPage();
      if (slide) {
        element = slide.getPageElementById(elementId);
        console.log(`Element search on current slide: ${element ? "Found" : "Not found"}`);
      }
    }


    // If not found, try all slides
    if (!element) {
      console.log("Searching all slides for element...");
      const allSlides = presentation.getSlides();
      for (let slide of allSlides) {
        element = slide.getPageElementById(elementId);
        if (element) {
          console.log(`Found element on slide: ${slide.getObjectId()}`);
          break;
        }
      }
    }


    // Handle element not found
    if (!element) {
      return { success: false, error: "Element not found in presentation." };
    }


    // Get current description and parse as JSON if exists
    let currentData = {};
    const description = element.getDescription();
    console.log(`Current description: ${description || "empty"}`);


    if (description && description.trim() !== "") {
      try {
        currentData = JSON.parse(description);
        console.log(`Successfully parsed existing JSON data: ${JSON.stringify(currentData)}`);
      } catch (e) {
        console.warn(`Invalid JSON in element description, treating as empty: ${e.message}`);
        // Continue with empty object if parsing fails
      }
    }


    // --- MERGE LOGIC ---
    // Create a base object with existing data
    let mergedData = { ...currentData };


    // Update or add the specified data type section
    if (newData.type === 'none') {
      // If setting type to 'none', remove that entire section
      delete mergedData[dataType];
      console.log(`Removed ${dataType} section because type is 'none'`);
    } else {
      // Otherwise, merge the new data into the appropriate section
      // Ensure the section exists before merging
      if (!mergedData[dataType]) {
          mergedData[dataType] = {};
      }

      // Preserve existing sub-properties if not overwritten by newData
      // Example: If saving animation, keep existing interaction data untouched.
      // If saving interaction, keep existing animation data untouched.
      // If saving interaction, ensure existing interaction sub-properties are kept unless overridden
      let existingDataTypeData = mergedData[dataType] || {};
      mergedData[dataType] = { ...existingDataTypeData, ...newData };


      // Clean up interaction-specific properties based on current state
      if (dataType === 'interaction') {
        const interaction = mergedData.interaction;

        // Clean up overlayText if showOverlayText is false/undefined
        if (!interaction.showOverlayText) {
          delete interaction.overlayText;
        }
        // Clean up customOpacity if useCustomOpacity is false/undefined
        if (!interaction.useCustomOpacity) {
          delete interaction.customOpacity;
        }
        // Remove overlayStyle entirely if interaction.useCustomOverlay (new flag) is false/undefined
        if (!interaction.useCustomOverlay) { // Assuming a new flag from sidebar indicates customisation
          delete interaction.overlayStyle;
        }
        // Remove the useCustomOverlay flag itself after cleanup
        delete interaction.useCustomOverlay;

      }
      console.log(`Updated ${dataType} section with new data: ${JSON.stringify(mergedData[dataType])}`);
    }


    // Check if the entire description should be cleared
    // Note: An interaction might exist but have no effect (e.g., only overlayStyle) - consider this if needed.
    // For now, keep if *any* data exists in either section.
    const interactionIsEmpty = !mergedData.interaction || Object.keys(mergedData.interaction).length === 0 || mergedData.interaction.type === 'none';
    const animationIsEmpty = !mergedData.animation || Object.keys(mergedData.animation).length === 0 || mergedData.animation.type === 'none';


    if (interactionIsEmpty && animationIsEmpty) {
      console.log("Both interaction and animation are 'none' or empty, clearing description");
      element.setDescription("");
      return {
        success: true,
        message: `${dataType.charAt(0).toUpperCase() + dataType.slice(1)} data cleared.`,
        elementId: elementId, // Return elementId for client reference
        verifiedFull: {} // Indicate cleared state
      };
    }


    // --- SAVE MERGED DATA ---
    const mergedJson = JSON.stringify(mergedData);
    console.log(`Setting new description: ${mergedJson}`);
    element.setDescription(mergedJson);


    // Verify the save worked by reading it back
    const verifyDesc = element.getDescription();
    console.log(`Verification - saved description: ${verifyDesc}`);

    // Verify parsing of the saved data
    let verifiedParsedData = null;
    try {
      verifiedParsedData = JSON.parse(verifyDesc);
    } catch(e) {
       console.error("Verification failed: Could not parse saved description back.");
    }


    return {
      success: true,
      message: `${dataType.charAt(0).toUpperCase() + dataType.slice(1)} data saved successfully.`,
      elementId: elementId, // Return elementId for client reference
      verifiedDataType: dataType,
      verifiedData: verifiedParsedData ? verifiedParsedData[dataType] : null, // Return the specific part from *verified* parse
      verifiedFull: verifiedParsedData // Return the full *verified* parsed object
    };
  } catch (e) {
    console.error(`Error in mergeElementData: ${e.message}\nStack: ${e.stack}`); // Log stack trace
    return {
      success: false,
      error: `Failed to merge data: ${e.message}`,
      elementId: elementId // Return elementId even on error
    };
  }
}




/**
 * Gets all elements on the current slide that have interaction or animation data stored.
 * Includes detailed parsing logging.
 *
 * @returns {object} An object containing an array of 'elements' or an 'error'.
 * Each element includes id, name, type, and parsed interaction/animation data.
 */
function getAllInteractiveElements() {
  try {
    const selection = SlidesApp.getActivePresentation().getSelection();
    if (!selection) { return { error: "No active presentation found." }; }


    // Use the current page from selection, or default to the first slide
    const slide = selection.getCurrentPage() || SlidesApp.getActivePresentation().getSlides()[0];
    if (!slide) { return { error: "Could not determine the active slide." }; }


    const elements = slide.getPageElements();
    const interactiveElements = [];


    console.log(`Checking ${elements.length} elements on slide ${slide.getObjectId()} for interactive data`);


    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      const desc = element.getDescription();


      // Log basic info for each element checked
      // console.log(`Element ${i}: ID=${element.getObjectId()}, Type=${element.getPageElementType()}, Has Description: ${!!desc}`);


      if (desc && desc.trim().startsWith('{') && desc.trim().endsWith('}')) { // Basic check for JSON
        let data = null;
        try {
          data = JSON.parse(desc);
          // console.log(`Parsed data for element ${element.getObjectId()}:`, JSON.stringify(data));


          // Check if data is an object and has *any* interaction or animation keys
          // (even if type is none, it might have overlayStyle)
           const hasInteractionData = data.interaction && typeof data.interaction === 'object';
           const hasAnimationData = data.animation && typeof data.animation === 'object';

           // Consider it "interactive" if it has either block, regardless of 'type' value
           if (hasInteractionData || hasAnimationData) {
            // Get element name with improved error handling
            let elementName = "";
            try {
              if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE && element.asShape().getText) {
                elementName = element.asShape().getText().asString().substring(0, 30).trim();
                if (element.asShape().getText().asString().length > 30) elementName += "...";
              } else if (element.getPageElementType() === SlidesApp.PageElementType.TEXT_BOX && element.asTextBox().getText) {
                elementName = element.asTextBox().getText().asString().substring(0, 30).trim();
                if (element.asTextBox().getText().asString().length > 30) elementName += "...";
              } else if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
                  elementName = "Image"; // Handle Images
              } else if (element.getPageElementType() === SlidesApp.PageElementType.VIDEO) {
                  elementName = "Video"; // Handle Videos
              }
              // Add other element types as needed
            } catch (e) {
              console.warn(`Could not get text/type name for element ${element.getObjectId()}: ${e.message}`);
            }


            if (!elementName) {
              // Fallback name using type and ID
              elementName = `${element.getPageElementType().toString()} (${element.getObjectId().substring(0, 5)}...)`;
            }


            // Add element to list, including the raw description for client-side parsing
            interactiveElements.push({
              id: element.getObjectId(),
              name: elementName,
              elementType: element.getPageElementType().toString(),
              description: desc // Pass raw description for client parsing
            });


            // console.log(`Added element ${element.getObjectId()} (${elementName}) to interactive elements list`);
          }
        } catch (e) {
          console.warn(`Element ${element.getObjectId()} has invalid JSON in description: ${e.message}`);
          console.warn(`Raw description: ${desc}`);
        }
      }
    }


    console.log(`Found ${interactiveElements.length} interactive elements on slide ${slide.getObjectId()}`);
    return { elements: interactiveElements };
  } catch (error) {
    console.error("Error in getAllInteractiveElements:", error);
    return { error: `Failed to get interactive elements: ${error.message}` };
  }
}




/**
 * Removes interaction data from a specific element by clearing its description.
 * Called by the sidebar when the delete button is clicked on a list item.
 * @param {string} elementId The ID of the element to modify.
 * @returns {object} An object indicating success or failure.
 */
function removeElementInteraction(elementId) {
  // NOTE: This function clears ALL data (interaction + animation).
  // If only one type should be removed, mergeElementData with type 'none' should be used.
  // Keeping this function as is for now, assuming 'Delete' means remove all configs.
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
     // Find the slide containing the element (check current slide first, then search all)
    let element = null;
    let slide = selection ? selection.getCurrentPage() : null;


    if (slide) {
        element = slide.getPageElementById(elementId);
    }


    // If not found on current slide, search all slides (more robust)
    if (!element) {
        const allSlides = presentation.getSlides();
        for (let s of allSlides) {
            element = s.getPageElementById(elementId);
            if (element) {
                slide = s; // Found the slide
                break;
            }
        }
    }


    if (!element) {
        return { success: false, error: "Element not found in the presentation." };
    }
    if (!slide) {
         return { success: false, error: "Could not determine the slide for the element." }; // Should not happen if element found
    }


    // Clear the description to remove ALL interaction/animation data.
    element.setDescription("");
    console.log(`Cleared description for element ${elementId}`);


    return {
      success: true,
      message: "All Interaction/Animation data removed successfully."
    };
  } catch (e) {
    console.error("Error removing element interaction:", e);
    return {
      success: false,
      error: `Failed to remove interaction: ${e.message}`
    };
  }
}


/**
 * Selects a specific element on its slide using its ID.
 * Called by the sidebar when an item in the interaction/animation list is clicked.
 * @param {string} elementId The ID of the element to select.
 * @returns {object} An object indicating success or failure.
 */
function selectElementById(elementId) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    let element = null;
    let targetSlide = null;


    // Find the element across all slides
    const allSlides = presentation.getSlides();
    for (let slide of allSlides) {
        element = slide.getPageElementById(elementId);
        if (element) {
            targetSlide = slide;
            break;
        }
    }


    if (!element) {
      return { success: false, error: "Element not found in the presentation." };
    }
    if (!targetSlide) {
         return { success: false, error: "Could not determine the slide for the element." };
    }


    // Switch to the target slide if it's not the current one
    if (targetSlide.getObjectId() !== presentation.getSelection()?.getCurrentPage()?.getObjectId()) {
        targetSlide.selectAsCurrentPage();
        // Brief pause to allow UI to update (might not be strictly necessary)
        Utilities.sleep(100);
    }


    // Select the specific element
    element.select(); // Simpler selection method


    return { success: true };
  } catch (e) {
    console.error("Error selecting element:", e);
    return { success: false, error: `Failed to select element: ${e.message}` };
  }
}




/**
 * Clears the description field of the selected element.
 * @returns {object} Result indicating success or failure.
 */
function clearElementDescription() {
  try {
    const info = getSelectedElementInfo();
    if (info.error) {
      return { success: false, error: info.error };
    }


    const elementId = info.id;


    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const slide = selection.getCurrentPage();
    const element = slide.getPageElementById(elementId);


    if (!element) {
       // If not on current slide, try finding it elsewhere before giving up
        let found = false;
        const allSlides = presentation.getSlides();
        for (let s of allSlides) {
            const el = s.getPageElementById(elementId);
            if (el) {
                element = el;
                found = true;
                break;
            }
        }
       if (!found) return { success: false, error: "Element not found in presentation." };
    }


    // Clear the description
    element.setDescription("");
     console.log(`Cleared description for element ${elementId}`);


    return {
      success: true,
      message: "Element description cleared successfully.",
      elementId: elementId
    };
  } catch (e) {
    return {
      success: false,
      error: `Failed to clear description: ${e.message}`
    };
  }
}


// --- Step Number Functions (REMOVED) ---
// Functions addStepNumber, resetStepCounter, getCurrentStepNumber were removed.


// --- NEW: Global Overlay Style Constants ---
const OVERLAY_SETTINGS_PREFIX = 'WEBAPP_OVERLAY_';
const OVERLAY_DEFAULTS = {
  shape: 'rectangle',
  color: '#e53935', // Default Red
  opacity: 15, // Default 15%
  outlineEnabled: true, // Default: Outline ON
  outlineColor: '#e53935', // Default Red
  outlineWidth: 1, // Default 1px
  outlineStyle: 'dashed', // Default dashed
  textColor: '#ffffff', // Default White
  textSize: 14, // Default 14px
  hoverText: 'Click here' // Default hover text
};


/**
 * Gets the global overlay style settings saved in Script Properties.
 * Returns defaults if a setting is not found.
 * @returns {object} An object containing all global overlay settings.
 */
function getGlobalOverlaySettings() {
  const props = PropertiesService.getScriptProperties();
  const settings = {};
  const keys = Object.keys(OVERLAY_DEFAULTS);

  keys.forEach(key => {
    const propKey = OVERLAY_SETTINGS_PREFIX + key.toUpperCase();
    let value = props.getProperty(propKey);

    if (value === null) {
      // Value not found, use default
      settings[key] = OVERLAY_DEFAULTS[key];
    } else {
      // Value found, parse it based on expected type
      const defaultValue = OVERLAY_DEFAULTS[key];
      if (typeof defaultValue === 'number') {
        const numValue = parseFloat(value);
        settings[key] = isNaN(numValue) ? defaultValue : numValue;
      } else if (typeof defaultValue === 'boolean') {
        settings[key] = (value === 'true');
      } else {
        // Assume string (color, shape, style, hoverText)
        settings[key] = value;
      }
    }
  });

  console.log(`getGlobalOverlaySettings: Returning ${JSON.stringify(settings)}`);
  return { success: true, settings: settings };
}


/**
 * Saves the global overlay style settings to Script Properties.
 * @param {object} settings An object containing the settings to save.
 * @returns {object} { success: boolean, message?: string, error?: string }
 */
function setGlobalOverlaySettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    const keys = Object.keys(OVERLAY_DEFAULTS); // Use defaults to know which keys to expect

    keys.forEach(key => {
      if (settings.hasOwnProperty(key)) {
        const propKey = OVERLAY_SETTINGS_PREFIX + key.toUpperCase();
        let valueToSave = settings[key];

        // Basic validation/type consistency (optional but recommended)
        const defaultValue = OVERLAY_DEFAULTS[key];
        if (typeof defaultValue === 'number') {
          valueToSave = parseFloat(valueToSave);
          if (isNaN(valueToSave)) valueToSave = defaultValue;
        } else if (typeof defaultValue === 'boolean') {
          valueToSave = !!valueToSave; // Ensure boolean
        } else if (typeof defaultValue === 'string') {
           valueToSave = String(valueToSave).trim(); // Ensure string, trim whitespace
           // Add specific validation e.g. for colors if needed
        }

        props.setProperty(propKey, String(valueToSave)); // Store all as strings
      } else {
        console.warn(`setGlobalOverlaySettings: Missing key '${key}' in input settings.`);
      }
    });

    console.log(`setGlobalOverlaySettings: Saved ${JSON.stringify(settings)}`);
    return { success: true, message: "Global overlay styles saved." };
  } catch (e) {
    console.error("Error setting global overlay settings:", e);
    return { success: false, error: `Failed to save global overlay settings: ${e.message}` };
  }
}


// --- DEPRECATED Overlay Style Functions ---
// These are replaced by getGlobalOverlaySettings / setGlobalOverlaySettings
// Keep them for a short while to avoid breaking older sidebar versions immediately,
// but eventually remove them.

const OVERLAY_OPACITY_KEY = 'WEBAPP_OVERLAY_OPACITY'; // DEPRECATED
const OVERLAY_SHADOW_KEY = 'WEBAPP_OVERLAY_SHADOW'; // DEPRECATED
const DEFAULT_OVERLAY_OPACITY = 15; // DEPRECATED
const DEFAULT_OVERLAY_SHADOW = false; // DEPRECATED

function getOverlayStyleSettings() {
   console.warn("DEPRECATED function getOverlayStyleSettings called. Use getGlobalOverlaySettings instead.");
   // Return structure compatible with old sidebar code for transition period
   const globalSettings = getGlobalOverlaySettings();
   return {
      success: globalSettings.success,
      opacity: globalSettings.settings?.opacity ?? DEFAULT_OVERLAY_OPACITY,
      // Map outlineEnabled to the old 'shadow' concept (true = shadow, false = no shadow)
      // This is an imperfect mapping, but provides some backward compatibility.
      // The new global settings provide more control (outline style, color, width).
      shadow: globalSettings.settings?.outlineEnabled ?? DEFAULT_OVERLAY_SHADOW,
      error: globalSettings.error
   };
}
function setOverlayStyleSettings(opacity, shadow) {
   console.warn("DEPRECATED function setOverlayStyleSettings called. Use setGlobalOverlaySettings instead.");
   // Update the *new* settings based on the old inputs
   const currentGlobal = getGlobalOverlaySettings().settings || {};
   currentGlobal.opacity = opacity;
   // Map old shadow bool to new outlineEnabled bool
   currentGlobal.outlineEnabled = !!shadow;
   // We don't have enough info to set outline color/style/width, so keep existing or defaults
   return setGlobalOverlaySettings(currentGlobal);
}


// --- Web App Functions ---


/**
 * Handles GET requests for the web app. Serves the WebApp.html file.
 * Passes the presentation ID to the template.
 * @param {object} e The event parameter for doGet, contains URL parameters.
 * @returns {HtmlOutput} The HTML page to be rendered.
 */
function doGet(e) {
  // Extract presentation ID from URL parameter 'presId', use default if missing.
  const presentationId = e?.parameter?.presId || "15sEovSiZVqWDdUdtatObclQPmB-LfzMCicjl3Wh9mxs"; // Replace with YOUR default ID if needed
  console.log(`doGet triggered for presentation ID: ${presentationId}`);


  // --- Optional: Debug/Loader Pages (Keep or Remove as needed) ---
  if (e?.parameter?.debug === "true") {
    console.log("Serving diagnostic page");
    return serveDiagnosticPage(e); // Use the dedicated function
  } else if (e?.parameter?.loader === "true") {
    console.log("Serving step loader page");
    return HtmlService.createHtmlOutputFromFile('StepLoader') // Ensure this file exists if used
        .setTitle('Interactive Training - Step Loader')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // --- End Optional ---


  try {
    // Verify access to the presentation ID. Throws error if inaccessible.
    SlidesApp.openById(presentationId);
    console.log(`Successfully verified access to presentation in doGet: ${presentationId}`);


    // Create an HTML template from WebApp.html.
    const template = HtmlService.createTemplateFromFile('WebApp');
    // Pass the presentation ID to the template's JavaScript scope.
    template.presentationId = presentationId;


    // Evaluate the template, set title and viewport meta tag.
    return template.evaluate()
        .setTitle('Interactive Training')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
      // Log detailed error if presentation access fails.
      console.error(`Error in doGet: Failed to open Presentation ID '${presentationId}'. Error: ${err.message}\nStack: ${err.stack}`);
      // Return a user-friendly error message.
      return HtmlService.createHtmlOutput(`Error: Could not load the presentation (ID: ${presentationId}). Please check the ID and ensure the script has permission to access it. See script logs for more details.`);
  }
}


/**
 * Fetches data for a specific slide to be displayed in the web app.
 * Includes global overlay style settings.
 * Called by the WebApp.html JavaScript (loadSlide).
 * @param {string} presentationId The ID of the Google Slides presentation.
 * @param {number} slideIndex The 0-based index of the slide to fetch.
 * @returns {object} An object containing slide data (ID, index, total, notes, elements, background, dimensions, globalOverlayDefaults) or an error.
 */
function getSlideDataForWebApp(presentationId, slideIndex) {
  console.log(`Entering getSlideDataForWebApp for ID: ${presentationId}, Index: ${slideIndex}`);


  // Validate parameters
  if (!presentationId) {
    console.error("getSlideDataForWebApp: Missing presentation ID parameter");
    return { error: "Presentation ID is required." };
  }
  slideIndex = parseInt(slideIndex, 10); // Ensure slideIndex is a number
  if (isNaN(slideIndex) || slideIndex < 0) {
    console.log("getSlideDataForWebApp: Invalid or missing slide index, defaulting to 0");
    slideIndex = 0; // Default to first slide if invalid
  }


  try {
    const presentation = SlidesApp.openById(presentationId);
    const slides = presentation.getSlides();
    console.log(`[GS] Found ${slides.length} slides.`);


    if (slideIndex >= slides.length) {
      console.error(`[GS] Invalid slide index: ${slideIndex}. Total slides: ${slides.length}`);
      return { error: `Invalid slide index (${slideIndex}). Presentation only has ${slides.length} slides.`, totalSlides: slides.length };
    }


    const slide = slides[slideIndex];
    console.log(`[GS] Processing slide ID: ${slide.getObjectId()}`);


    // Get global overlay settings
    const globalOverlayResult = getGlobalOverlaySettings();
    if (!globalOverlayResult.success) {
      console.warn("Could not retrieve global overlay settings, using defaults.");
      // Use defaults if retrieval failed, but don't stop the process
    }
    const globalOverlayDefaults = globalOverlayResult.settings || OVERLAY_DEFAULTS;


    // Prepare the data object to return
    const slideData = {
      slideId: slide.getObjectId(),
      currentSlideIndex: slideIndex,
      totalSlides: slides.length,
      slideNotes: "", // Initialize notes
      elements: [], // Initialize elements array
      backgroundUrl: null, // Initialize background URL
      slideWidth: presentation.getPageWidth(),
      slideHeight: presentation.getPageHeight(),
      // Include global overlay settings defaults
      globalOverlayDefaults: globalOverlayDefaults,
      // Deprecated fields - remove eventually
      // overlayOpacity: globalOverlayDefaults.opacity,
      // overlayShadow: globalOverlayDefaults.outlineEnabled
    };


    // Get background image data URL
    try {
        slideData.backgroundUrl = getSlideBackgroundAsDataUrl(slide);
        console.log(`[GS] Background URL fetched: ${slideData.backgroundUrl ? 'Success' : 'null'}`);
    } catch (bgError) {
         console.error(`[GS] Error fetching background: ${bgError}`);
         // Continue without background if it fails
    }


    // Get slide notes
    try {
      const notesPage = slide.getNotesPage();
      if (notesPage) {
        const speakerNotesShape = notesPage.getSpeakerNotesShape();
        if (speakerNotesShape) {
          slideData.slideNotes = speakerNotesShape.getText().asString();
          console.log("[GS] Notes fetched successfully.");
        }
      }
    } catch(notesError) {
      console.warn("[GS] Could not retrieve speaker notes: " + notesError);
    }


    // Process page elements to find interactive ones
    const pageElements = slide.getPageElements();
    console.log(`[GS] Found ${pageElements.length} elements on slide ${slideIndex}`);


    pageElements.forEach((element) => {
      const desc = element.getDescription();
      let elementDataJson = null;


      // Check if description contains JSON interaction data
      if (desc && desc.trim().startsWith('{') && desc.trim().endsWith('}')) {
        try {
          elementDataJson = JSON.parse(desc); // Keep as parsed JSON object


          // Get text content if available (best effort)
          let textContent = null;
          try {
            if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE && element.asShape().getText) {
              textContent = element.asShape().getText().asString();
            } else if (element.getPageElementType() === SlidesApp.PageElementType.TEXT_BOX && element.asTextBox().getText) {
              textContent = element.asTextBox().getText().asString();
            }
          } catch(textError){
            console.warn(`[GS] Could not get text for ${element.getObjectId()}: ${textError.message}`);
          }


          // Check if there's any interaction or animation data block defined
           const hasInteractionData = elementDataJson.interaction && typeof elementDataJson.interaction === 'object';
           const hasAnimationData = elementDataJson.animation && typeof elementDataJson.animation === 'object';


           if (hasInteractionData || hasAnimationData) { // Include if it has either data block
            // Add element data to the results array
            // The interaction and animation properties will contain the parsed objects directly
            // including showOverlayText, overlayText, overlayStyle etc. if present
            slideData.elements.push({
              id: element.getObjectId(),
              type: element.getPageElementType().toString(), // SHAPE, TEXT_BOX etc.
              left: element.getLeft(),
              top: element.getTop(),
              width: element.getWidth(),
              height: element.getHeight(),
              text: textContent, // Include element's own text content if found
              interaction: elementDataJson.interaction || null, // Include full interaction object
              animation: elementDataJson.animation || null // Include full animation object
            });


            // More detailed logging
            let logMsg = `[GS] Added element: ID=${element.getObjectId()}`;
            if (hasInteractionData) {
                logMsg += `, Interaction=${elementDataJson.interaction.type ?? 'N/A'}`;
                if (elementDataJson.interaction.showOverlayText) logMsg += ` [OText]`;
                if (elementDataJson.interaction.useCustomOpacity) logMsg += ` [COpacity]`;
                if (elementDataJson.interaction.overlayStyle) logMsg += ` [OStyle]`;
            }
             if (hasAnimationData) {
                 logMsg += `, Animation=${elementDataJson.animation.type ?? 'N/A'}`;
             }
             console.log(logMsg);
          }
        } catch (e) {
          console.warn(`[GS] Element ${element.getObjectId()} has invalid JSON: ${e.message}`);
        }
      }
    });


    console.log(`[GS] Finished processing slide ${slideIndex}. Found ${slideData.elements.length} elements with data. Global Overlay Defaults included.`);
    return slideData; // Return the compiled slide data


  } catch (e) {
    // Catch errors during presentation opening or processing
    console.error(`Error in getSlideDataForWebApp (ID: ${presentationId}, Index: ${slideIndex}): ${e.toString()}\nStack: ${e.stack}`);
    return { error: `An error occurred while loading slide data: ${e.message}. Check script logs.` };
  }
}




/**
 * Attempts to get the background of a slide as a Base64 Data URL.
 * Uses multiple methods for maximum compatibility.
 * @param {Slide} slide The Google Slides slide object.
 * @returns {string|null} Base64 Data URL of the background or null if not found/error.
 */
function getSlideBackgroundAsDataUrl(slide) {
    try {
        // console.log(`[GS] Attempting to get background for slide ${slide.getObjectId()}...`); // Less verbose logging
        const background = slide.getBackground();


        // Method 1: Try to get image using getPictureFill()
        try {
            const fill = background.getPictureFill();
            if (fill) {
                const blob = fill.getBlob();
                if (blob) {
                    // console.log("[GS] Found background via getPictureFill()");
                    return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
                }
            }
        } catch (e) {
            // console.log("[GS] Method 1 (getPictureFill) failed: " + e.message); // Optional logging
        }


        // Method 2: Check if it's a solid fill color (will return null)
        try {
            const solidFill = background.getSolidFill();
            if (solidFill) {
                // console.log("[GS] Background is a solid color, no image URL.");
                return null; // Indicate no image background
            }
        } catch(e) {
            // console.log("[GS] Method 2 (getSolidFill) check failed: " + e.message); // Optional logging
        }


        // Method 3: Use Slides API to get thumbnail as fallback (Requires Advanced Slides Service enabled)
        try {
            // Check if Slides API is available (avoids errors if not enabled)
            if (typeof Slides !== 'undefined' && Slides.Presentations && Slides.Presentations.Pages) {
                // console.log("[GS] Attempting background via Slides API thumbnail...");


                // Get presentation and slide IDs
                const presentationId = slide.getParent().getId();
                const slideId = slide.getObjectId();


                // Request thumbnail from Slides API
                const thumbnail = Slides.Presentations.Pages.getThumbnail(
                    presentationId,
                    slideId,
                    { 'thumbnailProperties.thumbnailSize': 'LARGE' } // Request a larger size
                );


                if (thumbnail && thumbnail.contentUrl) {
                    // console.log("[GS] Retrieved thumbnail URL from Slides API");
                    // Fetch the image URL and convert to Base64
                    // Note: Requires urlFetchScope permission
                    const response = UrlFetchApp.fetch(thumbnail.contentUrl);
                    const blob = response.getBlob();
                    if (blob) {
                        // console.log("[GS] Successfully fetched blob from thumbnail URL");
                        return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
                    }
                }
            } else {
                // console.log("[GS] Slides Advanced Service not enabled or available.");
            }
        } catch (e) {
            console.error("[GS] Method 3 (Slides API) failed: " + e.message);
            if (e.message && e.message.includes("requires authentication")) {
                console.error("[GS] Hint: Ensure 'Slides API' advanced service is enabled in Apps Script editor and permissions granted.");
            }
        }


        console.log("[GS] No background image could be retrieved for slide: " + slide.getObjectId());
        return null; // No method worked
    } catch (e) {
        console.error(`[GS] General error getting background for slide ${slide.getObjectId()}: ${e}`);
        return null;
    }
}


// --- Utility & Debug Functions ---


/**
 * Converts a hex color string (#RRGGBB or #RGB) to an RGB object for the Slides API.
 * @param {string} hex Hex color string.
 * @returns {object} {red, green, blue} object with values 0-1.
 */
function hexToRgb(hex) {
   // Implementation from previous version - looks correct
     hex = hex.replace(/^#/, '');
     let r = 0, g = 0, b = 0;
     if (hex.length === 3) {
        r = parseInt(hex.charAt(0) + hex.charAt(0), 16) / 255;
        g = parseInt(hex.charAt(1) + hex.charAt(1), 16) / 255;
        b = parseInt(hex.charAt(2) + hex.charAt(2), 16) / 255;
     } else if (hex.length === 6) {
        r = parseInt(hex.substring(0, 2), 16) / 255;
        g = parseInt(hex.substring(2, 4), 16) / 255;
        b = parseInt(hex.substring(4, 6), 16) / 255;
     }
     return { red: r, green: g, blue: b };
}


// --- Deployment ID Functions ---


/**
 * Shows a prompt dialog for the user to enter their Web App deployment ID.
 */
function setDeploymentId() {
    const ui = SlidesApp.getUi();
    ui.alert(
        'Set Deployment ID',
        'After deploying as a web app, you\'ll get a URL like:\n\n' +
        'https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxxxxxxxxxxxxx/exec\n\n' +
        'In the next prompt, paste ONLY the ID part (after /s/ and before /exec).',
        ui.ButtonSet.OK
    );
    const response = ui.prompt(
        'Enter Deployment ID',
        'Enter the deployment ID from your web app URL:',
        ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() === ui.Button.OK) {
        const deploymentId = response.getResponseText().trim();
        if (!deploymentId) {
            ui.alert('Error', 'Please enter a valid deployment ID.', ui.ButtonSet.OK);
            return;
        }
        // Basic validation
        if (!/^[a-zA-Z0-9_-]+$/.test(deploymentId) || deploymentId.length < 20) {
             ui.alert('Error', 'The entered ID seems invalid. Please paste the correct ID.', ui.ButtonSet.OK);
             return;
        }
        try {
            PropertiesService.getScriptProperties().setProperty('WEBAPP_DEPLOYMENT_ID', deploymentId);
            ui.alert('Success', 'Deployment ID saved successfully!', ui.ButtonSet.OK);
        } catch (e) {
            ui.alert('Error', 'Failed to save deployment ID: ' + e.message, ui.ButtonSet.OK);
        }
    }
}


/**
 * Generates and displays the shareable URL for the web app, including the current presentation ID.
 */
function showShareableUrl() {
     try {
        const presentationId = SlidesApp.getActivePresentation().getId();
        let deploymentId = PropertiesService.getScriptProperties().getProperty('WEBAPP_DEPLOYMENT_ID');


        if (!deploymentId) {
            SlidesApp.getUi().alert(
                'No Web App deployment found',
                'You need to set the deployment ID first.\nGo to Extensions > Onboarding Tools > Set Web App Deployment ID.',
                SlidesApp.getUi().ButtonSet.OK
            );
            return;
        }


        const baseUrl = `https://script.google.com/macros/s/${deploymentId}/exec`;
        const shareableUrl = `${baseUrl}?presId=${presentationId}`;


        const ui = SlidesApp.getUi();
        // Use a slightly larger modal for better readability
        const htmlContent = `<div style="font-family: Arial, sans-serif; padding: 10px;">
                               <p>Share this URL for the interactive presentation:</p>
                               <textarea rows="3" style="width:98%; margin-top:10px; font-family: monospace; font-size: 12px; border: 1px solid #ccc; border-radius: 4px; padding: 5px;" readonly onclick="this.select();">${shareableUrl}</textarea>
                               <br/><br/>
                               <button onclick="google.script.host.close()" style="padding: 8px 15px; background-color: #4285F4; color: white; border: none; border-radius: 4px; cursor: pointer;">Close</button>
                             </div>`;
        const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
            .setWidth(500) // Set width
            .setHeight(180); // Set height
        ui.showModalDialog(htmlOutput, 'Shareable URL');


    } catch (e) {
        console.error("Error generating shareable URL: " + e);
        SlidesApp.getUi().alert('Error generating URL: ' + e.message);
    }
}


// --- Diagnostic Functions (Unchanged) ---


/**
 * Creates a simple diagnostic page useful for debugging presentation rendering.
 * Shows information about the presentation, slides, and elements without trying to render them.
 * @param {object} e The event parameter for doGet, contains URL parameters.
 * @returns {HtmlOutput} The diagnostic HTML page.
 */
function serveDiagnosticPage(e) {
  const presentationId = e?.parameter?.presId || "";


  let htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Interactive Training - Diagnostic</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
        h1, h2, h3 { color: #1a73e8; }
        .section { margin-bottom: 20px; padding: 15px; border: 1px solid #e0e0e0; border-radius: 8px; }
        .info-item { margin-bottom: 10px; }
        .label { font-weight: bold; color: #666; }
        pre { background-color: #f5f5f5; padding: 10px; border-radius: 4px; overflow-x: auto; white-space: pre-wrap; word-wrap: break-word; }
        button { background-color: #1a73e8; color: white; border: none; padding: 8px 16px;
                border-radius: 4px; cursor: pointer; margin-right: 10px; margin-top: 5px; }
        button:hover { background-color: #1558b7; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }
        th { background-color: #f2f2f2; }
        .slide-thumbnail { max-width: 200px; height: auto; margin-right: 10px; border: 1px solid #ddd; vertical-align: top; }
        #slideDetails { border-top: 1px solid #eee; margin-top: 15px; padding-top: 15px; }
      </style>
    </head>
    <body>
      <h1>Interactive Training Diagnostic Page</h1>


      <div class="section">
        <h2>Presentation Information</h2>
        <div class="info-item">
          <span class="label">Requested Presentation ID:</span> ${presentationId || '<i style="color:red;">Not Provided!</i>'}
        </div>
        <div id="presentationDetails">Loading presentation details...</div>
        <button id="loadDetailsBtn" ${!presentationId ? 'disabled' : ''}>Load Presentation Details</button>
      </div>


      <div class="section">
        <h2>Slide Information</h2>
        <div class="info-item">
          <label for="slideIndex" style="margin-right: 5px;">Select Slide:</label>
          <select id="slideIndex" disabled>
            <option value="">Load Presentation First</option>
          </select>
          <button id="loadSlideBtn" disabled>Load Slide Details</button>
        </div>
        <div id="slideDetails">(No slide loaded)</div>
      </div>


      <div class="section">
        <h2>Environment Information</h2>
        <div class="info-item">
          <span class="label">User Agent:</span> <span id="userAgent"></span>
        </div>
        <div class="info-item">
          <span class="label">Window Size:</span> <span id="windowSize"></span>
        </div>
        <div class="info-item">
          <span class="label">Time Zone (Script):</span> ${Session.getScriptTimeZone()}
        </div>
      </div>


      <script>
        const currentPresentationId = "${presentationId}";


        // Display environment info immediately
        document.getElementById('userAgent').textContent = navigator.userAgent;
        document.getElementById('windowSize').textContent = window.innerWidth + 'x' + window.innerHeight;


        const loadDetailsBtn = document.getElementById('loadDetailsBtn');
        const loadSlideBtn = document.getElementById('loadSlideBtn');
        const slideSelect = document.getElementById('slideIndex');
        const presentationDetailsDiv = document.getElementById('presentationDetails');
        const slideDetailsDiv = document.getElementById('slideDetails');


        // Load presentation details
        loadDetailsBtn.addEventListener('click', function() {
          this.disabled = true;
          this.textContent = 'Loading...';
          presentationDetailsDiv.innerHTML = 'Loading...';
          slideSelect.innerHTML = '<option value="">Loading...</option>';
          slideSelect.disabled = true;
          loadSlideBtn.disabled = true;
          slideDetailsDiv.innerHTML = '(No slide loaded)';




          google.script.run
            .withSuccessHandler(function(result) {
              if (result.error) {
                presentationDetailsDiv.innerHTML = '<p style="color: red;">Error loading presentation: ' + result.error + '</p>';
                loadDetailsBtn.textContent = 'Load Presentation Details';
                loadDetailsBtn.disabled = false;
                return;
              }


              let html = '<div class="info-item"><span class="label">Title:</span> ' + (result.title || 'Untitled') + '</div>';
              html += '<div class="info-item"><span class="label">Total Slides:</span> ' + result.totalSlides + '</div>';
              html += '<div class="info-item"><span class="label">Dimensions:</span> ' + result.width + 'x' + result.height + '</div>';
              presentationDetailsDiv.innerHTML = html;


              // Populate slide dropdown
              slideSelect.innerHTML = ''; // Clear previous options
              if (result.totalSlides > 0) {
                 for (let i = 0; i < result.totalSlides; i++) {
                   const option = document.createElement('option');
                   option.value = i;
                   option.textContent = 'Slide ' + (i + 1);
                   slideSelect.appendChild(option);
                 }
                 slideSelect.disabled = false;
                 loadSlideBtn.disabled = false;
              } else {
                 slideSelect.innerHTML = '<option value="">No Slides Found</option>';
              }


              loadDetailsBtn.textContent = 'Reload Presentation Details';
              loadDetailsBtn.disabled = false;
            })
            .withFailureHandler(function(error) {
              presentationDetailsDiv.innerHTML =
                '<p style="color: red;">Failed to load presentation details: ' + error + '</p>';
              loadDetailsBtn.textContent = 'Load Presentation Details';
              loadDetailsBtn.disabled = false; // Re-enable button on failure
            })
            .getDiagnosticPresentationInfo(currentPresentationId);
        });


        // Load slide details
        loadSlideBtn.addEventListener('click', function() {
          const selectedIndex = slideSelect.value;
          if (selectedIndex === "") return; // No slide selected


          this.disabled = true;
          this.textContent = 'Loading Slide...';
          slideDetailsDiv.innerHTML = 'Loading slide details...';


          google.script.run
            .withSuccessHandler(function(result) {
              if (result.error) {
                slideDetailsDiv.innerHTML = '<p style="color: red;">Error loading slide: ' + result.error + '</p>';
                loadSlideBtn.textContent = 'Load Slide Details';
                loadSlideBtn.disabled = false;
                return;
              }


              let html = '<div class="info-item"><span class="label">Slide ID:</span> ' + result.slideId + '</div>';
              html += '<div class="info-item"><span class="label">Slide Index:</span> ' + result.currentSlideIndex + '</div>';
              html += '<div class="info-item"><span class="label">Notes:</span> <pre>' + (result.slideNotes || 'None') + '</pre></div>';
              // Display global overlay defaults
              if (result.globalOverlayDefaults) {
                html += '<h3>Global Overlay Defaults</h3>';
                html += '<pre>' + JSON.stringify(result.globalOverlayDefaults, null, 2) + '</pre>';
              } else {
                html += '<p style="color: orange;">Global overlay defaults missing.</p>';
              }
              // html += '<div class="info-item"><span class="label">Overlay Opacity:</span> ' + result.overlayOpacity + '%</div>'; // Deprecated
              // html += '<div class="info-item"><span class="label">Overlay Shadow:</span> ' + result.overlayShadow + '</div>'; // Deprecated


              html += '<div class="info-item"><span class="label">Background:</span><br>';
              if (result.backgroundUrl) {
                html += '<img src="' + result.backgroundUrl + '" class="slide-thumbnail" alt="Slide Background Preview"></div>';
              } else {
                html += '<span>Solid color or no background image</span></div>';
              }


              html += '<h3>Interactive Elements (' + (result.elements?.length || 0) + ')</h3>';


              if (result.elements && result.elements.length > 0) {
                html += '<table><thead><tr><th>ID</th><th>Type</th><th>Pos (L,T)</th><th>Size (W,H)</th><th>Interaction</th><th>Animation</th></tr></thead><tbody>';


                result.elements.forEach(element => {
                  html += '<tr>';
                  html += '<td>' + element.id.substring(0,8) + '...</td>';
                  html += '<td>' + element.type + '</td>';
                  html += '<td>' + Math.round(element.left) + ', ' + Math.round(element.top) + '</td>';
                  html += '<td>' + Math.round(element.width) + ', ' + Math.round(element.height) + '</td>';


                  const interactionDetails = element.interaction ? JSON.stringify(element.interaction, null, 2) : 'None';
                  html += '<td><pre>' + interactionDetails + '</pre></td>';


                  const animationDetails = element.animation ? JSON.stringify(element.animation, null, 2) : 'None';
                  html += '<td><pre>' + animationDetails + '</pre></td>';


                  html += '</tr>';
                });


                html += '</tbody></table>';
              } else {
                html += '<p>No interactive elements found on this slide.</p>';
              }


              slideDetailsDiv.innerHTML = html;
              loadSlideBtn.textContent = 'Reload Slide Details';
              loadSlideBtn.disabled = false;
            })
            .withFailureHandler(function(error) {
              slideDetailsDiv.innerHTML =
                '<p style="color: red;">Failed to load slide details: ' + error + '</p>';
              loadSlideBtn.textContent = 'Load Slide Details';
               loadSlideBtn.disabled = false; // Re-enable button on failure
            })
            .getSlideDataForWebApp(currentPresentationId, selectedIndex); // Use the main function here
        });


         // Trigger loading presentation details automatically if ID is present
         if(currentPresentationId) {
           loadDetailsBtn.click();
         }


      </script>
    </body>
    </html>
  `;


  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Interactive Training - Diagnostic')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/**
 * Gets presentation information for diagnostic purposes.
 * @param {string} presentationId The ID of the presentation to analyze.
 * @returns {object} Basic information about the presentation.
 */
function getDiagnosticPresentationInfo(presentationId) {
  if (!presentationId) {
     return { error: "No Presentation ID provided in the URL (?presId=...)." };
  }
  try {
    const presentation = SlidesApp.openById(presentationId);
    return {
      success: true,
      title: presentation.getName(),
      totalSlides: presentation.getSlides().length,
      width: presentation.getPageWidth(),
      height: presentation.getPageHeight()
    };
  } catch (e) {
    console.error(`Error in getDiagnosticPresentationInfo for ID '${presentationId}': ${e.message}`);
    let errorMsg = `Failed to get presentation info: ${e.message}.`;
     if (e.message && e.message.toLowerCase().includes('not found')) {
        errorMsg += ' Check if the Presentation ID is correct and exists.';
     } else if (e.message && e.message.toLowerCase().includes('access denied')) {
         errorMsg += ' Ensure the script has permission to access this presentation.';
     }
    return { error: errorMsg };
  }
}


/**
 * Debugging function to inspect element description directly.
 * Called from Sidebar.html for troubleshooting.
 * @param {string} elementId The ID of the element to inspect.
 * @returns {object} Debug information about the element's description.
 */
function debugElementDescription(elementId) {
  try {
    if (!elementId) return { success: false, error: "No element ID provided." };


    const presentation = SlidesApp.getActivePresentation();
    let element = null;
    let slideId = null;
    let slideIndex = -1;


    const allSlides = presentation.getSlides();
    for (let i = 0; i < allSlides.length; i++) {
        const s = allSlides[i];
        const el = s.getPageElementById(elementId);
        if (el) {
            element = el;
            slideId = s.getObjectId();
            slideIndex = i;
            break;
        }
    }


    if (!element) {
        return { success: false, error: `Element ID ${elementId} not found in presentation.` };
    }


    const description = element.getDescription();
    let parsedData = null;
    let parseError = null;
    let parsedSuccessfully = false;
    let hasInteractionField = false;
    let interactionType = 'N/A';
    let hasAnimationField = false;
    let animationType = 'N/A';
    let hasOverlayStyleField = false; // New check


    if (description && description.trim() !== "") {
        try {
            parsedData = JSON.parse(description);
            parsedSuccessfully = true;
            if (parsedData && typeof parsedData === 'object') {
                 hasInteractionField = parsedData.hasOwnProperty('interaction');
                 if (hasInteractionField && parsedData.interaction && typeof parsedData.interaction === 'object') {
                    interactionType = parsedData.interaction.type || 'Type missing';
                    hasOverlayStyleField = parsedData.interaction.hasOwnProperty('overlayStyle'); // Check within interaction
                 } else if (hasInteractionField) {
                    interactionType = 'Invalid format';
                 } else {
                    interactionType = 'Not present';
                 }


                 hasAnimationField = parsedData.hasOwnProperty('animation');
                 if (hasAnimationField && parsedData.animation && typeof parsedData.animation === 'object') {
                    animationType = parsedData.animation.type || 'Type missing';
                 } else if (hasAnimationField) {
                     animationType = 'Invalid format';
                 } else {
                     animationType = 'Not present';
                 }
            }
        } catch (e) {
            parseError = e.message;
        }
    }


    return {
        success: true,
        elementId: elementId,
        slideId: slideId,
        slideIndex: slideIndex,
        hasDescription: !!description,
        descriptionLength: description ? description.length : 0,
        rawDescription: description, // Include raw description for inspection
        parsedSuccessfully: parsedSuccessfully,
        parseError: parseError,
        hasInteractionField: hasInteractionField,
        interactionType: interactionType,
        hasAnimationField: hasAnimationField,
        animationType: animationType,
        hasOverlayStyleField: hasOverlayStyleField, // Include overlay style check
        parsedData: parsedData // Include parsed data if successful
    };


  } catch (e) {
      console.error(`Error in debugElementDescription for ${elementId}: ${e}`);
      return { success: false, error: `Debug failed: ${e.message}` };
  }
}
