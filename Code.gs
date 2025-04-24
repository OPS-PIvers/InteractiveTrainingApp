/**
 * @OnlyCurrentDoc
 */


// --- Add-on Menu & Sidebar ---

function onOpen() {
  const ui = SlidesApp.getUi();
  
  // Regular menu for all users
  const menu = ui.createMenu('Onboarding Tools')
      .addItem('Show Builder Panel', 'showSidebar')
      .addItem('Get Interactive Viewer Link', 'showViewerLink');
  
  // Check if running as owner/admin and add admin options
  const email = Session.getEffectiveUser().getEmail();
  const ownerEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
  
  if (!ownerEmail) {
    // First run, assume current user is admin
    PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', email);
    menu.addSeparator()
        .addItem('[Admin] Set Master Deployment ID', 'setMasterDeploymentId');
  } else if (email === ownerEmail) {
    // Admin user, show admin options
    menu.addSeparator()
        .addItem('[Admin] Set Master Deployment ID', 'setMasterDeploymentId');
  }
  
  menu.addToUi();
}

// --- Navigation Sequence Functions ---

/**
 * Property key constants for storing navigation sequences
 */
const NAV_SEQUENCE_PREFIX = 'SLIDE_NAV_SEQUENCE_';

/**
 * Saves a custom navigation sequence order for a specific slide
 * @param {string} slideId The ID of the slide
 * @param {string[]} elementIds Array of element IDs in the desired sequence order
 * @returns {object} Success/error result
 */
function saveNavigationSequence(slideId, elementIds) {
  try {
    if (!slideId) {
      return { success: false, error: "Missing slide ID" };
    }
    
    if (!Array.isArray(elementIds)) {
      return { success: false, error: "Element IDs must be provided as an array" };
    }
    
    // Check if elements exist
    if (elementIds.length === 0) {
      // User is clearing the sequence
      PropertiesService.getDocumentProperties().deleteProperty(NAV_SEQUENCE_PREFIX + slideId);
      console.log(`Cleared navigation sequence for slide: ${slideId}`);
      return { 
        success: true, 
        message: "Navigation sequence cleared" 
      };
    }

    // Store the sequence as a JSON string in document properties
    const sequenceData = JSON.stringify(elementIds);
    PropertiesService.getDocumentProperties().setProperty(NAV_SEQUENCE_PREFIX + slideId, sequenceData);
    
    console.log(`Saved navigation sequence for slide ${slideId}: ${sequenceData}`);
    return { 
      success: true, 
      message: "Navigation sequence saved successfully" 
    };
  } catch (e) {
    console.error(`Error saving navigation sequence: ${e.message}`);
    return { 
      success: false, 
      error: `Failed to save navigation sequence: ${e.message}` 
    };
  }
}

/**
 * Gets the saved navigation sequence for a specific slide
 * @param {string} slideId The ID of the slide
 * @returns {object} Object with success flag and either ordered element IDs or error
 */
function getNavigationSequence(slideId) {
  try {
    if (!slideId) {
      return { success: false, error: "Missing slide ID" };
    }
    
    const sequenceData = PropertiesService.getDocumentProperties().getProperty(NAV_SEQUENCE_PREFIX + slideId);
    
    if (!sequenceData) {
      // No custom sequence found - this is normal, not an error
      return { 
        success: true, 
        hasCustomSequence: false,
        elementIds: [] 
      };
    }
    
    // Parse the stored JSON string
    const elementIds = JSON.parse(sequenceData);
    
    return { 
      success: true,
      hasCustomSequence: true,
      elementIds: elementIds
    };
  } catch (e) {
    console.error(`Error getting navigation sequence: ${e.message}`);
    return { 
      success: false, 
      error: `Failed to get navigation sequence: ${e.message}`
    };
  }
}

/**
 * Clears the saved navigation sequence for a specific slide
 * @param {string} slideId The ID of the slide
 * @returns {object} Success/error result
 */
function clearNavigationSequence(slideId) {
  try {
    if (!slideId) {
      return { success: false, error: "Missing slide ID" };
    }
    
    PropertiesService.getDocumentProperties().deleteProperty(NAV_SEQUENCE_PREFIX + slideId);
    console.log(`Cleared navigation sequence for slide: ${slideId}`);
    
    return { 
      success: true, 
      message: "Navigation sequence cleared successfully" 
    };
  } catch (e) {
    console.error(`Error clearing navigation sequence: ${e.message}`);
    return { 
      success: false, 
      error: `Failed to clear navigation sequence: ${e.message}` 
    };
  }
}

function setMasterDeploymentId() {
  const ui = SlidesApp.getUi();
  const response = ui.prompt(
    'Set Master Deployment ID',
    'Enter the deployment ID from your web app deployment.\n\n' +
    'You can find this in the Google Apps Script editor by:\n' +
    '1. Go to Deploy > New deployment\n' +
    '2. Choose "Web app"\n' +
    '3. After deploying, copy the ID portion from the URL\n' +
    '   (e.g., "AKfycbxYgN..." from "https://script.google.com/macros/s/AKfycbxYgN.../exec")',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const deploymentId = response.getResponseText().trim();
    if (deploymentId) {
      PropertiesService.getScriptProperties().setProperty('MASTER_WEBAPP_ID', deploymentId);
      ui.alert('Success', 'Master deployment ID saved! Now anyone can use the "Get Interactive Viewer Link" menu item to generate a sharing link for their presentation.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No deployment ID was entered. Please try again with a valid deployment ID.', ui.ButtonSet.OK);
    }
  }
}

function getMasterWebAppUrl() {
  const deploymentId = PropertiesService.getScriptProperties().getProperty('MASTER_WEBAPP_ID');
  if (deploymentId) {
    return `https://script.google.com/macros/s/${deploymentId}/exec`;
  }
  return null;
}

function showViewerLink() {
  try {
    const presentationId = SlidesApp.getActivePresentation().getId();
    const masterUrl = getMasterWebAppUrl();
    
    if (!masterUrl) {
      // Guide the user through setup if the master web app URL is not configured
      const ui = SlidesApp.getUi();
      const result = ui.alert(
        'Setup Required',
        'The master web app URL has not been configured yet.\n\n' +
        'This is a one-time setup that needs to be done by the admin to enable the interactive viewer functionality.\n\n' +
        'If you are the admin, click "Yes" to set up the master deployment ID now.',
        ui.ButtonSet.YES_NO
      );
      
      // If user clicked "Yes" and is the admin, open the setup dialog
      if (result === ui.Button.YES) {
        const email = Session.getEffectiveUser().getEmail();
        const ownerEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
        
        if (email === ownerEmail || !ownerEmail) {
          // User is admin, show the deployment ID setup dialog
          setMasterDeploymentId();
        } else {
          // User is not admin, show message
          ui.alert(
            'Admin Access Required',
            'Sorry, only the admin (' + ownerEmail + ') can set up the master deployment ID.\n\n' +
            'Please contact the admin to complete this setup.',
            ui.ButtonSet.OK
          );
        }
      }
      return;
    }
    
    // Format the complete URL with this presentation's ID
    const viewerUrl = `${masterUrl}?presId=${presentationId}`;
    
    // Show the dialog with the link
    const ui = SlidesApp.getUi();
    const htmlContent = `<div style="font-family: Arial, sans-serif; padding: 10px;">
                        <p>Share this URL for the interactive presentation:</p>
                        <textarea rows="3" style="width:98%; margin-top:10px; font-family: monospace; font-size: 12px; border: 1px solid #ccc; border-radius: 4px; padding: 5px;" readonly onclick="this.select();">${viewerUrl}</textarea>
                        <p style="font-size: 12px; color: #666; margin-top: 8px;">
                          Anyone with this link can access the interactive version of your presentation
                          (regular sharing permissions still apply for the original presentation).
                        </p>
                        <button onclick="google.script.host.close()" style="padding: 8px 15px; background-color: #4285F4; color: white; border: none; border-radius: 4px; cursor: pointer;">Close</button>
                      </div>`;
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(500)
        .setHeight(220);
    ui.showModalDialog(htmlOutput, 'Interactive Viewer Link');
  } catch (e) {
    console.error("Error generating viewer link: " + e);
    SlidesApp.getUi().alert('Error generating link: ' + e.message);
  }
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
        return { error: "Could not determine the element's slide. Please try selecting the element again." };
      }

      const slideId = slide.getObjectId();
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
        canGoNext: canGoNext,  // Flag for next element navigation
        slideId: slideId  // Add the slide ID for context/saving
      };
    } else if (pageElements.length > 1) {
      // Error if multiple elements are selected.
      return { error: "Please select only one element. Multiple elements are currently selected." };
    }
  } else if (selectionType === SlidesApp.SelectionType.CURRENT_PAGE) {
    // Error if the entire slide is selected instead of a specific element.
    return { error: "Please select a specific element (shape, text box, image) rather than the entire slide." };
  } else if (selectionType === SlidesApp.SelectionType.NONE) {
    // Nothing is selected
    return { error: "No element is selected. Please click on a shape, text box, or image to select it.", canGoPrev: false, canGoNext: false };
  }
  
  // Default error if nothing suitable is selected.
  // Also return nav flags as false in this case.
  return { error: "Please select a single element on the slide (shape, text box, image) to configure interactions.", canGoPrev: false, canGoNext: false };
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
    // Remove sleep and immediately return element info
    return getSelectedElementInfo();
  } else {
    // Already at the first element, cannot go previous
    return getSelectedElementInfo();
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
    // Remove sleep and immediately return element info
    return getSelectedElementInfo();
  } else {
    // Already at the last element, cannot go next
    return getSelectedElementInfo();
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
       return { success: false, error: "No slide or element selected. Please select the slide containing your element." };
   }

   // Get the current slide from the selection.
   const slide = selection.getCurrentPage();
   if (!slide) {
       return { success: false, error: "Could not determine the active slide. Please select a slide." };
   }

   // Find the element on the current slide using its ID.
   const element = slide.getPageElementById(elementId);
   if (element) {
       // Call helper function to save the description
       return saveDescription(element, interactionDataJson);
   } else {
       // Element not found on the currently selected slide.
       return { success: false, error: "Element not found on the current slide. Please ensure the correct slide is active." };
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
 * Handles 'overlayText', 'showOverlayText', 'useCustomOpacity', 'customOpacity', 'overlayStyle',
 * 'disappearOnClick', 'appearanceBehavior', and 'appearAfterElementId' within the interaction object.
 *
 * @param {string} elementId The ID of the element to modify.
 * @param {object} newData The new data to merge (e.g., { type: 'showText', text: '...', overlayStyle: {...} }).
 * @param {string} dataType Either 'interaction' or 'animation' to indicate which part to update.
 * @returns {object} An object indicating success or failure, with message or error.
 */
function mergeElementData(elementId, newData, dataType) {
  try {
    // Get the current element description to check for existing data
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();

    // Ensure there's a selection context to find the slide
    if (!selection || selection.getSelectionType() === SlidesApp.SelectionType.NONE) {
      return { success: false, error: "No slide or element selected. Please select the slide containing your element." };
    }

    // Get the current slide from the selection
    const slide = selection.getCurrentPage();
    if (!slide) {
      return { success: false, error: "Could not determine the active slide. Please select a slide." };
    }

    // Find the element on the current slide using its ID
    const element = slide.getPageElementById(elementId);
    if (!element) {
      return { success: false, error: "Element not found on the current slide. Please ensure the correct slide is active." };
    }

    // Get current description and parse as JSON if exists
    let currentData = {};
    const description = element.getDescription();

    if (description && description.trim() !== "") {
      try {
        currentData = JSON.parse(description);
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
        
        // --- NEW: Handle appearance behavior properties ---
        if (interaction.appearanceBehavior) {
          // If the behavior is not 'afterPrevious', remove the dependency element ID
          if (interaction.appearanceBehavior !== 'afterPrevious') {
            delete interaction.appearAfterElementId;
          }
          // Otherwise ensure the ID exists if provided
          else if (interaction.appearAfterElementId === undefined || 
                  interaction.appearAfterElementId === null || 
                  interaction.appearAfterElementId === '') {
            // Provide a warning but don't block the save
            console.warn(`Element ${elementId} has 'afterPrevious' behavior but no valid target element ID`);
          }
        }
        // If no behavior specified, set to default
        else {
          interaction.appearanceBehavior = 'withPresentation';
          delete interaction.appearAfterElementId;
        }
      }
    }

    // Check if the entire description should be cleared
    // Note: An interaction might exist but have no effect (e.g., only overlayStyle) - consider this if needed.
    // For now, keep if *any* data exists in either section.
    const interactionIsEmpty = !mergedData.interaction || Object.keys(mergedData.interaction).length === 0 || mergedData.interaction.type === 'none';
    const animationIsEmpty = !mergedData.animation || Object.keys(mergedData.animation).length === 0 || mergedData.animation.type === 'none';

    if (interactionIsEmpty && animationIsEmpty) {
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
    element.setDescription(mergedJson);

    // Verify the save worked by reading it back
    const verifyDesc = element.getDescription();

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
    console.error(`Error in mergeElementData: ${e.message}`);
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
    if (!selection) { 
      return { 
        error: "No active presentation found. Please make sure you have a presentation open." 
      }; 
    }

    // Use the current page from selection, or default to the first slide
    const slide = selection.getCurrentPage() || SlidesApp.getActivePresentation().getSlides()[0];
    if (!slide) { 
      return { 
        error: "Could not determine the active slide. Please select a slide in your presentation." 
      }; 
    }

    const slideId = slide.getObjectId();
    const elements = slide.getPageElements();
    const interactiveElements = [];

    console.log(`Checking ${elements.length} elements on slide ${slideId} for interactive data`);

    // Get saved navigation sequence if it exists
    const sequenceResult = getNavigationSequence(slideId);
    const savedSequence = sequenceResult.success && sequenceResult.hasCustomSequence ? 
                          sequenceResult.elementIds : null;
    
    // Create index mapping for sequential display IDs
    const displayIdMap = new Map();

    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      const desc = element.getDescription();

      if (desc && desc.trim().startsWith('{') && desc.trim().endsWith('}')) { // Basic check for JSON
        let data = null;
        try {
          data = JSON.parse(desc);

          // Check if data is an object and has *any* interaction or animation keys
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

            // Store the element ID to assign display IDs later
            displayIdMap.set(element.getObjectId(), {
              index: interactiveElements.length,
              name: elementName
            });
            
            // Add element to list, including the raw description for client-side parsing
            interactiveElements.push({
              id: element.getObjectId(),
              name: elementName,
              elementType: element.getPageElementType().toString(),
              description: desc // Pass raw description for client parsing
            });
          }
        } catch (e) {
          console.warn(`Element ${element.getObjectId()} has invalid JSON in description: ${e.message}`);
          console.warn(`Raw description: ${desc}`);
        }
      }
    }
    
    // Assign sequential display IDs
    // First, reorder according to saved sequence if present
    let orderedElements = [...interactiveElements];
    if (savedSequence && savedSequence.length > 0) {
      // Reorder according to saved sequence
      const orderedTemp = [];
      const unorderedElements = [...interactiveElements]; // Copy to preserve original
      
      // First add elements in the saved sequence order
      savedSequence.forEach(id => {
        const elementIndex = unorderedElements.findIndex(el => el.id === id);
        if (elementIndex !== -1) {
          orderedTemp.push(unorderedElements[elementIndex]);
          unorderedElements.splice(elementIndex, 1); // Remove from unordered list
        }
      });
      
      // Then append any remaining elements (not in saved sequence)
      orderedElements = [...orderedTemp, ...unorderedElements];
    }
    
    // Assign display IDs to the final ordered list
    orderedElements.forEach((element, index) => {
      element.displayId = index + 1; // Assign sequential numbers starting from 1
    });
    
    console.log(`Found ${interactiveElements.length} interactive elements on slide ${slideId}`);
    return { 
      elements: orderedElements,
      slideId: slideId,
      hasCustomSequence: savedSequence !== null
    };
  } catch (error) {
    console.error("Error in getAllInteractiveElements:", error);
    return { 
      error: `Failed to get interactive elements: ${error.message}. Please check your slide content and try again.` 
    };
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
    
    // Ensure there's a selection context to find the slide
    if (!selection || selection.getSelectionType() === SlidesApp.SelectionType.NONE) {
      return { success: false, error: "No slide or element selected. Please select the slide containing your element." };
    }

    // Get the current slide from the selection
    const slide = selection.getCurrentPage();
    if (!slide) {
      return { success: false, error: "Could not determine the active slide. Please select a slide." };
    }

    // Find the element on the current slide using its ID
    const element = slide.getPageElementById(elementId);
    if (!element) {
      return { success: false, error: "Element not found on the current slide. Please ensure the correct slide is active." };
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
    
    // First try to find the element on the current slide
    const selection = presentation.getSelection();
    if (selection && selection.getCurrentPage()) {
      const currentSlide = selection.getCurrentPage();
      element = currentSlide.getPageElementById(elementId);
      
      if (element) {
        targetSlide = currentSlide;
        console.log(`Element ${elementId} found on the current slide.`);
      } else {
        console.log(`Element ${elementId} not found on current slide, searching all slides...`);
      }
    }

    // If not found on current slide, search all slides
    if (!element) {
      const allSlides = presentation.getSlides();
      for (let i = 0; i < allSlides.length; i++) {
        const s = allSlides[i];
        const el = s.getPageElementById(elementId);
        if (el) {
          element = el;
          targetSlide = s;
          console.log(`Element ${elementId} found on slide index ${i}`);
          break;
        }
      }
    }

    if (!element) {
      return { success: false, error: "Element not found in the presentation. It may have been deleted." };
    }
    if (!targetSlide) {
      return { success: false, error: "Could not determine the slide for the element." };
    }

    // Switch to the target slide if it's not the current one
    if (selection && selection.getCurrentPage() && 
        targetSlide.getObjectId() !== selection.getCurrentPage().getObjectId()) {
      targetSlide.selectAsCurrentPage();
    }

    // Select the specific element
    element.select();

    return { success: true, message: "Element selected successfully." };
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
    
    // Ensure there's a selection context to find the slide
    if (!selection || selection.getSelectionType() === SlidesApp.SelectionType.NONE) {
      return { success: false, error: "No slide or element selected. Please select the slide containing your element." };
    }

    // Get the current slide from the selection
    const slide = selection.getCurrentPage();
    if (!slide) {
      return { success: false, error: "Could not determine the active slide. Please select a slide." };
    }

    // Find the element on the current slide using its ID
    const element = slide.getPageElementById(elementId);
    if (!element) {
      return { success: false, error: "Element not found on the current slide. Please ensure the correct slide is active." };
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
  opacity: 0, // Default 0% (changed from 15%)
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
    const value = props.getProperty(propKey);

    // Important fix: Only use default if value is strictly null (not found)
    // This ensures '0' values (for opacity) are properly handled
    if (value === null) {
      // Value not found, use default
      settings[key] = OVERLAY_DEFAULTS[key];
      console.log(`Property ${propKey} not found, using default: ${OVERLAY_DEFAULTS[key]}`);
    } else {
      // Value found, parse it based on expected type
      const defaultValue = OVERLAY_DEFAULTS[key];
      if (typeof defaultValue === 'number') {
        const numValue = parseFloat(value);
        settings[key] = isNaN(numValue) ? defaultValue : numValue;
        console.log(`Loaded numeric property ${propKey}: ${numValue}`);
      } else if (typeof defaultValue === 'boolean') {
        settings[key] = (value === 'true');
        console.log(`Loaded boolean property ${propKey}: ${value}`);
      } else {
        // Assume string (color, shape, style, hoverText)
        settings[key] = value;
        console.log(`Loaded string property ${propKey}: ${value}`);
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


// --- Web App Functions ---

/**
 * Handles GET requests for the web app. Serves the WebApp.html file.
 * @param {object} e The event parameter for doGet, contains URL parameters.
 * @returns {HtmlOutput} The HTML page to be rendered.
 */
function doGet(e) {
  // Extract presentation ID from URL parameter 'presId'
  const presentationId = e?.parameter?.presId || "";
  console.log(`doGet triggered for presentation ID: ${presentationId}`);

  // If no presentation ID is provided, show a helpful intro page
  if (!presentationId) {
    return HtmlService.createHtmlOutput(`
      <html>
      <head>
        <title>Interactive Training - Welcome</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
          h1, h2 { color: #1a73e8; }
          .container { max-width: 800px; margin: 0 auto; }
          .section { margin-bottom: 30px; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; }
          .hint { background-color: #e8f0fe; padding: 15px; border-radius: 8px; margin-top: 20px; }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Interactive Training Viewer</h1>
          <div class="section">
            <h2>Welcome</h2>
            <p>This interactive viewer lets you view Google Slides presentations with interactive elements.</p>
            <p>To view a presentation, you need a link that includes the presentation ID.</p>
            <div class="hint">
              <p><strong>Need a viewer link?</strong> Open your presentation, click on the "Onboarding Tools" menu, 
              then select "Get Interactive Viewer Link" to get a shareable URL.</p>
            </div>
          </div>
        </div>
      </body>
      </html>
    `)
    .setTitle('Interactive Training - Welcome')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  try {
    // Try to open the presentation - this will throw an error if no access
    SlidesApp.openById(presentationId);
    console.log(`Successfully verified access to presentation: ${presentationId}`);

    // Create template and pass the presentation ID
    const template = HtmlService.createTemplateFromFile('WebApp');
    template.presentationId = presentationId;

    return template.evaluate()
        .setTitle('Interactive Training')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    console.error(`Error accessing presentation: ${err.message}`);
    
    // Provide more specific error messages
    let errorTitle = 'Error Loading Presentation';
    let errorHeading = 'Error';
    let errorDescription = `Could not load the presentation (ID: ${presentationId}).`;
    let errorSuggestion = 'The presentation may not exist or there was a problem loading it.';
    
    // Check if this is a permission error
    const isPermissionError = err.message && 
                             (err.message.includes("Access denied") || 
                              err.message.includes("permission") ||
                              err.message.includes("Insufficient"));
    
    if (isPermissionError) {
      errorHeading = 'Access Denied';
      errorDescription = `You don't have permission to access this presentation (ID: ${presentationId}).`;
      errorSuggestion = 'The presentation owner needs to share it with you or make it accessible to anyone with the link.';
    } else if (err.message && err.message.includes("not found")) {
      errorHeading = 'Presentation Not Found';
      errorDescription = `The presentation with ID: ${presentationId} could not be found.`;
      errorSuggestion = 'Check that the presentation ID in the URL is correct and that the presentation hasn\'t been deleted.';
    } else if (err.message && err.message.includes("quota")) {
      errorHeading = 'Service Quota Exceeded';
      errorDescription = 'The application has reached its usage limits.';
      errorSuggestion = 'Please try again later or contact the presentation owner.';
    }
    
    // Return a user-friendly error page
    return HtmlService.createHtmlOutput(`
      <html>
      <head>
        <title>Interactive Training - ${errorTitle}</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
          h1, h2 { color: #d93025; }
          .container { max-width: 800px; margin: 0 auto; }
          .error-section { margin-bottom: 30px; padding: 20px; border: 1px solid #f44336; border-radius: 8px; background-color: #fff0f0; }
          .suggestion { margin-top: 20px; padding: 15px; background-color: #fffde7; border-radius: 8px; }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>${errorTitle}</h1>
          <div class="error-section">
            <h2>${errorHeading}</h2>
            <p>${errorDescription}</p>
            <div class="suggestion">
              <p><strong>Suggestion:</strong> ${errorSuggestion}</p>
            </div>
          </div>
        </div>
      </body>
      </html>
    `)
    .setTitle(`Interactive Training - ${errorTitle}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
      console.error(`Slide index ${slideIndex} is out of range (max: ${slides.length - 1})`);
      return { error: `Invalid slide index. The presentation only has ${slides.length} slides.` };
    }

    const slide = slides[slideIndex];
    const slideId = slide.getObjectId();
    console.log(`[GS] Processing slide ID: ${slideId}`);

    // Get global overlay settings
    const globalOverlayResult = getGlobalOverlaySettings();
    if (!globalOverlayResult.success) {
      console.warn(`Failed to get global overlay settings: ${globalOverlayResult.error}`);
    }
    const globalOverlayDefaults = globalOverlayResult.settings || OVERLAY_DEFAULTS;

    // Prepare the data object to return
    const slideData = {
      id: slideId,
      index: slideIndex,
      total: slides.length,
      elements: [],
      globalOverlayDefaults: globalOverlayDefaults,
      dimensions: { 
        width: presentation.getPageWidth(),
        height: presentation.getPageHeight()
      }
    };

    // Get saved navigation sequence if it exists
    const sequenceResult = getNavigationSequence(slideId);
    const savedSequence = sequenceResult.success && sequenceResult.hasCustomSequence ? 
                          sequenceResult.elementIds : null;
    
    slideData.hasCustomSequence = savedSequence !== null;
    slideData.customSequence = savedSequence || [];

    // Get background image data URL
    try {
      slideData.backgroundUrl = getSlideBackgroundAsDataUrl(slide);
      console.log(`[GS] Background URL fetched: ${slideData.backgroundUrl ? 'Success' : 'null'}`);
    } catch (bgError) {
      console.error(`Error getting background: ${bgError.message}`);
    }

    // Get slide notes
    try {
      slideData.notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
    } catch(notesError) {
      console.warn(`No speaker notes found: ${notesError}`);
      slideData.notes = "";
    }

    // Process page elements to find interactive ones
    const pageElements = slide.getPageElements();
    console.log(`[GS] Found ${pageElements.length} elements on slide ${slideIndex}`);

    // Create a mapping of IDs for easier lookup and debugging
    const elementIdMap = new Map();
    
    // First pass: collect all elements IDs for complete mapping
    pageElements.forEach((element) => {
      elementIdMap.set(element.getObjectId(), {
        type: element.getPageElementType().toString(),
        left: element.getLeft(),
        top: element.getTop(),
        width: element.getWidth(), 
        height: element.getHeight()
      });
    });

    // Array to store all elements with interactive data
    const elementsWithData = [];

    // Second pass: process elements with interactive data
    pageElements.forEach((element) => {
      try {
        const desc = element.getDescription();
        if (!desc || !desc.trim()) return; // Skip if no description
        
        if (desc.trim().startsWith('{') && desc.trim().endsWith('}')) {
          try {
            const data = JSON.parse(desc);
            const hasInteractionData = data.interaction && typeof data.interaction === 'object';
            const hasAnimationData = data.animation && typeof data.animation === 'object';
            
            if (hasInteractionData || hasAnimationData) {
              // Element has valid interactive data, add it to our list
              const elementId = element.getObjectId();
              const elementData = {
                id: elementId,
                type: element.getPageElementType().toString(),
                left: element.getLeft(),
                top: element.getTop(),
                width: element.getWidth(),
                height: element.getHeight(),
                interaction: data.interaction || null,
                animation: data.animation || null
              };
              
              // Try to extract element text for identification (but don't fail if we can't)
              try {
                if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                  elementData.text = element.asShape().getText().asString();
                } else if (element.getPageElementType() === SlidesApp.PageElementType.TEXT_BOX) {
                  elementData.text = element.asTextBox().getText().asString();
                }
              } catch (e) {
                console.warn(`Couldn't get text for element ${elementId}: ${e.message}`);
              }
              
              elementsWithData.push(elementData);
            }
          } catch (e) {
            console.warn(`Invalid JSON in element ${element.getObjectId()}: ${e.message}`);
          }
        }
      } catch (e) {
        console.warn(`Error processing element: ${e.message}`);
      }
    });

    // Apply ordering based on saved sequence (if exists)
    if (savedSequence && savedSequence.length > 0) {
      // First add elements in the saved sequence
      const orderedElements = [];
      const remainingElements = [...elementsWithData]; // Make a copy to track what's left
      
      savedSequence.forEach(id => {
        const elementIndex = remainingElements.findIndex(el => el.id === id);
        if (elementIndex !== -1) {
          orderedElements.push(remainingElements[elementIndex]);
          remainingElements.splice(elementIndex, 1); // Remove from remaining
        }
      });
      
      // Add any elements not in the sequence at the end
      slideData.elements = [...orderedElements, ...remainingElements];
    } else {
      slideData.elements = elementsWithData;
    }
    
    // Assign sequential display IDs for easier reference
    slideData.elements.forEach((element, index) => {
      element.displayId = index + 1;
    });

    // Add debug information about all IDs for troubleshooting
    slideData._debug = {
      allElementIds: Array.from(elementIdMap.keys()),
      interactiveElementIds: slideData.elements.map(el => el.id)
    };
    
    console.log(`[GS] Finished processing slide ${slideIndex}. Found ${slideData.elements.length} elements with data.`);
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
        // Check cache first
        const cache = CacheService.getScriptCache();
        const slideId = slide.getObjectId();
        const cacheKey = 'bg_' + slideId;
        
        // Try to get cached background
        const cachedData = cache.get(cacheKey);
        if (cachedData) {
            console.log(`Retrieved slide background from cache for slide ${slideId}`);
            return cachedData;
        }

        // Not in cache, get via normal methods
        const background = slide.getBackground();
        let dataUrl = null;

        // Method 1: Try to get image using getPictureFill()
        try {
            const fill = background.getPictureFill();
            if (fill) {
                const blob = fill.getBlob();
                if (blob) {
                    dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
                }
            }
        } catch (e) {
            // Silent fail and continue to next method
        }

        // Method 2: Check if it's a solid fill color (will return null)
        if (!dataUrl) {
            try {
                const solidFill = background.getSolidFill();
                if (solidFill) {
                    // No image background, don't need to continue
                    // Still cache the null result to avoid repeated checks
                    cache.put(cacheKey, "null", 300); // Cache for 5 minutes
                    return null;
                }
            } catch(e) {
                // Silent fail and continue to next method
            }
        }

        // Method 3: Use Slides API to get thumbnail as fallback (Requires Advanced Slides Service enabled)
        if (!dataUrl && typeof Slides !== 'undefined' && Slides.Presentations && Slides.Presentations.Pages) {
            try {
                // Get presentation and slide IDs
                const presentationId = slide.getParent().getId();

                // Request thumbnail from Slides API with a timeout wrapper
                let thumbnailResponse;
                try {
                    // Set a timeout to prevent long-running execution
                    const start = new Date().getTime();
                    
                    thumbnailResponse = Slides.Presentations.Pages.getThumbnail(
                        presentationId,
                        slideId,
                        { 'thumbnailProperties.thumbnailSize': 'LARGE' }
                    );
                    
                    const elapsed = new Date().getTime() - start;
                    if (elapsed > 1000) {
                        console.warn(`Slides API thumbnail fetch took ${elapsed}ms, consider optimizing`);
                    }
                } catch (timeoutError) {
                    console.warn(`Thumbnail fetch took too long or failed: ${timeoutError.message}`);
                }

                if (thumbnailResponse && thumbnailResponse.contentUrl) {
                    // Add a timeout to the URL fetch to prevent excessive delays
                    const urlFetchOptions = {
                        'muteHttpExceptions': true,
                        'followRedirects': true,
                        'validateHttpsCertificates': true,
                        'timeout': 5000 // 5 second timeout
                    };
                    
                    try {
                        const response = UrlFetchApp.fetch(thumbnailResponse.contentUrl, urlFetchOptions);
                        if (response.getResponseCode() === 200) {
                            const blob = response.getBlob();
                            if (blob) {
                                dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
                            }
                        }
                    } catch (fetchError) {
                        console.warn(`URL fetch for thumbnail failed: ${fetchError.message}`);
                    }
                }
            } catch (e) {
                console.error("Method 3 (Slides API) failed: " + e.message);
            }
        }

        // Cache the result for future use (5 minutes)
        if (dataUrl) {
            console.log(`Caching slide background for slide ${slideId}`);
            cache.put(cacheKey, dataUrl, 300); // 5 minutes TTL
        } else {
            // Cache null result to avoid repeated lookups
            cache.put(cacheKey, "null", 300); 
        }

        return dataUrl;
    } catch (e) {
        console.error(`General error getting background for slide ${slide.getObjectId()}: ${e}`);
        return null;
    }
}


/**
 * Updates just the nickname of an element, preserving other data
 * @param {string} elementId The ID of the element to update
 * @param {string} nickname The new nickname to assign to this element
 * @returns {object} Result object with success/error information
 */
function mergeElementNickname(elementId, nickname) {
  try {
    // Get the current element description to check for existing data
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();

    // Ensure there's a selection context to find the slide
    if (!selection || selection.getSelectionType() === SlidesApp.SelectionType.NONE) {
      return { success: false, error: "No slide or element selected. Please select the slide containing your element." };
    }

    // Get the current slide from the selection
    const slide = selection.getCurrentPage();
    if (!slide) {
      return { success: false, error: "Could not determine the active slide. Please select a slide." };
    }

    // Find the element on the current slide using its ID
    const element = slide.getPageElementById(elementId);
    if (!element) {
      return { success: false, error: "Element not found on the current slide. Please ensure the correct slide is active." };
    }

    // Get current description and parse as JSON if exists
    let currentData = {};
    const description = element.getDescription();

    if (description && description.trim() !== "") {
      try {
        currentData = JSON.parse(description);
      } catch (e) {
        console.warn(`Invalid JSON in element description, treating as empty: ${e.message}`);
        // Continue with empty object if parsing fails
      }
    }

    // Add or update the nickname property
    currentData.nickname = nickname;
    
    // Save merged data
    const mergedJson = JSON.stringify(currentData);
    element.setDescription(mergedJson);

    // Verify the save worked by reading it back
    const verifyDesc = element.getDescription();
    let verifiedParsedData = null;
    try {
      verifiedParsedData = JSON.parse(verifyDesc);
    } catch(e) {
       console.error("Verification failed: Could not parse saved description back.");
    }

    return {
      success: true,
      message: `Nickname saved successfully.`,
      elementId: elementId,
      verifiedFull: verifiedParsedData
    };
  } catch (e) {
    console.error(`Error in mergeElementNickname: ${e.message}`);
    return {
      success: false,
      error: `Failed to save nickname: ${e.message}`,
      elementId: elementId
    };
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
                html += '<table><thead><tr><th>ID</th><th>Type</th><Pos (L,T)</th><th>Size (W,H)</th><th>Interaction</th><th>Animation</th></tr></thead><tbody>';


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
    const selection = presentation.getSelection();
    let element = null;
    let slideId = null;
    let slideIndex = -1;
    
    // First, check the current slide from selection
    if (selection && selection.getCurrentPage()) {
      const currentSlide = selection.getCurrentPage();
      element = currentSlide.getPageElementById(elementId);
      
      if (element) {
        slideId = currentSlide.getObjectId();
        // Find the slide index
        const slides = presentation.getSlides();
        for (let i = 0; i < slides.length; i++) {
          if (slides[i].getObjectId() === slideId) {
            slideIndex = i;
            break;
          }
        }
        console.warn(`Element ${elementId} found on the current slide.`);
      } else {
        console.warn(`Element ${elementId} not found on current slide, searching all slides...`);
      }
    }
    
    // If not found on current slide, search all slides
    if (!element) {
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
