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
 * @returns {object} An object indicating success or failure, with message or error, and the verified full data.
 */
function mergeElementData(elementId, newData, dataType) {
  try {
    // Log the incoming data for debugging
    console.log(`mergeElementData called for ${elementId}, dataType: ${dataType}`);
    console.log(`newData:`, JSON.stringify(newData));

    // Get the current element description to check for existing data
    const presentation = SlidesApp.getActivePresentation();
    let element = null;
    let slideFoundOn = null;

    // Find the element across all slides robustly
    const slides = presentation.getSlides();
    for(let i = 0; i < slides.length; i++) {
        try {
            const el = slides[i].getPageElementById(elementId);
            if (el) {
                element = el;
                slideFoundOn = slides[i]; // Keep track of the slide
                break;
            }
        } catch (e) { /* Ignore if element not on this slide */ }
    }

    if (!element) {
        console.error(`Element ID ${elementId} not found in presentation.`);
        return { success: false, error: `Element not found in presentation.` };
    }

    // Get current description and parse as JSON if exists
    let currentData = {};
    const description = element.getDescription();

    if (description && description.trim() !== "") {
      try {
        currentData = JSON.parse(description);
        console.log(`Found existing data:`, JSON.stringify(currentData));
      } catch (e) {
        console.warn(`Invalid JSON in element description for ${elementId}, treating as empty: ${e.message}`);
      }
    }

    // --- MERGE LOGIC ---
    let mergedData = { ...currentData };

    if (newData && newData.type === 'none') {
      // If type is 'none', delete the entire data block (animation or interaction)
      delete mergedData[dataType];
      console.log(`Type is 'none', removing entire ${dataType} block`);
    } else if (newData) {
      // Initialize the data type block if it doesn't exist
      if (!mergedData[dataType]) mergedData[dataType] = {};
      
      // Special handling for interaction with overlay style
      if (dataType === 'interaction') {
        const useCustomOverlay = newData.useCustomOverlay;
        
        // First merge everything except overlayStyle
        const { overlayStyle, useCustomOverlay: _, ...dataWithoutStyle } = newData;
        mergedData.interaction = { ...mergedData.interaction, ...dataWithoutStyle };
        
        // Handle overlay style specifically
        if (useCustomOverlay === true && overlayStyle) {
          // User wants custom style and provided style data
          mergedData.interaction.overlayStyle = overlayStyle;
          console.log(`Applied custom overlay style:`, JSON.stringify(overlayStyle));
        } else if (useCustomOverlay === false) {
          // User explicitly doesn't want custom style
          delete mergedData.interaction.overlayStyle;
          console.log(`Removed custom overlay style`);
        } else if (overlayStyle === null) {
          // Only remove if explicitly set to null
          delete mergedData.interaction.overlayStyle;
          console.log(`Removed custom overlay style (null case)`);
        }
        // Otherwise, don't touch the overlayStyle (undefined case)
        
        // Clean up
        if (!mergedData.interaction.showOverlayText) {
          delete mergedData.interaction.overlayText;
        }
      } else {
        // For animation or other types, simple merge
        mergedData[dataType] = { ...mergedData[dataType], ...newData };
      }
    } else {
       console.warn(`mergeElementData called with null or undefined newData for ${dataType}`);
    }

    // Log the merged operation for debugging
    //logMergeOperation(elementId, newData, dataType, currentData, mergedData);

    // Check if the entire description should be cleared
    const interactionIsEmpty = !mergedData.interaction || Object.keys(mergedData.interaction).length === 0 || 
                              (mergedData.interaction.type === 'none' && !mergedData.interaction.overlayStyle);
    const animationIsEmpty = !mergedData.animation || Object.keys(mergedData.animation).length === 0 || 
                             mergedData.animation.type === 'none';
    const nicknameIsEmpty = !mergedData.nickname;

    if (interactionIsEmpty && animationIsEmpty && nicknameIsEmpty) {
      element.setDescription("");
      console.log(`Cleared description for ${elementId} as all data was empty.`);
      return {
        success: true,
        message: `${dataType.charAt(0).toUpperCase() + dataType.slice(1)} data resulted in empty description.`,
        elementId: elementId,
        updatedDescription: "" // Explicitly return empty string
      };
    }

    // --- SAVE MERGED DATA ---
    const mergedJson = JSON.stringify(mergedData);
    element.setDescription(mergedJson);
    console.log(`Saved merged description for ${elementId}: ${mergedJson.substring(0, 200)}...`);

    // Return the updated description string for client-side cache update
    return {
      success: true,
      message: `${dataType.charAt(0).toUpperCase() + dataType.slice(1)} data saved successfully.`,
      elementId: elementId,
      updatedDescription: mergedJson // Return the actual saved string
    };
  } catch (e) {
    console.error(`Error in mergeElementData for ${elementId}: ${e.message}\nStack: ${e.stack}`);
    return {
      success: false,
      error: `Failed to merge data: ${e.message}`,
      elementId: elementId
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

    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      const elementId = element.getObjectId();
      const desc = element.getDescription();
      let data = null;
      let isInteractive = false; // Flag to determine if it should be listed

      // Basic info always included
      let elementBaseData = {
          id: elementId,
          type: element.getPageElementType().toString(),
          left: element.getLeft(),
          top: element.getTop(),
          width: element.getWidth(),
          height: element.getHeight(),
          interaction: null, // Initialize as null
          animation: null,
          nickname: null
      };

      // Try to get name/text
       try {
           const elementType = element.getPageElementType();
           if (elementType === SlidesApp.PageElementType.SHAPE && element.asShape().getText) {
               elementBaseData.name = element.asShape().getText().asString().trim();
           } else if (elementType === SlidesApp.PageElementType.TEXT_BOX && element.asTextBox().getText) {
               elementBaseData.name = element.asTextBox().getText().asString().trim();
           } else if (elementType === SlidesApp.PageElementType.IMAGE && element.getTitle) {
               elementBaseData.name = element.getTitle() || "Image";
           } else if (elementType === SlidesApp.PageElementType.VIDEO && element.getTitle) {
               elementBaseData.name = element.getTitle() || "Video";
           } else if (elementType === SlidesApp.PageElementType.GROUP) {
               elementBaseData.name = "Group";
           }
       } catch (e) {
         console.warn(`Could not get text/name for element ${elementId}: ${e.message}`);
       }
       // Fallback name
       if (!elementBaseData.name) {
           elementBaseData.name = `${elementBaseData.type} (${elementId.substring(0, 5)}...)`;
       }
       // Shorten long names
       if (elementBaseData.name && elementBaseData.name.length > 50) { // Increased length slightly
           elementBaseData.name = elementBaseData.name.substring(0, 50).trim() + "...";
       }

      // Try to parse description for interaction/animation/nickname data
      if (desc && desc.trim().startsWith('{') && desc.trim().endsWith('}')) {
        try {
          data = JSON.parse(desc);
          // Merge parsed data into base data
          if (data.interaction) elementBaseData.interaction = data.interaction;
          if (data.animation) elementBaseData.animation = data.animation;
          if (data.nickname) elementBaseData.nickname = data.nickname;

          // Mark as interactive if it has any relevant data block OR a nickname
          isInteractive = !!(elementBaseData.interaction || elementBaseData.animation || elementBaseData.nickname);

        } catch (e) {
          console.warn(`Element ${elementId} has invalid JSON in description: ${e.message}`);
          isInteractive = false; // Don't list if JSON is invalid
        }
      }

      // Always add the element with its base info, plus any merged data
      // Client-side will decide if it's "interactive" based on the presence of interaction/animation/nickname
      interactiveElements.push(elementBaseData);

    } // End loop through elements

    // Assign sequential display IDs based on the order found on the slide initially
    interactiveElements.forEach((element, index) => {
      element.displayId = index + 1; // Assign sequential numbers based on slide order
    });

    console.log(`Returning ${interactiveElements.length} total elements (with/without data) for slide ${slideId}`);
    return {
      elements: interactiveElements, // Return ALL elements with merged data
      slideId: slideId,
      hasCustomSequence: savedSequence !== null,
      customSequence: savedSequence // Include the saved sequence if it exists
    };
  } catch (error) {
    console.error("Error in getAllInteractiveElements:", error);
    return {
      error: `Failed to get interactive elements: ${error.message}. Please check your slide content and try again.`
    };
  }
}




/**
 * Removes all settings (description) from a specific element.
 * @param {string} elementId The ID of the element to modify.
 * @returns {object} An object indicating success or failure.
 */
function removeElementInteraction(elementId) { // Renamed for consistency, but clears all
  try {
    const presentation = SlidesApp.getActivePresentation();
    let element = null;

    // Find the element across all slides robustly
    const slides = presentation.getSlides();
    for(let i = 0; i < slides.length; i++) {
        try {
            const el = slides[i].getPageElementById(elementId);
            if (el) { element = el; break; }
        } catch (e) { /* Ignore */ }
    }

    if (!element) {
      return { success: false, error: "Element not found in presentation." };
    }

    // Clear the description to remove ALL interaction/animation/nickname data.
    element.setDescription("");
    console.log(`Cleared description for element ${elementId}`);

    return {
      success: true,
      message: "All settings removed successfully."
    };
  } catch (e) {
    console.error("Error removing element settings:", e);
    return {
      success: false,
      error: `Failed to remove settings: ${e.message}`
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
    let targetElement = null;
    let targetSlide = null;

    // Find the element across all slides
    const slides = presentation.getSlides();
    for (let i = 0; i < slides.length; i++) {
      try {
        const element = slides[i].getPageElementById(elementId);
        if (element) {
          targetElement = element;
          targetSlide = slides[i];
          break;
        }
      } catch (e) { /* Skip */ }
    }

    if (!targetElement || !targetSlide) {
      return {
        success: false,
        error: "Element not found in presentation. It may have been deleted."
      };
    }

    // Switch to target slide if needed
    const selection = presentation.getSelection();
    const currentSlide = selection?.getCurrentPage();
    if (currentSlide?.getObjectId() !== targetSlide.getObjectId()) {
      targetSlide.selectAsCurrentPage();
      Utilities.sleep(100); // Short delay for UI update
    }

    // Select the element
    targetElement.select();
    Utilities.sleep(50); // Short delay after selection

    return { success: true, message: "Element selected successfully." };
  } catch (e) {
    console.error("Error in selectElementById:", e);
    return { success: false, error: e.message };
  }
}




/**
 * Clears the description field of the currently selected element or specified element ID.
 * Called via Sidebar action button or delete functions.
 * @param {string} [elementId] Optional: ID of element to clear. If null, uses current selection.
 * @returns {object} Result indicating success or failure, includes updatedDescription.
 */
function clearElementDescription(elementId = null) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    let element = null;
    let currentElementId = elementId; // Use provided ID if available

    if (currentElementId) {
       // If ID provided, find it directly across all slides
       const slides = presentation.getSlides();
       for(let i=0; i<slides.length; i++) {
           try {
               const el = slides[i].getPageElementById(currentElementId);
               if (el) { element = el; break; }
           } catch(e){}
       }
    } else {
       // If no ID, rely on current selection
       if (selection?.getSelectionType() === SlidesApp.SelectionType.PAGE_ELEMENT) {
           const pageElements = selection.getPageElementRange().getPageElements();
           if (pageElements.length === 1) {
               element = pageElements[0];
               currentElementId = element.getObjectId(); // Get ID from selection
           } else if (pageElements.length > 1) {
                return { success: false, error: "Multiple elements selected. Please select only one." };
           }
       }
    }

    if (!element) {
      return { success: false, error: currentElementId ? `Element ID ${currentElementId} not found.` : "No single element selected." };
    }

    // Clear the description
    element.setDescription("");
    console.log(`Cleared description for element ${currentElementId}`);

    return {
      success: true,
      message: "Element settings cleared successfully.",
      elementId: currentElementId,
      updatedDescription: "" // Explicitly return empty string representing cleared state
    };
  } catch (e) {
    console.error(`Error clearing description for ${elementId}: ${e.message}`);
    return {
      success: false,
      error: `Failed to clear settings: ${e.message}`
    };
  }
}



// --- Global Overlay Style Constants ---
const OVERLAY_SETTINGS_PREFIX = 'WEBAPP_OVERLAY_';
const OVERLAY_DEFAULTS = {
  shape: 'rectangle',
  color: '#e53935',
  opacity: 0,
  outlineEnabled: true,
  outlineColor: '#e53935',
  outlineWidth: 1,
  outlineStyle: 'dashed',
  textColor: '#ffffff',
  textSize: 14,
  hoverText: 'Click here',
  textBackground: false
};


/**
 * Gets the global overlay style settings saved in Script Properties.
 * Returns defaults if a setting is not found.
 * @returns {object} An object containing all global overlay settings.
 */
function getGlobalOverlaySettings() {
  const props = PropertiesService.getScriptProperties();
  const settings = {};

  Object.keys(OVERLAY_DEFAULTS).forEach(key => {
    const propKey = OVERLAY_SETTINGS_PREFIX + key.toUpperCase();
    const rawValue = props.getProperty(propKey);

    if (rawValue === null) {
      settings[key] = OVERLAY_DEFAULTS[key];
    } else {
      const defaultValue = OVERLAY_DEFAULTS[key];
      if (typeof defaultValue === 'number') {
        const numValue = parseFloat(rawValue);
        settings[key] = isNaN(numValue) ? defaultValue : numValue;
      } else if (typeof defaultValue === 'boolean') {
        settings[key] = (rawValue === 'true');
      } else {
        settings[key] = rawValue; // String
      }
    }
  });

  // console.log(`getGlobalOverlaySettings: Returning ${JSON.stringify(settings)}`); // Reduce logging noise
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
    const keys = Object.keys(OVERLAY_DEFAULTS);

    keys.forEach(key => {
      const propKey = OVERLAY_SETTINGS_PREFIX + key.toUpperCase();
      if (settings.hasOwnProperty(key) && settings[key] !== undefined && settings[key] !== null) {
        let valueToSave = settings[key];
        const defaultValue = OVERLAY_DEFAULTS[key];
        if (typeof defaultValue === 'number') {
          valueToSave = parseFloat(valueToSave);
          if (isNaN(valueToSave)) valueToSave = defaultValue;
        } else if (typeof defaultValue === 'boolean') {
          valueToSave = !!valueToSave;
        } else if (typeof defaultValue === 'string') {
           valueToSave = String(valueToSave).trim();
        }
        props.setProperty(propKey, String(valueToSave));
      } else {
        // If key is missing or invalid in input, delete the property to revert to default
        props.deleteProperty(propKey);
        console.log(`setGlobalOverlaySettings: Key '${key}' missing or invalid, deleting property ${propKey}.`);
      }
    });

    console.log(`setGlobalOverlaySettings: Saved settings.`);
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

  // Check for diagnostic mode
  if (e?.parameter?.diag === 'true') {
      console.log("Serving diagnostic page.");
      return serveDiagnosticPage(e);
  }

  // If no presentation ID is provided, show a helpful intro page
  if (!presentationId) {
    return HtmlService.createHtmlOutputFromFile('Welcome') // Use separate Welcome.html file
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
    console.error(`Error accessing presentation ${presentationId}: ${err.message}`);

    // Provide more specific error messages
    let errorTitle = 'Error Loading Presentation';
    let errorHeading = 'Error';
    let errorDescription = `Could not load the presentation.`;
    let errorSuggestion = 'Please check the presentation ID in the URL. The presentation may not exist, have been deleted, or there was a server problem.';

    const isPermissionError = err.message && (
      err.message.includes("Access denied") ||
      err.message.includes("permission") ||
      err.message.includes("Insufficient") ||
      err.message.includes("not found")
    );

    if (isPermissionError) {
      errorHeading = 'Access Denied or Not Found';
      errorDescription = `Could not access the presentation (ID: ${presentationId.substring(0,15)}...).`;
      errorSuggestion = 'Please verify the link and ensure the presentation owner has shared it with you or made it accessible.';
    }

    // Return a user-friendly error page
    const errorTemplate = HtmlService.createTemplateFromFile('Error');
    errorTemplate.errorTitle = errorTitle;
    errorTemplate.errorHeading = errorHeading;
    errorTemplate.errorDescription = errorDescription;
    errorTemplate.errorSuggestion = errorSuggestion;
    return errorTemplate.evaluate()
      .setTitle(`Interactive Training - ${errorTitle}`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}


/**
 * Fetches data for a specific slide to be displayed in the web app.
 * Includes global overlay style settings and data for ALL elements on the slide.
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
  slideIndex = parseInt(slideIndex, 10);
  if (isNaN(slideIndex) || slideIndex < 0) {
    console.log("getSlideDataForWebApp: Invalid or missing slide index, defaulting to 0");
    slideIndex = 0;
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

    // Prepare the data object
    const slideData = {
      id: slideId,
      index: slideIndex,
      total: slides.length,
      elements: [], // **MODIFIED: Will contain ALL elements**
      globalOverlayDefaults: globalOverlayDefaults,
      slideWidth: presentation.getPageWidth(),
      slideHeight: presentation.getPageHeight(),
      notes: "" // Initialize notes
    };

    // Get slide notes
    try {
      const notesPage = slide.getNotesPage();
      if (notesPage) {
        const notesShape = notesPage.getSpeakerNotesShape();
        if (notesShape) slideData.notes = notesShape.getText().asString();
      }
    } catch(notesError) {
      console.warn(`No speaker notes found or error getting notes: ${notesError}`);
    }

    // Get background image
    slideData.backgroundUrl = getSlideBackgroundAsDataUrl_NoCache(slide);
    console.log(`[GS] Background URL fetched: ${slideData.backgroundUrl ? 'Success' : 'null'}`);

    // **MODIFIED: Process ALL page elements**
    const pageElements = slide.getPageElements();
    console.log(`[GS] Found ${pageElements.length} total elements on slide ${slideIndex}`);

    pageElements.forEach((element) => {
      try {
        const elementId = element.getObjectId();
        // Base data for EVERY element
        let elementData = {
          id: elementId,
          type: element.getPageElementType().toString(),
          left: element.getLeft(),
          top: element.getTop(),
          width: element.getWidth(),
          height: element.getHeight(),
          interaction: null, // Initialize
          animation: null,
          nickname: null
        };

        // Try to get description and merge data
        const desc = element.getDescription();
        if (desc && desc.trim().startsWith('{') && desc.trim().endsWith('}')) {
          try {
            const parsedDesc = JSON.parse(desc);
            // Merge known properties
            if (parsedDesc.interaction) elementData.interaction = parsedDesc.interaction;
            if (parsedDesc.animation) elementData.animation = parsedDesc.animation;
            if (parsedDesc.nickname) elementData.nickname = parsedDesc.nickname;
          } catch (e) {
            console.warn(`Element ${elementId} has invalid JSON in description: ${e.message}`);
          }
        }
        // Add the element data (basic or merged) to the list
        slideData.elements.push(elementData);

      } catch (e) {
        // Catch errors processing individual elements (e.g., unsupported types)
        console.warn(`Error processing element ${element?.getObjectId() || 'UNKNOWN'}: ${e.message}`);
      }
    }); // End loop through pageElements

    console.log(`[GS] Finished processing slide ${slideIndex}. Returning data for ${slideData.elements.length} elements.`);
    return slideData; // Return data for ALL elements

  } catch (e) {
    console.error(`Error in getSlideDataForWebApp (ID: ${presentationId}, Index: ${slideIndex}): ${e.toString()}\nStack: ${e.stack}`);
    return { error: `An error occurred while loading slide data: ${e.message}. Check script logs.` };
  }
}

/**
 * **FIXED VERSION WITHOUT CACHING**
 * Attempts to get the background of a slide as a Base64 Data URL.
 * Uses multiple methods for maximum compatibility.
 * Removed caching to prevent "Argument too large" errors.
 * @param {SlidesApp.Slide} slide The Google Slides slide object.
 * @returns {string|null} Base64 Data URL of the background or null if not found/error.
 */
function getSlideBackgroundAsDataUrl_NoCache(slide) {
  const slideId = slide.getObjectId();
  console.log(`[GS - NoCache] Getting background for slide ${slideId}`);
  let dataUrl = null;

  try {
    // Check for solid fill first
    try {
      const background = slide.getBackground();
      if (background.getSolidFill()) {
        console.log(`[GS - NoCache] Slide ${slideId} has solid fill background.`);
        return null;
      }
    } catch (e) {
      console.warn(`[GS - NoCache] Error checking solid fill for ${slideId}: ${e.message}`);
    }

    // Try picture fill next
    try {
      const background = slide.getBackground();
      const fill = background.getPictureFill();
      if (fill && fill.getBlob) {
        const blob = fill.getBlob();
        if (blob) {
          console.log(`[GS - NoCache] Found picture fill for ${slideId}. Encoding...`);
          dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
          console.log(`[GS - NoCache] Picture fill encoding complete for ${slideId}. Length: ${dataUrl ? dataUrl.length : 'N/A'}`);
        }
      }
    } catch (e) {
       console.warn(`[GS - NoCache] Error checking picture fill for ${slideId}: ${e.message}`);
    }

    // Last resort: Slides API (If Slides Service is enabled)
    if (!dataUrl && typeof Slides !== 'undefined' && Slides.Presentations && Slides.Presentations.Pages) {
      try {
        console.log(`[GS - NoCache] Trying Slides API thumbnail method for ${slideId}...`);
        const presentationId = slide.getParent().getId();
        const thumbnailResponse = Slides.Presentations.Pages.getThumbnail(
          presentationId,
          slideId,
          { 'thumbnailProperties.thumbnailSize': 'LARGE' } // Request large for better quality
        );

        if (thumbnailResponse && thumbnailResponse.contentUrl) {
          console.log(`[GS - NoCache] Fetching thumbnail URL: ${thumbnailResponse.contentUrl}`);
          const response = UrlFetchApp.fetch(thumbnailResponse.contentUrl, {
             muteHttpExceptions: true,
             followRedirects: true,
             validateHttpsCertificates: true,
             headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
             deadline: 10 // 10 second deadline
          });

          const responseCode = response.getResponseCode();
          console.log(`[GS - NoCache] Thumbnail fetch response code: ${responseCode}`);

          if (responseCode === 200) {
            const blob = response.getBlob();
            console.log(`[GS - NoCache] Slides API blob received for ${slideId}. Encoding...`);
            dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
            console.log(`[GS - NoCache] Slides API encoding complete for ${slideId}. Length: ${dataUrl ? dataUrl.length : 'N/A'}`);
          } else {
             console.warn(`[GS - NoCache] Slides API thumbnail fetch failed for ${slideId} with status ${responseCode}. Response: ${response.getContentText().substring(0, 200)}`);
          }
        } else {
            console.log(`[GS - NoCache] Slides API did not return a contentUrl for ${slideId}.`);
        }
      } catch (e) {
        console.error(`[GS - NoCache] Slides API thumbnail method failed for ${slideId}: ${e.message}`);
        if (e.message && e.message.toLowerCase().includes('api service')) {
           console.error("[GS - NoCache] Suggestion: Ensure the 'Google Slides API' advanced service is enabled in the Apps Script project settings.");
        }
      }
    } else if (!dataUrl) {
        console.log("[GS - NoCache] Slides API service not available or picture fill did not yield data.");
    }

    // **Crucial:** Check size before returning. If > ~1MB, return null to avoid issues.
    // 1MB Base64 is roughly 1.37 MB original. Adjust limit as needed. Cache limit is ~100KB.
    // Web app transfer limit is much higher (~50MB), but rendering large data URLs can be slow.
    // Let's set a practical limit of 2MB for the data URL string itself.
    const MAX_DATA_URL_LENGTH = 2 * 1024 * 1024; // 2 MB
    if (dataUrl && dataUrl.length > MAX_DATA_URL_LENGTH) {
        console.warn(`[GS - NoCache] Generated background data URL for ${slideId} is too large (${(dataUrl.length / 1024 / 1024).toFixed(2)} MB). Returning null.`);
        return null;
    }

    console.log(`[GS - NoCache] Returning dataUrl (length: ${dataUrl ? dataUrl.length : 'null'}) for ${slideId}`);
    return dataUrl;

  } catch (e) {
    console.error(`[GS - NoCache] Unexpected error in getSlideBackgroundAsDataUrl_NoCache for ${slideId}: ${e.message}\nStack: ${e.stack}`);
    return null;
  }
}



/**
 * Updates just the nickname of an element, preserving other data
 * @param {string} elementId The ID of the element to update
 * @param {string} nickname The new nickname to assign to this element
 * @returns {object} Result object with success/error information and updated description
 */
function mergeElementNickname(elementId, nickname) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    let element = null;

    // Find the element across all slides robustly
    const slides = presentation.getSlides();
    for(let i = 0; i < slides.length; i++) {
        try {
            const el = slides[i].getPageElementById(elementId);
            if (el) { element = el; break; }
        } catch (e) { /* Ignore */ }
    }

    if (!element) {
        return { success: false, error: "Element not found in presentation." };
    }

    // Get current description and parse as JSON if exists
    let currentData = {};
    const description = element.getDescription();
    if (description && description.trim().startsWith('{') && description.trim().endsWith('}')) {
      try {
        currentData = JSON.parse(description);
      } catch (e) {
        console.warn(`Invalid JSON in element description for ${elementId}, treating as empty: ${e.message}`);
      }
    } else if (description && description.trim()) {
        console.warn(`Overwriting plain text description for ${elementId} with new nickname data.`);
        currentData = {}; // Start fresh
    }

    // Add or update/remove the nickname property
    const trimmedNickname = nickname ? nickname.trim() : null; // Trim or nullify
    if (!trimmedNickname) {
        delete currentData.nickname;
    } else {
        currentData.nickname = trimmedNickname;
    }

    // Save merged data
    const mergedJson = JSON.stringify(currentData);

    if (mergedJson === "{}") {
        element.setDescription(""); // Clear description if only nickname existed and was removed
        console.log(`Cleared description for ${elementId} as only nickname existed and was removed.`);
        return {
            success: true,
            message: `Nickname removed successfully.`,
            elementId: elementId,
            updatedDescription: "" // Explicitly return empty string
        };
    } else {
        element.setDescription(mergedJson);
        console.log(`Saved nickname description for ${elementId}.`);
        return {
            success: true,
            message: `Nickname saved successfully.`,
            elementId: elementId,
            updatedDescription: mergedJson // Return the actual saved string
        };
    }

  } catch (e) {
    console.error(`Error in mergeElementNickname for ${elementId}: ${e.message}\nStack: ${e.stack}`);
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
        pre { background-color: #f5f5f5; padding: 10px; border-radius: 4px; overflow-x: auto; white-space: pre-wrap; word-wrap: break-word; font-size: 11px; }
        button { background-color: #1a73e8; color: white; border: none; padding: 8px 16px;
                border-radius: 4px; cursor: pointer; margin-right: 10px; margin-top: 5px; }
        button:hover { background-color: #1558b7; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; font-size: 12px; }
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
              loadSlideBtn.textContent = 'Reload Slide Details'; // Reset button text
              loadSlideBtn.disabled = false; // Re-enable button

              if (result.error) {
                slideDetailsDiv.innerHTML = '<p style="color: red;">Error loading slide: ' + result.error + '</p>';
                return;
              }

              let html = '<div class="info-item"><span class="label">Slide ID:</span> ' + result.id + '</div>'; // Use result.id
              html += '<div class="info-item"><span class="label">Slide Index:</span> ' + result.index + '</div>'; // Use result.index
              html += '<div class="info-item"><span class="label">Notes:</span> <pre>' + (result.notes || 'None') + '</pre></div>';
              // Display global overlay defaults
              if (result.globalOverlayDefaults) {
                html += '<h3>Global Overlay Defaults</h3>';
                html += '<pre>' + JSON.stringify(result.globalOverlayDefaults, null, 2) + '</pre>';
              } else {
                html += '<p style="color: orange;">Global overlay defaults missing.</p>';
              }
              html += '<div class="info-item"><span class="label">Background:</span><br>';
              if (result.backgroundUrl) {
                // Display first few chars of data url for confirmation, not the whole thing
                html += '<span>Data URL fetched (Length: ' + result.backgroundUrl.length + ')</span>';
                // html += '<img src="' + result.backgroundUrl + '" class="slide-thumbnail" alt="Slide Background Preview"></div>'; // Avoid rendering large images
              } else {
                html += '<span>Solid color or no background image</span></div>';
              }

              html += '<h3>Elements (' + (result.elements?.length || 0) + ')</h3>';

              if (result.elements && result.elements.length > 0) {
                 html += '<table><thead><tr><th>ID</th><th>Type</th><th>Pos (L,T)</th><th>Size (W,H)</th><th>Interaction</th><th>Animation</th><th>Nickname</th></tr></thead><tbody>';

                result.elements.forEach(element => {
                  html += '<tr>';
                  html += '<td>' + element.id.substring(0,8) + '...</td>';
                  html += '<td>' + element.type + '</td>';
                  html += '<td>' + Math.round(element.left) + ', ' + Math.round(element.top) + '</td>';
                  html += '<td>' + Math.round(element.width) + ', ' + Math.round(element.height) + '</td>';

                  const interactionDetails = element.interaction ? JSON.stringify(element.interaction, null, 2) : '<i>null</i>';
                  html += '<td><pre>' + interactionDetails + '</pre></td>';

                  const animationDetails = element.animation ? JSON.stringify(element.animation, null, 2) : '<i>null</i>';
                  html += '<td><pre>' + animationDetails + '</pre></td>';

                  const nicknameDetails = element.nickname || '<i>null</i>';
                  html += '<td><pre>' + nicknameDetails + '</pre></td>';

                  html += '</tr>';
                });

                html += '</tbody></table>';
              } else {
                html += '<p>No elements found on this slide.</p>';
              }

              slideDetailsDiv.innerHTML = html;
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
function getElementDescriptionForDebug(elementId) { // Renamed to match sidebar call
  try {
    if (!elementId) return { success: false, error: "No element ID provided." };
    const presentation = SlidesApp.getActivePresentation();
    let element = null;

     // Find the element across all slides robustly
    const slides = presentation.getSlides();
    for(let i = 0; i < slides.length; i++) {
        try {
            const el = slides[i].getPageElementById(elementId);
            if (el) { element = el; break; }
        } catch (e) { /* Ignore */ }
    }

    if (!element) {
        return { success: false, error: `Element ID ${elementId} not found in presentation.` };
    }

    const description = element.getDescription();

    return {
        success: true,
        id: elementId,
        description: description || "(No Description)",
    };
  } catch (e) {
      console.error(`Error in getElementDescriptionForDebug for ${elementId}: ${e}`);
      return { success: false, error: `Debug failed: ${e.message}` };
  }
}