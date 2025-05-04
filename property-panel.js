/**
 * property-panel.js - Element property editing interface for the Interactive Training Projects Web App
 * Provides a dynamic panel for editing element properties
 */

// Property panel state
let propertyPanelState = {
    currentElement: null,
    multipleSelection: false,
    groupOpen: {
      // Default state for groups (can be overridden by user interaction)
      Geometry: true,
      Appearance: true,
      Text: true,
      Interaction: true,
      Timeline: true,
      MultipleSelection: true
    }
  };
  
  /**
   * Initialize the property panel
   */
  function initPropertyPanel() {
    console.log('Initializing property panel');
    
    // Hide the properties container initially
    hidePropertyPanel();
  }
  
  /**
   * Hide the property panel or show empty state
   */
  function hidePropertyPanel() {
    const propertiesContainer = document.getElementById('properties-container');
    const noSelectionMessage = document.getElementById('no-selection-message');
    
    if (propertiesContainer) propertiesContainer.style.display = 'none';
    if (noSelectionMessage) noSelectionMessage.style.display = 'block';
    
    // Reset the current element
    propertyPanelState.currentElement = null;
    propertyPanelState.multipleSelection = false;
  }
  
  /**
   * Update the property panel for a selected object
   * * @param {Object} fabricObject - Fabric.js object
   */
  function updatePropertyPanel(fabricObject) {
    console.log('Updating property panel for object:', fabricObject);
    
    const propertiesContainer = document.getElementById('properties-container');
    const noSelectionMessage = document.getElementById('no-selection-message');
    
    if (!propertiesContainer || !noSelectionMessage) {
      console.error("Property panel elements not found in the DOM.");
      return;
    }
    
    // Show the properties container
    propertiesContainer.style.display = 'block';
    noSelectionMessage.style.display = 'none';
    
    // Store the current element
    propertyPanelState.currentElement = fabricObject;
    propertyPanelState.multipleSelection = false;
    
    // Get element data from project
    let elementData = null;
    if (fabricObject.elementId && editorState.project && editorState.project.elements) {
      elementData = editorState.project.elements.find(e => e.elementId === fabricObject.elementId);
    }
    
    if (!elementData) {
      console.warn(`Element data not found for elementId: ${fabricObject.elementId}`);
      // Optionally create a default element data structure or handle gracefully
    }
    
    // Clear the container
    propertiesContainer.innerHTML = '';
    
    // Add property groups based on object type
    // Geometry properties
    addGeometryProperties(propertiesContainer, fabricObject, elementData);
    
    // Appearance properties
    addAppearanceProperties(propertiesContainer, fabricObject, elementData);
    
    // Text properties (if applicable)
    if (fabricObject.type === 'text' || fabricObject.type === 'textbox') {
      addTextProperties(propertiesContainer, fabricObject, elementData);
    }
    
    // Interaction properties
    addInteractionProperties(propertiesContainer, fabricObject, elementData);
    
    // Timeline properties (if applicable)
    if (editorState.currentSlide && 
        (editorState.currentSlide.fileType === 'YouTube Video' || 
         editorState.currentSlide.fileType === 'Audio')) {
      addTimelineProperties(propertiesContainer, fabricObject, elementData);
    }
  }
  
  /**
   * Update the property panel for multiple selected objects
   * * @param {Array<Object>} objects - Array of Fabric.js objects
   */
  function updatePropertyPanelForMultipleSelection(objects) {
    console.log('Updating property panel for multiple selection:', objects);
    
    const propertiesContainer = document.getElementById('properties-container');
    const noSelectionMessage = document.getElementById('no-selection-message');
    
    if (!propertiesContainer || !noSelectionMessage) {
      console.error("Property panel elements not found in the DOM.");
      return;
    }
    
    // Show the properties container
    propertiesContainer.style.display = 'block';
    noSelectionMessage.style.display = 'none';
    
    // Store the multiple selection state
    propertyPanelState.currentElement = null; // No single current element
    propertyPanelState.multipleSelection = true;
    
    // Clear the container
    propertiesContainer.innerHTML = '';
    
    // Add header for multiple selection
    const header = document.createElement('div');
    header.className = 'property-group-header'; // Use existing style
    header.innerHTML = `Multiple Selection (${objects.length} objects)`;
    propertiesContainer.appendChild(header);
    
    // Add common properties that can be edited for multiple objects
    addMultipleSelectionProperties(propertiesContainer, objects);
  }
  
  // --- Property Group Creation Functions ---
  
  /**
   * Add geometry properties to the property panel
   * * @param {HTMLElement} container - Property panel container
   * @param {Object} fabricObject - Fabric.js object
   * @param {Object} elementData - Element data from project
   */
  function addGeometryProperties(container, fabricObject, elementData) {
    // Create property group
    const propertyGroup = createPropertyGroup('Geometry', 'geometry-group');
    
    // Add position properties
    addNumberProperty(propertyGroup, 'X Position', fabricObject.left, value => {
      fabricObject.set('left', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'left', value);
    });
    
    addNumberProperty(propertyGroup, 'Y Position', fabricObject.top, value => {
      fabricObject.set('top', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'top', value);
    });
    
    // Add size properties
    addNumberProperty(propertyGroup, 'Width', fabricObject.getScaledWidth(), value => {
      fabricObject.set({ scaleX: value / fabricObject.width });
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'width', value);
    }, 1); // Min width 1
    
    addNumberProperty(propertyGroup, 'Height', fabricObject.getScaledHeight(), value => {
      fabricObject.set({ scaleY: value / fabricObject.height });
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'height', value);
    }, 1); // Min height 1
    
    // Add rotation property
    addNumberProperty(propertyGroup, 'Rotation', fabricObject.angle, value => {
      fabricObject.set('angle', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'angle', value);
    }, 0, 360);
    
    // Add the property group to the container
    container.appendChild(propertyGroup);
  }
  
  /**
   * Add appearance properties to the property panel
   * * @param {HTMLElement} container - Property panel container
   * @param {Object} fabricObject - Fabric.js object
   * @param {Object} elementData - Element data from project
   */
  function addAppearanceProperties(container, fabricObject, elementData) {
    // Create property group
    const propertyGroup = createPropertyGroup('Appearance', 'appearance-group');
    
    // Add Nickname property
    addTextProperty(propertyGroup, 'Nickname', elementData?.nickname || '', value => {
      updateElementProperty(fabricObject.elementId, 'nickname', value);
      // Potentially update timeline label if timeline is visible
      if (hasTimelineContent()) {
        updateTimeline();
      }
    });
    
    // Add color property (fill for shapes, background for text)
    if (fabricObject.type !== 'image') {
      const colorLabel = (fabricObject.type === 'text' || fabricObject.type === 'textbox') ? 'Background Color' : 'Fill Color';
      const initialColor = (fabricObject.type === 'text' || fabricObject.type === 'textbox') ? fabricObject.backgroundColor : fabricObject.fill;
      addColorProperty(propertyGroup, colorLabel, initialColor, value => {
        if (fabricObject.type === 'text' || fabricObject.type === 'textbox') {
          fabricObject.set('backgroundColor', value);
        } else {
          fabricObject.set('fill', value);
        }
        editorState.canvas.renderAll();
        updateElementProperty(fabricObject.elementId, 'color', value);
      });
    }
    
    // Add opacity property
    const opacity = Math.round((fabricObject.opacity || 1) * 100);
    addRangeProperty(propertyGroup, 'Opacity', opacity, value => {
      const newOpacity = value / 100;
      fabricObject.set('opacity', newOpacity);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'opacity', value); // Store as 0-100
    }, 0, 100, 1);
    
    // Add outline properties
    const hasOutline = fabricObject.strokeWidth > 0;
    const outlineCheckbox = addCheckboxProperty(propertyGroup, 'Outline', hasOutline, value => {
      const outlineProps = document.getElementById('outline-properties');
      if (value) {
        const currentStroke = elementData?.outlineColor || ClientConfig.DEFAULT_COLORS.OUTLINE;
        const currentWidth = elementData?.outlineWidth || 1;
        fabricObject.set({
          stroke: currentStroke,
          strokeWidth: currentWidth
        });
        if (outlineProps) outlineProps.style.display = 'block';
        updateElementProperty(fabricObject.elementId, 'outline', true);
        // Update sub-properties if they don't exist yet
        if (elementData && elementData.outlineColor === undefined) {
          updateElementProperty(fabricObject.elementId, 'outlineColor', currentStroke);
        }
        if (elementData && elementData.outlineWidth === undefined) {
          updateElementProperty(fabricObject.elementId, 'outlineWidth', currentWidth);
        }
      } else {
        fabricObject.set({
          stroke: undefined,
          strokeWidth: 0
        });
        if (outlineProps) outlineProps.style.display = 'none';
        updateElementProperty(fabricObject.elementId, 'outline', false);
      }
      editorState.canvas.renderAll();
    });
    
    // Add outline properties container (visible only when outline is enabled)
    const outlineProperties = document.createElement('div');
    outlineProperties.id = 'outline-properties';
    outlineProperties.style.display = hasOutline ? 'block' : 'none';
    outlineProperties.className = 'property-subgroup'; // Add class for styling/indentation
    
    addColorProperty(outlineProperties, 'Outline Color', fabricObject.stroke || ClientConfig.DEFAULT_COLORS.OUTLINE, value => {
      fabricObject.set('stroke', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'outlineColor', value);
    });
    
    addNumberProperty(outlineProperties, 'Outline Width', fabricObject.strokeWidth || 1, value => {
      fabricObject.set('strokeWidth', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'outlineWidth', value);
    }, 1, 20); // Min 1, Max 20
    
    // Insert outline sub-properties after the checkbox row
    const outlineRow = outlineCheckbox.closest('.property-row');
    if (outlineRow && outlineRow.parentNode) {
      outlineRow.parentNode.insertBefore(outlineProperties, outlineRow.nextSibling);
    } else {
      propertyGroup.appendChild(outlineProperties); // Fallback
    }
    
    // Add shadow property
    const hasShadow = !!fabricObject.shadow; // Check for shadow object presence
    addCheckboxProperty(propertyGroup, 'Shadow', hasShadow, value => {
      if (value) {
        fabricObject.setShadow({
          color: 'rgba(0,0,0,0.3)',
          offsetX: 3,
          offsetY: 3,
          blur: 5
        });
        updateElementProperty(fabricObject.elementId, 'shadow', true);
      } else {
        fabricObject.setShadow(null);
        updateElementProperty(fabricObject.elementId, 'shadow', false);
      }
      editorState.canvas.renderAll();
    });
    
    // Add visibility property
    addCheckboxProperty(propertyGroup, 'Initially Hidden', elementData?.initiallyHidden || false, value => {
      // This only updates the data model; visibility on load is handled by loadElement
      updateElementProperty(fabricObject.elementId, 'initiallyHidden', value);
      // Optionally provide visual feedback in the editor (e.g., dim the object)
      // fabricObject.set('opacity', value ? 0.5 : (elementData?.opacity || 100) / 100);
      // editorState.canvas.renderAll();
    });
    
    // Add the property group to the container
    container.appendChild(propertyGroup);
  }
  
  /**
   * Add text properties to the property panel
   * * @param {HTMLElement} container - Property panel container
   * @param {Object} fabricObject - Fabric.js text/textbox object
   * @param {Object} elementData - Element data from project
   */
  function addTextProperties(container, fabricObject, elementData) {
    // Create property group
    const propertyGroup = createPropertyGroup('Text', 'text-group');
    
    // Add text content property (using textarea for multi-line)
    addTextareaProperty(propertyGroup, 'Content', fabricObject.text, value => {
      fabricObject.set('text', value);
      // Adjust height if it's a textbox to fit content
      if (fabricObject.type === 'textbox') {
        fabricObject.initDimensions(); // Recalculate width/height based on text
      }
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'text', value);
      // Also update width/height in data model if textbox adjusted
      if (fabricObject.type === 'textbox') {
          updateElementProperty(fabricObject.elementId, 'width', fabricObject.getScaledWidth());
          updateElementProperty(fabricObject.elementId, 'height', fabricObject.getScaledHeight());
      }
    });
    
    // Add font family property
    addSelectProperty(propertyGroup, 'Font', ClientConfig.FONTS, fabricObject.fontFamily, value => {
      fabricObject.set('fontFamily', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'font', value);
    });
    
    // Add font size property
    addNumberProperty(propertyGroup, 'Font Size', fabricObject.fontSize, value => {
      fabricObject.set('fontSize', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'fontSize', value);
    }, 6, 120); // Min 6, Max 120
    
    // Add font color property
    addColorProperty(propertyGroup, 'Font Color', fabricObject.fill, value => {
      fabricObject.set('fill', value);
      editorState.canvas.renderAll();
      updateElementProperty(fabricObject.elementId, 'fontColor', value);
    });
    
    // Add text alignment property (for textbox)
    if (fabricObject.type === 'textbox') {
      addSelectProperty(propertyGroup, 'Align', ['left', 'center', 'right', 'justify'], fabricObject.textAlign, value => {
        fabricObject.set('textAlign', value);
        editorState.canvas.renderAll();
        updateElementProperty(fabricObject.elementId, 'textAlign', value);
      });
    }
    
    // Add the property group to the container
    container.appendChild(propertyGroup);
  }
  
  /**
   * Add interaction properties to the property panel
   * * @param {HTMLElement} container - Property panel container
   * @param {Object} fabricObject - Fabric.js object
   * @param {Object} elementData - Element data from project
   */
  function addInteractionProperties(container, fabricObject, elementData) {
    // Create property group
    const propertyGroup = createPropertyGroup('Interaction', 'interaction-group');
    
    // Add Trigger Type property
    addSelectProperty(propertyGroup, 'Trigger', Object.values(ClientConfig.TRIGGER_TYPES), elementData?.triggers || ClientConfig.TRIGGER_TYPES.CLICK, value => {
      updateElementProperty(fabricObject.elementId, 'triggers', value);
      // Update fabric object property if needed for runtime behavior
      fabricObject.triggers = value;
    });
    
    // Add Interaction Type property
    addSelectProperty(propertyGroup, 'Action', Object.values(ClientConfig.INTERACTION_TYPES), elementData?.interactionType || ClientConfig.INTERACTION_TYPES.REVEAL, value => {
      updateElementProperty(fabricObject.elementId, 'interactionType', value);
      // Update fabric object property if needed for runtime behavior
      fabricObject.interactionType = value;
      // Show/hide specific fields based on action
      document.getElementById('text-modal-message-group').style.display = (value === ClientConfig.INTERACTION_TYPES.REVEAL) ? 'flex' : 'none';
    });
    
    // Add Text Modal Message property (only shown for Reveal interaction)
    const textModalGroup = addTextareaProperty(propertyGroup, 'Reveal Message', elementData?.textModalMessage || '', value => {
      updateElementProperty(fabricObject.elementId, 'textModalMessage', value);
    });
    textModalGroup.id = 'text-modal-message-group';
    textModalGroup.style.display = (elementData?.interactionType === ClientConfig.INTERACTION_TYPES.REVEAL || !elementData?.interactionType) ? 'flex' : 'none';
    
    // Add the property group to the container
    container.appendChild(propertyGroup);
  }
  
  /**
   * Add timeline properties to the property panel
   * * @param {HTMLElement} container - Property panel container
   * @param {Object} fabricObject - Fabric.js object
   * @param {Object} elementData - Element data from project
   */
  function addTimelineProperties(container, fabricObject, elementData) {
    // Create property group
    const propertyGroup = createPropertyGroup('Timeline', 'timeline-group');
    
    // Get timeline data or initialize if it doesn't exist
    let timelineData = elementData?.timeline || {};
    
    // Add Start Time property
    addNumberProperty(propertyGroup, 'Start Time (s)', timelineData.startTime || 0, value => {
      timelineData.startTime = value;
      updateElementProperty(fabricObject.elementId, 'timeline', timelineData);
      updateTimeline(); // Update visual timeline
    }, 0, 1000, 0.1); // Min 0, Max 1000, Step 0.1
    
    // Add End Time property
    addNumberProperty(propertyGroup, 'End Time (s)', timelineData.endTime || 0, value => {
      timelineData.endTime = value;
      updateElementProperty(fabricObject.elementId, 'timeline', timelineData);
      updateTimeline(); // Update visual timeline
    }, 0, 1000, 0.1); // Min 0, Max 1000, Step 0.1
    
    // Add Animation In property
    addSelectProperty(propertyGroup, 'Animation In', Object.values(ClientConfig.ANIMATION_IN_TYPES), timelineData.animationIn || ClientConfig.ANIMATION_IN_TYPES.NONE, value => {
      timelineData.animationIn = value;
      updateElementProperty(fabricObject.elementId, 'timeline', timelineData);
    });
    
    // Add Animation Out property
    addSelectProperty(propertyGroup, 'Animation Out', Object.values(ClientConfig.ANIMATION_OUT_TYPES), timelineData.animationOut || ClientConfig.ANIMATION_OUT_TYPES.NONE, value => {
      timelineData.animationOut = value;
      updateElementProperty(fabricObject.elementId, 'timeline', timelineData);
    });
    
    // Add the property group to the container
    container.appendChild(propertyGroup);
  }
  
  /**
   * Add properties that can be edited for multiple selected objects
   * * @param {HTMLElement} container - Property panel container
   * @param {Array<Object>} objects - Array of Fabric.js objects
   */
  function addMultipleSelectionProperties(container, objects) {
    // Create property group
    const propertyGroup = createPropertyGroup('Multiple Selection', 'multiple-selection-group', false); // Not collapsible
    
    // --- Common Geometry ---
    // Note: Editing X, Y, Width, Height for multiple objects is complex.
    // Usually, alignment/distribution buttons are provided instead.
    // For simplicity, we'll omit direct editing here but could add alignment later.
    
    // Add common Rotation property
    // Find if all objects have the same angle
    const firstAngle = objects[0].angle;
    const sameAngle = objects.every(obj => obj.angle === firstAngle);
    addNumberProperty(propertyGroup, 'Rotation', sameAngle ? firstAngle : '', value => {
      if (value === '' || isNaN(value)) return; // Handle empty or non-numeric input
      objects.forEach(obj => {
        obj.set('angle', value);
        updateElementProperty(obj.elementId, 'angle', value);
      });
      editorState.canvas.renderAll();
    }, 0, 360);
    
    // --- Common Appearance ---
    
    // Add common Fill Color property (if applicable)
    const canFill = objects.every(obj => obj.type !== 'image');
    if (canFill) {
      const firstFill = objects[0].type === 'text' || objects[0].type === 'textbox' ? objects[0].backgroundColor : objects[0].fill;
      const sameFill = objects.every(obj => (obj.type === 'text' || obj.type === 'textbox' ? obj.backgroundColor : obj.fill) === firstFill);
      addColorProperty(propertyGroup, 'Fill Color', sameFill ? firstFill : '', value => {
        if (value === '') return; // Handle empty input if color picker allows it
        objects.forEach(obj => {
          if (obj.type === 'text' || obj.type === 'textbox') {
            obj.set('backgroundColor', value);
          } else {
            obj.set('fill', value);
          }
          updateElementProperty(obj.elementId, 'color', value);
        });
        editorState.canvas.renderAll();
      });
    }
    
    // Add common Opacity property
    const firstOpacity = Math.round((objects[0].opacity || 1) * 100);
    const sameOpacity = objects.every(obj => Math.round((obj.opacity || 1) * 100) === firstOpacity);
    addRangeProperty(propertyGroup, 'Opacity', sameOpacity ? firstOpacity : 50, value => { // Default to 50 if different
      const newOpacity = value / 100;
      objects.forEach(obj => {
        obj.set('opacity', newOpacity);
        updateElementProperty(obj.elementId, 'opacity', value); // Store as 0-100
      });
      editorState.canvas.renderAll();
    }, 0, 100, 1);
    
    // --- Common Text (if all are text) ---
    const allText = objects.every(obj => obj.type === 'text' || obj.type === 'textbox');
    if (allText) {
      // Add common Font Family
      const firstFont = objects[0].fontFamily;
      const sameFont = objects.every(obj => obj.fontFamily === firstFont);
      addSelectProperty(propertyGroup, 'Font', ClientConfig.FONTS, sameFont ? firstFont : '', value => {
        if (value === '') return;
        objects.forEach(obj => {
          obj.set('fontFamily', value);
          updateElementProperty(obj.elementId, 'font', value);
        });
        editorState.canvas.renderAll();
      });
      
      // Add common Font Size
      const firstFontSize = objects[0].fontSize;
      const sameFontSize = objects.every(obj => obj.fontSize === firstFontSize);
      addNumberProperty(propertyGroup, 'Font Size', sameFontSize ? firstFontSize : '', value => {
        if (value === '' || isNaN(value)) return;
        objects.forEach(obj => {
          obj.set('fontSize', value);
          updateElementProperty(obj.elementId, 'fontSize', value);
        });
        editorState.canvas.renderAll();
      }, 6, 120);
      
      // Add common Font Color
      const firstFontColor = objects[0].fill;
      const sameFontColor = objects.every(obj => obj.fill === firstFontColor);
      addColorProperty(propertyGroup, 'Font Color', sameFontColor ? firstFontColor : '', value => {
        if (value === '') return;
        objects.forEach(obj => {
          obj.set('fill', value);
          updateElementProperty(obj.elementId, 'fontColor', value);
        });
        editorState.canvas.renderAll();
      });
    }
    
    // Add the property group to the container
    container.appendChild(propertyGroup);
  }
  
  // --- Helper Functions for Creating Property UI ---
  
  /**
   * Creates a collapsible property group container.
   * * @param {string} title - The title of the group.
   * @param {string} groupId - A unique ID for the group state.
   * @param {boolean} [collapsible=true] - Whether the group can be collapsed.
   * @returns {HTMLElement} The content div of the property group.
   */
  function createPropertyGroup(title, groupId, collapsible = true) {
    const groupDiv = document.createElement('div');
    groupDiv.className = 'property-group';
    
    const headerDiv = document.createElement('div');
    headerDiv.className = 'property-group-header';
    headerDiv.innerHTML = `${title} ${collapsible ? '<span class="toggle-icon">▼</span>' : ''}`;
    
    const contentDiv = document.createElement('div');
    contentDiv.className = 'property-group-content';
    
    // Set initial collapsed state based on propertyPanelState
    const isOpen = propertyPanelState.groupOpen[groupId] !== undefined ? propertyPanelState.groupOpen[groupId] : true; // Default to open
    contentDiv.style.display = isOpen ? 'block' : 'none';
    if (headerDiv.querySelector('.toggle-icon')) {
      headerDiv.querySelector('.toggle-icon').textContent = isOpen ? '▼' : '►';
    }
    
    if (collapsible) {
      headerDiv.onclick = () => {
        const currentlyOpen = contentDiv.style.display === 'block';
        contentDiv.style.display = currentlyOpen ? 'none' : 'block';
        propertyPanelState.groupOpen[groupId] = !currentlyOpen; // Update state
        if (headerDiv.querySelector('.toggle-icon')) {
          headerDiv.querySelector('.toggle-icon').textContent = currentlyOpen ? '►' : '▼';
        }
      };
    }
    
    groupDiv.appendChild(headerDiv);
    groupDiv.appendChild(contentDiv);
    
    return contentDiv; // Return the content div where properties will be added
  }
  
  /**
   * Creates a standard row for a property.
   * * @param {HTMLElement} parentGroup - The property group content div to add the row to.
   * @param {string} label - The label text for the property.
   * @returns {object} An object containing the row element and the field container element.
   */
  function addPropertyRow(parentGroup, label) {
    const rowDiv = document.createElement('div');
    rowDiv.className = 'property-row';
    
    const labelDiv = document.createElement('div');
    labelDiv.className = 'property-label';
    labelDiv.textContent = label;
    
    const fieldDiv = document.createElement('div');
    fieldDiv.className = 'property-field';
    
    rowDiv.appendChild(labelDiv);
    rowDiv.appendChild(fieldDiv);
    parentGroup.appendChild(rowDiv);
    
    return { row: rowDiv, field: fieldDiv };
  }
  
  /**
   * Adds a number input property.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {number} initialValue - The initial value.
   * @param {Function} onChange - Callback function when the value changes.
   * @param {number} [min] - Minimum allowed value.
   * @param {number} [max] - Maximum allowed value.
   * @param {number} [step=1] - Step increment.
   * @returns {HTMLElement} The input element.
   */
  function addNumberProperty(parentGroup, label, initialValue, onChange, min, max, step = 1) {
    const { field } = addPropertyRow(parentGroup, label);
    const input = document.createElement('input');
    input.type = 'number';
    input.className = 'form-control form-control-sm'; // Use Bootstrap classes if available
    input.value = initialValue !== null && initialValue !== undefined ? Number(initialValue).toFixed(step < 1 ? 1 : 0) : '';
    input.step = step;
    if (min !== undefined) input.min = min;
    if (max !== undefined) input.max = max;
    
    input.onchange = (e) => {
      let value = parseFloat(e.target.value);
      if (isNaN(value)) value = 0; // Default to 0 if invalid
      if (min !== undefined) value = Math.max(min, value);
      if (max !== undefined) value = Math.min(max, value);
      e.target.value = value.toFixed(step < 1 ? 1 : 0); // Update input display
      onChange(value);
    };
    // Also trigger on input for immediate feedback (optional)
    input.oninput = (e) => {
       let value = parseFloat(e.target.value);
       if (!isNaN(value)) {
          if (min !== undefined && value < min) return; // Prevent exceeding bounds during input
          if (max !== undefined && value > max) return;
          onChange(value);
       }
    };
    
    field.appendChild(input);
    return input;
  }
  
  /**
   * Adds a color picker property.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {string} initialValue - The initial color value (hex).
   * @param {Function} onChange - Callback function when the value changes.
   * @returns {HTMLElement} The input element.
   */
  function addColorProperty(parentGroup, label, initialValue, onChange) {
    const { field } = addPropertyRow(parentGroup, label);
    const input = document.createElement('input');
    input.type = 'color';
    input.value = initialValue || '#000000'; // Default to black if no value
    
    // Add a preview swatch
    const swatch = document.createElement('span');
    swatch.className = 'color-preview';
    swatch.style.backgroundColor = input.value;
    swatch.style.marginLeft = '5px'; // Add some space
    
    input.oninput = (e) => { // Use oninput for live preview
      swatch.style.backgroundColor = e.target.value;
      onChange(e.target.value);
    };
    
    field.appendChild(input);
    field.appendChild(swatch);
    return input;
  }
  
  /**
   * Adds a range slider property.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {number} initialValue - The initial value.
   * @param {Function} onChange - Callback function when the value changes.
   * @param {number} [min=0] - Minimum value.
   * @param {number} [max=100] - Maximum value.
   * @param {number} [step=1] - Step increment.
   * @returns {HTMLElement} The input element.
   */
  function addRangeProperty(parentGroup, label, initialValue, onChange, min = 0, max = 100, step = 1) {
    const { field } = addPropertyRow(parentGroup, label);
    const slider = document.createElement('input');
    slider.type = 'range';
    slider.min = min;
    slider.max = max;
    slider.step = step;
    slider.value = initialValue !== null && initialValue !== undefined ? initialValue : (min + max) / 2;
    
    const valueDisplay = document.createElement('span');
    valueDisplay.textContent = slider.value;
    valueDisplay.style.marginLeft = '10px';
    
    slider.oninput = (e) => {
      const value = parseInt(e.target.value, 10);
      valueDisplay.textContent = value;
      onChange(value);
    };
    
    field.appendChild(slider);
    field.appendChild(valueDisplay);
    return slider;
  }
  
  /**
   * Adds a checkbox property.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {boolean} initialValue - The initial checked state.
   * @param {Function} onChange - Callback function when the value changes.
   * @returns {HTMLElement} The input element.
   */
  function addCheckboxProperty(parentGroup, label, initialValue, onChange) {
    const { row, field } = addPropertyRow(parentGroup, label);
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = !!initialValue; // Ensure boolean
    
    input.onchange = (e) => {
      onChange(e.target.checked);
    };
    
    // Adjust styling for checkbox (remove label text, put checkbox in field)
    row.querySelector('.property-label').textContent = ''; // Clear default label text
    field.style.textAlign = 'left'; // Align checkbox left
    field.appendChild(input);
    // Add label text next to checkbox
    const labelText = document.createElement('label');
    labelText.textContent = label;
    labelText.style.marginLeft = '5px';
    labelText.style.verticalAlign = 'middle';
    // Link label to checkbox for accessibility
    const inputId = `checkbox-${label.replace(/\s+/g, '-')}-${Date.now()}`;
    input.id = inputId;
    labelText.htmlFor = inputId;
    field.appendChild(labelText);
    
    return input;
  }
  
  /**
   * Adds a text input property.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {string} initialValue - The initial text value.
   * @param {Function} onChange - Callback function when the value changes.
   * @returns {HTMLElement} The input element.
   */
  function addTextProperty(parentGroup, label, initialValue, onChange) {
    const { field } = addPropertyRow(parentGroup, label);
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'form-control form-control-sm';
    input.value = initialValue || '';
    
    input.onchange = (e) => { // Trigger on change (blur)
      onChange(e.target.value);
    };
    input.onkeyup = (e) => { // Optional: trigger on key up for faster feedback
       // onChange(e.target.value);
    };
    
    field.appendChild(input);
    return input;
  }
  
  /**
   * Adds a textarea property for multi-line text input.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {string} initialValue - The initial text value.
   * @param {Function} onChange - Callback function when the value changes.
   * @returns {HTMLElement} The textarea element.
   */
  function addTextareaProperty(parentGroup, label, initialValue, onChange) {
    const { row, field } = addPropertyRow(parentGroup, label);
    const textarea = document.createElement('textarea');
    textarea.className = 'form-control form-control-sm';
    textarea.rows = 3; // Adjust rows as needed
    textarea.value = initialValue || '';
    
    textarea.onchange = (e) => { // Trigger on change (blur)
      onChange(e.target.value);
    };
    textarea.onkeyup = (e) => { // Optional: trigger on key up
       // onChange(e.target.value);
    };
    
    field.appendChild(textarea);
    return row; // Return the row for potential hiding/showing
  }
  
  /**
   * Adds a select dropdown property.
   * * @param {HTMLElement} parentGroup - The property group content div.
   * @param {string} label - The property label.
   * @param {Array<string>} options - Array of option strings.
   * @param {string} initialValue - The initially selected value.
   * @param {Function} onChange - Callback function when the value changes.
   * @returns {HTMLElement} The select element.
   */
  function addSelectProperty(parentGroup, label, options, initialValue, onChange) {
    const { field } = addPropertyRow(parentGroup, label);
    const select = document.createElement('select');
    select.className = 'form-control form-control-sm';
    
    options.forEach(optionValue => {
      const option = document.createElement('option');
      option.value = optionValue;
      option.textContent = optionValue;
      if (optionValue === initialValue) {
        option.selected = true;
      }
      select.appendChild(option);
    });
    
    select.onchange = (e) => {
      onChange(e.target.value);
    };
    
    field.appendChild(select);
    return select;
  }
  
  
  // --- Data Update Function ---
  
  /**
   * Update a specific property of an element in the editorState.project.elements array.
   * * @param {string} elementId - The ID of the element to update.
   * @param {string} propertyName - The name of the property to update.
   * @param {*} value - The new value for the property.
   */
  function updateElementProperty(elementId, propertyName, value) {
    if (!editorState.project || !editorState.project.elements || !elementId) {
      console.error('Cannot update element property: Invalid state or elementId.');
      return;
    }
    
    const elementIndex = editorState.project.elements.findIndex(e => e.elementId === elementId);
    
    if (elementIndex === -1) {
      console.warn(`Element with ID ${elementId} not found in project data.`);
      return;
    }
    
    // Update the property
    editorState.project.elements[elementIndex][propertyName] = value;
    
    // Mark project as modified (optional, for save indication)
    // editorState.isModified = true; 
    
    console.log(`Updated property '${propertyName}' for element ${elementId} to:`, value);
    
    // Optionally, trigger an immediate save or add to undo stack here
    // addToUndoStack(); // Add state change to undo stack
  }
  