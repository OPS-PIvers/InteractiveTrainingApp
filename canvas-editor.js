/**
 * canvas-editor.js - Core canvas editing functionality for the Interactive Training Projects Web App
 * Integrates with Fabric.js to provide a canvas-based visual editor
 */

// Global editor state
let editorState = {
    project: null,
    currentSlide: null,
    canvas: null,
    selectedObjects: [],
    copiedObjects: [],
    undoStack: [],
    redoStack: [],
    isPlaying: false,
    currentTime: 0,
    zoomLevel: 1
  };
  
  /**
   * Initialize the editor with project data
   * Sets up the canvas, loads the project slides, and initializes the UI
   * 
   * @param {Object} project - Project data
   */
  function initEditor(project) {
    console.log('Initializing editor with project:', project);
    
    // Store project data
    editorState.project = project;
    
    // Initialize the canvas
    initializeCanvas();
    
    // Load slides
    loadSlides();
    
    // Set up event handlers for canvas interaction
    setupEventHandlers();
    
    // Initialize property panel
    initPropertyPanel();
    
    // Initialize timeline if needed
    if (hasTimelineContent()) {
      initializeTimeline();
    }
    
    // Set the first slide as active
    if (project.slides && project.slides.length > 0) {
      loadSlide(project.slides[0].slideId);
    }
    
    // Set up window resize handler
    window.addEventListener('resize', debounce(resizeCanvas, 250));
    
    console.log('Editor initialization complete');
  }
  
  /**
   * Initialize the canvas with Fabric.js
   */
  function initializeCanvas() {
    // Get the canvas element
    const canvasElement = document.getElementById('editor-canvas');
    
    // Create a new Fabric canvas
    editorState.canvas = new fabric.Canvas(canvasElement, {
      width: ClientConfig.UI_SETTINGS.CANVAS.DEFAULT_WIDTH,
      height: ClientConfig.UI_SETTINGS.CANVAS.DEFAULT_HEIGHT,
      backgroundColor: '#FFFFFF',
      preserveObjectStacking: true
    });
    
    // Set up canvas size to match the UI settings
    resizeCanvas();
    
    // Set up event listeners
    editorState.canvas.on({
      'object:modified': onObjectModified,
      'object:selected': onObjectSelected,
      'selection:created': onSelectionCreated,
      'selection:updated': onSelectionUpdated,
      'selection:cleared': onSelectionCleared,
      'mouse:down': onMouseDown,
      'mouse:move': onMouseMove,
      'mouse:up': onMouseUp
    });
    
    console.log('Canvas initialized');
  }
  
  /**
   * Resize the canvas based on container size and maintain aspect ratio
   */
  function resizeCanvas() {
    const containerWidth = document.getElementById('canvas-panel').clientWidth;
    const containerHeight = document.getElementById('canvas-panel').clientHeight;
    
    // Calculate the best fit for the canvas while maintaining aspect ratio
    const canvasWidth = ClientConfig.UI_SETTINGS.CANVAS.DEFAULT_WIDTH;
    const canvasHeight = ClientConfig.UI_SETTINGS.CANVAS.DEFAULT_HEIGHT;
    const canvasRatio = canvasWidth / canvasHeight;
    
    let scaleFactor;
    if (containerWidth / containerHeight > canvasRatio) {
      // Container is wider than canvas ratio
      scaleFactor = (containerHeight * 0.9) / canvasHeight;
    } else {
      // Container is taller than canvas ratio
      scaleFactor = (containerWidth * 0.9) / canvasWidth;
    }
    
    // Update zoom level
    editorState.zoomLevel = scaleFactor;
    
    // Apply scaling to canvas container
    const scaler = document.querySelector('.canvas-scaler');
    scaler.style.transform = `translate(-50%, -50%) scale(${scaleFactor})`;
    
    console.log(`Canvas resized with scale factor: ${scaleFactor}`);
  }
  
  /**
   * Load the slides from the project data
   * Populates the slides panel with slide thumbnails
   */
  function loadSlides() {
    const slidesList = document.getElementById('slides-list');
    slidesList.innerHTML = '';
    
    if (!editorState.project.slides || editorState.project.slides.length === 0) {
      slidesList.innerHTML = '<p>No slides found. Add a slide to get started.</p>';
      return;
    }
    
    // Create a slide thumbnail for each slide
    editorState.project.slides.forEach(slide => {
      const slideElement = document.createElement('div');
      slideElement.className = 'slide-thumbnail';
      slideElement.dataset.slideId = slide.slideId;
      
      // Create the thumbnail content
      slideElement.innerHTML = `
        <div class="slide-thumbnail-img" style="background-color: ${slide.backgroundColor || '#FFFFFF'};">
          ${slide.fileUrl ? '<img src="' + slide.fileUrl + '" alt="Slide Media" style="max-width: 100%; max-height: 100%;">' : ''}
        </div>
        <div class="slide-thumbnail-title">${slide.title || 'Untitled Slide'}</div>
        <div class="slide-thumbnail-actions">
          <button class="btn btn-secondary btn-sm edit-slide-btn" data-slide-id="${slide.slideId}">Edit</button>
          <button class="btn btn-secondary btn-sm delete-slide-btn" data-slide-id="${slide.slideId}">Delete</button>
        </div>
      `;
      
      // Add click event to load the slide
      slideElement.addEventListener('click', function() {
        loadSlide(slide.slideId);
      });
      
      // Add the slide thumbnail to the list
      slidesList.appendChild(slideElement);
    });
    
    console.log(`Loaded ${editorState.project.slides.length} slides`);
  }
  
  /**
   * Load a slide into the canvas
   * 
   * @param {string} slideId - ID of the slide to load
   */
  function loadSlide(slideId) {
    console.log(`Loading slide: ${slideId}`);
    
    // Find the slide in the project data
    const slide = editorState.project.slides.find(s => s.slideId === slideId);
    
    if (!slide) {
      console.error(`Slide not found: ${slideId}`);
      return;
    }
    
    // Store the current slide
    editorState.currentSlide = slide;
    
    // Clear the canvas
    editorState.canvas.clear();
    
    // Set the canvas background color
    editorState.canvas.setBackgroundColor(slide.backgroundColor || '#FFFFFF', editorState.canvas.renderAll.bind(editorState.canvas));
    
    // Load slide media if available
    if (slide.fileUrl) {
      if (slide.fileType === 'Image') {
        fabric.Image.fromURL(slide.fileUrl, (img) => {
          // Scale the image to fit the canvas
          const canvasWidth = editorState.canvas.width;
          const canvasHeight = editorState.canvas.height;
          
          // Calculate scale factor to fit the image within the canvas
          const scaleFactor = Math.min(
            canvasWidth / img.width,
            canvasHeight / img.height
          );
          
          // Apply scaling
          img.scale(scaleFactor);
          
          // Center the image on the canvas
          img.set({
            left: canvasWidth / 2,
            top: canvasHeight / 2,
            originX: 'center',
            originY: 'center',
            selectable: false
          });
          
          // Add the image to the canvas
          editorState.canvas.add(img);
          editorState.canvas.sendToBack(img);
          
          // Render the canvas
          editorState.canvas.renderAll();
        });
      } else if (slide.fileType === 'YouTube Video') {
        // For YouTube videos, we'll add a placeholder image for now
        // In a full implementation, this would be more sophisticated
        console.log('Loading YouTube video placeholder');
        
        // Create a placeholder rectangle
        const videoPlaceholder = new fabric.Rect({
          width: editorState.canvas.width * 0.8,
          height: editorState.canvas.height * 0.6,
          left: editorState.canvas.width / 2,
          top: editorState.canvas.height / 2,
          originX: 'center',
          originY: 'center',
          fill: '#000000',
          selectable: false
        });
        
        // Add play button icon in the center
        const playIcon = new fabric.Text('▶', {
          fontSize: 60,
          fill: '#FFFFFF',
          left: editorState.canvas.width / 2,
          top: editorState.canvas.height / 2,
          originX: 'center',
          originY: 'center',
          selectable: false
        });
        
        // Add to canvas
        editorState.canvas.add(videoPlaceholder);
        editorState.canvas.add(playIcon);
        editorState.canvas.sendToBack(videoPlaceholder);
        
        // Add video URL as text below
        const videoUrl = new fabric.Text(slide.fileUrl, {
          fontSize: 14,
          fill: '#000000',
          left: editorState.canvas.width / 2,
          top: editorState.canvas.height / 2 + videoPlaceholder.height / 2 + 20,
          originX: 'center',
          originY: 'top',
          selectable: false
        });
        
        editorState.canvas.add(videoUrl);
      }
    }
    
    // Load slide elements
    if (editorState.project.elements) {
      const slideElements = editorState.project.elements.filter(e => e.slideId === slideId);
      
      // Load each element
      slideElements.forEach(element => {
        loadElement(element);
      });
    }
    
    // Update the active slide in the slides panel
    const slideThumbnails = document.querySelectorAll('.slide-thumbnail');
    slideThumbnails.forEach(thumbnail => {
      if (thumbnail.dataset.slideId === slideId) {
        thumbnail.classList.add('active');
      } else {
        thumbnail.classList.remove('active');
      }
    });
    
    // Initialize or update timeline for this slide if needed
    if (hasTimelineContent()) {
      updateTimeline();
    }
    
    console.log(`Slide ${slideId} loaded successfully`);
  }
  
  /**
   * Load an element onto the canvas
   * 
   * @param {Object} element - Element data
   */
  function loadElement(element) {
    console.log(`Loading element: ${element.elementId}`, element);
    
    // Create the element based on its type
    let fabricObject;
    
    switch (element.type) {
      case 'Rectangle':
      case 'rectangle':
        fabricObject = createRectangleElement(element);
        break;
        
      case 'Rounded Rectangle':
      case 'roundedRectangle':
        fabricObject = createRoundedRectangleElement(element);
        break;
        
      case 'Circle':
      case 'circle':
        fabricObject = createCircleElement(element);
        break;
        
      case 'Arrow':
      case 'arrow':
        fabricObject = createArrowElement(element);
        break;
        
      case 'Text':
      case 'text':
        fabricObject = createTextElement(element);
        break;
        
      case 'Hotspot':
      case 'hotspot':
        fabricObject = createHotspotElement(element);
        break;
        
      default:
        console.warn(`Unknown element type: ${element.type}`);
        return;
    }
    
    // Set common properties
    if (fabricObject) {
      // Set the element ID
      fabricObject.elementId = element.elementId;
      
      // Set interaction properties
      fabricObject.interactionType = element.interactionType;
      fabricObject.triggers = element.triggers;
      
      // Check if the element should be initially hidden
      if (element.initiallyHidden) {
        fabricObject.visible = false;
      }
      
      // Add the object to the canvas
      editorState.canvas.add(fabricObject);
      
      console.log(`Element ${element.elementId} loaded successfully`);
    }
  }
  
  /**
   * Check if the current slide has timeline content
   * 
   * @return {boolean} True if the slide has timeline content
   */
  function hasTimelineContent() {
    // Check if current slide has a video or audio file
    if (editorState.currentSlide && 
        (editorState.currentSlide.fileType === 'YouTube Video' || 
         editorState.currentSlide.fileType === 'Audio')) {
      return true;
    }
    
    // Check if any elements have timeline properties
    if (editorState.project.elements) {
      const slideElements = editorState.project.elements.filter(e => 
        e.slideId === editorState.currentSlide?.slideId && e.timeline
      );
      
      return slideElements.length > 0;
    }
    
    return false;
  }
  
  /**
   * Initialize the timeline for the current slide
   */
  function initializeTimeline() {
    console.log('Initializing timeline');
    
    // Show the timeline panel
    document.getElementById('timeline-panel').style.display = 'block';
    
    // Set up event listeners for timeline controls
    document.getElementById('playButton').addEventListener('click', togglePlayback);
    document.getElementById('timeline-ruler').addEventListener('click', seekToPosition);
    
    // Initialize timeline state
    updateTimeline();
    
    console.log('Timeline initialized');
  }
  
  /**
   * Update the timeline display for the current slide
   */
  function updateTimeline() {
    console.log('Updating timeline');
    
    // Get the timeline tracks container
    const tracksContainer = document.getElementById('timeline-tracks');
    tracksContainer.innerHTML = '';
    
    // Get elements for this slide that have timeline properties
    if (editorState.project.elements && editorState.currentSlide) {
      const slideElements = editorState.project.elements.filter(e => 
        e.slideId === editorState.currentSlide.slideId && e.timeline
      );
      
      // Create a track for each element
      slideElements.forEach(element => {
        const track = document.createElement('div');
        track.className = 'timeline-track';
        track.dataset.elementId = element.elementId;
        
        // Add element name
        const trackLabel = document.createElement('div');
        trackLabel.className = 'timeline-track-label';
        trackLabel.textContent = element.nickname || `Element ${element.sequence}`;
        track.appendChild(trackLabel);
        
        // Add track content
        const trackContent = document.createElement('div');
        trackContent.className = 'timeline-track-content';
        
        // Add keyframes if available
        if (element.timeline) {
          if (element.timeline.startTime !== undefined) {
            const startKeyframe = document.createElement('div');
            startKeyframe.className = 'timeline-keyframe timeline-keyframe-start';
            startKeyframe.style.left = `${(element.timeline.startTime / 100)}%`;
            startKeyframe.dataset.time = element.timeline.startTime;
            startKeyframe.dataset.elementId = element.elementId;
            startKeyframe.dataset.keyframeType = 'start';
            trackContent.appendChild(startKeyframe);
          }
          
          if (element.timeline.endTime !== undefined) {
            const endKeyframe = document.createElement('div');
            endKeyframe.className = 'timeline-keyframe timeline-keyframe-end';
            endKeyframe.style.left = `${(element.timeline.endTime / 100)}%`;
            endKeyframe.dataset.time = element.timeline.endTime;
            endKeyframe.dataset.elementId = element.elementId;
            endKeyframe.dataset.keyframeType = 'end';
            trackContent.appendChild(endKeyframe);
          }
        }
        
        track.appendChild(trackContent);
        tracksContainer.appendChild(track);
      });
    }
    
    // Update time display
    document.getElementById('current-time').textContent = formatTime(editorState.currentTime);
    document.getElementById('total-time').textContent = formatTime(100); // Default to 100 seconds for now
    
    // Update playhead position
    document.getElementById('timeline-playhead').style.left = `${editorState.currentTime}%`;
  }
  
  /**
   * Format time in seconds to MM:SS.MSS format
   * 
   * @param {number} seconds - Time in seconds
   * @return {string} Formatted time string
   */
  function formatTime(seconds) {
    const mins = Math.floor(seconds / 60);
    const secs = Math.floor(seconds % 60);
    const ms = Math.floor((seconds % 1) * 1000);
    
    return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}.${ms.toString().padStart(3, '0')}`;
  }
  
  /**
   * Toggle playback of the timeline
   */
  function togglePlayback() {
    editorState.isPlaying = !editorState.isPlaying;
    
    const playButton = document.getElementById('playButton');
    playButton.textContent = editorState.isPlaying ? 'Pause' : 'Play';
    
    if (editorState.isPlaying) {
      playTimeline();
    } else {
      pauseTimeline();
    }
  }
  
  /**
   * Start playing the timeline
   */
  function playTimeline() {
    // Implement playback logic
    console.log('Playing timeline');
    
    // This would be a simple playback implementation
    // In a full implementation, this would control media playback as well
    editorState.playbackInterval = setInterval(() => {
      editorState.currentTime += 0.1;
      if (editorState.currentTime > 100) {
        editorState.currentTime = 0;
      }
      
      // Update UI
      document.getElementById('current-time').textContent = formatTime(editorState.currentTime);
      document.getElementById('timeline-playhead').style.left = `${editorState.currentTime}%`;
      
      // Check for keyframes at current time
      checkKeyframes();
    }, 100);
  }
  
  /**
   * Pause timeline playback
   */
  function pauseTimeline() {
    console.log('Pausing timeline');
    
    clearInterval(editorState.playbackInterval);
  }
  
  /**
   * Seek to a position in the timeline
   * 
   * @param {Event} event - Click event on the timeline
   */
  function seekToPosition(event) {
    const ruler = document.getElementById('timeline-ruler');
    const rulerRect = ruler.getBoundingClientRect();
    
    const clickX = event.clientX - rulerRect.left;
    const percentage = (clickX / rulerRect.width) * 100;
    
    editorState.currentTime = percentage;
    
    // Update UI
    document.getElementById('current-time').textContent = formatTime(editorState.currentTime);
    document.getElementById('timeline-playhead').style.left = `${percentage}%`;
    
    // Check for keyframes at current time
    checkKeyframes();
  }
  
  /**
   * Check for keyframes at the current time
   */
  function checkKeyframes() {
    // Get elements for this slide that have timeline properties
    if (editorState.project.elements && editorState.currentSlide) {
      const slideElements = editorState.project.elements.filter(e => 
        e.slideId === editorState.currentSlide.slideId && e.timeline
      );
      
      // Check each element's keyframes
      slideElements.forEach(element => {
        const fabricObject = getFabricObjectByElementId(element.elementId);
        
        if (fabricObject && element.timeline) {
          // Check if we're at or past the start time
          if (element.timeline.startTime !== undefined && 
              editorState.currentTime >= element.timeline.startTime) {
            // Check if we're before the end time or if there is no end time
            if (element.timeline.endTime === undefined || 
                editorState.currentTime <= element.timeline.endTime) {
              // Element should be visible
              fabricObject.visible = true;
            } else {
              // Past the end time, hide the element
              fabricObject.visible = false;
            }
          } else {
            // Before start time, hide the element
            fabricObject.visible = false;
          }
        }
      });
      
      // Render the canvas to show/hide elements
      editorState.canvas.renderAll();
    }
  }
  
  /**
   * Get a Fabric.js object by its element ID
   * 
   * @param {string} elementId - ID of the element
   * @return {Object} Fabric.js object
   */
  function getFabricObjectByElementId(elementId) {
    const objects = editorState.canvas.getObjects();
    return objects.find(obj => obj.elementId === elementId);
  }
  
  /**
   * Set up event handlers for editor interaction
   */
  function setupEventHandlers() {
    console.log('Setting up event handlers');
    
    // Example of keyboard shortcut handler
    document.addEventListener('keydown', function(event) {
      // Check if an input element is focused
      if (document.activeElement.tagName === 'INPUT' || 
          document.activeElement.tagName === 'TEXTAREA') {
        return;
      }
      
      // Handle keyboard shortcuts
      if (event.ctrlKey || event.metaKey) {
        switch (event.key) {
          case 's':
            // Save
            event.preventDefault();
            saveProject();
            break;
            
          case 'z':
            // Undo
            event.preventDefault();
            undo();
            break;
            
          case 'y':
            // Redo
            event.preventDefault();
            redo();
            break;
            
          case 'c':
            // Copy
            event.preventDefault();
            copySelectedObjects();
            break;
            
          case 'v':
            // Paste
            event.preventDefault();
            pasteObjects();
            break;
            
          case 'd':
            // Duplicate
            event.preventDefault();
            duplicateSelectedObjects();
            break;
        }
      } else {
        switch (event.key) {
          case 'Delete':
          case 'Backspace':
            // Delete
            if (editorState.canvas.getActiveObject()) {
              event.preventDefault();
              deleteSelectedObjects();
            }
            break;
        }
      }
    });
  }
  
  /**
   * Handle changes to objects on the canvas
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onObjectModified(event) {
    const modifiedObject = event.target;
    
    // Find the corresponding element in the project data
    if (modifiedObject.elementId && editorState.project.elements) {
      const element = editorState.project.elements.find(e => e.elementId === modifiedObject.elementId);
      
      if (element) {
        // Update element properties based on modified object
        element.left = modifiedObject.left;
        element.top = modifiedObject.top;
        element.width = modifiedObject.width * modifiedObject.scaleX;
        element.height = modifiedObject.height * modifiedObject.scaleY;
        element.angle = modifiedObject.angle;
        
        // Reset scale after applying it to width/height
        modifiedObject.scaleX = 1;
        modifiedObject.scaleY = 1;
        modifiedObject.width = element.width;
        modifiedObject.height = element.height;
        
        // For text elements, update text content
        if (modifiedObject.type === 'text' && element.type === 'Text') {
          element.text = modifiedObject.text;
        }
        
        console.log(`Element ${element.elementId} updated`);
        
        // Update property panel if this object is selected
        if (editorState.canvas.getActiveObject() === modifiedObject) {
          updatePropertyPanel(modifiedObject);
        }
      }
    }
    
    // Add to undo stack
    addToUndoStack();
  }
  
  /**
   * Handle object selection
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onObjectSelected(event) {
    const selectedObject = event.target;
    editorState.selectedObjects = [selectedObject];
    
    console.log('Object selected:', selectedObject);
    
    // Update the property panel
    updatePropertyPanel(selectedObject);
  }
  
  /**
   * Handle selection creation (multiple objects)
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onSelectionCreated(event) {
    const selection = event.target;
    editorState.selectedObjects = selection.getObjects();
    
    console.log('Selection created:', editorState.selectedObjects);
    
    // Update the property panel for multiple selection
    updatePropertyPanelForMultipleSelection(editorState.selectedObjects);
  }
  
  /**
   * Handle selection update
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onSelectionUpdated(event) {
    const selection = event.target;
    editorState.selectedObjects = selection.getObjects();
    
    console.log('Selection updated:', editorState.selectedObjects);
    
    // Update the property panel for multiple selection
    updatePropertyPanelForMultipleSelection(editorState.selectedObjects);
  }
  
  /**
   * Handle selection cleared
   */
  function onSelectionCleared() {
    editorState.selectedObjects = [];
    
    console.log('Selection cleared');
    
    // Hide property panel or show empty state
    hidePropertyPanel();
  }
  
  /**
   * Handle mouse down on canvas
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onMouseDown(event) {
    // Handle canvas click events
  }
  
  /**
   * Handle mouse move on canvas
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onMouseMove(event) {
    // Handle canvas mouse move events
  }
  
  /**
   * Handle mouse up on canvas
   * 
   * @param {Object} event - Fabric.js event object
   */
  function onMouseUp(event) {
    // Handle canvas mouse up events
  }
  
  /**
   * Add a slide to the project
   * 
   * @param {Object} slideData - Slide data
   */
  function addSlide(slideData) {
    // Generate a unique ID for the slide
    const slideId = generateUUID();
    
    // Create the slide object
    const slide = {
      slideId: slideId,
      title: slideData.title || 'Untitled Slide',
      backgroundColor: slideData.backgroundColor || '#FFFFFF',
      fileType: slideData.fileType || '',
      fileUrl: slideData.fileUrl || '',
      showControls: slideData.showControls || false,
      slideNumber: editorState.project.slides.length + 1
    };
    
    // Add the slide to the project
    editorState.project.slides.push(slide);
    
    // Reload the slides panel
    loadSlides();
    
    // Load the new slide
    loadSlide(slideId);
    
    // Hide the add slide modal
    hideAddSlideModal();
    
    console.log(`Slide ${slideId} added successfully`);
    
    // Show success message
    showAlert('Slide added successfully!', 'success');
  }
  
  /**
   * Add a media element to the canvas
   * 
   * @param {Object} mediaData - Media data
   */
  function addMediaElement(mediaData) {
    // This would be implemented in a full version
    console.log('Adding media element:', mediaData);
    
    // Hide the add media modal
    hideAddMediaModal();
    
    // Show success message
    showAlert('Media added successfully!', 'success');
  }
  
  /**
   * Add an image element to the canvas
   * 
   * @param {string} imageUrl - URL of the image
   */
  function addImageElement(imageUrl) {
    fabric.Image.fromURL(imageUrl, (img) => {
      // Scale the image to a reasonable size
      const maxSize = 300;
      const scale = Math.min(maxSize / img.width, maxSize / img.height, 1);
      
      img.scale(scale);
      
      // Position the image in the center of the canvas
      img.set({
        left: editorState.canvas.width / 2,
        top: editorState.canvas.height / 2,
        originX: 'center',
        originY: 'center',
        elementId: generateUUID()
      });
      
      // Add the image to the canvas
      editorState.canvas.add(img);
      editorState.canvas.setActiveObject(img);
      
      // Add the image to the project data
      const element = {
        elementId: img.elementId,
        nickname: 'Image',
        slideId: editorState.currentSlide.slideId,
        sequence: getNextElementSequence(),
        type: 'Image',
        left: img.left,
        top: img.top,
        width: img.width * scale,
        height: img.height * scale,
        angle: 0,
        initiallyHidden: false,
        opacity: 100,
        imageUrl: imageUrl
      };
      
      editorState.project.elements.push(element);
      
      // Update the property panel
      updatePropertyPanel(img);
      
      // Add to undo stack
      addToUndoStack();
      
      console.log(`Image element ${element.elementId} added successfully`);
    });
  }
  
  /**
   * Create a shape on the canvas
   * 
   * @param {string} shapeType - Type of shape to create
   */
  function createShape(shapeType) {
    console.log(`Creating shape: ${shapeType}`);
    
    // Generate a unique ID for the element
    const elementId = generateUUID();
    
    // Create the element object based on type
    let element = {
      elementId: elementId,
      nickname: capitalizeFirstLetter(shapeType),
      slideId: editorState.currentSlide.slideId,
      sequence: getNextElementSequence(),
      type: capitalizeFirstLetter(shapeType),
      left: editorState.canvas.width / 2,
      top: editorState.canvas.height / 2,
      angle: 0,
      initiallyHidden: false,
      opacity: 100,
      color: ClientConfig.DEFAULT_COLORS.ELEMENT,
      outline: false,
      outlineWidth: 1,
      outlineColor: ClientConfig.DEFAULT_COLORS.OUTLINE,
      shadow: false
    };
    
    // Set type-specific properties
    switch (shapeType) {
      case 'rectangle':
        element.width = ClientConfig.ELEMENT_TYPES.RECTANGLE.DEFAULT_WIDTH;
        element.height = ClientConfig.ELEMENT_TYPES.RECTANGLE.DEFAULT_HEIGHT;
        break;
        
      case 'circle':
        element.width = ClientConfig.ELEMENT_TYPES.CIRCLE.DEFAULT_RADIUS * 2;
        element.height = ClientConfig.ELEMENT_TYPES.CIRCLE.DEFAULT_RADIUS * 2;
        break;
        
      case 'text':
        element.width = ClientConfig.ELEMENT_TYPES.TEXT.DEFAULT_WIDTH;
        element.height = ClientConfig.ELEMENT_TYPES.TEXT.DEFAULT_HEIGHT;
        element.text = 'Double click to edit text';
        element.font = ClientConfig.ELEMENT_TYPES.TEXT.DEFAULT_FONT;
        element.fontSize = ClientConfig.ELEMENT_TYPES.TEXT.DEFAULT_FONT_SIZE;
        element.fontColor = ClientConfig.DEFAULT_COLORS.TEXT;
        break;
        
      case 'hotspot':
        element.width = ClientConfig.ELEMENT_TYPES.HOTSPOT.DEFAULT_WIDTH;
        element.height = ClientConfig.ELEMENT_TYPES.HOTSPOT.DEFAULT_HEIGHT;
        element.color = 'rgba(66, 133, 244, 0.3)';
        element.interactionType = INTERACTION_TYPES.REVEAL;
        element.triggers = TRIGGER_TYPES.CLICK;
        break;
        
      case 'arrow':
        element.width = ClientConfig.ELEMENT_TYPES.ARROW.DEFAULT_LENGTH;
        element.height = ClientConfig.LEMENT_TYPES.ARROW.DEFAULT_WIDTH;
        break;
    }
    
    // Add the element to the project data
    if (!editorState.project.elements) {
      editorState.project.elements = [];
    }
    
    editorState.project.elements.push(element);
    
    // Create the Fabric.js object
    let fabricObject;
    
    switch (shapeType) {
      case 'rectangle':
        fabricObject = createRectangleElement(element);
        break;
        
      case 'circle':
        fabricObject = createCircleElement(element);
        break;
        
      case 'text':
        fabricObject = createTextElement(element);
        break;
        
      case 'hotspot':
        fabricObject = createHotspotElement(element);
        break;
        
      case 'arrow':
        fabricObject = createArrowElement(element);
        break;
    }
    
    // Add the object to the canvas
    if (fabricObject) {
      fabricObject.elementId = elementId;
      editorState.canvas.add(fabricObject);
      editorState.canvas.setActiveObject(fabricObject);
      
      // Add to undo stack
      addToUndoStack();
      
      console.log(`Shape ${shapeType} created with ID ${elementId}`);
    }
  }
  
  /**
   * Delete selected objects from the canvas
   */
  function deleteSelectedObjects() {
    const activeObject = editorState.canvas.getActiveObject();
    
    if (!activeObject) {
      return;
    }
    
    if (activeObject.type === 'activeSelection') {
      // Multiple objects selected
      const objects = activeObject.getObjects();
      
      // Remove each object
      objects.forEach(object => {
        // Remove from project data
        if (object.elementId && editorState.project.elements) {
          editorState.project.elements = editorState.project.elements.filter(
            e => e.elementId !== object.elementId
          );
        }
        
        // Remove from canvas
        editorState.canvas.remove(object);
      });
      
      // Clear the selection
      editorState.canvas.discardActiveObject();
    } else {
      // Single object selected
      // Remove from project data
      if (activeObject.elementId && editorState.project.elements) {
        editorState.project.elements = editorState.project.elements.filter(
          e => e.elementId !== activeObject.elementId
        );
      }
      
      // Remove from canvas
      editorState.canvas.remove(activeObject);
    }
    
    // Render the canvas
    editorState.canvas.renderAll();
    
    // Add to undo stack
    addToUndoStack();
    
    console.log('Objects deleted');
  }
  
  /**
   * Copy selected objects to clipboard
   */
  function copySelectedObjects() {
    const activeObject = editorState.canvas.getActiveObject();
    
    if (!activeObject) {
      return;
    }
    
    // Clear the copied objects array
    editorState.copiedObjects = [];
    
    if (activeObject.type === 'activeSelection') {
      // Multiple objects selected
      const objects = activeObject.getObjects();
      
      // Copy each object
      objects.forEach(object => {
        editorState.copiedObjects.push(object);
      });
    } else {
      // Single object selected
      editorState.copiedObjects.push(activeObject);
    }
    
    console.log(`Copied ${editorState.copiedObjects.length} objects`);
    
    // Show success message
    showAlert(`Copied ${editorState.copiedObjects.length} objects`, 'success');
  }
  
  /**
   * Paste copied objects onto the canvas
   */
  function pasteObjects() {
    if (editorState.copiedObjects.length === 0) {
      return;
    }
    
    // Create new objects based on the copied objects
    editorState.copiedObjects.forEach(object => {
      // Clone the object
      object.clone(clonedObject => {
        // Generate a new element ID
        const newElementId = generateUUID();
        
        // Offset the position slightly
        clonedObject.set({
          left: clonedObject.left + 10,
          top: clonedObject.top + 10,
          elementId: newElementId
        });
        
        // Add to canvas
        editorState.canvas.add(clonedObject);
        
        // Add to project data
        if (editorState.project.elements) {
          // Find the original element
          const originalElement = editorState.project.elements.find(e => e.elementId === object.elementId);
          
          if (originalElement) {
            // Clone the element and update properties
            const newElement = JSON.parse(JSON.stringify(originalElement));
            newElement.elementId = newElementId;
            newElement.left = clonedObject.left;
            newElement.top = clonedObject.top;
            newElement.sequence = getNextElementSequence();
            
            // Add to project data
            editorState.project.elements.push(newElement);
          }
        }
      });
    });
    
    // Render the canvas
    editorState.canvas.renderAll();
    
    // Add to undo stack
    addToUndoStack();
    
    console.log(`Pasted ${editorState.copiedObjects.length} objects`);
    
    // Show success message
    showAlert(`Pasted ${editorState.copiedObjects.length} objects`, 'success');
  }
  
  /**
   * Duplicate selected objects
   */
  function duplicateSelectedObjects() {
    // First copy, then paste
    copySelectedObjects();
    pasteObjects();
  }
  
  /**
   * Bring selected objects forward in the stacking order
   */
  function bringForward() {
    const activeObject = editorState.canvas.getActiveObject();
    
    if (!activeObject) {
      return;
    }
    
    if (activeObject.type === 'activeSelection') {
      // Multiple objects selected
      const objects = activeObject.getObjects();
      
      // Bring each object forward
      objects.forEach(object => {
        editorState.canvas.bringForward(object);
      });
    } else {
      // Single object selected
      editorState.canvas.bringForward(activeObject);
    }
    
    // Render the canvas
    editorState.canvas.renderAll();
    
    console.log('Objects brought forward');
  }
  
  /**
   * Send selected objects backward in the stacking order
   */
  function sendBackward() {
    const activeObject = editorState.canvas.getActiveObject();
    
    if (!activeObject) {
      return;
    }
    
    if (activeObject.type === 'activeSelection') {
      // Multiple objects selected
      const objects = activeObject.getObjects();
      
      // Send each object backward
      objects.forEach(object => {
        editorState.canvas.sendBackward(object);
      });
    } else {
      // Single object selected
      editorState.canvas.sendBackward(activeObject);
    }
    
    // Render the canvas
    editorState.canvas.renderAll();
    
    console.log('Objects sent backward');
  }
  
  /**
   * Add the current canvas state to the undo stack
   */
  function addToUndoStack() {
    // In a full implementation, this would save the current state for undo/redo
    // For simplicity, we're not implementing the full undo/redo functionality
    
    // Clear the redo stack when a new action is performed
    editorState.redoStack = [];
  }
  
  /**
   * Undo the last action
   */
  function undo() {
    // This would be implemented in a full version
    console.log('Undo');
  }
  
  /**
   * Redo the last undone action
   */
  function redo() {
    // This would be implemented in a full version
    console.log('Redo');
  }
  
  /**
   * Serialize the project for saving
   * 
   * @return {Object} Serialized project data
   */
  function serializeProject() {
    // Create a copy of the project data to avoid modifying the original
    const project = JSON.parse(JSON.stringify(editorState.project));
    
    // Update any properties that might have changed in the canvas
    if (project.elements) {
      const canvasObjects = editorState.canvas.getObjects();
      
      canvasObjects.forEach(object => {
        if (object.elementId) {
          const element = project.elements.find(e => e.elementId === object.elementId);
          
          if (element) {
            // Update properties from the canvas object
            element.left = object.left;
            element.top = object.top;
            element.width = object.width * object.scaleX;
            element.height = object.height * object.scaleY;
            element.angle = object.angle;
            
            // For text elements, update text content
            if (object.type === 'text' && element.type === 'Text') {
              element.text = object.text;
            }
          }
        }
      });
    }
    
    return project;
  }
  
  /**
   * Get the next element sequence number
   * 
   * @return {number} Next sequence number
   */
  function getNextElementSequence() {
    if (!editorState.project.elements || editorState.project.elements.length === 0) {
      return 1;
    }
    
    // Find the maximum sequence number and add 1
    return Math.max(...editorState.project.elements.map(e => e.sequence || 0)) + 1;
  }
  
  /**
   * Create a debounced function
   * 
   * @param {Function} func - Function to debounce
   * @param {number} wait - Wait time in milliseconds
   * @return {Function} Debounced function
   */
  function debounce(func, wait) {
    let timeout;
    return function() {
      const context = this;
      const args = arguments;
      
      clearTimeout(timeout);
      
      timeout = setTimeout(() => {
        func.apply(context, args);
      }, wait);
    };
  }
  
  /**
   * Capitalize the first letter of a string
   * 
   * @param {string} string - String to capitalize
   * @return {string} Capitalized string
   */
  function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
  }
  
  /**
   * Generate a UUID for unique identifiers
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