/**
 * element-factory.js - Element creation and configuration for the Interactive Training Projects Web App
 * Creates and configures different types of elements for the canvas editor
 */

/**
 * Create a rectangle element
 * 
 * @param {Object} element - Element data
 * @return {Object} Fabric.js rectangle object
 */
function createRectangleElement(element) {
    // Create a new fabric.Rect object
    const rect = new fabric.Rect({
      left: element.left,
      top: element.top,
      width: element.width,
      height: element.height,
      fill: element.color || DEFAULT_COLORS.ELEMENT,
      stroke: element.outline ? (element.outlineColor || DEFAULT_COLORS.OUTLINE) : undefined,
      strokeWidth: element.outline ? (element.outlineWidth || 1) : 0,
      opacity: element.opacity ? element.opacity / 100 : 1,
      angle: element.angle || 0,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true
    });
    
    // Add shadow if specified
    if (element.shadow) {
      rect.setShadow({
        color: 'rgba(0,0,0,0.3)',
        offsetX: 3,
        offsetY: 3,
        blur: 5
      });
    }
    
    return rect;
  }
  
  /**
   * Create a rounded rectangle element
   * 
   * @param {Object} element - Element data
   * @return {Object} Fabric.js rectangle object with rounded corners
   */
  function createRoundedRectangleElement(element) {
    // Create a new fabric.Rect object with rounded corners
    const rect = new fabric.Rect({
      left: element.left,
      top: element.top,
      width: element.width,
      height: element.height,
      fill: element.color || DEFAULT_COLORS.ELEMENT,
      stroke: element.outline ? (element.outlineColor || DEFAULT_COLORS.OUTLINE) : undefined,
      strokeWidth: element.outline ? (element.outlineWidth || 1) : 0,
      opacity: element.opacity ? element.opacity / 100 : 1,
      angle: element.angle || 0,
      rx: ELEMENT_TYPES.ROUNDED_RECTANGLE.DEFAULT_RADIUS,
      ry: ELEMENT_TYPES.ROUNDED_RECTANGLE.DEFAULT_RADIUS,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true
    });
    
    // Add shadow if specified
    if (element.shadow) {
      rect.setShadow({
        color: 'rgba(0,0,0,0.3)',
        offsetX: 3,
        offsetY: 3,
        blur: 5
      });
    }
    
    return rect;
  }
  
  /**
   * Create a circle element
   * 
   * @param {Object} element - Element data
   * @return {Object} Fabric.js circle object
   */
  function createCircleElement(element) {
    // Calculate radius from width (assuming width equals height)
    const radius = Math.min(element.width, element.height) / 2;
    
    // Create a new fabric.Circle object
    const circle = new fabric.Circle({
      left: element.left,
      top: element.top,
      radius: radius,
      fill: element.color || DEFAULT_COLORS.ELEMENT,
      stroke: element.outline ? (element.outlineColor || DEFAULT_COLORS.OUTLINE) : undefined,
      strokeWidth: element.outline ? (element.outlineWidth || 1) : 0,
      opacity: element.opacity ? element.opacity / 100 : 1,
      angle: element.angle || 0,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true
    });
    
    // Add shadow if specified
    if (element.shadow) {
      circle.setShadow({
        color: 'rgba(0,0,0,0.3)',
        offsetX: 3,
        offsetY: 3,
        blur: 5
      });
    }
    
    return circle;
  }
  
  /**
   * Create a text element
   * 
   * @param {Object} element - Element data
   * @return {Object} Fabric.js text object
   */
  function createTextElement(element) {
    // Create a new fabric.Textbox object
    const text = new fabric.Textbox(element.text || 'Double click to edit text', {
      left: element.left,
      top: element.top,
      width: element.width,
      fill: element.fontColor || DEFAULT_COLORS.TEXT,
      fontFamily: element.font || FONTS[0],
      fontSize: element.fontSize || ELEMENT_TYPES.TEXT.DEFAULT_FONT_SIZE,
      textAlign: 'center',
      backgroundColor: element.color || 'transparent',
      opacity: element.opacity ? element.opacity / 100 : 1,
      angle: element.angle || 0,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true,
      editable: true
    });
    
    // Add border if specified
    if (element.outline) {
      text.set({
        stroke: element.outlineColor || DEFAULT_COLORS.OUTLINE,
        strokeWidth: element.outlineWidth || 1
      });
    }
    
    // Add shadow if specified
    if (element.shadow) {
      text.setShadow({
        color: 'rgba(0,0,0,0.3)',
        offsetX: 3,
        offsetY: 3,
        blur: 5
      });
    }
    
    return text;
  }
  
  /**
   * Create a hotspot element
   * 
   * @param {Object} element - Element data
   * @return {Object} Fabric.js rectangle object with transparency
   */
  function createHotspotElement(element) {
    // Create a new fabric.Rect object with transparency
    const hotspot = new fabric.Rect({
      left: element.left,
      top: element.top,
      width: element.width,
      height: element.height,
      fill: element.color || 'rgba(66, 133, 244, 0.3)',
      stroke: element.outline ? (element.outlineColor || '#4285F4') : '#4285F4',
      strokeWidth: element.outline ? (element.outlineWidth || 1) : 1,
      strokeDashArray: [5, 5], // Dashed line
      opacity: element.opacity ? element.opacity / 100 : 0.3,
      angle: element.angle || 0,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true
    });
    
    return hotspot;
  }
  
  /**
   * Create an arrow element
   * 
   * @param {Object} element - Element data
   * @return {Object} Fabric.js path object representing an arrow
   */
  function createArrowElement(element) {
    // Calculate arrow dimensions
    const arrowLength = element.width;
    const arrowWidth = element.height;
    const headWidth = arrowWidth * 3;
    const headLength = arrowLength * 0.3;
    
    // Create arrow path
    const pathString = `M 0 ${arrowWidth/2} 
                       L ${arrowLength - headLength} ${arrowWidth/2} 
                       L ${arrowLength - headLength} 0 
                       L ${arrowLength} ${arrowWidth*1.5} 
                       L ${arrowLength - headLength} ${arrowWidth*3} 
                       L ${arrowLength - headLength} ${arrowWidth*2.5} 
                       L 0 ${arrowWidth*2.5} z`;
    
    // Create a new fabric.Path object
    const arrow = new fabric.Path(pathString, {
      left: element.left,
      top: element.top,
      fill: element.color || DEFAULT_COLORS.ELEMENT,
      stroke: element.outline ? (element.outlineColor || DEFAULT_COLORS.OUTLINE) : undefined,
      strokeWidth: element.outline ? (element.outlineWidth || 1) : 0,
      opacity: element.opacity ? element.opacity / 100 : 1,
      angle: element.angle || 0,
      width: arrowLength,
      height: arrowWidth * 3,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true
    });
    
    // Add shadow if specified
    if (element.shadow) {
      arrow.setShadow({
        color: 'rgba(0,0,0,0.3)',
        offsetX: 3,
        offsetY: 3,
        blur: 5
      });
    }
    
    return arrow;
  }
  
  /**
   * Create a highlight area element
   * 
   * @param {Object} element - Element data
   * @return {Object} Fabric.js rectangle with spotlight effect
   */
  function createHighlightElement(element) {
    // Create a spotlight effect
    const highlight = new fabric.Rect({
      left: element.left,
      top: element.top,
      width: element.width,
      height: element.height,
      fill: element.color || 'rgba(255, 255, 0, 0.2)',
      stroke: element.outline ? (element.outlineColor || '#FFFF00') : '#FFFF00',
      strokeWidth: element.outline ? (element.outlineWidth || 2) : 2,
      strokeDashArray: [5, 5], // Dashed line
      opacity: element.opacity ? element.opacity / 100 : 0.5,
      angle: element.angle || 0,
      originX: 'center',
      originY: 'center',
      hasControls: true,
      hasBorders: true,
      selectable: true
    });
    
    return highlight;
  }
  
  /**
   * Create an image element
   * 
   * @param {Object} element - Element data
   * @param {Function} callback - Callback function after image is loaded
   * @return {Object} Fabric.js image object
   */
  function createImageElement(element, callback) {
    fabric.Image.fromURL(element.imageUrl, (img) => {
      // Set properties
      img.set({
        left: element.left,
        top: element.top,
        angle: element.angle || 0,
        opacity: element.opacity ? element.opacity / 100 : 1,
        originX: 'center',
        originY: 'center',
        hasControls: true,
        hasBorders: true,
        selectable: true
      });
      
      // Scale the image if width and height are specified
      if (element.width && element.height) {
        // Calculate scale factors
        const scaleX = element.width / img.width;
        const scaleY = element.height / img.height;
        
        img.scale(Math.min(scaleX, scaleY));
      }
      
      // Add border if specified
      if (element.outline) {
        img.set({
          stroke: element.outlineColor || DEFAULT_COLORS.OUTLINE,
          strokeWidth: element.outlineWidth || 1
        });
      }
      
      // Add shadow if specified
      if (element.shadow) {
        img.setShadow({
          color: 'rgba(0,0,0,0.3)',
          offsetX: 3,
          offsetY: 3,
          blur: 5
        });
      }
      
      if (callback) {
        callback(img);
      }
    });
  }