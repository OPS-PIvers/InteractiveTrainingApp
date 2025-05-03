/**
 * MediaProcessor class handles processing of different media types
 * for the Interactive Training Projects Web App
 */
class MediaProcessor {
    /**
     * Creates a new MediaProcessor instance
     * 
     * @param {DriveManager} driveManager - DriveManager instance for file operations
     */
    constructor(driveManager) {
      this.driveManager = driveManager;
    }
    
    /**
     * Processes an image file
     * 
     * @param {string} projectFolderId - ID of the project folder
     * @param {Blob} imageBlob - Image blob to process
     * @param {string} fileName - Name for the file
     * @param {Object} options - Processing options
     * @return {Object} Processed image info or null if failed
     */
    processImage(projectFolderId, imageBlob, fileName, options = {}) {
      try {
        // Set default options
        const defaultOptions = {
          createThumbnail: true,
          maxWidth: 1200,
          maxHeight: 800,
          quality: 0.85,
          resizeIfLarger: true
        };
        
        const processingOptions = { ...defaultOptions, ...options };
        
        // Validate file type
        if (!this.driveManager.isAllowedFileType(imageBlob, 'IMAGE')) {
          logError(`Invalid image file type: ${imageBlob.getContentType()}`);
          return null;
        }
        
        // Resize image if needed (currently not implemented in Apps Script)
        // In a future implementation, this would resize the image if it's larger than the max dimensions
        // For now, we'll just use the original image
        
        // Upload to Drive
        const uploadResult = this.driveManager.uploadImage(
          projectFolderId, 
          imageBlob, 
          fileName, 
          processingOptions.createThumbnail
        );
        
        if (!uploadResult) {
          logError(`Failed to upload image: ${fileName}`);
          return null;
        }
        
        // Format result
        const result = {
          type: 'image',
          fileName: fileName,
          fileId: uploadResult.fileId,
          fileUrl: uploadResult.fileUrl,
          downloadUrl: uploadResult.downloadUrl,
          contentType: imageBlob.getContentType(),
          size: imageBlob.getBytes().length
        };
        
        // Add thumbnail info if available
        if (uploadResult.thumbnailUrl) {
          result.thumbnailUrl = uploadResult.thumbnailUrl;
          result.thumbnailId = uploadResult.thumbnailId;
        }
        
        logInfo(`Processed image: ${fileName}, ID: ${result.fileId}`);
        return result;
      } catch (error) {
        logError(`Failed to process image: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets image dimensions from a blob
     * Note: This is limited in Apps Script which doesn't have native image processing
     * This is a placeholder for future implementation using advanced services
     * 
     * @param {Blob} imageBlob - Image blob to analyze
     * @return {Object} Object with width and height properties, or null if failed
     */
    getImageDimensions(imageBlob) {
      try {
        // Apps Script doesn't have native image processing
        // In a full implementation, we'd use advanced services or external APIs
        
        // For now, return placeholder dimensions
        return {
          width: 800,
          height: 600,
          estimatedOnly: true
        };
      } catch (error) {
        logError(`Failed to get image dimensions: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Processes a YouTube video URL
     * 
     * @param {string} videoUrl - YouTube video URL
     * @return {Object} Processed video info or null if failed
     */
    processYouTubeVideo(videoUrl) {
      try {
        // Extract video ID from URL
        const videoId = this.extractYouTubeVideoId(videoUrl);
        
        if (!videoId) {
          logError(`Failed to extract YouTube video ID from URL: ${videoUrl}`);
          return null;
        }
        
        // Generate embedded URL
        const embedUrl = `https://www.youtube.com/embed/${videoId}`;
        
        // Generate thumbnail URLs
        const thumbnailUrlDefault = `https://img.youtube.com/vi/${videoId}/default.jpg`;
        const thumbnailUrlMq = `https://img.youtube.com/vi/${videoId}/mqdefault.jpg`;
        const thumbnailUrlHq = `https://img.youtube.com/vi/${videoId}/hqdefault.jpg`;
        const thumbnailUrlSd = `https://img.youtube.com/vi/${videoId}/sddefault.jpg`;
        const thumbnailUrlMax = `https://img.youtube.com/vi/${videoId}/maxresdefault.jpg`;
        
        // Format result
        const result = {
          type: 'youtube',
          videoId: videoId,
          videoUrl: videoUrl,
          embedUrl: embedUrl,
          thumbnails: {
            default: thumbnailUrlDefault,
            medium: thumbnailUrlMq,
            high: thumbnailUrlHq,
            standard: thumbnailUrlSd,
            maxres: thumbnailUrlMax
          }
        };
        
        logInfo(`Processed YouTube video: ${videoId}`);
        return result;
      } catch (error) {
        logError(`Failed to process YouTube video: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Extracts YouTube video ID from various URL formats
     * 
     * @param {string} url - YouTube URL
     * @return {string|null} YouTube video ID or null if invalid
     */
    extractYouTubeVideoId(url) {
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
     * Processes an audio file
     * 
     * @param {string} projectFolderId - ID of the project folder
     * @param {Blob} audioBlob - Audio blob to process
     * @param {string} fileName - Name for the file
     * @param {Object} options - Processing options
     * @return {Object} Processed audio info or null if failed
     */
    processAudio(projectFolderId, audioBlob, fileName, options = {}) {
      try {
        // Set default options
        const defaultOptions = {
          generateWaveform: false // Future enhancement
        };
        
        const processingOptions = { ...defaultOptions, ...options };
        
        // Validate file type
        if (!this.driveManager.isAllowedFileType(audioBlob, 'AUDIO')) {
          logError(`Invalid audio file type: ${audioBlob.getContentType()}`);
          return null;
        }
        
        // Get media folder
        const mediaFolder = this.driveManager.getMediaFolder(projectFolderId);
        
        if (!mediaFolder) {
          logError(`Media folder not found for project: ${projectFolderId}`);
          return null;
        }
        
        // Upload to Drive
        const folder = this.driveManager.getFolder(mediaFolder.id);
        if (!folder) {
          logError(`Failed to get media folder: ${mediaFolder.id}`);
          return null;
        }
        
        const file = folder.createFile(audioBlob);
        file.setName(fileName);
        
        // Format result
        const result = {
          type: 'audio',
          fileName: fileName,
          fileId: file.getId(),
          fileUrl: file.getUrl(),
          downloadUrl: file.getDownloadUrl(),
          contentType: audioBlob.getContentType(),
          size: audioBlob.getBytes().length
        };
        
        // Generate placeholder image for audio (future enhancement)
        if (processingOptions.generateWaveform) {
          // Placeholder for waveform generation
          // This would require advanced services or external APIs
        }
        
        logInfo(`Processed audio: ${fileName}, ID: ${result.fileId}`);
        return result;
      } catch (error) {
        logError(`Failed to process audio: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets audio metadata from a blob
     * Note: This is limited in Apps Script which doesn't have native audio processing
     * This is a placeholder for future implementation using advanced services
     * 
     * @param {Blob} audioBlob - Audio blob to analyze
     * @return {Object} Object with audio metadata, or null if failed
     */
    getAudioMetadata(audioBlob) {
      try {
        // Apps Script doesn't have native audio processing
        // In a full implementation, we'd use advanced services or external APIs
        
        // For now, return placeholder metadata
        return {
          duration: 0,
          bitrate: 0,
          channels: 2,
          estimatedOnly: true
        };
      } catch (error) {
        logError(`Failed to get audio metadata: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Processes a URL to determine the media type and extract relevant information
     * 
     * @param {string} url - URL to process
     * @return {Object} Processed media info or null if failed
     */
    processMediaUrl(url) {
      try {
        // Check if it's a YouTube video
        const youtubeVideoId = this.extractYouTubeVideoId(url);
        if (youtubeVideoId) {
          return this.processYouTubeVideo(url);
        }
        
        // Check if it's an image URL
        if (url.match(/\.(jpeg|jpg|gif|png)$/i)) {
          return {
            type: 'image',
            imageUrl: url,
            isExternal: true
          };
        }
        
        // Check if it's an audio URL
        if (url.match(/\.(mp3|wav|ogg)$/i)) {
          return {
            type: 'audio',
            audioUrl: url,
            isExternal: true
          };
        }
        
        // Unknown media type
        return {
          type: 'unknown',
          url: url,
          isExternal: true
        };
      } catch (error) {
        logError(`Failed to process media URL: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Generates a placeholder image for media types that don't have visuals
     * 
     * @param {string} mediaType - Type of media (audio, etc)
     * @param {Object} options - Options for the placeholder
     * @return {Blob} Image blob for the placeholder, or null if failed
     */
    generatePlaceholderImage(mediaType, options = {}) {
      try {
        // Apps Script doesn't have native image generation
        // In a full implementation, we'd use advanced services or external APIs
        
        // For now, return null as a placeholder
        logWarning('Placeholder image generation not implemented yet');
        return null;
      } catch (error) {
        logError(`Failed to generate placeholder image: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Creates a thumbnail for an image
     * 
     * @param {Blob} imageBlob - Original image blob
     * @param {Object} options - Options for the thumbnail
     * @return {Blob} Thumbnail blob, or null if failed
     */
    createThumbnail(imageBlob, options = {}) {
      try {
        // Set default options
        const defaultOptions = {
          maxWidth: 200,
          maxHeight: 200,
          quality: 0.8
        };
        
        const thumbnailOptions = { ...defaultOptions, ...options };
        
        // Apps Script doesn't have built-in image resizing
        // In a full implementation, we'd use advanced services or external APIs
        
        // For now, just return the original image as a placeholder
        logWarning('Thumbnail creation using original image as placeholder');
        return imageBlob;
      } catch (error) {
        logError(`Failed to create thumbnail: ${error.message}`);
        return null;
      }
    }
  }