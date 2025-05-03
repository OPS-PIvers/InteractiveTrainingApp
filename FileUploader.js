/**
 * FileUploader class handles file uploads for the Interactive Training Projects Web App
 * Manages validation, processing, and storage of uploaded files
 */
class FileUploader {
    /**
     * Creates a new FileUploader instance
     * 
     * @param {DriveManager} driveManager - DriveManager instance for file operations
     * @param {MediaProcessor} mediaProcessor - MediaProcessor instance for media processing
     */
    constructor(driveManager, mediaProcessor) {
      this.driveManager = driveManager;
      this.mediaProcessor = mediaProcessor;
    }
    
    /**
     * Uploads a file to a project
     * 
     * @param {string} projectId - ID of the project
     * @param {Blob} fileBlob - File blob to upload
     * @param {string} fileName - Name for the file
     * @param {Object} options - Upload options
     * @return {Object} Upload result info or null if failed
     */
    uploadFile(projectId, fileBlob, fileName, options = {}) {
      try {
        // Set default options
        const defaultOptions = {
          fileType: null, // Auto-detect if null
          createThumbnail: true,
          processMedia: true,
          maxFileSizeMB: 10
        };
        
        const uploadOptions = { ...defaultOptions, ...options };
        
        // Validate file size
        const fileSizeMB = fileBlob.getBytes().length / (1024 * 1024);
        if (fileSizeMB > uploadOptions.maxFileSizeMB) {
          return { 
            success: false, 
            message: `File size exceeds maximum allowed (${uploadOptions.maxFileSizeMB}MB)`,
            fileSize: fileSizeMB 
          };
        }
        
        // Get project folder
        const project = this.getProjectInfo(projectId);
        
        if (!project || !project.folderId) {
          return { 
            success: false, 
            message: 'Project folder not found'
          };
        }
        
        // Determine file type if not specified
        let fileType = uploadOptions.fileType;
        
        if (!fileType) {
          fileType = this.detectFileType(fileBlob, fileName);
        }
        
        // Process file based on type
        let result;
        
        switch (fileType) {
          case 'image':
            result = this.mediaProcessor.processImage(
              project.folderId, 
              fileBlob, 
              fileName,
              { createThumbnail: uploadOptions.createThumbnail }
            );
            break;
            
          case 'audio':
            result = this.mediaProcessor.processAudio(
              project.folderId, 
              fileBlob, 
              fileName
            );
            break;
            
          default:
            // Generic file upload
            const mediaFolder = this.driveManager.getMediaFolder(project.folderId);
            
            if (!mediaFolder) {
              return { 
                success: false, 
                message: 'Media folder not found' 
              };
            }
            
            const file = this.driveManager.createFile(
              mediaFolder.id,
              fileName,
              fileBlob
            );
            
            if (!file) {
              return { 
                success: false, 
                message: 'Failed to create file'
              };
            }
            
            result = {
              type: 'file',
              fileName: fileName,
              fileId: file.fileId,
              fileUrl: file.fileUrl,
              downloadUrl: file.downloadUrl,
              contentType: fileBlob.getContentType(),
              size: fileBlob.getBytes().length
            };
        }
        
        if (!result) {
          return { 
            success: false, 
            message: 'Failed to process file'
          };
        }
        
        // Return success result
        return { 
          success: true, 
          message: 'File uploaded successfully',
          fileType: fileType,
          file: result
        };
      } catch (error) {
        logError(`Failed to upload file: ${error.message}`);
        return { 
          success: false, 
          message: `Error: ${error.message}`
        };
      }
    }
    
    /**
     * Processes a URL to upload or reference external media
     * 
     * @param {string} projectId - ID of the project
     * @param {string} url - URL to process
     * @param {Object} options - Processing options
     * @return {Object} Processing result info
     */
    processUrl(projectId, url, options = {}) {
      try {
        // Set default options
        const defaultOptions = {
          saveLocally: true, // Whether to download and save the file locally
          validateUrl: true  // Whether to validate the URL is accessible
        };
        
        const processOptions = { ...defaultOptions, ...options };
        
        // Process URL to determine media type
        const mediaInfo = this.mediaProcessor.processMediaUrl(url);
        
        if (!mediaInfo) {
          return { 
            success: false, 
            message: 'Failed to process URL'
          };
        }
        
        // For YouTube videos, we don't need to download anything
        if (mediaInfo.type === 'youtube') {
          return { 
            success: true, 
            message: 'YouTube video processed successfully',
            media: mediaInfo
          };
        }
        
        // For other URL types, optionally download the file
        if (processOptions.saveLocally && (mediaInfo.type === 'image' || mediaInfo.type === 'audio')) {
          // Try to fetch the file
          try {
            const response = UrlFetchApp.fetch(url);
            const fileBlob = response.getBlob();
            const fileName = this.getFileNameFromUrl(url);
            
            // Upload the file
            const uploadResult = this.uploadFile(projectId, fileBlob, fileName, {
              fileType: mediaInfo.type
            });
            
            if (uploadResult.success) {
              return { 
                success: true, 
                message: 'Media downloaded and saved successfully',
                media: uploadResult.file,
                originalUrl: url
              };
            } else {
              // If download failed, still return the URL reference
              logWarning(`Failed to download media, using URL reference: ${url}`);
              return { 
                success: true, 
                message: 'Using URL reference (download failed)',
                media: mediaInfo,
                downloadFailed: true
              };
            }
          } catch (fetchError) {
            logWarning(`Failed to fetch URL, using reference only: ${fetchError.message}`);
            return { 
              success: true, 
              message: 'Using URL reference (fetch failed)',
              media: mediaInfo,
              fetchFailed: true
            };
          }
        }
        
        // If not downloading or for other types, just return the media info
        return { 
          success: true, 
          message: 'URL processed as reference',
          media: mediaInfo
        };
      } catch (error) {
        logError(`Failed to process URL: ${error.message}`);
        return { 
          success: false, 
          message: `Error: ${error.message}`
        };
      }
    }
    
    /**
     * Validates an upload based on file type and contents
     * 
     * @param {Blob} fileBlob - File blob to validate
     * @param {string} fileName - Name of the file
     * @param {string} allowedTypes - Array of allowed file types (mimeTypes or categories)
     * @return {Object} Validation result with success flag and message
     */
    validateUpload(fileBlob, fileName, allowedTypes = []) {
      try {
        // Check file type based on content type
        const contentType = fileBlob.getContentType();
        const fileType = this.detectFileType(fileBlob, fileName);
        
        // If allowedTypes is empty, allow all file types
        if (allowedTypes.length === 0) {
          return { 
            success: true, 
            message: 'File validated successfully',
            fileType: fileType,
            contentType: contentType
          };
        }
        
        // Check if the file type is in the allowed types
        let isAllowed = false;
        
        // Check by category
        if (allowedTypes.includes('image') && fileType === 'image') {
          isAllowed = true;
        } else if (allowedTypes.includes('audio') && fileType === 'audio') {
          isAllowed = true;
        } else if (allowedTypes.includes('video') && fileType === 'video') {
          isAllowed = true;
        } else if (allowedTypes.includes('document') && fileType === 'document') {
          isAllowed = true;
        }
        
        // Check by mime type
        if (!isAllowed && allowedTypes.includes(contentType)) {
          isAllowed = true;
        }
        
        // Check by file extension
        const extension = fileName.split('.').pop().toLowerCase();
        if (!isAllowed && allowedTypes.includes(`.${extension}`)) {
          isAllowed = true;
        }
        
        if (!isAllowed) {
          return { 
            success: false, 
            message: `File type not allowed. Allowed types: ${allowedTypes.join(', ')}`,
            fileType: fileType,
            contentType: contentType
          };
        }
        
        return { 
          success: true, 
          message: 'File validated successfully',
          fileType: fileType,
          contentType: contentType
        };
      } catch (error) {
        logError(`Failed to validate upload: ${error.message}`);
        return { 
          success: false, 
          message: `Error: ${error.message}`
        };
      }
    }
    
    /**
     * Detects file type based on content type and file name
     * 
     * @param {Blob} fileBlob - File blob to check
     * @param {string} fileName - Name of the file
     * @return {string} File type category (image, audio, video, document, etc.)
     */
    detectFileType(fileBlob, fileName) {
      try {
        const contentType = fileBlob.getContentType();
        const extension = fileName.split('.').pop().toLowerCase();
        
        // Check by content type
        if (contentType.startsWith('image/')) {
          return 'image';
        } else if (contentType.startsWith('audio/')) {
          return 'audio';
        } else if (contentType.startsWith('video/')) {
          return 'video';
        } else if (contentType.startsWith('text/')) {
          return 'document';
        } else if (contentType.includes('pdf')) {
          return 'document';
        } else if (contentType.includes('spreadsheet') || contentType.includes('excel')) {
          return 'spreadsheet';
        } else if (contentType.includes('presentation') || contentType.includes('powerpoint')) {
          return 'presentation';
        }
        
        // Check by file extension
        const imageExtensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg'];
        const audioExtensions = ['mp3', 'wav', 'ogg', 'aac', 'm4a'];
        const videoExtensions = ['mp4', 'webm', 'mov', 'avi', 'mkv'];
        const documentExtensions = ['doc', 'docx', 'txt', 'rtf', 'pdf', 'odt'];
        const spreadsheetExtensions = ['xls', 'xlsx', 'csv', 'ods'];
        const presentationExtensions = ['ppt', 'pptx', 'odp'];
        
        if (imageExtensions.includes(extension)) {
          return 'image';
        } else if (audioExtensions.includes(extension)) {
          return 'audio';
        } else if (videoExtensions.includes(extension)) {
          return 'video';
        } else if (documentExtensions.includes(extension)) {
          return 'document';
        } else if (spreadsheetExtensions.includes(extension)) {
          return 'spreadsheet';
        } else if (presentationExtensions.includes(extension)) {
          return 'presentation';
        }
        
        // Default to generic file type
        return 'file';
      } catch (error) {
        logError(`Failed to detect file type: ${error.message}`);
        return 'file'; // Default to generic file type on error
      }
    }
    
    /**
     * Gets file name from a URL
     * 
     * @param {string} url - URL to extract file name from
     * @return {string} File name or a generated name if extraction fails
     */
    getFileNameFromUrl(url) {
      try {
        // Try to extract file name from URL
        const urlPath = url.split('?')[0]; // Remove query parameters
        const pathParts = urlPath.split('/');
        let fileName = pathParts[pathParts.length - 1];
        
        // If file name is empty or has no extension, generate a name
        if (!fileName || !fileName.includes('.')) {
          const timestamp = new Date().getTime();
          fileName = `file_${timestamp}`;
          
          // Try to add appropriate extension based on URL
          if (url.includes('.jpg') || url.includes('.jpeg')) {
            fileName += '.jpg';
          } else if (url.includes('.png')) {
            fileName += '.png';
          } else if (url.includes('.gif')) {
            fileName += '.gif';
          } else if (url.includes('.mp3')) {
            fileName += '.mp3';
          } else if (url.includes('.mp4')) {
            fileName += '.mp4';
          } else if (url.includes('.pdf')) {
            fileName += '.pdf';
          }
        }
        
        return fileName;
      } catch (error) {
        logError(`Failed to get file name from URL: ${error.message}`);
        return `file_${new Date().getTime()}`; // Generate a generic name on error
      }
    }
    
    /**
     * Gets basic project info needed for uploads
     * 
     * @param {string} projectId - ID of the project
     * @return {Object} Project info with folder IDs
     */
    getProjectInfo(projectId) {
      try {
        // Find the project in the index
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const indexSheet = ss.getSheetByName(SHEET_STRUCTURE.PROJECT_INDEX.TAB_NAME);
        
        if (!indexSheet) {
          logError('Project index sheet not found');
          return null;
        }
        
        // Find project row
        const projectIdCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.PROJECT_ID;
        const projectTitleCol = SHEET_STRUCTURE.PROJECT_INDEX.COLUMNS.TITLE;
        
        const data = indexSheet.getDataRange().getValues();
        let projectRow = null;
        
        for (let i = 1; i < data.length; i++) { // Skip header row
          if (data[i][projectIdCol] === projectId) {
            projectRow = data[i];
            break;
          }
        }
        
        if (!projectRow) {
          logError(`Project not found in index: ${projectId}`);
          return null;
        }
        
        const projectTitle = projectRow[projectTitleCol];
        const projectTabName = sanitizeSheetName(projectTitle);
        
        // Get project sheet
        const projectSheet = ss.getSheetByName(projectTabName);
        
        if (!projectSheet) {
          logError(`Project sheet not found: ${projectTabName}`);
          return null;
        }
        
        // Get project folder ID
        const folderIdRow = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.ROW;
        const folderIdCol = SHEET_STRUCTURE.PROJECT_TAB.PROJECT_INFO.FIELDS.PROJECT_FOLDER_ID.VALUE_COL;
        const folderId = projectSheet.getRange(folderIdRow, columnToIndex(folderIdCol) + 1).getValue();
        
        return {
          projectId: projectId,
          title: projectTitle,
          folderId: folderId
        };
      } catch (error) {
        logError(`Failed to get project info: ${error.message}`);
        return null;
      }
    }
  }