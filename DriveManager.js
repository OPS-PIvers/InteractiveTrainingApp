/**
 * DriveManager class handles Google Drive operations for project files
 * Manages folder creation, file uploads, permissions, and cleanup
 */
class DriveManager {
    /**
     * Creates a new DriveManager instance
     */
    constructor() {
      this.rootFolderId = null;
      this.initRootFolder();
    }
    
    /**
     * Initializes the root folder for the application
     * Creates it if it doesn't exist
     */
    initRootFolder() {
      try {
        // Try to find existing root folder
        const rootFolderName = DRIVE_SETTINGS.ROOT_FOLDER_NAME;
        
        // Search for folders with the correct name
        const folderIterator = DriveApp.getFoldersByName(rootFolderName);
        
        if (folderIterator.hasNext()) {
          // Use existing folder
          const folder = folderIterator.next();
          this.rootFolderId = folder.getId();
          logDebug(`Found existing root folder: ${rootFolderName} (${this.rootFolderId})`);
        } else {
          // Create new root folder
          const folder = DriveApp.createFolder(rootFolderName);
          this.rootFolderId = folder.getId();
          logInfo(`Created new root folder: ${rootFolderName} (${this.rootFolderId})`);
        }
      } catch (error) {
        logError(`Failed to initialize root folder: ${error.message}`);
        throw error;
      }
    }
    
    /**
     * Gets the root folder for the application
     * 
     * @return {Folder} Google Drive Folder object
     */
    getRootFolder() {
      try {
        if (!this.rootFolderId) {
          this.initRootFolder();
        }
        
        return DriveApp.getFolderById(this.rootFolderId);
      } catch (error) {
        logError(`Failed to get root folder: ${error.message}`);
        throw error;
      }
    }
    
    /**
     * Creates a project folder structure
     * 
     * @param {string} projectId - Unique ID for the project
     * @param {string} projectName - Name of the project
     * @return {Object} Object with folder IDs for the project structure
     */
    createProjectFolders(projectId, projectName) {
      try {
        const rootFolder = this.getRootFolder();
        
        // Create main project folder
        const projectFolderName = `${projectName} (${projectId})`;
        const projectFolder = rootFolder.createFolder(projectFolderName);
        const projectFolderId = projectFolder.getId();
        
        // Create media assets subfolder
        const mediaFolder = projectFolder.createFolder(DRIVE_SETTINGS.MEDIA_FOLDER_NAME);
        const mediaFolderId = mediaFolder.getId();
        
        // Create thumbnails subfolder
        const thumbnailsFolder = mediaFolder.createFolder(DRIVE_SETTINGS.THUMBNAIL_FOLDER_NAME);
        const thumbnailsFolderId = thumbnailsFolder.getId();
        
        logInfo(`Created project folders for: ${projectName} (${projectId})`);
        
        return {
          projectFolderId,
          mediaFolderId,
          thumbnailsFolderId
        };
      } catch (error) {
        logError(`Failed to create project folders: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets a folder by ID
     * 
     * @param {string} folderId - ID of the folder to get
     * @return {Folder} Google Drive Folder object
     */
    getFolder(folderId) {
      try {
        return DriveApp.getFolderById(folderId);
      } catch (error) {
        logError(`Failed to get folder (${folderId}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets a file by ID
     * 
     * @param {string} fileId - ID of the file to get
     * @return {File} Google Drive File object
     */
    getFile(fileId) {
      try {
        return DriveApp.getFileById(fileId);
      } catch (error) {
        logError(`Failed to get file (${fileId}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Creates a file in a specific folder
     * 
     * @param {string} folderId - ID of the folder to create the file in
     * @param {string} fileName - Name for the new file
     * @param {Blob|string} content - Content for the file (Blob or string)
     * @param {string} mimeType - MIME type for the file (optional)
     * @return {Object} Object with file ID and URL, or null if failed
     */
    createFile(folderId, fileName, content, mimeType = null) {
      try {
        const folder = this.getFolder(folderId);
        if (!folder) {
          logError(`Folder not found: ${folderId}`);
          return null;
        }
        
        let file;
        
        if (content instanceof Blob) {
          file = folder.createFile(content);
          file.setName(fileName);
        } else if (mimeType) {
          file = folder.createFile(fileName, content, mimeType);
        } else {
          file = folder.createFile(fileName, content, 'text/plain');
        }
        
        logInfo(`Created file: ${fileName} in folder ${folderId}`);
        
        return {
          fileId: file.getId(),
          fileUrl: file.getUrl(),
          downloadUrl: file.getDownloadUrl()
        };
      } catch (error) {
        logError(`Failed to create file (${fileName}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Uploads an image file to the media folder of a project
     * 
     * @param {string} projectFolderId - ID of the project folder
     * @param {Blob} imageBlob - Image blob to upload
     * @param {string} fileName - Name for the file
     * @param {boolean} createThumbnail - Whether to create a thumbnail
     * @return {Object} Object with file info, or null if failed
     */
    uploadImage(projectFolderId, imageBlob, fileName, createThumbnail = true) {
      try {
        const projectFolder = this.getFolder(projectFolderId);
        if (!projectFolder) {
          logError(`Project folder not found: ${projectFolderId}`);
          return null;
        }
        
        // Find media folder
        const mediaFolderName = DRIVE_SETTINGS.MEDIA_FOLDER_NAME;
        let mediaFolder = null;
        
        const folderIterator = projectFolder.getFolders();
        while (folderIterator.hasNext()) {
          const folder = folderIterator.next();
          if (folder.getName() === mediaFolderName) {
            mediaFolder = folder;
            break;
          }
        }
        
        if (!mediaFolder) {
          // Create media folder if it doesn't exist
          mediaFolder = projectFolder.createFolder(mediaFolderName);
        }
        
        // Create file in media folder
        const file = mediaFolder.createFile(imageBlob);
        file.setName(fileName);
        
        const fileInfo = {
          fileId: file.getId(),
          fileUrl: file.getUrl(),
          downloadUrl: file.getDownloadUrl(),
          thumbnailUrl: null
        };
        
        // Create thumbnail if requested
        if (createThumbnail) {
          // Find or create thumbnails folder
          const thumbnailFolderName = DRIVE_SETTINGS.THUMBNAIL_FOLDER_NAME;
          let thumbnailFolder = null;
          
          const subFolderIterator = mediaFolder.getFolders();
          while (subFolderIterator.hasNext()) {
            const folder = subFolderIterator.next();
            if (folder.getName() === thumbnailFolderName) {
              thumbnailFolder = folder;
              break;
            }
          }
          
          if (!thumbnailFolder) {
            thumbnailFolder = mediaFolder.createFolder(thumbnailFolderName);
          }
          
          // Create thumbnail
          try {
            const thumbnailBlob = this.createThumbnail(imageBlob);
            if (thumbnailBlob) {
              const thumbnailFile = thumbnailFolder.createFile(thumbnailBlob);
              thumbnailFile.setName(`thumb_${fileName}`);
              
              fileInfo.thumbnailUrl = thumbnailFile.getDownloadUrl();
              fileInfo.thumbnailId = thumbnailFile.getId();
            }
          } catch (thumbError) {
            logWarning(`Failed to create thumbnail for ${fileName}: ${thumbError.message}`);
          }
        }
        
        logInfo(`Uploaded image: ${fileName} to project folder ${projectFolderId}`);
        return fileInfo;
      } catch (error) {
        logError(`Failed to upload image (${fileName}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Creates a thumbnail from an image blob
     * 
     * @param {Blob} imageBlob - Original image blob
     * @param {number} maxWidth - Maximum width for thumbnail
     * @param {number} maxHeight - Maximum height for thumbnail
     * @return {Blob} Thumbnail blob, or null if failed
     */
    createThumbnail(imageBlob, maxWidth = 200, maxHeight = 200) {
      try {
        // Currently, Apps Script doesn't have built-in image manipulation
        // For thumbnails, we would normally use an external service or advanced API
        // For simplicity, we'll just return the original image until we implement proper thumbnail generation
        
        // In a full implementation, consider using the Advanced Drive API to generate thumbnails
        
        return imageBlob;
      } catch (error) {
        logWarning(`Failed to create thumbnail: ${error.message}`);
        return null;
      }
    }
    
    /**
     * Deletes a file by ID
     * 
     * @param {string} fileId - ID of the file to delete
     * @return {boolean} True if successful
     */
    deleteFile(fileId) {
      try {
        const file = this.getFile(fileId);
        if (!file) {
          logWarning(`File not found for deletion: ${fileId}`);
          return false;
        }
        
        // Move to trash instead of permanent deletion
        file.setTrashed(true);
        
        logInfo(`Deleted file: ${fileId}`);
        return true;
      } catch (error) {
        logError(`Failed to delete file (${fileId}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Gets all files in a folder
     * 
     * @param {string} folderId - ID of the folder
     * @return {Array<Object>} Array of file objects with id, name, and url properties
     */
    getFolderFiles(folderId) {
      try {
        const folder = this.getFolder(folderId);
        if (!folder) {
          logError(`Folder not found: ${folderId}`);
          return [];
        }
        
        const files = [];
        const fileIterator = folder.getFiles();
        
        while (fileIterator.hasNext()) {
          const file = fileIterator.next();
          files.push({
            id: file.getId(),
            name: file.getName(),
            url: file.getUrl(),
            downloadUrl: file.getDownloadUrl(),
            mimeType: file.getMimeType(),
            size: file.getSize(),
            dateCreated: file.getDateCreated(),
            lastUpdated: file.getLastUpdated()
          });
        }
        
        return files;
      } catch (error) {
        logError(`Failed to get folder files (${folderId}): ${error.message}`);
        return [];
      }
    }
    
    /**
     * Sets sharing permissions for a file or folder
     * 
     * @param {string} fileOrFolderId - ID of the file or folder
     * @param {boolean} anyoneWithLink - Whether to allow anyone with link to access
     * @param {string} access - Access level: 'VIEW', 'COMMENT', or 'EDIT'
     * @return {boolean} True if successful
     */
    setSharing(fileOrFolderId, anyoneWithLink = true, access = 'VIEW') {
      try {
        let item;
        
        // Try to get as file first
        try {
          item = DriveApp.getFileById(fileOrFolderId);
        } catch (fileError) {
          // If not a file, try as folder
          try {
            item = DriveApp.getFolderById(fileOrFolderId);
          } catch (folderError) {
            logError(`Item not found for sharing: ${fileOrFolderId}`);
            return false;
          }
        }
        
        if (anyoneWithLink) {
          // Set sharing for anyone with link
          switch (access) {
            case 'VIEW':
              item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              break;
            case 'COMMENT':
              item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
              break;
            case 'EDIT':
              item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
              break;
            default:
              item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          }
        } else {
          // Private - only specific people can access
          item.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
        }
        
        logInfo(`Set sharing for item: ${fileOrFolderId} (anyoneWithLink: ${anyoneWithLink}, access: ${access})`);
        return true;
      } catch (error) {
        logError(`Failed to set sharing (${fileOrFolderId}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Adds a specific user as an editor to a file or folder
     * 
     * @param {string} fileOrFolderId - ID of the file or folder
     * @param {string} email - Email address of the user to add
     * @return {boolean} True if successful
     */
    addEditor(fileOrFolderId, email) {
      try {
        if (!isValidEmail(email)) {
          logError(`Invalid email address: ${email}`);
          return false;
        }
        
        let item;
        
        // Try to get as file first
        try {
          item = DriveApp.getFileById(fileOrFolderId);
        } catch (fileError) {
          // If not a file, try as folder
          try {
            item = DriveApp.getFolderById(fileOrFolderId);
          } catch (folderError) {
            logError(`Item not found for adding editor: ${fileOrFolderId}`);
            return false;
          }
        }
        
        item.addEditor(email);
        logInfo(`Added editor ${email} to item: ${fileOrFolderId}`);
        return true;
      } catch (error) {
        logError(`Failed to add editor (${email} to ${fileOrFolderId}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Adds a specific user as a viewer to a file or folder
     * 
     * @param {string} fileOrFolderId - ID of the file or folder
     * @param {string} email - Email address of the user to add
     * @return {boolean} True if successful
     */
    addViewer(fileOrFolderId, email) {
      try {
        if (!isValidEmail(email)) {
          logError(`Invalid email address: ${email}`);
          return false;
        }
        
        let item;
        
        // Try to get as file first
        try {
          item = DriveApp.getFileById(fileOrFolderId);
        } catch (fileError) {
          // If not a file, try as folder
          try {
            item = DriveApp.getFolderById(fileOrFolderId);
          } catch (folderError) {
            logError(`Item not found for adding viewer: ${fileOrFolderId}`);
            return false;
          }
        }
        
        item.addViewer(email);
        logInfo(`Added viewer ${email} to item: ${fileOrFolderId}`);
        return true;
      } catch (error) {
        logError(`Failed to add viewer (${email} to ${fileOrFolderId}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Lists all project folders
     * 
     * @return {Array<Object>} Array of project folder objects
     */
    listProjectFolders() {
      try {
        const rootFolder = this.getRootFolder();
        const projectFolders = [];
        
        const folderIterator = rootFolder.getFolders();
        
        while (folderIterator.hasNext()) {
          const folder = folderIterator.next();
          const folderName = folder.getName();
          
          // Extract project ID from folder name (if it follows the pattern)
          const projectIdMatch = folderName.match(/\(([\w-]+)\)$/);
          const projectId = projectIdMatch ? projectIdMatch[1] : null;
          
          projectFolders.push({
            id: folder.getId(),
            name: folderName,
            projectId: projectId,
            dateCreated: folder.getDateCreated(),
            lastUpdated: folder.getLastUpdated(),
            url: folder.getUrl()
          });
        }
        
        return projectFolders;
      } catch (error) {
        logError(`Failed to list project folders: ${error.message}`);
        return [];
      }
    }
    
    /**
     * Gets a project folder by project ID
     * 
     * @param {string} projectId - ID of the project
     * @return {Object} Project folder object, or null if not found
     */
    getProjectFolder(projectId) {
      try {
        const rootFolder = this.getRootFolder();
        const folderIterator = rootFolder.getFolders();
        
        while (folderIterator.hasNext()) {
          const folder = folderIterator.next();
          const folderName = folder.getName();
          
          // Check if folder name contains the project ID
          if (folderName.includes(`(${projectId})`)) {
            return {
              id: folder.getId(),
              name: folderName,
              projectId: projectId,
              dateCreated: folder.getDateCreated(),
              lastUpdated: folder.getLastUpdated(),
              url: folder.getUrl()
            };
          }
        }
        
        logWarning(`Project folder not found for project ID: ${projectId}`);
        return null;
      } catch (error) {
        logError(`Failed to get project folder (${projectId}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets the media folder for a project
     * 
     * @param {string} projectFolderId - ID of the project folder
     * @return {Object} Media folder object, or null if not found
     */
    getMediaFolder(projectFolderId) {
      try {
        const projectFolder = this.getFolder(projectFolderId);
        if (!projectFolder) {
          logError(`Project folder not found: ${projectFolderId}`);
          return null;
        }
        
        const mediaFolderName = DRIVE_SETTINGS.MEDIA_FOLDER_NAME;
        const folderIterator = projectFolder.getFolders();
        
        while (folderIterator.hasNext()) {
          const folder = folderIterator.next();
          if (folder.getName() === mediaFolderName) {
            return {
              id: folder.getId(),
              name: folder.getName(),
              dateCreated: folder.getDateCreated(),
              lastUpdated: folder.getLastUpdated(),
              url: folder.getUrl()
            };
          }
        }
        
        // If media folder doesn't exist, create it
        const mediaFolder = projectFolder.createFolder(mediaFolderName);
        
        return {
          id: mediaFolder.getId(),
          name: mediaFolder.getName(),
          dateCreated: mediaFolder.getDateCreated(),
          lastUpdated: mediaFolder.getLastUpdated(),
          url: mediaFolder.getUrl()
        };
      } catch (error) {
        logError(`Failed to get media folder (${projectFolderId}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Gets the thumbnails folder for a project
     * 
     * @param {string} projectFolderId - ID of the project folder
     * @return {Object} Thumbnails folder object, or null if not found
     */
    getThumbnailsFolder(projectFolderId) {
      try {
        const mediaFolder = this.getMediaFolder(projectFolderId);
        if (!mediaFolder) {
          logError(`Media folder not found for project: ${projectFolderId}`);
          return null;
        }
        
        const mediaFolderObj = this.getFolder(mediaFolder.id);
        if (!mediaFolderObj) {
          logError(`Failed to get media folder object: ${mediaFolder.id}`);
          return null;
        }
        
        const thumbnailFolderName = DRIVE_SETTINGS.THUMBNAIL_FOLDER_NAME;
        const folderIterator = mediaFolderObj.getFolders();
        
        while (folderIterator.hasNext()) {
          const folder = folderIterator.next();
          if (folder.getName() === thumbnailFolderName) {
            return {
              id: folder.getId(),
              name: folder.getName(),
              dateCreated: folder.getDateCreated(),
              lastUpdated: folder.getLastUpdated(),
              url: folder.getUrl()
            };
          }
        }
        
        // If thumbnails folder doesn't exist, create it
        const thumbnailsFolder = mediaFolderObj.createFolder(thumbnailFolderName);
        
        return {
          id: thumbnailsFolder.getId(),
          name: thumbnailsFolder.getName(),
          dateCreated: thumbnailsFolder.getDateCreated(),
          lastUpdated: thumbnailsFolder.getLastUpdated(),
          url: thumbnailsFolder.getUrl()
        };
      } catch (error) {
        logError(`Failed to get thumbnails folder (${projectFolderId}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Checks if a file exists by name in a folder
     * 
     * @param {string} folderId - ID of the folder to check
     * @param {string} fileName - Name of the file to find
     * @return {boolean} True if file exists, false if not
     */
    fileExistsByName(folderId, fileName) {
      try {
        const folder = this.getFolder(folderId);
        if (!folder) {
          logError(`Folder not found: ${folderId}`);
          return false;
        }
        
        const fileIterator = folder.getFilesByName(fileName);
        return fileIterator.hasNext();
      } catch (error) {
        logError(`Failed to check if file exists (${fileName}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Gets a file by name from a folder
     * 
     * @param {string} folderId - ID of the folder
     * @param {string} fileName - Name of the file to find
     * @return {Object} File object, or null if not found
     */
    getFileByName(folderId, fileName) {
      try {
        const folder = this.getFolder(folderId);
        if (!folder) {
          logError(`Folder not found: ${folderId}`);
          return null;
        }
        
        const fileIterator = folder.getFilesByName(fileName);
        
        if (fileIterator.hasNext()) {
          const file = fileIterator.next();
          return {
            id: file.getId(),
            name: file.getName(),
            mimeType: file.getMimeType(),
            size: file.getSize(),
            dateCreated: file.getDateCreated(),
            lastUpdated: file.getLastUpdated(),
            url: file.getUrl(),
            downloadUrl: file.getDownloadUrl()
          };
        }
        
        logWarning(`File not found by name: ${fileName} in folder ${folderId}`);
        return null;
      } catch (error) {
        logError(`Failed to get file by name (${fileName}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Renames a file
     * 
     * @param {string} fileId - ID of the file to rename
     * @param {string} newName - New name for the file
     * @return {boolean} True if successful
     */
    renameFile(fileId, newName) {
      try {
        const file = this.getFile(fileId);
        if (!file) {
          logError(`File not found for renaming: ${fileId}`);
          return false;
        }
        
        file.setName(newName);
        logInfo(`Renamed file ${fileId} to ${newName}`);
        return true;
      } catch (error) {
        logError(`Failed to rename file (${fileId}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Moves a file to a different folder
     * 
     * @param {string} fileId - ID of the file to move
     * @param {string} targetFolderId - ID of the destination folder
     * @param {boolean} removeFromSource - Whether to remove the file from its source folder
     * @return {boolean} True if successful
     */
    moveFile(fileId, targetFolderId, removeFromSource = true) {
      try {
        const file = this.getFile(fileId);
        if (!file) {
          logError(`File not found for moving: ${fileId}`);
          return false;
        }
        
        const targetFolder = this.getFolder(targetFolderId);
        if (!targetFolder) {
          logError(`Target folder not found: ${targetFolderId}`);
          return false;
        }
        
        // Add file to target folder
        targetFolder.addFile(file);
        
        if (removeFromSource) {
          // Get parent folders and remove file from them
          const parents = file.getParents();
          while (parents.hasNext()) {
            const parent = parents.next();
            if (parent.getId() !== targetFolderId) {
              parent.removeFile(file);
            }
          }
        }
        
        logInfo(`Moved file ${fileId} to folder ${targetFolderId}`);
        return true;
      } catch (error) {
        logError(`Failed to move file (${fileId}): ${error.message}`);
        return false;
      }
    }
    
    /**
     * Creates a copy of a file in the same or different folder
     * 
     * @param {string} fileId - ID of the file to copy
     * @param {string} newName - Name for the copy
     * @param {string} targetFolderId - ID of the destination folder (optional)
     * @return {Object} New file object, or null if failed
     */
    copyFile(fileId, newName, targetFolderId = null) {
      try {
        const file = this.getFile(fileId);
        if (!file) {
          logError(`File not found for copying: ${fileId}`);
          return null;
        }
        
        // Create a copy
        const copy = file.makeCopy(newName);
        
        // Move to target folder if specified
        if (targetFolderId) {
          const targetFolder = this.getFolder(targetFolderId);
          if (!targetFolder) {
            logError(`Target folder not found: ${targetFolderId}`);
            return null;
          }
          
          targetFolder.addFile(copy);
          
          // Remove from original folder
          const parents = copy.getParents();
          while (parents.hasNext()) {
            const parent = parents.next();
            if (parent.getId() !== targetFolderId) {
              parent.removeFile(copy);
            }
          }
        }
        
        logInfo(`Created copy of file ${fileId} named ${newName}`);
        
        return {
          id: copy.getId(),
          name: copy.getName(),
          mimeType: copy.getMimeType(),
          size: copy.getSize(),
          dateCreated: copy.getDateCreated(),
          lastUpdated: copy.getLastUpdated(),
          url: copy.getUrl(),
          downloadUrl: copy.getDownloadUrl()
        };
      } catch (error) {
        logError(`Failed to copy file (${fileId}): ${error.message}`);
        return null;
      }
    }
    
    /**
     * Clean up unused files in a project folder
     * 
     * @param {string} projectFolderId - ID of the project folder
     * @param {Array<string>} usedFileIds - Array of file IDs that are in use
     * @param {boolean} moveToTrash - Whether to move unused files to trash or just log them
     * @return {Object} Result object with counts of processed and removed files
     */
    cleanupUnusedFiles(projectFolderId, usedFileIds = [], moveToTrash = false) {
      try {
        const projectFolder = this.getFolder(projectFolderId);
        if (!projectFolder) {
          logError(`Project folder not found: ${projectFolderId}`);
          return { processed: 0, removed: 0 };
        }
        
        // Get all files in the project recursively
        const allFiles = this.getAllFilesInFolder(projectFolderId);
        let removedCount = 0;
        
        // Create a set of used file IDs for faster lookup
        const usedFileIdSet = new Set(usedFileIds);
        
        // Process each file
        for (const file of allFiles) {
          if (!usedFileIdSet.has(file.id)) {
            // File is not in use
            if (moveToTrash) {
              const fileObj = this.getFile(file.id);
              if (fileObj) {
                fileObj.setTrashed(true);
                removedCount++;
                logInfo(`Moved unused file to trash: ${file.name} (${file.id})`);
              }
            } else {
              logInfo(`Found unused file: ${file.name} (${file.id})`);
            }
          }
        }
        
        return {
          processed: allFiles.length,
          removed: removedCount
        };
      } catch (error) {
        logError(`Failed to cleanup unused files (${projectFolderId}): ${error.message}`);
        return { processed: 0, removed: 0 };
      }
    }
    
    /**
     * Gets all files in a folder recursively
     * 
     * @param {string} folderId - ID of the folder
     * @return {Array<Object>} Array of file objects
     */
    getAllFilesInFolder(folderId) {
      try {
        const folder = this.getFolder(folderId);
        if (!folder) {
          logError(`Folder not found: ${folderId}`);
          return [];
        }
        
        let allFiles = [];
        
        // Get files directly in this folder
        const fileIterator = folder.getFiles();
        while (fileIterator.hasNext()) {
          const file = fileIterator.next();
          allFiles.push({
            id: file.getId(),
            name: file.getName(),
            mimeType: file.getMimeType(),
            size: file.getSize(),
            dateCreated: file.getDateCreated(),
            lastUpdated: file.getLastUpdated(),
            url: file.getUrl()
          });
        }
        
        // Process subfolders recursively
        const folderIterator = folder.getFolders();
        while (folderIterator.hasNext()) {
          const subFolder = folderIterator.next();
          const subFolderFiles = this.getAllFilesInFolder(subFolder.getId());
          allFiles = allFiles.concat(subFolderFiles);
        }
        
        return allFiles;
      } catch (error) {
        logError(`Failed to get all files in folder (${folderId}): ${error.message}`);
        return [];
      }
    }
    
    /**
     * Validates that a file has an allowed MIME type
     * 
     * @param {Blob|Object} file - File blob or object with mimeType property
     * @param {string} type - Type category ('IMAGE', 'AUDIO', etc.)
     * @return {boolean} True if file has an allowed MIME type
     */
    isAllowedFileType(file, type) {
      try {
        const mimeType = file.mimeType || file.getContentType();
        
        if (!mimeType) {
          logError('No MIME type found for file');
          return false;
        }
        
        if (!DRIVE_SETTINGS.ALLOWED_MIME_TYPES[type]) {
          logError(`No allowed MIME types defined for type: ${type}`);
          return false;
        }
        
        return DRIVE_SETTINGS.ALLOWED_MIME_TYPES[type].includes(mimeType);
      } catch (error) {
        logError(`Failed to validate file type: ${error.message}`);
        return false;
      }
    }
  }