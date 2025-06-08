/**
 * Gets the contents (files and subfolders) of a given folder ID.
 * This version uses caching to improve performance on repeated requests.
 * @param {string} folderId The ID of the folder to list.
 * @return {Object} An object containing lists of files and folders.
 */
function getFolderContents(folderId) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `contents_${folderId}`;
    const cachedContents = cache.get(cacheKey);

    if (cachedContents) {
        Logger.log(`Cache hit for folderId: ${folderId}`);
        return JSON.parse(cachedContents);
    }
    Logger.log(`Cache miss for folderId: ${folderId}. Fetching from Drive.`);

    const folder = DriveApp.getFolderById(folderId);
    const folders = [];
    const files = [];

    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const sub = subfolders.next();
      folders.push({ id: sub.getId(), name: sub.getName() });
    }

    const fileIterator = folder.getFiles();
    const supportedMimeTypes = [MimeType.GOOGLE_DOCS, MimeType.JPEG, MimeType.PNG, MimeType.GIF];
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      if (supportedMimeTypes.includes(file.getMimeType())) {
        files.push({ id: file.getId(), name: file.getName(), mimeType: file.getMimeType() });
      }
    }

    // Sort alphabetically
    folders.sort((a, b) => a.name.localeCompare(b.name));
    files.sort((a, b) => a.name.localeCompare(b.name));

    const contents = { folders, files };
    cache.put(cacheKey, JSON.stringify(contents), 300);

    return contents;
}

/**
 * Gets the entire info data object for a given folder.
 * @param {string} folderId The ID of the folder.
 * @return {Object} The parsed JSON data from the info file, or an empty object.
 */
function getInfoData(folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const fileName = ".nik-template-info.json";
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
        try {
            const content = files.next().getBlob().getDataAsString();
            return JSON.parse(content);
        } catch (e) {
            console.error(`Error parsing info file in folder ${folderId}: ${e.message}`);
            return {}; // Return empty object on parsing error
        }
    }
    return {}; // Return empty object if file doesn't exist
}

/**
 * Gets the description for a single file from the folder's info file.
 * @param {string} fileId The ID of the file.
 * @param {string} parentFolderId The ID of the folder containing the file.
 * @return {string} The stored description or an empty string.
 */
function getFileDescription(fileId, parentFolderId) {
  const infoData = getInfoData(parentFolderId);
  return infoData[fileId] || "";
}

/**
 * Removes entries from the info file for templates that no longer exist.
 * @param {string} folderId The ID of the folder to clean up.
 */
function pruneInfoFile(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const fileName = ".nik-template-info.json";
  const infoFiles = folder.getFilesByName(fileName);

  if (!infoFiles.hasNext()) {
      Logger.log(`No info file to prune in folder ${folderId}`);
      return; // Nothing to do
  }
  
  const infoFile = infoFiles.next();
  let infoData;
  try {
      infoData = JSON.parse(infoFile.getBlob().getDataAsString());
  } catch(e) {
      console.error(`Could not parse info file for pruning: ${e.message}`);
      return; // Cannot prune a broken file
  }

  const currentFileIds = new Set();
  const filesInFolder = folder.getFiles();
  while(filesInFolder.hasNext()){
      currentFileIds.add(filesInFolder.next().getId());
  }
  
  let wasModified = false;
  for (const idInInfoFile in infoData) {
      if (!currentFileIds.has(idInInfoFile)) {
          console.log(`Pruning stale description for deleted file ID: ${idInInfoFile}`);
          delete infoData[idInInfoFile];
          wasModified = true;
      }
  }

  // Only rewrite the file if changes were made
  if (wasModified) {
      console.log(`Saving pruned info file in folder ${folderId}`);
      const content = JSON.stringify(infoData, null, 2);
      infoFile.setContent(content);
  }
}

/**
 * Saves the description for a single file to the folder's info file.
 * @param {string} fileId The ID of the file.
 * @param {string} parentFolderId The ID of the folder containing the file.
 * @param {string} description The new description text.
 */
function saveFileDescription(fileId, parentFolderId, description) {
  const folder = DriveApp.getFolderById(parentFolderId);
  const fileName = ".nik-template-info.json";
  let infoData = getInfoData(parentFolderId);

  // Update or add the new description
  infoData[fileId] = description;

  // Save the updated data back to the file
  const content = JSON.stringify(infoData, null, 2);
  const infoFiles = folder.getFilesByName(fileName);
  if (infoFiles.hasNext()) {
    infoFiles.next().setContent(content);
  } else {
    folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  }

  // After saving, also prune any other stale entries
  pruneInfoFile(parentFolderId);
}

/**
 * Creates a new Google Doc in the specified folder.
 * @param {string} folderId The ID of the target folder.
 * @return {Document} The newly created Google Doc object.
 */
function createNewTemplateInFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const docName = `Untitled Template - ${new Date().toLocaleString()}`;
  const newDoc = DocumentApp.create(docName);
  const newDocFile = DriveApp.getFileById(newDoc.getId());
  
  newDocFile.moveTo(folder);
  
  Logger.log(`Created new doc "${docName}" (${newDoc.getId()}) in folder "${folder.getName()}" (${folderId})`);
  return newDoc;
}

/**
 * Inserts content from a Google Doc or an image into the active document.
 * @param {string} fileId The ID of the file to insert.
 * @param {string} mimeType The MIME type of the file.
 */
function insertContent(fileId, mimeType) {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error('Cannot insert content. Please place your cursor in the document.');
  }

  if (mimeType.includes('image')) {
    const imageBlob = DriveApp.getFileById(fileId).getBlob();
    cursor.insertImage(imageBlob);
    return;
  }

  if (mimeType === MimeType.GOOGLE_DOCS) {
    const sourceBody = DocumentApp.openById(fileId).getBody();
    let insertionPoint = cursor.getElement();
    let container;

    if (insertionPoint) {
      container = insertionPoint.getParent();
      while (container && container.getParent() && container.getType() !== DocumentApp.ElementType.BODY_SECTION) {
        insertionPoint = container;
        container = insertionPoint.getParent();
      }
    }
    
    if (!container || typeof container.insertParagraph !== 'function') {
        container = doc.getBody();
        insertionPoint = null;
    }
    
    const insertionIndex = insertionPoint ? container.getChildIndex(insertionPoint) + 1 : container.getNumChildren();

    for (let i = sourceBody.getNumChildren() - 1; i >= 0; i--) {
        const originalElement = sourceBody.getChild(i);
        const type = originalElement.getType();
        
        if (type === DocumentApp.ElementType.PARAGRAPH) {
            const sourcePara = originalElement.asParagraph();
            const targetPara = container.insertParagraph(insertionIndex, "");
            
            targetPara.setAttributes(sourcePara.getAttributes());
            
            for (let j = 0; j < sourcePara.getNumChildren(); j++) {
                const child = sourcePara.getChild(j);
                const childType = child.getType();

                if (childType === DocumentApp.ElementType.TEXT) {
                    targetPara.append(child.copy());
                } else if (childType === DocumentApp.ElementType.INLINE_IMAGE) {
                    targetPara.append(child.copy());
                }
            }

        } else if (type === DocumentApp.ElementType.TABLE) {
            container.insertTable(insertionIndex, originalElement.copy().asTable());
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
            container.insertListItem(insertionIndex, originalElement.copy().asListItem());
        } else {
            const elementCopy = originalElement.copy();
            if (elementCopy) {
                if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
                    container.insertHorizontalRule(insertionIndex);
                }
            } else {
                console.warn(`Skipping an unsupported element type: ${type}`);
            }
        }
    }
  }
}
