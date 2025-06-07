/**
 * @OnlyCurrentDoc
 */

// --- Main Add-on Entry Point ---

/**
 * Runs when the add-on is opened from the sidebar icon.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onHomepage(e) {
  const userProps = PropertiesService.getUserProperties();
  const folderId = userProps.getProperty('rootFolderId');

  if (folderId) {
    try {
      DriveApp.getFolderById(folderId); // Validate that the folder is still accessible.
      return createBrowserCard(folderId, [], {}); // Start with no file selected
    } catch (error) {
      userProps.deleteProperty('rootFolderId');
      return createErrorCard(error, 'The previously saved folder could not be accessed.');
    }
  }
  return createConfigurationCard();
}

/**
 * Callback for the Docs-specific homepage trigger.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onDocsHomepage(e) {
    return onHomepage(e);
}


// --- Card Creation Functions ---

/**
 * Creates the initial configuration card prompting the user to set a folder.
 * @return {CardService.Card}
 */
function createConfigurationCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Setup Required'));
  const section = CardService.newCardSection();

  section.addWidget(CardService.newDecoratedText()
    .setText('To get started, please set a root folder for the content inserter.')
    .setWrapText(true));

  section.addWidget(CardService.newTextInput()
    .setFieldName('folderUrl')
    .setTitle('Google Drive Folder URL or ID')
    .setHint('Paste the folder URL or just the ID here'));

  const saveAction = CardService.newAction().setFunctionName('handleSaveFolderAction');

  section.addWidget(CardService.newTextButton()
    .setText('Save and Start Browsing')
    .setOnClickAction(saveAction));

  card.addSection(section);
  return card.build();
}

/**
 * Creates the main card for browsing files and folders with the new selection model.
 * @param {string} folderId - The ID of the folder to display.
 * @param {Array<Object>} path - An array of {id, name} objects representing the navigation path.
 * @param {Object} selection - An object like {id, name, mimeType} representing the selected file.
 * @return {CardService.Card} The browser card.
 */
function createBrowserCard(folderId, path, selection) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Drive Content Inserter'));

  const resetAction = CardService.newAction().setFunctionName('handleResetAction');
  card.addCardAction(CardService.newCardAction()
    .setText('Reset Root Folder')
    .setOnClickAction(resetAction));

  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderContents = getFolderContents(folderId);
    
    // --- Breadcrumb Section ---
    const breadcrumbSection = CardService.newCardSection().setHeader("Path");
    const breadcrumbSet = CardService.newButtonSet();
    const rootId = PropertiesService.getUserProperties().getProperty('rootFolderId');

    if (rootId) {
        const rootFolder = DriveApp.getFolderById(rootId);
        breadcrumbSet.addButton(CardService.newTextButton()
          .setText(rootFolder.getName())
          .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
          .setOnClickAction(CardService.newAction().setFunctionName('handleNavigation').setParameters({ folderId: rootId, path: '[]', selection: '{}' })));

        path.forEach((p, index) => {
          breadcrumbSet.addButton(CardService.newTextButton()
            .setText(p.name)
            .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
            .setOnClickAction(CardService.newAction().setFunctionName('handleNavigation').setParameters({ folderId: p.id, path: JSON.stringify(path.slice(0, index)), selection: '{}' })));
        });
        breadcrumbSection.addWidget(breadcrumbSet);
    }
    breadcrumbSection.addWidget(CardService.newDecoratedText().setText(folder.getName()).setTopLabel("Current Folder"));
    card.addSection(breadcrumbSection);


    // --- Content Section ---
    const contentSection = CardService.newCardSection().setHeader("Contents");
    let hasContent = false;

    // Subfolders
    folderContents.folders.forEach(subfolder => {
        hasContent = true;
        const newPathForSubfolder = [...path, { id: folderId, name: folder.getName() }];
        contentSection.addWidget(CardService.newDecoratedText()
            .setText(`📁 ${subfolder.name}`) // Add folder emoji
            .setWrapText(true)
            .setOnClickAction(CardService.newAction().setFunctionName('handleNavigation').setParameters({ folderId: subfolder.id, path: JSON.stringify(newPathForSubfolder), selection: '{}' })));
    });
    
    // Files
    folderContents.files.forEach(file => {
      hasContent = true;
      const isSelected = selection && selection.id === file.id;

      const fileWidget = createFileWidget(file, folderId, path, selection);
      contentSection.addWidget(fileWidget);

      // If this file is the selected one, show the action buttons and description underneath it.
      if (isSelected) {
        const actionWidget = createActionWidget(selection, folderId, path);
        contentSection.addWidget(actionWidget);
        
        let description = getFileDescription(selection.id, folderId);
        if (description && description.trim() !== "No description provided." && description.trim() !== "") {
            if (description.length > 255) {
                description = description.substring(0, 252) + '...';
            }
            const formattedDescription = `<i><font color="#888888">${description}</font></i>`;
            contentSection.addWidget(CardService.newTextParagraph().setText(formattedDescription));
        }
      }
    });

    if (!hasContent) {
      contentSection.addWidget(CardService.newTextParagraph().setText("This folder is empty."));
    }
    card.addSection(contentSection);

  } catch (err) {
    return createErrorCard(err, `Error building browser for folderId: ${folderId}`);
  }
  return card.build();
}

/**
 * Creates the action button widget for the selected file.
 * @param {Object} selection - The selected file object.
 * @param {string} parentFolderId - The ID of the current folder.
 * @param {Array<Object>} path - The current breadcrumb path.
 * @return {CardService.ButtonSet}
 */
function createActionWidget(selection, parentFolderId, path) {
    const insertAction = CardService.newAction().setFunctionName('handleInsertAction')
        .setParameters({ fileId: selection.id, mimeType: selection.mimeType });
    const insertButton = CardService.newTextButton()
        .setText("⬇️ Insert")
        .setOnClickAction(insertAction);

    const infoAction = CardService.newAction().setFunctionName('handleInfoAction')
        .setParameters({ fileId: selection.id, parentFolderId: parentFolderId, path: JSON.stringify(path), selection: JSON.stringify(selection) });
    const infoButton = CardService.newTextButton()
        .setText("💬 Info")
        .setOnClickAction(infoAction);

    return CardService.newButtonSet().addButton(insertButton).addButton(infoButton);
}

/**
 * Creates a simple widget for a single file row that handles selection.
 * @param {Object} file - A file object containing id, name, and mimeType.
 * @param {string} parentFolderId - The ID of the current folder.
 * @param {Array<Object>} path - The current breadcrumb path.
 * @param {Object} selection - The currently selected file object.
 * @return {CardService.DecoratedText|null} The widget for the file row.
 */
function createFileWidget(file, parentFolderId, path, selection) {
  const isSelected = selection && selection.id === file.id;
  const newSelection = isSelected ? {} : { id: file.id, name: file.name, mimeType: file.mimeType };
  
  const selectionAction = CardService.newAction().setFunctionName('handleFileSelection')
    .setParameters({
        folderId: parentFolderId,
        path: JSON.stringify(path),
        selection: JSON.stringify(newSelection)
    });

  let icon = '📄'; // Default file icon
  if (file.mimeType.includes('image')) {
    icon = '🌅';
  }
  
  let fileNameText = file.name;

  if (isSelected) {
    icon = '✅';
    fileNameText = `<b>${file.name}</b>`; // Bold the text when selected
  }
  
  const textWidget = CardService.newDecoratedText()
    .setText(`${icon} ${fileNameText}`)
    .setWrapText(true)
    .setOnClickAction(selectionAction);
  
  return textWidget;
}

/**
 * Creates a card that displays a detailed error message.
 * @param {Error} error The error object.
 * @param {string} contextMessage A user-friendly message about what failed.
 * @return {CardService.Card} The error card.
 */
function createErrorCard(error, contextMessage) {
  console.error(`${contextMessage}: ${error.message}`, error.stack);
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('An Error Occurred').setSubtitle(contextMessage));
  const section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(`<b>Details:</b> ${error.message}`));
  const resetAction = CardService.newAction().setFunctionName('handleResetAction');
  section.addWidget(CardService.newTextButton().setText('Reset and Start Over').setOnClickAction(resetAction));
  card.addSection(section);
  return card.build();
}


// --- Action Handlers ---

function handleSaveFolderAction(e) {
  const folderUrl = e.formInput.folderUrl.trim();
  const folderId = folderUrl.includes('folders/') ? folderUrl.substring(folderUrl.lastIndexOf('folders/') + 8).split('/')[0] : folderUrl;
  
  try {
    DriveApp.getFolderById(folderId);
    PropertiesService.getUserProperties().setProperty('rootFolderId', folderId);
    const newCard = createBrowserCard(folderId, [], {});
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(newCard)).build();
  } catch (err) {
    const errorCard = createErrorCard(err, `Could not access folder with ID "${folderId}".`);
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(errorCard)).build();
  }
}

function handleResetAction() {
  PropertiesService.getUserProperties().deleteProperty('rootFolderId');
  const card = createConfigurationCard();
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(card)).build();
}

function handleNavigation(e) {
  const folderId = e.parameters.folderId;
  const path = JSON.parse(e.parameters.path);
  const selection = JSON.parse(e.parameters.selection);
  const newCard = createBrowserCard(folderId, path, selection);
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(newCard)).build();
}

function handleFileSelection(e) {
    const folderId = e.parameters.folderId;
    const path = JSON.parse(e.parameters.path);
    const selection = JSON.parse(e.parameters.selection);
    const newCard = createBrowserCard(folderId, path, selection);
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(newCard)).build();
}

function handleInsertAction(e) {
  try {
    insertContent(e.parameters.fileId, e.parameters.mimeType);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Content inserted.")).build();
  } catch (err) {
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText(`Insertion Failed: ${err.message}`)).build();
  }
}

function handleInfoAction(e) {
  const params = e.parameters;
  const file = DriveApp.getFileById(params.fileId);
  const description = getFileDescription(params.fileId, params.parentFolderId);
  
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(`Description for ${file.getName()}`))
    .addSection(CardService.newCardSection().addWidget(
      CardService.newTextInput()
        .setFieldName('description_text')
        .setTitle('Edit Description')
        .setValue(description)
        .setMultiline(true)))
    .setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Save')
        .setOnClickAction(CardService.newAction().setFunctionName('handleSaveDescriptionAction').setParameters(params))))
    .build();

  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
}

function handleSaveDescriptionAction(e) {
  const params = e.parameters;
  saveFileDescription(params.fileId, params.parentFolderId, e.formInput.description_text);
  
  // Rebuild the card to return to the same state.
  const path = JSON.parse(params.path);
  const selection = JSON.parse(params.selection);
  const card = createBrowserCard(params.parentFolderId, path, selection);
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Description saved.'))
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}


// --- Backend Logic Functions ---

function getFolderContents(folderId) {
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

    return { folders, files };
}

function getFileDescription(fileId, parentFolderId) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const fileName = `.${fileId}.description.json`;
  const files = parentFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    try {
      return JSON.parse(files.next().getBlob().getDataAsString()).description;
    } catch (e) { console.error(`Error parsing description for ${fileId}: ${e.message}`); }
  }
  return "";
}

function saveFileDescription(fileId, parentFolderId, description) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const fileName = `.${fileId}.description.json`;
  const files = parentFolder.getFilesByName(fileName);
  const data = { description: description, lastUpdated: new Date().toISOString() };
  const content = JSON.stringify(data, null, 2);

  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    parentFolder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  }
}

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
      while (container.getParent() &&
             container.getType() !== DocumentApp.ElementType.BODY_SECTION &&
             container.getType() !== DocumentApp.ElementType.HEADER_SECTION &&
             container.getType() !== DocumentApp.ElementType.FOOTER_SECTION &&
             container.getType() !== DocumentApp.ElementType.TABLE_CELL) {
        insertionPoint = container;
        container = insertionPoint.getParent();
      }
    } else {
      container = doc.getBody();
      insertionPoint = null;
    }
    
    const insertionIndex = insertionPoint ? container.getChildIndex(insertionPoint) + 1 : container.getNumChildren();

    for (let i = sourceBody.getNumChildren() - 1; i >= 0; i--) {
      const originalElement = sourceBody.getChild(i);
      const elementToCopy = originalElement.copy();
      
      if (elementToCopy) {
        const type = elementToCopy.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH) {
          container.insertParagraph(insertionIndex, elementToCopy.asParagraph());
        } else if (type === DocumentApp.ElementType.TABLE) {
          container.insertTable(insertionIndex, elementToCopy.asTable());
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          container.insertListItem(insertionIndex, elementToCopy.asListItem());
        } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
          container.insertHorizontalRule(insertionIndex);
        }
      } else {
        console.warn(`Skipping an unsupported element type: ${originalElement.getType()} at source index ${i}.`);
      }
    }
  }
}
