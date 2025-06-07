/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file
 * access for this add-on to only the current document.
 */

/**
 * Callback for rendering the add-on's homepage. This is the main entry point.
 * It checks for configuration and routes the user to the correct card.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onHomepage(e) {
  console.log('onHomepage trigger fired.');
  const userProps = PropertiesService.getUserProperties();
  const folderId = userProps.getProperty('rootFolderId');
  console.log(`Retrieved rootFolderId: ${folderId}`);

  if (folderId) {
    try {
      // Verify the saved folder exists and is accessible before proceeding.
      DriveApp.getFolderById(folderId);
      console.log(`Successfully accessed folderId: ${folderId}. Building browser card.`);
      return createBrowserCard(folderId, []);
    } catch (error) {
      // The saved folder ID is no longer valid or accessible.
      console.error(`Failed to access saved folderId: ${folderId}. Error: ${error.message}`, error.stack);
      userProps.deleteProperty('rootFolderId');
      return createErrorCard(error); // Display a detailed error card.
    }
  }
  // No folder has been configured yet. Show the initial setup card.
  console.log('No rootFolderId found. Building configuration card.');
  return createConfigurationCard(false);
}

/**
 * Callback for the Docs-specific homepage trigger. This is good practice
 * to have when the manifest specifies it.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onDocsHomepage(e) {
    console.log('onDocsHomepage trigger fired.');
    return onHomepage(e);
}

/**
 * Creates a card that displays a detailed error message for debugging.
 * @param {Error} error The error object caught by a try-catch block.
 * @return {CardService.Card} The error card.
 */
function createErrorCard(error) {
  console.log(`Creating error card for error: ${error.message}`);
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('An Unexpected Error Occurred')
    .setSubtitle('Please share this information for troubleshooting.')
    .setImageUrl('https://ssl.gstatic.com/docs/script/images/icons/error_red_64dp.png'));

  const errorSection = CardService.newCardSection()
    .setHeader('Error Details');

  // Add the main error message
  errorSection.addWidget(CardService.newDecoratedText()
    .setTopLabel('Message')
    .setText(error.message)
    .setWrapText(true));

  // Add the stack trace for more context
  if (error.stack) {
    errorSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('Stack Trace')
      .setText(error.stack)
      .setWrapText(true));
  }

  // Add a button to reset the configuration and try again.
  const resetAction = CardService.newAction().setFunctionName('handleResetAction');
  errorSection.addWidget(CardService.newTextButton()
    .setText('Reset and Start Over')
    .setOnClickAction(resetAction));

  card.addSection(errorSection);
  return card.build();
}


/**
 * Creates a card that tells the user they need to authorize the add-on.
 * @return {CardService.Card} The authorization card.
 */
function createAuthorizationCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Authorization Required'));
  const section = CardService.newCardSection();
  section.addWidget(CardService.newDecoratedText()
    .setText('This add-on requires your permission to access Google Drive. Please reload the add-on and grant access when prompted. If the problem persists, you may need to re-install the add-on from the Google Workspace Marketplace.')
    .setWrapText(true)
    .setTopLabel("Permission Error"));
  card.addSection(section);
  return card.build();
}


/**
 * Creates a card that prompts the user to configure a root folder.
 * This is shown on first run or when the saved folder is invalid.
 * @param {boolean} showError - If true, displays a generic error message.
 * @return {CardService.Card} The configuration card.
 */
function createConfigurationCard(showError) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Setup Required'));
  const section = CardService.newCardSection();

  if (showError) {
    section.addWidget(CardService.newDecoratedText()
      .setText('The previously saved folder could not be accessed. It may have been deleted, or you may no longer have permission to view it. Please configure a new folder.')
      .setWrapText(true)
      .setTopLabel("Error"));
  } else {
    section.addWidget(CardService.newDecoratedText()
      .setText('To get started, please set a root folder for the content inserter. You can get the folder URL or ID from its page in Google Drive.')
      .setWrapText(true));
  }

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
 * Creates the main card for browsing files and folders.
 * @param {string} folderId - The ID of the folder to display.
 * @param {Array<Object>} path - An array of {id, name} objects representing the navigation path.
 * @return {CardService.Card} The browser card.
 */
function createBrowserCard(folderId, path) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Drive Content Inserter'));

  const resetAction = CardService.newAction().setFunctionName('handleResetAction');
  card.addCardAction(CardService.newCardAction()
    .setText('Reset Root Folder')
    .setOnClickAction(resetAction));

  try {
    const folder = DriveApp.getFolderById(folderId);

    // --- Breadcrumb Section ---
    const breadcrumbSection = CardService.newCardSection().setHeader("Path");
    const breadcrumbSet = CardService.newButtonSet();
    const rootId = PropertiesService.getUserProperties().getProperty('rootFolderId');

    if (rootId) {
        const rootFolder = DriveApp.getFolderById(rootId);
        // Root folder button
        breadcrumbSet.addButton(CardService.newTextButton()
          .setText(rootFolder.getName())
          .setTextButtonStyle(CardService.TextButtonStyle.TEXT) // Makes it look more like a link
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleNavigation')
            .setParameters({
              folderId: rootId,
              path: JSON.stringify([])
            })));

        // Path component buttons
        path.forEach((p, index) => {
          breadcrumbSet.addButton(CardService.newTextButton()
            .setText(p.name)
            .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
            .setOnClickAction(CardService.newAction()
              .setFunctionName('handleNavigation')
              .setParameters({
                folderId: p.id,
                path: JSON.stringify(path.slice(0, index + 1))
              })));
        });

        breadcrumbSection.addWidget(breadcrumbSet);
        // Display current folder name for context
        breadcrumbSection.addWidget(CardService.newDecoratedText().setText(folder.getName()).setTopLabel("Current Folder"));
        card.addSection(breadcrumbSection);
    }


    // --- Content Section ---
    const contentSection = CardService.newCardSection().setHeader("Contents");
    let hasContent = false;

    // --- List Subfolders ---
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      hasContent = true;
      const subfolder = subfolders.next();
      const newPathForSubfolder = [...path, {
        id: folder.getId(),
        name: folder.getName()
      }];

      const navAction = CardService.newAction()
        .setFunctionName('handleNavigation')
        .setParameters({
          folderId: subfolder.getId(),
          path: JSON.stringify(newPathForSubfolder)
        });

      contentSection.addWidget(CardService.newDecoratedText()
        .setText(subfolder.getName())
        .setWrapText(true)
        .setIcon(CardService.Icon.FOLDER)
        .setOnClickAction(navAction));
    }

    // --- List Files ---
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileWidget = createFileWidget(file, folderId);
      if (fileWidget) {
        hasContent = true;
        contentSection.addWidget(fileWidget);
      }
    }

    if (!hasContent) {
      contentSection.addWidget(CardService.newTextParagraph().setText("This folder is empty."));
    }
    card.addSection(contentSection);

  } catch (err) {
    console.error(`Error building browser card for folderId: ${folderId}. Error: ${err.message}`, err.stack);
    return createErrorCard(err);
  }

  return card.build();
}

/**
 * Creates a widget for a single file. Since DecoratedText only supports one
 * button, this widget only includes the "Insert" action.
 * @param {DriveApp.File} file - The file object.
 * @param {string} parentFolderId - The ID of the parent folder.
 * @return {CardService.DecoratedText|null} The widget for the file or null if unsupported.
 */
function createFileWidget(file, parentFolderId) {
  const fileId = file.getId();
  const fileName = file.getName();
  const mimeType = file.getMimeType();

  const supportedMimeTypes = [MimeType.GOOGLE_DOCS, MimeType.JPEG, MimeType.PNG, MimeType.GIF];
  if (!supportedMimeTypes.includes(mimeType)) return null;

  const actionParams = {
    fileId: fileId,
    mimeType: mimeType,
    parentFolderId: parentFolderId
  };
    
  // --- Create the interactive button ---
  const insertAction = CardService.newAction().setFunctionName('handleInsertAction').setParameters(actionParams);
  const insertButton = CardService.newTextButton()
      .setText('Insert')
      .setOnClickAction(insertAction);

  // --- Create the main text widget and add the button to it ---
  const decoratedText = CardService.newDecoratedText()
    .setText(fileName)
    .setWrapText(true)
    .setButton(insertButton);
  
  // Set icon based on MIME type
  if (mimeType === MimeType.GOOGLE_DOCS) {
    decoratedText.setIcon(CardService.Icon.DESCRIPTION);
  } else if (mimeType === MimeType.JPEG || mimeType === MimeType.PNG || mimeType === MimeType.GIF) {
    decoratedText.setIcon(CardService.Icon.IMAGE);
  }
  
  return decoratedText;
}

/**
 * Handles saving the root folder from the configuration card.
 * @param {Object} e The event object containing form input.
 * @return {CardService.ActionResponse} The response to navigate to the browser or show an error.
 */
function handleSaveFolderAction(e) {
  const folderUrl = e.formInput.folderUrl;
  console.log(`handleSaveFolderAction called with URL: ${folderUrl}`);

  if (!folderUrl) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Folder URL or ID cannot be empty.'))
      .build();
  }

  let folderId;
  const match = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
  folderId = (match && match[1]) ? match[1] : folderUrl;
  console.log(`Extracted folderId: ${folderId}`);

  try {
    const folder = DriveApp.getFolderById(folderId); // This validates the ID and permissions.
    console.log(`Successfully validated folder: ${folder.getName()}`);
    PropertiesService.getUserProperties().setProperty('rootFolderId', folderId);
    console.log(`Set rootFolderId property to: ${folderId}`);

    const browserCard = createBrowserCard(folderId, []);
    const navigation = CardService.newNavigation().updateCard(browserCard);

    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .setNotification(CardService.newNotification().setText(`Root folder set to "${folder.getName()}".`))
      .build();
  } catch (error) {
    console.error(`Failed to validate or save folderId: ${folderId}. Error: ${error.message}`, error.stack);
    // Instead of just a notification, show the full error card for immediate feedback.
    const errorCard = createErrorCard(error);
    const navigation = CardService.newNavigation().updateCard(errorCard);
    return CardService.newActionResponseBuilder().setNavigation(navigation).build();
  }
}

/**
 * Handles the "Reset" action, clearing the saved folder and showing the config card.
 * @param {Object} e The event object.
 * @return {CardService.ActionResponse} A response to show the config card.
 */
function handleResetAction(e) {
  console.log('handleResetAction called.');
  PropertiesService.getUserProperties().deleteProperty('rootFolderId');
  const configCard = createConfigurationCard(false);
  const navigation = CardService.newNavigation().updateCard(configCard);

  return CardService.newActionResponseBuilder()
    .setNavigation(navigation)
    .setNotification(CardService.newNotification().setText('Settings have been reset.'))
    .build();
}

/**
 * Handles navigation clicks from the breadcrumb or folder list.
 * @param {Object} e The event object from a user click.
 * @return {CardService.ActionResponse} A navigation object to update the card.
 */
function handleNavigation(e) {
  const params = e.parameters;
  const folderId = params.folderId;
  const path = params.path ? JSON.parse(params.path) : [];
  console.log(`Navigating to folderId: ${folderId}`);

  const newCard = createBrowserCard(folderId, path);
  const navigation = CardService.newNavigation().updateCard(newCard);
  return CardService.newActionResponseBuilder().setNavigation(navigation).build();
}

/**
 * Handles the click of an "Insert" button. This function has been rewritten
 * for robustness to handle different cursor positions and content types.
 * @param {Object} e The event object from a user click.
 * @return {CardService.ActionResponse} A response with a notification for the user.
 */
function handleInsertAction(e) {
  const params = e.parameters;
  const fileId = params.fileId;
  const mimeType = params.mimeType;
  console.log(`Attempting to insert fileId: ${fileId} (MIME: ${mimeType})`);
  
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    console.warn('Cannot insert content; cursor not found.');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Cannot insert content. Please place your cursor in the document.'))
      .build();
  }

  try {
    if (mimeType === MimeType.GOOGLE_DOCS) {
      const sourceBody = DocumentApp.openById(fileId).getBody();
      
      // Determine the insertion point and its container element in a robust way.
      let insertionPoint = cursor.getElement();
      let container;

      if (insertionPoint) {
        container = insertionPoint.getParent();
        // Traverse up to find a container that supports insertion methods.
        while (container.getParent() &&
               container.getType() !== DocumentApp.ElementType.BODY_SECTION &&
               container.getType() !== DocumentApp.ElementType.HEADER_SECTION &&
               container.getType() !== DocumentApp.ElementType.FOOTER_SECTION &&
               container.getType() !== DocumentApp.ElementType.TABLE_CELL
              ) {
          insertionPoint = container;
          container = insertionPoint.getParent();
        }
      } else {
        // Cursor is not in an element, so insert into the main body at the end.
        container = doc.getBody();
        insertionPoint = null; // No specific element to insert after.
      }
      
      // Calculate the insertion index based on the found element.
      const insertionIndex = insertionPoint ? container.getChildIndex(insertionPoint) + 1 : container.getNumChildren();

      // Loop backwards through the source elements to insert them in the correct order.
      for (let i = sourceBody.getNumChildren() - 1; i >= 0; i--) {
        const originalElement = sourceBody.getChild(i);
        const elementToCopy = originalElement.copy();
        
        // IMPORTANT: Check if the element was copied successfully.
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
            console.warn(`Skipping an unsupported element type: ${originalElement.getType().name()} at source index ${i}.`);
        }
      }

    } else if (mimeType === MimeType.JPEG || mimeType === MimeType.PNG || mimeType === MimeType.GIF) {
      // For simple image files, just insert the image at the cursor.
      const imageBlob = DriveApp.getFileById(fileId).getBlob();
      cursor.insertImage(imageBlob);
    }
    
    console.log('Content inserted successfully.');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Content inserted successfully.'))
      .build();
  } catch (error) {
    console.error(`Failed to insert fileId: ${fileId}. Error: ${error.message}`, error.stack);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Failed to insert file: ${error.message}`))
      .build();
  }
}

/**
 * Handles the click of an "Info" button. Shows a modal card with the file's description.
 * @param {Object} e The event object from a user click.
 * @return {CardService.ActionResponse} The response to open the info card.
 */
function handleInfoAction(e) {
  const params = e.parameters;
  const file = DriveApp.getFileById(params.fileId);
  const parentFolder = DriveApp.getFolderById(params.parentFolderId);
  const descriptionFileName = `.${params.fileId}.description.json`;
  const existingFiles = parentFolder.getFilesByName(descriptionFileName);
  let description = "No description provided.";

  if (existingFiles.hasNext()) {
    try {
      const data = JSON.parse(existingFiles.next().getBlob().getDataAsString());
      description = data.description;
    } catch (err) { /* Ignore parsing errors */ }
  }

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(`Description for ${file.getName()}`))
    .addSection(CardService.newCardSection().addWidget(
      CardService.newTextInput()
      .setFieldName('description_text')
      .setTitle('Edit Description')
      .setValue(description)
      .setMultiline(true)
    ))
    .setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
      .setText('Save')
      .setOnClickAction(CardService.newAction().setFunctionName('handleSaveDescriptionAction').setParameters(params))
    )).build();

  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
}

/**
 * Saves the description from the info modal.
 * @param {Object} e The event object from a user click.
 * @return {CardService.ActionResponse} A response that closes the modal.
 */
function handleSaveDescriptionAction(e) {
  const params = e.parameters;
  const newDescription = e.formInput.description_text;
  const descriptionFileName = `.${params.fileId}.description.json`;
  const parentFolder = DriveApp.getFolderById(params.parentFolderId);
  const existingFiles = parentFolder.getFilesByName(descriptionFileName);

  const descriptionData = {
    description: newDescription,
    lastUpdated: new Date().toISOString()
  };
  const content = JSON.stringify(descriptionData, null, 2);

  if (existingFiles.hasNext()) {
    existingFiles.next().setContent(content);
  } else {
    parentFolder.createFile(descriptionFileName, content, MimeType.PLAIN_TEXT);
  }

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Description saved.'))
    .setNavigation(CardService.newNavigation().popCard())
    .build();
}
