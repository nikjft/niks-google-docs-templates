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
    .setTitle('Google Drive™ Folder URL or ID')
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

  // --- Add items to the ... card menu ---
  const resetAction = CardService.newAction().setFunctionName('handleResetAction');
  card.addCardAction(CardService.newCardAction()
    .setText('Reset Root Folder')
    .setOnClickAction(resetAction));
    
  const newTemplateAction = CardService.newAction().setFunctionName('handleCreateNewTemplate')
      .setParameters({ folderId: folderId, path: JSON.stringify(path), selection: JSON.stringify(selection) });
  card.addCardAction(CardService.newCardAction()
      .setText('New Template...')
      .setOnClickAction(newTemplateAction));
      
  const refreshAction = CardService.newAction().setFunctionName('handleRefreshAction')
      .setParameters({ folderId: folderId, path: JSON.stringify(path), selection: JSON.stringify(selection) });
  card.addCardAction(CardService.newCardAction()
      .setText('Reload Folder')
      .setOnClickAction(refreshAction));

  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderContents = getFolderContents(folderId);
    
    // --- Breadcrumb & Controls Section ---
    const controlsSection = CardService.newCardSection().setHeader("Path");
    const breadcrumbSet = CardService.newButtonSet();
    const rootId = PropertiesService.getUserProperties().getProperty('rootFolderId');

    if (rootId) {
        // The path passed in is the list of ancestors. The current folder is `folder`.
        // We construct the full path for display by combining them.
        const fullDisplayPath = [...path, { id: folder.getId(), name: folder.getName() }];

        fullDisplayPath.forEach((p, index) => {
            const isLastItem = index === fullDisplayPath.length - 1;
            
            // The path to navigate to is the list of items *before* the clicked item.
            const navigationPath = fullDisplayPath.slice(0, index);

            const button = CardService.newTextButton()
                .setText(p.name)
                .setDisabled(isLastItem) // Disable the last button (current folder)
                .setTextButtonStyle(isLastItem ? CardService.TextButtonStyle.FILLED : CardService.TextButtonStyle.TEXT)
                .setOnClickAction(CardService.newAction()
                    .setFunctionName('handleNavigation')
                    .setParameters({
                        folderId: p.id,
                        path: JSON.stringify(navigationPath), 
                        selection: '{}'
                    }));
            breadcrumbSet.addButton(button);
        });
        
        controlsSection.addWidget(breadcrumbSet);
    }
    card.addSection(controlsSection);

    // --- Content Section ---
    const contentSection = CardService.newCardSection().setHeader("Contents");
    let hasContent = false;

    // Subfolders
    folderContents.folders.forEach(subfolder => {
        hasContent = true;
        const newPathForSubfolder = [...path, { id: folderId, name: folder.getName() }];
        contentSection.addWidget(CardService.newDecoratedText()
            .setText(`📁 ${subfolder.name}`)
            .setWrapText(true)
            .setOnClickAction(CardService.newAction().setFunctionName('handleNavigation').setParameters({ folderId: subfolder.id, path: JSON.stringify(newPathForSubfolder), selection: '{}' })));
    });
    
    // Files
    folderContents.files.forEach(file => {
      hasContent = true;
      const isSelected = selection && selection.id === file.id;

      const fileWidget = createFileWidget(file, folderId, path, selection);
      contentSection.addWidget(fileWidget);

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
        .setText("✅")
        .setOnClickAction(insertAction);

    const infoAction = CardService.newAction().setFunctionName('handleInfoAction')
        .setParameters({ fileId: selection.id, parentFolderId: parentFolderId, path: JSON.stringify(path), selection: JSON.stringify(selection) });
    const infoButton = CardService.newTextButton()
        .setText("💬")
        .setOnClickAction(infoAction);

    const editAction = CardService.newAction().setFunctionName('handleEditAction')
        .setParameters({ fileId: selection.id });
    const editButton = CardService.newTextButton()
        .setText("✍️")
        .setOnClickAction(editAction);

    return CardService.newButtonSet().addButton(insertButton).addButton(infoButton).addButton(editButton);
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
    icon = '🟢';
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
