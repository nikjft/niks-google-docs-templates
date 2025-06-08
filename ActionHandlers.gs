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

function handleRefreshAction(e) {
    const params = e.parameters;
    const folderId = params.folderId;
    Logger.log(`Clearing cache for folderId: ${folderId}`);
    const cacheKey = `contents_${folderId}`;
    CacheService.getScriptCache().remove(cacheKey);
    
    const path = JSON.parse(params.path);
    const selection = JSON.parse(params.selection);
    const newCard = createBrowserCard(folderId, path, selection);

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Folder refreshed."))
        .setNavigation(CardService.newNavigation().updateCard(newCard))
        .build();
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

function handleEditAction(e) {
    const fileId = e.parameters.fileId;
    const file = DriveApp.getFileById(fileId);
    const url = file.getUrl();
    const openLink = CardService.newOpenLink().setUrl(url);
    return CardService.newActionResponseBuilder().setOpenLink(openLink).build();
}

function handleSaveDescriptionAction(e) {
  const params = e.parameters;
  saveFileDescription(params.fileId, params.parentFolderId, e.formInput.description_text);
  
  const path = JSON.parse(params.path);
  const selection = JSON.parse(params.selection);
  const card = createBrowserCard(params.parentFolderId, path, selection);
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Description saved.'))
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}
