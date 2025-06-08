/**
 * @OnlyCurrentDoc
 */

/**
 * Runs when the add-on is installed to create the menu.
 * @param {Object} e The event object.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the document is opened. Adds a custom menu to the Add-ons menu.
 * @param {Object} e The event object.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Reset Add-on', 'handleMenuReset')
      .addSeparator()
      .addItem('Help', 'handleMenuHelp')
      .addToUi();
}

/**
 * Runs when the add-on is opened from the sidebar icon. This is the main
 * entry point defined by the homepageTrigger in the manifest.
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
 * Callback for the Docs-specific homepage trigger. This is required by the manifest
 * for file-scope permissions, even if it just calls the main homepage.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onDocsHomepage(e) {
    return onHomepage(e);
}

/**
 * Resets the add-on's configuration from the menu.
 */
function handleMenuReset() {
  try {
    // Clear the stored root folder setting
    PropertiesService.getUserProperties().deleteProperty('rootFolderId');
    
    // Clear the script cache
    CacheService.getScriptCache().removeAll(['contents_root']); // Attempt to clear root cache
    
    const message = 'The add-on has been reset. The next time you open it, you will be asked to configure a root folder.\n\nTo fully remove all permissions, please visit your Google Account settings.';
    DocumentApp.getUi().alert('Reset Complete', message, DocumentApp.getUi().ButtonSet.OK);
  } catch(e) {
    DocumentApp.getUi().alert(`An error occurred during reset: ${e.message}`);
  }
}

/**
 * Shows a help dialog with a link to the GitHub repository.
 */
function handleMenuHelp() {
  const html = HtmlService.createHtmlOutput(
    '<p>You can find help, view the source code, and report issues on the GitHub repository.</p>' +
    '<a href="https://github.com/nikjft/niks-google-docs-templates" target="_blank">Open GitHub Page</a>'
  ).setWidth(300).setHeight(100);
  
  DocumentApp.getUi().showModalDialog(html, 'Help & Support');
}
