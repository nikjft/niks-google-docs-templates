/**
 * @OnlyCurrentDoc
 */

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
