// This file contains handlers for data the add-in gets from the Office API

export function getIdentifiers() {
  // Get references to the mailbox and the current item
  let mailbox: Office.Mailbox = Office.context.mailbox;

  // The item.itemId changes if the item is moved to a different folder
  // We use the item.conversationId instead and append the creation date of the mail
  const uniqueMailId = mailbox.item.conversationId + "_" + new Date(mailbox.item.dateTimeCreated).toISOString();

  // Return the identifiers as an object
  return {
    mailId: uniqueMailId,
    senderId: mailbox.item.from.emailAddress,
    conversationId: mailbox.item.conversationId,
    itemSubject: mailbox.item.subject,
    itemNormalizedSubject: mailbox.item.normalizedSubject,
  };
}

export function getSettings() {
  // All notes are saved in the 'settings' object
  return Office.context.roamingSettings;
}

export function setupCategoryMasterList() {
  console.log("Setting up category master list...");
  // First check if the "Mail Notes" category already exists, and if not, add it
  Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    const masterCategories = asyncResult.value;
    if (!masterCategories.find((category) => category.displayName === "Mail Notes")) {
      const masterCategoriesToAdd = [
        {
          displayName: "Mail Notes",
          color: "Preset7",
        },
      ];

      Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd);
    }
  });
}
