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
