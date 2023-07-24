// This file contains handlers for data the add-in gets from the Office API

export function getIdentifiers() {
  // Get references to the mailbox and the current item
  let mailbox: Office.Mailbox = Office.context.mailbox;

  // Return the identifiers as an object
  return {
    mailId: mailbox.item.itemId,
    senderId: mailbox.item.from.emailAddress,
    conversationId: mailbox.item.conversationId
  };
}

export function getSettings() {
  // All notes are saved in the 'settings' object
  return Office.context.roamingSettings;
}
