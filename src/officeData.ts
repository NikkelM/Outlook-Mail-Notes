// This file contains handlers for data the add-in gets from the Office API
/* global Office */

import { DEFAULT_ADDIN_CATEGORIES } from "./constants";

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

export async function setupCategoryMasterList() {
  // Get the categories saved in settings
  const settings = getSettings();
  let userAddInCategories = await settings.get("addinCategories");
  // If there are no categories saved in settings, use the default categories
  if (!userAddInCategories) {
    settings.set("addinCategories", DEFAULT_ADDIN_CATEGORIES);
    settings.saveAsync();
    userAddInCategories = DEFAULT_ADDIN_CATEGORIES;
  }

  // For each category, make sure it exists in the master list
  const addinCategories = Object.values(userAddInCategories);
  Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    const masterCategories = asyncResult.value;
    // Add all categories that don't exist yet
    let categoriesToAdd = [];
    addinCategories.forEach((category: any) => {
      if (
        !masterCategories.find(
          (masterCategory) =>
            masterCategory.displayName === category.displayName && masterCategory.color === category.color
        )
      ) {
        categoriesToAdd.push(category);
      }
    });

    Office.context.mailbox.masterCategories.addAsync(categoriesToAdd);
  });
}
