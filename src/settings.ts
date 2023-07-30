// Contains all logic concerning the settings menu

import { ADDIN_VERSION } from "./version";
import { focusEditor, manageNoteCategories } from "./editor";
import { getSettings, getIdentifiers } from "./officeData";

export async function setupApplicationSettings() {
  const settings = getSettings();

  setupSettingsButtonAndVersionNumber();
  await setupCategoryDropdowns(settings);
}

function setupSettingsButtonAndVersionNumber() {
  const settingsButton = document.getElementById("settingsButton");
  const settingsContentDiv = document.getElementById("settingContentDiv");

  const versionNumber = document.getElementById("versionNumber");
  versionNumber.textContent = `v${ADDIN_VERSION}`;

  settingsButton.addEventListener("click", () => {
    if (!settingsContentDiv.classList.contains("show")) {
      settingsContentDiv.style.pointerEvents = "all";

      settingsContentDiv.classList.toggle("show");
      versionNumber.classList.toggle("show");

      settingsContentDiv.style.animation = "fadeIn 0.5s forwards";
      versionNumber.style.animation = "fadeIn 0.5s forwards";
    } else {
      settingsContentDiv.style.pointerEvents = "none";
      focusEditor();

      settingsContentDiv.style.animation = "fadeOut 0.5s forwards";
      versionNumber.style.animation = "fadeOut 0.5s forwards";

      setTimeout(() => {
        settingsContentDiv.classList.toggle("show");
        versionNumber.classList.toggle("show");
      }, 500);
    }
  });
}

async function setupCategoryDropdowns(settings: Office.RoamingSettings) {
  const categoryDropdownsDiv: HTMLDivElement = document.getElementById("categoryDropdownsDiv") as HTMLDivElement;
  const messageCategoriesDropdown: HTMLSelectElement = categoryDropdownsDiv.children.namedItem(
    "messageCategoriesDropdown"
  ) as HTMLSelectElement;
  const categoryContextsDropdown: HTMLSelectElement = categoryDropdownsDiv.children.namedItem(
    "categoryContextsDropdown"
  ) as HTMLSelectElement;

  // Set the message categories dropdown to the saved setting
  const savedMessageCategories = settings.get("messageCategories");
  if (savedMessageCategories) {
    messageCategoriesDropdown.value = savedMessageCategories;
    if (savedMessageCategories === "noCategories") {
      categoryContextsDropdown.classList.add("hidden");
    }
  } else {
    messageCategoriesDropdown.value = "mailNotes";
    settings.set("messageCategories", "mailNotes");
    settings.saveAsync();
  }

  // Set the category context dropdown to the saved setting
  const savedCategoryContexts = settings.get("categoryContexts");
  if (savedCategoryContexts) {
    categoryContextsDropdown.value = savedCategoryContexts;
  } else {
    categoryContextsDropdown.value = "all";
    settings.set("categoryContexts", "all");
    settings.saveAsync();
  }

  const allNotes = await settings.get("notes");
  const { mailId, senderId, conversationId, itemSubject, itemNormalizedSubject } = getIdentifiers();

  // If the user changes the message categories dropdown
  messageCategoriesDropdown.addEventListener("change", function () {
    const selectedCategory = messageCategoriesDropdown.value;

    // Hide the category context dropdown if the user selects "No Categories"
    if (selectedCategory === "noCategories") {
      categoryContextsDropdown.classList.add("hidden");
    } else {
      categoryContextsDropdown.classList.remove("hidden");
    }

    // Save the selected setting to the Office Roaming Settings
    settings.set("messageCategories", selectedCategory);
    settings.saveAsync();

    // Update the categories live
    manageNoteCategories(allNotes[mailId], allNotes[conversationId], allNotes[senderId]);
  });

  // If the user changes the category context dropdown
  categoryContextsDropdown.addEventListener("change", function () {
    const selectedContext = categoryContextsDropdown.value;

    // Save the selected setting to the Office Roaming Settings
    settings.set("categoryContexts", selectedContext);
    settings.saveAsync();

    // Update the categories live
    manageNoteCategories(allNotes[mailId], allNotes[conversationId], allNotes[senderId]);
  });
}
