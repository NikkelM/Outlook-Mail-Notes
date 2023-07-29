// Contains all logic concerning the settings menu

import { ADDIN_VERSION } from "./version";
import { focusEditor } from "./editor";
import { getSettings } from "./officeData";

export function setupApplicationSettings() {
  const settings = getSettings();

  setupSettingsButtonAndVersionNumber();
  setupCategoryDropdowns(settings);
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

function setupCategoryDropdowns(settings: Office.RoamingSettings) {
  const categoryDropdownsDiv: HTMLDivElement = document.getElementById("categoryDropdownsDiv") as HTMLDivElement;
  const messageCategoriesDropdown: HTMLSelectElement = categoryDropdownsDiv.children.namedItem(
    "messageCategoriesDropdown"
  ) as HTMLSelectElement;
  const categoryContextDropdown: HTMLSelectElement = categoryDropdownsDiv.children.namedItem(
    "categoryContextDropdown"
  ) as HTMLSelectElement;

  // Set the message categories dropdown to the saved setting
  const savedMessageCategories = settings.get("messageCategories");
  if (savedMessageCategories) {
    messageCategoriesDropdown.value = savedMessageCategories;
    if (savedMessageCategories === "noCategories") {
      categoryContextDropdown.classList.add("hidden");
    }
  } else {
    messageCategoriesDropdown.value = "mailNotes";
    settings.set("messageCategories", "mailNotes");
    settings.saveAsync();
  }

  // Set the category context dropdown to the saved setting
  const savedCategoryContext = settings.get("categoryContext");
  if (savedCategoryContext) {
    categoryContextDropdown.value = savedCategoryContext;
  } else {
    categoryContextDropdown.value = "all";
    settings.set("categoryContext", "all");
    settings.saveAsync();
  }

  // If the user changes the message categories dropdown
  messageCategoriesDropdown.addEventListener("change", () => {
    const selectedCategory = messageCategoriesDropdown.value;

    // Hide the category context dropdown if the user selects "No Categories"
    if (selectedCategory === "noCategories") {
      categoryContextDropdown.classList.add("hidden");
    } else {
      categoryContextDropdown.classList.remove("hidden");
    }

    // Save the selected setting to the Office Roaming Settings
    settings.set("messageCategories", selectedCategory);
    settings.saveAsync();
  });

  // If the user changes the category context dropdown
  categoryContextDropdown.addEventListener("change", () => {
    const selectedContext = categoryContextDropdown.value;

    // Save the selected setting to the Office Roaming Settings
    settings.set("categoryContext", selectedContext);
    settings.saveAsync();
  });
}
