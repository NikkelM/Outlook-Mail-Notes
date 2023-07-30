// Contains all logic concerning the settings menu

import { ADDIN_VERSION } from "./version";
import { CATEGORY_COLORS } from "./constants";
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

  setupCategoryColorPicker();

  // Set the message categories dropdown to the saved setting
  const savedMessageCategories = settings.get("messageCategories");
  if (savedMessageCategories) {
    messageCategoriesDropdown.value = savedMessageCategories;
    if (savedMessageCategories === "noCategories") {
      categoryContextsDropdown.classList.add("removed");
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
      categoryContextsDropdown.classList.add("removed");
    } else {
      categoryContextsDropdown.classList.remove("removed");
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

function setupCategoryColorPicker() {
  const settings = getSettings();
  // Get the category input elements
  const categoryInputs: NodeListOf<HTMLInputElement> = document.querySelectorAll(".category-input");
  const addinCategories = settings.get("addinCategories");

  categoryInputs.forEach((categoryInput) => {
    // The key in the settings object is the same as the id of the input minus the 'NameInput' suffix
    const inputId = categoryInput.id.slice(0, -9);

    // Create a color picker container
    const colorPicker = document.createElement("div");
    colorPicker.classList.add("color-picker");

    // Create a color picker button
    const colorPickerButton = document.createElement("button");
    colorPickerButton.classList.add("color-picker-button");
    colorPickerButton.title = "Select a color";
    colorPicker.appendChild(colorPickerButton);
    colorPickerButton.style.backgroundColor = CATEGORY_COLORS.find(
      (c) => c.preset === addinCategories[inputId].color
    ).value;

    // Create a dropdown
    const colorPickerDropdown = document.createElement("div");
    colorPickerDropdown.classList.add("color-picker-dropdown");
    colorPickerDropdown.style.display = "none";

    // Create the grid within the dropdown
    const colorPickerGrid = document.createElement("div");
    colorPickerGrid.classList.add("color-picker-grid");

    // Loop through the colors and create a color picker cell element for each color
    CATEGORY_COLORS.forEach((color) => {
      const colorPickerCell = document.createElement("div");
      colorPickerCell.classList.add("color-picker-cell");
      colorPickerCell.style.backgroundColor = color.value;
      colorPickerCell.title = color.name;
      colorPickerCell.addEventListener("click", () => {
        // Update the background color of the button
        colorPickerButton.style.backgroundColor = color.value;
        // Hide the dropdown
        colorPickerDropdown.style.display = "none";

        // Save the new color to the settings
        addinCategories[inputId].color = color.preset;
        settings.set("addinCategories", addinCategories);
        settings.saveAsync();
      });
      colorPickerGrid.appendChild(colorPickerCell);
    });

    colorPickerDropdown.appendChild(colorPickerGrid);
    colorPicker.appendChild(colorPickerDropdown);

    // Insert the color picker after the category input
    categoryInput.parentNode.insertBefore(colorPicker, categoryInput.nextSibling);

    // Add an event listener to the button
    colorPickerButton.addEventListener("click", () => {
      // Toggle the display of the dropdown
      if (colorPickerDropdown.style.display === "none") {
        colorPickerDropdown.style.display = "block";
      } else {
        colorPickerDropdown.style.display = "none";
      }
    });
  });
}
