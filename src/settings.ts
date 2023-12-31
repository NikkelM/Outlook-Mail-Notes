// Contains all logic concerning the settings menu
/* global document, Office, console, setTimeout */

import { ADDIN_VERSION } from "./version";
import { CATEGORY_COLORS, IS_IE } from "./constants";
import { focusEditor, manageNoteCategories } from "./editor";
import { getSettings, getIdentifiers } from "./officeData";

export async function setupApplicationSettings() {
  const settings = getSettings();

  setupSettingsButtonAndVersionNumber();
  await setupCategoryDropdowns(settings);

  setupCategoryCustomization(settings);
}

function setupSettingsButtonAndVersionNumber() {
  const settingsButton = document.getElementById("settingsButton");
  const settingsContentDiv = document.getElementById("settingsContentDiv");
  const dynamicSettingsFooter = document.getElementById("dynamicSettingsFooter");

  // Get the current version number
  const versionNumber = document.getElementById("versionNumber");
  versionNumber.textContent = `v${ADDIN_VERSION}`;

  // Toggle the settings menu and the dynamic footer content when the settings button is clicked
  settingsButton.addEventListener("click", () => {
    if (!settingsContentDiv.classList.contains("show")) {
      settingsContentDiv.style.pointerEvents = "all";

      settingsContentDiv.classList.toggle("show");
      dynamicSettingsFooter.classList.toggle("show");

      settingsContentDiv.style.animation = "fadeIn 0.5s forwards";
      dynamicSettingsFooter.style.animation = "fadeIn 0.5s forwards";
    } else {
      settingsContentDiv.style.pointerEvents = "none";
      focusEditor();

      settingsContentDiv.style.animation = "fadeOut 0.5s forwards";
      dynamicSettingsFooter.style.animation = "fadeOut 0.5s forwards";

      setTimeout(() => {
        settingsContentDiv.classList.toggle("show");
        dynamicSettingsFooter.classList.toggle("show");
      }, 500);
    }
  });
}

function setupCategoryCustomization(settings: Office.RoamingSettings) {
  showRelevantCustomization(settings);
  setupCategoryNameInputs(settings);
  setupCategoryColorPicker(settings);
}

async function showRelevantCustomization(settings: Office.RoamingSettings) {
  const categoryNamesInputsDiv = document.getElementById("categoryNamesInputsDiv") as HTMLDivElement;
  const generalCategoryCustomization = document.getElementById("generalCategoryCustomization") as HTMLDivElement;
  const messageCategoryCustomization = document.getElementById("messageCategoryCustomization") as HTMLDivElement;
  const conversationCategoryCustomization = document.getElementById(
    "conversationCategoryCustomization"
  ) as HTMLDivElement;
  const senderCategoryCustomization = document.getElementById("senderCategoryCustomization") as HTMLDivElement;

  const messageCategories = await settings.get("messageCategories");
  const categoryContexts = await settings.get("categoryContexts");

  // By default, hide all customization options
  generalCategoryCustomization.classList.add("removed");
  messageCategoryCustomization.classList.add("removed");
  conversationCategoryCustomization.classList.add("removed");
  senderCategoryCustomization.classList.add("removed");
  categoryNamesInputsDiv.classList.add("removed");

  switch (messageCategories) {
    case "mailNotes":
      generalCategoryCustomization.classList.remove("removed");
      messageCategoryCustomization.classList.add("removed");
      conversationCategoryCustomization.classList.add("removed");
      senderCategoryCustomization.classList.add("removed");
      categoryNamesInputsDiv.classList.remove("removed");
      break;
    case "unique":
      categoryNamesInputsDiv.classList.remove("removed");
      if (categoryContexts === "all") {
        messageCategoryCustomization.classList.remove("removed");
        conversationCategoryCustomization.classList.remove("removed");
        senderCategoryCustomization.classList.remove("removed");
      } else if (categoryContexts === "messagesConversations") {
        messageCategoryCustomization.classList.remove("removed");
        conversationCategoryCustomization.classList.remove("removed");
        senderCategoryCustomization.classList.add("removed");
      } else if (categoryContexts === "messages") {
        messageCategoryCustomization.classList.remove("removed");
        conversationCategoryCustomization.classList.add("removed");
        senderCategoryCustomization.classList.add("removed");
      } else {
        throw new Error("Invalid category context");
      }
      break;
    case "noCategories":
      break;
    default:
      throw new Error("Invalid message categories");
  }
}

async function setupCategoryDropdowns(settings: Office.RoamingSettings) {
  const categoryDropdownsDiv = document.getElementById("categoryDropdownsDiv") as HTMLDivElement;
  const messageCategoriesDropdown = categoryDropdownsDiv.children.namedItem(
    "messageCategoriesDropdown"
  ) as HTMLSelectElement;
  const categoryContextsDropdown = categoryDropdownsDiv.children.namedItem(
    "categoryContextsDropdown"
  ) as HTMLSelectElement;

  // Set the message categories dropdown to the saved setting
  const savedMessageCategories = await settings.get("messageCategories");
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
  const savedCategoryContexts = await settings.get("categoryContexts");
  if (savedCategoryContexts) {
    categoryContextsDropdown.value = savedCategoryContexts;
  } else {
    categoryContextsDropdown.value = "all";
    settings.set("categoryContexts", "all");
    settings.saveAsync();
  }

  const allNotes = await settings.get("notes");
  const { mailId, senderId, conversationId } = getIdentifiers();

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

    showRelevantCustomization(settings);

    // Update the categories live
    manageNoteCategories(allNotes[mailId], allNotes[conversationId], allNotes[senderId]);
  });

  // If the user changes the category context dropdown
  categoryContextsDropdown.addEventListener("change", function () {
    const selectedContext = categoryContextsDropdown.value;

    // Save the selected setting to the Office Roaming Settings
    settings.set("categoryContexts", selectedContext);
    settings.saveAsync();

    showRelevantCustomization(settings);

    // Update the categories live
    manageNoteCategories(allNotes[mailId], allNotes[conversationId], allNotes[senderId]);
  });
}

async function setupCategoryNameInputs(settings: Office.RoamingSettings) {
  // Get the category input elements
  const categoryInputs: NodeListOf<HTMLInputElement> = document.querySelectorAll(".category-input");
  const addinCategories = await settings.get("addinCategories");

  categoryInputs.forEach((categoryInput) => {
    // The key in the settings object is the same as the id of the input minus the 'NameInput' suffix
    const inputId = categoryInput.id.slice(0, -9);

    // Set the input value to the saved setting
    categoryInput.value = addinCategories[inputId].displayName;

    // If the user changes the category name input, update the master category list
    categoryInput.addEventListener("change", async function () {
      // If the name was not changed, do nothing
      if (categoryInput.value === addinCategories[inputId].displayName) {
        return;
      }

      const newCategory = {
        displayName: categoryInput.value,
        color: addinCategories[inputId].color,
      };
      await updateMasterCategories(addinCategories[inputId], newCategory);

      // Save the new name to the settings
      addinCategories[inputId].displayName = categoryInput.value;
      settings.set("addinCategories", addinCategories);
      settings.saveAsync();
    });
  });
}

async function setupCategoryColorPicker(settings: Office.RoamingSettings) {
  if (IS_IE) {
    console.log("Customizing category colours is not supported in Internet Explorer.");
    return;
  }

  // Get the category input elements
  const categoryInputs: NodeListOf<HTMLInputElement> = document.querySelectorAll(".category-input");
  const addinCategories = await settings.get("addinCategories");

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

      // If this color is the active color, add the .active-color class
      if (color.preset === addinCategories[inputId].color) {
        colorPickerCell.classList.add("active-color");
      }

      colorPickerCell.addEventListener("click", async function () {
        // Hide the dropdown
        colorPickerDropdown.style.display = "none";

        if (color.preset === addinCategories[inputId].color) {
          return;
        }

        // Update the background color of the button
        colorPickerButton.style.backgroundColor = color.value;

        // Change the active-color class
        const activeColor = colorPickerGrid.querySelector(".active-color");
        activeColor.classList.remove("active-color");
        colorPickerCell.classList.add("active-color");

        // Update the master category list
        const newCategory = {
          displayName: addinCategories[inputId].displayName,
          color: color.preset,
        };
        await updateMasterCategories(addinCategories[inputId], newCategory);

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
      // While it is open, clicking anywhere not in the dropdown will close it
      document.addEventListener("click", function closeDropdown(event: any) {
        if (!colorPicker.contains(event.target)) {
          colorPickerDropdown.style.display = "none";
          document.removeEventListener("click", closeDropdown);
        }
      });
    });
  });
}

async function updateMasterCategories(oldCategory: any, newCategory: any) {
  // Only remove the master category if the names are different
  const categoryToRemove = [oldCategory.displayName];
  const categoryToAdd = [newCategory];

  Office.context.mailbox.item.categories.removeAsync(categoryToRemove, function () {
    Office.context.mailbox.masterCategories.removeAsync(categoryToRemove, async function () {
      const options: Office.AsyncContextOptions = {
        asyncContext: [],
      };

      while (
        options.asyncContext.length === 0 ||
        options.asyncContext.find(
          (masterCategory) =>
            masterCategory.displayName === oldCategory.displayName && masterCategory.color === oldCategory.color
        )
      ) {
        await new Promise((resolve) => setTimeout(resolve, 5));
        Office.context.mailbox.masterCategories.getAsync(options, function (asyncResult) {
          const newMasterCategories = asyncResult.value;
          options.asyncContext = newMasterCategories;
        });
      }

      Office.context.mailbox.masterCategories.addAsync(categoryToAdd, async function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log("Adding master category failed with error: " + asyncResult.error.message);
        }

        const options: Office.AsyncContextOptions = {
          asyncContext: [],
        };
        while (
          !options.asyncContext.find(
            (masterCategory) =>
              masterCategory.displayName === newCategory.displayName && masterCategory.color === newCategory.color
          )
        ) {
          await new Promise((resolve) => setTimeout(resolve, 5));
          Office.context.mailbox.masterCategories.getAsync(options, async function (asyncResult) {
            const newMasterCategories = asyncResult.value;
            options.asyncContext = newMasterCategories;
          });
        }

        // Add the new category to the current item
        const { mailId, senderId, conversationId } = getIdentifiers();
        const settings = getSettings();
        const allNotes = await settings.get("notes");
        manageNoteCategories(allNotes[mailId], allNotes[conversationId], allNotes[senderId]);
      });
    });
  });
}
