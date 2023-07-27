import Quill from "quill";
var Delta = Quill.import("delta");

import { getSettings, getIdentifiers } from "./officeData";
import { getActiveContext, switchToContext } from "./context";

export let quill: Quill;
let mailId: string, senderId: string, conversationId: string, itemSubject: string, itemNormalizedSubject: string;
let settings: Office.RoamingSettings;
// Used to determine whether or not to show the autosave icon
let previousContext: string;

// Set up the Quill editor even before the Office.onReady event fires, so that the editor is ready to use as soon as possible
setupQuill();

// ----- Setup -----
export async function setupEditor(): Promise<void> {
  // Get the identifiers for the current item
  ({ mailId, senderId, conversationId, itemSubject, itemNormalizedSubject } = getIdentifiers());

  settings = getSettings();

  await displayInitialNote();

  // Start the autosave timer
  previousContext = getActiveContext();
  autosaveNote();
}

function setupQuill(): void {
  var Link = Quill.import("formats/link");
  class MyLink extends Link {
    static create(value) {
      if (!value.startsWith("http://") && !value.startsWith("https://")) {
        value = "http://" + value;
      }
      const node = super.create(value);
      return node;
    }
  }
  Quill.register(MyLink, true);

  // All options that should be displayed in the editor toolbar
  var toolbarOptions = [["bold", "italic", "underline", "strike"], ["link"], [{ list: "ordered" }, { list: "bullet" }]];

  // Defines the Quill editor
  quill = new Quill("#noteInput", {
    modules: {
      toolbar: toolbarOptions,
    },
    placeholder:
      "Jot down some notes here - your changes are automatically saved.\nUse the toolbar above to style your text!",
    theme: "snow",
  });
}

async function displayInitialNote(): Promise<void> {
  // Try to get an existing note for any of the contexts, in descending priority/specificity
  const allNotes = await settings.get("notes");
  let mailNote = allNotes[mailId];
  const conversationNote = allNotes[conversationId];
  const senderNote = allNotes[senderId];

  // With v1.2.0, the mailId was changed from using the item.itemId to use item.conversationId_item.dateTimeCreated.toISOString()
  // We need to check if the current item is still using the old ID format, and if so, update it
  const pre1_2_0Notes = await settings.get("pre1_2_0Notes");
  if (pre1_2_0Notes) {
    console.log("Checking for pre-1.2.0 note");
    mailNote = await pre1_2_0Update(Office.context.mailbox.item.itemId, allNotes, pre1_2_0Notes, settings);
  }

  if (mailNote) {
    await switchToContext("mail", quill, mailId, settings);
  } else if (conversationNote) {
    await switchToContext("conversation", quill, conversationId, settings);
  } else if (senderNote) {
    await switchToContext("sender", quill, senderId, settings);
  } else {
    // The default context is the mail context
    await switchToContext("mail", quill, mailId, settings);
  }
}

async function pre1_2_0Update(mailId: string, allNotes: any, pre1_2_0Notes: any, settings: Office.RoamingSettings): Promise<string> {
  // Generate the new ItemId
  const newItemId =
    Office.context.mailbox.item.conversationId +
    "_" +
    new Date(Office.context.mailbox.item.dateTimeCreated).toISOString();

  if (pre1_2_0Notes[mailId]) {
    allNotes[newItemId] = {};
    allNotes[newItemId].noteContents = allNotes[mailId].noteContents;
    allNotes[newItemId].lastEdited = allNotes[mailId].lastEdited;
    delete pre1_2_0Notes[mailId];
    delete allNotes[mailId];

    if (Object.keys(pre1_2_0Notes).length === 0) {
      settings.remove("pre1_2_0Notes");
    } else {
      settings.set("pre1_2_0Notes", pre1_2_0Notes);
    }

    settings.set("notes", allNotes);
    settings.saveAsync();
  }
  return newItemId;
}

// ----- Note saving -----
let autosaveTimeout;
function autosaveNote() {
  let accumulatedChanges = new Delta();

  quill.on("text-change", function (delta) {
    // If the context was changed, we do not want to display the saving icon
    if (getActiveContext() !== previousContext) {
      previousContext = getActiveContext();
      savingIcon.style.visibility = "hidden";
      return;
    }

    toggleIconSpinner(true);
    savingIcon.style.visibility = "visible";

    accumulatedChanges = accumulatedChanges.compose(delta);

    // Changes are saved after a period of inactivity
    clearTimeout(autosaveTimeout);
    autosaveTimeout = setTimeout(function () {
      saveNote();
      toggleIconSpinner(false);
      accumulatedChanges = new Delta();
    }, 750);
  });

  // Changes are always saved after a set timeout, even if the user is still typing, but only if there are changes to save
  setInterval(function () {
    if (accumulatedChanges.length() > 0) {
      saveNote();
      accumulatedChanges = new Delta();
      // Don't change the icon appearance, as it would get switched back to the spinner by the next text-change event immediately
    }
  }, 5000);
}

// The autosave icon
const savingIcon = document.getElementById("savingIcon");

function toggleIconSpinner(toSpinner: boolean): void {
  savingIcon.classList.remove(toSpinner ? "tick" : "spinner");
  savingIcon.classList.add(toSpinner ? "spinner" : "tick");
}

async function saveNote(): Promise<void> {
  const icon = document.getElementById("savingNotice");
  icon.style.visibility = "visible";

  const newNoteContents = quill.getContents();
  const activeContext = getActiveContext();
  const allNotes = await settings.get("notes");

  const contextMapping = {
    mail: mailId,
    sender: senderId,
    conversation: conversationId,
  };

  if (newNoteContents.length() === 1 && newNoteContents.ops[0].insert === "\n") {
    delete allNotes[contextMapping[activeContext]];

    manageItemCategories(contextMapping[activeContext], activeContext, false);
  } else {
    allNotes[contextMapping[activeContext]] = allNotes[contextMapping[activeContext]] ?? {};
    allNotes[contextMapping[activeContext]].noteContents = newNoteContents;
    allNotes[contextMapping[activeContext]].lastEdited = new Date().toISOString().split("T")[0];

    manageItemCategories(contextMapping[activeContext], activeContext, true);
  }

  // Save the note to storage
  settings.set("notes", allNotes);
  settings.saveAsync();

  // Hide the icon after a timeout
  setTimeout(() => {
    icon.style.visibility = "hidden";
  }, 1000);
}

function manageItemCategories(item: string, activeContext: string, shouldAdd: boolean): void {
  // Remove the category from the item if the note is empty
  if (!shouldAdd) {
    Office.context.mailbox.item.categories.removeAsync(["Mail Notes"], function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Removing category failed with error: " + asyncResult.error.message);
      }
    });
  } else {
    Office.context.mailbox.item.categories.addAsync(["Mail Notes"], function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Setting category failed with error: " + asyncResult.error.message);
      }
    });
  }
}
