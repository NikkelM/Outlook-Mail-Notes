import Quill from "quill";
var Delta = Quill.import("delta");

import { getSettings, getIdentifiers } from "./officeData";
import { getActiveContext, switchToContext } from "./context";

export let quill: Quill;
let mailId: string, senderId: string, conversationId: string;
let settings: Office.RoamingSettings;
// Used to determine whether or not to show the autosave icon
let previousContext: string;

// Set up the Quill editor even before the Office.onReady event fires, so that the editor is ready to use as soon as possible
setupQuill();

// ----- Setup -----
export async function setupEditor(): Promise<void> {
  // Get the identifiers for the current item
  ({ mailId, senderId, conversationId } = getIdentifiers());

  settings = getSettings();

  await displayInitialNote();

  // Start the autosave timer
  previousContext = getActiveContext();
  autosaveNote();
}

function setupQuill(): void {
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
  const mailNote = allNotes[mailId];
  const conversationNote = allNotes[conversationId];
  const senderNote = allNotes[senderId];

  if (mailNote) {
    switchToContext("mail", quill, mailId, settings);
  } else if (conversationNote) {
    switchToContext("conversation", quill, conversationId, settings);
  } else if (senderNote) {
    switchToContext("sender", quill, senderId, settings);
  }
}

// ----- Note saving -----
let autosaveTimeout;
function autosaveNote() {
  let accumulatedChanges = new Delta();

  quill.on("text-change", function (delta) {
    // If the context was changed, we do not want to display the saving icon
    if (getActiveContext() !== previousContext) {
      console.log("change");
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
  } else {
    allNotes[contextMapping[activeContext]] = allNotes[contextMapping[activeContext]] ?? {};
    allNotes[contextMapping[activeContext]].noteContents = newNoteContents;
    allNotes[contextMapping[activeContext]].lastEdited = new Date().toISOString();
  }

  // Save the note to storage
  settings.set("notes", allNotes);
  settings.saveAsync();

  // Hide the icon after a timeout
  setTimeout(() => {
    icon.style.visibility = "hidden";
  }, 1000);
}
