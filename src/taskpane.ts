// Contains the logic for the main Add-In taskpane
/* global document, Office */
import Quill from "quill";
var Delta = Quill.import("delta");

import { getIdentifiers, getSettings } from "./officeData";
import { updateVersion } from "./versionUpdate";
import { setActiveContext, getActiveContext } from "./context";

let mailId: string, senderId: string, conversationId: string;
let settings: Office.RoamingSettings;
let quill: Quill;

// Set up the Quill editor even before the Office.onReady event fires, so that the editor is ready to use as soon as possible
setupQuill();

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Get the identifiers for the current item
    ({ mailId, senderId, conversationId } = getIdentifiers());

    settings = getSettings();
    updateVersion(settings);

    // Load a possibly already existing note from storage
    await displayExistingNote();

    fadeOutOverlay();

    // Start the autosave timer
    autosaveNote();
  } else {
    console.log("This add-in only supports Outlook clients!");
    document.getElementById("outsideOutlook").style.display = "block";
    document.getElementById("insideOutlook").style.display = "none";
  }
});

// TODO: Move this to another file
function setupQuill(): void {
  // All options that should be displayed in the editor toolbar
  var toolbarOptions = [
    ["bold", "italic", "underline", "strike"],
    ["link"],
    [{ list: "ordered" }, { list: "bullet" }]
  ];

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

async function displayExistingNote(): Promise<void> {
  // Try to get an existing note for any of the contexts, in descending priority/specificity
  const allNotes = await settings.get("notes");
  const mailNote = allNotes[mailId];
  const conversationNote = allNotes[conversationId];
  const senderNote = allNotes[senderId];

  if (mailNote) {
    quill.setContents(mailNote.noteContents);
    setActiveContext("mail");
  } else if (conversationNote) {
    quill.setContents(conversationNote.noteContents);
    setActiveContext("conversation");
  } else if (senderNote) {
    quill.setContents(senderNote.noteContents);
    setActiveContext("sender");
  }
}

function fadeOutOverlay(): void {
  const overlay = document.getElementById("overlay");
  overlay.classList.add("fade-out");
}

async function saveNote(): Promise<void> {
  const icon = document.getElementById("savingNotice");
  icon.style.visibility = "visible";

  const newNoteContents = quill.getContents();
  const activeContext = getActiveContext();
  const allNotes = await settings.get("notes");

  const updateNote = (noteId: string) => {
    allNotes[noteId] = allNotes[noteId] ?? {};
    allNotes[noteId].noteContents = newNoteContents;
    allNotes[noteId].lastEdited = new Date().toISOString();
  };

  switch (activeContext) {
    case "mail":
      if (newNoteContents.length() === 1 && newNoteContents.ops[0].insert === "\n") {
        delete allNotes[mailId];
      } else {
        updateNote(mailId);
      }
      break;
    case "sender":
      if (newNoteContents.length() === 1 && newNoteContents.ops[0].insert === "\n") {
        delete allNotes[senderId];
      } else {
        updateNote(senderId);
      }
      break;
    case "conversation":
      if (newNoteContents.length() === 1 && newNoteContents.ops[0].insert === "\n") {
        delete allNotes[conversationId];
      } else {
        updateNote(conversationId);
      }
      break;
  }

  // Save the note to storage
  settings.set("notes", allNotes);
  settings.saveAsync();

  // Hide the icon after a timeout
  setTimeout(() => {
    icon.style.visibility = "hidden";
  }, 1000);
}

let autosaveTimeout, previousContext: string;

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

// TODO: Move this to another file
const savingIcon = document.getElementById("savingIcon");

function toggleIconSpinner(toSpinner: boolean): void {
  savingIcon.classList.remove(toSpinner ? "tick" : "spinner");
  savingIcon.classList.add(toSpinner ? "spinner" : "tick");
}
