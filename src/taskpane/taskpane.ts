/* global document, Office */
import Quill from "quill";
var Delta = Quill.import("delta");

let mailbox: Office.Mailbox;
let settings: Office.RoamingSettings;
let mailItem: string;
// let conversation, sender;
let quill: Quill;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    setupQuill();

    // Set up references to the mailbox and the current item
    mailbox = Office.context.mailbox;
    settings = Office.context.roamingSettings;
    mailItem = mailbox.item.itemId;
    // conversation = mailbox.item.conversationId;
    // sender = mailbox.item.from.emailAddress;

    // Load a possible existing note from storage
    await displayExistingNote();
    // Start the autosave timer
    autosaveNote();
  }
});

function setupQuill(): void {
  // All options that should be displayed in the editor toolbar
  var toolbarOptions = [
    [{ size: ["small", false, "large", "huge"] }],
    ["bold", "italic", "underline", "strike"],
    ["link", "image"],
    [{ color: [] }, { background: [] }],
    [{ list: "ordered" }, { list: "bullet" }],
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
  const note = await settings.get(mailItem);
  if (note) {
    quill.setContents(note);
  }
}

async function saveNote(): Promise<void> {
  const button: HTMLButtonElement = document.getElementById("savingNoteButton") as HTMLButtonElement;
  button.style.display = "inline-block";

  const note = quill.getContents();

  // Save the note to storage
  settings.set(mailItem, note);
  settings.saveAsync();

  // Hide the button after a timeout
  setTimeout(() => {
    button.style.display = "none";
  }, 1000);
}

let autosaveTimeout: NodeJS.Timeout;

function autosaveNote() {
  let accumulatedChanges = new Delta();

  quill.on("text-change", function (delta) {
    accumulatedChanges = accumulatedChanges.compose(delta);

    // Changes are saved after a period of inactivity
    clearTimeout(autosaveTimeout);
    autosaveTimeout = setTimeout(function () {
      saveNote();
      accumulatedChanges = new Delta();
    }, 750);
  });

  // Changes are always saved after a set timeout, even if the user is still typing, but only if there are changes to save
  setInterval(function () {
    if (accumulatedChanges.length() > 0) {
      saveNote();
      accumulatedChanges = new Delta();
    }
  }, 5000);
}
