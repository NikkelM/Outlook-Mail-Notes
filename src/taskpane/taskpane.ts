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
  } else {
    console.log("This add-in only supports Outlook clients!");
    document.getElementById("outsideOutlook").style.display = "block";
    document.getElementById("insideOutlook").style.display = "none";
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
  const icon = document.getElementById("savingNotice");
  icon.style.visibility = "visible";

  const note = quill.getContents();

  // Save the note to storage
  settings.set(mailItem, note);
  settings.saveAsync();

  // Hide the icon after a timeout
  setTimeout(() => {
    icon.style.visibility = "hidden";
  }, 1000);
}

let autosaveTimeout;

function autosaveNote() {
  let accumulatedChanges = new Delta();

  quill.on("text-change", function (delta) {
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
      // Don't change the icon appearance, as it would get switched back to the spinner by the next text-change event immediately
      saveNote();
      accumulatedChanges = new Delta();
    }
  }, 5000);
}

const savingIcon = document.getElementById("savingIcon");

function toggleIconSpinner(toSpinner: boolean): void {
  savingIcon.classList.remove(toSpinner ? "tick" : "spinner");
  savingIcon.classList.add(toSpinner ? "spinner" : "tick");
}
