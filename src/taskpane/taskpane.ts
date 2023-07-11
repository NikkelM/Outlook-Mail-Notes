/* global document, Office */
import Quill from "quill";

let mailbox: Office.Mailbox;
let settings: Office.RoamingSettings;
let mailItem;
let quill: Quill;
// let conversation, sender;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
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
      placeholder: "Jot down some notes here.\nUse the toolbar above to style your text!",
      theme: "snow",
    });

    document.getElementById("saveNoteButton").onclick = saveNote;
    mailbox = Office.context.mailbox;
    settings = Office.context.roamingSettings;
    mailItem = mailbox.item.itemId;
    // conversation = mailbox.item.conversationId;
    // sender = mailbox.item.from.emailAddress;

    displayExistingNote();
  }
});

async function displayExistingNote() {
  const note = await settings.get(mailItem);
  if (note) {
    quill.setContents(note);
  }
}

async function saveNote() {
  const button: HTMLButtonElement = document.getElementById("saveNoteButton") as HTMLButtonElement;
  
  const note = quill.getContents();
  settings.set(mailItem, note);
  settings.saveAsync();
  button.textContent = "Saved";
  setTimeout(() => {
    button.textContent = "Save note";
  }, 2000);
}
