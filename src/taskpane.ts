// Contains the logic for the main Add-In taskpane
/* global document, Office */
import Quill from "quill";
var Delta = Quill.import("delta");

import { getIdentifiers, getSettings } from "./officeData";
import { updateVersion } from "./versionUpdate";
import { setActiveContext } from "./context";
import { quill, setupEditor } from "./editor";

let mailId: string, senderId: string, conversationId: string;
let settings: Office.RoamingSettings;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Get the identifiers for the current item
    ({ mailId, senderId, conversationId } = getIdentifiers());

    settings = getSettings();
    updateVersion(settings);

    // Load a possibly already existing note from storage
    await displayExistingNote();

    setupEditor();

    fadeOutOverlay();
  } else {
    console.log("This add-in only supports Outlook clients!");
    document.getElementById("outsideOutlook").style.display = "block";
    document.getElementById("insideOutlook").style.display = "none";
  }
});

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
