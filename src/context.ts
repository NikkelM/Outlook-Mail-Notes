// Contains logic to handle context buttons
/* global document, Office */

import Quill from "quill";
import { getIdentifiers, getSettings } from "./officeData";

const contextButtons = {
  mail: document.getElementById("emailContextButton"),
  sender: document.getElementById("senderContextButton"),
  conversation: document.getElementById("conversationContextButton"),
};
let activeContext;

for (const [key, button] of Object.entries(contextButtons)) {
  button.addEventListener("click", () => {
    setActiveContext(key);
    loadNoteForContext(key);
  });
}

export function initContextButtons() {
  // Set the initial active button
  contextButtons.mail.classList.add("active");
  activeContext = "mail";
}

export function setActiveContext(context: string) {
  const button = contextButtons[context];
  if (!button) {
    throw new Error("Invalid context");
  }

  for (const [key, value] of Object.entries(contextButtons)) {
    if (key === context) {
      value.classList.add("active");
    } else {
      value.classList.remove("active");
    }
  }

  activeContext = context;
}

export function getActiveContext(): string {
  return activeContext;
}

export async function loadNoteForContext(
  context: string,
  quill?: Quill,
  itemId?: string,
  settings?: Office.RoamingSettings
) {
  if (!settings) {
    settings = getSettings();
  }

  if (!quill) {
    quill = Quill.find(document.getElementById("noteInput"));
  }

  if (!itemId) {
    const identifiers = getIdentifiers();
    switch (context) {
      case "mail":
        itemId = identifiers.mailId;
        break;
      case "sender":
        itemId = identifiers.senderId;
        break;
      case "conversation":
        itemId = identifiers.conversationId;
        break;
      default:
        throw new Error("Invalid context");
    }
  }

  const allNotes = await settings.get("notes");

  let noteContents = allNotes[itemId]?.noteContents ?? null;
  quill.setContents(noteContents);
}
