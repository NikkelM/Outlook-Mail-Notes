// Contains logic to handle context buttons
/* global document, Office */

import Quill from "quill";
import { focusEditor } from "./editor";
import { getIdentifiers, getSettings } from "./officeData";

// Text element showing when the current note was last edited
const lastEditedNotice: HTMLParagraphElement = document.getElementById("lastEditedNotice") as HTMLParagraphElement;

const contextButtons = {
  mail: document.getElementById("emailContextButton"),
  sender: document.getElementById("senderContextButton"),
  conversation: document.getElementById("conversationContextButton"),
};
let activeContext;

export function setupContextButtons(): void {
  for (const [key, button] of Object.entries(contextButtons)) {
    button.addEventListener("click", () => {
      switchToContext(key);
    });
  }
}

export function getActiveContext(): string {
  return activeContext;
}

export async function switchToContext(
  context: string,
  quill?: Quill,
  itemId?: string,
  settings?: Office.RoamingSettings
): Promise<void> {
  if (context === activeContext) {
    return;
  }
  setActiveContext(context);
  await loadNoteForContext(context, quill, itemId, settings);
}

function setActiveContext(context: string) {
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

async function loadNoteForContext(context: string, quill?: Quill, itemId?: string, settings?: Office.RoamingSettings) {
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

  const noteContents = allNotes[itemId]?.noteContents ?? null;
  quill.setContents(noteContents);

  updateLastEditedNotice(itemId, allNotes);

  focusEditor();
}

export function updateLastEditedNotice(itemId: string, allNotes: any) {
  let lastEdited = allNotes[itemId]?.lastEdited ?? null;

  if (lastEdited) {
    const date: any = new Date(lastEdited);
    const today: any = new Date();
    const yesterday: any = new Date(today);
    yesterday.setDate(today.getDate() - 1);

    console.log(navigator.language);

    const dayDiff = Math.floor((today - date) / (1000 * 60 * 60 * 24));

    let formattedDate;
    if (date.toDateString() === today.toDateString()) {
      formattedDate = "Today";
    } else if (date.toDateString() === yesterday.toDateString()) {
      formattedDate = "Yesterday";
    } else if (dayDiff < 7) {
      formattedDate = date.toLocaleDateString(navigator.language, { weekday: "long", month: "long", day: "numeric" });
    } else {
      formattedDate = date.toLocaleDateString(navigator.language, {
        year: "numeric",
        month: "long",
        day: "numeric",
      });
    }

    lastEditedNotice.innerHTML = `<i>Last edited: ${formattedDate}</i>`;
  } else {
    lastEditedNotice.innerText = "";
  }
}
