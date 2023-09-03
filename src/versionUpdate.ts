// Contains logic that handles the version update process
/* global Office, console */

import { ADDIN_VERSION } from "./version";

export async function updateVersion(settings: Office.RoamingSettings) {
  // Get the current version from the settings object
  const currentVersion = (await settings.get("version")) ?? "0.0.0";

  if (currentVersion < ADDIN_VERSION) {
    console.log("Updating Add-In from version " + currentVersion + " to version " + ADDIN_VERSION);

    if (currentVersion < "0.0.2") {
      // Move all note data from the root of the settings object to the "notes" property
      let notes = {};
      for (const key in settings["settingsData"]) {
        if (key !== "version") {
          notes[key] = {};
          notes[key]["noteContents"] = settings["settingsData"][key];
          notes[key]["lastEdited"] = new Date().toISOString();
          delete settings["settingsData"][key];
        }
      }
      settings.set("notes", notes);
    }

    if (currentVersion < "1.2.0") {
      // The mailId was changed from using the item.itemId to use item.conversationId_item.dateTimeCreated.toISOString()
      let pre1_2_0Notes = {};
      const allNotes = await settings.get("notes");

      // For each note, check if the ID is an itemId
      for (const key in allNotes) {
        if (key.startsWith("AQM")) {
          pre1_2_0Notes[key] = 1;
        }
      }

      if (Object.keys(pre1_2_0Notes).length > 0) {
        settings.set("pre1_2_0Notes", pre1_2_0Notes);
      }
      settings.set("notes", allNotes);
    }

    // Update the version in the settings object
    settings.set("version", ADDIN_VERSION);

    // Update the settings object
    settings.saveAsync();
  }
}
