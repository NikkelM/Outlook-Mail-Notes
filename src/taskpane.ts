// Contains the logic for the main Add-In taskpane
/* global document, Office, console */

import { getSettings, setupCategoryMasterList } from "./officeData";
import { updateVersion } from "./versionUpdate";
import { setupEditor } from "./editor";
import { setupApplicationSettings } from "./settings";
import { setupContextButtons } from "./context";
import { setupNoteExport } from "./export";

let settings: Office.RoamingSettings;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Check if the add-in has been updated since the last time it was opened
    settings = getSettings();
    await updateVersion(settings);

    // Make sure the category master list is set up correctly
    // Users can manually delete categories, so we need to re-add them if necessary
    await setupCategoryMasterList();

    await setupApplicationSettings();

    await setupEditor();

    setupNoteExport();

    setupContextButtons();

    fadeOutOverlay();
  } else {
    console.log("This add-in only supports Outlook clients!");
    document.getElementById("outsideOutlook").style.display = "block";
    document.getElementById("insideOutlook").style.display = "none";
  }
});

function fadeOutOverlay(): void {
  const overlay = document.getElementById("overlay");
  overlay.classList.add("fade-out");
}
