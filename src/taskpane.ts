// Contains the logic for the main Add-In taskpane
/* global document, Office */

import { getSettings } from "./officeData";
import { updateVersion } from "./versionUpdate";
import { setupEditor } from "./editor";

let settings: Office.RoamingSettings;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Check if the add-in has been updated since the last time it was opened
    settings = getSettings();
    updateVersion(settings);

    await setupEditor();

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