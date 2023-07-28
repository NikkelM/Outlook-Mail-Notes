// Contains the logic for the main Add-In taskpane
/* global document, Office */

import { getSettings } from "./officeData";
import { updateVersion } from "./versionUpdate";
import { setupEditor } from "./editor";
import { ADDIN_VERSION } from "./version";

let settings: Office.RoamingSettings;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Check if the add-in has been updated since the last time it was opened
    settings = getSettings();
    await updateVersion(settings);

    await setupEditor();

    const settingsButton = document.getElementById("settingsButton");
    const settingContentDiv = document.getElementById("settingContentDiv");
    const versionNumber = document.getElementById("versionNumber");

    versionNumber.textContent = `v${ADDIN_VERSION}`;

    settingsButton.addEventListener("click", () => {

      if (!settingContentDiv.classList.contains("show")) {
        settingContentDiv.style.pointerEvents = "all";
        settingContentDiv.classList.toggle("show");
        versionNumber.classList.toggle("show");
        settingContentDiv.style.animation = "fadeIn 0.5s forwards";
        versionNumber.style.animation = "fadeIn 0.5s forwards";
      } else {
        // TODO: Focus editor
        settingContentDiv.style.animation = "fadeOut 0.5s forwards";
        settingContentDiv.style.pointerEvents = "none";
        versionNumber.style.animation = "fadeOut 0.5s forwards";
        setTimeout(() => {
          settingContentDiv.classList.toggle("show");
          versionNumber.classList.toggle("show");
        }, 500);
      }
    });

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
