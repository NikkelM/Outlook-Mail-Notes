import { ADDIN_VERSION } from "./version";
import { focusEditor } from "./editor";

// TODO: Async?
export async function setupSettings() {
  const settingsButton = document.getElementById("settingsButton");
  const settingsContentDiv = document.getElementById("settingContentDiv");

  const versionNumber = document.getElementById("versionNumber");
  versionNumber.textContent = `v${ADDIN_VERSION}`;

  settingsButton.addEventListener("click", () => {
    if (!settingsContentDiv.classList.contains("show")) {
      settingsContentDiv.style.pointerEvents = "all";

      settingsContentDiv.classList.toggle("show");
      versionNumber.classList.toggle("show");

      settingsContentDiv.style.animation = "fadeIn 0.5s forwards";
      versionNumber.style.animation = "fadeIn 0.5s forwards";
    } else {
      settingsContentDiv.style.pointerEvents = "none";
	  focusEditor();

      settingsContentDiv.style.animation = "fadeOut 0.5s forwards";
      versionNumber.style.animation = "fadeOut 0.5s forwards";

      setTimeout(() => {
        settingsContentDiv.classList.toggle("show");
        versionNumber.classList.toggle("show");
      }, 500);
    }
  });
}
