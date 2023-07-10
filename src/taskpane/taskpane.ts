/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("saveNoteButton").onclick = saveNote;
  }
});

export async function saveNote() {
  const button: HTMLButtonElement = document.getElementById("saveNoteButton") as HTMLButtonElement;
  button.textContent = "Saving...";
  button.disabled = true;
}
