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

const stylingOptions = {
  bold: false,
  italic: false,
  underline: false,
  fontSize: 16,
};

const stylingDiv = document.getElementById("textStyling");
const boldButton = stylingDiv.querySelector("#boldButton");
const italicButton = stylingDiv.querySelector("#italicButton");
const underlineButton = stylingDiv.querySelector("#underlineButton");
const fontSizeInput = stylingDiv.querySelector("#fontSizeInput");
