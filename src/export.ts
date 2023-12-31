// Contains the logic for importing and exporting notes
/* global document */

import { getSettings } from "./officeData";

export function setupNoteExport(): void {
  // Get the exportButton and add exportToCSV as a click event listener
  const exportButton = document.getElementById("exportButton");
  exportButton.addEventListener("click", async () => {
    await exportToCSV();
  });
}

async function exportToCSV(): Promise<void> {
  const settings = getSettings();
  const allNotes = await settings.get("notes");

  // Create the CSV file
  let csvFile = "Note ID,Note,Last edited,Internal representation (must exist when importing notes)\n";

  // For each note, add a row to the CSV file
  Object.keys(allNotes).forEach((noteID) => {
    const note = allNotes[noteID];

    const delta = note.noteContents.ops;
    const plaintext = getPlainText(delta);
    const lastEdited = note.lastEdited;

    // Escape commas and double quotes in plaintext
    const escapedPlaintext = plaintext.replace(/\n/g, "\\n").replace(/"/g, '""');
    const csvPlaintext = `"${escapedPlaintext}"`;

    const deltaString = `"${JSON.stringify(delta).replace(/\n/g, "\\n").replace(/"/g, '""')}"`;

    const escapedNoteID = `"${noteID.toString().replace(/"/g, '""')}"`;
    const escapedLastEdited = `"${lastEdited.toString().replace(/"/g, '""')}"`;

    csvFile += `${escapedNoteID},${csvPlaintext},${escapedLastEdited},${deltaString}\n`;
  });

  // Download the CSV file
  const fileName = `MailNotes_export_${new Date().toISOString().slice(0, -5)}.csv`;
  downloadCSVFile(csvFile, fileName);
}

function getPlainText(delta): string {
  let plaintext = "";
  delta.forEach((op) => {
    if (op.insert) {
      if (typeof op.insert === "string") {
        plaintext += op.insert;
      } else {
        plaintext += " ";
      }
    }
    if (op.delete) {
      plaintext = plaintext.slice(0, -op.delete);
    }
  });
  // If the last character is a newline, omit it
  if (plaintext.slice(-1) === "\n") {
    plaintext = plaintext.slice(0, -1);
  }
  return plaintext;
}

function downloadCSVFile(data: string, fileName: string) {
  const blob = new Blob([data], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}
