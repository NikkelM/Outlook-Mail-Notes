// Contains the logic for importing and exporting notes

export function exportToCSV(allNotes): void {
  // Create the CSV file
  let csvFile = "Note ID,Note plaintext,Note styled text,Last edited\n";

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

    csvFile += `${escapedNoteID},${csvPlaintext},${deltaString},${escapedLastEdited}\n`;
  });

  // Download the CSV file
  downloadCSVFile(csvFile, new Date().toISOString() + ".csv");
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
