const fs = require("fs");
const path = require("path");

async function run() {
  const templatePath = path.join(__dirname, "test-template.docx");
  const dataPath = path.join(__dirname, "test-data.json");

  const form = new FormData();
  const templateBlob = new Blob([fs.readFileSync(templatePath)], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  form.append("template", templateBlob, "test-template.docx");
  form.append("data", fs.readFileSync(dataPath, "utf8"));

  const response = await fetch("http://localhost:3000/render-docx", {
    method: "POST",
    body: form,
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`HTTP ${response.status}: ${errText}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  fs.writeFileSync(path.join(__dirname, "output.docx"), Buffer.from(arrayBuffer));

  console.log("Fertig: output.docx wurde erzeugt.");
}

run().catch(err => {
  console.error("Fehler:", err.message);
});