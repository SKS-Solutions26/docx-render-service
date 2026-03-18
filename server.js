const express = require("express");
const multer = require("multer");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
const upload = multer();

app.post("/render-docx", upload.single("template"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        error: "Keine Template-Datei empfangen. Erwartet wird das Feld 'template'.",
      });
    }

    if (!req.body.data) {
      return res.status(400).json({
        error: "Keine JSON-Daten empfangen. Erwartet wird das Feld 'data'.",
      });
    }

    const templateBuffer = req.file.buffer;
    const data = JSON.parse(req.body.data);

    const zip = new PizZip(templateBuffer);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.render(data);

    const outputBuffer = doc.getZip().generate({
      type: "nodebuffer",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="output.docx"'
    );

    res.send(outputBuffer);
  } catch (error) {
    console.error("Fehler beim Rendern:", error);

    res.status(500).json({
      error: "Rendering fehlgeschlagen",
      details: error.message,
    });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Render-Service läuft auf http://localhost:${PORT}`);
});