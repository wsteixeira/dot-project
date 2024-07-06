const express = require("express");
const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");

const app = express();
const dotxFilePath = "c:/Temp/modelo.dotx";

app.get("/api/convert", async (req, res) => {
  try {
    const buffer = fs.readFileSync(dotxFilePath);

    // Converter .dotx para HTML usando mammoth
    const { value: html } = await mammoth.convertToHtml({ buffer });

    // Converter o HTML de volta para um documento .docx
    const docxBuffer = Buffer.from(html, "utf-8");

    res.set({
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": "attachment; filename=modelo.docx",
    });

    res.send(docxBuffer);
  } catch (error) {
    console.error("Error converting file:", error);
    res.status(500).send("Error converting file");
  }
});

app.listen(3000, () => {
  console.log("Server is running on port 3000");
});
