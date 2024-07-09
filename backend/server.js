const express = require("express");
const app = express();
const fs = require("fs");
const path = require("path");

// Definindo a porta
const PORT = 3000;

app.get("/api/download-model", (req, res) => {
  const filePath = "c:\\Temp\\modelo_cnd.dotx";
  const fileName = path.basename(filePath);

  // Verifica se o arquivo existe
  fs.access(filePath, fs.constants.F_OK, (err) => {
    if (err) {
      console.error("File doesn't exist:", err);
      return res.status(404).send("File not found");
    }

    // Configura os headers para download do arquivo
    res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    // Envia o arquivo
    fs.createReadStream(filePath).pipe(res);
  });
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
