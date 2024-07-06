const express = require("express");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const bodyParser = require("body-parser");

const app = express();
const port = 3000;

// Middleware para parsear JSON no corpo da requisição
app.use(bodyParser.json());

app.post("/api/generate-doc", (req, res) => {
  // Dados recebidos do frontend
  const { name, age } = req.body;

  // Carregar o modelo DOTX
  const content = fs.readFileSync("C:\\Temp\\modelo.docx", "binary");

  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // Injetar os dados no modelo
  doc.setData({
    name: name,
    age: age,
  });

  try {
    doc.render();
  } catch (error) {
    console.error(error);
    return res.status(500).send("Erro ao gerar o documento");
  }

  const buf = doc.getZip().generate({ type: "nodebuffer" });

  // Salvar o documento gerado no caminho especificado
  const outputPath = "C:\\temp\\documento.docx";
  fs.writeFileSync(outputPath, buf);

  // Responder ao frontend que o documento foi criado com sucesso
  res.send({ message: "Documento gerado com sucesso" });
});

app.listen(port, () => {
  console.log(`Servidor rodando na porta ${port}`);
});
