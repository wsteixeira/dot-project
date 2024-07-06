const express = require("express");
const fs = require("fs");
const officegen = require("officegen");
const bodyParser = require("body-parser");

const app = express();
const port = 3000;

// Middleware para parsear JSON no corpo da requisição
app.use(bodyParser.json());

app.post("/api/generate-doc", async (req, res) => {
  // Dados recebidos do frontend
  const { name, age } = req.body;

  // Criar um novo documento do Word
  let docx = officegen({
    type: "docx",
    orientation: "portrait",
    // Opções específicas para 'docx'
    subject: "Exemplo de Documento",
    keywords: "exemplo, documento",
    description: "Um documento gerado como exemplo.",
  });

  // Tratamento de erros
  docx.on("error", (err) => {
    console.error(err);
    res.status(500).send("Erro ao gerar o documento");
  });

  // Adicionar conteúdo ao documento
  // Aqui você precisará adaptar para manipular DOCVARIABLE conforme necessário
  let pObj = docx.createP();
  pObj.addText(`Nome: ${name}`);
  pObj.addLineBreak();
  pObj.addText(`Idade: ${age}`);

  // Salvar o documento gerado no caminho especificado
  const outputPath = "C:\\temp\\documento.docx";
  let out = fs.createWriteStream(outputPath);

  out.on("error", (err) => {
    console.error(err);
    res.status(500).send("Erro ao salvar o documento");
  });

  // Gera o documento usando officegen e salva no sistema de arquivos
  docx.generate(out);

  out.on("close", () => {
    // Responder ao frontend que o documento foi criado com sucesso
    res.send({ message: "Documento gerado com sucesso" });
  });
});

app.listen(port, () => {
  console.log(`Servidor rodando na porta ${port}`);
});
