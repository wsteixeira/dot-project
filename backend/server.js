const express = require("express");
const app = express();
const port = 3000;

// Middleware para analisar JSON
app.use(express.json());

// Rota para gerar documento (código anterior)
app.post("/api/generate-doc", async (req, res) => {
  // Implementação omitida para brevidade
});

// Rota para download do modelo.dotx
app.get("/api/template", (req, res) => {
  const templatePath = "C:\\Temp\\modelo_x.dotx";
  res.download(templatePath); // Isso fará o download do arquivo quando acessado
});

// Iniciar o servidor
app.listen(port, () => {
  console.log(`Servidor rodando em http://localhost:${port}`);
});
