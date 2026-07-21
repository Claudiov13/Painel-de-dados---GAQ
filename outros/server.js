const express = require('express');
const path    = require('path');
const app     = express();
const PORT    = process.env.PORT || 3000;

// Serve todos os arquivos estáticos da pasta do projeto
app.use(express.static(__dirname));

// Rota principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Evita que erros em requisições derrubem o servidor
app.use((err, req, res, next) => {
  console.error('[ERRO]', new Date().toLocaleString('pt-BR'), err.message);
  res.status(500).send('Erro interno do servidor.');
});

const server = app.listen(PORT, () => {
  console.log('');
  console.log('  ✔  Sistema de Compras — Painel GAQ');
  console.log(`  ✔  Rodando em: http://localhost:${PORT}`);
  console.log('');
  console.log('  Para parar o servidor: pressione Ctrl+C');
  console.log('');
});

server.on('error', (err) => {
  console.error('[ERRO SERVIDOR]', err.code, err.message);
  process.exit(1);
});

server.on('close', () => {
  console.error('[SERVIDOR FECHADO]', new Date().toLocaleString('pt-BR'));
});

process.on('exit', (code) => {
  console.log(`[PROCESSO ENCERRADO] codigo: ${code} — ${new Date().toLocaleString('pt-BR')}`);
});

process.on('uncaughtException', (err) => {
  console.error('[EXCEÇÃO NÃO TRATADA]', new Date().toLocaleString('pt-BR'), err.message);
});

process.on('unhandledRejection', (reason) => {
  console.error('[PROMISE REJEITADA]', new Date().toLocaleString('pt-BR'), reason);
});
