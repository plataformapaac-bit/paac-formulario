// ═══════════════════════════════════════════════════════════════════
// BLOCO DE INICIALIZAÇÃO — Cria o usuarios.json a partir da
// variável de ambiente USUARIOS_JSON (configurada no Railway).
// ═══════════════════════════════════════════════════════════════════
const fs_init = require('fs');
const path_init = require('path');
const arquivoUsuarios = path_init.join(__dirname, 'usuarios.json');

if (process.env.USUARIOS_JSON) {
  try {
    JSON.parse(process.env.USUARIOS_JSON);
    fs_init.writeFileSync(arquivoUsuarios, process.env.USUARIOS_JSON, 'utf-8');
    console.log('usuarios.json criado a partir da variavel de ambiente.');
  } catch (e) {
    console.error('ERRO: A variavel USUARIOS_JSON tem JSON invalido:', e.message);
    console.error('Conteudo recebido (primeiros 200 chars):', process.env.USUARIOS_JSON.substring(0, 200));
  }
} else {
  console.log('AVISO: Variavel USUARIOS_JSON nao encontrada.');
  if (!fs_init.existsSync(arquivoUsuarios)) {
    console.error('ERRO CRITICO: usuarios.json nao existe e USUARIOS_JSON nao foi definida.');
  }
}

// ═══════════════════════════════════════════════════════════════════
// PAAC — server.js v2 — Login unificado + Sistema de papéis (roles)
