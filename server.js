// ═══════════════════════════════════════════════════════════════════
// BLOCO DE INICIALIZAÇÃO — Cria o usuarios.json a partir da
// variável de ambiente USUARIOS_JSON (configurada no Railway).
// Isso é necessário porque o Railway não permite subir arquivos
// manualmente — então o conteúdo fica numa variável e o servidor
// cria o arquivo ao iniciar.
// ═══════════════════════════════════════════════════════════════════
const fs_init = require('fs');
const path_init = require('path');
const arquivoUsuarios = path_init.join(__dirname, 'usuarios.json');

if (process.env.USUARIOS_JSON && !fs_init.existsSync(arquivoUsuarios)) {
  fs_init.writeFileSync(arquivoUsuarios, process.env.USUARIOS_JSON, 'utf-8');
  console.log('usuarios.json criado a partir da variavel de ambiente.');
}

// ═══════════════════════════════════════════════════════════════════
// PAAC — server.js v2 — Login unificado + Sistema de papéis (roles)
