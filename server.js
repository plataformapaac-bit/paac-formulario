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
  }
} else {
  console.log('AVISO: Variavel USUARIOS_JSON nao encontrada. Usando arquivo local.');
}

// ═══════════════════════════════════════════════════════════════════
// PAAC — server.js v2 — Login unificado + Sistema de papéis (roles)
// ═══════════════════════════════════════════════════════════════════
//
// ARQUITETURA:
// 1. Tela de login ÚNICA para todos (e-mail + senha)
// 2. Após login:
//    - Usuário comum → vai direto para o formulário (index.html)
//    - Usuário admin  → vê tela com 2 opções:
//      a) "Acessar Formulário" (usa como usuário normal)
//      b) "Painel Administrativo" (pede SEGUNDA senha exclusiva)
// 3. Painel Admin: cadastro de usuários, promoção a admin, etc.
//
// REGRA INVIOLÁVEL DE SEGURANÇA:
// - Nenhum admin pode excluir, bloquear ou trocar a senha de outro admin
// - Cada admin só gerencia: usuários comuns + seus próprios dados
// - Um admin PODE cadastrar outro admin quando decidir (sem obrigação)
// - Após criado, o novo admin torna-se mutuamente intocável
//
// TRÊS SENHAS DISTINTAS:
// 1. "senhaHash" de cada usuário → para fazer login na plataforma
// 2. "senhaAdminHash" global → segunda senha para entrar no painel admin
// 3. Cada admin pode trocar APENAS sua própria senha de login
// ═══════════════════════════════════════════════════════════════════

const express = require('express');
const bcrypt = require('bcryptjs');
const cookieParser = require('cookie-parser');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const COOKIE_SECRET = process.env.COOKIE_SECRET || 'paac-e0-segredo-temporario-2026';

// ── Limite de tentativas de login (proteção contra força bruta) ──
// Armazena em memória: { "email": { tentativas: N, bloqueadoAte: timestamp } }
const tentativasLogin = {};
const MAX_TENTATIVAS = 5;          // máximo de erros permitidos
const TEMPO_BLOQUEIO = 15 * 60000; // 15 minutos em milissegundos

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser(COOKIE_SECRET));

// ═══════════════════════════════════════════════════════════════════
// FUNÇÕES AUXILIARES
// ═══════════════════════════════════════════════════════════════════

function lerDados() {
  try {
    return JSON.parse(fs.readFileSync(path.join(__dirname, 'usuarios.json'), 'utf-8'));
  } catch (e) {
    console.error('Erro ao ler usuarios.json:', e.message);
    return { senhaAdminHash: '', usuarios: [] };
  }
}

function salvarDados(dados) {
  fs.writeFileSync(
    path.join(__dirname, 'usuarios.json'),
    JSON.stringify(dados, null, 2),
    'utf-8'
  );
}

function buscarUsuario(email) {
  const dados = lerDados();
  return dados.usuarios.find(u => u.email.toLowerCase() === email.toLowerCase().trim());
}

// Verifica se o e-mail está bloqueado por excesso de tentativas
function verificarBloqueio(email) {
  const reg = tentativasLogin[email.toLowerCase()];
  if (!reg) return false;
  if (reg.bloqueadoAte && Date.now() < reg.bloqueadoAte) {
    const minutosRestantes = Math.ceil((reg.bloqueadoAte - Date.now()) / 60000);
    return minutosRestantes;
  }
  // Bloqueio expirou, limpa o registro
  if (reg.bloqueadoAte && Date.now() >= reg.bloqueadoAte) {
    delete tentativasLogin[email.toLowerCase()];
  }
  return false;
}

// Registra uma tentativa falha de login
function registrarFalha(email) {
  const chave = email.toLowerCase();
  if (!tentativasLogin[chave]) tentativasLogin[chave] = { tentativas: 0 };
  tentativasLogin[chave].tentativas++;
  if (tentativasLogin[chave].tentativas >= MAX_TENTATIVAS) {
    tentativasLogin[chave].bloqueadoAte = Date.now() + TEMPO_BLOQUEIO;
  }
}

// Limpa as tentativas após login bem-sucedido
function limparTentativas(email) {
  delete tentativasLogin[email.toLowerCase()];
}

// Lê o cookie de sessão e retorna os dados ou null
function lerSessao(req) {
  try {
    const cookie = req.signedCookies.paac_sessao;
    if (!cookie) return null;
    return JSON.parse(cookie);
  } catch (e) {
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════
// ROTA PRINCIPAL: GET /
// ═══════════════════════════════════════════════════════════════════

app.get('/', (req, res) => {
  const sessao = lerSessao(req);

  // Sem sessão → tela de login
  if (!sessao || !sessao.email) {
    return res.send(htmlLogin());
  }

  const usuario = buscarUsuario(sessao.email);

  // Sessão existe mas usuário foi removido ou desativado
  if (!usuario || !usuario.ativo) {
    res.clearCookie('paac_sessao');
    return res.send(htmlLogin('Sua conta foi desativada. Contate o administrador.'));
  }

  // Usuário admin → tela com 2 opções
  if (usuario.papel === 'admin') {
    return res.send(htmlEscolhaAdmin(usuario.nome));
  }

  // Usuário comum → formulário direto
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ═══════════════════════════════════════════════════════════════════
// ROTA: POST /login — Autenticação unificada (todos passam por aqui)
// ═══════════════════════════════════════════════════════════════════

app.post('/login', async (req, res) => {
  const email = (req.body.email || '').trim();
  const senha = req.body.senha || '';

  if (!email || !senha) {
    return res.send(htmlLogin('Informe e-mail e senha.'));
  }

  // Verifica bloqueio por tentativas excessivas
  const bloqueado = verificarBloqueio(email);
  if (bloqueado) {
    return res.send(htmlLogin(
      'Login bloqueado por excesso de tentativas. Aguarde ' + bloqueado + ' minuto(s).'
    ));
  }

  const usuario = buscarUsuario(email);

  if (!usuario) {
    registrarFalha(email);
    return res.send(htmlLogin('E-mail ou senha incorretos.'));
  }

  if (!usuario.ativo) {
    return res.send(htmlLogin('Sua conta está desativada. Contate o administrador.'));
  }

  // Compara a senha digitada com o hash armazenado
  const senhaOk = await bcrypt.compare(senha, usuario.senhaHash);
  if (!senhaOk) {
    registrarFalha(email);
    const reg = tentativasLogin[email.toLowerCase()];
    const restantes = MAX_TENTATIVAS - (reg ? reg.tentativas : 0);
    const aviso = restantes > 0
      ? ' (' + restantes + ' tentativa(s) restante(s))'
      : '';
    return res.send(htmlLogin('E-mail ou senha incorretos.' + aviso));
  }

  // Login OK
  limparTentativas(email);

  // Registra último acesso
  const dados = lerDados();
  const idx = dados.usuarios.findIndex(u => u.email.toLowerCase() === email.toLowerCase());
  if (idx >= 0) {
    dados.usuarios[idx].ultimoAcesso = new Date().toISOString();
    salvarDados(dados);
  }

  // Cria cookie de sessão (24 horas)
  res.cookie('paac_sessao', JSON.stringify({
    email: usuario.email,
    nome: usuario.nome,
    papel: usuario.papel
  }), {
    signed: true,
    httpOnly: true,
    maxAge: 24 * 60 * 60 * 1000,
    sameSite: 'lax'
  });

  res.redirect('/');
});

// ═══════════════════════════════════════════════════════════════════
// ROTA: GET /formulario — Acesso ao index.html (para admins que
// clicaram em "Acessar Formulário" na tela de escolha)
// ═══════════════════════════════════════════════════════════════════

app.get('/formulario', (req, res) => {
  const sessao = lerSessao(req);
  if (!sessao) return res.redirect('/');

  const usuario = buscarUsuario(sessao.email);
  if (!usuario || !usuario.ativo) return res.redirect('/');

  res.sendFile(path.join(__dirname, 'index.html'));
});

// ═══════════════════════════════════════════════════════════════════
// ROTA: GET /sair — Encerra sessão
// ═══════════════════════════════════════════════════════════════════

app.get('/sair', (req, res) => {
  res.clearCookie('paac_sessao');
  res.clearCookie('paac_admin');
  res.redirect('/');
});

// ═══════════════════════════════════════════════════════════════════
// ROTAS DO PAINEL ADMIN
// ═══════════════════════════════════════════════════════════════════

// GET /painel — Pede a segunda senha OU mostra o painel
app.get('/painel', (req, res) => {
  const sessao = lerSessao(req);
  if (!sessao) return res.redirect('/');

  const usuario = buscarUsuario(sessao.email);
  if (!usuario || usuario.papel !== 'admin') return res.redirect('/');

  // Verifica se já digitou a segunda senha (cookie separado)
  if (req.signedCookies.paac_admin === 'autenticado') {
    return res.send(htmlPainelAdmin(usuario));
  }

  // Ainda não digitou → mostra tela da segunda senha
  res.send(htmlSegundaSenha());
});

// POST /painel/auth — Valida a segunda senha do admin
app.post('/painel/auth', async (req, res) => {
  const sessao = lerSessao(req);
  if (!sessao) return res.redirect('/');

  const usuario = buscarUsuario(sessao.email);
  if (!usuario || usuario.papel !== 'admin') return res.redirect('/');

  const senhaAdmin = req.body.senha_admin || '';
  const dados = lerDados();

  const ok = await bcrypt.compare(senhaAdmin, dados.senhaAdminHash);
  if (!ok) {
    return res.send(htmlSegundaSenha('Senha do painel incorreta.'));
  }

  // Segunda senha OK → cookie de admin (4 horas)
  res.cookie('paac_admin', 'autenticado', {
    signed: true,
    httpOnly: true,
    maxAge: 4 * 60 * 60 * 1000,
    sameSite: 'lax'
  });

  res.redirect('/painel');
});

// ═══════════════════════════════════════════════════════════════════
// API DO PAINEL ADMIN — Gerenciamento de usuários
// ═══════════════════════════════════════════════════════════════════

// Middleware: exige sessão de admin + segunda senha validada
function exigirAdmin(req, res, next) {
  const sessao = lerSessao(req);
  if (!sessao) return res.status(401).json({ erro: 'Sessão expirada.' });

  const usuario = buscarUsuario(sessao.email);
  if (!usuario || usuario.papel !== 'admin') {
    return res.status(403).json({ erro: 'Sem permissão.' });
  }

  if (req.signedCookies.paac_admin !== 'autenticado') {
    return res.status(401).json({ erro: 'Segunda senha não validada.' });
  }

  // Passa o e-mail do admin logado para as rotas usarem
  req.adminEmail = sessao.email.toLowerCase();
  next();
}

// GET /painel/api/usuarios — Lista todos os usuários
app.get('/painel/api/usuarios', exigirAdmin, (req, res) => {
  const dados = lerDados();
  // Retorna a lista SEM os hashes de senha (segurança)
  const lista = dados.usuarios.map(u => ({
    email: u.email,
    nome: u.nome,
    papel: u.papel,
    observacao: u.observacao || '',
    ativo: u.ativo,
    protegido: u.protegido || false,
    cadastradoEm: u.cadastradoEm,
    ultimoAcesso: u.ultimoAcesso
  }));
  res.json(lista);
});

// POST /painel/api/usuarios — Cadastra novo usuário (comum ou admin)
app.post('/painel/api/usuarios', exigirAdmin, async (req, res) => {
  const { email, nome, senha, observacao, papel } = req.body;

  if (!email || !nome || !senha) {
    return res.status(400).json({ erro: 'E-mail, nome e senha são obrigatórios.' });
  }

  if (senha.length < 6) {
    return res.status(400).json({ erro: 'Senha deve ter no mínimo 6 caracteres.' });
  }

  const dados = lerDados();
  const emailLower = email.toLowerCase().trim();

  if (dados.usuarios.some(u => u.email.toLowerCase() === emailLower)) {
    return res.status(409).json({ erro: 'Este e-mail já está cadastrado.' });
  }

  // Valida o papel: só aceita "usuario" ou "admin"
  const papelFinal = (papel === 'admin') ? 'admin' : 'usuario';

  // Limite: máximo 2 admins no sistema (regra E0)
  if (papelFinal === 'admin') {
    const qtdAdmins = dados.usuarios.filter(u => u.papel === 'admin').length;
    if (qtdAdmins >= 2) {
      return res.status(400).json({ erro: 'Limite de 2 administradores atingido.' });
    }
  }

  // Criptografa a senha com bcrypt (10 salt rounds)
  const senhaHash = await bcrypt.hash(senha, 10);

  dados.usuarios.push({
    email: emailLower,
    nome: nome.trim(),
    senhaHash: senhaHash,
    papel: papelFinal,
    observacao: (observacao || '').trim(),
    ativo: true,
    protegido: papelFinal === 'admin', // admins são automaticamente protegidos
    cadastradoEm: new Date().toISOString(),
    ultimoAcesso: null
  });

  salvarDados(dados);
  res.json({ ok: true, mensagem: papelFinal === 'admin'
    ? 'Administrador cadastrado. Ele agora é mutuamente protegido.'
    : 'Usuário cadastrado com sucesso.'
  });
});

// PUT /painel/api/usuarios/:email/toggle — Ativa/desativa usuário
app.put('/painel/api/usuarios/:email/toggle', exigirAdmin, (req, res) => {
  const dados = lerDados();
  const emailAlvo = decodeURIComponent(req.params.email).toLowerCase();
  const idx = dados.usuarios.findIndex(u => u.email.toLowerCase() === emailAlvo);

  if (idx < 0) return res.status(404).json({ erro: 'Usuário não encontrado.' });

  // REGRA INVIOLÁVEL: admin não pode bloquear outro admin
  if (dados.usuarios[idx].papel === 'admin') {
    return res.status(403).json({
      erro: 'Operação proibida: não é possível desativar outro administrador.'
    });
  }

  dados.usuarios[idx].ativo = !dados.usuarios[idx].ativo;
  salvarDados(dados);
  res.json({ ok: true, ativo: dados.usuarios[idx].ativo });
});

// DELETE /painel/api/usuarios/:email — Remove usuário
app.delete('/painel/api/usuarios/:email', exigirAdmin, (req, res) => {
  const dados = lerDados();
  const emailAlvo = decodeURIComponent(req.params.email).toLowerCase();
  const idx = dados.usuarios.findIndex(u => u.email.toLowerCase() === emailAlvo);

  if (idx < 0) return res.status(404).json({ erro: 'Usuário não encontrado.' });

  // REGRA INVIOLÁVEL: admin não pode excluir outro admin
  if (dados.usuarios[idx].papel === 'admin') {
    return res.status(403).json({
      erro: 'Operação proibida: não é possível remover outro administrador.'
    });
  }

  dados.usuarios.splice(idx, 1);
  salvarDados(dados);
  res.json({ ok: true, mensagem: 'Usuário removido.' });
});

// PUT /painel/api/minha-senha — Admin troca APENAS sua própria senha de login
app.put('/painel/api/minha-senha', exigirAdmin, async (req, res) => {
  const { senhaAtual, senhaNova } = req.body;

  if (!senhaAtual || !senhaNova) {
    return res.status(400).json({ erro: 'Informe a senha atual e a nova.' });
  }
  if (senhaNova.length < 6) {
    return res.status(400).json({ erro: 'Nova senha deve ter no mínimo 6 caracteres.' });
  }

  const dados = lerDados();
  const idx = dados.usuarios.findIndex(u => u.email.toLowerCase() === req.adminEmail);
  if (idx < 0) return res.status(404).json({ erro: 'Usuário não encontrado.' });

  const ok = await bcrypt.compare(senhaAtual, dados.usuarios[idx].senhaHash);
  if (!ok) return res.status(401).json({ erro: 'Senha atual incorreta.' });

  dados.usuarios[idx].senhaHash = await bcrypt.hash(senhaNova, 10);
  salvarDados(dados);
  res.json({ ok: true, mensagem: 'Sua senha de login foi alterada.' });
});

// ═══════════════════════════════════════════════════════════════════
// HTML — TELA DE LOGIN (igual para todos)
// ═══════════════════════════════════════════════════════════════════
function htmlLogin(msgErro) {
  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PAAC — Acesso</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:linear-gradient(135deg,#0d1f3c 0%,#1a3a6b 60%,#0a1628 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.4);width:100%;max-width:440px;overflow:hidden}
.hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);border-bottom:3px solid #c9a84c;padding:22px 28px;text-align:center}
.sigla{font-family:'Cormorant Garamond',serif;font-size:2.8rem;font-weight:700;color:#c9a84c;letter-spacing:5px}
.sub{color:rgba(255,255,255,.7);font-size:.75rem;margin-top:4px}
.body{padding:24px 28px}
.aviso{background:#fff8ec;border:1px solid rgba(201,168,76,.4);border-radius:10px;padding:10px 13px;font-size:.74rem;color:#7a5a1a;margin-bottom:16px;line-height:1.5}
.aviso strong{display:block;margin-bottom:3px;color:#5a3a00}
.erro{background:#fdf2f2;border:1px solid rgba(192,57,43,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#c0392b;margin-bottom:12px}
.campo{display:flex;flex-direction:column;gap:4px;margin-bottom:12px}
.campo label{font-size:.77rem;font-weight:600;color:#0d1f3c}
.campo input{border:1.5px solid #dde2ec;border-radius:10px;padding:10px 13px;font-size:.87rem;font-family:'DM Sans',sans-serif;outline:none;transition:border-color .2s}
.campo input:focus{border-color:#2a5298;box-shadow:0 0 0 3px rgba(42,82,152,.1)}
.btn{width:100%;background:linear-gradient(135deg,#c9a84c,#a87d2a);color:#0d1f3c;border:none;border-radius:10px;padding:12px;font-family:'DM Sans',sans-serif;font-size:.93rem;font-weight:700;cursor:pointer;margin-top:4px;box-shadow:0 4px 14px rgba(201,168,76,.35)}
.btn:hover{box-shadow:0 6px 20px rgba(201,168,76,.5)}
.fase{background:#fff8ec;border:1px solid rgba(201,168,76,.4);border-radius:10px;padding:7px 11px;font-size:.73rem;color:#7a5a1a;text-align:center;margin-top:12px}
.olho{position:absolute;right:11px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:#6b7a99;font-size:.95rem;padding:0}
.senha-wrap{position:relative}
</style></head><body>
<div class="card">
  <div class="hdr">
    <div class="sigla">PAAC</div>
    <div class="sub">Plataforma de Automacao e Auditoria Cartorial</div>
  </div>
  <div class="body">
    <div class="aviso">
      <strong>Bem-vindo a PAAC</strong>
      Informe seu e-mail e senha para acessar a plataforma.
    </div>
    ${msgErro ? '<div class="erro">' + msgErro + '</div>' : ''}
    <form method="POST" action="/login">
      <div class="campo">
        <label>E-mail</label>
        <input type="email" name="email" placeholder="seu@email.com" required autofocus>
      </div>
      <div class="campo">
        <label>Senha</label>
        <div class="senha-wrap">
          <input type="password" name="senha" id="senhaInput" placeholder="Sua senha" required>
          <button type="button" class="olho" onclick="var i=document.getElementById('senhaInput');i.type=i.type==='password'?'text':'password'">&#128065;</button>
        </div>
      </div>
      <button type="submit" class="btn">Entrar na Plataforma</button>
    </form>
    <div class="fase">Fase E0 — Acesso para testes e validacao.</div>
  </div>
</div>
</body></html>`;
}

// ═══════════════════════════════════════════════════════════════════
// HTML — TELA DE ESCOLHA (só aparece para admins após login)
// ═══════════════════════════════════════════════════════════════════
function htmlEscolhaAdmin(nome) {
  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PAAC — Escolha de acesso</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:linear-gradient(135deg,#0d1f3c 0%,#1a3a6b 60%,#0a1628 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.4);width:100%;max-width:500px;overflow:hidden}
.hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);border-bottom:3px solid #c9a84c;padding:22px 28px;text-align:center}
.sigla{font-family:'Cormorant Garamond',serif;font-size:2.2rem;font-weight:700;color:#c9a84c;letter-spacing:4px}
.sub{color:rgba(255,255,255,.7);font-size:.75rem;margin-top:4px}
.body{padding:28px}
.saudacao{font-size:.88rem;color:#0d1f3c;margin-bottom:20px;line-height:1.5}
.saudacao strong{color:#1a3a6b}
.opcoes{display:flex;flex-direction:column;gap:14px}
.opcao{display:block;text-decoration:none;border:2px solid #dde2ec;border-radius:12px;padding:18px 20px;transition:all .2s;cursor:pointer}
.opcao:hover{border-color:#2a5298;background:#f4f6fb}
.opcao-titulo{font-weight:700;font-size:.95rem;color:#0d1f3c;margin-bottom:4px}
.opcao-sub{font-size:.76rem;color:#6b7a99;line-height:1.4}
.opcao-admin{border-color:rgba(192,57,43,.25)}
.opcao-admin:hover{border-color:#c0392b;background:#fdf8f8}
.opcao-admin .opcao-titulo{color:#8b1a1a}
.rodape{display:flex;justify-content:flex-end;margin-top:18px}
.btn-sair{background:none;border:1px solid #dde2ec;color:#6b7a99;padding:6px 14px;border-radius:8px;font-size:.76rem;cursor:pointer;font-family:'DM Sans',sans-serif;text-decoration:none}
.btn-sair:hover{border-color:#c0392b;color:#c0392b}
</style></head><body>
<div class="card">
  <div class="hdr">
    <div class="sigla">PAAC</div>
    <div class="sub">Plataforma de Automacao e Auditoria Cartorial</div>
  </div>
  <div class="body">
    <div class="saudacao">Ola, <strong>${nome}</strong>. Escolha como deseja acessar:</div>
    <div class="opcoes">
      <a href="/formulario" class="opcao">
        <div class="opcao-titulo">Acessar Formulario</div>
        <div class="opcao-sub">Preencher dados e gerar documentos de incorporacao como qualquer usuario da plataforma.</div>
      </a>
      <a href="/painel" class="opcao opcao-admin">
        <div class="opcao-titulo">Painel Administrativo</div>
        <div class="opcao-sub">Gerenciar usuarios, visualizar dados e configurar a plataforma. Requer senha adicional.</div>
      </a>
    </div>
    <div class="rodape"><a href="/sair" class="btn-sair">Sair</a></div>
  </div>
</div>
</body></html>`;
}

// ═══════════════════════════════════════════════════════════════════
// HTML — TELA DA SEGUNDA SENHA (acesso ao painel admin)
// ═══════════════════════════════════════════════════════════════════
function htmlSegundaSenha(msgErro) {
  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PAAC — Autenticacao Admin</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:linear-gradient(135deg,#1a0a0a 0%,#3d1515 60%,#0a0808 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.5);width:100%;max-width:420px;overflow:hidden}
.hdr{background:linear-gradient(135deg,#3d1515,#6b1a1a);border-bottom:3px solid #c9a84c;padding:20px 28px;text-align:center}
.sigla{font-family:'Cormorant Garamond',serif;font-size:1.8rem;font-weight:700;color:#c9a84c;letter-spacing:3px}
.sub{color:rgba(255,255,255,.7);font-size:.73rem;margin-top:4px}
.body{padding:24px 28px}
.aviso{background:#fdf2f2;border:1px solid rgba(192,57,43,.25);border-radius:10px;padding:10px 13px;font-size:.76rem;color:#7a3a3a;margin-bottom:16px;line-height:1.5}
.erro{background:#fdf2f2;border:1px solid rgba(192,57,43,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#c0392b;margin-bottom:12px}
.campo{display:flex;flex-direction:column;gap:4px;margin-bottom:14px}
.campo label{font-size:.77rem;font-weight:600;color:#3d1515}
.campo input{border:1.5px solid #dde2ec;border-radius:10px;padding:10px 13px;font-size:.87rem;font-family:'DM Sans',sans-serif;outline:none}
.campo input:focus{border-color:#c0392b;box-shadow:0 0 0 3px rgba(192,57,43,.1)}
.btn{width:100%;background:linear-gradient(135deg,#c0392b,#8b1a1a);color:#fff;border:none;border-radius:10px;padding:12px;font-family:'DM Sans',sans-serif;font-size:.9rem;font-weight:700;cursor:pointer}
.btn:hover{box-shadow:0 4px 16px rgba(192,57,43,.4)}
.voltar{display:block;text-align:center;margin-top:12px;font-size:.76rem;color:#6b7a99;text-decoration:none}
.voltar:hover{color:#c0392b}
</style></head><body>
<div class="card">
  <div class="hdr">
    <div class="sigla">PAAC ADMIN</div>
    <div class="sub">Verificacao de seguranca adicional</div>
  </div>
  <div class="body">
    <div class="aviso">Para acessar o painel administrativo, informe a senha exclusiva de administrador.</div>
    ${msgErro ? '<div class="erro">' + msgErro + '</div>' : ''}
    <form method="POST" action="/painel/auth">
      <div class="campo">
        <label>Senha do Painel Administrativo</label>
        <input type="password" name="senha_admin" placeholder="Senha exclusiva do admin" required autofocus>
      </div>
      <button type="submit" class="btn">Acessar Painel Admin</button>
    </form>
    <a href="/" class="voltar">Voltar</a>
  </div>
</div>
</body></html>`;
}

// ═══════════════════════════════════════════════════════════════════
// HTML — PAINEL ADMINISTRATIVO COMPLETO
// ═══════════════════════════════════════════════════════════════════
function htmlPainelAdmin(adminLogado) {
  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PAAC — Painel Admin</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:#f4f6fb;color:#0d1f3c;min-height:100vh}
.hdr{background:linear-gradient(135deg,#3d1515,#6b1a1a);border-bottom:3px solid #c9a84c;padding:14px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.hdr-left{display:flex;align-items:center;gap:12px}
.hdr-sigla{font-family:'Cormorant Garamond',serif;font-size:1.5rem;font-weight:700;color:#c9a84c;letter-spacing:3px}
.hdr-sub{color:rgba(255,255,255,.7);font-size:.72rem}
.hdr-right{display:flex;align-items:center;gap:10px}
.badge{background:rgba(192,57,43,.3);border:1px solid rgba(192,57,43,.5);color:#ffb3b3;padding:3px 10px;border-radius:20px;font-size:.65rem;letter-spacing:1px}
.btn-sair{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.7);padding:5px 12px;border-radius:6px;font-size:.72rem;cursor:pointer;font-family:'DM Sans',sans-serif;text-decoration:none}
.btn-sair:hover{border-color:#fff;color:#fff}
.btn-voltar{background:rgba(201,168,76,.15);border:1px solid rgba(201,168,76,.4);color:#c9a84c;padding:5px 12px;border-radius:6px;font-size:.72rem;cursor:pointer;font-family:'DM Sans',sans-serif;text-decoration:none}
.main{max-width:860px;margin:0 auto;padding:24px 16px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
.stat{background:#fff;border-radius:12px;padding:14px 16px;box-shadow:0 2px 12px rgba(0,0,0,.06);text-align:center}
.stat-num{font-size:1.5rem;font-weight:700;color:#0d1f3c}
.stat-label{font-size:.68rem;color:#6b7a99;margin-top:2px}
.card{background:#fff;border-radius:16px;box-shadow:0 4px 24px rgba(13,31,60,.1);margin-bottom:20px;overflow:hidden}
.card-hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);padding:14px 22px;display:flex;align-items:center;gap:10px}
.card-icon{font-size:1rem}
.card-title{font-family:'Cormorant Garamond',serif;font-size:1.1rem;font-weight:600;color:#fff}
.card-body{padding:18px 22px}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px}
.form-row.tri{grid-template-columns:1fr 1fr 1fr}
.form-row.full{grid-template-columns:1fr}
.campo{display:flex;flex-direction:column;gap:3px}
.campo label{font-size:.73rem;font-weight:600}
.campo input,.campo select{border:1.5px solid #dde2ec;border-radius:8px;padding:8px 11px;font-size:.84rem;font-family:'DM Sans',sans-serif;outline:none}
.campo input:focus,.campo select:focus{border-color:#2a5298;box-shadow:0 0 0 3px rgba(42,82,152,.1)}
.campo .dica{font-size:.67rem;color:#6b7a99}
.btn-add{background:linear-gradient(135deg,#c9a84c,#a87d2a);color:#0d1f3c;border:none;border-radius:10px;padding:10px 24px;font-family:'DM Sans',sans-serif;font-size:.86rem;font-weight:700;cursor:pointer;width:100%;margin-top:6px}
.btn-add:hover{box-shadow:0 4px 14px rgba(201,168,76,.4)}
.msg{padding:8px 12px;border-radius:8px;font-size:.78rem;margin-bottom:10px;display:none}
.msg.ok{display:block;background:#e8f7ef;border:1px solid rgba(26,122,74,.3);color:#1a7a4a}
.msg.erro{display:block;background:#fdf2f2;border:1px solid rgba(192,57,43,.3);color:#c0392b}
.tabela{width:100%;border-collapse:collapse;font-size:.8rem}
.tabela th{background:#0d1f3c;color:#fff;padding:9px 10px;text-align:left;font-weight:500;font-size:.7rem;letter-spacing:.5px}
.tabela td{padding:8px 10px;border-bottom:1px solid #dde2ec;vertical-align:middle}
.tabela tr:hover td{background:#f4f6fb}
.status{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.68rem;font-weight:600}
.status.ativo{background:#e8f7ef;color:#1a7a4a}
.status.inativo{background:#fdf2f2;color:#c0392b}
.papel-tag{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.68rem;font-weight:600}
.papel-tag.admin{background:#fdf2f2;color:#8b1a1a;border:1px solid rgba(139,26,26,.2)}
.papel-tag.usuario{background:#ebf3fb;color:#2a5298;border:1px solid rgba(42,82,152,.2)}
.btn-sm{border:none;border-radius:6px;padding:4px 10px;font-size:.7rem;cursor:pointer;font-family:'DM Sans',sans-serif;font-weight:500}
.btn-toggle{background:#ebf3fb;color:#2a5298;border:1px solid rgba(42,82,152,.2)}
.btn-del{background:#fdf2f2;color:#c0392b;border:1px solid rgba(192,57,43,.2)}
.btn-sm:hover{opacity:.8}
.btn-sm:disabled{opacity:.4;cursor:not-allowed}
.vazio{text-align:center;color:#6b7a99;padding:24px;font-size:.84rem}
.protegido-tag{font-size:.6rem;color:#b7701a;background:#fff8ec;padding:1px 5px;border-radius:4px;margin-left:4px}
@media(max-width:640px){.stats{grid-template-columns:1fr 1fr}.form-row,.form-row.tri{grid-template-columns:1fr}}
</style></head><body>

<div class="hdr">
  <div class="hdr-left">
    <div class="hdr-sigla">PAAC</div>
    <div class="hdr-sub">Painel Administrativo — ${adminLogado.nome}</div>
  </div>
  <div class="hdr-right">
    <span class="badge">ADMIN E0</span>
    <a href="/" class="btn-voltar">Voltar a plataforma</a>
    <a href="/sair" class="btn-sair">Sair</a>
  </div>
</div>

<div class="main">

  <div class="stats">
    <div class="stat"><div class="stat-num" id="stTotal">0</div><div class="stat-label">Total</div></div>
    <div class="stat"><div class="stat-num" id="stAtivos">0</div><div class="stat-label">Ativos</div></div>
    <div class="stat"><div class="stat-num" id="stAdmins">0</div><div class="stat-label">Admins</div></div>
    <div class="stat"><div class="stat-num" id="stInativos">0</div><div class="stat-label">Inativos</div></div>
  </div>

  <!-- CADASTRAR NOVO -->
  <div class="card">
    <div class="card-hdr"><span class="card-icon">+</span><span class="card-title">Cadastrar Novo Usuario</span></div>
    <div class="card-body">
      <div class="msg" id="msgCad"></div>
      <div class="form-row">
        <div class="campo"><label>Nome Completo *</label><input type="text" id="cNome" placeholder="Ex: Joao da Silva"></div>
        <div class="campo"><label>E-mail *</label><input type="email" id="cEmail" placeholder="joao@email.com"></div>
      </div>
      <div class="form-row tri">
        <div class="campo"><label>Senha inicial *</label><input type="password" id="cSenha" placeholder="Minimo 6 caracteres"></div>
        <div class="campo"><label>Papel</label>
          <select id="cPapel"><option value="usuario">Usuario comum</option><option value="admin">Administrador</option></select>
          <span class="dica">Admin: maximo 2 no sistema. Protecao mutua automatica.</span>
        </div>
        <div class="campo"><label>Observacao</label><input type="text" id="cObs" placeholder="Opcional"></div>
      </div>
      <button class="btn-add" onclick="cadastrar()">Cadastrar</button>
    </div>
  </div>

  <!-- TROCAR MINHA SENHA -->
  <div class="card">
    <div class="card-hdr"><span class="card-icon">&#128274;</span><span class="card-title">Alterar Minha Senha de Login</span></div>
    <div class="card-body">
      <div class="msg" id="msgSenha"></div>
      <div class="form-row tri">
        <div class="campo"><label>Senha atual *</label><input type="password" id="sAtual"></div>
        <div class="campo"><label>Nova senha *</label><input type="password" id="sNova" placeholder="Minimo 6 caracteres"></div>
        <div class="campo" style="justify-content:flex-end"><button class="btn-add" style="margin-top:0" onclick="trocarSenha()">Alterar</button></div>
      </div>
    </div>
  </div>

  <!-- LISTA -->
  <div class="card">
    <div class="card-hdr"><span class="card-icon">&#128101;</span><span class="card-title">Usuarios Cadastrados</span></div>
    <div class="card-body" style="padding:0;overflow-x:auto">
      <table class="tabela">
        <thead><tr><th>Nome</th><th>E-mail</th><th>Papel</th><th>Status</th><th>Ultimo acesso</th><th>Acoes</th></tr></thead>
        <tbody id="corpo"><tr><td colspan="6" class="vazio">Carregando...</td></tr></tbody>
      </table>
    </div>
  </div>

</div>

<script>
var MEU_EMAIL = '${adminLogado.email}';

async function carregar() {
  var r = await fetch('/painel/api/usuarios');
  var lista = await r.json();
  document.getElementById('stTotal').textContent = lista.length;
  document.getElementById('stAtivos').textContent = lista.filter(function(u){return u.ativo}).length;
  document.getElementById('stAdmins').textContent = lista.filter(function(u){return u.papel==='admin'}).length;
  document.getElementById('stInativos').textContent = lista.filter(function(u){return !u.ativo}).length;

  var corpo = document.getElementById('corpo');
  if (lista.length === 0) { corpo.innerHTML = '<tr><td colspan="6" class="vazio">Nenhum usuario cadastrado.</td></tr>'; return; }

  corpo.innerHTML = lista.map(function(u) {
    var acesso = u.ultimoAcesso ? new Date(u.ultimoAcesso).toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit'}) : '—';
    var ehAdmin = u.papel === 'admin';
    var ehEu = u.email.toLowerCase() === MEU_EMAIL.toLowerCase();
    var podeAgir = !ehAdmin;
    var protTag = u.protegido ? '<span class="protegido-tag">protegido</span>' : '';
    return '<tr>' +
      '<td><strong>' + u.nome + '</strong>' + protTag + (u.observacao ? '<br><small style="color:#6b7a99">' + u.observacao + '</small>' : '') + '</td>' +
      '<td>' + u.email + (ehEu ? ' <small style="color:#c9a84c">(voce)</small>' : '') + '</td>' +
      '<td><span class="papel-tag ' + u.papel + '">' + (ehAdmin ? 'Admin' : 'Usuario') + '</span></td>' +
      '<td><span class="status ' + (u.ativo ? 'ativo' : 'inativo') + '">' + (u.ativo ? 'Ativo' : 'Inativo') + '</span></td>' +
      '<td style="font-size:.74rem;color:#6b7a99">' + acesso + '</td>' +
      '<td style="white-space:nowrap">' +
        (podeAgir
          ? '<button class="btn-sm btn-toggle" onclick="toggle(\\'' + u.email + '\\')">' + (u.ativo ? 'Desativar' : 'Ativar') + '</button> ' +
            '<button class="btn-sm btn-del" onclick="remover(\\'' + u.email + '\\')">Remover</button>'
          : '<span style="font-size:.68rem;color:#b7701a">Admin protegido</span>')
      + '</td></tr>';
  }).join('');
}

async function cadastrar() {
  var nome = document.getElementById('cNome').value.trim();
  var email = document.getElementById('cEmail').value.trim();
  var senha = document.getElementById('cSenha').value;
  var papel = document.getElementById('cPapel').value;
  var obs = document.getElementById('cObs').value.trim();
  var msg = document.getElementById('msgCad');
  if (!nome || !email || !senha) { msg.className='msg erro'; msg.textContent='Nome, e-mail e senha sao obrigatorios.'; return; }
  var r = await fetch('/painel/api/usuarios', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({email:email,nome:nome,senha:senha,papel:papel,observacao:obs})});
  var res = await r.json();
  if (res.ok) {
    msg.className='msg ok'; msg.textContent=res.mensagem;
    document.getElementById('cNome').value=''; document.getElementById('cEmail').value='';
    document.getElementById('cSenha').value=''; document.getElementById('cObs').value='';
    document.getElementById('cPapel').value='usuario';
    carregar();
  } else { msg.className='msg erro'; msg.textContent=res.erro; }
}

async function toggle(email) {
  var r = await fetch('/painel/api/usuarios/' + encodeURIComponent(email) + '/toggle', {method:'PUT'});
  var res = await r.json();
  if (res.erro) { alert(res.erro); } else { carregar(); }
}

async function remover(email) {
  if (!confirm('Tem certeza que deseja REMOVER o acesso de ' + email + '?')) return;
  var r = await fetch('/painel/api/usuarios/' + encodeURIComponent(email), {method:'DELETE'});
  var res = await r.json();
  if (res.erro) { alert(res.erro); } else { carregar(); }
}

async function trocarSenha() {
  var atual = document.getElementById('sAtual').value;
  var nova = document.getElementById('sNova').value;
  var msg = document.getElementById('msgSenha');
  if (!atual || !nova) { msg.className='msg erro'; msg.textContent='Informe a senha atual e a nova.'; return; }
  var r = await fetch('/painel/api/minha-senha', {method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify({senhaAtual:atual,senhaNova:nova})});
  var res = await r.json();
  if (res.ok) { msg.className='msg ok'; msg.textContent=res.mensagem; document.getElementById('sAtual').value=''; document.getElementById('sNova').value=''; }
  else { msg.className='msg erro'; msg.textContent=res.erro; }
}

carregar();
</script>
</body></html>`;
}

// ═══════════════════════════════════════════════════════════════════
// INICIAR O SERVIDOR
// ═══════════════════════════════════════════════════════════════════
app.listen(PORT, '0.0.0.0', () => {
  console.log('');
  console.log('=====================================================');
  console.log('  PAAC v2 — Servidor rodando na porta ' + PORT);
  console.log('  Login:           http://localhost:' + PORT);
  console.log('  Painel Admin:    http://localhost:' + PORT + '/painel');
  console.log('  (acessivel apenas apos login com conta admin)');
  console.log('=====================================================');
  console.log('');
});
