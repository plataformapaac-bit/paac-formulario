// BLOCO DE INICIALIZACAO
const fs_init = require('fs');
const path_init = require('path');
const arquivoUsuarios = path_init.join(__dirname, 'usuarios.json');
if (process.env.USUARIOS_JSON) {
  try { JSON.parse(process.env.USUARIOS_JSON); fs_init.writeFileSync(arquivoUsuarios, process.env.USUARIOS_JSON, 'utf-8'); console.log('usuarios.json criado.'); }
  catch (e) { console.error('ERRO JSON:', e.message); }
} else { console.log('AVISO: USUARIOS_JSON nao encontrada.'); }

const express = require('express');
const bcrypt = require('bcryptjs');
const cookieParser = require('cookie-parser');
const fs = require('fs');
const path = require('path');
const app = express();

// ═══ SUPABASE ═══
const { createClient } = require('@supabase/supabase-js');
const SUPABASE_URL = process.env.SUPABASE_URL || '';
const SUPABASE_KEY = process.env.SUPABASE_KEY || '';
var supabase = null;
if (SUPABASE_URL && SUPABASE_KEY) {
  supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
  console.log('Supabase conectado: ' + SUPABASE_URL);
} else {
  console.log('AVISO: SUPABASE_URL ou SUPABASE_KEY nao configuradas — projetos desativados.');
}

const PORT = process.env.PORT || 3000;
const COOKIE_SECRET = process.env.COOKIE_SECRET || 'paac-e0-temp';
const tentativasLogin = {};
const MAX_TENTATIVAS = 5;
const TEMPO_BLOQUEIO = 15 * 60000;
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser(COOKIE_SECRET));

function lerDados() { try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'usuarios.json'), 'utf-8')); } catch(e) { return { senhaAdminHash:'', senhaMestraHash:'', usuarios:[] }; } }
function salvarDados(d) { fs.writeFileSync(path.join(__dirname, 'usuarios.json'), JSON.stringify(d,null,2), 'utf-8'); }
function buscarUsuario(email) { return lerDados().usuarios.find(u => u.email.toLowerCase() === email.toLowerCase().trim()); }
function verificarBloqueio(email) { var r=tentativasLogin[email.toLowerCase()]; if(!r)return false; if(r.bloqueadoAte&&Date.now()<r.bloqueadoAte)return Math.ceil((r.bloqueadoAte-Date.now())/60000); if(r.bloqueadoAte)delete tentativasLogin[email.toLowerCase()]; return false; }
function registrarFalha(email) { var c=email.toLowerCase(); if(!tentativasLogin[c])tentativasLogin[c]={tentativas:0}; tentativasLogin[c].tentativas++; if(tentativasLogin[c].tentativas>=MAX_TENTATIVAS)tentativasLogin[c].bloqueadoAte=Date.now()+TEMPO_BLOQUEIO; }
function limparTentativas(email) { delete tentativasLogin[email.toLowerCase()]; }
function lerSessao(req) { try { var c=req.signedCookies.paac_sessao; return c?JSON.parse(c):null; } catch(e){return null;} }

var fonts = '<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:wght@400;500;600&display=swap" rel="stylesheet">';
var olhoScript = '<script>document.addEventListener("click",function(e){var b=e.target.closest("[data-toggle-pw]");if(!b)return;var id=b.getAttribute("data-toggle-pw");var i=document.getElementById(id);if(i)i.type=i.type==="password"?"text":"password"});<\/script>';

function senhaInput(id, name, ph, req) {
  return '<div style="position:relative"><input type="password" name="' + name + '" id="' + id + '" placeholder="' + ph + '" ' + (req ? 'required' : '') + ' style="padding-right:40px;width:100%"><button type="button" data-toggle-pw="' + id + '" style="position:absolute;right:11px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:#6b7a99;font-size:1.1rem">&#128065;</button></div>';
}

// ═══ ROTAS ═══
app.get('/', function(req,res) {
  var s=lerSessao(req); if(!s||!s.email)return res.send(htmlLogin());
  var u=buscarUsuario(s.email); if(!u||!u.ativo){res.clearCookie('paac_sessao');return res.send(htmlLogin('Conta desativada.'));}
  if(u.papel==='admin')return res.send(htmlEscolhaAdmin(u.nome));
  res.sendFile(path.join(__dirname,'index.html'));
});

app.post('/login', async function(req,res) {
  var email=(req.body.email||'').trim(), senha=req.body.senha||'';
  if(!email)return res.send(htmlLogin('Informe seu e-mail.'));
  var bloq=verificarBloqueio(email); if(bloq)return res.send(htmlLogin('Bloqueado. Aguarde '+bloq+' min.'));
  var u=buscarUsuario(email); if(!u){registrarFalha(email);return res.send(htmlLogin('E-mail ou senha incorretos.'));}
  if(!u.ativo)return res.send(htmlLogin('Conta desativada.'));
  if(!u.senhaHash||u.senhaHash===''){if(senha!=='')return res.send(htmlLogin('Primeiro acesso: deixe senha em branco.')); return res.send(htmlCriarSenha(email));}
  if(!senha)return res.send(htmlLogin('Informe sua senha.'));
  var ok=await bcrypt.compare(senha,u.senhaHash);
  if(!ok){registrarFalha(email);var rr=tentativasLogin[email.toLowerCase()];var rest=MAX_TENTATIVAS-(rr?rr.tentativas:0);return res.send(htmlLogin('E-mail ou senha incorretos.'+(rest>0?' ('+rest+' tentativas)':'')));}
  limparTentativas(email); var d=lerDados(); var idx=d.usuarios.findIndex(x => x.email.toLowerCase()===email.toLowerCase()); if(idx>=0){d.usuarios[idx].ultimoAcesso=new Date().toISOString();salvarDados(d);}
  res.cookie('paac_sessao',JSON.stringify({email:u.email,nome:u.nome,papel:u.papel}),{signed:true,httpOnly:true,maxAge:24*60*60*1000,sameSite:'lax'});
  res.redirect('/');
});

app.post('/criar-senha', async function(req,res) {
  var email=(req.body.email||'').trim(),senha=req.body.senha||'',conf=req.body.confirmar||'';
  if(!email||!senha)return res.send(htmlCriarSenha(email,'Preencha todos os campos.'));
  if(senha.length<6)return res.send(htmlCriarSenha(email,'Minimo 6 caracteres.'));
  if(senha!==conf)return res.send(htmlCriarSenha(email,'Senhas nao conferem.'));
  var d=lerDados();var idx=d.usuarios.findIndex(x => x.email.toLowerCase()===email.toLowerCase());
  if(idx<0)return res.send(htmlLogin('E-mail nao encontrado.'));
  if(d.usuarios[idx].senhaHash&&d.usuarios[idx].senhaHash!=='')return res.send(htmlLogin('Conta ja possui senha.'));
  d.usuarios[idx].senhaHash=await bcrypt.hash(senha,12);salvarDados(d);
  res.send(htmlLogin('Senha criada! Faca login.'));
});

app.get('/formulario', function(req,res) { var s=lerSessao(req);if(!s)return res.redirect('/');var u=buscarUsuario(s.email);if(!u||!u.ativo)return res.redirect('/');res.sendFile(path.join(__dirname,'index.html')); });

app.get('/minha-conta', function(req,res) { var s=lerSessao(req);if(!s)return res.redirect('/');var u=buscarUsuario(s.email);if(!u||!u.ativo)return res.redirect('/');res.send(htmlMinhaConta(u)); });
app.post('/minha-conta', async function(req,res) {
  var s=lerSessao(req);if(!s)return res.redirect('/');var u=buscarUsuario(s.email);if(!u)return res.redirect('/');
  var atual=req.body.senha_atual||'',nova=req.body.senha_nova||'',conf=req.body.confirmar||'';
  if(!atual||!nova)return res.send(htmlMinhaConta(u,'Preencha todos os campos.'));
  if(nova.length<6)return res.send(htmlMinhaConta(u,'Minimo 6 caracteres.'));
  if(nova!==conf)return res.send(htmlMinhaConta(u,'Senhas nao conferem.'));
  var ok=await bcrypt.compare(atual,u.senhaHash);if(!ok)return res.send(htmlMinhaConta(u,'Senha atual incorreta.'));
  var d=lerDados();var idx=d.usuarios.findIndex(x => x.email.toLowerCase()===s.email.toLowerCase());
  d.usuarios[idx].senhaHash=await bcrypt.hash(nova,12);salvarDados(d);
  res.send(htmlMinhaConta(u,null,'Senha alterada!'));
});

app.get('/sair', function(req,res) { res.clearCookie('paac_sessao');res.clearCookie('paac_admin');res.redirect('/'); });

app.get('/painel', function(req,res) {
  var s=lerSessao(req);if(!s)return res.redirect('/');var u=buscarUsuario(s.email);if(!u||u.papel!=='admin')return res.redirect('/');
  if(req.signedCookies.paac_admin==='autenticado')return res.send(htmlPainelAdmin(u));
  res.send(htmlSegundaSenha());
});
app.post('/painel/auth', async function(req,res) {
  var s=lerSessao(req);if(!s)return res.redirect('/');var u=buscarUsuario(s.email);if(!u||u.papel!=='admin')return res.redirect('/');
  var d=lerDados();var ok=await bcrypt.compare(req.body.senha_admin||'',d.senhaAdminHash);if(!ok)return res.send(htmlSegundaSenha('Senha incorreta.'));
  res.cookie('paac_admin','autenticado',{signed:true,httpOnly:true,maxAge:4*60*60*1000,sameSite:'lax'});res.redirect('/painel');
});

app.get('/emergencia', function(req,res) { res.send(htmlEmergencia()); });
app.post('/emergencia', async function(req,res) {
  var email=(req.body.email||'').trim(),sm=req.body.senha_mestra||'';
  if(!email||!sm)return res.send(htmlEmergencia('Preencha tudo.'));
  var d=lerDados();if(!d.senhaMestraHash)return res.send(htmlEmergencia('Senha-mestra nao configurada.'));
  var ok=await bcrypt.compare(sm,d.senhaMestraHash);if(!ok)return res.send(htmlEmergencia('Senha-mestra incorreta.'));
  var idx=d.usuarios.findIndex(x => x.email.toLowerCase()===email.toLowerCase());
  if(idx<0)return res.send(htmlEmergencia('E-mail nao encontrado.'));
  if(d.usuarios[idx].papel!=='admin')return res.send(htmlEmergencia('Somente para admins.'));
  d.usuarios[idx].senhaHash='';salvarDados(d);
  res.send(htmlEmergencia(null,'Senha resetada. Admin criara nova no proximo acesso.'));
});

// ═══ API ADMIN ═══
function exigirAdmin(req,res,next) {
  var s=lerSessao(req);if(!s)return res.status(401).json({erro:'Sessao expirada.'});
  var u=buscarUsuario(s.email);if(!u||u.papel!=='admin')return res.status(403).json({erro:'Sem permissao.'});
  if(req.signedCookies.paac_admin!=='autenticado')return res.status(401).json({erro:'Segunda senha.'});
  req.adminEmail=s.email.toLowerCase();next();
}
app.get('/painel/api/usuarios',exigirAdmin,function(req,res){var d=lerDados();res.json(d.usuarios.map(u=>({email:u.email,nome:u.nome,papel:u.papel,observacao:u.observacao||'',ativo:u.ativo,protegido:u.protegido||false,temSenha:!!(u.senhaHash&&u.senhaHash!==''),cadastradoEm:u.cadastradoEm,ultimoAcesso:u.ultimoAcesso})));});
app.post('/painel/api/usuarios',exigirAdmin,async function(req,res){
  var {email:e,nome:n,senha:s,observacao:o,papel:p}=req.body;
  if(!e||!n)return res.status(400).json({erro:'E-mail e nome obrigatorios.'});
  var d=lerDados();var el=e.toLowerCase().trim();
  if(d.usuarios.some(u=>u.email.toLowerCase()===el))return res.status(409).json({erro:'E-mail ja cadastrado.'});
  var pf=(p==='admin')?'admin':'usuario';
  if(pf==='admin'&&d.usuarios.filter(u=>u.papel==='admin').length>=3)return res.status(400).json({erro:'Max 3 admins.'});
  var sh='';if(s&&s.length>=6)sh=await bcrypt.hash(s,12);else if(s&&s.length>0&&s.length<6)return res.status(400).json({erro:'Senha: min 6 chars.'});
  d.usuarios.push({email:el,nome:n.trim(),senhaHash:sh,papel:pf,observacao:(o||'').trim(),ativo:true,protegido:pf==='admin',cadastradoEm:new Date().toISOString(),ultimoAcesso:null});
  salvarDados(d);var msg=pf==='admin'?'Admin cadastrado.':'Usuario cadastrado.';if(!sh)msg+=' Criara senha no 1o acesso.';
  res.json({ok:true,mensagem:msg});
});
app.put('/painel/api/usuarios/:email/toggle',exigirAdmin,function(req,res){var d=lerDados();var ea=decodeURIComponent(req.params.email).toLowerCase();var idx=d.usuarios.findIndex(u=>u.email.toLowerCase()===ea);if(idx<0)return res.status(404).json({erro:'Nao encontrado.'});if(d.usuarios[idx].papel==='admin')return res.status(403).json({erro:'Nao pode desativar admin.'});d.usuarios[idx].ativo=!d.usuarios[idx].ativo;salvarDados(d);res.json({ok:true,ativo:d.usuarios[idx].ativo});});
app.delete('/painel/api/usuarios/:email',exigirAdmin,function(req,res){var d=lerDados();var ea=decodeURIComponent(req.params.email).toLowerCase();var idx=d.usuarios.findIndex(u=>u.email.toLowerCase()===ea);if(idx<0)return res.status(404).json({erro:'Nao encontrado.'});if(d.usuarios[idx].papel==='admin')return res.status(403).json({erro:'Nao pode remover admin.'});d.usuarios.splice(idx,1);salvarDados(d);res.json({ok:true});});
app.put('/painel/api/usuarios/:email/reset-senha',exigirAdmin,function(req,res){var d=lerDados();var ea=decodeURIComponent(req.params.email).toLowerCase();var idx=d.usuarios.findIndex(u=>u.email.toLowerCase()===ea);if(idx<0)return res.status(404).json({erro:'Nao encontrado.'});if(ea===req.adminEmail)return res.status(400).json({erro:'Use Alterar Minha Senha.'});d.usuarios[idx].senhaHash='';salvarDados(d);res.json({ok:true,mensagem:'Senha resetada.'});});
app.put('/painel/api/minha-senha',exigirAdmin,async function(req,res){var{senhaAtual:sa,senhaNova:sn}=req.body;if(!sa||!sn)return res.status(400).json({erro:'Preencha ambos.'});if(sn.length<6)return res.status(400).json({erro:'Min 6 chars.'});var d=lerDados();var idx=d.usuarios.findIndex(u=>u.email.toLowerCase()===req.adminEmail);if(idx<0)return res.status(404).json({erro:'Nao encontrado.'});var ok=await bcrypt.compare(sa,d.usuarios[idx].senhaHash);if(!ok)return res.status(401).json({erro:'Senha incorreta.'});d.usuarios[idx].senhaHash=await bcrypt.hash(sn,12);salvarDados(d);res.json({ok:true,mensagem:'Senha alterada.'});});
app.get('/painel/admin.js', exigirAdmin, function(req, res) { res.sendFile(path.join(__dirname, 'painel-admin.js')); });

// ═══ HTML FUNCTIONS ═══
var cssBase = '*{box-sizing:border-box;margin:0;padding:0}body{font-family:"DM Sans",sans-serif}';
var cssBgAzul = 'body{background:linear-gradient(135deg,#0d1f3c 0%,#1a3a6b 60%,#0a1628 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}';
var cssBgVerm = 'body{background:linear-gradient(135deg,#1a0a0a 0%,#3d1515 60%,#0a0808 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}';
var cssCard = '.card{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.4);width:100%;max-width:500px;overflow:hidden}';
var cssHdrAzul = '.hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);border-bottom:3px solid #c9a84c;padding:22px 28px;text-align:center}.sigla{font-family:"Cormorant Garamond",serif;font-size:2.8rem;font-weight:700;color:#c9a84c;letter-spacing:5px}.sub{color:rgba(255,255,255,.7);font-size:.75rem;margin-top:4px}';
var cssHdrVerm = '.hdr{background:linear-gradient(135deg,#3d1515,#6b1a1a);border-bottom:3px solid #c9a84c;padding:20px 28px;text-align:center}.sigla{font-family:"Cormorant Garamond",serif;font-size:1.8rem;font-weight:700;color:#c9a84c;letter-spacing:3px}.sub{color:rgba(255,255,255,.7);font-size:.73rem;margin-top:4px}';
var cssCampo = '.campo{display:flex;flex-direction:column;gap:4px;margin-bottom:12px}.campo label{font-size:.77rem;font-weight:600;color:#0d1f3c}.campo input{border:1.5px solid #dde2ec;border-radius:10px;padding:10px 13px;font-size:.87rem;font-family:"DM Sans",sans-serif;outline:none;width:100%}.campo input:focus{border-color:#2a5298;box-shadow:0 0 0 3px rgba(42,82,152,.1)}';
var cssBtn = '.btn{width:100%;background:linear-gradient(135deg,#c9a84c,#a87d2a);color:#0d1f3c;border:none;border-radius:10px;padding:12px;font-family:"DM Sans",sans-serif;font-size:.93rem;font-weight:700;cursor:pointer;margin-top:4px}.btn:hover{box-shadow:0 6px 20px rgba(201,168,76,.5)}';
var cssBtnVerm = '.btn{width:100%;background:linear-gradient(135deg,#c0392b,#8b1a1a);color:#fff;border:none;border-radius:10px;padding:12px;font-family:"DM Sans",sans-serif;font-size:.9rem;font-weight:700;cursor:pointer}';

function htmlLogin(msg) {
  var ehOk = msg && msg.indexOf('Senha criada') >= 0;
  var corMsg = ehOk ? 'background:#e8f7ef;border:1px solid rgba(26,122,74,.3);color:#1a7a4a' : 'background:#fdf2f2;border:1px solid rgba(192,57,43,.3);color:#c0392b';
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC</title>' + fonts + '<style>' + cssBase + cssBgAzul + cssCard + cssHdrAzul + cssCampo + cssBtn + '.body{padding:24px 28px}.aviso{background:#fff8ec;border:1px solid rgba(201,168,76,.4);border-radius:10px;padding:10px 13px;font-size:.74rem;color:#7a5a1a;margin-bottom:16px;line-height:1.5}.aviso strong{display:block;margin-bottom:3px;color:#5a3a00}.msg{border-radius:10px;padding:8px 12px;font-size:.76rem;margin-bottom:12px}.link{text-align:center;margin-top:10px}.link a{color:#2a5298;font-size:.74rem}</style></head><body><div class="card"><div class="hdr"><div class="sigla">PAAC</div><div class="sub">Plataforma de Automacao e Auditoria Cartorial</div></div><div class="body"><div class="aviso"><strong>PAAC — Plataforma de Automacao e Auditoria Cartorial</strong>Sistema especializado na geracao automatica de documentos de incorporacao imobiliaria, com base nas normas da ABNT NBR 12721, Lei 4.591/64 e Codigo Civil Brasileiro. Acesse com seu e-mail e senha cadastrados.</div>' + (msg ? '<div class="msg" style="' + corMsg + '">' + msg + '</div>' : '') + '<form method="POST" action="/login"><div class="campo"><label>E-mail</label><input type="email" name="email" placeholder="seu@email.com" required autofocus></div><div class="campo"><label>Senha</label>' + senhaInput('s1','senha','Sua senha (branco se 1o acesso)',false) + '<span style="font-size:.67rem;color:#6b7a99;margin-top:2px">Primeiro acesso? Deixe em branco.</span></div><button type="submit" class="btn">Entrar na Plataforma</button></form><div class="link"><a href="/emergencia">Esqueci minha senha (admin)</a></div></div></div>' + olhoScript + '</body></html>';
}
function htmlCriarSenha(email, msg) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC — Criar Senha</title>' + fonts + '<style>' + cssBase + cssBgAzul + cssCard + '.hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);border-bottom:3px solid #c9a84c;padding:22px 28px;text-align:center}.sigla{font-family:"Cormorant Garamond",serif;font-size:2.2rem;font-weight:700;color:#c9a84c;letter-spacing:4px}.sub{color:rgba(255,255,255,.7);font-size:.75rem;margin-top:4px}' + cssCampo + cssBtn + '.body{padding:24px 28px}.aviso{background:#e8f7ef;border:1px solid rgba(26,122,74,.3);border-radius:10px;padding:10px 13px;font-size:.76rem;color:#1a7a4a;margin-bottom:16px}.erro{background:#fdf2f2;border:1px solid rgba(192,57,43,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#c0392b;margin-bottom:12px}</style></head><body><div class="card"><div class="hdr"><div class="sigla">PAAC</div><div class="sub">Criar sua senha</div></div><div class="body"><div class="aviso">Primeiro acesso ou senha resetada. Crie sua senha.</div>' + (msg ? '<div class="erro">' + msg + '</div>' : '') + '<form method="POST" action="/criar-senha"><input type="hidden" name="email" value="' + email + '"><div class="campo"><label>E-mail</label><input type="email" value="' + email + '" disabled></div><div class="campo"><label>Nova Senha (min 6 caracteres)</label>' + senhaInput('cs1','senha','Crie sua senha',true) + '</div><div class="campo"><label>Confirmar</label>' + senhaInput('cs2','confirmar','Repita',true) + '</div><button type="submit" class="btn">Criar Senha</button></form></div></div>' + olhoScript + '</body></html>';
}
function htmlEscolhaAdmin(nome) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC</title>' + fonts + '<style>' + cssBase + cssBgAzul + '.card{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.4);width:100%;max-width:500px;overflow:hidden}.hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);border-bottom:3px solid #c9a84c;padding:22px 28px;text-align:center}.sigla{font-family:"Cormorant Garamond",serif;font-size:2.2rem;font-weight:700;color:#c9a84c;letter-spacing:4px}.sub{color:rgba(255,255,255,.7);font-size:.75rem;margin-top:4px}.body{padding:28px}.saudacao{font-size:.88rem;margin-bottom:20px}.saudacao strong{color:#1a3a6b}.opcoes{display:flex;flex-direction:column;gap:14px}.opcao{display:block;text-decoration:none;border:2px solid #dde2ec;border-radius:12px;padding:18px 20px;transition:all .2s}.opcao:hover{border-color:#2a5298;background:#f4f6fb}.opcao-titulo{font-weight:700;font-size:.95rem;color:#0d1f3c;margin-bottom:4px}.opcao-sub{font-size:.76rem;color:#6b7a99}.opcao-admin{border-color:rgba(192,57,43,.25)}.opcao-admin:hover{border-color:#c0392b;background:#fdf8f8}.opcao-admin .opcao-titulo{color:#8b1a1a}.rodape{display:flex;justify-content:flex-end;gap:10px;margin-top:18px}.rodape a{padding:6px 14px;border-radius:8px;font-size:.76rem;text-decoration:none;font-family:"DM Sans",sans-serif}.btn-conta{border:1px solid #2a5298;color:#2a5298}.btn-sair{border:1px solid #dde2ec;color:#6b7a99}</style></head><body><div class="card"><div class="hdr"><div class="sigla">PAAC</div><div class="sub">Plataforma de Automacao e Auditoria Cartorial</div></div><div class="body"><div class="saudacao">Ola, <strong>' + nome + '</strong>. Escolha como acessar:</div><div class="opcoes"><a href="/formulario" class="opcao"><div class="opcao-titulo">Acessar Formulario</div><div class="opcao-sub">Preencher dados e gerar documentos de incorporacao.</div></a><a href="/painel" class="opcao opcao-admin"><div class="opcao-titulo">Painel Administrativo</div><div class="opcao-sub">Gerenciar usuarios e configurar. Requer senha adicional.</div></a></div><div class="rodape"><a href="/minha-conta" class="btn-conta">Minha Conta</a><a href="/sair" class="btn-sair">Sair</a></div></div></div></body></html>';
}
function htmlSegundaSenha(msg) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC Admin</title>' + fonts + '<style>' + cssBase + cssBgVerm + cssCard + cssHdrVerm + cssCampo + cssBtnVerm + '.campo label{color:#3d1515}.campo input:focus{border-color:#c0392b;box-shadow:0 0 0 3px rgba(192,57,43,.1)}.body{padding:24px 28px}.aviso{background:#fdf2f2;border:1px solid rgba(192,57,43,.25);border-radius:10px;padding:10px 13px;font-size:.76rem;color:#7a3a3a;margin-bottom:16px}.erro{background:#fdf2f2;border:1px solid rgba(192,57,43,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#c0392b;margin-bottom:12px}.voltar{display:block;text-align:center;margin-top:12px;font-size:.76rem;color:#6b7a99;text-decoration:none}</style></head><body><div class="card"><div class="hdr"><div class="sigla">PAAC ADMIN</div><div class="sub">Verificacao adicional</div></div><div class="body"><div class="aviso">Informe a senha exclusiva do painel.</div>' + (msg ? '<div class="erro">' + msg + '</div>' : '') + '<form method="POST" action="/painel/auth"><div class="campo"><label>Senha do Painel</label>' + senhaInput('sa','senha_admin','Senha do admin',true) + '</div><button type="submit" class="btn">Acessar Painel</button></form><a href="/" class="voltar">Voltar</a></div></div>' + olhoScript + '</body></html>';
}
function htmlMinhaConta(usuario, msgErro, msgOk) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC — Minha Conta</title>' + fonts + '<style>' + cssBase + cssBgAzul + '.card{background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.4);width:100%;max-width:480px;overflow:hidden}.hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);border-bottom:3px solid #c9a84c;padding:20px 28px;text-align:center}.sigla{font-family:"Cormorant Garamond",serif;font-size:1.8rem;font-weight:700;color:#c9a84c;letter-spacing:3px}.sub{color:rgba(255,255,255,.7);font-size:.73rem;margin-top:4px}' + cssCampo + cssBtn + '.body{padding:24px 28px}.info{background:#ebf3fb;border:1px solid rgba(42,82,152,.2);border-radius:10px;padding:10px 13px;font-size:.78rem;color:#1a3a6b;margin-bottom:16px}.info strong{display:block;margin-bottom:2px}.erro{background:#fdf2f2;border:1px solid rgba(192,57,43,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#c0392b;margin-bottom:12px}.ok{background:#e8f7ef;border:1px solid rgba(26,122,74,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#1a7a4a;margin-bottom:12px}.voltar{display:block;text-align:center;margin-top:14px;font-size:.76rem;color:#6b7a99;text-decoration:none}</style></head><body><div class="card"><div class="hdr"><div class="sigla">PAAC</div><div class="sub">Minha Conta</div></div><div class="body"><div class="info"><strong>' + usuario.nome + '</strong>E-mail: ' + usuario.email + '<br>Papel: ' + (usuario.papel === 'admin' ? 'Administrador' : 'Usuario') + '</div>' + (msgErro ? '<div class="erro">' + msgErro + '</div>' : '') + (msgOk ? '<div class="ok">' + msgOk + '</div>' : '') + '<form method="POST" action="/minha-conta"><div class="campo"><label>Senha Atual</label>' + senhaInput('mc1','senha_atual','Senha atual',true) + '</div><div class="campo"><label>Nova Senha (min 6)</label>' + senhaInput('mc2','senha_nova','Nova senha',true) + '</div><div class="campo"><label>Confirmar</label>' + senhaInput('mc3','confirmar','Repita',true) + '</div><button type="submit" class="btn">Alterar Senha</button></form><a href="/" class="voltar">Voltar</a></div></div>' + olhoScript + '</body></html>';
}
function htmlEmergencia(msg, msgOk) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC — Emergencia</title>' + fonts + '<style>' + cssBase + cssBgVerm + cssCard + cssHdrVerm + cssCampo + cssBtnVerm + '.sigla{font-size:1.6rem}.campo label{color:#3d1515}.body{padding:24px 28px}.aviso{background:#fff8ec;border:1px solid rgba(201,168,76,.4);border-radius:10px;padding:10px 13px;font-size:.74rem;color:#7a5a1a;margin-bottom:16px;line-height:1.5}.aviso strong{display:block;margin-bottom:3px;color:#5a3a00}.erro{background:#fdf2f2;border:1px solid rgba(192,57,43,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#c0392b;margin-bottom:12px}.ok{background:#e8f7ef;border:1px solid rgba(26,122,74,.3);border-radius:10px;padding:8px 12px;font-size:.76rem;color:#1a7a4a;margin-bottom:12px}.voltar{display:block;text-align:center;margin-top:12px;font-size:.76rem;color:#6b7a99;text-decoration:none}</style></head><body><div class="card"><div class="hdr"><div class="sigla">EMERGENCIA</div><div class="sub">Reset de admin</div></div><div class="body"><div class="aviso"><strong>Recurso de emergencia</strong>Use a senha-mestra para resetar a senha de um admin.</div>' + (msg ? '<div class="erro">' + msg + '</div>' : '') + (msgOk ? '<div class="ok">' + msgOk + '</div>' : '') + '<form method="POST" action="/emergencia"><div class="campo"><label>E-mail do Admin</label><input type="email" name="email" placeholder="admin@email.com" required></div><div class="campo"><label>Senha-Mestra</label>' + senhaInput('sm','senha_mestra','Senha-mestra',true) + '</div><button type="submit" class="btn">Resetar Senha</button></form><a href="/" class="voltar">Voltar</a></div></div>' + olhoScript + '</body></html>';
}
function htmlPainelAdmin(adminLogado) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PAAC Painel</title>' + fonts + '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:"DM Sans",sans-serif;background:#f4f6fb;color:#0d1f3c;min-height:100vh}.hdr{background:linear-gradient(135deg,#3d1515,#6b1a1a);border-bottom:3px solid #c9a84c;padding:14px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}.hdr-left{display:flex;align-items:center;gap:12px}.hdr-sigla{font-family:"Cormorant Garamond",serif;font-size:1.5rem;font-weight:700;color:#c9a84c;letter-spacing:3px}.hdr-sub{color:rgba(255,255,255,.7);font-size:.72rem}.hdr-right{display:flex;align-items:center;gap:10px}.badge{background:rgba(192,57,43,.3);border:1px solid rgba(192,57,43,.5);color:#ffb3b3;padding:3px 10px;border-radius:20px;font-size:.65rem}.btn-sair,.btn-voltar{padding:5px 12px;border-radius:6px;font-size:.72rem;cursor:pointer;font-family:"DM Sans",sans-serif;text-decoration:none;border:1px solid rgba(255,255,255,.2)}.btn-sair{background:rgba(255,255,255,.1);color:rgba(255,255,255,.7)}.btn-voltar{background:rgba(201,168,76,.15);border-color:rgba(201,168,76,.4);color:#c9a84c}.main{max-width:860px;margin:0 auto;padding:24px 16px}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}.stat{background:#fff;border-radius:12px;padding:14px;box-shadow:0 2px 12px rgba(0,0,0,.06);text-align:center}.stat-num{font-size:1.5rem;font-weight:700}.stat-label{font-size:.68rem;color:#6b7a99;margin-top:2px}.card{background:#fff;border-radius:16px;box-shadow:0 4px 24px rgba(13,31,60,.1);margin-bottom:20px;overflow:hidden}.card-hdr{background:linear-gradient(135deg,#0d1f3c,#1a3a6b);padding:14px 22px;display:flex;align-items:center;gap:10px}.card-icon{font-size:1rem}.card-title{font-family:"Cormorant Garamond",serif;font-size:1.1rem;font-weight:600;color:#fff}.card-body{padding:18px 22px}.form-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px}.form-row.tri{grid-template-columns:1fr 1fr 1fr}.campo{display:flex;flex-direction:column;gap:3px}.campo label{font-size:.73rem;font-weight:600}.campo input,.campo select{border:1.5px solid #dde2ec;border-radius:8px;padding:8px 11px;font-size:.84rem;font-family:"DM Sans",sans-serif;outline:none;width:100%}.campo .dica{font-size:.67rem;color:#6b7a99}.btn-add{background:linear-gradient(135deg,#c9a84c,#a87d2a);color:#0d1f3c;border:none;border-radius:10px;padding:10px;font-family:"DM Sans",sans-serif;font-size:.86rem;font-weight:700;cursor:pointer;width:100%;margin-top:6px}.msg{padding:8px 12px;border-radius:8px;font-size:.78rem;margin-bottom:10px;display:none}.msg.ok{display:block;background:#e8f7ef;border:1px solid rgba(26,122,74,.3);color:#1a7a4a}.msg.erro{display:block;background:#fdf2f2;border:1px solid rgba(192,57,43,.3);color:#c0392b}.tabela{width:100%;border-collapse:collapse;font-size:.8rem}.tabela th{background:#0d1f3c;color:#fff;padding:9px 10px;text-align:left;font-size:.7rem}.tabela td{padding:8px 10px;border-bottom:1px solid #dde2ec;vertical-align:middle}.tabela tr:hover td{background:#f4f6fb}.status,.papel-tag{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.68rem;font-weight:600}.status.ativo{background:#e8f7ef;color:#1a7a4a}.status.inativo{background:#fdf2f2;color:#c0392b}.papel-tag.admin{background:#fdf2f2;color:#8b1a1a;border:1px solid rgba(139,26,26,.2)}.papel-tag.usuario{background:#ebf3fb;color:#2a5298;border:1px solid rgba(42,82,152,.2)}.btn-sm{border:none;border-radius:6px;padding:4px 10px;font-size:.7rem;cursor:pointer;font-family:"DM Sans",sans-serif;margin-right:3px}.btn-toggle{background:#ebf3fb;color:#2a5298;border:1px solid rgba(42,82,152,.2)}.btn-del{background:#fdf2f2;color:#c0392b;border:1px solid rgba(192,57,43,.2)}.btn-reset{background:#fff8ec;color:#b7701a;border:1px solid rgba(183,112,26,.3)}.vazio{text-align:center;color:#6b7a99;padding:24px}.protegido-tag{font-size:.6rem;color:#b7701a;background:#fff8ec;padding:1px 5px;border-radius:4px;margin-left:4px}.sem-senha-tag{font-size:.6rem;color:#2a5298;background:#ebf3fb;padding:1px 5px;border-radius:4px;margin-left:4px}@media(max-width:640px){.stats{grid-template-columns:1fr 1fr}.form-row,.form-row.tri{grid-template-columns:1fr}}</style></head><body><div class="hdr"><div class="hdr-left"><div class="hdr-sigla">PAAC</div><div class="hdr-sub">Painel — ' + adminLogado.nome + '</div></div><div class="hdr-right"><span class="badge">ADMIN E0</span><a href="/" class="btn-voltar">Voltar</a><a href="/sair" class="btn-sair">Sair</a></div></div><div class="main"><div class="stats"><div class="stat"><div class="stat-num" id="stTotal">0</div><div class="stat-label">Total</div></div><div class="stat"><div class="stat-num" id="stAtivos">0</div><div class="stat-label">Ativos</div></div><div class="stat"><div class="stat-num" id="stAdmins">0</div><div class="stat-label">Admins</div></div><div class="stat"><div class="stat-num" id="stInativos">0</div><div class="stat-label">Inativos</div></div></div><div class="card"><div class="card-hdr"><span class="card-icon">+</span><span class="card-title">Cadastrar Novo Usuario</span></div><div class="card-body"><div class="msg" id="msgCad"></div><div class="form-row"><div class="campo"><label>Nome *</label><input type="text" id="cNome" placeholder="Joao da Silva"></div><div class="campo"><label>E-mail *</label><input type="email" id="cEmail" placeholder="joao@email.com"></div></div><div class="form-row tri"><div class="campo"><label>Senha (opcional)</label>' + senhaInput('cSenha','','Branco = cria no 1o acesso',false).replace('name=""','') + '<span class="dica">Em branco: cria no 1o acesso.</span></div><div class="campo"><label>Papel</label><select id="cPapel"><option value="usuario">Usuario</option><option value="admin">Admin</option></select><span class="dica">Max 3 admins.</span></div><div class="campo"><label>Obs</label><input type="text" id="cObs" placeholder="Opcional"></div></div><button class="btn-add" id="btnCadastrar">Cadastrar</button></div></div><div class="card"><div class="card-hdr"><span class="card-icon">&#128274;</span><span class="card-title">Alterar Minha Senha</span></div><div class="card-body"><div class="msg" id="msgSenha"></div><div class="form-row tri"><div class="campo"><label>Atual *</label>' + senhaInput('sAtual','','','false') + '</div><div class="campo"><label>Nova *</label>' + senhaInput('sNova','','Min 6 chars',false) + '</div><div class="campo" style="justify-content:flex-end"><button class="btn-add" style="margin-top:0" id="btnAltSenha">Alterar</button></div></div></div></div><div class="card"><div class="card-hdr"><span class="card-icon">&#128101;</span><span class="card-title">Usuarios Cadastrados</span></div><div class="card-body" style="padding:0;overflow-x:auto"><table class="tabela"><thead><tr><th>Nome</th><th>E-mail</th><th>Papel</th><th>Status</th><th>Senha</th><th>Ultimo acesso</th><th>Acoes</th></tr></thead><tbody id="corpo"><tr><td colspan="7" class="vazio">Carregando...</td></tr></tbody></table></div></div></div><script>window._ADMIN_EMAIL="' + adminLogado.email + '";</script><script src="/painel/admin.js"></script>' + olhoScript + '</body></html>';
}

// ═══ ROTA /api/me ═══
app.get('/api/me', function(req, res) {
  var s = lerSessao(req);
  if (!s || !s.email) return res.status(401).json({ erro: 'Nao autenticado.' });
  var u = buscarUsuario(s.email);
  if (!u || !u.ativo) return res.status(401).json({ erro: 'Conta inativa.' });
  res.json({ ok: true, email: u.email, nome: u.nome, papel: u.papel });
});

// ═══ ROTA /api/verificar-senha — usada pelo modal de exclusão de projeto ═══
// Reaproveita verificarBloqueio/registrarFalha/limparTentativas para proteção anti-bruteforce
app.post('/api/verificar-senha', async function(req, res) {
  var s = lerSessao(req);
  if (!s || !s.email) return res.status(401).json({ ok: false, erro: 'Sessao expirada.' });
  var u = buscarUsuario(s.email);
  if (!u || !u.ativo) return res.status(401).json({ ok: false, erro: 'Conta inativa.' });

  // Verifica se o usuario ja esta bloqueado
  var bloq = verificarBloqueio(u.email);
  if (bloq) return res.status(429).json({ ok: false, bloqueado: true, minutos: bloq, erro: 'Bloqueado por seguranca. Aguarde ' + bloq + ' min.' });

  var senha = (req.body && req.body.senha) ? String(req.body.senha) : '';
  if (!senha) return res.status(400).json({ ok: false, erro: 'Senha obrigatoria.' });
  if (!u.senhaHash) return res.status(400).json({ ok: false, erro: 'Usuario sem senha cadastrada.' });

  try {
    var ok = await bcrypt.compare(senha, u.senhaHash);
    if (!ok) {
      registrarFalha(u.email);
      var rr = tentativasLogin[u.email.toLowerCase()];
      var rest = MAX_TENTATIVAS - (rr ? rr.tentativas : 0);
      if (rest <= 0) {
        return res.status(429).json({ ok: false, bloqueado: true, minutos: Math.ceil(TEMPO_BLOQUEIO / 60000), erro: 'Bloqueado por seguranca. Aguarde ' + Math.ceil(TEMPO_BLOQUEIO / 60000) + ' min.' });
      }
      return res.status(401).json({ ok: false, restantes: rest, erro: 'Senha incorreta. Restam ' + rest + ' tentativas.' });
    }
    limparTentativas(u.email);
    res.json({ ok: true });
  } catch (e) {
    console.error('Erro ao verificar senha:', e.message);
    res.status(500).json({ ok: false, erro: 'Erro interno.' });
  }
});

// ═══ ROTAS DE PROJETOS (Supabase) ═══
function exigirLogin(req, res, next) {
  var s = lerSessao(req);
  if (!s || !s.email) return res.status(401).json({ erro: 'Sessao expirada.' });
  var u = buscarUsuario(s.email);
  if (!u || !u.ativo) return res.status(401).json({ erro: 'Conta inativa.' });
  req.usuarioEmail = s.email.toLowerCase();
  req.usuarioNome = u.nome;
  next();
}

// GET /api/projetos
app.get('/api/projetos', exigirLogin, async function(req, res) {
  if (!supabase) return res.status(503).json({ erro: 'Banco de dados nao configurado.' });
  try {
    var { data, error } = await supabase
      .from('projetos')
      .select('id, nome_projeto, atualizado_em, created_at')
      .eq('usuario_email', req.usuarioEmail)
      .order('atualizado_em', { ascending: false });
    if (error) throw error;
    res.json({ ok: true, projetos: data || [] });
  } catch (e) {
    console.error('Erro ao listar projetos:', e.message);
    res.status(500).json({ erro: 'Erro ao buscar projetos.' });
  }
});

// GET /api/projetos/:id
app.get('/api/projetos/:id', exigirLogin, async function(req, res) {
  if (!supabase) return res.status(503).json({ erro: 'Banco de dados nao configurado.' });
  try {
    var { data, error } = await supabase
      .from('projetos')
      .select('*')
      .eq('id', req.params.id)
      .eq('usuario_email', req.usuarioEmail)
      .single();
    if (error) throw error;
    if (!data) return res.status(404).json({ erro: 'Projeto nao encontrado.' });
    res.json({ ok: true, projeto: data });
  } catch (e) {
    console.error('Erro ao carregar projeto:', e.message);
    res.status(500).json({ erro: 'Erro ao carregar projeto.' });
  }
});

// POST /api/projetos — cria ou atualiza (SEM .single() no update)
app.post('/api/projetos', exigirLogin, async function(req, res) {
  if (!supabase) return res.status(503).json({ erro: 'Banco de dados nao configurado.' });
  var { id, nome_projeto, dados } = req.body;
  if (!dados) return res.status(400).json({ erro: 'Dados do projeto obrigatorios.' });
  var nome = (nome_projeto || 'Sem nome').trim();
  var agora = new Date().toISOString();
  try {
    var resultado;
    if (id) {
      // Atualiza projeto existente — sem .single() para evitar erro
      var { data, error } = await supabase
        .from('projetos')
        .update({ nome_projeto: nome, dados: dados, atualizado_em: agora })
        .eq('id', id)
        .eq('usuario_email', req.usuarioEmail)
        .select();
      if (error) throw error;
      if (data && data.length > 0) {
        resultado = data[0];
      } else {
        // Registro não encontrado — criar novo com o mesmo id
        var { data: dataIns, error: errIns } = await supabase
          .from('projetos')
          .insert({ usuario_email: req.usuarioEmail, nome_projeto: nome, dados: dados, atualizado_em: agora })
          .select();
        if (errIns) throw errIns;
        resultado = dataIns[0];
      }
    } else {
      // Cria projeto novo — sem .single()
      var { data, error } = await supabase
        .from('projetos')
        .insert({ usuario_email: req.usuarioEmail, nome_projeto: nome, dados: dados, atualizado_em: agora })
        .select();
      if (error) throw error;
      resultado = data[0];
    }
    res.json({ ok: true, projeto: resultado });
  } catch (e) {
    console.error('Erro ao salvar projeto:', e.message);
    res.status(500).json({ erro: 'Erro ao salvar projeto.' });
  }
});

// DELETE /api/projetos/:id
app.delete('/api/projetos/:id', exigirLogin, async function(req, res) {
  if (!supabase) return res.status(503).json({ erro: 'Banco de dados nao configurado.' });
  try {
    var { error } = await supabase
      .from('projetos')
      .delete()
      .eq('id', req.params.id)
      .eq('usuario_email', req.usuarioEmail);
    if (error) throw error;
    res.json({ ok: true });
  } catch (e) {
    console.error('Erro ao excluir projeto:', e.message);
    res.status(500).json({ erro: 'Erro ao excluir projeto.' });
  }
});

// GET /api/nbr-status
app.get('/api/nbr-status', exigirLogin, async function(req, res) {
  if (!supabase) return res.status(503).json({ erro: 'Banco de dados nao configurado.' });
  var projeto_id = req.query.projeto_id;
  if (!projeto_id) return res.status(400).json({ erro: 'projeto_id obrigatorio.' });
  try {
    var { data, error } = await supabase
      .from('nbr_resultados')
      .select('status, resultados_qi_qiv, valor_total_obra, processado_em, erro_mensagem')
      .eq('projeto_id', projeto_id)
      .eq('usuario_email', req.usuarioEmail)
      .order('processado_em', { ascending: false })
      .limit(1)
      .single();
    if (error) {
      if (error.code === 'PGRST116') return res.json({ status: 'pendente' });
      throw error;
    }
    res.json({ status: data.status, resultados: data.resultados_qi_qiv || null, valor_total_obra: data.valor_total_obra || null, processado_em: data.processado_em || null, erro_mensagem: data.erro_mensagem || null });
  } catch (e) {
    console.error('Erro ao verificar status NBR:', e.message);
    res.status(500).json({ erro: 'Erro ao verificar status.' });
  }
});

// POST /api/nbr-resultado
app.post('/api/nbr-resultado', async function(req, res) {
  var token = req.headers['x-paac-token'];
  if (token !== (process.env.WEBHOOK_TOKEN || 'paac-token-e0-2026')) return res.status(401).json({ erro: 'Token invalido.' });
  if (!supabase) return res.status(503).json({ erro: 'Banco de dados nao configurado.' });
  var { projeto_id, usuario_email, status, resultados_qi_qiv, valor_total_obra, hash_dados_entrada, erro_mensagem } = req.body;
  if (!projeto_id || !usuario_email) return res.status(400).json({ erro: 'projeto_id e usuario_email obrigatorios.' });
  var agora = new Date().toISOString();
  try {
    var { data, error } = await supabase
      .from('nbr_resultados')
      .upsert({ projeto_id, usuario_email: usuario_email.toLowerCase(), status: status || 'nbr_parte1_concluida', resultados_qi_qiv: resultados_qi_qiv || null, valor_total_obra: valor_total_obra || null, hash_dados_entrada: hash_dados_entrada || null, processado_em: agora, erro_mensagem: erro_mensagem || null }, { onConflict: 'projeto_id' })
      .select()
      .single();
    if (error) throw error;
    res.json({ ok: true, processado_em: agora });
  } catch (e) {
    console.error('Erro ao salvar resultado NBR:', e.message);
    res.status(500).json({ erro: 'Erro ao salvar resultado.' });
  }
});

app.listen(PORT, '0.0.0.0', function() {
  console.log('');
  console.log('  PAAC v5.2 rodando na porta ' + PORT);
  console.log('');
});
