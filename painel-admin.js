var MEU_EMAIL = window._ADMIN_EMAIL;

async function carregar() {
  var r = await fetch('/painel/api/usuarios');
  var lista = await r.json();
  document.getElementById('stTotal').textContent = lista.length;
  document.getElementById('stAtivos').textContent = lista.filter(function(u){return u.ativo}).length;
  document.getElementById('stAdmins').textContent = lista.filter(function(u){return u.papel==='admin'}).length;
  document.getElementById('stInativos').textContent = lista.filter(function(u){return !u.ativo}).length;
  var corpo = document.getElementById('corpo');
  if (lista.length === 0) { corpo.innerHTML = '<tr><td colspan=7 class=vazio>Nenhum usuario.</td></tr>'; return; }
  var html = '';
  for (var i = 0; i < lista.length; i++) {
    var u = lista[i];
    var acesso = u.ultimoAcesso ? new Date(u.ultimoAcesso).toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit'}) : '\u2014';
    var ehAdmin = u.papel === 'admin';
    var ehEu = u.email.toLowerCase() === MEU_EMAIL.toLowerCase();
    var prot = u.protegido ? '<span class=protegido-tag>protegido</span>' : '';
    var acoes = '';
    if (!ehAdmin) {
      acoes = '<button class="btn-sm btn-toggle" onclick="toggle(\'' + u.email + '\')">' + (u.ativo ? 'Desativar' : 'Ativar') + '</button>';
      acoes += '<button class="btn-sm btn-del" onclick="remover(\'' + u.email + '\')">Remover</button>';
      acoes += '<button class="btn-sm btn-reset" onclick="resetSenha(\'' + u.email + '\')">Reset senha</button>';
    } else if (!ehEu) {
      acoes = '<button class="btn-sm btn-reset" onclick="resetSenha(\'' + u.email + '\')">Reset senha</button>';
    } else {
      acoes = '<span style="font-size:.68rem;color:#6b7a99">Voce</span>';
    }
    html += '<tr><td><strong>' + u.nome + '</strong>' + prot + (u.observacao ? '<br><small style="color:#6b7a99">' + u.observacao + '</small>' : '') + '</td>';
    html += '<td>' + u.email + (ehEu ? ' <small style="color:#c9a84c">(voce)</small>' : '') + '</td>';
    html += '<td><span class="papel-tag ' + u.papel + '">' + (ehAdmin ? 'Admin' : 'Usuario') + '</span></td>';
    html += '<td><span class="status ' + (u.ativo ? 'ativo' : 'inativo') + '">' + (u.ativo ? 'Ativo' : 'Inativo') + '</span></td>';
    html += '<td>' + (u.temSenha ? 'Sim' : '<span class="sem-senha-tag">Pendente</span>') + '</td>';
    html += '<td style="font-size:.74rem;color:#6b7a99">' + acesso + '</td>';
    html += '<td style="white-space:nowrap">' + acoes + '</td></tr>';
  }
  corpo.innerHTML = html;
}

async function cadastrar() {
  var nome = document.getElementById('cNome').value.trim();
  var email = document.getElementById('cEmail').value.trim();
  var senha = document.getElementById('cSenha').value;
  var papel = document.getElementById('cPapel').value;
  var obs = document.getElementById('cObs').value.trim();
  var msg = document.getElementById('msgCad');
  if (!nome || !email) { msg.className = 'msg erro'; msg.textContent = 'Nome e e-mail sao obrigatorios.'; return; }
  var body = { email: email, nome: nome, papel: papel, observacao: obs };
  if (senha) body.senha = senha;
  var r = await fetch('/painel/api/usuarios', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  var res = await r.json();
  if (res.ok) { msg.className = 'msg ok'; msg.textContent = res.mensagem; document.getElementById('cNome').value = ''; document.getElementById('cEmail').value = ''; document.getElementById('cSenha').value = ''; document.getElementById('cObs').value = ''; document.getElementById('cPapel').value = 'usuario'; carregar(); }
  else { msg.className = 'msg erro'; msg.textContent = res.erro; }
}

async function toggle(email) {
  var r = await fetch('/painel/api/usuarios/' + encodeURIComponent(email) + '/toggle', { method: 'PUT' });
  var res = await r.json();
  if (res.erro) alert(res.erro); else carregar();
}

async function remover(email) {
  if (!confirm('Remover o acesso de ' + email + '?')) return;
  var r = await fetch('/painel/api/usuarios/' + encodeURIComponent(email), { method: 'DELETE' });
  var res = await r.json();
  if (res.erro) alert(res.erro); else carregar();
}

async function resetSenha(email) {
  if (!confirm('Resetar a senha de ' + email + '? Ele tera que criar nova senha no proximo acesso.')) return;
  var r = await fetch('/painel/api/usuarios/' + encodeURIComponent(email) + '/reset-senha', { method: 'PUT' });
  var res = await r.json();
  if (res.erro) alert(res.erro); else { alert(res.mensagem); carregar(); }
}

async function trocarSenha() {
  var atual = document.getElementById('sAtual').value;
  var nova = document.getElementById('sNova').value;
  var msg = document.getElementById('msgSenha');
  if (!atual || !nova) { msg.className = 'msg erro'; msg.textContent = 'Preencha os dois campos.'; return; }
  var r = await fetch('/painel/api/minha-senha', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ senhaAtual: atual, senhaNova: nova }) });
  var res = await r.json();
  if (res.ok) { msg.className = 'msg ok'; msg.textContent = res.mensagem; document.getElementById('sAtual').value = ''; document.getElementById('sNova').value = ''; }
  else { msg.className = 'msg erro'; msg.textContent = res.erro; }
}

function toggleOlho(id) { var i = document.getElementById(id); i.type = i.type === 'password' ? 'text' : 'password'; }

carregar();
document.getElementById('btnCadastrar').addEventListener('click', cadastrar);
document.getElementById('btnAltSenha').addEventListener('click', trocarSenha);
