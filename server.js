const express = require('express');
const multer  = require('multer');
const PizZip  = require('pizzip');
const ExcelJS = require('exceljs');

const app  = express();
const PORT = process.env.PORT || 3000;

const storage = multer.memoryStorage();
const upload  = multer({ storage });
app.use(express.json({ limit: '50mb' }));

app.get('/health', (req, res) => res.json({ status: 'online', engine: 'exceljs', version: '7.2' }));

function limparCampos(raw) {
  if (typeof raw === 'object' && raw !== null) return raw;
  let txt = String(raw);
  if (txt.startsWith('=')) txt = txt.slice(1);
  if (txt.startsWith('"') && txt.endsWith('"')) txt = txt.slice(1, -1);
  txt = txt.replace(/\\"/g, '"');
  txt = txt.replace(/\r\n/g, ' ').replace(/\r/g, ' ').replace(/\n/g, ' ').replace(/\t/g, ' ');
  txt = txt.replace(/\u201C/g, "'").replace(/\u201D/g, "'");
  txt = txt.replace(/\u201E/g, "'").replace(/\u2018/g, "'").replace(/\u2019/g, "'");
  txt = txt.replace(/\\\\u([0-9a-fA-F]{4})/g, '\\u$1');
  try { return JSON.parse(txt); }
  catch (e) { throw new Error('JSON invalido: ' + e.message + ' | Inicio: ' + txt.substring(0, 300)); }
}

function toFloat(s) {
  if (s === null || s === undefined || s === '') return 0;
  if (typeof s === 'number') return s;
  let limpo = String(s).replace(/R\$\s*/g, '').trim();
  return parseFloat(limpo.replace(/\./g, '').replace(',', '.')) || 0;
}

function safeWrite(ws, celRef, valor, avisos) {
  const cell = ws.getCell(celRef);
  if (cell.formula || cell.sharedFormula) {
    avisos.push(`${celRef} ignorada — fórmula`);
    return false;
  }
  cell.value = valor;
  return true;
}

function forceWrite(ws, celRef, valor) {
  const cell = ws.getCell(celRef);
  cell.value = valor;
  return true;
}

function escXML(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// ══════════════════════════════════════════════════
// GERADORES DE BLOCOS DINÂMICOS (DOCX)
// ══════════════════════════════════════════════════

function xmlParagrafo(texto) {
  return `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">${escXML(texto)}</w:t></w:r></w:p>`;
}

function xmlParagrafoEspacamento(texto) {
  return `<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${escXML(texto)}</w:t></w:r></w:p>`;
}

function gerarParagrafosPavimentos(pavimentos) {
  if (!pavimentos || !pavimentos.length) return xmlParagrafo('[PENDENTE — nenhum pavimento]');
  return pavimentos.map(function(pav, idx) {
    var nome = pav.PAVIMENTO || ('Pavimento ' + (idx + 1));
    var descRaw = pav['DESC_CONVENÇÃO'] || pav['DESC_CONVENCAO'] || pav['DESC_MEMORIAL'] || pav.DESC || '[PENDENTE]';
    var desc = String(descRaw).replace(/\r\n/g, ' ').replace(/\r/g, ' ').replace(/\n/g, ' ').replace(/[ ]{2,}/g, ' ').trim();
    var texto = '\u00a7 ' + (idx + 1) + '. ' + nome + ': ' + desc;
    return xmlParagrafoEspacamento(texto);
  }).join('');
}

function gerarTabelaUnidades(unidades) {
  if (!unidades || !unidades.length) return xmlParagrafo('[PENDENTE — nenhuma unidade]');
  var W_LABEL = 5600, W_VALOR = 1600, W_UNID = 800;
  var W_TOT = W_LABEL + W_VALOR + W_UNID;
  var bcI = '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="BBBBBB"/><w:left w:val="none"/><w:bottom w:val="single" w:sz="4" w:color="BBBBBB"/><w:right w:val="none"/></w:tcBorders>';
  var shdC = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>';
  var shdB = '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>';
  function pPr15(jc) { return '<w:pPr><w:spacing w:line="360" w:lineRule="auto" w:before="20" w:after="20"/>' + (jc ? '<w:jc w:val="' + jc + '"/>' : '') + '</w:pPr>'; }
  function rPrTabela() { return '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>'; }
  function tcPr(w, shd) { return '<w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + bcI + shd + '<w:vAlign w:val="center"/></w:tcPr>'; }
  function linhaM2(label, valor, par) {
    var shd = par ? shdC : shdB; var val = String(valor || '[E1]');
    return '<w:tr><w:tc>' + tcPr(W_LABEL, shd) + '<w:p>' + pPr15() + '<w:r>' + rPrTabela() + '<w:t xml:space="preserve">' + escXML(label) + '</w:t></w:r></w:p></w:tc><w:tc>' + tcPr(W_VALOR, shd) + '<w:p>' + pPr15('right') + '<w:r>' + rPrTabela() + '<w:t xml:space="preserve">' + escXML(val) + '</w:t></w:r></w:p></w:tc><w:tc>' + tcPr(W_UNID, shd) + '<w:p>' + pPr15() + '<w:r>' + rPrTabela() + '<w:t xml:space="preserve"> m\u00b2</w:t></w:r></w:p></w:tc></w:tr>';
  }
  function linhaSemM2(label, valor, par) {
    var shd = par ? shdC : shdB; var val = String(valor || '[E1]');
    return '<w:tr><w:tc>' + tcPr(W_LABEL, shd) + '<w:p>' + pPr15() + '<w:r>' + rPrTabela() + '<w:t xml:space="preserve">' + escXML(label) + '</w:t></w:r></w:p></w:tc><w:tc>' + tcPr(W_VALOR + W_UNID, shd) + '<w:p>' + pPr15('right') + '<w:r>' + rPrTabela() + '<w:t xml:space="preserve">' + escXML(val) + '</w:t></w:r></w:p></w:tc></w:tr>';
  }
  return unidades.map(function(u) {
    var txRaw = u.TEXTO_DESCRITIVO || u.TEXTO || '';
    var tx = String(txRaw).replace(/\r\n/g,' ').replace(/\r/g,' ').replace(/\n/g,' ').replace(/[ ]{2,}/g,' ').trim();
    if (!tx) tx = 'A unidade aut\u00f4noma de n\u00ba ' + (u.DESIGNACAO||'') + ', localizada no ' + (u.PAVIMENTO||'').toLowerCase() + '.';
    var paraDesc = '<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto" w:after="160"/><w:ind w:firstLine="709"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">' + escXML(tx) + '</w:t></w:r></w:p>';
    var aPrivP = u.AREA_PRIV_PRINCIPAL || u.AREA_PP || '[E1]';
    var aOutras = parseFloat(String(u.AREA_PRIV_ACESS_OUTRAS||'0').replace(',','.')) || 0;
    var aVagas  = parseFloat(String(u.AREA_VAGAS_PRIV||'0').replace(',','.')) || 0;
    var aPrivA = u.AREA_PRIV_ACESS || u.AREA_PA ? (u.AREA_PRIV_ACESS || u.AREA_PA) : (aOutras + aVagas > 0 ? (aOutras+aVagas).toFixed(3).replace('.',',') : '[E1]');
    var aPP = parseFloat(String(aPrivP).replace(',','.')) || 0;
    var aPA = parseFloat(String(aPrivA).replace(',','.')) || 0;
    var aPrivTotal = (aPP > 0 && aPA > 0) ? (aPP+aPA).toFixed(3).replace('.',',') : '[E1]';
    var aCom = u.AREA_USO_COMUM || u.AREA_COMUM || ''; if (!aCom || aCom === '0') aCom = '[E1]';
    var aTotal = u.AREA_TOTAL || u.AREA_REAL_TOTAL || '[E1]';
    var coef   = u.COEFICIENTE || u.COEF_PROP || '[E1]';
    var colGrid = '<w:tblGrid><w:gridCol w:w="' + W_LABEL + '"/><w:gridCol w:w="' + W_VALOR + '"/><w:gridCol w:w="' + W_UNID + '"/></w:tblGrid>';
    var tabAreas = '<w:tbl><w:tblPr><w:tblW w:w="' + W_TOT + '" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="single" w:sz="4" w:color="BBBBBB"/><w:insideV w:val="none"/></w:tblBorders><w:tblIndent w:w="709" w:type="dxa"/></w:tblPr>' + colGrid + linhaM2('\u00c1rea Real Privativa Principal', aPrivP, true) + linhaM2('\u00c1rea Privativa Acess\u00f3ria', aPrivA, false) + linhaM2('\u00c1rea Privativa Total', aPrivTotal, true) + linhaM2('\u00c1rea de Uso Comum', aCom, false) + linhaM2('\u00c1rea Real Total', aTotal, true) + linhaSemM2('Coeficiente de Proporcionalidade', coef, false) + '</w:tbl>';
    var espacador = '<w:p><w:pPr><w:spacing w:after="240"/></w:pPr></w:p>';
    return paraDesc + tabAreas + espacador;
  }).join('');
}

function gerarParagrafosUnidades(unidades) {
  if (!unidades || !unidades.length) return xmlParagrafo('[PENDENTE — nenhuma unidade]');
  return unidades.map(u => {
    const txt = `UNIDADE ${u.DESIGNACAO||''} — ${u.PAVIMENTO||''}: composta de ${u.TEXTO_DESCRITIVO||u.TEXTO||'[PENDENTE]'}. ` +
      `Área privativa principal: ${u.AREA_PRIV_PRINCIPAL||u.AREA_PP||'[E1]'} m²; área privativa acessória: ${u.AREA_PRIV_ACESS||u.AREA_PA||'[E1]'} m²; área privativa total: ${u.AREA_TOTAL||'[E1]'} m².`;
    return xmlParagrafo(txt);
  }).join('');
}

function gerarLinhasTabVagas(vagas) {
  if (!vagas || !vagas.length) return `<w:tr><w:tc><w:p><w:r><w:t>Nenhuma vaga cadastrada.</w:t></w:r></w:p></w:tc></w:tr>`;
  const bc = `<w:tcBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>`;
  const map = {};
  vagas.forEach(v => {
    const u = v.unidade || v.UNIDADE || '';
    if (!u || u === '?' || u.trim() === '') return;
    if (!map[u]) map[u] = [];
    map[u].push(v);
  });
  const unidsOrdenadas = Object.keys(map).sort((a, b) => {
    const na = parseInt(a) || 0, nb = parseInt(b) || 0;
    if (na && nb) return na - nb;
    return a.localeCompare(b);
  });
  function descVaga(v) {
    const cob = (v.cobertura||'').toLowerCase();
    const cobTxt = cob === 'coberta' ? 'coberta' : cob === 'descoberta' ? 'descoberta' : cob === 'semicoberta' ? 'semicoberta' : cob || 'coberta';
    const tipo = (v.tipo||'').toLowerCase();
    if (tipo.includes('rotativo')) return 'Direito ao uso de 01 (uma) vaga de garagem de uso comum e rotativo';
    const usoTxt = tipo.includes('comum') && tipo.includes('proporcional') ? 'de uso comum de divisão proporcional' : tipo.includes('comum') ? 'de uso comum de divisão não-proporcional' : 'de uso Privativo';
    let desc = `Uma vaga de garagem ${cobTxt} ${usoTxt} localizada no ${v.pavimento||''}`;
    if (v.confinada) desc += ', sendo a primeira de livre acesso e a segunda confinada por esta primeira';
    if (v.pcd)   desc += ', adaptada ao uso PCD';
    if (v.idoso) desc += ' — discriminada como vaga de Idoso';
    return desc;
  }
  return unidsOrdenadas.map(unidade => {
    const vs = map[unidade];
    const desc = vs.map(v => descVaga(v)).join('; ');
    return `<w:tr><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/>${bc}</w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Apt. ${escXML(unidade)}</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="7560" w:type="dxa"/>${bc}</w:tcPr><w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">${escXML(desc)}</w:t></w:r></w:p></w:tc></w:tr>`;
  }).join('');
}

function gerarBlocoAssinaturas(unidades) {
  if (!unidades || !unidades.length) return xmlParagrafo('[PENDENTE — nenhuma unidade]');
  var designacoes = unidades.map(function(u) { return u.DESIGNACAO || ''; }).filter(Boolean);
  designacoes.sort(function(a, b) { var na = parseInt(a)||0, nb=parseInt(b)||0; return na&&nb?na-nb:a.localeCompare(b); });
  var pares = [];
  for (var i = 0; i < designacoes.length; i += 2) pares.push([designacoes[i], designacoes[i+1]||null]);
  var TAB_POS = '4500';
  var rPrAss = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>';
  var rPrLbl = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:i/></w:rPr>';
  var LINHA = '________________________';
  function blocoAss(esq, dir) {
    var pLinha = '<w:p><w:pPr><w:spacing w:before="600" w:after="0"/><w:tabs><w:tab w:val="left" w:pos="' + TAB_POS + '"/></w:tabs></w:pPr><w:r>' + rPrAss + '<w:t xml:space="preserve">' + LINHA + '</w:t></w:r>';
    if (dir) pLinha += '<w:r>' + rPrAss + '<w:tab/></w:r><w:r>' + rPrAss + '<w:t xml:space="preserve">' + LINHA + '</w:t></w:r>';
    pLinha += '</w:p>';
    var pLabel = '<w:p><w:pPr><w:spacing w:before="0" w:after="200"/><w:tabs><w:tab w:val="left" w:pos="' + TAB_POS + '"/></w:tabs></w:pPr><w:r>' + rPrLbl + '<w:t xml:space="preserve">Apt\u00ba ' + escXML(esq) + '</w:t></w:r>';
    if (dir) pLabel += '<w:r>' + rPrLbl + '<w:tab/></w:r><w:r>' + rPrLbl + '<w:t xml:space="preserve">Apt\u00ba ' + escXML(dir) + '</w:t></w:r>';
    pLabel += '</w:p>';
    return pLinha + pLabel;
  }
  return pares.map(function(par) { return blocoAss(par[0], par[1]); }).join('');
}

function gerarTabelaRateio(unidades, vlrTerreno) {
  if (!unidades || !unidades.length) return '';
  const bc = `<w:tcBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>`;
  function cel(txt, w, align, bold) {
    const b = bold ? '<w:b/>' : '';
    return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${bc}</w:tcPr><w:p><w:pPr><w:jc w:val="${align||'left'}"/></w:pPr><w:r><w:rPr>${b}<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${escXML(txt)}</w:t></w:r></w:p></w:tc>`;
  }
  const vlr = toFloat(vlrTerreno);
  const linhas = unidades.map(u => {
    const coef = u.COEFICIENTE || u.COEF_PROP || '[E1]';
    const coefNum = parseFloat(String(coef).replace(',','.')) || 0;
    const rateio = vlr && coefNum ? 'R$ ' + (vlr * coefNum).toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2}) : '[FASE E1]';
    return `<w:tr>${cel(u.DESIGNACAO||'',2000,'center')}${cel(coef,2500,'center')}${cel(rateio,4860,'center')}</w:tr>`;
  }).join('');
  return linhas;
}

function gerarRateioConstrucao(unidades) {
  if (!unidades || !unidades.length) return xmlParagrafo('[FASE E1 — Quadro III NBR 12721]');
  return unidades.map(u => {
    const txt = `Valor da construção da unidade ${u.DESIGNACAO||''} — ${u.PAVIMENTO||''}: R$ [FASE E1 — Quadro III NBR 12721].`;
    return xmlParagrafo(txt);
  }).join('') + xmlParagrafo('OBSERVAÇÃO: O valor de construção de cada unidade é o resultado obtido com a multiplicação do valor de construção do empreendimento (Quadro III da NBR 12.721/2006) pelo coeficiente de proporcionalidade de cada unidade (Quadro IV-A da NBR 12.721/2006).');
}

function gerarRateioTerreno(unidades, vlrTerreno) {
  if (!unidades || !unidades.length) return xmlParagrafo('[PENDENTE — nenhuma unidade]');
  const vlr = toFloat(vlrTerreno);
  const linhas = unidades.map(u => {
    const coef = u.COEFICIENTE || u.COEF_PROP || null;
    const coefNum = coef ? parseFloat(String(coef).replace(',','.')) : 0;
    const rateio = vlr && coefNum ? 'R$ ' + (vlr * coefNum).toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2}) : 'R$ [FASE E1]';
    const txt = `Valor da fração de terreno pertencente à unidade ${u.DESIGNACAO||''} — ${u.PAVIMENTO||''}: ${rateio};`;
    return xmlParagrafo(txt);
  }).join('');
  return linhas + xmlParagrafo('OBSERVAÇÃO: O valor correspondente a cada unidade é o valor do terreno multiplicado pelo coeficiente de proporcionalidade de cada unidade (Quadro II da NBR 12.721/2006).');
}

function processarBlocos(xml, campos) {
  const pavs  = campos._pavimentos || [];
  const unis  = campos._unidades   || [];
  const vagas = campos._vagas      || [];
  function subParagrafo(xmlIn, chave, xmlGerado) {
    const re = new RegExp(`<w:p\\b[^>]*>(?:(?!<\\/w:p>)[\\s\\S])*?\\{\\{${chave}\\}\\}[\\s\\S]*?<\\/w:p>`, 'g');
    return xmlIn.replace(re, xmlGerado);
  }
  function subTr(xmlIn, chave, xmlGerado) {
    const re = new RegExp(`<w:tr\\b[^>]*>(?:(?!<\\/w:tr>)[\\s\\S])*?\\{\\{${chave}\\}\\}[\\s\\S]*?<\\/w:tr>`, 'g');
    return xmlIn.replace(re, xmlGerado);
  }
  if (xml.includes('{{PARAGRAFOS_PAVIMENTOS}}'))   xml = subParagrafo(xml, 'PARAGRAFOS_PAVIMENTOS',   gerarParagrafosPavimentos(pavs));
  if (xml.includes('{{TABELA_UNIDADES}}'))         xml = subParagrafo(xml, 'TABELA_UNIDADES',         gerarTabelaUnidades(unis));
  if (xml.includes('{{PARAGRAFOS_UNIDADES}}'))     xml = subParagrafo(xml, 'PARAGRAFOS_UNIDADES',     gerarParagrafosUnidades(unis));
  if (xml.includes('{{TABELA_VAGAS}}'))            xml = subTr(xml,        'TABELA_VAGAS',            gerarLinhasTabVagas(vagas));
  if (xml.includes('{{BLOCO_ASSINATURAS}}'))       xml = subParagrafo(xml, 'BLOCO_ASSINATURAS',       gerarBlocoAssinaturas(unis));
  const vlrTerreno = campos.VALOR_TERRENO || campos.VLR_TERRENO || '';
  if (xml.includes('{{TABELA_RATEIO}}'))           xml = subTr(xml,        'TABELA_RATEIO',           gerarTabelaRateio(unis, vlrTerreno));
  if (xml.includes('{{TABELA_RATEIO_CONSTRUCAO}}'))xml = subParagrafo(xml, 'TABELA_RATEIO_CONSTRUCAO',gerarRateioConstrucao(unis));
  if (xml.includes('{{TABELA_RATEIO_TERRENO}}'))   xml = subParagrafo(xml, 'TABELA_RATEIO_TERRENO',   gerarRateioTerreno(unis, vlrTerreno));
  return xml;
}

// ══════════════════════════════════════════════════
// CABEÇALHO E RODAPÉ
// ══════════════════════════════════════════════════

function xmlHeader(nomeDoc, nomeEmp) {
  const texto = escXML(nomeDoc) + ' | ' + escXML(nomeEmp);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
  <w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2E74B5"/></w:pBdr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:color w:val="404040"/></w:rPr><w:t xml:space="preserve">${texto}</w:t></w:r></w:p>
</w:hdr>`;
}

function xmlFooter() {
  const textoPAAC = 'Documento gerado com o auxílio da Plataforma PAAC — Automação e Auditoria Cartorial';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
  <w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="2E74B5"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="707070"/><w:i/></w:rPr><w:t xml:space="preserve">${escXML(textoPAAC)}</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="707070"/></w:rPr><w:tab/><w:t xml:space="preserve">Página </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="404040"/><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="404040"/><w:b/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="404040"/><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="707070"/></w:rPr><w:t xml:space="preserve"> de </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="404040"/><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="404040"/><w:b/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="14"/><w:color w:val="404040"/><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>
</w:ftr>`;
}

function injetarHeaderFooter(zip, nomeDoc, nomeEmp) {
  zip.file('word/header1.xml', xmlHeader(nomeDoc, nomeEmp));
  zip.file('word/footer1.xml', xmlFooter());
  const relsFile = zip.file('word/_rels/document.xml.rels');
  let rels = relsFile ? relsFile.asText() : '';
  if (!rels.includes('header1.xml')) {
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>');
    zip.file('word/_rels/document.xml.rels', rels);
  }
  const ctFile = zip.file('[Content_Types].xml');
  let ct = ctFile ? ctFile.asText() : '';
  if (!ct.includes('header1.xml')) {
    ct = ct.replace('</Types>', '<Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" PartName="/word/header1.xml"/><Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" PartName="/word/footer1.xml"/></Types>');
    zip.file('[Content_Types].xml', ct);
  }
  const docFile = zip.file('word/document.xml');
  if (docFile) {
    let doc = docFile.asText();
    doc = doc.replace(/<w:headerReference[^\/]*\/>/g, '');
    doc = doc.replace(/<w:footerReference[^\/]*\/>/g, '');
    doc = doc.replace('<w:sectPr>', '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/>');
    if (!doc.includes('xmlns:r=')) doc = doc.replace('<w:document ', '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    zip.file('word/document.xml', doc);
  }
}

const NOMES_DOCS = {
  'memorial_incorporacao': 'Memorial de Incorporação',
  'memorial_instituicao':  'Memorial de Instituição de Condomínio Edilício',
  'convencao':             'Convenção de Condomínio Edilício',
  'declaracao_vagas':      'Declaração de Vagas de Garagem',
};

function detectarNomeDoc(campos, nomeArquivo) {
  if (campos._nome_template) return NOMES_DOCS[campos._nome_template] || campos._nome_template;
  if (nomeArquivo) {
    const arq = nomeArquivo.toLowerCase();
    if (arq.includes('incorporacao') || arq.includes('incorporação')) return NOMES_DOCS['memorial_incorporacao'];
    if (arq.includes('instituicao')  || arq.includes('instituição'))  return NOMES_DOCS['memorial_instituicao'];
    if (arq.includes('convencao')    || arq.includes('convenção'))    return NOMES_DOCS['convencao'];
    if (arq.includes('vaga')         || arq.includes('declaracao'))   return NOMES_DOCS['declaracao_vagas'];
  }
  var tipoMem = campos.tipo_memorial || campos.TIPO_MEMORIAL_SEL || '';
  if (tipoMem === 'instituicao')  return NOMES_DOCS['memorial_instituicao'];
  if (tipoMem === 'incorporacao') return NOMES_DOCS['memorial_incorporacao'];
  if (campos.NUMERO_HABITESE || campos.NUM_HABITESE) return NOMES_DOCS['memorial_instituicao'];
  if (campos.REGIME_INCORPORACAO || campos.REGIME)   return NOMES_DOCS['memorial_incorporacao'];
  return 'Documento';
}

function consolidarPlaceholders(xml) {
  let resultado = xml.replace(
    /\{\{<\/w:t><\/w:r>([\s\S]*?)\}\}/g,
    function(match, meio) {
      const textos = [];
      const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
      let m;
      while ((m = tRegex.exec(meio)) !== null) textos.push(m[1]);
      if (textos.length > 0) return '{{' + textos.join('') + '}}';
      return match;
    }
  );
  return resultado;
}

// ══════════════════════════════════════════════════
// ENDPOINT 1 — /preencher-docx
// ══════════════════════════════════════════════════
app.post('/preencher-docx', (req, res) => {
  try {
    let body = req.body;
    if (Array.isArray(body) && body.length > 0) body = body[0];
    let templateBuffer;
    if (req.file) { templateBuffer = req.file.buffer; }
    else if (body && body.template_base64) { templateBuffer = Buffer.from(body.template_base64, 'base64'); }
    else { return res.status(400).json({ erro: 'Nenhum template recebido.', body_type: typeof body, keys: body ? Object.keys(body).slice(0,5) : [] }); }
    let campos;
    try { campos = limparCampos(body?.campos || ''); }
    catch (e) { return res.status(400).json({ erro: 'Campo "campos" inválido.', detalhe: e.message }); }
    if (body?.blocos) { try { const blocos = limparCampos(body.blocos); Object.assign(campos, blocos); } catch (e) {} }
    const zip = new PizZip(templateBuffer);
    const zipFiles = zip.files || {};
    const xmlFiles = [];
    Object.keys(zipFiles).forEach(function(relativePath) {
      var file = zipFiles[relativePath];
      if (relativePath.startsWith('word/') && relativePath.endsWith('.xml') && !file.dir) {
        var nome = relativePath.toLowerCase();
        if (nome.includes('document') || nome.includes('header') || nome.includes('footer')) xmlFiles.push(relativePath);
      }
    });
    if (xmlFiles.length === 0) return res.status(500).json({ erro: 'Nenhum XML encontrado no template DOCX.' });
    let substituicoes = 0;
    xmlFiles.forEach(function(xmlPath) {
      const xmlFileItem = zip.file(xmlPath);
      if (!xmlFileItem) return;
      let xml = xmlFileItem.asText();
      xml = consolidarPlaceholders(xml);
      if (xmlPath.includes('document')) xml = processarBlocos(xml, campos);
      for (const [chave, valor] of Object.entries(campos)) {
        if (chave.startsWith('_')) continue;
        const seguro = String(valor ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
        const regex = new RegExp(`\\{\\{${chave}\\}\\}`, 'g');
        const antes = xml;
        xml = xml.replace(regex, seguro);
        if (xml !== antes) substituicoes++;
      }
      zip.file(xmlPath, xml);
    });
    const docxFinal = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    res.json({ sucesso: true, substituicoes, docx_base64: docxFinal.toString('base64') });
  } catch (err) {
    console.error('Erro /preencher-docx:', err);
    res.status(500).json({ erro: 'Erro interno.', detalhe: err.message });
  }
});

// ══════════════════════════════════════════════════
// ENDPOINT 2 — /preencher-xlsx  v7.0
// ══════════════════════════════════════════════════
app.post('/preencher-xlsx', upload.single('planilha'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ erro: 'Nenhuma planilha recebida.' });
    let dados;
    try { dados = limparCampos(req.body?.dados || ''); }
    catch (e) { return res.status(400).json({ erro: 'Campo dados invalido.', detalhe: e.message }); }

    const config      = dados.config      || dados || {};
    const pavimentos  = dados.pavimentos  || [];
    const unidades    = dados.unidades    || [];
    const custos      = dados.custos      || {};
    const memorial    = dados.memorial    || {};
    const equipamentos = dados.equipamentos || {};
    const acabamentos  = dados.acabamentos  || {};
    const avisos = [];
    let celulasEscritas = 0;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const findSheet = (re) => wb.worksheets.find(ws => re.test(ws.name));
    const sw = (ws, cel, val) => { if (safeWrite(ws, cel, val, avisos)) celulasEscritas++; };

    // ── P4: Reforçar orientação de página de cada aba ──────────────────
    // O ExcelJS pode perder as orientações ao re-salvar.
    // Mapeamento fixo baseado no template NBR_12721_P36_UNIDv02:
    const orientacoes = {
      'portrait':  [/Inf Prel/i, /Q[\s_-]*III/i, /Q[\s_-]*V-1/i, /Q[\s_-]*V-2/i],
      'landscape': [/Q[\s_-]*I/i, /Q[\s_-]*II/i, /Q[\s_-]*IV[\s_-]*A/i,
                    /Q[\s_-]*IV[\s_-]*B/i, /Q[\s_-]*VI/i, /Q[\s_-]*VII/i, /Q[\s_-]*VIII/i],
    };
    wb.worksheets.forEach(function(ws) {
      if (!ws.pageSetup) ws.pageSetup = {};
      ws.pageSetup.paperSize = 9; // A4
      for (const [orient, patterns] of Object.entries(orientacoes)) {
        if (patterns.some(function(re){ return re.test(ws.name); })) {
          ws.pageSetup.orientation = orient;
          break;
        }
      }
      // Abas Q V-2 criadas dinamicamente também são portrait
      if (/Q[\s_-]*V-2/i.test(ws.name)) ws.pageSetup.orientation = 'portrait';
    });

    // ── DADOS_REF ──────────────────────────────────────────────────────────
    const shRef = wb.getWorksheet('DADOS_REF');
    if (shRef) {
      sw(shRef, 'B4', toFloat(config.FATOR_REDUCAO || config.FATOR_RED || '0.25'));
      sw(shRef, 'B5', config.DATA_DOCUMENTO || new Date().toLocaleDateString('pt-BR'));
    } else avisos.push('Aba DADOS_REF não encontrada.');

    // ── INF PREL (Folha 1) ────────────────────────────────────────────────
    const shInf = findSheet(/Inf Prel/i);
    if (shInf) {
      sw(shInf,'D5', config.NOME_INCORPORADOR||config.NOME_INCORP||'');
      sw(shInf,'E6', config.CNPJ_CPF_INCORPORADOR||config.CNPJ_CPF||'');
      sw(shInf,'E7', config.ENDERECO_INCORPORADOR||config.END_INCORP||'');
      const pref = (config.TITULO_RT||'Arquiteto').startsWith('Arq') ? 'Arq.' : 'Eng.';
      sw(shInf,'G10', `${pref} ${config.NOME_RT||''}`);
      sw(shInf,'G11', config.REGISTRO_RT||'');
      sw(shInf,'G12', config.RRT_ART||'');
      sw(shInf,'E13', config.ENDERECO_RT||config.END_RT||'');
      sw(shInf,'F16', config.NOME_EDIFICIO||'');
      sw(shInf,'F17', config.LOCAL_CONSTRUCAO||config.ENDERECO_COMPLETO||'');

      // ── CORREÇÃO A: UF — fallback robusto para campo vazio ou não selecionado ──
      const ufDireto = config.UF && String(config.UF).trim() && config.UF !== '[PENDENTE]'
        ? String(config.UF).trim() : '';
      const cidadeUfPartes = String(config.CIDADE_UF || '').split(' - ');
      const ufFallback = cidadeUfPartes.length > 1 ? cidadeUfPartes[cidadeUfPartes.length - 1].trim() : '';
      const uf = ufDireto || ufFallback || '';
      sw(shInf,'F18', (config.CIDADE || '') + (uf ? ' - ' + uf : ''));

      const pp = (config.PROJETO_PADRAO||config.PADRAO_NBR||'').toUpperCase();
      let ppL='', ppCS='', ppCL='';
      if(pp.startsWith('CSL')){ppCS='CS';ppCL='CL';}else if(pp.startsWith('CS')){ppCS='CS';}else if(pp.startsWith('CL')){ppCL='CL';}else if(pp.startsWith('R')){ppL='R';}
      sw(shInf,'G20',ppL); sw(shInf,'I20',ppCS); sw(shInf,'L20',ppCL);
      const pad=(config.PADRAO_ACABAMENTO||config.PADRAO_ACAB||'').toUpperCase();
      sw(shInf,'G21',`${config.PROJETO_PADRAO||config.PADRAO_NBR||''} - ${pad?pad.charAt(0):''}`);
      sw(shInf,'G22',`${config.QTD_UNIDADES||config.QTD_UNID||''} (${config.QTD_UNIDADES_EXTENSO||config.QTD_UNID_EXT||''})`);
      sw(shInf,'G23',config.PADRAO_ACABAMENTO||config.PADRAO_ACAB||'');
      sw(shInf,'G24',config.NUM_PAVIMENTOS||config.NUM_PAV||'');
      sw(shInf,'H27',config.DESC_VAGAS||'');
      sw(shInf,'H29',toFloat(config.AREA_TERRENO));
      sw(shInf,'G30',config.DATA_ALVARA||'');
      sw(shInf,'G31',config.NUMERO_ALVARA||config.NUM_ALVARA||'');
      sw(shInf,'G32',config.DATA_HABITESE||'');
      sw(shInf,'G33',config.NUMERO_HABITESE||config.NUM_HABITESE||'');
      sw(shInf,'E41',config.DATA_DOCUMENTO_EXTENSO||config.DATA_DOCUMENTO||new Date().toLocaleDateString('pt-BR'));
      sw(shInf,'C55',config.OBS_PAVIMENTOS||'');
      sw(shInf,'N1',1);
    } else avisos.push('Aba Inf Prel não encontrada.');

    // ── QUADRO I ───────────────────────────────────────────────────────────
    const shQI = findSheet(/Q[\s_-]*I[\s_-]*[-–]/i)||findSheet(/Q[\s_-]*I\b/i);
    if (shQI && pavimentos.length > 0) {
      sw(shQI,'B12',config.DATA_DOCUMENTO_EXTENSO||config.DATA_DOCUMENTO||'');
      sw(shQI,'L12',config.TELEFONE_RT||config.TEL_RT||'');
      pavimentos.forEach((pav,idx) => {
        const ln = 29 + idx; if (ln > 45) { avisos.push(`Pav ${idx+1} ignorado`); return; }
        sw(shQI,`A${ln}`,pav.PAVIMENTO||`Pavimento ${idx+1}`);
        sw(shQI,`B${ln}`,toFloat(pav.AREA_PRIV_COB_PAD)); sw(shQI,`C${ln}`,toFloat(pav.AREA_PRIV_COB_DIF));
        sw(shQI,`G${ln}`,toFloat(pav.AREA_COM_COB_PAD));  sw(shQI,`H${ln}`,toFloat(pav.AREA_COM_COB_DIF));
        sw(shQI,`L${ln}`,toFloat(pav.AREA_PROP_COB_PAD)); sw(shQI,`M${ln}`,toFloat(pav.AREA_PROP_COB_DIF));
      });
    } else if (!shQI) avisos.push('Aba Q I não encontrada.');

    // ── QUADRO II ──────────────────────────────────────────────────────────
    const shQII = findSheet(/Q[\s_-]*II[\s_-]*[-–]/i)||findSheet(/Q[\s_-]*II\b/i);
    if (shQII && unidades.length > 0) {
      sw(shQII,'M8',config.TELEFONE_RT||config.TEL_RT||'');
      const max = Math.min(unidades.length, 36);
      for (let i = 0; i < max; i++) {
        const u = unidades[i], ln = 21 + i;
        sw(shQII,`A${ln}`,u.DESIGNACAO||`Unidade ${i+1}`);
        sw(shQII,`B${ln}`,toFloat(u.AREA_PRIV_PRINCIPAL||u.AREA_PP));
        sw(shQII,`C${ln}`,toFloat(u.AREA_PRIV_ACESS||u.AREA_PA));
        sw(shQII,`G${ln}`,toFloat(u.AREA_COM_NP_COB_PAD||0));
        sw(shQII,`H${ln}`,toFloat(u.AREA_COM_NP_COB_DIF||0));
      }
      // ── P2: Garantir que células não usadas permaneçam null (não zero) ──
      // O ExcelJS pode converter None→0 ao salvar, fazendo o formato mostrar "0,000"
      // em vez do traço "-" definido pelo formato de número do template.
      for (let i = max; i < 36; i++) {
        const ln = 21 + i;
        ['B','C','G','H'].forEach(function(col) {
          const cell = shQII.getCell(col + ln);
          if (!cell.formula && !cell.sharedFormula && (cell.value === 0 || cell.value === null)) {
            cell.value = null;
          }
        });
      }
      if (unidades.length > 36) avisos.push(`${unidades.length} unidades excedem 36/folha.`);
    } else if (!shQII) avisos.push('Aba Q II não encontrada.');

    // ── QUADRO III ─────────────────────────────────────────────────────────
    const shQIII = findSheet(/Q[\s_-]*III/i);
    if (shQIII) {
      // ── P4: mapeamento correto das células do Q III ──────────────────────
      // C24 = designação projeto-padrão (ex: R8-N)
      // D24 = padrão de acabamento (ex: NORMAL)
      // E24 = número de pavimentos
      sw(shQIII,'C24',config.PROJETO_PADRAO||config.PADRAO_NBR||'');
      sw(shQIII,'D24',config.PADRAO_ACABAMENTO||config.PADRAO_ACAB||'');
      sw(shQIII,'E24',config.NUM_PAVIMENTOS||config.NUM_PAV||'');
      sw(shQIII,'F28',config.CUB_MES_REFERENCIA||config.CUB_MES||'');
      sw(shQIII,'K28',toFloat(config.CUB_VALOR||config.CUB_VLR));
      // Tipologias: quartos, salas, banheiros, empregados da primeira tipologia
      if(config.QUARTOS_PAD !== undefined)   sw(shQIII,'G23', parseInt(config.QUARTOS_PAD)||0);
      if(config.SALAS_PAD !== undefined)     sw(shQIII,'I23', parseInt(config.SALAS_PAD)||0);
      if(config.BANHEIROS_PAD !== undefined) sw(shQIII,'K23', parseInt(config.BANHEIROS_PAD)||0);
      if(config.EMPREG_PAD !== undefined)    sw(shQIII,'L23', parseInt(config.EMPREG_PAD)||0);
      if(custos.adicionais){[45,46,48,49,50,51,52,53,54,55,56,58,59,60,61,62,64,65,66,67].forEach((ln,i)=>{if(custos.adicionais[i])sw(shQIII,`L${ln}`,toFloat(custos.adicionais[i]));});}
      if(custos.impostos){[71,72].forEach((ln,i)=>{if(custos.impostos[i])sw(shQIII,`L${ln}`,toFloat(custos.impostos[i]));});}
      if(custos.projetos){[75,76,77].forEach((ln,i)=>{if(custos.projetos[i])sw(shQIII,`L${ln}`,toFloat(custos.projetos[i]));});}
      if(custos.remuneracoes){[80,81].forEach((ln,i)=>{if(custos.remuneracoes[i])sw(shQIII,`L${ln}`,toFloat(custos.remuneracoes[i]));});}
      if(custos.perc_materiais) sw(shQIII,'N41',toFloat(custos.perc_materiais));
      if(custos.perc_mao_obra)  sw(shQIII,'N42',toFloat(custos.perc_mao_obra));
    }

    // ── QUADRO IV-B ────────────────────────────────────────────────────────
    const shQIVB = findSheet(/Q[\s_-]*IV[\s_-]*B/i);
    const FMT_COEF5 = '_(* #,##0.00000_);_(* (#,##0.00000);_(* "-"??_);_(@_)';
    const FMT_AREA3 = '_(* #,##0.000_);_(* (#,##0.000);_(* "-"??_);_(@_)';
    if (shQIVB) {
      sw(shQIVB,'F11',config.REGISTRO_RT||'');
      // Coeficiente de Proporcionalidade — 5 casas decimais SOMENTE na coluna G
      // Colunas de área (B, C, D, E, F) — 3 casas decimais (padrão NBR)
      for (let r = 17; r <= 52; r++) {
        shQIVB.getCell('B' + r).numFmt = FMT_AREA3;
        shQIVB.getCell('C' + r).numFmt = FMT_AREA3;
        shQIVB.getCell('D' + r).numFmt = FMT_AREA3;
        shQIVB.getCell('E' + r).numFmt = FMT_AREA3;
        shQIVB.getCell('F' + r).numFmt = FMT_AREA3;
        shQIVB.getCell('G' + r).numFmt = FMT_COEF5;
      }
    }

    // ── Coeficiente 5 casas decimais no Q II (coluna M) ──
    if (shQII) {
      for (let r = 21; r <= 57; r++) {
        shQII.getCell('M' + r).numFmt = FMT_COEF5;
      }
    }

    // ── Coeficiente 5 casas decimais no Q IV-A (coluna D) ──
    const shQIVA = findSheet(/Q[\s_-]*IV[\s_-]*A/i);
    if (shQIVA) {
      for (let r = 23; r <= 58; r++) {
        shQIVA.getCell('D' + r).numFmt = FMT_COEF5;
      }
    }

    // ── P6: QUADRO V-1 — mapeamento célula a célula com objeto memorial estruturado ──
    // O template tem estrutura por item (a, b, c, d, e, f, g):
    //   a) Tipo de edificação  → C18  (célula de valor após rótulo B18)
    //   b) Nomes de pavimentos → D19  (célula de valor após rótulo B19:C19)
    //   c) Unidades/pavimento  → B22  (linha após rótulo B21)
    //   d) Numeração unidades  → B24  (linha após rótulo B23)
    //   e) Descrição pavtos    → B26..B30 (linhas após rótulo B25)
    //   f) Data/alvará         → B32..B36 (linhas após rótulo B31)
    //   g) Outras indicações   → B38..B42 (linhas após rótulo B37)
    //   Área terreno (ha)      → M38  |  Reservatório → M40
    const shQV1 = findSheet(/Q[\s_-]*V-1/i)||findSheet(/Q[\s_-]*V[\s_-]*1/i);
    if (shQV1) {
      const mem = memorial || {};

      function escQV1(celRef, valor) {
        if (valor === null || valor === undefined || valor === '') return;
        const cell = shQV1.getCell(celRef);
        if (cell.type === 'Formula') return; // não sobrescrever fórmulas
        try {
          cell.value = String(valor);
          cell.alignment = { wrapText: true, vertical: 'top' };
          celulasEscritas++;
        } catch(e) { avisos.push('Q V-1 ' + celRef + ': ' + e.message); }
      }

      // a) Tipo de edificação
      if (mem.tipo_edificacao) escQV1('C18', mem.tipo_edificacao);
      else {
        const itemTipo = (mem.descritivo||[]).find(function(t){ return t && /tipo de edifica/i.test(t); });
        if (itemTipo) escQV1('C18', itemTipo.replace(/^tipo de edifica[çc][aã]o:\s*/i,'').trim());
      }

      // b) Número e nome de pavimentos
      if (mem.nomes_pavimentos) {
        const nomesStr = Array.isArray(mem.nomes_pavimentos) ? mem.nomes_pavimentos.join(', ') : mem.nomes_pavimentos;
        escQV1('D19', nomesStr);
      } else {
        const itemPav = (mem.descritivo||[]).find(function(t){ return t && /número e nome/i.test(t); });
        if (itemPav) escQV1('D19', itemPav.replace(/^n[úu]mero e nome de pavimentos?:\s*/i,'').trim());
      }

      // c) Unidades autônomas por pavimento — aglutinado (enunciado + descricao)
      if (mem.desc_blocos) {
        const blocoStr = Array.isArray(mem.desc_blocos) ? mem.desc_blocos.join(' / ') : mem.desc_blocos;
        escQV1('B21', blocoStr);
      }

      // d) Numeracao das unidades — linha 22 (ajuste de template 04/abr/2026)
      if (mem.nomes_unidades) {
        const unidStr = Array.isArray(mem.nomes_unidades) ? mem.nomes_unidades.join(', ') : mem.nomes_unidades;
        escQV1('B22', unidStr);
        // Garantir que B23 e B24 nao ficam com residuo de versoes anteriores
        forceWrite(shQV1, 'B23', null);
        forceWrite(shQV1, 'B24', null);
      }

      // e) Descricao dos pavimentos — comeca na linha 23 (livre apos item D ir para 22)
      // Cada pavimento ocupa uma linha; Pavimento Coberta SEMPRE inserido por ultimo.
      // O Script02 garante que pavimentos_desc ja contem Coberta no final.
      const pavDescs = (mem.pavimentos_desc || []);
      const itensPavFallback = (mem.descritivo||[]).filter(function(t){ return t && /^§\s*\d+\./.test(t); });
      const fonteDesc = pavDescs.length > 0 ? pavDescs : itensPavFallback;
      fonteDesc.forEach(function(txt, idx) {
        // Linha base 23; cada pavimento adicional avanca uma linha
        escQV1('B' + (23 + idx), txt);
      });

      // f) Data de aprovação / alvará / habite-se / responsáveis técnicos
      const linhasF = [32, 33, 34, 35, 36];
      const itensF = (mem.itens_f || []);
      // fallback: buscar item que começa com "Alvará"
      if (itensF.length === 0) {
        const itemAlv = (mem.descritivo||[]).find(function(t){ return t && /alvará/i.test(t); });
        if (itemAlv) escQV1('B32', itemAlv);
      } else {
        itensF.forEach(function(txt, idx) { if (idx < linhasF.length) escQV1('B' + linhasF[idx], txt); });
      }

      // g) Outras indicações
      const linhasG = [38, 39, 40, 41, 42];
      const itensG = (mem.itens_g || []);
      if (itensG.length === 0) {
        let gIdx = 0;
        if (mem.area_terreno_ha) { escQV1('B38', 'Área do terreno: ' + toFloat(mem.area_terreno_ha).toFixed(4) + ' ha'); gIdx++; }
        if (mem.capacidade_reservatorio) { escQV1('B' + linhasG[gIdx], 'Capacidade total dos reservatórios: ' + toFloat(mem.capacidade_reservatorio).toFixed(1) + ' m³'); }
      } else {
        itensG.forEach(function(txt, idx) { if (idx < linhasG.length) escQV1('B' + linhasG[idx], txt); });
      }

      // Campos numéricos M38 e M40
      if (mem.area_terreno_ha) {
        try { const c = shQV1.getCell('M38'); if (c.type !== 'Formula') { c.value = toFloat(mem.area_terreno_ha); celulasEscritas++; } } catch(e) {}
      }
      if (mem.capacidade_reservatorio) {
        try { const c = shQV1.getCell('M40'); if (c.type !== 'Formula') { c.value = toFloat(mem.capacidade_reservatorio); celulasEscritas++; } } catch(e) {}
      }

    } else avisos.push('Aba Q V-1 não encontrada.');

    // ── CORREÇÃO C: QUADRO V-2 — suporte a múltiplas folhas ──────────────
    // Cada aba comporta 4 unidades (posições de texto: B19, B28, B37, B46)
    // Para mais de 4 unidades, o docserver copia a aba original e ajusta
    // as fórmulas de áreas para apontar às linhas corretas do Q IV-B.
    //
    // Estrutura de fórmulas por folha:
    //   Folha 1 (original): Q IV-B linhas 17-20 (unidades 0-3)
    //   Folha 2 (cópia):    Q IV-B linhas 21-24 (unidades 4-7)
    //   Folha 3 (cópia):    Q IV-B linhas 25-28 (unidades 8-11)
    //   etc.
    //
    // Posições de texto descritivo dentro de cada aba (fixas):
    //   slot 0: B19, slot 1: B28, slot 2: B37, slot 3: B46
    // Posições de designação (DESIGNACAO) dentro de cada aba:
    //   slot 0: A19, slot 1: A28, slot 2: A37, slot 3: A46

    const shQV2original = findSheet(/Q[\s_-]*V-2/i)||findSheet(/Q[\s_-]*V[\s_-]*2/i);

    if (shQV2original && unidades.length > 0) {

      // Posições fixas de texto e designação dentro de uma aba Q V-2
      const SLOTS_TEXTO  = [19, 28, 37, 46]; // linha do texto descritivo
      const SLOTS_DESIG  = [19, 28, 37, 46]; // mesma linha — coluna A para designação, B para texto
      // Linhas de fórmula de área por slot (dentro da aba)
      const SLOTS_AREAS  = [
        { privP:21, privA:22, privT:23, com:24, total:25, coef:26 },
        { privP:30, privA:31, privT:32, com:33, total:34, coef:35 },
        { privP:39, privA:40, privT:41, com:42, total:43, coef:44 },
        { privP:48, privA:49, privT:50, com:51, total:52, coef:53 },
      ];
      // Linha base do Q IV-B para a folha 1 (unidade 0 = linha 17)
      const IVBLINE_BASE = 17;

      // Calcular número de folhas necessárias
      const totalFolhas = Math.ceil(unidades.length / 4);

      // Função auxiliar: escrever valores de área diretamente nas células da aba
      // Escreve apenas os valores que chegam no payload (privativa principal e acessória).
      // Área de uso comum e total real dependem de fórmulas do Q II — são calculadas pelo Excel.
      // Coeficiente também vem de fórmula — não escrevemos para não corromper o cálculo.
      function escreverAreasQV2(ws, unidadesDaFolha, baseSlots) {
        unidadesDaFolha.forEach(function(u, slotIdx) {
          if (slotIdx >= 4) return;
          const s = baseSlots[slotIdx];
          function w(ref, val) {
            const cell = ws.getCell(ref);
            // Não sobrescrever células com fórmulas — elas calculam automaticamente
            if (cell.formula || cell.sharedFormula) return;
            cell.value = val;
            celulasEscritas++;
          }
          const privP = toFloat(u.AREA_PRIV_PRINCIPAL || u.AREA_PP);
          const privA = toFloat(u.AREA_PRIV_ACESS     || u.AREA_PA);
          const privT = privP + privA;
          // Escrever apenas os valores que chegam do formulário
          if (privP > 0) w('I' + s.privP, privP);
          if (privA > 0) w('I' + s.privA, privA);
          if (privT > 0) w('I' + s.privT, privT);
          // Uso comum (UCNP25 + UCNPD26) — chegam no payload por unidade
          const ucnp  = toFloat(u.UCNP25 || u.AREA_COM_NP_COB_PAD || 0);
          const ucnpd = toFloat(u.UCNPD26 || u.AREA_COM_NP_COB_DIF || 0);
          const comNP = ucnp + ucnpd;
          // Área proporcional de uso comum não está no payload — deixar para fórmula do Excel
          // Escrever apenas UCNP se houver valor (maioria dos casos é zero)
          if (comNP > 0) w('I' + s.com, comNP);
        });
      }

      // Função: limpar todos os 4 slots de uma aba (necessário antes de preencher cópia)
      function limparSlotsQV2(ws) {
        SLOTS_TEXTO.forEach(function(linhaTexto) {
          var cA = ws.getCell('A' + linhaTexto);
          var cB = ws.getCell('B' + linhaTexto);
          if (!cA.formula) cA.value = null;
          if (!cB.formula) { cB.value = null; cB.alignment = { wrapText: true, vertical: 'top' }; }
        });
      }

      // Função: preencher os slots de texto e designação em uma aba
      function preencherSlotsQV2(ws, unidadesDaFolha) {
        unidadesDaFolha.forEach(function(u, slotIdx) {
          if (slotIdx >= 4) return; // máximo 4 por aba
          const linhaTexto = SLOTS_TEXTO[slotIdx];
          const celDesig = ws.getCell('A' + linhaTexto);
          if (!celDesig.formula) celDesig.value = u.DESIGNACAO || '';
          const celTxt = ws.getCell('B' + linhaTexto);
          if (!celTxt.formula) {
            const txt = String(u.TEXTO_DESCRITIVO || u.TEXTO || '').trim();
            celTxt.value = txt;
            celTxt.alignment = { wrapText: true, vertical: 'top' };
          }
        });
      }

      // Estrutura de slots de área dentro de cada aba (fixas — não mudam entre folhas)
      const QV2_SLOTS = [
        { privP:21, privA:22, privT:23, com:24, total:25, coef:26 },
        { privP:30, privA:31, privT:32, com:33, total:34, coef:35 },
        { privP:39, privA:40, privT:41, com:42, total:43, coef:44 },
        { privP:48, privA:49, privT:50, com:51, total:52, coef:53 },
      ];

      // Folha 1 — aba original: preencher unidades 0-3
      const grupo0 = unidades.slice(0, 4);
      preencherSlotsQV2(shQV2original, grupo0);
      escreverAreasQV2(shQV2original, grupo0, QV2_SLOTS);
      // Coeficiente 5 casas decimais na aba original
      QV2_SLOTS.forEach(function(s) {
        shQV2original.getCell('I' + s.coef).numFmt = '0.00000';
      });

      // Folhas adicionais — copiar aba original e ajustar fórmulas
      for (let folhaIdx = 1; folhaIdx < totalFolhas; folhaIdx++) {
        const grupoUnids = unidades.slice(folhaIdx * 4, folhaIdx * 4 + 4);
        if (grupoUnids.length === 0) break;

        // Criar nova aba com nome sequencial
        const nomeNovaAba = `Q V-2-f ${shQV2original.name.match(/f\s*(\d+)/i)?.[1] ? (parseInt(shQV2original.name.match(/f\s*(\d+)/i)[1]) + folhaIdx) : ('8-' + folhaIdx)}`;

        // Copiar a aba original: ExcelJS não tem clone nativo,
        // então adicionamos a aba após a original e copiamos propriedades
        const novaAba = wb.addWorksheet(nomeNovaAba);

        // Copiar dimensões de colunas
        shQV2original.columns.forEach(function(col, idx) {
          if (col.width) novaAba.getColumn(idx + 1).width = col.width;
        });

        // Copiar célula por célula (valores, fórmulas, estilos, mesclagens)
        shQV2original.eachRow({ includeEmpty: true }, function(row, rowNumber) {
          const novaLinha = novaAba.getRow(rowNumber);
          novaLinha.height = row.height;
          row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
            const novaCell = novaLinha.getCell(colNumber);
            // Copiar valor ou fórmula
            if (cell.formula) {
              novaCell.formula = cell.formula;
            } else if (cell.value !== null && cell.value !== undefined) {
              novaCell.value = cell.value;
            }
            // Copiar estilo (clone para não compartilhar referência)
            if (cell.style) {
              novaCell.style = JSON.parse(JSON.stringify(cell.style));
            }
          });
        });

        // Copiar mesclagens
        shQV2original.mergeCells; // garantir acesso
        if (shQV2original.model && shQV2original.model.merges) {
          shQV2original.model.merges.forEach(function(mergeRange) {
            try { novaAba.mergeCells(mergeRange); } catch(e) { /* ignorar se já mesclado */ }
          });
        }

        // Limpar slots da aba copiada (que herdou conteúdo da aba original)
        limparSlotsQV2(novaAba);
        // Preencher texto e designação das unidades deste grupo
        preencherSlotsQV2(novaAba, grupoUnids);

        // ── Reescrever fórmulas de área para apontar linhas corretas do Q IV-B ──
        // Folha original: slot 0 → Q IV-B linha 17, slot 1 → 18, slot 2 → 19, slot 3 → 20
        // Folha extra folhaIdx=1: slot 0 → 21, slot 1 → 22, etc.
        const nomeQIVB = "'Q IV B - f 6 - OU'";
        grupoUnids.forEach(function(u, slotIdx) {
          if (slotIdx >= 4) return;
          const s = QV2_SLOTS[slotIdx];
          const qivbLinha = IVBLINE_BASE + (folhaIdx * 4) + slotIdx;
          function setF(celRef, formula) {
            const cell = novaAba.getCell(celRef);
            cell.value = { formula: formula };
          }
          setF('I' + s.privP, nomeQIVB + '!B' + qivbLinha);
          setF('I' + s.privA, nomeQIVB + '!C' + qivbLinha);
          setF('I' + s.privT, 'SUM(I' + s.privP + ':I' + s.privA + ')');
          setF('I' + s.com,   nomeQIVB + '!E' + qivbLinha);
          setF('I' + s.total, 'SUM(I' + s.privT + ':I' + s.com + ')');
          setF('I' + s.coef,  nomeQIVB + '!G' + qivbLinha);
          // Garantir formato 4 casas no coeficiente
          novaAba.getCell('I' + s.coef).numFmt = '0.00000';
        });

        // ── Limpar slots sobrando na última aba (sem unidade correspondente) ──
        for (var slotExtra = grupoUnids.length; slotExtra < 4; slotExtra++) {
          var sE = QV2_SLOTS[slotExtra];
          // Limpar designação e texto
          var cA = novaAba.getCell('A' + SLOTS_TEXTO[slotExtra]);
          var cB = novaAba.getCell('B' + SLOTS_TEXTO[slotExtra]);
          cA.value = null;
          cB.value = null;
          // Limpar TODAS as células de área e rótulos do slot vazio
          ['privP','privA','privT','com','total','coef'].forEach(function(campo) {
            novaAba.getCell('I' + sE[campo]).value = null;
          });
          // Limpar rótulos (B col) e fills de toda a região do slot
          var primeiraLinha = SLOTS_TEXTO[slotExtra] - 1;
          var ultimaLinha = sE.coef + 1;
          for (var r = primeiraLinha; r <= ultimaLinha; r++) {
            ['A','B','C','D','E','F','G','H','I'].forEach(function(col) {
              var c2 = novaAba.getCell(col + r);
              if (!c2.formula && !c2.sharedFormula) {
                c2.value = null;
              }
              try { c2.fill = { type: 'pattern', pattern: 'none' }; } catch(e) {}
            });
          }
        }

        celulasEscritas += grupoUnids.length * 2; // designação + texto por unidade
      }

      if (totalFolhas > 1) {
        // ── P9: Reposicionar abas Q V-2 extras logo após a aba original ──────
        // Método: renumerar orderNo de todas as abas e reconstruir _worksheets
        try {
          // Montar lista ordenada: todas as abas na ordem desejada
          const abasQV2Extras = [];
          for (let f = 1; f < totalFolhas; f++) {
            const nomeExtra = 'Q V-2-f ' + (parseInt((shQV2original.name.match(/\d+$/)||['8'])[0]) + f);
            const wsE = wb.getWorksheet(nomeExtra);
            if (wsE) abasQV2Extras.push(wsE);
          }
          // Construir ordem desejada: tudo antes de Q VI-1, inserir extras após Q V-2 original
          const ordemFinal = [];
          wb._worksheets.forEach(function(ws) {
            if (!ws) return;
            if (abasQV2Extras.indexOf(ws) >= 0) return; // pular extras (serão inseridas)
            ordemFinal.push(ws);
            // Após a aba Q V-2 original, inserir as extras
            if (ws.name === shQV2original.name) {
              abasQV2Extras.forEach(function(e) { ordemFinal.push(e); });
            }
          });
          // Renumerar orderNo e reconstruir _worksheets
          const novoWS = [undefined]; // slot 0
          ordemFinal.forEach(function(ws, idx) {
            ws.orderNo = idx;
            novoWS.push(ws);
          });
          wb._worksheets = novoWS;
          avisos.push('Q V-2: abas reposicionadas com sucesso.');
        } catch(ePos) { avisos.push('Reposicionamento Q V-2: ' + ePos.message); }
        avisos.push('Q V-2: ' + unidades.length + ' unidades → ' + totalFolhas + ' folha(s) gerada(s).');
      }

    } else if (!shQV2original) avisos.push('Aba Q V-2 não encontrada.');

    // ── QUADROS VI-1, VI-2, VII, VIII ─────────────────────────────────────
    const shQVI1 = findSheet(/Q[\s_-]*VI-1/i)||findSheet(/Q[\s_-]*VI[\s_-]*1/i);
    if (shQVI1 && equipamentos.lista) equipamentos.lista.forEach(item => {
      if (item.linha) {
        if (item.tipo)      sw(shQVI1,`B${item.linha}`,item.tipo);
        if (item.acabamento)sw(shQVI1,`C${item.linha}`,item.acabamento);
        if (item.detalhes)  sw(shQVI1,`E${item.linha}`,item.detalhes);
      }
    });

    const shQVI2 = findSheet(/Q[\s_-]*VI-2/i)||findSheet(/Q[\s_-]*VI[\s_-]*2/i);
    if (shQVI2 && equipamentos.acabamentos) equipamentos.acabamentos.forEach(item => {
      if (item.linha) {
        if (item.tipo)      sw(shQVI2,`B${item.linha}`,item.tipo);
        if (item.acabamento)sw(shQVI2,`C${item.linha}`,item.acabamento);
        if (item.detalhes)  sw(shQVI2,`E${item.linha}`,item.detalhes);
      }
    });

    const shQVII = findSheet(/Q[\s_-]*VII\b/i);
    if (shQVII && acabamentos.privativos) acabamentos.privativos.forEach(item => {
      if (item.linha) {
        if (item.nome)    sw(shQVII,`A${item.linha}`,item.nome);
        if (item.piso)    sw(shQVII,`B${item.linha}`,item.piso);
        if (item.paredes) sw(shQVII,`E${item.linha}`,item.paredes);
        if (item.tetos)   sw(shQVII,`H${item.linha}`,item.tetos);
      }
    });

    const shQVIII = findSheet(/Q[\s_-]*VIII\b/i);
    if (shQVIII && acabamentos.comuns) acabamentos.comuns.forEach(item => {
      if (item.linha) {
        if (item.nome)    sw(shQVIII,`A${item.linha}`,item.nome);
        if (item.piso)    sw(shQVIII,`B${item.linha}`,item.piso);
        if (item.paredes) sw(shQVIII,`E${item.linha}`,item.paredes);
        if (item.tetos)   sw(shQVIII,`H${item.linha}`,item.tetos);
      }
    });

    // ── CABEÇALHOS: LOCAL DO IMÓVEL, INCORPORADOR, RT, CAU, DATA, TELEFONE ──
    // Monta endereço completo da obra
    const cidadeVal = config.CIDADE || '';
    const ufVal2 = (function(){
      var u = config.UF && String(config.UF).trim() && config.UF !== '[PENDENTE]' ? String(config.UF).trim() : '';
      if (!u) { var p = String(config.CIDADE_UF||'').split(' - '); u = p.length > 1 ? p[p.length-1].trim() : ''; }
      return u;
    })();
    const endObra = (config.LOCAL_CONSTRUCAO||config.ENDERECO_COMPLETO||'') + (cidadeVal ? ' - ' + cidadeVal : '') + (ufVal2 ? '/' + ufVal2 : '');
    const nomeIncorp = config.NOME_INCORPORADOR||config.NOME_INCORP||'';
    const prefRT = (config.TITULO_RT||'Arquiteto').startsWith('Arq') ? 'Arq.' : 'Eng.';
    const nomeRT = prefRT + ' ' + (config.NOME_RT||'');
    const cauRT = config.REGISTRO_RT||'';
    const telRT = config.TELEFONE_RT||config.TEL_RT||'';
    const dataDoc = config.DATA_DOCUMENTO || new Date().toLocaleDateString('pt-BR');

    // Mapeamento definitivo por aba: {local, incorp, rt, cau, data, tel, limpar[]}
    const cabMap = [
      { re:/Q[\s_-]*I[\s_-]*[-–]/i,  local:'B8',  incorp:'B11', rt:'L11', cau:'K13', data:'C12', tel:'L12', limpar:['B12'] },
      { re:/Q[\s_-]*II[\s_-]*[-–]/i, local:'C5',  incorp:'B7',  rt:'M7',  cau:'M9',  data:'J7',  tel:'M8',  limpar:['I7'] },
      { re:/Q[\s_-]*III/i,           local:'D7',  incorp:'C12', rt:'H12', cau:'H13', data:'B13', tel:null,  limpar:[] },
      { re:/Q[\s_-]*IV[\s_-]*A/i,    local:'B8',  incorp:'B10', rt:'H10', cau:'H11', data:'B11', tel:'H12', limpar:['M6','B7','I7','H9'] },
      { re:/Q[\s_-]*IV[\s_-]*B/i,    local:'B8',  incorp:'B10', rt:'F10', cau:'F11', data:'B11', tel:null,  limpar:['B7','H7'] },
      { re:/Q[\s_-]*V-1/i,           local:'C7',  incorp:'B12', rt:'G12', cau:'G13', data:'C13', tel:null,  limpar:['B13'] },
      { re:/Q[\s_-]*VI-1/i,          local:'B8',  incorp:'B12', rt:'D12', cau:'D13', data:'D14', tel:null,  limpar:[] },
      { re:/Q[\s_-]*VI-2/i,          local:'B8',  incorp:'B12', rt:'D12', cau:'D13', data:'D14', tel:null,  limpar:[] },
      { re:/Q[\s_-]*VII\b/i,         local:'B8',  incorp:'B12', rt:'G12', cau:'G13', data:'C13', tel:null,  limpar:['B13','D12','D13'] },
      { re:/Q[\s_-]*VIII\b/i,        local:'B6',  incorp:'B9',  rt:'G9',  cau:'G10', data:'B10', tel:null,  limpar:['B12','D12','B13'] },
    ];

    // Aplicar cabeçalhos em cada aba encontrada
    cabMap.forEach(function(m) {
      const ws = findSheet(m.re);
      if (!ws) return;
      // 1. Limpar fórmulas residuais (força null mesmo em fórmulas)
      m.limpar.forEach(function(cel) { forceWrite(ws, cel, null); });
      // 2. Escrever campos
      forceWrite(ws, m.local, endObra);
      // 3. Alinhamento vertical centralizado na célula de endereço
      ws.getCell(m.local).alignment = { vertical: 'middle', wrapText: true };
      forceWrite(ws, m.incorp, nomeIncorp);
      forceWrite(ws, m.rt, nomeRT);
      forceWrite(ws, m.cau, cauRT);
      forceWrite(ws, m.data, dataDoc);
      if (m.tel) forceWrite(ws, m.tel, telRT);
      celulasEscritas += 5 + (m.tel ? 1 : 0);
    });

    // Cabeçalhos nas abas Q V-2 (original + extras)
    wb.worksheets.forEach(function(ws) {
      if (!/Q[\s_-]*V-2/i.test(ws.name)) return;
      forceWrite(ws, 'C7', endObra);
      ws.getCell('C7').alignment = { vertical: 'middle', wrapText: true };
      forceWrite(ws, 'B12', nomeIncorp);
      forceWrite(ws, 'G12', nomeRT);
      forceWrite(ws, 'G13', cauRT);
      forceWrite(ws, 'B13', null);  // limpar fórmula residual
      forceWrite(ws, 'C13', dataDoc);
      celulasEscritas += 5;
    });

    // ── Total de folhas ────────────────────────────────────────────────────
    const totalFolhasWB = wb.worksheets.length - 1; // -1 por DADOS_REF
    if (shRef) sw(shRef,'B6',totalFolhasWB);
    if (shInf) sw(shInf,'F36',`${totalFolhasWB}`);

    // ── NUMERAÇÃO DE FOLHAS — dinâmica com base na ordem final das abas ──
    // Cada aba recebe seu número sequencial e o total
    // Mapeamento: {regex da aba → {num: célula do nº, total: célula do total}}
    const folhaMap = [
      { re:/Inf Prel/i,          num:'N1',  total:null },
      { re:/Q[\s_-]*I[\s_-]*[-–]/i, num:null, total:null },  // Q I não tem célula de número no template
      { re:/Q[\s_-]*II/i,        num:null,  total:null },     // Q II idem
      { re:/Q[\s_-]*III/i,       num:'M5',  total:'M6' },
      { re:/Q[\s_-]*IV[\s_-]*A/i,num:'N4',  total:'N5' },
      { re:/Q[\s_-]*IV[\s_-]*B/i,num:'I5',  total:'I6' },
      { re:/Q[\s_-]*V-1/i,       num:'K5',  total:'K6' },
      { re:/Q[\s_-]*V-2/i,       num:'K5',  total:'K6' },
      { re:/Q[\s_-]*VI-1/i,      num:'J5',  total:'J6' },
      { re:/Q[\s_-]*VI-2/i,      num:'J5',  total:'J6' },
      { re:/Q[\s_-]*VII\b/i,     num:'M5',  total:'M6' },
      { re:/Q[\s_-]*VIII\b/i,    num:'M4',  total:'M5' },
    ];
    // Percorrer abas na ordem atual (já reposicionadas) e numerar
    let numFolha = 1;
    wb.worksheets.forEach(function(ws) {
      if (/DADOS_REF/i.test(ws.name)) return; // pular DADOS_REF
      // Encontrar mapeamento correspondente
      for (let fm = 0; fm < folhaMap.length; fm++) {
        if (folhaMap[fm].re.test(ws.name)) {
          if (folhaMap[fm].num) forceWrite(ws, folhaMap[fm].num, numFolha);
          if (folhaMap[fm].total) forceWrite(ws, folhaMap[fm].total, totalFolhasWB);
          break;
        }
      }
      numFolha++;
    });

    const outputBuffer = await wb.xlsx.writeBuffer();
    res.json({ sucesso: true, celulas_escritas: celulasEscritas, avisos, xlsx_base64: Buffer.from(outputBuffer).toString('base64') });

  } catch (err) {
    console.error('Erro /preencher-xlsx:', err);
    res.status(500).json({ erro: 'Erro interno ao processar a NBR.', detalhe: err.message });
  }
});

app.listen(PORT, () => console.log(`PAAC DocServer v7.0 rodando na porta ${PORT}`));
