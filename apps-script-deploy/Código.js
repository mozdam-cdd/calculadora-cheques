/**
 * Calculadora de Cheques — backend de tokens (Apps Script Web App)
 *
 * SETUP (1 vez só):
 *   1. Editar a constante ADMIN_SECRET abaixo e trocar pelo seu valor real
 *      OU rodar setupSecret('sua_senha') uma vez no editor
 *   2. Implantar como Web App: Execute as = Me, Access = Anyone
 *
 * SHEET_ID está fixo no código (planilha do projeto).
 * ADMIN_SECRET é lido de PropertiesService (não fica em texto plano).
 */

const SHEET_ID = '1nhMgJk4ZQIu0Cn-F4SkI7onZXXMxsZZRsvXaIZd05dI';

/**
 * SETUP — rodar UMA vez no editor antes do primeiro uso.
 * Edite SOMENTE o valor da constante SENHA_ADMIN abaixo, salve e clique em Run.
 * No primeiro Run, o Google pede autorização para acessar o Sheets — autorize.
 * Depois disso o Web App responde para qualquer pessoa.
 */
function setup() {
  const SENHA_ADMIN = 'uf-2026-root';
  PropertiesService.getScriptProperties().setProperty('ADMIN_SECRET', SENHA_ADMIN);
  const nome = SpreadsheetApp.openById(SHEET_ID).getName();
  return 'Setup concluído. Planilha: ' + nome;
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ status: 'erro', erro: 'json_invalido' });
  }

  try {
    switch (body.action) {
      case 'redeem':       return json(redeem(body));
      case 'admin_list':   return json(adminList(body));
      case 'admin_create': return json(adminCreate(body));
      case 'admin_revoke': return json(adminRevoke(body));
      default:             return json({ status: 'erro', erro: 'acao_desconhecida' });
    }
  } catch (err) {
    return json({ status: 'erro', erro: String(err && err.message || err) });
  }
}

function doGet() {
  return json({ status: 'ok', service: 'calculadora-tokens' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getProp(name) {
  return PropertiesService.getScriptProperties().getProperty(name);
}

function checkAdmin(body) {
  const secret = getProp('ADMIN_SECRET');
  return !!secret && body.secret === secret;
}

function abrirPlanilha() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function lerTabela(sheet) {
  const dados = sheet.getDataRange().getValues();
  if (dados.length === 0) return { headers: [], linhas: [] };
  const headers = dados[0].map(h => String(h).trim());
  const linhas = dados.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 };
    headers.forEach((h, j) => obj[h] = row[j]);
    return obj;
  });
  return { headers, linhas };
}

function colIndex(headers, name) {
  const idx = headers.indexOf(name);
  if (idx < 0) throw new Error('coluna_ausente:' + name);
  return idx + 1;
}

function redeem(body) {
  const token = String(body.token || '').trim().toUpperCase();
  const deviceId = String(body.deviceId || '').trim();
  const userAgent = String(body.userAgent || '').slice(0, 500);

  if (!token || !deviceId) return { status: 'invalido' };

  const ss = abrirPlanilha();
  const tokensSheet = ss.getSheetByName('tokens');
  const redSheet = ss.getSheetByName('redencoes');
  if (!tokensSheet || !redSheet) return { status: 'erro', erro: 'aba_ausente' };

  const { linhas: tokensLinhas } = lerTabela(tokensSheet);
  const tokenRow = tokensLinhas.find(r => String(r.token).trim().toUpperCase() === token);
  if (!tokenRow) return { status: 'invalido' };

  const revogado = String(tokenRow.revogado || '').trim().toLowerCase();
  if (revogado === 'sim' || revogado === 'true') return { status: 'revogado' };

  const { headers: redHeaders, linhas: redLinhas } = lerTabela(redSheet);
  const colUltimo = colIndex(redHeaders, 'ultimo_acesso');

  const reden = redLinhas.find(r => String(r.token).trim().toUpperCase() === token);
  const agora = new Date().toISOString();

  if (reden) {
    if (String(reden.device_id).trim() !== deviceId) {
      return { status: 'usado' };
    }
    redSheet.getRange(reden._rowIndex, colUltimo).setValue(agora);
    return { status: 'ok' };
  }

  redSheet.appendRow([token, deviceId, agora, agora, userAgent]);
  return { status: 'ok' };
}

function adminList(body) {
  if (!checkAdmin(body)) return { status: 'unauthorized' };
  const ss = abrirPlanilha();
  const tokensSheet = ss.getSheetByName('tokens');
  const redSheet = ss.getSheetByName('redencoes');
  if (!tokensSheet || !redSheet) return { status: 'erro', erro: 'aba_ausente' };

  const limpa = arr => arr.map(o => {
    const c = {};
    Object.keys(o).forEach(k => { if (k !== '_rowIndex') c[k] = o[k]; });
    return c;
  });

  return {
    status: 'ok',
    tokens:    limpa(lerTabela(tokensSheet).linhas),
    redencoes: limpa(lerTabela(redSheet).linhas)
  };
}

function adminCreate(body) {
  if (!checkAdmin(body)) return { status: 'unauthorized' };
  const token = String(body.token || '').trim().toUpperCase();
  const cliente = String(body.cliente || '').trim();
  const obs = String(body.obs || '').trim();

  if (!token) return { status: 'erro', erro: 'token_vazio' };

  const ss = abrirPlanilha();
  const sheet = ss.getSheetByName('tokens');
  if (!sheet) return { status: 'erro', erro: 'aba_ausente' };

  const { linhas } = lerTabela(sheet);
  if (linhas.some(r => String(r.token).trim().toUpperCase() === token)) {
    return { status: 'erro', erro: 'token_duplicado' };
  }

  sheet.appendRow([token, new Date().toISOString(), cliente, '', obs]);
  return { status: 'ok', token };
}

function adminRevoke(body) {
  if (!checkAdmin(body)) return { status: 'unauthorized' };
  const token = String(body.token || '').trim().toUpperCase();
  if (!token) return { status: 'erro', erro: 'token_vazio' };

  const ss = abrirPlanilha();
  const sheet = ss.getSheetByName('tokens');
  if (!sheet) return { status: 'erro', erro: 'aba_ausente' };

  const { headers, linhas } = lerTabela(sheet);
  const colRev = colIndex(headers, 'revogado');

  const row = linhas.find(r => String(r.token).trim().toUpperCase() === token);
  if (!row) return { status: 'erro', erro: 'token_nao_encontrado' };

  sheet.getRange(row._rowIndex, colRev).setValue('sim');
  return { status: 'ok' };
}
