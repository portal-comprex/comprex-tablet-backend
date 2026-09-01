/**
 * Autenticação de APLICATIVO (client credentials) com o Microsoft Graph.
 *
 * Diferente do portal_comprex.html (que usa o login MSAL da PESSOA logada),
 * aqui não existe usuário Microsoft nenhum — o backend inteiro grava no
 * SharePoint "como o aplicativo", usando um Client ID + Client Secret
 * criados no Entra ID (App Registration) com permissão de APLICATIVO
 * (Sites.ReadWrite.All, tipo Application, com consentimento de admin).
 *
 * Isso é obrigatório: sem isso não tem como o backend gravar Checklists/
 * Parte Diária em nome de operadores que não têm conta Microsoft.
 */
const TENANT_ID = process.env.AAD_TENANT_ID;
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_ID = process.env.SP_SITE_ID;

let tokenCache = { token: null, expiraEm: 0 };

async function obterTokenDeAplicativo() {
  const agora = Date.now();
  if (tokenCache.token && agora < tokenCache.expiraEm - 60000) return tokenCache.token;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Configuração incompleta: defina AAD_TENANT_ID, AAD_CLIENT_ID e AAD_CLIENT_SECRET nas configurações do Function App.');
  }
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  if (!r.ok) {
    const texto = await r.text();
    throw new Error('Falha ao obter token de aplicativo (' + r.status + '): ' + texto);
  }
  const dados = await r.json();
  tokenCache = { token: dados.access_token, expiraEm: agora + dados.expires_in * 1000 };
  return tokenCache.token;
}

async function graphGet(caminho) {
  const token = await obterTokenDeAplicativo();
  const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, {
    headers: { Authorization: 'Bearer ' + token }
  });
  if (!r.ok) {
    const texto = await r.text();
    throw new Error('Graph GET ' + caminho + ' falhou (' + r.status + '): ' + texto);
  }
  return r.json();
}

async function graphPost(caminho, corpo) {
  const token = await obterTokenDeAplicativo();
  const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(corpo)
  });
  if (!r.ok) {
    const texto = await r.text();
    throw new Error('Graph POST ' + caminho + ' falhou (' + r.status + '): ' + texto);
  }
  return r.json();
}

async function graphPatch(caminho, corpo) {
  const token = await obterTokenDeAplicativo();
  const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, {
    method: 'PATCH',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(corpo)
  });
  if (!r.ok) {
    const texto = await r.text();
    throw new Error('Graph PATCH ' + caminho + ' falhou (' + r.status + '): ' + texto);
  }
  return r.json();
}

module.exports = { obterTokenDeAplicativo, graphGet, graphPost, graphPatch, SITE_ID };
