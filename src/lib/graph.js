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
// Enquanto uma renovação de token está em andamento, guarda a MESMA promise
// aqui — assim, se 24 chamadas a graphGet() chegarem juntas (caso clássico:
// a Controladoria carrega ~24 listas em paralelo) com o token expirado/ainda
// não obtido, todas esperam essa ÚNICA requisição em vez de cada uma disparar
// a sua própria pro Azure AD ao mesmo tempo. Sem isso, 24 pedidos de token
// simultâneos competem e alguns levam mais tempo (ou até tomam throttling do
// Azure AD), o que é uma causa real da Controladoria demorar mais pra abrir
// do que telas que carregam só 1-3 listas (Checklist/Movimentações).
let obtencaoEmAndamento = null;

async function obterTokenDeAplicativo() {
  const agora = Date.now();
  if (tokenCache.token && agora < tokenCache.expiraEm - 60000) return tokenCache.token;
  if (obtencaoEmAndamento) return obtencaoEmAndamento;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Configuração incompleta: defina AAD_TENANT_ID, AAD_CLIENT_ID e AAD_CLIENT_SECRET nas configurações do Function App.');
  }
  obtencaoEmAndamento = (async () => {
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
    tokenCache = { token: dados.access_token, expiraEm: Date.now() + dados.expires_in * 1000 };
    return tokenCache.token;
  })();
  try {
    return await obtencaoEmAndamento;
  } finally {
    obtencaoEmAndamento = null;
  }
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

async function graphDelete(caminho) {
  const token = await obterTokenDeAplicativo();
  const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, {
    method: 'DELETE',
    headers: { Authorization: 'Bearer ' + token }
  });
  if (!r.ok && r.status !== 404) {
    const texto = await r.text();
    throw new Error('Graph DELETE ' + caminho + ' falhou (' + r.status + '): ' + texto);
  }
  return true;
}

module.exports = { obterTokenDeAplicativo, graphGet, graphPost, graphPatch, graphDelete, SITE_ID };
