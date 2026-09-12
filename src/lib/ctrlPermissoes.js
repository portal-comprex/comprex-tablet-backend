/**
 * Confere se a conta Microsoft já autenticada (ver msalAuth.js) tem
 * permissão de usar a Controladoria, e com qual perfil. Substitui a tabela
 * de usuários que antes vivia no Supabase Auth.
 *
 * A lista "Permissões da Controladoria" no SharePoint tem uma linha por
 * pessoa autorizada: Title = e-mail (minúsculo), Dados = JSON com
 * { perfil: 'controladoria' | 'escritorio' } (mesmos nomes de perfil que o
 * sistema já usa hoje).
 *
 * Cache simples em memória (1 minuto) pra não bater no Graph a cada clique —
 * mudanças de permissão demoram no máximo esse tempo pra valer.
 */
const { graphGet } = require('./graph');
const { idDaLista } = require('./ctrlLists');

const LISTA_PERMISSOES_ENV = 'LISTA_CTRL_PERMISSOES_ID';
let cache = { porEmail: null, carregadoEm: 0 };

async function carregarPermissoes() {
  const agora = Date.now();
  if (cache.porEmail && agora - cache.carregadoEm < 60000) return cache.porEmail;
  const listaId = process.env[LISTA_PERMISSOES_ENV];
  if (!listaId) {
    throw Object.assign(new Error('Backend sem ' + LISTA_PERMISSOES_ENV + ' configurada.'), { status: 500 });
  }
  const SITE_ID = require('./graph').SITE_ID;
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + listaId + '/items?$expand=fields&$top=999');
  const porEmail = {};
  (d.value || []).forEach(function (item) {
    const f = item.fields || {};
    const email = String(f.Title || '').toLowerCase().trim();
    if (!email) return;
    let dados = {};
    try { dados = JSON.parse(f.Dados || '{}'); } catch (e) { dados = {}; }
    porEmail[email] = { perfil: dados.perfil || 'escritorio' };
  });
  cache = { porEmail: porEmail, carregadoEm: agora };
  return porEmail;
}

/** Lança erro (status 403) se o e-mail não estiver autorizado. Devolve { perfil }. */
async function exigirPermissaoControladoria(email) {
  const porEmail = await carregarPermissoes();
  const permissao = porEmail[String(email || '').toLowerCase().trim()];
  if (!permissao) {
    throw Object.assign(new Error('Sua conta (' + email + ') não tem acesso à Controladoria. Peça pra alguém do time te cadastrar.'), { status: 403 });
  }
  return permissao;
}

module.exports = { exigirPermissaoControladoria, MAPA_ENV: LISTA_PERMISSOES_ENV };
