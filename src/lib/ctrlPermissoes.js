/**
 * Confere se a conta Microsoft já autenticada (ver msalAuth.js) tem
 * permissão de usar a Controladoria, e com quais telas.
 *
 * ATUALIZAÇÃO: a Controladoria deixou de ter sua própria lista de permissões
 * ("Controladoria - Permissoes de Acesso") e passou a usar a MESMA tela
 * "Gerenciar Acessos" que o Portal já tem (Checklists/Movimentações) — uma
 * pessoa só, cadastrada uma vez, com permissões granulares por sistema.
 * Isso é lido da lista "Configuracoes" do SharePoint (a mesma que o Portal
 * já grava, ver portal_comprex.html / LISTA_CONFIGURACOES_ID), chaves:
 *
 *   EMAILS_ADMIN        "fulano@comprex.com.br,ciclana@comprex.com.br"
 *                        -> acesso total a tudo (Checklists, Movimentações
 *                        E Controladoria), igual já funciona hoje.
 *
 *   ACESSOS_LIMITADOS    "fulano@comprex.com.br:checklist,controladoria/apropriacao;..."
 *                        -> cada pessoa tem uma lista de "tokens" separados
 *                        por vírgula. Um token é ou uma seção inteira
 *                        ("controladoria" = todas as telas da Controladoria)
 *                        ou uma tela específica ("controladoria/apropriacao").
 *
 * Cache simples em memória (1 minuto) — mesmo padrão do arquivo anterior.
 */
const { graphGet, SITE_ID } = require('./graph');

const LISTA_CONFIGURACOES_ENV = 'LISTA_CONFIGURACOES_ID';
let cache = { dados: null, carregadoEm: 0 };

async function carregarConfiguracoes() {
  const agora = Date.now();
  if (cache.dados && agora - cache.carregadoEm < 60000) return cache.dados;
  const listaId = process.env[LISTA_CONFIGURACOES_ENV];
  if (!listaId) {
    throw Object.assign(new Error('Backend sem ' + LISTA_CONFIGURACOES_ENV + ' configurada.'), { status: 500 });
  }
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + listaId + '/items?$expand=fields&$top=999');
  const porChave = {};
  (d.value || []).forEach(function (item) {
    const chave = String((item.fields || {}).Title || '').trim().toUpperCase();
    if (chave) porChave[chave] = String((item.fields || {}).Valor || '').trim();
  });
  const emailsAdmin = (porChave['EMAILS_ADMIN'] || '')
    .split(',').map(function (e) { return e.trim().toLowerCase(); }).filter(Boolean);
  const acessosLimitados = {};
  (porChave['ACESSOS_LIMITADOS'] || '').split(';').forEach(function (par) {
    const idx = par.indexOf(':');
    if (idx === -1) return;
    const email = par.slice(0, idx).trim().toLowerCase();
    const tokens = par.slice(idx + 1).split(',').map(function (t) { return t.trim(); }).filter(Boolean);
    if (email && tokens.length) acessosLimitados[email] = tokens;
  });
  cache = { dados: { emailsAdmin: emailsAdmin, acessosLimitados: acessosLimitados }, carregadoEm: agora };
  return cache.dados;
}

/** Lança erro (status 403) se o e-mail não estiver autorizado a usar a
 *  Controladoria. Devolve { tipo: 'admin' | 'limitado', telas: null | string[] }
 *  — telas null = todas liberadas (admin, ou "controladoria" inteiro marcado). */
async function exigirPermissaoControladoria(email) {
  const cfg = await carregarConfiguracoes();
  const emailNorm = String(email || '').toLowerCase().trim();
  if (cfg.emailsAdmin.indexOf(emailNorm) !== -1) {
    return { tipo: 'admin', telas: null };
  }
  const tokens = cfg.acessosLimitados[emailNorm] || [];
  const tokensControladoria = tokens.filter(function (t) {
    return t === 'controladoria' || t.indexOf('controladoria/') === 0;
  });
  if (!tokensControladoria.length) {
    throw Object.assign(new Error(
      'Sua conta (' + email + ') não tem acesso à Controladoria. Peça pra um administrador te liberar em "Gerenciar Acessos".'
    ), { status: 403 });
  }
  const secaoInteiraLiberada = tokensControladoria.indexOf('controladoria') !== -1;
  const telas = secaoInteiraLiberada ? null : tokensControladoria.map(function (t) { return t.split('/')[1]; }).filter(Boolean);
  return { tipo: 'limitado', telas: telas };
}

module.exports = { exigirPermissaoControladoria };
