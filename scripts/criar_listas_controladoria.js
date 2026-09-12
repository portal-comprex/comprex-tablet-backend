#!/usr/bin/env node
/**
 * Cria, no SharePoint da Comprex, todas as listas novas que substituem o
 * Supabase da Controladoria (Etapa 3). Cada lista tem só 2 colunas: a
 * "Título" que toda lista do SharePoint já tem, e "Dados" (texto longo)
 * onde o registro inteiro fica em JSON.
 *
 * COMO RODAR:
 *   1. Entre na pasta do backend (tablet-project/backend).
 *   2. Garanta que as variáveis AAD_TENANT_ID, AAD_CLIENT_ID,
 *      AAD_CLIENT_SECRET e SP_SITE_ID estão definidas (as mesmas já usadas
 *      em produção — no Render, ou num arquivo .env local).
 *   3. Rode:  node scripts/criar_listas_controladoria.js
 *
 * PERMISSÃO NECESSÁRIA: o App Registration do backend precisa ter a
 * permissão de aplicativo "Sites.Manage.All" (ou "Sites.ReadWrite.All"),
 * com consentimento de administrador — sem isso o Graph recusa criar listas
 * novas (só continua funcionando o que já existe hoje: frota/operadores/
 * checklists/parte diária).
 *
 * O script é seguro pra rodar mais de uma vez: se uma lista com aquele nome
 * já existir no site, ele pula a criação dela (não duplica).
 *
 * NO FINAL, o script imprime as variáveis de ambiente prontas pra colar no
 * Render (LISTA_CTRL_ORDENS_ID=..., etc.) — sem isso o backend não sabe o
 * ID de cada lista.
 */
const { graphGet, graphPost, SITE_ID } = require('../src/lib/graph');
const { MAPA_LISTAS } = require('../src/lib/ctrlLists');

// Nome que vai aparecer no SharePoint (visível pra quem olhar o site) para
// cada variável de ambiente que o backend espera.
const LISTAS_A_CRIAR = Object.assign({}, MAPA_LISTAS, {
  // não faz parte do MAPA_LISTAS (que é só as "genéricas") — é a lista de
  // permissões de acesso à Controladoria, ver lib/ctrlPermissoes.js.
});
const NOMES_EXIBICAO = {
  equipamentos: 'Controladoria - Equipamentos',
  operadores: 'Controladoria - Operadores',
  ordens: 'Controladoria - Ordens',
  apropriacoes: 'Controladoria - Apropriacoes',
  apropriacoesRascunhos: 'Controladoria - Apropriacoes Rascunhos',
  tarifasHora: 'Controladoria - Tarifas por Hora',
  tarifasProducao: 'Controladoria - Tarifas de Producao',
  tarifasVerba: 'Controladoria - Tarifas de Verba',
  motivosParada: 'Controladoria - Motivos de Parada',
  escalas: 'Controladoria - Escalas',
  producoes: 'Controladoria - Producoes',
  producaoNotas: 'Controladoria - Notas de Producao',
  statusFaturamento: 'Controladoria - Status de Faturamento',
  custosOrdem: 'Controladoria - Custos por Ordem',
  clientes: 'Controladoria - Clientes',
  logAuditoria: 'Controladoria - Log de Auditoria',
  configuracoesCtrl: 'Controladoria - Configuracoes'
};
// Lista separada (não entra no CRUD genérico — ver lib/ctrlPermissoes.js)
const LISTA_PERMISSOES = { chaveEnv: 'LISTA_CTRL_PERMISSOES_ID', nomeExibicao: 'Controladoria - Permissoes de Acesso' };

async function listaJaExiste(nomeExibicao) {
  const d = await graphGet('/sites/' + SITE_ID + '/lists?$select=id,displayName&$top=200');
  const achada = (d.value || []).find(function (l) { return l.displayName === nomeExibicao; });
  return achada ? achada.id : null;
}

async function criarLista(nomeExibicao) {
  const criada = await graphPost('/sites/' + SITE_ID + '/lists', {
    displayName: nomeExibicao,
    list: { template: 'genericList' }
  });
  // Coluna "Dados": texto longo (registro inteiro em JSON). "Título" já existe
  // por padrão em toda lista — é nela que guardamos o id lógico do registro.
  await graphPost('/sites/' + SITE_ID + '/lists/' + criada.id + '/columns', {
    name: 'Dados',
    text: { allowMultipleLines: true, linesForEditing: 6 }
  });
  return criada.id;
}

async function garantirLista(chaveEnv, nomeExibicao) {
  const existenteId = await listaJaExiste(nomeExibicao);
  if (existenteId) {
    console.log('  já existe: ' + nomeExibicao + ' (id ' + existenteId + ')');
    return existenteId;
  }
  const novoId = await criarLista(nomeExibicao);
  console.log('  criada:    ' + nomeExibicao + ' (id ' + novoId + ')');
  return novoId;
}

async function main() {
  console.log('Site: ' + SITE_ID);
  console.log('Criando/confirmando listas da Controladoria...\n');

  const variaveisFinais = {};

  for (const chaveLogica of Object.keys(LISTAS_A_CRIAR)) {
    const chaveEnv = LISTAS_A_CRIAR[chaveLogica];
    const nomeExibicao = NOMES_EXIBICAO[chaveLogica];
    const id = await garantirLista(chaveEnv, nomeExibicao);
    variaveisFinais[chaveEnv] = id;
  }

  const idPermissoes = await garantirLista(LISTA_PERMISSOES.chaveEnv, LISTA_PERMISSOES.nomeExibicao);
  variaveisFinais[LISTA_PERMISSOES.chaveEnv] = idPermissoes;

  console.log('\nPronto. Cole estas variáveis de ambiente no backend (Render):\n');
  Object.keys(variaveisFinais).forEach(function (chave) {
    console.log(chave + '=' + variaveisFinais[chave]);
  });

  console.log('\nAlém dessas, o backend também precisa de PORTAL_CLIENT_ID (o Client ID do');
  console.log('App Registration do Portal, usado pra confirmar a sessão Microsoft — o');
  console.log('mesmo valor que já está em portal_comprex.html/portal/index.html como clientId do MSAL).');

  console.log('\nIMPORTANTE: depois de criar a lista "' + LISTA_PERMISSOES.nomeExibicao + '", cadastre nela');
  console.log('manualmente quem pode acessar a Controladoria: uma linha por pessoa, com o');
  console.log('Título = e-mail da conta Microsoft dela (minúsculo) e a coluna Dados = ');
  console.log('{"perfil":"controladoria"} ou {"perfil":"escritorio"} (mesmos perfis de sempre).');
}

main().catch(function (err) {
  console.error('\nERRO: ' + err.message);
  process.exit(1);
});
