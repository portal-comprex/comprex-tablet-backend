/**
 * Mapa das listas "genéricas" da Controladoria no SharePoint — substituem o
 * antigo armazenamento no Supabase (tabela única app_storage). Cada uma
 * dessas listas tem só 2 colunas: Title (o id do registro) e Dados (o
 * registro inteiro em JSON) — o mesmo formato que a Parte Diária do tablet
 * já usa pra Atividades/Paradas, então não precisamos recriar manualmente
 * cada campo de cada tipo de cadastro da Controladoria.
 *
 * A chave da esquerda é o nome que o SISPOC usa pra pedir os dados (ver
 * loadList/saveList no parte-diaria.html); a variável de ambiente da direita
 * é o ID real da lista no SharePoint (definido depois de criar as listas —
 * ver build_criar_listas_controladoria.js).
 */
const MAPA_LISTAS = {
  // Nota: por ora estes ficam em listas PRÓPRIAS da Controladoria (mesmo
  // padrão genérico das demais), e não reaproveitando frota/operadores do
  // backend do tablet — unificar os dois cadastros é um passo à parte (os
  // campos são diferentes hoje: um serve o login do tablet, o outro não),
  // fica registrado como próximo passo depois de tirar o Supabase do ar.
  equipamentos: 'LISTA_CTRL_EQUIPAMENTOS_ID',
  operadores: 'LISTA_CTRL_OPERADORES_ID',
  ordens: 'LISTA_CTRL_ORDENS_ID',
  apropriacoes: 'LISTA_CTRL_APROPRIACOES_ID',
  apropriacoesRascunhos: 'LISTA_CTRL_RASCUNHOS_ID',
  tarifasHora: 'LISTA_CTRL_TARIFAS_HORA_ID',
  tarifasProducao: 'LISTA_CTRL_TARIFAS_PRODUCAO_ID',
  tarifasVerba: 'LISTA_CTRL_TARIFAS_VERBA_ID',
  motivosParada: 'LISTA_CTRL_MOTIVOS_PARADA_ID',
  escalas: 'LISTA_CTRL_ESCALAS_ID',
  producoes: 'LISTA_CTRL_PRODUCOES_ID',
  producaoNotas: 'LISTA_CTRL_PRODUCAO_NOTAS_ID',
  statusFaturamento: 'LISTA_CTRL_STATUS_FATURAMENTO_ID',
  custosOrdem: 'LISTA_CTRL_CUSTOS_ORDEM_ID',
  clientes: 'LISTA_CTRL_CLIENTES_ID',
  logAuditoria: 'LISTA_CTRL_LOG_AUDITORIA_ID',
  // "Configurações soltas" que no Supabase eram um objeto único (não uma lista de
  // itens com id) — cliente_detalhes, cliente_numeros, materiais de pesagem,
  // serviços de hora/verba, metas mensais, descrições de boletim, ordens
  // restritas por usuário. Cada uma vira UM item nesta lista, identificada
  // pelo próprio nome da configuração (ver rota /api/ctrl-config).
  configuracoesCtrl: 'LISTA_CTRL_CONFIG_ID'
};

function idDaLista(nomeLogico) {
  const varAmbiente = MAPA_LISTAS[nomeLogico];
  if (!varAmbiente) return null;
  return process.env[varAmbiente] || null;
}

module.exports = { MAPA_LISTAS, idDaLista };
