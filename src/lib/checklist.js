const { graphGet, SITE_ID } = require('./graph');
const { MAPA_CAMPOS } = require('./mapaCampos');

const LISTA_CHECKLISTS_ID = process.env.LISTA_CHECKLISTS_ID;
const LISTA_ITENS_CHECKLIST_ID = process.env.LISTA_ITENS_CHECKLIST_ID;

/** Separa uma célula "várias linhas de texto" tipo "MÍNIMO, VAZIO" em ['MÍNIMO','VAZIO'].
 *  Cópia da mesma função do portal_comprex.html. */
function dividirRespostas(txt) {
  return String(txt || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

/** Lê a lista "ItensChecklist" do SharePoint — mesma lógica de
 *  carregarItensChecklistDoSharePoint() no portal_comprex.html — pra manter
 *  as perguntas do checklist do tablet sempre iguais às do app principal. */
async function carregarItensChecklist() {
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_ITENS_CHECKLIST_ID + '/items?$expand=fields&$top=999');
  const linhas = (d.value || []).map(function (item) { return item.fields; })
    .filter(function (f) {
      const ativo = String(f.Ativo || 'Sim').trim().toLowerCase();
      return ativo !== 'não' && ativo !== 'nao';
    });
  linhas.sort(function (a, b) { return (parseFloat(a.Ordem) || 0) - (parseFloat(b.Ordem) || 0); });
  const itens = [];
  let extraData = null;
  let labelParada = null;
  linhas.forEach(function (f) {
    const titulo = String(f.Title || '').trim();
    if (!titulo) return;
    const opcoes = [];
    dividirRespostas(f.RespostasBoas).forEach(function (v) { opcoes.push([v, 'ok']); });
    dividirRespostas(f.RespostasAtencao).forEach(function (v) { opcoes.push([v, 'at']); });
    dividirRespostas(f.RespostasRuins).forEach(function (v) { opcoes.push([v, 'ru']); });
    dividirRespostas(f.RespostasNaoSeAplica).forEach(function (v) { opcoes.push([v, 'na']); });
    if (!opcoes.length) return;
    itens.push([titulo, opcoes]);
    if (!extraData && String(f.CampoExtraTipo || '').trim().toLowerCase() === 'data') {
      extraData = {
        label: titulo,
        rotulo: String(f.CampoExtraRotulo || 'Data').trim() || 'Data',
        opcionalSe: String(f.CampoExtraOpcionalSe || '').trim()
      };
    }
    if (!labelParada && String(f.EhItemDeParada || '').trim().toLowerCase() === 'sim') {
      labelParada = titulo;
    }
  });
  return { itens, extraData, labelParada };
}

/** Busca as colunas reais da lista "Checklists" e monta displayName -> nome interno.
 *  Mesma lógica de carregarColunasChecklist() no portal_comprex.html. */
async function carregarColunasChecklist() {
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_CHECKLISTS_ID + '/columns');
  const colunasInternas = {};
  (d.value || []).forEach(function (col) {
    colunasInternas[col.displayName.trim().toLowerCase()] = col.name;
  });
  return colunasInternas;
}

/** Monta o payload de "fields" pra gravar o checklist, convertendo rótulo -> nome
 *  interno real via MAPA_CAMPOS + colunas ao vivo. Mesma lógica de
 *  montarFieldsChecklist() no portal_comprex.html — mantém os dois sempre
 *  gravando exatamente da mesma forma. */
function montarFieldsChecklist(dadosPorRotulo, colunasInternas) {
  const fields = {};
  const faltando = [];
  Object.keys(dadosPorRotulo).forEach(function (rotulo) {
    const displayName = MAPA_CAMPOS[rotulo];
    if (!displayName) return; // rótulo desconhecido — ignora, igual ao app principal
    const nomeInterno = colunasInternas[displayName.trim().toLowerCase()];
    if (!nomeInterno) { faltando.push(displayName); return; }
    const valor = dadosPorRotulo[rotulo];
    if (valor !== '' && valor != null) fields[nomeInterno] = valor;
  });
  return { fields, faltando };
}

module.exports = { carregarItensChecklist, carregarColunasChecklist, montarFieldsChecklist, dividirRespostas };
