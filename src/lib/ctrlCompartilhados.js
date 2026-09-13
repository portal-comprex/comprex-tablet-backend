/**
 * Cadastro COMPARTILHADO entre o tablet/Portal e a Controladoria.
 *
 * Equipamentos e Operadores deixam de ter uma lista própria e vazia da
 * Controladoria (LISTA_CTRL_EQUIPAMENTOS_ID / LISTA_CTRL_OPERADORES_ID) e
 * passam a ler/gravar DIRETO nas mesmas listas que o tablet já usa (Frota e
 * Operadores) — um cadastro só, editável de qualquer um dos três lugares
 * (tablet, Portal, Controladoria).
 *
 * Duas colunas novas em cada lista do SharePoint (crie-as antes de usar):
 *
 *   Frota:
 *     - TipoEquipamento   (Texto de linha única) — ex.: "ESCAVADEIRA VOLVO EC360"
 *     - ObservacaoCtrl    (Várias linhas de texto)
 *
 *   Operadores:
 *     - Matricula  (Texto de linha única)
 *     - Cargo      (Texto de linha única)
 *
 * Locais entra no mesmo esquema, mas sem coluna nova nenhuma — a lista
 * Locais já é só um Title (nome do local), igual o Frota/Portal já usa pro
 * dropdown de local do Checklist. Vira o dropdown de "Local" dentro de uma
 * Ordem na Controladoria (antes era texto livre).
 *
 * O "id" que a Controladoria usa (inclusive como chave estrangeira em
 * apropriações — equipamentoId/operadorId) passa a ser o ID REAL do item no
 * SharePoint (item.id, uma string), não mais um uid() gerado no navegador.
 * Como as duas listas da Controladoria criadas na Etapa 3 nunca chegaram a
 * ser usadas (vazias), não existe nenhum dado antigo pra migrar.
 *
 * Exclusão:
 *   - Operador: NÃO apaga o item — só marca Ativo="Não" (mesmo padrão que o
 *     Portal já usa pra desativar operador). Preserva login/histórico.
 *   - Equipamento: apaga o item de verdade no Frota. Isso é uma mudança de
 *     comportamento real: excluir um equipamento pela Controladoria agora
 *     também some da lista do Frota/checklist do tablet. Se isso não for
 *     desejado, é fácil trocar pra um soft-delete também (mesmo padrão do
 *     operador) — falar com quem mantém este código antes de usar em
 *     produção se tiver dúvida.
 */
const { graphGet, graphPost, graphPatch, graphDelete, SITE_ID } = require('./graph');

const LISTA_FROTA_ID = process.env.LISTA_FROTA_ID;
const LISTA_OPERADORES_ID = process.env.LISTA_OPERADORES_ID;
const LISTA_LOCAIS_ID = process.env.LISTA_LOCAIS_ID;

const NOMES_COMPARTILHADOS = { equipamentos: true, operadores: true, locais: true };
function ehListaCompartilhada(nomeLogico) {
  return !!NOMES_COMPARTILHADOS[nomeLogico];
}

// ---------------------------------------------------------------------
// Equipamentos <-> lista Frota
// ---------------------------------------------------------------------
function equipamentoDoItem(item) {
  const f = item.fields || {};
  return {
    id: String(item.id),
    codigo: String(f.Title || '').trim(),
    tipo: String(f.TipoEquipamento || '').trim(),
    observacao: String(f.ObservacaoCtrl || '').trim(),
    // Local ATUAL do equipamento (mesmo campo que o tablet/Portal usam pra
    // Movimentações) — a Controladoria usa isso pra achar sozinha a Ordem
    // certa quando alguém escolhe um equipamento em Lançar Parte Diária.
    local: String(f.LocalAtual || '').trim()
  };
}

async function carregarEquipamentos() {
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items?$expand=fields&$top=999');
  return (d.value || []).map(equipamentoDoItem).filter(function (e) { return e.codigo; });
}

async function criarEquipamento(item) {
  await graphPost('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items', {
    fields: {
      Title: String(item.codigo || '').trim(),
      TipoEquipamento: String(item.tipo || '').trim(),
      ObservacaoCtrl: String(item.observacao || '').trim(),
      // Mantém a Frota utilizável no checklist do tablet mesmo pra quem
      // cadastrar só pela Controladoria (o tablet exige TipoItem == "implemento"
      // pra tratar como implemento; qualquer outro valor já conta como
      // "equipamento" normal, então não precisamos escrever nada aqui).
    }
  });
}

async function atualizarEquipamento(idItem, item) {
  await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items/' + idItem + '/fields', {
    Title: String(item.codigo || '').trim(),
    TipoEquipamento: String(item.tipo || '').trim(),
    ObservacaoCtrl: String(item.observacao || '').trim()
  });
}

async function excluirEquipamento(idItem) {
  await graphDelete('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items/' + idItem);
}

// ---------------------------------------------------------------------
// Operadores <-> lista Operadores
// ---------------------------------------------------------------------
function operadorDoItem(item) {
  const f = item.fields || {};
  const ativoTexto = String(f.Ativo === undefined ? 'Sim' : f.Ativo).trim().toLowerCase();
  return {
    id: String(item.id),
    matricula: String(f.Matricula || '').trim(),
    nome: String(f.Title || '').trim(),
    cargo: String(f.Cargo || '').trim(),
    ativo: !(ativoTexto === 'não' || ativoTexto === 'nao' || ativoTexto === 'false')
  };
}

async function carregarOperadores() {
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999');
  return (d.value || []).map(operadorDoItem).filter(function (o) { return o.nome; });
}

async function criarOperador(item) {
  await graphPost('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items', {
    fields: {
      Title: String(item.nome || '').trim(),
      Matricula: String(item.matricula || '').trim(),
      Cargo: String(item.cargo || '').trim(),
      Ativo: item.ativo === false ? 'Não' : 'Sim'
    }
  });
}

async function atualizarOperador(idItem, item) {
  await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + idItem + '/fields', {
    Title: String(item.nome || '').trim(),
    Matricula: String(item.matricula || '').trim(),
    Cargo: String(item.cargo || '').trim(),
    Ativo: item.ativo === false ? 'Não' : 'Sim'
  });
}

async function desativarOperador(idItem) {
  await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + idItem + '/fields', {
    Ativo: 'Não'
  });
}

// ---------------------------------------------------------------------
// Locais <-> lista Locais (só Title — mesmo dropdown do Checklist)
// ---------------------------------------------------------------------
function localDoItem(item) {
  const f = item.fields || {};
  return { id: String(item.id), nome: String(f.Title || '').trim() };
}

async function carregarLocais() {
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_LOCAIS_ID + '/items?$expand=fields&$top=999');
  return (d.value || []).map(localDoItem).filter(function (l) { return l.nome; });
}

async function criarLocal(item) {
  await graphPost('/sites/' + SITE_ID + '/lists/' + LISTA_LOCAIS_ID + '/items', {
    fields: { Title: String(item.nome || '').trim() }
  });
}

async function atualizarLocal(idItem, item) {
  await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_LOCAIS_ID + '/items/' + idItem + '/fields', {
    Title: String(item.nome || '').trim()
  });
}

async function excluirLocal(idItem) {
  await graphDelete('/sites/' + SITE_ID + '/lists/' + LISTA_LOCAIS_ID + '/items/' + idItem);
}

// ---------------------------------------------------------------------
// Fachada única, pelo nome lógico ("equipamentos" | "operadores" | "locais")
// ---------------------------------------------------------------------
async function carregar(nomeLogico) {
  if (nomeLogico === 'equipamentos') return carregarEquipamentos();
  if (nomeLogico === 'locais') return carregarLocais();
  return carregarOperadores();
}
async function criar(nomeLogico, item) {
  if (nomeLogico === 'equipamentos') return criarEquipamento(item);
  if (nomeLogico === 'locais') return criarLocal(item);
  return criarOperador(item);
}
async function atualizar(nomeLogico, idItem, item) {
  if (nomeLogico === 'equipamentos') return atualizarEquipamento(idItem, item);
  if (nomeLogico === 'locais') return atualizarLocal(idItem, item);
  return atualizarOperador(idItem, item);
}
async function excluir(nomeLogico, idItem) {
  if (nomeLogico === 'equipamentos') return excluirEquipamento(idItem);
  // Locais: apaga de verdade (mesmo comportamento do equipamento) — é só um
  // nome, sem histórico dependente dele como o operador tem (login/apontamento).
  if (nomeLogico === 'locais') return excluirLocal(idItem);
  return desativarOperador(idItem);
}

module.exports = { ehListaCompartilhada, carregar, criar, atualizar, excluir };
