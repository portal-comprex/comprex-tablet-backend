const express = require('express');
const cors = require('cors');

const { graphGet, graphPost, graphPatch, graphDelete, SITE_ID } = require('./lib/graph');
const {
  assinarToken, conferirSenha, exigirOperadorLogado, exigirChaveAdmin, hashSenha,
  statusBloqueioLogin, registrarTentativaFalha, limparTentativas, infoPresenca
} = require('./lib/auth');
const { carregarItensChecklist, carregarColunasChecklist, montarFieldsChecklist } = require('./lib/checklist');
const { buscarUrlFotoEquipamento } = require('./lib/fotos');
const { exigirContaMicrosoft } = require('./lib/msalAuth');
const { exigirPermissaoControladoria } = require('./lib/ctrlPermissoes');
const { idDaLista } = require('./lib/ctrlLists');

const LISTA_OPERADORES_ID = process.env.LISTA_OPERADORES_ID;
const LISTA_FROTA_ID = process.env.LISTA_FROTA_ID;
const LISTA_LOCAIS_ID = process.env.LISTA_LOCAIS_ID;
const LISTA_CONFIGURACOES_ID = process.env.LISTA_CONFIGURACOES_ID;
const LISTA_CHECKLISTS_ID = process.env.LISTA_CHECKLISTS_ID;
const LISTA_PARTE_DIARIA_ID = process.env.LISTA_PARTE_DIARIA_ID;
const TIPOS_PARADA_PADRAO = ['ALMOÇO', 'REVISÃO PREVENTIVA', 'MANUTENÇÃO CORRETIVA', 'PARADA TÉCNICA', 'EM DISPONIBILIDADE'];

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));

// Endpoint simples pra confirmar que o backend está de pé (usado no
// navegador ou pelo Render/Railway pra "health check").
app.get('/', (req, res) => res.json({ ok: true, servico: 'comprex-tablet-backend' }));
app.get('/api/health', (req, res) => res.json({ ok: true }));

// ---------------------------------------------------------------------
// POST /api/login  { usuario, senha } -> { token, nome }
// ---------------------------------------------------------------------
app.post('/api/login', async (req, res) => {
  const body = req.body || {};
  const usuario = String(body.usuario || '').trim().toLowerCase();
  const senha = String(body.senha || '');
  if (!usuario || !senha) {
    return res.status(400).json({ erro: 'Informe usuário e senha.' });
  }
  // Bloqueia tentativas repetidas de senha errada (ataque de força bruta) —
  // depois de LIMITE_TENTATIVAS erros seguidos, esse usuário fica travado
  // por um tempo, mesmo que a senha certa seja informada nesse meio-tempo.
  const bloqueio = statusBloqueioLogin(usuario);
  if (bloqueio.bloqueado) {
    return res.status(429).json({ erro: 'Muitas tentativas erradas. Tente novamente em ' + bloqueio.minutosRestantes + ' minuto(s).' });
  }
  try {
    const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999');
    const item = (d.value || []).find(function (it) {
      return String((it.fields || {}).Usuario || '').trim().toLowerCase() === usuario;
    });
    const ativo = item && String((item.fields.Ativo === undefined ? 'Sim' : item.fields.Ativo)).trim().toLowerCase();
    const inativo = ativo === 'não' || ativo === 'nao' || ativo === 'false';
    if (!item || inativo || !item.fields.SenhaHash || !conferirSenha(senha, item.fields.SenhaHash)) {
      registrarTentativaFalha(usuario);
      return res.status(401).json({ erro: 'Usuário ou senha inválidos.' });
    }
    limparTentativas(usuario);
    const nome = String(item.fields.Title || usuario).trim();
    const token = assinarToken({ sub: item.id, usuario, nome, role: 'operador' });
    return res.json({ token, nome });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao verificar login. Tente novamente em instantes.' });
  }
});

// ---------------------------------------------------------------------
// GET /api/ping  (Authorization: Bearer <token>)
// Só serve pra marcar "presença" (visto pela última vez agora) e confirmar
// que o token ainda é válido — o tablet chama isso periodicamente enquanto
// o operador está logado, sem depender de ele estar realmente enviando
// checklists/parte diária no momento.
// ---------------------------------------------------------------------
app.get('/api/ping', (req, res) => {
  try { exigirOperadorLogado(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  return res.json({ ok: true });
});

// ---------------------------------------------------------------------
// GET /api/dados-iniciais  (Authorization: Bearer <token>)
// ---------------------------------------------------------------------
async function carregarTiposParada() {
  if (!LISTA_CONFIGURACOES_ID) return TIPOS_PARADA_PADRAO;
  try {
    const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_CONFIGURACOES_ID + '/items?$expand=fields&$top=999');
    const item = (d.value || []).find(function (it) {
      return String((it.fields || {}).Title || '').trim().toUpperCase() === 'TIPOS_PARADA';
    });
    const valor = item && String((item.fields || {}).Valor || '').trim();
    if (!valor) return TIPOS_PARADA_PADRAO;
    const lista = valor.split(',').map(function (t) { return t.trim(); }).filter(Boolean);
    return lista.length ? lista : TIPOS_PARADA_PADRAO;
  } catch (e) {
    return TIPOS_PARADA_PADRAO;
  }
}

app.get('/api/dados-iniciais', async (req, res) => {
  let sessao;
  try { sessao = exigirOperadorLogado(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  try {
    const [frotaResp, locaisResp, operadoresResp, checklistDef, tiposParada] = await Promise.all([
      graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items?$expand=fields&$top=999'),
      graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_LOCAIS_ID + '/items?$expand=fields&$top=999'),
      graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999'),
      carregarItensChecklist(),
      carregarTiposParada()
    ]);

    const frotas = (frotaResp.value || []).map(function (item) {
      const f = item.fields || {};
      return {
        nome: String(f.Title || '').trim(),
        tipoItem: f.TipoItem === 'implemento' ? 'implemento' : 'equipamento',
        local: f.LocalAtual || '',
        descricao: f.Descricao || ''
      };
    }).filter(function (f) { return f.nome; });

    const locais = (locaisResp.value || []).map(function (item) {
      return String((item.fields || {}).Title || '').trim();
    }).filter(Boolean);

    const operadores = (operadoresResp.value || []).map(function (item) {
      const f = item.fields || {};
      const ativo = String(f.Ativo === undefined ? 'Sim' : f.Ativo).trim().toLowerCase();
      if (ativo === 'não' || ativo === 'nao' || ativo === 'false') return null;
      return String(f.Title || '').trim();
    }).filter(Boolean);

    return res.json({
      operadorLogado: sessao.nome,
      frotas,
      locais,
      operadores,
      tiposParada,
      checklist: checklistDef
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao carregar dados iniciais: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// GET /api/foto-equipamento/:codigo  (Authorization: Bearer <token>)
// Devolve { url } com o link de download da foto do equipamento (biblioteca
// "FotosEquipamentos" no SharePoint), ou { url: null } se não tiver foto —
// nunca dá erro só por falta de foto, pra não travar o checklist.
// ---------------------------------------------------------------------
app.get('/api/foto-equipamento/:codigo', async (req, res) => {
  try { exigirOperadorLogado(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  try {
    const url = await buscarUrlFotoEquipamento(req.params.codigo);
    return res.json({ url });
  } catch (err) {
    console.error(err);
    return res.json({ url: null });
  }
});

// ---------------------------------------------------------------------
// POST /api/checklist  (Authorization: Bearer <token>)
// ---------------------------------------------------------------------
app.post('/api/checklist', async (req, res) => {
  let sessao;
  try { sessao = exigirOperadorLogado(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const body = req.body || {};
  const dadosPorRotulo = body.dadosPorRotulo;
  if (!dadosPorRotulo || typeof dadosPorRotulo !== 'object') {
    return res.status(400).json({ erro: 'Envie "dadosPorRotulo" com os dados do checklist.' });
  }
  // Confere o valor ANTES de converter pros nomes internos do SharePoint —
  // depois da conversão a chave já não se chama mais "Equipamento" (vira o
  // nome interno real da coluna, ex.: "field_2"), então checar fields.Equipamento
  // ali na frente sempre falhava mesmo com o campo preenchido certinho.
  if (!String(dadosPorRotulo.Equipamento || '').trim()) {
    return res.status(400).json({ erro: 'Informe o equipamento antes de enviar.' });
  }
  dadosPorRotulo.Operador = sessao.nome;

  try {
    const colunasInternas = await carregarColunasChecklist();
    const { fields, faltando } = montarFieldsChecklist(dadosPorRotulo, colunasInternas);
    await graphPost('/sites/' + SITE_ID + '/lists/' + LISTA_CHECKLISTS_ID + '/items', { fields });
    return res.json({ ok: true, colunasNaoEncontradas: faltando });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Falha ao gravar o checklist: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// POST /api/parte-diaria  (Authorization: Bearer <token>)
// ---------------------------------------------------------------------
app.post('/api/parte-diaria', async (req, res) => {
  let sessao;
  try { sessao = exigirOperadorLogado(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const body = req.body || {};
  const equipamento = String(body.equipamento || '').trim().toUpperCase();
  const data = String(body.data || '').trim();
  const local = String(body.local || '').trim();
  const observacoes = String(body.observacoes || '').trim();
  const atividades = Array.isArray(body.atividades) ? body.atividades : [];
  const paradas = Array.isArray(body.paradas) ? body.paradas : [];
  const totais = body.totais || {};
  const horimetro = atividades.length ? String(atividades[0].horIni || '').trim() : '';

  if (!equipamento || !data || !local) {
    return res.status(400).json({ erro: 'Preencha equipamento, data e local.' });
  }
  if (!atividades.length && !paradas.length) {
    return res.status(400).json({ erro: 'Adicione ao menos uma atividade ou uma parada.' });
  }
  if (!horimetro) {
    return res.status(400).json({ erro: 'Informe o horímetro inicial na primeira atividade.' });
  }

  const num = function (v) { return typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.')) || 0; };
  const fields = {
    Equipamento: equipamento,
    Horimetro: horimetro,
    Data: data + 'T00:00:00Z',
    Operador: sessao.nome,
    Local: local,
    Atividades: JSON.stringify(atividades),
    Paradas: JSON.stringify(paradas),
    Observacoes: observacoes,
    TotalTrabalhadas: num(totais.trabalhadas).toFixed(2).replace('.', ','),
    TotalParadas: num(totais.paradas).toFixed(2).replace('.', ','),
    HorasLiquidas: num(totais.liquidas).toFixed(2).replace('.', ',')
  };

  try {
    await graphPost('/sites/' + SITE_ID + '/lists/' + LISTA_PARTE_DIARIA_ID + '/items', { fields });
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Falha ao gravar a parte diária: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// GET/POST /api/admin/operadores  (X-Admin-Key)
// ---------------------------------------------------------------------
app.get('/api/admin/operadores', async (req, res) => {
  try { exigirChaveAdmin(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  try {
    const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999');
    const lista = (d.value || []).map(function (item) {
      const f = item.fields || {};
      const ativo = String(f.Ativo === undefined ? 'Sim' : f.Ativo).trim().toLowerCase();
      const presenca = infoPresenca(f.Usuario);
      return {
        id: item.id,
        nome: f.Title || '',
        usuario: f.Usuario || '',
        temLoginTablet: !!f.Usuario,
        ativo: ativo !== 'não' && ativo !== 'nao' && ativo !== 'false',
        online: presenca.online,
        ultimoAcesso: presenca.ultimoAcesso // timestamp em ms, ou null se nunca logou desde que o backend subiu
      };
    });
    return res.json({ operadores: lista });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao listar operadores: ' + err.message });
  }
});

app.post('/api/admin/operadores', async (req, res) => {
  try { exigirChaveAdmin(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const body = req.body || {};
  const nome = String(body.nome || '').trim();
  const usuario = String(body.usuario || '').trim().toLowerCase();
  const senha = String(body.senha || '');
  if (!nome || !usuario || !senha) {
    return res.status(400).json({ erro: 'Informe nome, usuário e senha.' });
  }
  if (senha.length < 6) {
    return res.status(400).json({ erro: 'A senha precisa ter pelo menos 6 caracteres.' });
  }
  try {
    const existentes = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999');
    const jaExiste = (existentes.value || []).some(function (it) {
      return String((it.fields || {}).Usuario || '').trim().toLowerCase() === usuario;
    });
    if (jaExiste) {
      return res.status(409).json({ erro: 'Já existe um operador com esse nome de usuário.' });
    }
    const criado = await graphPost('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items', {
      fields: { Title: nome, Usuario: usuario, SenhaHash: hashSenha(senha), Ativo: 'Sim' }
    });
    return res.json({ ok: true, id: criado.id });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao criar operador: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// PATCH /api/admin/operadores/:id  { senha?, ativo? }  (X-Admin-Key)
// ---------------------------------------------------------------------
app.patch('/api/admin/operadores/:id', async (req, res) => {
  try { exigirChaveAdmin(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const id = req.params.id;
  const body = req.body || {};
  const campos = {};
  if (typeof body.senha === 'string' && body.senha) {
    if (body.senha.length < 6) return res.status(400).json({ erro: 'A senha precisa ter pelo menos 6 caracteres.' });
    campos.SenhaHash = hashSenha(body.senha);
  }
  if (typeof body.ativo === 'boolean') {
    campos.Ativo = body.ativo ? 'Sim' : 'Não';
  }
  if (!Object.keys(campos).length) {
    return res.status(400).json({ erro: 'Nada para atualizar — envie "senha" e/ou "ativo".' });
  }
  try {
    await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + id + '/fields', campos);
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao atualizar operador: ' + err.message });
  }
});

// =======================================================================
// CONTROLADORIA — dados que antes viviam no Supabase (ver Etapa 3).
// Autenticação: confia na sessão Microsoft que a pessoa já tem no Portal
// (nada de segunda senha) + uma lista de permissões no SharePoint dizendo
// quem pode entrar e com qual perfil.
// =======================================================================
// A tela da Controladoria agora é aberta em dois lugares: dentro do Portal
// (pessoa com conta Microsoft) e dentro do tablet de campo (operador com o
// login próprio do tablet, sem conta Microsoft nenhuma). Por isso a sessão
// aceita os dois tipos de token — primeiro tenta como operador de campo
// (mais rápido, não depende de rede da Microsoft); se não for esse tipo de
// token, tenta como conta Microsoft do Portal.
async function exigirSessaoControladoria(req) {
  try {
    const payload = exigirOperadorLogado(req);
    return { email: null, nome: payload.nome, perfil: 'campo' };
  } catch (eOperador) {
    // não era um token nosso — segue pra tentativa como conta Microsoft
  }
  const conta = await exigirContaMicrosoft(req);
  const permissao = await exigirPermissaoControladoria(conta.email);
  return { email: conta.email, nome: conta.nome, perfil: permissao.perfil };
}

// Operador de campo só pode ESCREVER nas duas listas que são dele mesmo
// (o próprio lançamento de parte diária) — as demais (ordens, tarifas,
// cadastros...) ele só LÊ, pra preencher o formulário. Perfis de escritório/
// controladoria continuam com acesso completo, como sempre.
function permiteEscritaControladoria(perfil, nomeLista) {
  if (perfil !== 'campo') return true;
  return nomeLista === 'apropriacoes' || nomeLista === 'apropriacoesRascunhos';
}

// GET /api/ctrl/sessao  (Authorization: Bearer <token MSAL do Portal>)
// Confirma que a pessoa pode entrar na Controladoria e com qual perfil —
// é a "tela de login" da Controladoria, só que sem pedir nada digitado.
app.get('/api/ctrl/sessao', async (req, res) => {
  try {
    const sessao = await exigirSessaoControladoria(req);
    return res.json(sessao);
  } catch (e) {
    return res.status(e.status || 401).json({ erro: e.message });
  }
});

async function localizarItemPorId(listaId, idLogico) {
  const filtro = encodeURIComponent("fields/Title eq '" + String(idLogico).replace(/'/g, "''") + "'");
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + listaId + '/items?$expand=fields&$filter=' + filtro);
  return (d.value || [])[0] || null;
}

// GET /api/ctrl/:lista  -> { itens: [...] }
app.get('/api/ctrl/:lista', async (req, res) => {
  try { await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const listaId = idDaLista(req.params.lista);
  if (!listaId) return res.status(500).json({ erro: 'Lista "' + req.params.lista + '" não está configurada no backend.' });

  try {
    const d = await graphGet('/sites/' + SITE_ID + '/lists/' + listaId + '/items?$expand=fields&$top=999');
    const itens = (d.value || []).map(function (item) {
      try { return JSON.parse((item.fields || {}).Dados || 'null'); }
      catch (e) { return null; }
    }).filter(function (x) { return x !== null; });
    return res.json({ itens: itens });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao carregar "' + req.params.lista + '": ' + err.message });
  }
});

// POST /api/ctrl/:lista  { item }  -> cria um item novo (item.id definido pelo cliente)
app.post('/api/ctrl/:lista', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  if (!permiteEscritaControladoria(sessao.perfil, req.params.lista)) {
    return res.status(403).json({ erro: 'Seu perfil não tem permissão de gravar em "' + req.params.lista + '".' });
  }

  const listaId = idDaLista(req.params.lista);
  if (!listaId) return res.status(500).json({ erro: 'Lista "' + req.params.lista + '" não está configurada no backend.' });

  const item = (req.body || {}).item;
  if (!item || item.id === undefined || item.id === null || item.id === '') {
    return res.status(400).json({ erro: 'Envie "item" com um campo "id".' });
  }
  try {
    await graphPost('/sites/' + SITE_ID + '/lists/' + listaId + '/items', {
      fields: { Title: String(item.id), Dados: JSON.stringify(item) }
    });
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao criar item em "' + req.params.lista + '": ' + err.message });
  }
});

// PATCH /api/ctrl/:lista/:id  { item }  -> substitui o conteúdo de um item existente
app.patch('/api/ctrl/:lista/:id', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  if (!permiteEscritaControladoria(sessao.perfil, req.params.lista)) {
    return res.status(403).json({ erro: 'Seu perfil não tem permissão de gravar em "' + req.params.lista + '".' });
  }

  const listaId = idDaLista(req.params.lista);
  if (!listaId) return res.status(500).json({ erro: 'Lista "' + req.params.lista + '" não está configurada no backend.' });

  const item = (req.body || {}).item;
  if (!item) return res.status(400).json({ erro: 'Envie "item".' });

  try {
    const achado = await localizarItemPorId(listaId, req.params.id);
    if (!achado) return res.status(404).json({ erro: 'Item não encontrado.' });
    await graphPatch('/sites/' + SITE_ID + '/lists/' + listaId + '/items/' + achado.id + '/fields', { Dados: JSON.stringify(item) });
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao atualizar item em "' + req.params.lista + '": ' + err.message });
  }
});

// DELETE /api/ctrl/:lista/:id
app.delete('/api/ctrl/:lista/:id', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  if (!permiteEscritaControladoria(sessao.perfil, req.params.lista)) {
    return res.status(403).json({ erro: 'Seu perfil não tem permissão de excluir em "' + req.params.lista + '".' });
  }

  const listaId = idDaLista(req.params.lista);
  if (!listaId) return res.status(500).json({ erro: 'Lista "' + req.params.lista + '" não está configurada no backend.' });

  try {
    const achado = await localizarItemPorId(listaId, req.params.id);
    if (!achado) return res.json({ ok: true }); // já não existe — segue o jogo
    await graphDelete('/sites/' + SITE_ID + '/lists/' + listaId + '/items/' + achado.id);
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao excluir item em "' + req.params.lista + '": ' + err.message });
  }
});

// ---------------------------------------------------------------------
// "Configurações soltas" da Controladoria — o que no Supabase era um objeto
// único (não uma lista de itens com id): detalhes/números de clientes,
// materiais de pesagem, serviços de hora/verba, metas mensais, descrições
// de boletim, ordens restritas por usuário. Cada uma é UM item aqui,
// identificada pelo próprio nome da configuração.
//   GET  /api/ctrl-config/:chave   -> { valor: <objeto ou null> }
//   PUT  /api/ctrl-config/:chave   { valor }
// ---------------------------------------------------------------------
app.get('/api/ctrl-config/:chave', async (req, res) => {
  try { await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const listaId = idDaLista('configuracoesCtrl');
  if (!listaId) return res.status(500).json({ erro: 'Lista de configurações da Controladoria não está configurada no backend.' });

  try {
    const achado = await localizarItemPorId(listaId, req.params.chave);
    if (!achado) return res.json({ valor: null });
    let valor = null;
    try { valor = JSON.parse((achado.fields || {}).Dados || 'null'); } catch (e) { valor = null; }
    return res.json({ valor: valor });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao carregar configuração "' + req.params.chave + '": ' + err.message });
  }
});

app.put('/api/ctrl-config/:chave', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  if (sessao.perfil === 'campo') {
    return res.status(403).json({ erro: 'Seu perfil não tem permissão de alterar configurações.' });
  }

  const listaId = idDaLista('configuracoesCtrl');
  if (!listaId) return res.status(500).json({ erro: 'Lista de configurações da Controladoria não está configurada no backend.' });

  const valor = (req.body || {}).valor;
  if (valor === undefined) return res.status(400).json({ erro: 'Envie "valor".' });

  try {
    const achado = await localizarItemPorId(listaId, req.params.chave);
    if (achado) {
      await graphPatch('/sites/' + SITE_ID + '/lists/' + listaId + '/items/' + achado.id + '/fields', { Dados: JSON.stringify(valor) });
    } else {
      await graphPost('/sites/' + SITE_ID + '/lists/' + listaId + '/items', {
        fields: { Title: req.params.chave, Dados: JSON.stringify(valor) }
      });
    }
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao salvar configuração "' + req.params.chave + '": ' + err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('Comprex tablet backend rodando na porta ' + PORT);
});
