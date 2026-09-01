const express = require('express');
const cors = require('cors');

const { graphGet, graphPost, graphPatch, SITE_ID } = require('./lib/graph');
const { assinarToken, conferirSenha, exigirOperadorLogado, exigirChaveAdmin, hashSenha } = require('./lib/auth');
const { carregarItensChecklist, carregarColunasChecklist, montarFieldsChecklist } = require('./lib/checklist');

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
// TEMPORÁRIO — diagnóstico da lista ItensChecklist (protegido por X-Admin-Key).
// Remover depois que o problema do checklist vazio for resolvido.
// ---------------------------------------------------------------------
app.get('/api/admin/debug-itens-checklist', async (req, res) => {
  try { exigirChaveAdmin(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  try {
    const LISTA_ITENS_CHECKLIST_ID = process.env.LISTA_ITENS_CHECKLIST_ID;
    const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_ITENS_CHECKLIST_ID + '/items?$expand=fields&$top=999');
    const linhas = (d.value || []).map(function (item) { return item.fields; });
    return res.json({ totalLinhas: linhas.length, amostra: linhas.slice(0, 5) });
  } catch (err) {
    return res.status(500).json({ erro: 'Falha ao ler ItensChecklist: ' + err.message });
  }
});

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
  try {
    const d = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999');
    const item = (d.value || []).find(function (it) {
      return String((it.fields || {}).Usuario || '').trim().toLowerCase() === usuario;
    });
    const ativo = item && String((item.fields.Ativo === undefined ? 'Sim' : item.fields.Ativo)).trim().toLowerCase();
    const inativo = ativo === 'não' || ativo === 'nao' || ativo === 'false';
    if (!item || inativo || !item.fields.SenhaHash || !conferirSenha(senha, item.fields.SenhaHash)) {
      return res.status(401).json({ erro: 'Usuário ou senha inválidos.' });
    }
    const nome = String(item.fields.Title || usuario).trim();
    const token = assinarToken({ sub: item.id, usuario, nome, role: 'operador' });
    return res.json({ token, nome });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao verificar login. Tente novamente em instantes.' });
  }
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
  dadosPorRotulo.Operador = sessao.nome;

  try {
    const colunasInternas = await carregarColunasChecklist();
    const { fields, faltando } = montarFieldsChecklist(dadosPorRotulo, colunasInternas);
    if (!fields.Equipamento) {
      return res.status(400).json({ erro: 'Informe o equipamento antes de enviar.' });
    }
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
      return {
        id: item.id,
        nome: f.Title || '',
        usuario: f.Usuario || '',
        temLoginTablet: !!f.Usuario,
        ativo: ativo !== 'não' && ativo !== 'nao' && ativo !== 'false'
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('Comprex tablet backend rodando na porta ' + PORT);
});
