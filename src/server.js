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
const ctrlCompartilhados = require('./lib/ctrlCompartilhados');

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

    // Último equipamento que ESTE operador (o do token) usou — pra pré-selecionar
    // sozinho no Checklist/Parte Diária, sem precisar escolher de novo todo dia.
    // Guardado pelo código do equipamento (não um id), porque é o que bate entre
    // esta lista de frota e a lista de equipamentos da Controladoria.
    const itemOperadorLogado = (operadoresResp.value || []).find(function (item) {
      return String(item.id) === String(sessao.sub);
    });
    const meuUltimoEquipamento = itemOperadorLogado ? String((itemOperadorLogado.fields || {}).UltimoEquipamentoCodigo || '').trim() : '';

    return res.json({
      operadorLogado: sessao.nome,
      meuUltimoEquipamento,
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
// PATCH /api/operador/ultimo-equipamento  { equipamentoCodigo }  (Authorization: Bearer <token>)
// Usado pelo Checklist do tablet (que só tem o token do operador, não o id
// dele no SharePoint) pra guardar qual foi o último equipamento usado — pra
// pré-selecionar sozinho da próxima vez, em qualquer aparelho. Mesma ideia
// (e mesmo campo) do endpoint /api/ctrl/operadores/:id/ultimo-equipamento
// usado pela Parte Diária dentro da Controladoria.
// ---------------------------------------------------------------------
app.patch('/api/operador/ultimo-equipamento', async (req, res) => {
  let sessao;
  try { sessao = exigirOperadorLogado(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  const equipamentoCodigo = String((req.body || {}).equipamentoCodigo || '').trim();
  try {
    await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + sessao.sub + '/fields', { UltimoEquipamentoCodigo: equipamentoCodigo });
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao salvar último equipamento: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// GET /api/foto-equipamento/:codigo  (Authorization: Bearer <token>)
// Devolve { url } com o link de download da foto do equipamento (biblioteca
// "FotosEquipamentos" no SharePoint), ou { url: null } se não tiver foto —
// nunca dá erro só por falta de foto, pra não travar o checklist.
// ---------------------------------------------------------------------
app.get('/api/foto-equipamento/:codigo', async (req, res) => {
  // Aceita tanto o token do operador de campo (tablet) quanto a sessão
  // Microsoft/Portal da Controladoria — a Controladoria também usa esta foto
  // (mesma ideia do Checklist) ao escolher um equipamento em Lançar Parte Diária.
  try { await exigirSessaoControladoria(req); }
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
    // Esta MESMA lista do SharePoint também serve o cadastro de Operadores da
    // Controladoria (matrícula/cargo, sem login nenhum) — por isso só entram
    // aqui os registros que têm de fato um login de tablet (campo Usuario
    // preenchido). Sem esse filtro, todo operador cadastrado só pela
    // Controladoria aparecia aqui como se tivesse login.
    const lista = (d.value || [])
      .filter(function (item) { return !!(item.fields || {}).Usuario; })
      .map(function (item) {
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

// ---------------------------------------------------------------------
// DELETE /api/admin/operadores/:id  (X-Admin-Key)
// Remove o LOGIN do tablet desse operador (apaga usuário/senha) — mas não
// o registro inteiro, porque essa mesma lista do SharePoint também é o
// cadastro de Operadores da Controladoria (matrícula/cargo) e pode ter
// apropriações/checklists históricos vinculados ao nome dele. Um delete
// físico do item apagaria esse histórico de cadastro também; em vez disso
// só limpamos Usuario/SenhaHash, o que já tira essa pessoa da lista de
// "Login Tablet" (ver filtro no GET acima) e libera o nome de usuário pra
// ser reaproveitado depois.
// ---------------------------------------------------------------------
app.delete('/api/admin/operadores/:id', async (req, res) => {
  try { exigirChaveAdmin(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const id = req.params.id;
  try {
    await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + id + '/fields', {
      Usuario: '', SenhaHash: ''
    });
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao remover operador: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// PATCH /api/ctrl/operadores/:id/senha  { usuario?, senha }
// Define/redefine o login do tablet de um operador diretamente pela tela de
// Cadastros da Controladoria (mesmo registro da lista Operadores, mesma
// coisa que o botão "Redefinir senha" do Portal faz) — só que sem precisar
// abrir o painel separado "Login Tablet". Só administrador da Controladoria
// pode usar (mesmo tratamento de permissão que Cadastros já tem).
// ---------------------------------------------------------------------
app.patch('/api/ctrl/operadores/:id/senha', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  if (sessao.perfil !== 'admin') {
    return res.status(403).json({ erro: 'Só administrador pode definir login de tablet.' });
  }
  const usuario = String((req.body || {}).usuario || '').trim().toLowerCase();
  const senha = String((req.body || {}).senha || '');
  if (!senha) return res.status(400).json({ erro: 'Informe a senha.' });
  if (senha.length < 6) return res.status(400).json({ erro: 'A senha precisa ter pelo menos 6 caracteres.' });
  const id = req.params.id;
  try {
    const campos = { SenhaHash: hashSenha(senha) };
    if (usuario) {
      const existentes = await graphGet('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items?$expand=fields&$top=999');
      // Antes a mensagem só dizia "já existe" sem apontar QUEM já está com esse
      // usuário — quem via a tela não tinha como saber se é um operador ativo
      // de verdade, um duplicado esquecido, ou um cadastro desativado. Agora
      // nomeia o dono atual do usuário, pra dar pra resolver (renomear o outro,
      // ou perceber que é um registro duplicado que devia ter sido removido).
      const conflito = (existentes.value || []).find(function (it) {
        return String(it.id) !== String(id) && String((it.fields || {}).Usuario || '').trim().toLowerCase() === usuario;
      });
      if (conflito) {
        const nomeConflito = String((conflito.fields || {}).Title || '').trim();
        const ativoConflito = String(conflito.fields.Ativo === undefined ? 'Sim' : conflito.fields.Ativo).trim().toLowerCase();
        const inativoConflito = ativoConflito === 'não' || ativoConflito === 'nao' || ativoConflito === 'false';
        return res.status(409).json({
          erro: 'Esse nome de usuário já está em uso' + (nomeConflito ? (' por "' + nomeConflito + '"' + (inativoConflito ? ' (desativado)' : '')) : '') + '. Escolha outro nome de usuário.'
        });
      }
      campos.Usuario = usuario;
    }
    await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + id + '/fields', campos);
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao definir login: ' + err.message });
  }
});

// ---------------------------------------------------------------------
// PATCH /api/ctrl/operadores/:id/ultimo-equipamento  { equipamentoCodigo }
// Guarda qual foi o último equipamento que esse operador usou (no Checklist
// ou na Parte Diária) — só pra pré-selecionar sozinho da próxima vez, em
// qualquer aparelho que ele entrar (fica salvo no cadastro dele, não no
// tablet). Aceita tanto o token do próprio operador de campo (só pode
// atualizar o dele mesmo) quanto a sessão da Controladoria como admin (ex.:
// suporte ajustando por outra pessoa).
// ---------------------------------------------------------------------
app.patch('/api/ctrl/operadores/:id/ultimo-equipamento', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  const id = req.params.id;
  const podeComoAdmin = sessao.perfil === 'admin';
  const ehOProprioOperador = sessao.perfil === 'campo' && String(sessao.operadorId) === String(id);
  if (!podeComoAdmin && !ehOProprioOperador) {
    return res.status(403).json({ erro: 'Você só pode atualizar o último equipamento do seu próprio cadastro.' });
  }
  const equipamentoCodigo = String((req.body || {}).equipamentoCodigo || '').trim();
  try {
    await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_OPERADORES_ID + '/items/' + id + '/fields', { UltimoEquipamentoCodigo: equipamentoCodigo });
    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao salvar último equipamento: ' + err.message });
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
    return { email: null, nome: payload.nome, perfil: 'campo', telas: null, operadorId: payload.sub };
  } catch (eOperador) {
    // não era um token nosso — segue pra tentativa como conta Microsoft
  }
  const conta = await exigirContaMicrosoft(req);
  // permissao.tipo: 'admin' (tudo liberado) ou 'limitado' (só as telas da
  // Controladoria marcadas em "Gerenciar Acessos" no Portal — mesma tela que
  // já é usada pra Checklists/Movimentações). telas null = sem restrição.
  const permissao = await exigirPermissaoControladoria(conta.email);
  return { email: conta.email, nome: conta.nome, perfil: permissao.tipo, telas: permissao.telas, operadorId: null };
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

async function carregarListaGenericaCtrl(nomeLogico) {
  if (ctrlCompartilhados.ehListaCompartilhada(nomeLogico)) {
    return ctrlCompartilhados.carregar(nomeLogico);
  }
  const listaId = idDaLista(nomeLogico);
  if (!listaId) return null;
  const d = await graphGet('/sites/' + SITE_ID + '/lists/' + listaId + '/items?$expand=fields&$top=999');
  return (d.value || []).map(function (item) {
    try { return JSON.parse((item.fields || {}).Dados || 'null'); }
    catch (e) { return null; }
  }).filter(function (x) { return x !== null; });
}

async function carregarConfigGenericaCtrl(chave) {
  const listaId = idDaLista('configuracoesCtrl');
  if (!listaId) return null;
  const achado = await localizarItemPorId(listaId, chave);
  if (!achado) return null;
  try { return JSON.parse((achado.fields || {}).Dados || 'null'); }
  catch (e) { return null; }
}

// GET /api/ctrl/dados-iniciais?modoCampo=1  -> { perfil, listas:{...}, configs:{...} }
//
// Junta num pacote só as listas que o operador de campo precisa pra abrir a
// Parte Diária — antes eram ~11 requisições separadas (uma por lista, cada
// uma disparando sua própria chamada ao SharePoint), cada uma tendo que
// esperar a mesma instância do Render acordar; agora é 1 requisição só, com
// as chamadas ao SharePoint em paralelo aqui dentro (mesmo padrão que já
// funciona bem pro Checklist em /api/dados-iniciais). PRECISA vir definida
// ANTES de "/api/ctrl/:lista" abaixo, senão o Express entenderia
// "dados-iniciais" como se fosse o nome de uma lista.
const CTRL_BUNDLE_CHAVES_CAMPO = ['equipamentos', 'ordens', 'apropriacoes', 'apropriacoesRascunhos', 'operadores', 'tarifasHora', 'tarifasProducao', 'motivosParada', 'escalas', 'clientes'];
const CTRL_BUNDLE_CHAVES_EXTRA = ['tarifasVerba', 'usuarios', 'producoes', 'producaoNotas', 'statusFaturamento', 'custosOrdem', 'logAuditoria', 'locais'];
const CTRL_BUNDLE_CONFIG_CAMPO = ['cliente_numeros'];
const CTRL_BUNDLE_CONFIG_EXTRA = ['cliente_detalhes', 'materiais_pesagem', 'servicos_hora', 'servicos_verba', 'metas_mensais', 'descricoes_boletim', 'usuarios_ordens_restritas'];
app.get('/api/ctrl/dados-iniciais', async (req, res) => {
  let sessao;
  try { sessao = await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  const modoCampo = String(req.query.modoCampo) === '1';
  const chaves = modoCampo ? CTRL_BUNDLE_CHAVES_CAMPO : CTRL_BUNDLE_CHAVES_CAMPO.concat(CTRL_BUNDLE_CHAVES_EXTRA);
  const configs = modoCampo ? CTRL_BUNDLE_CONFIG_CAMPO : CTRL_BUNDLE_CONFIG_CAMPO.concat(CTRL_BUNDLE_CONFIG_EXTRA);

  try {
    const [listasResp, configsResp] = await Promise.all([
      Promise.all(chaves.map(function (k) { return carregarListaGenericaCtrl(k).catch(function () { return null; }); })),
      Promise.all(configs.map(function (k) { return carregarConfigGenericaCtrl(k).catch(function () { return null; }); }))
    ]);
    const listas = {};
    chaves.forEach(function (k, i) { listas[k] = listasResp[i]; });
    const configsOut = {};
    configs.forEach(function (k, i) { configsOut[k] = configsResp[i]; });
    return res.json({ perfil: sessao.perfil, listas: listas, configs: configsOut });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ erro: 'Erro ao carregar dados iniciais da Controladoria: ' + err.message });
  }
});

// GET /api/ctrl/:lista  -> { itens: [...] }
app.get('/api/ctrl/:lista', async (req, res) => {
  try { await exigirSessaoControladoria(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }

  // Equipamentos e Operadores são cadastro COMPARTILHADO com o tablet/Portal
  // (mesmas listas Frota/Operadores) — ver src/lib/ctrlCompartilhados.js.
  if (ctrlCompartilhados.ehListaCompartilhada(req.params.lista)) {
    try {
      const itens = await ctrlCompartilhados.carregar(req.params.lista);
      return res.json({ itens: itens });
    } catch (err) {
      console.error(err);
      return res.status(500).json({ erro: 'Erro ao carregar "' + req.params.lista + '": ' + err.message });
    }
  }

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

  const item = (req.body || {}).item;
  if (!item || item.id === undefined || item.id === null || item.id === '') {
    return res.status(400).json({ erro: 'Envie "item" com um campo "id".' });
  }

  if (ctrlCompartilhados.ehListaCompartilhada(req.params.lista)) {
    try {
      await ctrlCompartilhados.criar(req.params.lista, item);
      return res.json({ ok: true });
    } catch (err) {
      console.error(err);
      return res.status(500).json({ erro: 'Erro ao criar item em "' + req.params.lista + '": ' + err.message });
    }
  }

  const listaId = idDaLista(req.params.lista);
  if (!listaId) return res.status(500).json({ erro: 'Lista "' + req.params.lista + '" não está configurada no backend.' });

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

  const item = (req.body || {}).item;
  if (!item) return res.status(400).json({ erro: 'Envie "item".' });

  if (ctrlCompartilhados.ehListaCompartilhada(req.params.lista)) {
    try {
      await ctrlCompartilhados.atualizar(req.params.lista, req.params.id, item);
      return res.json({ ok: true });
    } catch (err) {
      console.error(err);
      return res.status(500).json({ erro: 'Erro ao atualizar item em "' + req.params.lista + '": ' + err.message });
    }
  }

  const listaId = idDaLista(req.params.lista);
  if (!listaId) return res.status(500).json({ erro: 'Lista "' + req.params.lista + '" não está configurada no backend.' });

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

  if (ctrlCompartilhados.ehListaCompartilhada(req.params.lista)) {
    try {
      await ctrlCompartilhados.excluir(req.params.lista, req.params.id);
      return res.json({ ok: true });
    } catch (err) {
      console.error(err);
      return res.status(500).json({ erro: 'Erro ao excluir item em "' + req.params.lista + '": ' + err.message });
    }
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

// ---------------------------------------------------------------------
// Migração única dos dados que ainda estão no Supabase (tabela
// "app_storage", chave/valor) para as listas novas do SharePoint. Roda
// AQUI (no próprio backend, já em produção) em vez de precisar de Node.js
// no computador de quem está fazendo a migração — basta visitar a URL com
// a chave de admin (a mesma X-Admin-Key do painel de operadores).
//
//   POST /api/admin/migrar-supabase?chave=...   -> inicia (responde na hora,
//        a migração continua rodando em segundo plano — pode demorar
//        minutos se houver muitos lançamentos)
//   GET  /api/admin/migrar-supabase?chave=...   -> mostra o andamento/resultado
// ---------------------------------------------------------------------
const SUPABASE_URL_MIGRACAO = process.env.SUPABASE_URL || 'https://klpgnaefqhsjbreiqjwe.supabase.co';
const SUPABASE_KEY_MIGRACAO = process.env.SUPABASE_KEY || 'sb_publishable_gDdp4YFLFyggngUronBBsQ_7IawVcmO';
const CHAVES_LISTA_MIGRACAO = {
  equipamentos_v6: 'equipamentos', operadores_v4: 'operadores', ordens_v9: 'ordens',
  apropriacoes_v7: 'apropriacoes', apropriacoes_rascunhos_v1: 'apropriacoesRascunhos',
  tarifas_hora_v2: 'tarifasHora', tarifas_producao_v3: 'tarifasProducao', tarifas_verba_v1: 'tarifasVerba',
  motivos_parada_v1: 'motivosParada', escalas_v1: 'escalas', usuarios_v1: 'usuarios',
  producoes_v1: 'producoes', producao_notas_v1: 'producaoNotas', status_faturamento_v1: 'statusFaturamento',
  custos_ordem_v1: 'custosOrdem', clientes_v2: 'clientes', log_auditoria_v1: 'logAuditoria'
};
const CHAVES_CONFIG_MIGRACAO = {
  cliente_detalhes_v1: 'cliente_detalhes', cliente_numeros_v3: 'cliente_numeros',
  materiais_pesagem_v1: 'materiais_pesagem', servicos_hora_v1: 'servicos_hora',
  servicos_verba_v1: 'servicos_verba', metas_mensais_v1: 'metas_mensais',
  descricoes_boletim_v1: 'descricoes_boletim', usuarios_ordens_restritas_v2: 'usuarios_ordens_restritas'
};
function ehListaDeNomesMigracao(chaveInterna) { return chaveInterna === 'clientes_v2'; }

let estadoMigracao = { rodando: false, concluida: false, log: [], erro: null };

function exigirChaveAdminSimples(req) {
  const chave = req.headers['x-admin-key'] || req.query.chave || '';
  if (!process.env.ADMIN_API_KEY || chave !== process.env.ADMIN_API_KEY) {
    const erro = new Error('Chave de administração ausente ou inválida.');
    erro.status = 401;
    throw erro;
  }
}

async function idsExistentesMigracao(listaId) {
  const existentes = new Set();
  let caminho = '/sites/' + SITE_ID + '/lists/' + listaId + '/items?$select=id&$expand=fields($select=Title)&$top=999';
  while (caminho) {
    const d = await graphGet(caminho);
    (d.value || []).forEach(function (item) {
      const titulo = (item.fields || {}).Title;
      if (titulo !== undefined && titulo !== null) existentes.add(String(titulo));
    });
    const proximo = d['@odata.nextLink'];
    caminho = proximo ? proximo.replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  return existentes;
}

async function rodarMigracaoSupabase() {
  estadoMigracao = { rodando: true, concluida: false, log: [], erro: null };
  const registrar = (linha) => { estadoMigracao.log.push(linha); console.log('[migração] ' + linha); };
  try {
    registrar('Lendo dados do Supabase...');
    const r = await fetch(SUPABASE_URL_MIGRACAO + '/rest/v1/app_storage?select=key,value', {
      headers: { apikey: SUPABASE_KEY_MIGRACAO, Authorization: 'Bearer ' + SUPABASE_KEY_MIGRACAO }
    });
    if (!r.ok) throw new Error('Falha ao ler do Supabase (' + r.status + '): ' + (await r.text()));
    const linhas = await r.json();
    const dados = {};
    linhas.forEach((linha) => { dados[linha.key] = linha.value; });
    registrar(Object.keys(dados).length + ' chaves encontradas no Supabase.');

    for (const chaveInterna of Object.keys(CHAVES_LISTA_MIGRACAO)) {
      const nomeLogico = CHAVES_LISTA_MIGRACAO[chaveInterna];
      if (!(chaveInterna in dados)) continue;
      const listaId = idDaLista(nomeLogico);
      if (!listaId) { registrar('PULADO: lista "' + nomeLogico + '" não tem LISTA_CTRL_..._ID configurada.'); continue; }
      const valor = dados[chaveInterna];
      if (!Array.isArray(valor)) { registrar('PULADO: "' + chaveInterna + '" não é uma lista.'); continue; }
      const itens = ehListaDeNomesMigracao(chaveInterna)
        ? valor.filter(Boolean).map((nome) => ({ id: nome, nome: nome }))
        : valor;
      const existentes = await idsExistentesMigracao(listaId);
      let criados = 0, pulados = 0, semId = 0, erros = 0;
      for (const item of itens) {
        if (!item || item.id === undefined || item.id === null || item.id === '') { semId++; continue; }
        const id = String(item.id);
        if (existentes.has(id)) { pulados++; continue; }
        try {
          await graphPost('/sites/' + SITE_ID + '/lists/' + listaId + '/items', { fields: { Title: id, Dados: JSON.stringify(item) } });
          existentes.add(id);
          criados++;
        } catch (e) { erros++; registrar('  erro ao migrar item "' + id + '" de "' + nomeLogico + '": ' + e.message); }
      }
      registrar('"' + nomeLogico + '": ' + criados + ' criados, ' + pulados + ' já existiam, ' + semId + ' sem id, ' + erros + ' com erro.');
    }

    const listaConfigId = idDaLista('configuracoesCtrl');
    for (const chaveInterna of Object.keys(CHAVES_CONFIG_MIGRACAO)) {
      const chaveConfig = CHAVES_CONFIG_MIGRACAO[chaveInterna];
      if (!(chaveInterna in dados)) continue;
      const valor = dados[chaveInterna];
      if (valor === null || valor === undefined) continue;
      if (!listaConfigId) { registrar('PULADO: configuração "' + chaveConfig + '" — lista de configurações não está configurada.'); continue; }
      const existentes = await idsExistentesMigracao(listaConfigId);
      if (existentes.has(chaveConfig)) { registrar('"' + chaveConfig + '": já existia — não sobrescrito.'); continue; }
      try {
        await graphPost('/sites/' + SITE_ID + '/lists/' + listaConfigId + '/items', { fields: { Title: chaveConfig, Dados: JSON.stringify(valor) } });
        registrar('"' + chaveConfig + '": criado.');
      } catch (e) { registrar('  erro ao migrar configuração "' + chaveConfig + '": ' + e.message); }
    }

    registrar('Migração concluída.');
    estadoMigracao.concluida = true;
  } catch (e) {
    estadoMigracao.erro = e.message;
    estadoMigracao.log.push('ERRO FATAL: ' + e.message);
  } finally {
    estadoMigracao.rodando = false;
  }
}

app.post('/api/admin/migrar-supabase', (req, res) => {
  try { exigirChaveAdminSimples(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  if (estadoMigracao.rodando) return res.json({ ok: true, mensagem: 'Já está rodando — consulte com GET.' });
  rodarMigracaoSupabase(); // não espera terminar — roda em segundo plano
  return res.json({ ok: true, mensagem: 'Migração iniciada. Consulte o andamento com GET nesta mesma URL.' });
});

app.get('/api/admin/migrar-supabase', (req, res) => {
  try { exigirChaveAdminSimples(req); }
  catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
  // Visitar a URL pelo navegador já é suficiente pra iniciar (não precisa de
  // POST) — assim quem for rodar isso só precisa colar o link uma vez, e
  // apertar F5 depois pra ver o andamento. "?reiniciar=1" força rodar de novo.
  if (!estadoMigracao.rodando && (!estadoMigracao.concluida || req.query.reiniciar)) {
    rodarMigracaoSupabase();
    return res.json({ ok: true, mensagem: 'Migração iniciada — atualize a página em alguns segundos para ver o andamento.' });
  }
  return res.json(estadoMigracao);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('Comprex tablet backend rodando na porta ' + PORT);
});
