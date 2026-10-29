/* Rotas de MANUTENÇÃO do tablet (paradas, intervenções, liberação).
 * Grava nas MESMAS listas do Portal (ManutOcorrencias, ManutPecas, ManutEventos,
 * Frota e Movimentações), então tudo aparece na Visão Geral e na ficha do equipamento.
 *
 * Variáveis de ambiente (todas opcionais — sem elas as listas são achadas pelo nome):
 *   LISTA_MANUT_OCORRENCIAS_ID, LISTA_MANUT_PECAS_ID, LISTA_MANUT_EVENTOS_ID,
 *   LISTA_MOVIMENTACOES_ID, LISTA_MECANICOS_ID
 */
const TZ = 'America/Sao_Paulo';
const NOMES = { ocorrencias: 'ManutOcorrencias', pecas: 'ManutPecas', eventos: 'ManutEventos', mecanicos: 'Mecanicos' };
const ENV_IDS = { ocorrencias: 'LISTA_MANUT_OCORRENCIAS_ID', pecas: 'LISTA_MANUT_PECAS_ID', eventos: 'LISTA_MANUT_EVENTOS_ID', mecanicos: 'LISTA_MECANICOS_ID' };
const NOMES_MOVIMENTACOES = ['Movimentações', 'Movimentacoes', 'Movimentação', 'Movimentacao'];
const ETAPAS = ['Diagnóstico', 'Em manutenção', 'Aguardando peça', 'Reparo externo', 'Serviço Téc Especializado', 'Montagem/Testes'];
// colunas que o Portal exige em cada lista; as demais são opcionais (ignoradas se não existirem)
const OBRIGATORIAS = {
  ocorrencias: ['Title', 'DataEntrada', 'Etapa', 'Motivo', 'Tipo', 'Responsavel', 'PrevisaoRetorno', 'DataSaida', 'Situacao', 'Observacao'],
  pecas: ['Title', 'OcorrenciaId', 'Equipamento', 'Situacao', 'Destino', 'DataEnvio', 'PrevisaoOriginal', 'PrevisaoRetorno', 'RetornoReal', 'Documento', 'Custo'],
  eventos: ['Title', 'OcorrenciaId', 'Equipamento', 'Data', 'Responsavel', 'Realizado'],
  mecanicos: []
};

function criar(deps) {
  const { graphGet, graphPost, graphPatch, SITE_ID, exigirOperadorLogado } = deps;
  const LISTA_FROTA_ID = deps.LISTA_FROTA_ID;
  const cache = { listas: null, ids: {}, cols: {}, mov: undefined };

  function norm(t) { return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase(); }
  function erro(status, msg) { const e = new Error(msg); e.status = status; return e; }

  async function lerTudo(caminho) {
    const todos = [];
    while (caminho) {
      const d = await graphGet(caminho);
      (d.value || []).forEach(function (i) { todos.push(i); });
      const prox = d['@odata.nextLink'];
      caminho = prox ? prox.replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    return todos;
  }
  async function listasDoSite() {
    if (!cache.listas) {
      const d = await graphGet('/sites/' + SITE_ID + '/lists?$select=id,displayName&$top=500');
      cache.listas = {};
      (d.value || []).forEach(function (l) { cache.listas[String(l.displayName || '').trim().toLowerCase()] = l.id; });
    }
    return cache.listas;
  }
  async function idLista(chave) {
    if (cache.ids[chave]) return cache.ids[chave];
    let id = process.env[ENV_IDS[chave]] || null;
    if (!id) {
      const l = await listasDoSite();
      id = l[NOMES[chave].toLowerCase()] || null;
    }
    if (!id) throw erro(500, 'A lista "' + NOMES[chave] + '" não foi encontrada no SharePoint.');
    cache.ids[chave] = id;
    return id;
  }
  async function idMovimentacoes() {
    if (cache.mov !== undefined) return cache.mov;
    let id = process.env.LISTA_MOVIMENTACOES_ID || null;
    if (!id) {
      const l = await listasDoSite();
      for (const n of NOMES_MOVIMENTACOES) { if (l[n.toLowerCase()]) { id = l[n.toLowerCase()]; break; } }
    }
    cache.mov = id;
    return id;
  }

  /* nome interno real de cada coluna (o SharePoint às vezes cria "Tipo0" etc.) */
  async function colunas(chave) {
    if (cache.cols[chave]) return cache.cols[chave];
    const id = await idLista(chave);
    let defs = null;
    try {
      const d = await graphGet('/sites/' + SITE_ID + '/lists/' + id + '/columns?$select=name,displayName,readOnly,text,number,dateTime&$top=500');
      defs = (d.value || []).filter(function (c) { return !c.readOnly || c.name === 'Title'; });
    } catch (e) { defs = null; }
    const mapa = {};
    function achar(L) {
      if (defs === null) return { name: L };
      const n = norm(L);
      return defs.find(function (x) { return norm(x.name) === n; }) || defs.find(function (x) { return norm(x.displayName) === n; })
        || defs.find(function (x) { return new RegExp('^' + n + '[0-9]+$').test(norm(x.name)); }) || null;
    }
    const todos = OBRIGATORIAS[chave].concat(chave === 'ocorrencias'
      ? ['DataInicio', 'Prioridade', 'HoraEntrada', 'HoraSaida', 'HorimetroEntrada', 'HorimetroSaida', 'DataParada', 'HoraParada', 'HoraInicio', 'DestinoLiberacao', 'ContratoDestino', 'DataEnvioContrato', 'HoraEnvioContrato']
      : chave === 'pecas' ? ['TipoPeca'] : chave === 'mecanicos' ? ['Matricula', 'Ativo'] : []);
    todos.forEach(function (L) { const c = achar(L); if (c) mapa[L] = c; });
    cache.cols[chave] = { mapa: mapa, obrig: OBRIGATORIAS[chave] };
    return cache.cols[chave];
  }
  function coagir(def, v) {
    if (v === undefined) return undefined;
    if (def && def.number) { if (v === '' || v === null) return null; return typeof v === 'number' ? v : Number(String(v).replace(',', '.')); }
    if (def && def.text && typeof v === 'number') return String(v);
    return v;
  }
  async function gravar(chave, campos) {
    const c = await colunas(chave), out = {}, faltam = [];
    Object.keys(campos).forEach(function (L) {
      const def = c.mapa[L];
      if (def) { const v = coagir(def, campos[L]); if (v !== undefined) out[def.name] = v; }
      else if (c.obrig.indexOf(L) !== -1) faltam.push(L);
    });
    if (faltam.length) throw erro(500, 'Falta a coluna "' + faltam.join('", "') + '" na lista ' + NOMES[chave] + '.');
    return out;
  }
  async function ler(chave, item) {
    const c = await colunas(chave), f = item.fields || {}, out = {};
    Object.keys(f).forEach(function (k) { out[k] = f[k]; });
    Object.keys(c.mapa).forEach(function (L) { if (c.mapa[L].name !== L && f[c.mapa[L].name] !== undefined) out[L] = f[c.mapa[L].name]; });
    return out;
  }
  async function criarItem(chave, campos) {
    const id = await idLista(chave);
    return graphPost('/sites/' + SITE_ID + '/lists/' + id + '/items', { fields: await gravar(chave, campos) });
  }
  async function atualizarItem(chave, itemId, campos) {
    const id = await idLista(chave);
    return graphPatch('/sites/' + SITE_ID + '/lists/' + id + '/items/' + itemId + '/fields', await gravar(chave, campos));
  }
  async function itensDe(chave) { return lerTudo('/sites/' + SITE_ID + '/lists/' + await idLista(chave) + '/items?$expand=fields&$top=999'); }

  /* ---------- Frota ---------- */
  const frotaAusentes = [];
  async function acharFrota(codigo) {
    const todos = await lerTudo('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items?$expand=fields&$top=999');
    const alvo = String(codigo || '').trim().toUpperCase();
    return todos.find(function (i) { return String((i.fields || {}).Title || '').trim().toUpperCase() === alvo; }) || null;
  }
  async function patchFrota(itemId, campos) {
    const c = Object.assign({}, campos);
    frotaAusentes.forEach(function (k) { delete c[k]; });
    for (let i = 0; i < 8; i++) {
      try { return await graphPatch('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items/' + itemId + '/fields', c); }
      catch (e) {
        const m = /Field '([^']+)' is not recognized/i.exec(String(e && e.message || e));
        if (!m || !(m[1] in c) || m[1] === 'LocalAtual' || m[1] === 'StatusAtual') throw e;
        delete c[m[1]]; if (frotaAusentes.indexOf(m[1]) < 0) frotaAusentes.push(m[1]);
      }
    }
  }
  async function registrarMovimentacao(campos) {
    const id = await idMovimentacoes();
    if (!id) { console.warn('[manutencao] lista de Movimentações não encontrada — movimentação não registrada.'); return; }
    await graphPost('/sites/' + SITE_ID + '/lists/' + id + '/items', { fields: campos });
  }

  /* ---------- datas (Brasil) ---------- */
  function partesSP(iso) {
    const d = iso ? new Date(iso) : new Date();
    const dt = isNaN(d.getTime()) ? new Date() : d;
    const p = {};
    new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
      .formatToParts(dt).forEach(function (x) { p[x.type] = x.value; });
    return { d: p.year + '-' + p.month + '-' + p.day, h: p.hour + ':' + p.minute, ms: dt.getTime() };
  }
  function paraIso(d, h) { return d ? d.split('T')[0] + 'T' + (h && /^\d{1,2}:\d{2}/.test(h) ? h.slice(0, 5) : '00:00') + ':00-03:00' : ''; }
  function ymd(v) { return v ? String(v).split('T')[0] : ''; }
  function num(v) { if (v === null || v === undefined || v === '') return null; const n = Number(String(v).replace(',', '.')); return isFinite(n) ? n : null; }
  function txt(v, max) { const t = String(v == null ? '' : v).trim(); return max && t.length > max ? t.slice(0, max - 1) + '…' : t; }

  async function ocorrenciaPorId(id) {
    const itens = await itensDe('ocorrencias');
    const it = itens.find(function (i) { return String(i.id) === String(id); });
    if (!it) throw erro(404, 'Parada não encontrada — talvez já tenha sido encerrada.');
    const f = await ler('ocorrencias', it);
    if (String(f.Situacao || 'Aberta') === 'Concluída') throw erro(409, 'Essa parada já foi encerrada.');
    return { id: it.id, f: f };
  }
  async function evento(oc, eq, texto, nome, dataYmd) {
    await criarItem('eventos', { Title: txt(texto, 250), OcorrenciaId: String(oc), Equipamento: eq, Data: dataYmd || partesSP().d, Responsavel: nome, Realizado: true });
  }
  function quem(b) { return txt(b.mecanicoNome, 120) || 'não informado'; }
  function auth(handler) {
    return async function (req, res) {
      let sessao;
      try { sessao = exigirOperadorLogado(req); } catch (e) { return res.status(e.status || 401).json({ erro: e.message }); }
      try { await handler(req, res, sessao); }
      catch (err) { console.error(err); res.status(err.status || 500).json({ erro: err.message || 'Erro ao processar.' }); }
    };
  }

  function registrar(app) {
    // ---- abrir parada ----
    app.post('/api/manutencao', auth(async function (req, res, sessao) {
      const b = req.body || {};
      const eq = txt(b.equipamento), local = txt(b.local), motivo = txt(b.motivo, 250);
      if (!eq || !local || !motivo) throw erro(400, 'Informe equipamento, local e motivo.');
      const quando = partesSP(b.paradaEm);
      const frota = await acharFrota(eq);
      if (!frota) throw erro(400, 'Equipamento "' + eq + '" não encontrado na Frota.');
      const abertas = (await itensDe('ocorrencias'));
      for (const it of abertas) {
        const f = await ler('ocorrencias', it);
        if (String(f.Title || '').trim().toUpperCase() === eq.toUpperCase() && String(f.Situacao || 'Aberta') !== 'Concluída') {
          throw erro(409, eq + ' já tem uma parada em aberto. Atualize a existente em "Paradas em aberto".');
        }
      }
      const ff = frota.fields || {};
      const localAntes = ff.LocalAtual || '';
      const jaEstava = String(ff.StatusAtual || '') === 'EM MANUTENÇÃO';
      await patchFrota(frota.id, { LocalAtual: local, StatusAtual: 'EM MANUTENÇÃO', ResponsavelTecnico: quem(b), ObservacaoManutencao: motivo, TipoManutencao: '' });
      if (!jaEstava || localAntes !== local) {
        await registrarMovimentacao({ Equipamento: eq, ItemVinculado: ff.ItemVinculado || '', LocalOrigem: localAntes || '—', LocalDestino: local, Status: 'EM MANUTENÇÃO',
          Data: quando.d + 'T12:00:00', Usuario: sessao.nome || '', Observacao: 'Entrada na manutenção — lançada no tablet', ResponsavelTecnico: quem(b), TipoManutencao: '' });
      }
      const horim = num(b.horimetro);
      await criarItem('ocorrencias', { Title: eq, DataEntrada: quando.d, Etapa: 'Diagnóstico', Motivo: motivo, Tipo: '', Responsavel: quem(b), PrevisaoRetorno: null, Situacao: 'Aberta',
        Observacao: txt(b.observacao), Prioridade: 'Normal', HoraEntrada: quando.h, DataParada: quando.d, HoraParada: quando.h, HorimetroEntrada: horim });
      res.json({ ok: true });
    }));

    // ---- paradas em aberto ----
    app.get('/api/manutencao/abertas', auth(async function (req, res) {
      const [itens, frotas] = await Promise.all([itensDe('ocorrencias'), lerTudo('/sites/' + SITE_ID + '/lists/' + LISTA_FROTA_ID + '/items?$expand=fields&$top=999')]);
      const desc = {};
      frotas.forEach(function (i) { const f = i.fields || {}; desc[String(f.Title || '').trim().toUpperCase()] = f.Descricao || ''; });
      const abertas = [];
      for (const it of itens) {
        const f = await ler('ocorrencias', it);
        if (String(f.Situacao || 'Aberta') === 'Concluída') continue;
        const eq = String(f.Title || '').trim();
        abertas.push({ id: it.id, equipamento: eq, descricao: desc[eq.toUpperCase()] || '', motivo: f.Motivo || '', mecanico: f.Responsavel || '',
          etapa: f.Etapa || 'Diagnóstico', paradaEm: paraIso(ymd(f.DataParada) || ymd(f.DataEntrada), f.HoraParada || f.HoraEntrada) });
      }
      abertas.sort(function (a, b) { return String(a.paradaEm).localeCompare(String(b.paradaEm)); });
      res.json({ abertas: abertas });
    }));

    // ---- atualizar situação (etapa) ----
    app.post('/api/manutencao/etapa', auth(async function (req, res) {
      const b = req.body || {};
      const etapa = String(b.etapa || '');
      if (ETAPAS.indexOf(etapa) === -1) throw erro(400, 'Situação inválida.');
      const oc = await ocorrenciaPorId(b.ocorrenciaId);
      const eq = String(oc.f.Title || b.equipamento || '').trim();
      const hoje = partesSP(b.quando);
      const campos = { Etapa: etapa };
      if (etapa === 'Aguardando peça' && ymd(b.previsao)) campos.PrevisaoRetorno = ymd(b.previsao);
      await atualizarItem('ocorrencias', oc.id, campos);
      if (etapa === 'Aguardando peça' && txt(b.peca)) {
        await criarItem('pecas', { Title: txt(b.peca, 250), OcorrenciaId: String(oc.id), Equipamento: eq, Situacao: 'Aguardando', TipoPeca: 'Compra', Destino: '',
          PrevisaoOriginal: ymd(b.previsao) || null, PrevisaoRetorno: ymd(b.previsao) || null, Custo: 0 });
      }
      await evento(oc.id, eq, 'Situação alterada para "' + etapa + '"' + (etapa === 'Aguardando peça' && txt(b.peca) ? ' — peça: ' + txt(b.peca) : '') + ' — por ' + quem(b), quem(b), hoje.d);
      res.json({ ok: true });
    }));

    // ---- intervenção concluída (equipamento continua parado) ----
    app.post('/api/manutencao/intervencao', auth(async function (req, res) {
      const b = req.body || {};
      if (!txt(b.servico)) throw erro(400, 'Descreva o serviço executado.');
      const oc = await ocorrenciaPorId(b.ocorrenciaId);
      const eq = String(oc.f.Title || b.equipamento || '').trim();
      const fim = partesSP(b.fimEm);
      await evento(oc.id, eq, 'Intervenção concluída' + (b.terceiro ? ' (prestador: ' : ' (') + quem(b) + ') — ' + txt(b.servico), quem(b), fim.d);
      res.json({ ok: true });
    }));

    // ---- novo problema encontrado ----
    app.post('/api/manutencao/novo-problema', auth(async function (req, res) {
      const b = req.body || {};
      const motivo = txt(b.motivo, 250);
      if (!motivo) throw erro(400, 'Informe o motivo do novo problema.');
      const oc = await ocorrenciaPorId(b.ocorrenciaId);
      const eq = String(oc.f.Title || b.equipamento || '').trim();
      const q = partesSP(b.quando);
      const anterior = oc.f.Motivo || '';
      const nota = '[Novo problema ' + q.d.split('-').reverse().join('/') + '] ' + motivo + ' (responsável: ' + quem(b) + ')';
      await atualizarItem('ocorrencias', oc.id, { Etapa: 'Em manutenção', Motivo: motivo, Responsavel: quem(b), Observacao: [oc.f.Observacao || '', nota].filter(Boolean).join('\n') });
      const frota = await acharFrota(eq);
      if (frota) await patchFrota(frota.id, { ResponsavelTecnico: quem(b), ObservacaoManutencao: motivo });
      await evento(oc.id, eq, 'Novo problema encontrado: ' + motivo + ' — responsável: ' + quem(b) + (anterior ? ' (problema anterior: ' + anterior + ')' : ''), quem(b), q.d);
      res.json({ ok: true });
    }));

    // ---- liberar equipamento (encerra a parada) ----
    app.post('/api/manutencao/concluir', auth(async function (req, res, sessao) {
      const b = req.body || {};
      if (!txt(b.servico)) throw erro(400, 'Descreva o serviço executado.');
      const comPend = !!b.comPendencia;
      if (comPend && !txt(b.pendencia)) throw erro(400, 'Descreva a pendência.');
      const oc = await ocorrenciaPorId(b.ocorrenciaId);
      const eq = String(oc.f.Title || b.equipamento || '').trim();
      const lib = partesSP(b.liberadoEm);
      if (lib.d + 'T' + lib.h < ymd(oc.f.DataEntrada) + 'T' + (String(oc.f.HoraEntrada || '00:00').slice(0, 5))) throw erro(400, 'A liberação não pode ser antes da chegada na manutenção.');
      const horimS = num(b.horimetro), horimE = num(oc.f.HorimetroEntrada);
      if (horimS !== null && horimE !== null && horimS < horimE) throw erro(400, 'O horímetro na liberação não pode ser menor que o da entrada (' + horimE + ').');
      const frota = await acharFrota(eq);
      if (frota) {
        const ff = frota.fields || {};
        const local = ff.LocalAtual || '';
        await patchFrota(frota.id, { StatusAtual: 'DISPONÍVEL', TipoManutencao: '', ResponsavelTecnico: '', ObservacaoManutencao: '', PrevisaoReparo: null });
        await registrarMovimentacao({ Equipamento: eq, ItemVinculado: ff.ItemVinculado || '', LocalOrigem: local || '—', LocalDestino: local || '—', Status: 'DISPONÍVEL',
          Data: lib.d + 'T12:00:00', Usuario: sessao.nome || '', Observacao: 'Manutenção concluída — lançada no tablet · disponível na manutenção' + (comPend ? ' · liberado COM pendência: ' + txt(b.pendencia) : ' · liberado sem pendência'),
          TipoManutencao: oc.f.Tipo || '', ResponsavelTecnico: oc.f.Responsavel || '' });
      }
      await atualizarItem('ocorrencias', oc.id, { Situacao: 'Concluída', DataSaida: lib.d, HoraSaida: lib.h, HorimetroSaida: horimS, DestinoLiberacao: 'Disponível na manutenção', ContratoDestino: '',
        Observacao: [oc.f.Observacao || '', 'Conclusão: ' + txt(b.servico), comPend ? '[PENDÊNCIA] ' + txt(b.pendencia).replace(/\n+/g, ' ') : ''].filter(Boolean).join('\n') });
      await evento(oc.id, eq, 'Equipamento liberado ' + (comPend ? 'COM pendência' : 'sem pendência') + ' — por ' + quem(b) + ' — ' + txt(b.servico), quem(b), lib.d);
      res.json({ ok: true });
    }));
  }

  /* lista de mecânicos ativos para o tablet (vem em /api/dados-iniciais) */
  async function carregarMecanicos() {
    try {
      const itens = await itensDe('mecanicos');
      const out = [];
      for (const it of itens) {
        const f = await ler('mecanicos', it);
        const ativo = String(f.Ativo === undefined || f.Ativo === null ? 'Sim' : f.Ativo).trim().toLowerCase();
        if (ativo === 'não' || ativo === 'nao' || ativo === 'false') continue;
        const nome = String(f.Title || '').replace(/\s+/g, ' ').trim();
        if (nome) out.push({ nome: nome, matricula: String(f.Matricula || '').replace(/\D/g, '').replace(/^0+/, '') });
      }
      return out;
    } catch (e) { return []; }
  }

  return { registrar, carregarMecanicos, _partesSP: partesSP };
}

module.exports = { criar };
