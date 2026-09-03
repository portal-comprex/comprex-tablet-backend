const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');

const JWT_SECRET = process.env.JWT_SECRET;
const ADMIN_API_KEY = process.env.ADMIN_API_KEY;

/** ---------------------------------------------------------------------
 *  Bloqueio de login por tentativas erradas + presença ("quem está online
 *  agora"). Tudo guardado em memória do próprio processo — não precisa de
 *  banco nem de coluna nova no SharePoint, mas também some se o serviço
 *  reiniciar (no plano free do Render isso acontece depois de um tempo
 *  parado). Pra bloqueio de senha isso é aceitável (o objetivo é atrapalhar
 *  um ataque automatizado, não guardar histórico); pra presença, "online
 *  agora" é por natureza uma informação passageira mesmo.
 *  ------------------------------------------------------------------- */
const LIMITE_TENTATIVAS = 5;
const BLOQUEIO_MINUTOS = 15;
const tentativasLogin = {}; // usuario -> { falhas, bloqueadoAte }

function statusBloqueioLogin(usuario) {
  const chave = String(usuario || '').trim().toLowerCase();
  const info = tentativasLogin[chave];
  if (info && info.bloqueadoAte && info.bloqueadoAte > Date.now()) {
    return { bloqueado: true, minutosRestantes: Math.ceil((info.bloqueadoAte - Date.now()) / 60000) };
  }
  return { bloqueado: false };
}
function registrarTentativaFalha(usuario) {
  const chave = String(usuario || '').trim().toLowerCase();
  const info = tentativasLogin[chave] || { falhas: 0, bloqueadoAte: null };
  info.falhas += 1;
  if (info.falhas >= LIMITE_TENTATIVAS) {
    info.bloqueadoAte = Date.now() + BLOQUEIO_MINUTOS * 60000;
    info.falhas = 0;
  }
  tentativasLogin[chave] = info;
}
function limparTentativas(usuario) {
  delete tentativasLogin[String(usuario || '').trim().toLowerCase()];
}

const LIMITE_ONLINE_MINUTOS = 3; // tablet manda um "ping" a cada ~60s enquanto logado
const ultimoAcesso = {}; // usuario -> timestamp (ms) do último acesso autenticado
function registrarAcesso(usuario) {
  if (usuario) ultimoAcesso[String(usuario).trim().toLowerCase()] = Date.now();
}
function infoPresenca(usuario) {
  const ts = ultimoAcesso[String(usuario || '').trim().toLowerCase()];
  if (!ts) return { online: false, ultimoAcesso: null };
  return { online: (Date.now() - ts) < LIMITE_ONLINE_MINUTOS * 60000, ultimoAcesso: ts };
}

function assinarToken(payload) {
  if (!JWT_SECRET) throw new Error('Configuração incompleta: defina JWT_SECRET nas configurações do Function App.');
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '14h' });
}

function verificarToken(token) {
  if (!JWT_SECRET) throw new Error('Configuração incompleta: defina JWT_SECRET nas configurações do Function App.');
  return jwt.verify(token, JWT_SECRET);
}

/** Lê o token "Bearer ..." do header Authorization e devolve o payload, ou lança erro.
 *  Compatível com req.headers do Express (objeto simples, chaves em minúsculo). */
function exigirOperadorLogado(request) {
  const h = request.headers || {};
  const cabecalho = h.authorization || h.Authorization || '';
  const partes = cabecalho.split(' ');
  if (partes.length !== 2 || partes[0].toLowerCase() !== 'bearer') {
    const erro = new Error('Token ausente. Faça login novamente.');
    erro.status = 401;
    throw erro;
  }
  try {
    const payload = verificarToken(partes[1]);
    registrarAcesso(payload.usuario); // qualquer chamada autenticada já conta como "online agora"
    return payload;
  } catch (e) {
    const erro = new Error('Sessão expirada ou inválida. Faça login novamente.');
    erro.status = 401;
    throw erro;
  }
}

/** Endpoints de administração (criar/desativar operador) usam uma chave fixa
 *  configurada no Function App — não um login — porque quem chama é o painel
 *  administrativo do Portal Comprex (portal_comprex.html), já protegido por
 *  login Microsoft + checagem de admin no próprio app. */
function exigirChaveAdmin(request) {
  const h = request.headers || {};
  const chave = h['x-admin-key'] || h['X-Admin-Key'] || '';
  if (!ADMIN_API_KEY || chave !== ADMIN_API_KEY) {
    const erro = new Error('Chave de administração ausente ou inválida.');
    erro.status = 401;
    throw erro;
  }
}

function hashSenha(senha) {
  return bcrypt.hashSync(String(senha), 10);
}

function conferirSenha(senha, hash) {
  if (!hash) return false;
  return bcrypt.compareSync(String(senha), String(hash));
}

module.exports = {
  assinarToken, verificarToken, exigirOperadorLogado, exigirChaveAdmin, hashSenha, conferirSenha,
  statusBloqueioLogin, registrarTentativaFalha, limparTentativas, infoPresenca
};
