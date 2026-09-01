const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');

const JWT_SECRET = process.env.JWT_SECRET;
const ADMIN_API_KEY = process.env.ADMIN_API_KEY;

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
    return verificarToken(partes[1]);
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

module.exports = { assinarToken, verificarToken, exigirOperadorLogado, exigirChaveAdmin, hashSenha, conferirSenha };
