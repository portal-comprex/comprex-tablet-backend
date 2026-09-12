/**
 * Confere o token de acesso Microsoft (MSAL) que o Portal já tem — a MESMA
 * sessão que a pessoa usou pra entrar no Portal, sem pedir senha nenhuma de
 * novo. Isso é o que permite eliminar o login separado da Controladoria: em
 * vez de outro sistema de contas (Supabase Auth), a Controladoria passa a
 * confiar na conta Microsoft que já está logada.
 *
 * O QUE ESTE ARQUIVO FAZ: recebe o token que o navegador manda no cabeçalho
 * Authorization, confere a ASSINATURA dele contra as chaves públicas da
 * Microsoft (então ninguém pode forjar um token sem a senha da pessoa), e
 * confere que esse token foi emitido PARA O NOSSO Portal (mesmo Client ID)
 * e para o NOSSO tenant (mesma empresa) — assim um token de outro aplicativo
 * Microsoft qualquer não abre a porta aqui.
 *
 * Variáveis de ambiente necessárias (mesmo tenant já usado no App
 * Registration do Portal — ver portal_comprex.html, MSAL_CLIENT_ID lá):
 *   AAD_TENANT_ID        (já existe, reaproveitado)
 *   PORTAL_CLIENT_ID     (o Client ID do Portal — "aud" esperado no token)
 */
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const TENANT_ID = process.env.AAD_TENANT_ID;
const PORTAL_CLIENT_ID = process.env.PORTAL_CLIENT_ID;

const client = jwksClient({
  jwksUri: 'https://login.microsoftonline.com/' + TENANT_ID + '/discovery/v2.0/keys',
  cache: true,
  cacheMaxAge: 12 * 60 * 60 * 1000, // 12h — chaves da Microsoft mudam raramente
  rateLimit: true
});

function obterChave(header, callback) {
  client.getSigningKey(header.kid, function (err, key) {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

/** Devolve { email, nome } se o token for válido, ou lança erro (status 401). */
function validarTokenMicrosoft(tokenBruto) {
  return new Promise(function (resolve, reject) {
    if (!TENANT_ID || !PORTAL_CLIENT_ID) {
      return reject(Object.assign(new Error('Backend sem AAD_TENANT_ID/PORTAL_CLIENT_ID configurados.'), { status: 500 }));
    }
    jwt.verify(tokenBruto, obterChave, {
      audience: PORTAL_CLIENT_ID,
      issuer: 'https://login.microsoftonline.com/' + TENANT_ID + '/v2.0',
      algorithms: ['RS256']
    }, function (err, payload) {
      if (err) return reject(Object.assign(new Error('Sessão do Portal inválida ou expirada — saia e entre de novo.'), { status: 401 }));
      const email = String(payload.preferred_username || payload.upn || payload.email || '').toLowerCase().trim();
      if (!email) return reject(Object.assign(new Error('Token do Portal sem e-mail identificável.'), { status: 401 }));
      resolve({ email: email, nome: payload.name || email });
    });
  });
}

/** Lê "Authorization: Bearer <token>" do request e valida. Lança erro (status) se faltar/for inválido. */
async function exigirContaMicrosoft(req) {
  const auth = req.headers.authorization || '';
  const partes = auth.split(' ');
  if (partes.length !== 2 || partes[0] !== 'Bearer' || !partes[1]) {
    throw Object.assign(new Error('Faça login no Portal para acessar a Controladoria.'), { status: 401 });
  }
  return validarTokenMicrosoft(partes[1]);
}

module.exports = { exigirContaMicrosoft };
