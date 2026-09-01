const { graphGet, SITE_ID } = require('./graph');

/** Fotos dos equipamentos ficam numa biblioteca de documentos do SharePoint
 *  chamada "FotosEquipamentos", um arquivo .jpg por equipamento (nome do
 *  arquivo = código do equipamento, ex.: "TR-04.jpg") — mesma biblioteca já
 *  usada pelo portal_comprex.html (buscarUrlFotoEquipamento). Aqui o backend
 *  busca via Graph com o token de aplicativo e devolve só a URL de download,
 *  já que o tablet não tem login Microsoft pra fazer essa chamada sozinho. */
let driveFotosEquipId = null;
const cacheFotoEquip = {}; // código (maiúsculo) -> url (ou null se não encontrada)

async function obterDriveFotosEquipamento() {
  if (driveFotosEquipId) return driveFotosEquipId;
  const d = await graphGet('/sites/' + SITE_ID + '/drives');
  const lib = (d.value || []).find(function (x) { return x.name === 'FotosEquipamentos'; });
  if (!lib) throw new Error('Biblioteca "FotosEquipamentos" não encontrada no SharePoint.');
  driveFotosEquipId = lib.id;
  return driveFotosEquipId;
}

async function buscarUrlFotoEquipamento(codigoEquip) {
  const chave = String(codigoEquip || '').trim().toUpperCase();
  if (!chave) return null;
  if (cacheFotoEquip[chave] !== undefined) return cacheFotoEquip[chave];
  try {
    const driveId = await obterDriveFotosEquipamento();
    const nomeArquivo = encodeURIComponent(chave) + '.jpg';
    const d = await graphGet('/drives/' + driveId + '/root:/' + nomeArquivo);
    const url = d['@microsoft.graph.downloadUrl'] || null;
    cacheFotoEquip[chave] = url;
    return url;
  } catch (err) {
    cacheFotoEquip[chave] = null; // sem foto cadastrada (ou biblioteca ausente) — não é erro, só não mostra foto
    return null;
  }
}

module.exports = { buscarUrlFotoEquipamento };
