/**
 * Cópia EXATA do MAPA_CAMPOS do portal_comprex.html — mapeia o "rótulo" de
 * cada campo/pergunta do checklist para o nome de exibição da coluna real
 * na lista "Checklists" do SharePoint. Se um dia esse mapa mudar no app
 * principal, replique a mudança aqui também (os dois precisam ficar iguais).
 */
const MAPA_CAMPOS = {
  'Equipamento': 'Equipamento',
  'Horimetro': 'Horímetro',
  'Operador': 'Operador',
  'Local': 'Local/ Pátio',
  'Observacao': 'Observações',
  'DATA DE VALIDADE DO EXTINTOR': 'Validade do extintor',
  'NÍVEL DE ÁGUA DO RADIADOR': 'Nível de água do radiador',
  'NÍVEL DE ÓLEO HIDRÁULICO': 'Nível de óleo hidráulico',
  'NÍVEL DE ÓLEO DE TRANSMISSÃO': 'Nível de óleo de transmissão',
  'ESTADO DAS CORREIAS': 'Estado das correias',
  'EXTINTOR DE INCÊNDIO': 'Estado do extintor',
  'FILTRO DE AR DO MOTOR (Interno e Externo)': 'Filtro de ar do motor',
  'LUBRIFICAÇÃO (GRAXA)': 'Lubrificação (Graxa)',
  'ESTADO DAS MANGUEIRAS': 'Estado das mangueiras',
  'EQUIPAMENTO POSSUI ALGUM VAZAMENTO?': 'Equipamento possui algum vazamento?',
  'CILINDROS HIDRÁULICOS': 'Cilindros hidráulicos',
  'ESTADO DAS LANTERNAS': 'Estado das lanternas',
  'PNEUS/CALIBRAGEM': 'Pneus/ Calibragem',
  'TENSÃO DAS ESTEIRAS': 'Tensão das esteiras',
  'FUNCIONAMENTO DO IMPLEMENTO': 'Funcionamento do implemento',
  'RETROVISORES': 'Retrovisores',
  'VIDROS': 'Vidros',
  'MARTELINHO DE SEGURANÇA': 'Martelinho de segurança',
  'CINTO DE SEGURANÇA': 'Cinto de segurança',
  'LIMPADOR DE PARA-BRISA': 'Limpador de parabrisa',
  'IGNIÇÃO/SISTEMA DE ARRANQUE': 'Ignição/ Sistema de arranque',
  'FREIOS': 'Freios',
  'COMANDOS DA MÁQUINA/RESPOSTA': 'Comandos da máquin/ resposta',
  'ESTRIBOS/ESCADAS/CORRIMÃO': 'Estribo/ Escadas/ Corrimão',
  'LATARIA/ESTRUTURA/PINTURA': 'Lataria/ Estrutura/ Pintura',
  'BATERIA': 'Bateria',
  'AR-CONDICIONADO': 'Ar condicionado',
  'CONSERVAÇÃO E LIMPEZA GERAL': 'Conservação e limpeza geral',
  'GIROFLEX': 'Giroflex',
  'ALARME SONORO PARA CINTO DE SEGURANÇA (FUNCIONA)': 'Alarme sonoro do cinto de segurança',
  'ALARME SONORO E/OU VISUAL PARA FREIO DE ESTACIONAMENTO (FUNCIONA)': 'Alarme sonoro e/ou visual para freio de estacionamento',
  'ALARME DE RÉ (FUNCIONA)': 'Alarme de ré',
  'SENSOR DE RÉ (FUNCIONA)': 'Sensor de ré',
  'CÂMERA DE RÉ (FUNCIONA)': 'Câmera de ré',
  'SISTEMA DE DETECÇÃO DE PROXIMIDADE (FUNCIONA)': 'Detecção de proxiidade',
  'NECESSITA PARALIZAR O EQUIPAMENTO?': 'Necessita paralisar?'
};

module.exports = { MAPA_CAMPOS };
