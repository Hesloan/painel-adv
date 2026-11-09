/**
 * Painel Giovanna Adv — backend (Google Apps Script)
 *
 * Arquivos do projeto:
 *   Codigo.gs   → publicação do app, API chamada pelas telas, setup inicial
 *   Repo.gs     → camada de dados (planilha hoje; banco de verdade amanhã)
 *   Drive.gs    → pastas e documentos de cada cliente no Google Drive
 *   Index.html, Estilos.html, App.html → as telas
 *
 * Instalação: veja o guia INSTALACAO.md.
 */

const APP_VERSION = '1.3.2';

const DEFAULT_CONFIG = {
  categorias_gasto: ['Anúncios', 'Mentoria', 'Curso', 'Software / assinatura', 'Projeto / freelancer', 'Material / escritório', 'Impostos e taxas', 'Outro'],
  origens: ['Meta Ads (Instagram/Facebook)', 'Google Ads', 'Instagram orgânico', 'Site / Google orgânico', 'Indicação', 'Cliente antigo', 'Outro'],
  areas: ['Família', 'Patrimonial', 'Previdenciário', 'Outro'],
  plataformas: ['Meta Ads', 'Google Ads', 'Orgânico', 'Indicação', 'Outro'],
  tipos_prazo: ['Prazo processual', 'Audiência', 'Reunião', 'Tarefa', 'Lembrete'],
  categorias_receita: ['Consulta', 'Parecer', 'Honorários avulsos', 'Outro']
};

const DEFAULT_PRODUTOS = [
  ['Divórcio consensual', 'Família'],
  ['Divórcio litigioso', 'Família'],
  ['Guarda de filhos', 'Família'],
  ['Pensão alimentícia', 'Família'],
  ['União estável', 'Família'],
  ['Planejamento matrimonial (pacto antenupcial)', 'Patrimonial'],
  ['Proteção patrimonial', 'Patrimonial'],
  ['Salário-maternidade', 'Previdenciário']
];

/** Origens/campanhas iniciais: [nome, plataforma, serviço, início]. */
const DEFAULT_CAMPANHAS = [
  ['Orgânico', 'Orgânico', '', ''],
  ['Indicação', 'Indicação', '', ''],
  ['Google Ads - Divórcio Consensual', 'Google Ads', 'Divórcio consensual', ''],
  ['Meta Ads - 2 Salário Maternidade - 26/10', 'Meta Ads', 'Salário-maternidade', '2026-10-26'],
  ['Meta Ads - 2 Salário Maternidade - 03/11', 'Meta Ads', 'Salário-maternidade', '2026-11-03']
];

/** Tabelas que as telas podem gravar diretamente. */
const WRITABLE = ['clientes', 'produtos', 'contratos', 'parcelas', 'gastos', 'campanhas', 'prazos'];

// ===========================================================================
// Publicação
// ===========================================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Painel · Giovanna Adv')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Usado pelo Index.html para incluir os outros arquivos HTML. */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ===========================================================================
// Setup (rodar UMA vez no editor, logado na conta da Giovanna)
// ===========================================================================

function setup() {
  const props = PropertiesService.getScriptProperties();

  // 1. Pasta do sistema e planilha-banco
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('Painel Giovanna Adv · Banco de dados');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
    const sysFolder = DriveApp.createFolder('Painel Giovanna Adv · Sistema');
    DriveApp.getFileById(ssId).moveTo(sysFolder);
  }

  // 2. Abas com cabeçalho, formato e proteção
  Object.keys(SCHEMA).forEach(table => {
    const name = SHEET_NAMES[table];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const cols = SCHEMA[table];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#F9F2E7');
    sh.setFrozenRows(1);
    cols.forEach((c, i) => {
      sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat(columnFormat_(table, c));
    });
    if (sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length === 0) {
      sh.protect().setDescription('Gerenciado pelo Painel — edite pelo app').setWarningOnly(true);
    }
  });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1') || ss.getSheetByName('Planilha1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // 3. Listas padrão (só cria o que ainda não existe)
  const cfg = Repo.getConfig();
  Object.keys(DEFAULT_CONFIG).forEach(k => {
    if (!cfg[k]) Repo.setConfig(k, DEFAULT_CONFIG[k]);
  });

  // 4. Catálogo inicial de produtos (só se estiver vazio)
  if (Repo.all('produtos').length === 0) {
    DEFAULT_PRODUTOS.forEach(p => Repo.upsert('produtos', { nome: p[0], area: p[1], valor_referencia: null, ativo: true }));
  }

  // 5. Origens e campanhas iniciais (só se estiver vazio)
  if (Repo.all('campanhas').length === 0) {
    const prods = Repo.all('produtos');
    DEFAULT_CAMPANHAS.forEach(c => {
      const prod = prods.find(p => p.nome === c[2]);
      Repo.upsert('campanhas', { nome: c[0], plataforma: c[1], produto_id: prod ? prod.id : '', inicio: c[3], fim: '', status: 'ativa', orcamento: null, observacoes: '' });
    });
  }

  // 6. Pasta-mãe dos clientes no Drive
  if (!props.getProperty('DRIVE_ROOT_ID')) {
    const root = DriveApp.createFolder('Giovanna Adv · Clientes');
    props.setProperty('DRIVE_ROOT_ID', root.getId());
  }

  props.setProperty('SCHEMA_VERSION', APP_VERSION);
  Logger.log('Pronto. Planilha: ' + ss.getUrl());
  Logger.log('Pasta de clientes: ' + DriveApp.getFolderById(props.getProperty('DRIVE_ROOT_ID')).getUrl());
}

// ===========================================================================
// API — chamada pelas telas via google.script.run
// Cada função retorna apenas dados simples (texto, número, listas, objetos).
// ===========================================================================

/** Carrega tudo de uma vez na abertura do painel. */
function api_bootstrap() {
  const props = PropertiesService.getScriptProperties();
  // Migração automática: na primeira abertura de uma versão nova, ajusta a planilha
  // (colunas e listas novas). Não apaga nem altera dados existentes.
  if (props.getProperty('SPREADSHEET_ID') && props.getProperty('SCHEMA_VERSION') !== APP_VERSION) {
    withLock_(() => { if (props.getProperty('SCHEMA_VERSION') !== APP_VERSION) setup(); });
  }
  const rootId = props.getProperty('DRIVE_ROOT_ID');
  return {
    tables: Repo.allTables(),
    config: Object.assign({}, DEFAULT_CONFIG, Repo.getConfig()),
    meta: {
      version: APP_VERSION,
      hoje: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      planilhaUrl: Repo.ss().getUrl(),
      driveRootUrl: rootId ? 'https://drive.google.com/drive/folders/' + rootId : ''
    }
  };
}

/** Salva (cria ou atualiza) um registro. */
function api_save(table, record) {
  if (WRITABLE.indexOf(table) < 0) throw new Error('Tabela não permitida: ' + table);
  return withLock_(() => {
    if (table === 'clientes') return saveCliente_(record);
    return Repo.upsert(table, record);
  });
}

/** Salva um processo (contrato) com todas as parcelas dele numa única operação.
 *  Na primeira vez, cria a pasta do processo dentro da pasta do cliente no Drive. */
function api_saveContrato(contrato, parcelas) {
  return withLock_(() => {
    let saved = Repo.upsert('contratos', contrato);
    if (!validFolder_(saved.drive_folder_id)) {
      try { saved = ensureFolder_('processo', saved.id).record.data; } catch (e) { /* segue; a pasta é criada ao abrir os documentos */ }
    }
    const keep = {};
    (parcelas || []).forEach(p => { if (p.id) keep[p.id] = true; });
    // Remove parcelas que saíram do contrato
    Repo.all('parcelas')
      .filter(p => p.contrato_id === saved.id && !keep[p.id])
      .forEach(p => Repo.remove('parcelas', p.id));
    const savedParcelas = (parcelas || []).map(p => Repo.upsert('parcelas', Object.assign({}, p, { contrato_id: saved.id })));
    return { contrato: saved, parcelas: savedParcelas };
  });
}

/** Exclui um registro (com cascata: cliente → contratos → parcelas). */
function api_delete(table, id) {
  if (WRITABLE.indexOf(table) < 0) throw new Error('Tabela não permitida: ' + table);
  return withLock_(() => {
    const removed = { contratos: [], parcelas: [], prazos: [] };
    if (table === 'clientes') {
      Repo.all('contratos').filter(c => c.cliente_id === id).forEach(c => {
        removed.parcelas = removed.parcelas.concat(Repo.removeWhere('parcelas', 'contrato_id', c.id));
        Repo.remove('contratos', c.id);
        removed.contratos.push(c.id);
      });
      removed.prazos = Repo.removeWhere('prazos', 'cliente_id', id);
      // Recebimentos avulsos do cliente continuam existindo, só perdem o vínculo (guardam o nome).
      const cli = Repo.get('clientes', id) || {};
      Repo.all('parcelas').filter(p => !p.contrato_id && p.cliente_id === id)
        .forEach(p => { p.pessoa = p.pessoa || cli.nome || ''; p.cliente_id = ''; Repo.upsert('parcelas', p); });
    }
    if (table === 'contratos') {
      removed.parcelas = Repo.removeWhere('parcelas', 'contrato_id', id);
      removed.prazos = Repo.removeWhere('prazos', 'contrato_id', id);
    }
    Repo.remove(table, id);
    // Obs.: a pasta do cliente no Drive NUNCA é apagada pelo painel.
    return removed;
  });
}

/** Atualiza uma lista de configuração (categorias, origens etc.). */
function api_saveConfig(chave, valor) {
  if (!DEFAULT_CONFIG[chave]) throw new Error('Configuração desconhecida: ' + chave);
  if (!Array.isArray(valor)) throw new Error('Valor inválido.');
  return withLock_(() => { Repo.setConfig(chave, valor); return valor; });
}

// ===========================================================================
// Internos
// ===========================================================================

function saveCliente_(record) {
  const before = record.id ? Repo.get('clientes', record.id) : null;
  const rec = Object.assign({}, record);
  // Cria a pasta do cliente no Drive no primeiro salvamento
  if (!rec.drive_folder_id) {
    try { rec.drive_folder_id = createClientFolder_(rec.nome).getId(); } catch (e) { /* segue sem pasta */ }
  } else if (before && before.nome !== rec.nome && before.drive_folder_id === rec.drive_folder_id) {
    // Nome mudou: renomeia a pasta junto
    try { DriveApp.getFolderById(rec.drive_folder_id).setName(clientFolderName_(rec.nome)); } catch (e) { /* ignora */ }
  }
  return Repo.upsert('clientes', rec);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
