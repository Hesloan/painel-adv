/**
 * Drive.gs — documentos de clientes e processos.
 *
 * Estrutura no Drive da Giovanna:
 *   Giovanna Adv · Clientes/            ← pasta-mãe
 *     Ana Paula Ribeiro/                ← pasta do cliente (documentos gerais: RG, CPF…)
 *       Divórcio consensual (12-10-2026)/   ← pasta de cada processo
 *
 * O painel NÃO guarda cópia de nada: ele lê as pastas na hora. O que for
 * adicionado, editado ou apagado direto no Drive aparece no painel, e o que
 * for enviado pelo painel aparece no Drive.
 *
 * `kind` é 'cliente' ou 'processo'; `id` é o id do cliente ou do processo.
 */

function driveRoot_() {
  const id = PropertiesService.getScriptProperties().getProperty('DRIVE_ROOT_ID');
  if (!id) throw new Error('Pasta de clientes não configurada. Rode setup().');
  return DriveApp.getFolderById(id);
}

function clientFolderName_(nome) {
  return String(nome || 'Cliente sem nome').trim();
}

function processFolderName_(proc) {
  const prod = proc.produto_id ? Repo.get('produtos', proc.produto_id) : null;
  const nome = prod ? prod.nome : 'Processo';
  const ref = proc.numero_processo || (proc.data_fechamento ? proc.data_fechamento.split('-').reverse().join('-') : '');
  return ref ? nome + ' (' + ref + ')' : nome;
}

function createClientFolder_(nome) {
  return driveRoot_().createFolder(clientFolderName_(nome));
}

function validFolder_(id) {
  if (!id) return null;
  try {
    const f = DriveApp.getFolderById(id);
    return f.isTrashed() ? null : f;
  } catch (e) { return null; }
}

/** Garante a pasta do cliente ou do processo; cria se não existir. */
function ensureFolder_(kind, id) {
  if (kind === 'cliente') {
    const cliente = Repo.get('clientes', id);
    if (!cliente) throw new Error('Cliente não encontrado.');
    const f = validFolder_(cliente.drive_folder_id);
    if (f) return { folder: f, record: null };
    const folder = createClientFolder_(cliente.nome);
    cliente.drive_folder_id = folder.getId();
    return { folder: folder, record: { table: 'clientes', data: Repo.upsert('clientes', cliente) } };
  }
  if (kind === 'processo') {
    const proc = Repo.get('contratos', id);
    if (!proc) throw new Error('Processo não encontrado.');
    const f = validFolder_(proc.drive_folder_id);
    if (f) return { folder: f, record: null };
    const parent = ensureFolder_('cliente', proc.cliente_id).folder;
    const folder = parent.createFolder(processFolderName_(proc));
    proc.drive_folder_id = folder.getId();
    return { folder: folder, record: { table: 'contratos', data: Repo.upsert('contratos', proc) } };
  }
  throw new Error('Tipo de pasta desconhecido: ' + kind);
}

function fileToItem_(f, isFolder) {
  return {
    id: f.getId(),
    name: f.getName(),
    url: f.getUrl(),
    mime: isFolder ? 'folder' : f.getMimeType(),
    size: isFolder ? null : f.getSize(),
    updated: f.getLastUpdated().toISOString(),
    isFolder: !!isFolder
  };
}

/** Lista o conteúdo da pasta (subpastas primeiro, depois arquivos mais recentes). */
function api_listFiles(kind, id) {
  return withLock_(() => {
    const r = ensureFolder_(kind, id);
    const items = [];
    const folders = r.folder.getFolders();
    while (folders.hasNext()) {
      const sub = folders.next();
      if (!sub.isTrashed()) items.push(fileToItem_(sub, true));
    }
    const files = r.folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (!f.isTrashed()) items.push(fileToItem_(f, false));
    }
    items.sort((a, b) => (a.isFolder === b.isFolder ? (a.updated < b.updated ? 1 : -1) : (a.isFolder ? -1 : 1)));
    return {
      folder: { id: r.folder.getId(), url: r.folder.getUrl(), name: r.folder.getName() },
      record: r.record,
      items: items
    };
  });
}

/** Recebe um arquivo (base64) e salva na pasta do cliente ou do processo. */
function api_upload(kind, id, file) {
  if (!file || !file.data || !file.name) throw new Error('Arquivo inválido.');
  const r = withLock_(() => ensureFolder_(kind, id));
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mime || 'application/octet-stream', file.name);
  const created = r.folder.createFile(blob);
  return fileToItem_(created, false);
}

/** Vincula uma pasta que já existe no Drive (cola o link) a um cliente ou processo. */
function api_linkFolder(kind, id, urlOrId) {
  const folderId = extractDriveId_(urlOrId);
  const folder = DriveApp.getFolderById(folderId); // lança erro se não existir/sem acesso
  const table = kind === 'processo' ? 'contratos' : 'clientes';
  return withLock_(() => {
    const rec = Repo.get(table, id);
    if (!rec) throw new Error('Registro não encontrado.');
    rec.drive_folder_id = folder.getId();
    return { table: table, data: Repo.upsert(table, rec) };
  });
}

/** Troca a pasta-mãe onde as pastas de clientes novos são criadas. */
function api_setDriveRoot(urlOrId) {
  const id = extractDriveId_(urlOrId);
  const folder = DriveApp.getFolderById(id);
  PropertiesService.getScriptProperties().setProperty('DRIVE_ROOT_ID', folder.getId());
  return 'https://drive.google.com/drive/folders/' + folder.getId();
}

function extractDriveId_(s) {
  const m = String(s || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('Não reconheci esse link do Drive.');
  return m[0];
}
