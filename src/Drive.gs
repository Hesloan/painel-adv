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

function listItems_(folder) {
  const items = [];
  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const sub = folders.next();
    if (!sub.isTrashed()) items.push(fileToItem_(sub, true));
  }
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (!f.isTrashed()) items.push(fileToItem_(f, false));
  }
  items.sort((a, b) => (a.isFolder === b.isFolder ? (a.updated < b.updated ? 1 : -1) : (a.isFolder ? -1 : 1)));
  return items;
}

/** Só permite abrir/enviar em pastas que estão dentro da pasta-mãe dos clientes. */
function folderInsideRoot_(folderId) {
  const rootId = driveRoot_().getId();
  const folder = DriveApp.getFolderById(folderId);
  let level = [folder];
  for (let depth = 0; depth < 8 && level.length; depth++) {
    const next = [];
    for (const f of level) {
      if (f.getId() === rootId) return folder;
      const ps = f.getParents();
      while (ps.hasNext()) next.push(ps.next());
    }
    level = next;
  }
  // Pastas vinculadas manualmente a um cliente/processo também valem, mesmo fora da pasta-mãe.
  const linked = Repo.all('clientes').some(c => c.drive_folder_id === folderId) || Repo.all('contratos').some(c => c.drive_folder_id === folderId);
  if (linked) return folder;
  throw new Error('Esta pasta está fora da pasta de clientes do painel.');
}

/** Lista o conteúdo da pasta do cliente ou do processo (subpastas primeiro, depois arquivos recentes). */
function api_listFiles(kind, id) {
  return withLock_(() => {
    const r = ensureFolder_(kind, id);
    return {
      folder: { id: r.folder.getId(), url: r.folder.getUrl(), name: r.folder.getName() },
      record: r.record,
      items: listItems_(r.folder)
    };
  });
}

/** Lista uma subpasta qualquer (navegação dentro das pastas no painel). */
function api_listFolder(folderId) {
  const folder = folderInsideRoot_(folderId);
  return { folder: { id: folder.getId(), url: folder.getUrl(), name: folder.getName() }, items: listItems_(folder) };
}

/** Recebe um arquivo (base64) e salva na pasta do cliente/processo, ou numa subpasta dela. */
function api_upload(kind, id, file, folderId) {
  if (!file || !file.data || !file.name) throw new Error('Arquivo inválido.');
  const folder = folderId ? folderInsideRoot_(folderId) : withLock_(() => ensureFolder_(kind, id)).folder;
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mime || 'application/octet-stream', file.name);
  const created = folder.createFile(blob);
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
