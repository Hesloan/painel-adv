/**
 * Repo.gs — camada de acesso a dados.
 *
 * TODO o acesso ao armazenamento passa por aqui. Hoje o armazenamento é uma
 * planilha Google (uma aba por tabela). Na migração para um banco de verdade
 * (ex.: Cloudflare D1), só este arquivo é reescrito: as funções mantêm a mesma
 * assinatura e o restante do sistema (API e telas) não muda.
 *
 * Convenções (iguais às de um banco SQL, para a migração ser mecânica):
 *  - toda tabela tem `id` (texto) como chave primária;
 *  - datas são texto ISO `AAAA-MM-DD`; carimbos de tempo são ISO completo;
 *  - valores em reais são números (ponto decimal), nunca texto formatado;
 *  - relacionamentos usam `<tabela>_id`.
 */

const SCHEMA = {
  clientes:  ['id', 'nome', 'whatsapp', 'email', 'cidade', 'origem', 'campanha_id', 'drive_folder_id', 'status', 'observacoes', 'criado_em', 'atualizado_em'],
  produtos:  ['id', 'nome', 'area', 'valor_referencia', 'ativo', 'criado_em', 'atualizado_em'],
  // contratos = processos (cada processo tem seu contrato de honorários)
  contratos: ['id', 'cliente_id', 'produto_id', 'data_fechamento', 'valor_total', 'forma_pagamento', 'status', 'campanha_id', 'observacoes', 'criado_em', 'atualizado_em', 'numero_processo', 'orgao', 'andamento', 'drive_folder_id'],
  parcelas:  ['id', 'contrato_id', 'numero', 'valor', 'vencimento', 'tipo', 'condicao', 'status', 'data_pagamento', 'criado_em', 'atualizado_em'],
  gastos:    ['id', 'data', 'categoria', 'descricao', 'valor', 'campanha_id', 'fornecedor', 'criado_em', 'atualizado_em', 'produto_id'],
  campanhas: ['id', 'nome', 'plataforma', 'produto_id', 'inicio', 'fim', 'status', 'orcamento', 'observacoes', 'criado_em', 'atualizado_em'],
  config:    ['chave', 'valor']
};

/** Nome da aba de cada tabela na planilha. */
const SHEET_NAMES = {
  clientes: 'Clientes', produtos: 'Produtos', contratos: 'Contratos', parcelas: 'Parcelas',
  gastos: 'Gastos', campanhas: 'Campanhas', config: 'Config'
};

/** Colunas numéricas por tabela. */
const NUMERIC = {
  produtos: ['valor_referencia'],
  contratos: ['valor_total'],
  parcelas: ['numero', 'valor'],
  gastos: ['valor'],
  campanhas: ['orcamento']
};

/** Colunas booleanas por tabela. */
const BOOLEAN = { produtos: ['ativo'] };

const DATA_TABLES = ['clientes', 'produtos', 'contratos', 'parcelas', 'gastos', 'campanhas'];

const Repo = {
  _cache: {},

  ss() {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('O painel ainda não foi configurado. Rode a função setup() no editor do Apps Script.');
    if (!this._cache.ss) this._cache.ss = SpreadsheetApp.openById(id);
    return this._cache.ss;
  },

  sheet(table) {
    assertTable_(table);
    const sh = this.ss().getSheetByName(SHEET_NAMES[table]);
    if (!sh) throw new Error('Aba "' + SHEET_NAMES[table] + '" não encontrada. Rode setup() de novo.');
    return sh;
  },

  /** Lê a tabela inteira como lista de objetos. */
  all(table) {
    const sh = this.sheet(table);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const cols = SCHEMA[table];
    const values = sh.getRange(2, 1, last - 1, cols.length).getValues();
    const out = [];
    for (const row of values) {
      if (row[0] === '' || row[0] === null) continue;
      out.push(rowToObj_(table, row));
    }
    return out;
  },

  /** Lê todas as tabelas de dados de uma vez (uma ida ao servidor). */
  allTables() {
    const out = {};
    DATA_TABLES.forEach(t => { out[t] = this.all(t); });
    return out;
  },

  get(table, id) {
    const pos = this._findRow(table, id);
    if (!pos) return null;
    const cols = SCHEMA[table];
    const row = this.sheet(table).getRange(pos, 1, 1, cols.length).getValues()[0];
    return rowToObj_(table, row);
  },

  /** Insere ou atualiza pelo `id`. Retorna o registro salvo. */
  upsert(table, obj) {
    const key = SCHEMA[table][0];
    const now = new Date().toISOString();
    const rec = Object.assign({}, obj);
    if (!rec[key]) rec[key] = Utilities.getUuid();
    if (table !== 'config') {
      if (!rec.criado_em) rec.criado_em = now;
      rec.atualizado_em = now;
    }
    const sh = this.sheet(table);
    const cols = SCHEMA[table];
    const pos = this._findRow(table, rec[key]);
    const rowNum = pos || Math.max(sh.getLastRow(), 1) + 1;
    if (pos && table !== 'config') {
      // Preserva a data de criação original.
      const created = sh.getRange(pos, cols.indexOf('criado_em') + 1).getValue();
      if (created) rec.criado_em = String(created);
    }
    const range = sh.getRange(rowNum, 1, 1, cols.length);
    range.setNumberFormats([cols.map(c => columnFormat_(table, c))]);
    range.setValues([objToRow_(table, rec)]);
    return rowToObj_(table, objToRow_(table, rec));
  },

  upsertMany(table, list) {
    return list.map(o => this.upsert(table, o));
  },

  remove(table, id) {
    const pos = this._findRow(table, id);
    if (pos) this.sheet(table).deleteRow(pos);
    return !!pos;
  },

  /** Remove todas as linhas onde `field === value`. Retorna os ids removidos. */
  removeWhere(table, field, value) {
    const sh = this.sheet(table);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const cols = SCHEMA[table];
    const fi = cols.indexOf(field);
    const values = sh.getRange(2, 1, last - 1, cols.length).getValues();
    const removed = [];
    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][fi]) === String(value)) {
        removed.push(String(values[i][0]));
        sh.deleteRow(i + 2);
      }
    }
    return removed;
  },

  _findRow(table, id) {
    const sh = this.sheet(table);
    const last = sh.getLastRow();
    if (last < 2 || !id) return null;
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) return i + 2;
    }
    return null;
  },

  // --- Config (chave/valor; listas guardadas como JSON) ---
  getConfig() {
    const out = {};
    this.all('config').forEach(r => {
      try { out[r.chave] = JSON.parse(r.valor); } catch (e) { out[r.chave] = r.valor; }
    });
    return out;
  },

  setConfig(chave, valor) {
    return this.upsert('config', { chave: chave, valor: JSON.stringify(valor) });
  }
};

// --- helpers ---------------------------------------------------------------

function assertTable_(table) {
  if (!SCHEMA[table]) throw new Error('Tabela desconhecida: ' + table);
}

function columnFormat_(table, col) {
  return (NUMERIC[table] || []).indexOf(col) >= 0 ? '0.00' : '@';
}

function rowToObj_(table, row) {
  const cols = SCHEMA[table];
  const nums = NUMERIC[table] || [];
  const bools = BOOLEAN[table] || [];
  const obj = {};
  cols.forEach((c, i) => {
    let v = row[i];
    if (v instanceof Date) {
      // Se alguém editar a planilha à mão e o Sheets converter para data.
      v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (nums.indexOf(c) >= 0) {
      v = (v === '' || v === null) ? null : Number(String(v).replace(',', '.'));
      if (v !== null && isNaN(v)) v = null;
    } else if (bools.indexOf(c) >= 0) {
      v = !(v === false || v === 'false' || v === 'FALSE' || v === 0 || v === '0');
    } else {
      v = (v === null || v === undefined) ? '' : String(v);
    }
    obj[c] = v;
  });
  return obj;
}

function objToRow_(table, obj) {
  const nums = NUMERIC[table] || [];
  const bools = BOOLEAN[table] || [];
  return SCHEMA[table].map(c => {
    const v = obj[c];
    if (nums.indexOf(c) >= 0) {
      if (v === '' || v === null || v === undefined) return '';
      const n = Number(v);
      return isNaN(n) ? '' : n;
    }
    if (bools.indexOf(c) >= 0) return v === false || v === 'false' ? 'false' : 'true';
    return v === null || v === undefined ? '' : String(v);
  });
}
