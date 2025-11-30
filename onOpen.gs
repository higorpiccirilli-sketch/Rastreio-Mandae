/**
 * =============================================================================
 * File: onOpen.gs
 * Módulo: Abertura, Sidebar e Processos (Importar/Exportar)
 * Versão: 2.2.2 (visual do painel preservado; Log RT inalterado)
 * Autores: Higor Piccirilli Soares Batista & GPT-5 Thinking
 * Data: 2025-10-25
 * Descrição:
 *   - Mantém todo o comportamento original (menu, sidebar, importação/exportação).
 *   - NÃO altera o Log em Tempo Real (UI/fluxo do Log permanece o mesmo).
 *   - Atualizações desta versão:
 *       1) Exportação só bloqueia quando (Nome + Chave) JÁ possuem link na coluna D do Log.
 *       2) Estrutura de pastas por ANO/MÊS (ex.: 2025/10-Outubro). Cria se não existir.
 *       3) Nome do arquivo: N- Rastreio_mandae - AAAA-MM-DD.xlsx (N = contador diário).
 *       4) Log coluna D recebe link de download direto (uc?export=download&id=...).
 *       5) **NOVO:** Log coluna E recebe o NOME DO ARQUIVO gerado (exibido no painel).
 *       6) “Últimos Arquivos” usa o nome da coluna E (se existir); caso contrário, nome sintético.
 * =============================================================================
 * Observação:
 *   - Se os arquivos do Log RT não estiverem presentes, este arquivo funciona
 *     normalmente (Sidebar/menu/processos). Apenas os botões que chamam Log
 *     não devem ser acionados nesse caso.
 * =============================================================================
 */

// --- CONFIGURAÇÃO PRINCIPAL ---
const CONFIG = {
  mainSpreadsheetId: "1f1SLMQt4dVw15Mk1s72CiA9hDVG2sLs76fBxOOURCOo",
  templateSpreadsheetId: "1TA7H8oXTuVjjQQomHbJT1yQJpklktRzh8GktKPn9Jfk",
  exportFolderId: "1b5SuEcivuWqC7-9P0zBkh8dHzcFvFSxY",
  importSheetName: "Sheet1",
  logSheetName: "Log",
  headerRow: 2,
  timeZone: "America/Sao_Paulo"
};

/** Menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NF-e Tools')
    .addItem('Abrir Painel de Controle', 'showSidebar')
    .addToUi();
}

/** Sidebar (HTML direto; Log RT e visual separados) */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Painel de Controle NF-e');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ===================================================================================================
// FUNÇÕES CHAMADAS PELA SIDEBAR (originais, com ajustes onde indicado)
// ===================================================================================================

function listXmlFiles() {
  const xmlFiles = [];
  const fileIds = new Set();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  function addFilesFrom(mimeType) {
    const filesIterator = DriveApp.getFilesByType(mimeType);
    while (filesIterator.hasNext()) {
      const file = filesIterator.next();
      if (file.getLastUpdated() >= today && !fileIds.has(file.getId())) {
        fileIds.add(file.getId());
        xmlFiles.push({ name: file.getName(), id: file.getId() });
      }
    }
  }

  addFilesFrom('application/xml');
  addFilesFrom('text/xml');
  return xmlFiles;
}

/** Importa múltiplos XMLs (mantida) */
function processMultipleXmlFiles(fileIds) {
  limparDadosSheet1Interna();

  const summary = { successCount: 0, duplicateCount: 0, errorCount: 0 };
  const loggedKeys = getLoggedKeys();
  for (const id of fileIds) {
    try {
      const result = processSingleXmlFile(id, loggedKeys);
      if (result.status === 'SUCCESS') {
        summary.successCount++;
        loggedKeys.add(result.chaveNf);
      } else if (result.status === 'DUPLICATE') {
        summary.duplicateCount++;
      } else {
        summary.errorCount++;
      }
    } catch (e) {
      Logger.log(`Erro não capturado ao processar ID ${id}: ${e.message}`);
      summary.errorCount++;
    }
  }
  return summary;
}

/**
 * Exporta os dados para .xlsx usando a planilha template.
 * - Bloqueio só quando (Nome + Chave) já possuem URL na coluna D (ou seja, já geraram .xlsx).
 * - Pastas ANO/MÊS e nome N- Rastreio_mandae - AAAA-MM-DD.xlsx.
 * - Atualiza Log: D=download direto, E=nome do arquivo gerado.
 */
function generateExportFile() {
  try {
    // --- Etapa 1: Ler dados da planilha principal ---
    const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(CONFIG.importSheetName);
    const lastDataRow = getActualLastRow(sourceSheet);
    if (lastDataRow < CONFIG.headerRow + 1) {
      return 'Nenhum dado para exportar na "Sheet1".';
    }

    const sourceRange = sourceSheet.getRange(
      CONFIG.headerRow, 1,
      lastDataRow - CONFIG.headerRow + 1,
      sourceSheet.getLastColumn()
    );
    const dataToCopy = sourceRange.getDisplayValues(); // [headers, row1, ...]
    const headers = dataToCopy[0].map(h => String(h || '').trim());
    const headerIdx = indexMap(headers);

    const COL_NOME  = headerIdx['NOME DO DESTINATÁRIO*'];
    const COL_CHAVE = headerIdx['CHAVE NF'];
    if (COL_NOME === undefined || COL_CHAVE === undefined) {
      throw new Error('Erro: Colunas "NOME DO DESTINATÁRIO*" e/ou "CHAVE NF" não encontradas na Sheet1.');
    }

    // --- Etapa 1.1: Bloqueio apenas se já existe D (download) no Log para (Nome+Chave) ---
    const logSheet = sourceSpreadsheet.getSheetByName(CONFIG.logSheetName);
    const exportedPairsSet = readExportedPairsSet_(logSheet); // nome||chave com D preenchido
    const duplicates = [];
    const seenInBatch = new Set();

    for (let i = 1; i < dataToCopy.length; i++) {
      const row = dataToCopy[i];
      const nome  = String(row[COL_NOME]  || '').trim();
      const chave = String(row[COL_CHAVE] || '').trim();
      if (!nome || !chave) continue;
      const k = nome + '||' + chave;

      if (exportedPairsSet.has(k)) {
        if (duplicates.length < 20) duplicates.push(`• ${nome} | ${chave}`);
      }
      if (seenInBatch.has(k)) {
        if (duplicates.length < 20) duplicates.push(`• (Duplicado no lote) ${nome} | ${chave}`);
      }
      seenInBatch.add(k);
    }

    if (duplicates.length > 0) {
      throw new Error(
        '[ERRO] Exportação bloqueada: já existe arquivo gerado (coluna D do Log preenchida) para os pares abaixo.\n' +
        'Exemplos:\n' + duplicates.join('\n')
      );
    }

    // --- Etapa 1.2: Preparar template e colar dados (valores apenas) ---
    const templateSpreadsheet = SpreadsheetApp.openById(CONFIG.templateSpreadsheetId);
    const templateSheet = templateSpreadsheet.getSheets()[0];

    templateSheet.getRange(CONFIG.headerRow, 1,
      templateSheet.getMaxRows() - CONFIG.headerRow + 1,
      templateSheet.getMaxColumns()
    ).clearContent();

    templateSheet.getRange(CONFIG.headerRow, 1, dataToCopy.length, dataToCopy[0].length).setValues(dataToCopy);
    SpreadsheetApp.flush();

    // --- Etapa 2: Pastas ANO/MÊS e criação do arquivo ---
    const now = new Date();
    const tz = CONFIG.timeZone || 'America/Sao_Paulo';
    const yyyy       = Utilities.formatDate(now, tz, 'yyyy');
    const yyyy_mm_dd = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    const rootFolder  = DriveApp.getFolderById(CONFIG.exportFolderId);
    const yearFolder  = getOrCreateSubfolder_(rootFolder, yyyy);
    const monthFolder = getOrCreateSubfolder_(yearFolder, getMonthFolderName_(now, tz)); // "10-Outubro"

    const countToday = countFilesCreatedOn_(monthFolder, now, tz);
    const N = countToday + 1;

    const exportUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.templateSpreadsheetId}/export?format=xlsx&gid=${templateSheet.getSheetId()}`;
    const blob      = UrlFetchApp.fetch(exportUrl, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();

    const fileName  = `${N}- Rastreio_mandae - ${yyyy_mm_dd}.xlsx`;
    const file      = monthFolder.createFile(blob.setName(fileName));
    const downloadUrl = buildDownloadUrlFromId_(file.getId()); // uc?export=download&id=<ID>

    // --- Etapa 3: Atualizar Log (D=downloadUrl, E=fileName) para todas as chaves exportadas ---
    const chaveNfColIndex = headers.indexOf('CHAVE NF');
    const exportedKeys = new Set(dataToCopy.slice(1).map(r => r[chaveNfColIndex]));
    updateLogWithLink(exportedKeys, downloadUrl, fileName);

    return `Arquivo "${fileName}" gerado com sucesso! O Log foi atualizado.`;

  } catch (e) {
    Logger.log(`ERRO ao gerar exportação: ${e.stack}`);
    throw e;
  }
}

/**
 * “Últimos Arquivos” pelo Log (A:D[+E]), sem consultar Drive:
 *  - 5 últimos DownloadUrl distintos (coluna D), com timestamp máximo (C).
 *  - Se existir coluna E (nome do arquivo), usa E como "name"; senão, usa nome sintético.
 *  - url (view) é derivada do DownloadUrl. downloadUrl é retornado (será re-normalizado no front).
 */
function getRecentExportLinks() {
  try {
    const ss       = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId);
    const logSheet = ss.getSheetByName(CONFIG.logSheetName);
    if (!logSheet) return [];
    const last = logSheet.getLastRow();
    if (last < 2) return [];

    const lastCol = logSheet.getLastColumn();
    const colsToRead = Math.min(5, Math.max(4, lastCol)); // lê A:D ou A:E se existir
    const values = logSheet.getRange(2, 1, last - 1, colsToRead).getValues(); // [A..D,(E?)]
    const hasNameCol = colsToRead >= 5;

    const tz = CONFIG.timeZone || 'America/Sao_Paulo';

    // Mapa D->max timestamp
    const maxTsByD = new Map();
    for (let i = 0; i < values.length; i++) {
      const tsStr = String(values[i][2] || '').trim(); // C
      const dUrl  = String(values[i][3] || '').trim(); // D
      if (!dUrl || !tsStr) continue;
      const dt = parseLogTimestamp_(tsStr, tz);
      const prev = maxTsByD.get(dUrl);
      if (!prev || dt > prev) maxTsByD.set(dUrl, dt);
    }

    // Varre de baixo pra cima pegando D distintos
    const picked = new Set();
    const out = [];
    for (let i = values.length - 1; i >= 0; i--) {
      const dUrl = String(values[i][3] || '').trim();
      if (!dUrl) continue;
      if (picked.has(dUrl)) continue;
      picked.add(dUrl);

      const dt = maxTsByD.get(dUrl) || new Date();
      const tsLabel = Utilities.formatDate(dt, tz, 'dd/MM HH:mm');
      const ymd     = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
      const openUrl = deriveOpenUrlFromDownloadUrl_(dUrl);

      // Nome do arquivo: usa E, se existir; senão, sintético
      const nameFromLog = hasNameCol ? String(values[i][4] || '').trim() : '';
      const name = nameFromLog || `Rastreio_mandae - ${ymd}.xlsx`;

      out.push({ name, url: openUrl, downloadUrl: dUrl, timestamp: tsLabel });
      if (out.length >= 5) break;
    }
    return out;

  } catch (e) {
    Logger.log(`Erro ao montar Últimos Arquivos pelo Log: ${e.message}`);
    return [];
  }
}

// ===================================================================================================
// FUNÇÕES INTERNAS E AUXILIARES
// ===================================================================================================

function processSingleXmlFile(fileId, loggedKeys) {
  const mainSpreadsheet = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId);
  const sheet    = mainSpreadsheet.getSheetByName(CONFIG.importSheetName);
  const logSheet = mainSpreadsheet.getSheetByName(CONFIG.logSheetName);

  try {
    const file = DriveApp.getFileById(fileId);
    const xmlContent = file.getBlob().getDataAsString();
    const doc   = XmlService.parse(xmlContent);
    const root  = doc.getRootElement();
    const nfeNs = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');
    const protNFe = root.getChild('protNFe', nfeNs);
    const chaveNf = protNFe ? protNFe.getChild('infProt', nfeNs).getChildText('chNFe', nfeNs) : '';

    if (!chaveNf) throw new Error(`Chave NF não encontrada no arquivo ${file.getName()}`);
    if (loggedKeys.has(chaveNf)) return {status: 'DUPLICATE'};

    const infNFe   = root.getChild('NFe', nfeNs).getChild('infNFe', nfeNs);
    const dest     = infNFe.getChild('dest', nfeNs);
    const enderDest= dest.getChild('enderDest', nfeNs);
    const nomeDestinatario = dest.getChildText('xNome', nfeNs);

    // >>> ADIÇÃO: captura de e-mail e telefone com fallback (ignora namespace se necessário)
    // E-mail costuma vir em <dest><email>; telefone normalmente em <enderDest><fone>, às vezes em <dest><fone>.
    const emailRaw = getTextIgnoreNs_(dest, 'email', nfeNs).trim();
    const foneRaw  = (getTextIgnoreNs_(enderDest, 'fone', nfeNs) || getTextIgnoreNs_(dest, 'fone', nfeNs)).trim();

    // Normalizações leves
    const email    = emailRaw ? emailRaw.toLowerCase() : "";
    const telefone = foneRaw  ? foneRaw.replace(/\D+/g, '') : "";

    const rowData = {
      "NOME DO DESTINATÁRIO*": nomeDestinatario, "NOME DA EMPRESA (EM CASO DE ENDEREÇO COMERCIAL)": "",
      "E-MAIL": email, "TELEFONE": telefone,
      "CPF / CNPJ CLIENTE*": dest.getChildText('CPF', nfeNs) || dest.getChildText('CNPJ', nfeNs) || '',
      "INSCR. ESTADUAL": "", "CEP*": enderDest.getChildText('CEP', nfeNs),
      "LOGRADOURO*": enderDest.getChildText('xLgr', nfeNs), "NÚMERO*": enderDest.getChildText('nro', nfeNs),
      "COMPLEMENTO": enderDest.getChildText('xCpl', nfeNs), "BAIRRO*": enderDest.getChildText('xBairro', nfeNs),
      "CIDADE*": enderDest.getChildText('xMun', nfeNs), "ESTADO*": enderDest.getChildText('UF', nfeNs),
      "A ENCOMENDA POSSUI NF?*": 'Sim', "CHAVE NF": chaveNf
    };

    const headers    = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextFreeRow= getActualLastRow(sheet) + 1;

    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i] || '').trim();
      if (header === "VOLUMES*") continue;
      if (rowData.hasOwnProperty(header)) {
        sheet.getRange(nextFreeRow, i + 1).setValue(rowData[header]);
      }
    }

    const timestamp = Utilities.formatDate(new Date(), CONFIG.timeZone, "dd/MM/yyyy HH:mm:ss");
    logSheet.appendRow([nomeDestinatario, chaveNf, timestamp]);

    return {status: 'SUCCESS', chaveNf: chaveNf};
  } catch (e) {
    Logger.log(`ERRO ao processar ${fileId}: ${e.stack}`);
    return {status: 'ERROR'};
  }
}

/**
 * Atualiza Log:
 *  - Coluna D (download) para as chaves exportadas.
 *  - Coluna E (nome do arquivo gerado), se fileName fornecido.
 * Usa escrita por colunas (B/D/E) para evitar mismatch com DataRange.
 */
function updateLogWithLink(exportedKeys, downloadLink, fileName) {
  const ss = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId);
  const logSheet = ss.getSheetByName(CONFIG.logSheetName);
  const totalRows = logSheet.getLastRow();
  if (totalRows < 2) return;

  // Coluna B (chave), D (download), E (nome arquivo)
  const rows = totalRows - 1;
  const colB = logSheet.getRange(2, 2, rows, 1).getValues(); // B2:B
  const colDRange = logSheet.getRange(2, 4, rows, 1);         // D2:D
  const colD = colDRange.getValues();
  let colERange = null, colE = null;

  if (fileName) {
    colERange = logSheet.getRange(2, 5, rows, 1); // E2:E (cria se não existir conteúdo)
    colE = colERange.getValues();
  }

  for (let i = 0; i < rows; i++) {
    const chave = String(colB[i][0] || '').trim();
    if (exportedKeys.has(chave)) {
      colD[i][0] = downloadLink;
      if (colE) colE[i][0] = fileName;
    }
  }
  colDRange.setValues(colD);
  if (colERange && colE) colERange.setValues(colE);
}

function getLoggedKeys() {
  try {
    const logSheet = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId).getSheetByName(CONFIG.logSheetName);
    if (!logSheet || logSheet.getLastRow() < 2) return new Set();
    const keys = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues().flat();
    return new Set(keys.map(String).filter(k => k));
  } catch (e) { return new Set(); }
}

function getActualLastRow(sheet) {
  const maxRows = sheet.getMaxRows();
  const columnAValues = sheet.getRange("A1:A" + maxRows).getValues();
  for (let i = columnAValues.length - 1; i >= 0; i--) {
    if (columnAValues[i][0] !== "") return i + 1;
  }
  return CONFIG.headerRow - 1;
}

// ===================================================================================================
// Atalhos do Log RT (inalterados)
// ===================================================================================================

function showLogWindowImportXml(fileIds) {
  if (typeof showLogWindowGeneric !== 'function') {
    throw new Error('showLogWindowGeneric não encontrado. Verifique se ProcessWrappers.gs foi adicionado.');
  }
  if (typeof LOGRT === 'undefined') {
    throw new Error('LOGRT não encontrado. Verifique se ConfigRT.gs foi adicionado.');
  }
  showLogWindowGeneric('executarImportacaoXml', [fileIds], LOGRT.UI.MODAL_TITLE_IMPORT);
}

function showLogWindowExportacao() {
  if (typeof showLogWindowGeneric !== 'function') {
    throw new Error('showLogWindowGeneric não encontrado. Verifique se ProcessWrappers.gs foi adicionado.');
  }
  if (typeof LOGRT === 'undefined') {
    throw new Error('LOGRT não encontrado. Verifique se ConfigRT.gs foi adicionado.');
  }
  showLogWindowGeneric('executarExportacaoXlsx', [], LOGRT.UI.MODAL_TITLE_EXPORT);
}

// ===================================================================================================
// Utilidades
// ===================================================================================================

function indexMap(headers) {
  const map = {};
  headers.forEach((h, i) => { map[String(h || '').trim()] = i; });
  return map;
}

/** Set de (nome||chave) apenas quando coluna D está preenchida (já gerados) */
function readExportedPairsSet_(logSheet) {
  const set = new Set();
  if (!logSheet) return set;
  const last = logSheet.getLastRow();
  if (last < 2) return set;

  const values = logSheet.getRange(2, 1, last - 1, 4).getValues(); // A:D
  for (const [nome, chave, , dUrl] of values) {
    const n = String(nome  || '').trim();
    const c = String(chave || '').trim();
    const d = String(dUrl  || '').trim();
    if (n && c && d) set.add(n + '||' + c);
  }
  return set;
}

function getOrCreateSubfolder_(parentFolder, name) {
  const iter = parentFolder.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parentFolder.createFolder(name);
}

function getMonthFolderName_(date, tz) {
  const mm = Utilities.formatDate(date, tz, 'MM');
  const m  = parseInt(Utilities.formatDate(date, tz, 'M'), 10);
  const nomes = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return `${mm}-${nomes[m - 1]}`;
}

function countFilesCreatedOn_(folder, date, tz) {
  const start = new Date(date); start.setHours(0,0,0,0);
  const end   = new Date(date); end.setHours(23,59,59,999);
  let count = 0;
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const dc = f.getDateCreated();
    if (dc >= start && dc <= end) count++;
  }
  return count;
}

function buildDownloadUrlFromId_(id) {
  return `https://drive.google.com/uc?export=download&id=${id}`;
}

function deriveOpenUrlFromDownloadUrl_(downloadUrl) {
  const m = String(downloadUrl).match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (m) return `https://drive.google.com/file/d/${m[1]}/view?usp=drivesdk`;
  const m2 = String(downloadUrl).match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m2) return `https://docs.google.com/spreadsheets/d/${m2[1]}/edit`;
  return downloadUrl;
}

function parseLogTimestamp_(str, tz) {
  const m = String(str).match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return new Date();
  const dd = parseInt(m[1], 10);
  const MM = parseInt(m[2], 10);
  const yyyy = parseInt(m[3], 10);
  const HH = parseInt(m[4], 10);
  const mm = parseInt(m[5], 10);
  const ss = m[6] ? parseInt(m[6], 10) : 0;
  return new Date(yyyy, MM - 1, dd, HH, mm, ss);
}

/** Helper para embutir HTML por nome (caso você ainda use em outras telas) */
function incluir(nome) {
  var base = String(nome || '').replace(/\.html?$/i, '');
  return HtmlService.createHtmlOutputFromFile(base).getContent();
}

/** =========================
 *  ADIÇÕES PARA EXTRAÇÃO NS
 *  =========================
 */

/** Busca recursiva pelo primeiro elemento com esse nome (ignorando namespace) e retorna o texto. */
function findFirstDescendantText_(element, localName) {
  if (!element) return '';
  if (element.getName && element.getName() === localName) {
    return element.getText() || '';
  }
  const children = element.getChildren();
  for (var i = 0; i < children.length; i++) {
    var hit = findFirstDescendantText_(children[i], localName);
    if (hit) return hit;
  }
  return '';
}

/** Tenta pegar o texto de um filho direto (com namespace) e, se não achar, faz busca recursiva ignorando ns. */
function getTextIgnoreNs_(parentEl, localName, ns) {
  if (!parentEl) return '';
  try {
    var direct = parentEl.getChildText(localName, ns);
    if (direct) return direct;
  } catch (e) {}
  return findFirstDescendantText_(parentEl, localName);
}
