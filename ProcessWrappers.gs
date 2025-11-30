/**
 * =============================================================================
 * File: ProcessWrappers.gs
 * Módulo: Wrappers com Log para processos existentes (Importação/Exportação)
 * Versão: 1.0.0
 * Autores: Higor Piccirilli Soares Batista & GPT-5 Thinking
 * Data: 2025-10-23
 * Descrição:
 *   - NÃO altera suas funções originais nem assinaturas.
 *   - Apenas adiciona wrappers "executar..." que:
 *       1) limpam estado do logger
 *       2) emitem mensagens RT.push(...)
 *       3) chamam suas funções atuais
 *       4) retornam mensagens finais para o LogUI
 * =============================================================================
 * Dependências:
 *   - LogRT.gs (RT.push/clear)
 *   - onOpen.gs (listXmlFiles, processSingleXmlFile, generateExportFile, getLoggedKeys, limparDadosSheet1Interna, etc.)
 * =============================================================================
 */

/**
 * Abre janela de Log para qualquer wrapper (genérico).
 * @param {string} processFn Nome do wrapper (ex.: 'executarImportacaoXml' ou 'executarExportacaoXlsx')
 * @param {Array=} args Lista de argumentos que o wrapper espera após executionId
 * @param {string=} modalTitle Título opcional
 */
function showLogWindowGeneric(processFn, args, modalTitle) {
  const t = HtmlService.createTemplateFromFile('LogUI');
  t.processFn = processFn;
  t.serializedArgs = args ? JSON.stringify(args) : '';

  const html = t.evaluate()
    .setWidth(LOGRT.UI.WIDTH)
    .setHeight(LOGRT.UI.HEIGHT);
  SpreadsheetApp.getUi().showModalDialog(html, modalTitle || 'Execução — Log em Tempo Real');
}

/**
 * Atalho específico para importar XMLs (recebe os fileIds do front).
 */
function showLogWindowImportXml(fileIds) {
  showLogWindowGeneric('executarImportacaoXml', [fileIds], LOGRT.UI.MODAL_TITLE_IMPORT);
}

/**
 * Atalho específico para exportar XLSX.
 */
function showLogWindowExportacao() {
  showLogWindowGeneric('executarExportacaoXlsx', [], LOGRT.UI.MODAL_TITLE_EXPORT);
}

/**
 * Wrapper com Log — Importação de XMLs (equivalente ao processMultipleXmlFiles + logs por arquivo).
 * NÃO altera a sua função existente. Só replica a mesma lógica, emitindo status.
 */
function executarImportacaoXml(executionId, fileIds) {
  RT.clear(executionId);
  try {
    if (!Array.isArray(fileIds) || fileIds.length === 0) {
      RT.push(executionId, '[ERRO] Nenhum arquivo selecionado para importar.');
      throw new Error('Nenhum arquivo selecionado.');
    }

    RT.push(executionId, `[EM ANDAMENTO] Preparando ambiente de importação...`);
    limparDadosSheet1Interna(); // mantém seu comportamento atual

    const summary = { successCount: 0, duplicateCount: 0, errorCount: 0 };
    const loggedKeys = getLoggedKeys();

    RT.push(executionId, `[EM ANDAMENTO] Importando ${fileIds.length} arquivo(s)...`);

    for (let idx = 0; idx < fileIds.length; idx++) {
      const id = fileIds[idx];
      try {
        RT.push(executionId, `[EM ANDAMENTO] (${idx + 1}/${fileIds.length}) Lendo arquivo ${id}...`);
        const result = processSingleXmlFile(id, loggedKeys);

        if (result.status === 'SUCCESS') {
          summary.successCount++;
          loggedKeys.add(result.chaveNf);
          RT.push(executionId, `[OK] NF importada: ${result.chaveNf}`);
        } else if (result.status === 'DUPLICATE') {
          summary.duplicateCount++;
          RT.push(executionId, `[OK] Duplicada ignorada (já no Log).`);
        } else {
          summary.errorCount++;
          RT.push(executionId, `[ERRO] Falha ao processar arquivo ${id}.`);
        }
      } catch (e) {
        summary.errorCount++;
        RT.push(executionId, `[ERRO] Exceção ao processar ${id}: ${e.message || e}`);
      }
    }

    RT.push(executionId, `[OK] Importação concluída. Sucesso: ${summary.successCount} | Duplicadas: ${summary.duplicateCount} | Erros: ${summary.errorCount}`);
    return `✅ ${summary.successCount} nota(s) importada(s).\n⚠️ ${summary.duplicateCount} nota(s) já existiam.\n❌ ${summary.errorCount} arquivo(s) com erro.`;
  } catch (err) {
    RT.push(executionId, `[ERRO] ${err.message || err}`);
    throw err;
  }
}

/**
 * Wrapper com Log — Exportação XLSX (chama sua função original).
 * NÃO altera a sua generateExportFile(); apenas loga etapas.
 */
function executarExportacaoXlsx(executionId) {
  RT.clear(executionId);
  try {
    RT.push(executionId, `[EM ANDAMENTO] Preparando exportação...`);
    RT.push(executionId, `[EM ANDAMENTO] Lendo dados da planilha e aplicando displayValues...`);
    const resultMsg = generateExportFile();
    RT.push(executionId, `[OK] Exportação finalizada.`);
    return resultMsg; // exibido como mensagem final no LogUI
  } catch (err) {
    RT.push(executionId, `[ERRO] ${err.message || err}`);
    throw err;
  }
}
