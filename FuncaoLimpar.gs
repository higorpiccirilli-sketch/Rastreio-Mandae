/**
 * Função para limpar os dados da Sheet1, preservando o cabeçalho e as colunas com fórmulas.
 * Esta função deve ser associada a um botão/imagem na sua planilha.
 */
// VERSÃO ATUALIZADA
function limparDadosSheet1() {
  const ui = SpreadsheetApp.getUi();
  const confirmacao = ui.alert(
    'Confirmar Limpeza',
    'Você tem certeza que deseja apagar TODOS os dados de envio da "Sheet1"?\n\nEsta ação não pode ser desfeita.\n(As fórmulas nas colunas protegidas serão preservadas).',
    ui.ButtonSet.YES_NO
  );

  if (confirmacao !== ui.Button.YES) {
    ui.alert('Operação cancelada.');
    return;
  }

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId).getSheetByName(CONFIG.importSheetName);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < CONFIG.headerRow + 1) {
      ui.alert('Limpeza não necessária', 'A planilha já está limpa.', ui.ButtonSet.OK);
      return;
    }
    
    const headers = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colunasProtegidasPorNome = ["VOLUMES*", "SERVIÇO DE ENVIO*"];
    const colunasProtegidasPorIndice = [15, 19]; // Coluna O = 15, Coluna S = 19
    const numLinhasParaLimpar = lastRow - CONFIG.headerRow;

    for (let i = 0; i < headers.length; i++) {
      const headerAtual = headers[i].trim();
      const colunaAtual = i + 1; // O índice da coluna

      if (colunasProtegidasPorNome.indexOf(headerAtual) === -1 && colunasProtegidasPorIndice.indexOf(colunaAtual) === -1) {
        sheet.getRange(CONFIG.headerRow + 1, colunaAtual, numLinhasParaLimpar, 1).clearContent();
      }
    }
    
    ui.alert('Sucesso!', 'Os dados da "Sheet1" foram limpos e as fórmulas foram preservadas.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Erro', `Ocorreu um erro durante a limpeza: ${e.message}`, ui.ButtonSet.OK);
    Logger.log("Erro na limpeza: " + e.stack);
  }
}


/**
 * Limpa os dados da Sheet1 de forma interna, sem alertas para o usuário.
 * Protege colunas por nome e por índice (O e S).
 */
function limparDadosSheet1Interna() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.mainSpreadsheetId).getSheetByName(CONFIG.importSheetName);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < CONFIG.headerRow + 1) {
      Logger.log('Limpeza não necessária, a planilha já está limpa.');
      return; // A planilha já está vazia, não faz nada.
    }
    
    const headers = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colunasProtegidasPorNome = ["VOLUMES*", "SERVIÇO DE ENVIO*"];
    const colunasProtegidasPorIndice = [15, 19]; // Coluna O = 15, Coluna S = 19
    const numLinhasParaLimpar = lastRow - CONFIG.headerRow;

    for (let i = 0; i < headers.length; i++) {
      const headerAtual = headers[i].trim();
      const colunaAtual = i + 1; // O índice da coluna (começando em 1)
      
      // Verifica se a coluna NÃO está na lista de nomes E NÃO está na lista de índices para proteger
      if (colunasProtegidasPorNome.indexOf(headerAtual) === -1 && colunasProtegidasPorIndice.indexOf(colunaAtual) === -1) {
        // Se não estiver protegida, limpa o conteúdo abaixo do cabeçalho
        sheet.getRange(CONFIG.headerRow + 1, colunaAtual, numLinhasParaLimpar, 1).clearContent();
      }
    }
     Logger.log('Limpeza interna da Sheet1 executada com sucesso.');
  } catch (e) {
    Logger.log("Erro na limpeza interna: " + e.stack);
    // Lança o erro para que a função principal possa capturá-lo se necessário
    throw new Error(`Ocorreu um erro ao tentar limpar a planilha antes da importação: ${e.message}`);
  }
}
