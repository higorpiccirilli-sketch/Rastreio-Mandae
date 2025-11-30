/**
 * =============================================================================
 * File: ConfigRT.gs
 * Módulo: Configurações do Logger em Tempo Real (não conflita com CONFIG existente)
 * Versão: 1.0.0
 * Autores: Higor Piccirilli Soares Batista & GPT-5 Thinking
 * Data: 2025-10-23
 * Descrição:
 *   Parâmetros e strings auxiliares para o recurso de "Log em tempo real"
 *   acoplado às funções já existentes do projeto.
 * =============================================================================
 * Compatibilidade:
 *   - Não altera a constante CONFIG já usada no projeto (onOpen.gs).
 *   - Pode ser removido sem impactar a lógica original — apenas o Log deixará de existir.
 * =============================================================================
 */

const LOGRT = Object.freeze({
  CACHE_KEY_BASE: 'lanc_status_',
  CACHE_TTL_SEC: 6 * 60 * 60,               // 6 horas
  UI: {
    MODAL_TITLE_IMPORT: 'Importação XML — Log em Tempo Real',
    MODAL_TITLE_EXPORT: 'Exportação XLSX — Log em Tempo Real',
    WIDTH: 900,
    HEIGHT: 560,
  }
});
