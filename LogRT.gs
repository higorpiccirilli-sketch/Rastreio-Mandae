/**
 * =============================================================================
 * File: LogRT.gs
 * Módulo: Logger em Tempo Real (Cache + Polling)
 * Versão: 1.0.0
 * Autores: Higor Piccirilli Soares Batista & GPT-5 Thinking
 * Data: 2025-10-23
 * Descrição:
 *   Fornece:
 *     - API para acumular mensagens de status por execução (executionId)
 *     - Endpoint polled pelo front-end (obterStatusDoLancamento)
 *   Seguro (textContent no front) e não invasivo (não mexe nas funções existentes).
 * =============================================================================
 */

const RT = {
  _key(id) { return `${LOGRT.CACHE_KEY_BASE}${id}`; },

  push(executionId, msg) {
    const cache = CacheService.getUserCache();
    const arr = JSON.parse(cache.get(this._key(executionId)) || '[]');
    arr.push(msg);
    cache.put(this._key(executionId), JSON.stringify(arr), LOGRT.CACHE_TTL_SEC);
  },

  last(executionId) {
    const cache = CacheService.getUserCache();
    const arr = JSON.parse(cache.get(this._key(executionId)) || '[]');
    return arr.length ? arr[arr.length - 1] : null;
  },

  clear(executionId) {
    CacheService.getUserCache().remove(this._key(executionId));
  }
};

/**
 * Endpoint polled pelo HTML do Log.
 * Retorna apenas a última linha (o front só acrescenta quando muda).
 */
function obterStatusDoLancamento(executionId) {
  return RT.last(executionId);
}
