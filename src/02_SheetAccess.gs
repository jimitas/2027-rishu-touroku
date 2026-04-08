/**
 * 02_SheetAccess.gs - キャッシュ付きデータアクセス基盤
 *
 * 元コードで4つのキャッシュ取得関数（getSubmissionSheetDataCached,
 * getCourseDataCached, getTeacherDataCached, getStudentRosterDataCached）が
 * 同一パターンで重複していた。これを汎用キャッシュラッパーで統一する。
 *
 * 3段階キャッシュ:
 *   1. インメモリキャッシュ（同一実行内で有効）
 *   2. GAS ScriptCache（複数実行間で共有、TTL付き）
 *   3. シートから直接取得（フォールバック）
 */

// インメモリキャッシュ（GASは各実行が独立V8インスタンスのため古いデータのリスクなし）
let _memCache = {};

const SheetAccess = {

  // --- スプレッドシート・シート参照の遅延取得 ---

  /** @private */
  _ss: null,

  /** @private シート名→シートオブジェクトのキャッシュ */
  _sheets: {},

  /**
   * スプレッドシートの遅延取得
   * @returns {SpreadsheetApp.Spreadsheet}
   */
  getSpreadsheet() {
    if (!this._ss) {
      this._ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    return this._ss;
  },

  /**
   * シートの遅延取得（名前→シートのマッピング）
   * 元コード: getCourseSheet(), getTeacherSheet() 等5つの関数を統一
   *
   * @param {string} sheetName - シート名
   * @returns {SpreadsheetApp.Sheet|null}
   */
  getSheet(sheetName) {
    if (!this._sheets[sheetName]) {
      this._sheets[sheetName] = this.getSpreadsheet().getSheetByName(sheetName);
    }
    return this._sheets[sheetName];
  },

  /**
   * シートの存在確認（複数シート一括）
   *
   * @param {Array<string>} sheetNames - チェックするシート名の配列
   * @returns {{ valid: boolean, missing: Array<string> }}
   */
  validateSheets(sheetNames) {
    const missing = sheetNames.filter(name => !this.getSheet(name));
    return { valid: missing.length === 0, missing };
  },

  // --- 汎用キャッシュ付きデータ取得 ---

  /**
   * 汎用キャッシュ付きデータ取得
   * 3段階フォールバック: インメモリ → ScriptCache → シート直接取得
   *
   * @param {string} sheetName - シート名（CONFIG.SHEETS の値）
   * @param {string} cacheKey - CacheServiceのキー（CONFIG.CACHE_KEYS の値）
   * @param {number} ttlSec - キャッシュ有効期間（秒）
   * @returns {Array<Array>} シートデータ（2次元配列）
   */
  getCachedData(sheetName, cacheKey, ttlSec) {
    // 1. インメモリキャッシュ
    if (_memCache[cacheKey]) {
      return _memCache[cacheKey];
    }

    // 2. ScriptCache
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        _memCache[cacheKey] = parsed;
        return parsed;
      }
    } catch (e) {
      // CacheService障害時はフォールバック
      Logger_.log(`ScriptCache読み取り失敗 (${cacheKey}): ${e.toString()}`, 'WARN');
    }

    // 3. シートから直接取得
    const sheet = this.getSheet(sheetName);
    if (!sheet) {
      Logger_.log(`シートが見つかりません: ${sheetName}`, 'WARN');
      return [];
    }

    const data = this.getDataWithRetry(sheet, sheetName);

    // ScriptCacheに保存（サイズ制限チェック付き）
    this._putCache(cacheKey, data, ttlSec);

    _memCache[cacheKey] = data;
    return data;
  },

  /**
   * リトライ付きデータ取得
   * スプレッドシートAPIの一時的なエラーに対応
   *
   * @param {SpreadsheetApp.Sheet} sheet - シートオブジェクト
   * @param {string} context - ログ用コンテキスト名
   * @param {number} maxRetries - 最大リトライ回数
   * @returns {Array<Array>} シートデータ
   */
  getDataWithRetry(sheet, context = '', maxRetries = 2) {
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        return sheet.getDataRange().getValues();
      } catch (e) {
        if (attempt === maxRetries) {
          Logger_.log(`データ取得失敗 (${context}): ${e.toString()}`, 'ERROR');
          throw e;
        }
        Logger_.log(`データ取得リトライ ${attempt + 1}/${maxRetries} (${context})`, 'WARN');
        Utilities.sleep(500 * (attempt + 1));
      }
    }
  },

  // --- キャッシュ管理 ---

  /**
   * ScriptCacheにデータを保存（サイズ制限チェック付き）
   * @private
   */
  _putCache(cacheKey, data, ttlSec) {
    try {
      const json = JSON.stringify(data);
      if (json.length < CONFIG.CACHE_MAX_SIZE) {
        CacheService.getScriptCache().put(cacheKey, json, ttlSec);
      } else {
        Logger_.log(`キャッシュサイズ超過 (${cacheKey}): ${json.length} bytes`, 'WARN');
      }
    } catch (e) {
      Logger_.log(`ScriptCache書き込み失敗 (${cacheKey}): ${e.toString()}`, 'WARN');
    }
  },

  /**
   * 特定のキャッシュをクリア
   * @param {string} cacheKey - クリアするキャッシュキー
   */
  clearCache(cacheKey) {
    delete _memCache[cacheKey];
    try {
      CacheService.getScriptCache().remove(cacheKey);
    } catch (e) {
      // 無視
    }
  },

  /**
   * 全キャッシュをクリア
   */
  clearAllCaches() {
    _memCache = {};
    try {
      const allKeys = Object.values(CONFIG.CACHE_KEYS);
      CacheService.getScriptCache().removeAll(allKeys);
    } catch (e) {
      // 無視
    }
  },

  // --- 便利メソッド（各シートデータの取得）---

  /** 科目データを取得（キャッシュ付き、TTL長め） */
  getCourseData() {
    return this.getCachedData(CONFIG.SHEETS.COURSE, CONFIG.CACHE_KEYS.COURSE, 3600);
  },

  /** 登録データを取得（キャッシュ付き、TTL短め） */
  getSubmissionData() {
    return this.getCachedData(CONFIG.SHEETS.SUBMISSION, CONFIG.CACHE_KEYS.SUBMISSION, CONFIG.CACHE_EXPIRY_SEC);
  },

  /** 教職員データを取得（キャッシュ付き、TTL長め） */
  getTeacherData() {
    return this.getCachedData(CONFIG.SHEETS.TEACHER, CONFIG.CACHE_KEYS.TEACHER, 3600);
  },

  /** 生徒名簿を取得（キャッシュ付き、TTL長め） */
  getRosterData() {
    return this.getCachedData(CONFIG.SHEETS.ROSTER, CONFIG.CACHE_KEYS.ROSTER, 3600);
  },

  /** 設定データを取得（キャッシュ付き） */
  getSettingsData() {
    return this.getCachedData(CONFIG.SHEETS.SETTINGS, CONFIG.CACHE_KEYS.SETTINGS, CONFIG.CACHE_EXPIRY_SEC);
  },

  // --- シートへの書き込みヘルパー ---

  /**
   * シートの特定行を更新
   *
   * @param {string} sheetName - シート名
   * @param {number} rowIndex - 行番号（1ベース）
   * @param {Array} rowData - 書き込むデータ
   */
  updateRow(sheetName, rowIndex, rowData) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) throw new Error(`シートが見つかりません: ${sheetName}`);
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  },

  /**
   * シートの末尾に行を追加
   *
   * @param {string} sheetName - シート名
   * @param {Array<Array>} rows - 追加する行データの配列
   */
  appendRows(sheetName, rows) {
    if (!rows || rows.length === 0) return;
    const sheet = this.getSheet(sheetName);
    if (!sheet) throw new Error(`シートが見つかりません: ${sheetName}`);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  },
};
