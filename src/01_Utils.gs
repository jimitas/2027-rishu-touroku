/**
 * 01_Utils.gs - 共通ユーティリティ
 *
 * 元コードで20箇所以上重複していたパターンを共通関数として提供する。
 * - ヘッダーMap構築（headers.indexOf() の O(1) 化）
 * - 行検索（for-loop パターンの統一）
 * - データ変換（配列 → オブジェクト配列）
 * - タイムスタンプ、ログ、エラーハンドリング
 * - セル値のサニタイズ
 */

// =============================================
// SheetUtils: シートデータ操作ユーティリティ
// =============================================
const SheetUtils = {

  /**
   * ヘッダー配列から列名→インデックスのMapを構築
   * 元コード: headers.indexOf('学籍番号') が20箇所以上で繰り返されていたパターンを O(1) 化
   *
   * @param {Array<string>} headers - ヘッダー行の配列
   * @returns {Map<string, number>} 列名 → 0ベースインデックス
   */
  buildHeaderMap(headers) {
    const map = new Map();
    if (!headers) return map;
    headers.forEach((h, i) => {
      if (h != null && h !== '') map.set(String(h), i);
    });
    return map;
  },

  /**
   * ヘッダーMapから列インデックスを取得（0ベース）
   * 存在しない場合は -1 を返す
   *
   * @param {Map} headerMap - buildHeaderMap() の戻り値
   * @param {string} columnName - 列名
   * @returns {number} 0ベースインデックス、存在しない場合は -1
   */
  getColIndex(headerMap, columnName) {
    const idx = headerMap.get(columnName);
    return idx !== undefined ? idx : -1;
  },

  /**
   * 必須列の存在を一括チェック
   * 不足があればエラーをスローする
   *
   * @param {Map} headerMap - buildHeaderMap() の戻り値
   * @param {Array<string>} columnNames - 必須列名の配列
   * @throws {Error} 不足列がある場合
   */
  requireColumns(headerMap, columnNames) {
    const missing = columnNames.filter(name => !headerMap.has(name));
    if (missing.length > 0) {
      throw new Error(`必須列が見つかりません: ${missing.join(', ')}`);
    }
  },

  /**
   * データ配列から条件に合う行を検索
   * 元コード: for (let i = 1; ...) if (data[i][idx] === value) パターンが10箇所以上
   *
   * @param {Array<Array>} data - シートデータ（[0]がヘッダー）
   * @param {Map} headerMap - buildHeaderMap() の戻り値
   * @param {string} columnName - 検索対象の列名
   * @param {*} searchValue - 検索値
   * @returns {{ rowIndex: number, dataIndex: number, rowData: Array } | null}
   *   rowIndex: シート上の行番号（1ベース）、dataIndex: 配列上のインデックス
   */
  findRow(data, headerMap, columnName, searchValue) {
    const colIdx = headerMap.get(columnName);
    if (colIdx === undefined) return null;
    const searchStr = String(searchValue);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdx]) === searchStr) {
        return { rowIndex: i + 1, dataIndex: i, rowData: data[i] };
      }
    }
    return null;
  },

  /**
   * データ配列から条件に合う全行を検索
   *
   * @param {Array<Array>} data - シートデータ（[0]がヘッダー）
   * @param {Map} headerMap - buildHeaderMap() の戻り値
   * @param {string} columnName - 検索対象の列名
   * @param {*} searchValue - 検索値
   * @returns {Array<{ rowIndex: number, dataIndex: number, rowData: Array }>}
   */
  findAllRows(data, headerMap, columnName, searchValue) {
    const colIdx = headerMap.get(columnName);
    if (colIdx === undefined) return [];
    const searchStr = String(searchValue);
    const results = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colIdx]) === searchStr) {
        results.push({ rowIndex: i + 1, dataIndex: i, rowData: data[i] });
      }
    }
    return results;
  },

  /**
   * データ配列（[0]がヘッダー）をオブジェクト配列に変換
   * 元コード: getCourseList, getStaffList等で個別に行っていたパターンを統一
   *
   * @param {Array<Array>} data - シートデータ
   * @param {Object} options - オプション
   * @param {boolean} options.skipEmpty - 空行をスキップするか（デフォルト: true）
   * @returns {Array<Object>} オブジェクト配列
   */
  toObjects(data, options = {}) {
    const { skipEmpty = true } = options;
    if (!data || data.length < 2) return [];
    const headers = data[0];
    return data.slice(1)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h != null && h !== '') obj[String(h)] = row[i];
        });
        return obj;
      })
      .filter(obj => {
        if (!skipEmpty) return true;
        return Object.values(obj).some(v => v !== '' && v != null);
      });
  },

  /**
   * オブジェクト配列をシート書き込み用の2次元配列に変換
   *
   * @param {Array<Object>} objects - オブジェクト配列
   * @param {Array<string>} headers - ヘッダー順序
   * @returns {Array<Array>} 2次元配列（ヘッダー行なし）
   */
  toRows(objects, headers) {
    return objects.map(obj =>
      headers.map(h => obj[h] != null ? obj[h] : '')
    );
  },

  /**
   * Lookup Map を構築（学籍番号→行データ等の高速検索用）
   * 元コード: O(n^2) ネストループを O(n) に改善するために使用
   *
   * @param {Array<Array>} data - シートデータ
   * @param {Map} headerMap - buildHeaderMap() の戻り値
   * @param {string} keyColumn - キーとなる列名
   * @returns {Map<string, Array>} キー → 行データ
   */
  buildLookupMap(data, headerMap, keyColumn) {
    const colIdx = headerMap.get(keyColumn);
    if (colIdx === undefined) return new Map();
    const map = new Map();
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][colIdx]);
      if (key) map.set(key, data[i]);
    }
    return map;
  },

  /**
   * セル値のサニタイズ（数式インジェクション防止）
   * スプレッドシートに書き込む文字列値から危険な先頭文字を除去する
   *
   * @param {*} value - セル値
   * @returns {string} サニタイズ済みの値
   */
  sanitizeCellValue(value) {
    if (value == null) return '';
    const str = String(value);
    if (/^[=+\-@]/.test(str)) return "'" + str;
    return str;
  },
};

// =============================================
// Logger: ログ・タイムスタンプ・エラーハンドリング
// =============================================
const Logger_ = {

  /**
   * 日本時間のタイムスタンプを取得
   * @returns {string} yyyy-MM-dd HH:mm:ss 形式
   */
  getJSTTimestamp() {
    return Utilities.formatDate(
      new Date(),
      CONFIG.DEFAULTS.TIMEZONE,
      CONFIG.DEFAULTS.DATE_FORMAT
    );
  },

  /**
   * ログ出力
   * @param {string} message - メッセージ
   * @param {string} level - ログレベル (INFO, WARN, ERROR)
   */
  log(message, level = 'INFO') {
    const timestamp = this.getJSTTimestamp();
    console.log(`[${timestamp}] ${level}: ${message}`);
  },

  /**
   * エラーハンドリング用ヘルパー
   * 統一されたエラーレスポンスを生成する
   *
   * @param {Error} error - エラーオブジェクト
   * @param {string} context - エラー発生コンテキスト
   * @returns {{ success: false, error: string, timestamp: string }}
   */
  handleError(error, context) {
    const errorMsg = `${context}: ${error.toString()}`;
    this.log(errorMsg, 'ERROR');
    return {
      success: false,
      error: errorMsg,
      timestamp: this.getJSTTimestamp(),
    };
  },

  /**
   * 成功レスポンスを生成
   *
   * @param {Object} data - レスポンスデータ
   * @param {string} message - オプションのメッセージ
   * @returns {{ success: true, timestamp: string, ... }}
   */
  successResponse(data = {}, message = '') {
    return {
      success: true,
      timestamp: this.getJSTTimestamp(),
      message,
      ...data,
    };
  },
};

// =============================================
// Validator: 汎用バリデーション
// =============================================
const Validator = {

  /**
   * 学年値を安全にパース
   * @param {*} value - 学年値
   * @param {number} defaultValue - デフォルト値
   * @returns {number} パースされた学年（0-3）
   */
  safeParseGrade(value, defaultValue = 1) {
    if (value == null || value === '') return defaultValue;
    const parsed = parseInt(value, 10);
    if (isNaN(parsed) || parsed < 0 || parsed > 3) return defaultValue;
    return parsed;
  },

  /**
   * カンマ区切り文字列をトリム済み配列に変換
   * 元コード: String(x).split(',').map(s => s.trim()) が多数重複
   *
   * @param {string} str - カンマ区切り文字列
   * @returns {Array<string>} トリム済みの配列（空文字列を除外）
   */
  splitCSV(str) {
    if (!str) return [];
    return String(str).split(',').map(s => s.trim()).filter(Boolean);
  },

  /**
   * カンマ区切り文字列を数値配列に変換
   * @param {string} str - カンマ区切り文字列
   * @returns {Array<number>}
   */
  splitCSVNumbers(str) {
    return this.splitCSV(str).map(Number).filter(n => !isNaN(n));
  },

  /**
   * 不可視文字を除去
   * 元コード: index.html内で5回ローカル定義されていたパターン
   *
   * @param {*} str - 入力文字列
   * @returns {string} クリーンな文字列
   */
  stripInvisible(str) {
    if (!str) return '';
    return String(str).replace(/[\uFEFF\u200B\u200C\u200D\uFFFE]/g, '').trim();
  },
};
