/**
 * 08_SettingsService.gs - 設定管理
 *
 * 設定シート（Key-Value）の読み書きと、
 * サーバー側の機能制限チェックを行う。
 *
 * 依存: 02_SheetAccess.gs, 03_Auth.gs
 */

const SettingsService = {

  /**
   * 設定を取得（Key-Valueオブジェクト）
   * キャッシュ付き
   *
   * @returns {Object} 設定オブジェクト { 設定名: 内容, ... }
   */
  getSettings() {
    try {
      const data = SheetAccess.getSettingsData();
      if (data.length <= 1) return {};

      const settings = {};
      for (let i = 1; i < data.length; i++) {
        const key = data[i][0];
        const value = data[i][1];
        if (key) settings[String(key)] = value;
      }
      return settings;

    } catch (error) {
      Logger_.log(`設定データ取得エラー: ${error.toString()}`, 'ERROR');
      return { [CONFIG.SETTING_KEYS.SCHOOL_NAME]: CONFIG.DEFAULTS.SCHOOL_NAME };
    }
  },

  /**
   * 設定を更新（管理者のみ）
   *
   * @param {Object} settingsData - { 設定名: 内容, ... }
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  updateSettings(settingsData) {
    try {
      Auth.requireAdmin('設定の更新');

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.SETTINGS);
      if (!sheet) throw new Error('設定シートが見つかりません');

      const data = sheet.getDataRange().getValues();

      // 既存項目を更新
      for (let i = 1; i < data.length; i++) {
        const settingName = data[i][0];
        if (settingName && settingsData.hasOwnProperty(settingName)) {
          sheet.getRange(i + 1, 2).setValue(settingsData[settingName]);
        }
      }

      // 新規項目を追加
      const existingKeys = new Set(data.map(row => row[0]));
      Object.keys(settingsData).forEach(key => {
        if (!existingKeys.has(key)) {
          sheet.appendRow([key, settingsData[key]]);
        }
      });

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SETTINGS);

      return Logger_.successResponse({}, '設定を更新しました');

    } catch (error) {
      return Logger_.handleError(error, '設定更新');
    }
  },

  /**
   * 設定シートの初期化（存在しない場合に作成）
   */
  ensureSettingsSheet() {
    let sheet = SheetAccess.getSheet(CONFIG.SHEETS.SETTINGS);
    if (!sheet) {
      sheet = SheetAccess.getSpreadsheet().insertSheet(CONFIG.SHEETS.SETTINGS);
      const initData = CONFIG.INITIAL_SETTINGS;
      sheet.getRange(1, 1, initData.length, initData[0].length).setValues(initData);
      Logger_.log('設定シートを作成しました');
    }
    return sheet;
  },

  /**
   * サーバー側の機能制限をチェック
   *
   * @param {string} userRole - ユーザーのロール
   * @param {string} featureKey - 機能キー（設定シートのキー名）
   * @returns {{ allowed: boolean, reason: string }}
   */
  checkRestriction(userRole, featureKey) {
    const settings = this.getSettings();

    // メンテナンスモードチェック
    const maintenance = settings[CONFIG.SETTING_KEYS.MAINTENANCE] || '無効';
    if (maintenance === '管理者のみ' && !Auth.isAdmin(userRole)) {
      return { allowed: false, reason: 'メンテナンスモード中です（管理者のみアクセス可）' };
    }
    if (maintenance === '教員以上' && !Auth.isTeacherOrAbove(userRole)) {
      return { allowed: false, reason: 'メンテナンスモード中です（教員以上のみアクセス可）' };
    }

    // 機能別チェック
    if (featureKey) {
      const value = settings[featureKey];
      if (value === '不可') {
        return { allowed: false, reason: `${featureKey}は現在無効です` };
      }
    }

    return { allowed: true, reason: '' };
  },
};

// --- 後方互換ラッパー ---
function getSettings() { return SettingsService.getSettings(); }
function updateSettings(settingsData) { return SettingsService.updateSettings(settingsData); }
