/**
 * 07_StaffService.gs - 教職員データ管理サービス
 *
 * 教職員データのCRUD操作、メールアドレス一覧取得などを提供する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs, 03_Auth.gs
 */

const StaffService = {

  /**
   * 教職員一覧を取得（管理画面用）
   * @returns {{ success: boolean, staff?: Array, headers?: Array, error?: string }}
   */
  getStaffList() {
    Auth.requireAdmin('教職員一覧取得');
    try {
      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.TEACHER);
      if (!sheet) {
        throw new Error('教職員データシートが見つかりません');
      }
      const data = SheetAccess.getDataWithRetry(sheet, '教職員一覧');
      if (data.length === 0) {
        return { success: true, staff: [], headers: [] };
      }
      const headers = data[0];
      const staff = SheetUtils.toObjects(data).filter(m => m['メールアドレス']);
      return { success: true, staff: staff, headers: headers };
    } catch (error) {
      return Logger_.handleError(error, '教職員一覧取得');
    }
  },

  /**
   * 教職員を追加
   * @param {Object} staffData - 教職員データ
   * @returns {{ success: boolean, error?: string }}
   */
  addStaff(staffData) {
    Auth.requireAdmin('教職員追加');
    try {
      if (!staffData || !staffData['メールアドレス']) {
        throw new Error('メールアドレスは必須です');
      }
      const email = String(staffData['メールアドレス']).trim().toLowerCase();
      if (!email) {
        throw new Error('メールアドレスは必須です');
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.TEACHER);
      if (!sheet) {
        throw new Error('教職員データシートが見つかりません');
      }
      const allData = SheetAccess.getDataWithRetry(sheet, '教職員追加');
      const headers = allData[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      // メールアドレス重複チェック
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      if (emailIdx === -1) {
        throw new Error('教職員データシートにメールアドレス列が見つかりません');
      }
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][emailIdx]).trim().toLowerCase() === email) {
          return { success: false, error: '同じメールアドレスの教職員が既に存在します: ' + email };
        }
      }

      // ヘッダー順に新しい行データを構築
      const newRow = headers.map(header =>
        staffData.hasOwnProperty(header) ? staffData[header] : ''
      );
      sheet.appendRow(newRow);

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('教職員追加: ' + email);
      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '教職員追加');
    }
  },

  /**
   * 教職員を更新
   * @param {string} originalEmail - 元のメールアドレス
   * @param {Object} staffData - 更新後の教職員データ
   * @returns {{ success: boolean, error?: string }}
   */
  updateStaff(originalEmail, staffData) {
    Auth.requireAdmin('教職員更新');
    try {
      if (!originalEmail || !staffData || !staffData['メールアドレス']) {
        throw new Error('メールアドレスは必須です');
      }
      const newEmail = String(staffData['メールアドレス']).trim().toLowerCase();
      if (!newEmail) {
        throw new Error('メールアドレスは必須です');
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.TEACHER);
      if (!sheet) {
        throw new Error('教職員データシートが見つかりません');
      }
      const allData = SheetAccess.getDataWithRetry(sheet, '教職員更新');
      const headers = allData[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      if (emailIdx === -1) {
        throw new Error('教職員データシートにメールアドレス列が見つかりません');
      }

      // 該当行を検索
      let targetRow = -1;
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][emailIdx]).trim().toLowerCase() === String(originalEmail).trim().toLowerCase()) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow === -1) {
        throw new Error('教職員が見つかりません: ' + originalEmail);
      }

      // メールアドレス変更の場合は新メールアドレスの重複チェック
      if (newEmail !== String(originalEmail).trim().toLowerCase()) {
        for (let i = 1; i < allData.length; i++) {
          if (i + 1 === targetRow) continue;
          if (String(allData[i][emailIdx]).trim().toLowerCase() === newEmail) {
            return { success: false, error: '同じメールアドレスの教職員が既に存在します: ' + newEmail };
          }
        }
      }

      // 各列の値を更新
      headers.forEach((header, colIndex) => {
        if (staffData.hasOwnProperty(header)) {
          sheet.getRange(targetRow, colIndex + 1).setValue(staffData[header]);
        }
      });

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('教職員更新: ' + originalEmail + (newEmail !== String(originalEmail).trim().toLowerCase() ? ' → ' + newEmail : ''));
      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '教職員更新');
    }
  },

  /**
   * 教職員を削除
   * @param {string} email - メールアドレス
   * @returns {{ success: boolean, error?: string }}
   */
  deleteStaff(email) {
    Auth.requireAdmin('教職員削除');
    try {
      if (!email) {
        throw new Error('メールアドレスが指定されていません');
      }
      email = String(email).trim().toLowerCase();

      // 自分自身の削除を禁止
      const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
      if (email === currentUserEmail) {
        return { success: false, error: '自分自身を削除することはできません' };
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.TEACHER);
      if (!sheet) {
        throw new Error('教職員データシートが見つかりません');
      }
      const allData = SheetAccess.getDataWithRetry(sheet, '教職員削除');
      const headers = allData[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      if (emailIdx === -1) {
        throw new Error('教職員データシートにメールアドレス列が見つかりません');
      }

      let targetRow = -1;
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][emailIdx]).trim().toLowerCase() === email) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow === -1) {
        throw new Error('教職員が見つかりません: ' + email);
      }

      sheet.deleteRow(targetRow);

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('教職員削除: ' + email);
      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '教職員削除');
    }
  },

  /**
   * 教職員データを一括置換
   * @param {Array<Object>} staffArray - 教職員データの配列
   * @returns {{ success: boolean, count?: number, error?: string }}
   */
  replaceAllStaff(staffArray) {
    Auth.requireAdmin('教職員データ一括保存');
    try {
      if (!Array.isArray(staffArray) || staffArray.length === 0) {
        throw new Error('教職員データが空です');
      }

      // 自分自身のメールが含まれているか確認
      const currentEmail = Session.getActiveUser().getEmail().trim().toLowerCase();
      let found = false;
      const emailSet = {};
      for (let i = 0; i < staffArray.length; i++) {
        const e = String(staffArray[i]['メールアドレス'] || '').trim().toLowerCase();
        if (!e) {
          throw new Error((i + 1) + '行目: メールアドレスが空です');
        }
        if (emailSet[e]) {
          throw new Error('メールアドレスが重複しています: ' + e);
        }
        emailSet[e] = true;
        if (e === currentEmail) found = true;
      }
      if (!found) {
        throw new Error('自分自身のメールアドレスが含まれていません。自分を削除することはできません。');
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.TEACHER);
      if (!sheet) {
        throw new Error('教職員データシートが見つかりません');
      }
      const allData = sheet.getDataRange().getValues();
      const headers = allData[0];

      // 既存データ行をすべて削除（ヘッダー行は残す）
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }

      // 新データを書き込み
      const rows = SheetUtils.toRows(staffArray, headers);
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('教職員データ一括保存: ' + rows.length + '件');
      return { success: true, count: rows.length };
    } catch (error) {
      return Logger_.handleError(error, '教職員データ一括保存');
    }
  },

  /**
   * 教職員データからメールアドレスで教員情報を取得
   * @param {string} email - メールアドレス
   * @returns {{ name: string, 学年?: *, 組?: * }|null}
   */
  getTeacherData(email) {
    try {
      const data = SheetAccess.getTeacherData();
      if (!data || data.length < 2) return null;

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      let nameIdx = SheetUtils.getColIndex(headerMap, '氏名');
      if (nameIdx === -1) nameIdx = SheetUtils.getColIndex(headerMap, '名前');
      const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
      const classIdx = SheetUtils.getColIndex(headerMap, '組');

      if (emailIdx === -1 || nameIdx === -1) return null;

      const normalizedEmail = String(email).trim().toLowerCase();

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (String(row[emailIdx]).trim().toLowerCase() === normalizedEmail) {
          return {
            name: row[nameIdx],
            学年: gradeIdx !== -1 ? row[gradeIdx] : undefined,
            組: classIdx !== -1 ? row[classIdx] : undefined
          };
        }
      }

      return null;

    } catch (error) {
      Logger_.log(`教員データ取得エラー: ${error.toString()}`, 'ERROR');
      return null;
    }
  },

  /**
   * 教職員名を取得（下位互換）
   * @param {string} email - メールアドレス
   * @returns {string|null}
   */
  getTeacherName(email) {
    const teacherData = this.getTeacherData(email);
    return teacherData ? teacherData.name : null;
  },

  /**
   * 教職員メールアドレス一覧を取得
   * @returns {Array<string>} メールアドレスの配列
   */
  getTeacherEmails() {
    Auth.requireTeacherOrAdmin('教職員メール取得');
    const data = SheetAccess.getTeacherData();
    if (data.length < 2) return [];

    const headerMap = SheetUtils.buildHeaderMap(data[0]);
    const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
    if (emailIdx === -1) return [];

    const emails = [];
    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][emailIdx] || '').trim().toLowerCase();
      if (email) emails.push(email);
    }
    return emails;
  },
};

// --- 後方互換ラッパー ---
function getStaffList() { return StaffService.getStaffList(); }
function addStaff(staffData) { return StaffService.addStaff(staffData); }
function updateStaff(originalEmail, staffData) { return StaffService.updateStaff(originalEmail, staffData); }
function deleteStaff(email) { return StaffService.deleteStaff(email); }
function replaceAllStaff(staffArray) { return StaffService.replaceAllStaff(staffArray); }
function getTeacherData(email) { return StaffService.getTeacherData(email); }
function getTeacherName(email) { return StaffService.getTeacherName(email); }
function getTeacherEmails() { return StaffService.getTeacherEmails(); }
