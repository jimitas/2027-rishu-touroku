/**
 * 03_Auth.gs - 認証・認可
 *
 * Google Workspace認証を利用したユーザー情報の取得、
 * ロール判定、権限チェックを行う。
 *
 * ロール階層:
 *   admin > kyomu > teacher > student
 *
 * 依存: 02_SheetAccess.gs (SheetAccess.getTeacherData())
 */

const Auth = {

  // --- ロール判定ヘルパー ---

  /**
   * ロールが教員以上かを判定
   * @param {string} role
   * @returns {boolean}
   */
  isTeacherOrAbove(role) {
    return role === CONFIG.ROLES.TEACHER
        || role === CONFIG.ROLES.KYOMU
        || role === CONFIG.ROLES.ADMIN;
  },

  /**
   * ロールが教務以上かを判定
   * @param {string} role
   * @returns {boolean}
   */
  isKyomuOrAbove(role) {
    return role === CONFIG.ROLES.KYOMU || role === CONFIG.ROLES.ADMIN;
  },

  /**
   * ロールが管理者かを判定
   * @param {string} role
   * @returns {boolean}
   */
  isAdmin(role) {
    return role === CONFIG.ROLES.ADMIN;
  },

  // --- ユーザー情報取得 ---

  /**
   * 現在のセッションからユーザー情報を取得
   * @returns {{ success: boolean, user: Object, timestamp: string }}
   */
  getUserInfo() {
    try {
      const email = Session.getActiveUser().getEmail();
      const userInfo = this.parseEmailForUserInfo(email);
      return Logger_.successResponse({ user: userInfo });
    } catch (error) {
      return Logger_.handleError(error, 'ユーザー認証');
    }
  },

  /**
   * メールアドレスからユーザー情報を解析
   * 教職員データシート → 登録データシートの順で検索
   *
   * @param {string} email - メールアドレス
   * @returns {Object} ユーザー情報
   * @throws {Error} どのデータにも該当しない場合
   */
  parseEmailForUserInfo(email) {
    // 1. 教職員データシートで検索（優先）
    const teacherResult = this.determineUserRole(email);
    if (teacherResult !== null && this.isTeacherOrAbove(teacherResult.role)) {
      return {
        email: email,
        role: teacherResult.role,
        name: teacherResult.name || '教職員',
        assignedClasses: teacherResult.assignedClasses,
      };
    }

    // 2. 登録データシートで生徒として検索
    const studentId = this._getStudentIdFromEmail(email);
    if (studentId) {
      const studentData = this._getStudentDataByIdInternal(studentId);
      return {
        studentId: studentId,
        name: studentData?.name || `学生${studentId}`,
        grade: studentData?.grade ?? 1,
        class: studentData?.class || '',
        number: studentData?.number || '',
        role: CONFIG.ROLES.STUDENT,
        email: email,
      };
    }

    // 3. どちらにも該当しない場合
    throw new Error('このアカウントはシステムに登録されていません。管理者に連絡してください。');
  },

  /**
   * メールアドレスからロール・名前・担当クラスを判定
   *
   * @param {string} email - メールアドレス
   * @returns {{ role: string, name: string|null, assignedClasses: string|null } | null}
   */
  determineUserRole(email) {
    try {
      const teacherData = SheetAccess.getTeacherData();
      if (teacherData.length <= 1) {
        return { role: CONFIG.ROLES.STUDENT, name: null, assignedClasses: null };
      }

      const headerMap = SheetUtils.buildHeaderMap(teacherData[0]);
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      if (emailIdx === -1) {
        return { role: CONFIG.ROLES.STUDENT, name: null, assignedClasses: null };
      }

      const normalizedEmail = String(email).trim().toLowerCase();

      for (let i = 1; i < teacherData.length; i++) {
        const row = teacherData[i];
        if (String(row[emailIdx]).trim().toLowerCase() === normalizedEmail) {
          // ロール判定
          const roleIdx = SheetUtils.getColIndex(headerMap, '権限');
          let role = CONFIG.ROLES.TEACHER;
          if (roleIdx !== -1) {
            const roleValue = row[roleIdx];
            if (roleValue === '管理者') role = CONFIG.ROLES.ADMIN;
            else if (roleValue === '教務') role = CONFIG.ROLES.KYOMU;
          }

          // 名前取得（「氏名」「名前」両方に対応）
          let nameIdx = SheetUtils.getColIndex(headerMap, '氏名');
          if (nameIdx === -1) nameIdx = SheetUtils.getColIndex(headerMap, '名前');

          // 担当クラス
          const classIdx = SheetUtils.getColIndex(headerMap, 'クラス');

          return {
            role: role,
            name: nameIdx !== -1 ? row[nameIdx] : null,
            assignedClasses: classIdx !== -1 ? row[classIdx] : null,
          };
        }
      }

      // 教職員データに未登録 → 生徒
      return { role: CONFIG.ROLES.STUDENT, name: null, assignedClasses: null };

    } catch (error) {
      Logger_.log(`権限判定エラー: ${error.toString()}`, 'ERROR');
      return null;
    }
  },

  // --- 権限ガード関数 ---

  /**
   * 教員以上の権限を要求（不足時はthrow）
   * @param {string} context - エラーメッセージに含めるコンテキスト
   * @returns {Object} userInfo
   */
  requireTeacherOrAdmin(context) {
    const userInfo = this.getUserInfo();
    if (!userInfo.success || !this.isTeacherOrAbove(userInfo.user.role)) {
      throw new Error(`${context}には教員権限が必要です`);
    }
    return userInfo;
  },

  /**
   * 管理者権限を要求（不足時はthrow）
   * @param {string} context - エラーメッセージに含めるコンテキスト
   * @returns {Object} userInfo
   */
  requireAdmin(context) {
    const userInfo = this.getUserInfo();
    if (!userInfo.success || !this.isAdmin(userInfo.user.role)) {
      throw new Error(`権限エラー: ${context}には管理者権限が必要です`);
    }
    return userInfo;
  },

  /**
   * 本人または教員以上の権限を要求
   * @param {string} requestedStudentId - リクエストされた学籍番号
   * @param {string} context - エラーメッセージに含めるコンテキスト
   * @returns {Object} userInfo
   */
  requireSelfOrTeacher(requestedStudentId, context) {
    const userInfo = this.getUserInfo();
    if (!userInfo.success) {
      throw new Error('ユーザー認証に失敗しました');
    }
    const user = userInfo.user;
    if (this.isTeacherOrAbove(user.role)) return userInfo;
    if (user.role === CONFIG.ROLES.STUDENT && String(user.studentId) === String(requestedStudentId)) {
      return userInfo;
    }
    throw new Error(`${context}への権限がありません`);
  },

  // --- クラス情報表示制御 ---

  /**
   * クラス情報を非表示にすべきかを判定
   * 設定「生徒への表示」が「学籍番号」の場合、クラス情報を返さない
   * @returns {boolean}
   */
  shouldHideClass() {
    const settings = SettingsService.getSettings();
    return settings[CONFIG.SETTING_KEYS.DISPLAY_MODE] === '学籍番号';
  },

  /**
   * {header, row} 形式のデータから組・番号をクリア
   * @param {Object} data - {header: [...], row: [...]}
   * @returns {Object}
   */
  stripClassInfoFromHeaderRow(data) {
    if (!data || !data.header || !data.row) return data;
    const classIdx = data.header.indexOf('組');
    const numberIdx = data.header.indexOf('番号');
    const newRow = [...data.row];
    if (classIdx !== -1) newRow[classIdx] = '';
    if (numberIdx !== -1) newRow[numberIdx] = '';
    return { ...data, row: newRow };
  },

  /**
   * 生徒データオブジェクトからclass/numberをクリア
   * @param {Object} data - 生徒データ
   * @returns {Object}
   */
  stripClassInfoFromStudentData(data) {
    if (!data) return data;
    return { ...data, class: '', number: '' };
  },

  // --- 内部ヘルパー（他サービスへの委譲） ---

  /**
   * メールアドレスから学籍番号を取得
   * @private
   * @param {string} email
   * @returns {string|null}
   */
  _getStudentIdFromEmail(email) {
    const submissionData = SheetAccess.getSubmissionData();
    if (submissionData.length <= 1) return null;

    const headerMap = SheetUtils.buildHeaderMap(submissionData[0]);
    const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
    const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
    if (emailIdx === -1 || studentIdIdx === -1) return null;

    const found = SheetUtils.findRow(submissionData, headerMap, 'メールアドレス', email);
    return found ? found.rowData[studentIdIdx] : null;
  },

  /**
   * 学籍番号から生徒基本データを取得（内部用）
   * @private
   * @param {string} studentId
   * @returns {{ name: string, grade: number, class: string, number: string } | null}
   */
  _getStudentDataByIdInternal(studentId) {
    const submissionData = SheetAccess.getSubmissionData();
    if (submissionData.length <= 1) return null;

    const headerMap = SheetUtils.buildHeaderMap(submissionData[0]);
    const found = SheetUtils.findRow(submissionData, headerMap, '学籍番号', studentId);
    if (!found) return null;

    const row = found.rowData;
    return {
      name: row[SheetUtils.getColIndex(headerMap, '名前')] || '',
      grade: Validator.safeParseGrade(row[SheetUtils.getColIndex(headerMap, '学年')]),
      class: row[SheetUtils.getColIndex(headerMap, '組')] || '',
      number: row[SheetUtils.getColIndex(headerMap, '番号')] || '',
    };
  },
};

// --- 後方互換ラッパー（既存フロントエンドからの呼び出し用）---
// Phase 3 でフロントエンドを書き換えたら削除可能

function getUserInfo() { return Auth.getUserInfo(); }
function isTeacherOrAbove(role) { return Auth.isTeacherOrAbove(role); }
function isKyomuOrAbove(role) { return Auth.isKyomuOrAbove(role); }
