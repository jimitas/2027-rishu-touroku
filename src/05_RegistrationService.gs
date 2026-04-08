/**
 * 05_RegistrationService.gs - 履修登録サービス
 *
 * 履修データの保存・仮登録・承認・差戻し・本登録など、
 * 登録ワークフロー全体を管理する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs, 03_Auth.gs,
 *        06_StudentService.gs, 08_SettingsService.gs
 */

const RegistrationService = {

  /**
   * 履修データを一時保存
   * @param {Object} registrationData - 履修データ
   * @returns {{ success: boolean, message?: string, timestamp?: string, error?: string }}
   */
  saveRegistrationData(registrationData) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success) {
        throw new Error('ユーザー認証に失敗');
      }
      const restriction = SettingsService.checkRestriction(userInfo.user.role, CONFIG.SETTING_KEYS.TEMP_SAVE);
      if (!restriction.allowed) {
        throw new Error(restriction.reason);
      }

      const studentId = userInfo.user.studentId;
      // ステータスは常に「保存」に強制（クライアント指定を無視し、ステータス注入を防止）
      this.saveToSubmissionSheet(studentId, registrationData, CONFIG.STATUSES.DRAFT, userInfo.user.email);

      Logger_.log(`履修データ保存完了: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '履修データを保存しました');

    } catch (error) {
      return Logger_.handleError(error, '履修データ保存');
    }
  },

  /**
   * 履修データを仮登録
   * @param {Object} registrationData - 履修データ
   * @returns {{ success: boolean, message?: string, timestamp?: string, error?: string }}
   */
  submitRegistrationData(registrationData) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success) {
        throw new Error('ユーザー認証に失敗');
      }
      const restriction = SettingsService.checkRestriction(userInfo.user.role, CONFIG.SETTING_KEYS.PROVISIONAL_REG);
      if (!restriction.allowed) {
        throw new Error(restriction.reason);
      }

      const studentId = userInfo.user.studentId;

      // 次年度の学年を算出
      const currentGrade = Validator.safeParseGrade(userInfo.user.grade, 1);
      const registrationYear = this.getNextGradeForStudent(studentId, currentGrade);

      // データ整合性チェック
      const validation = this.validateRegistrationData(registrationData, registrationYear);
      if (!validation.valid) {
        throw new Error(`データ整合性エラー: ${validation.errors.join(', ')}`);
      }

      this.saveToSubmissionSheet(studentId, registrationData, CONFIG.STATUSES.PROVISIONAL, userInfo.user.email);

      Logger_.log(`履修データ仮登録完了: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '履修データを仮登録しました');

    } catch (error) {
      return Logger_.handleError(error, '履修データ仮登録');
    }
  },

  /**
   * 教務が生徒の代わりに履修データを仮登録
   * @param {string} studentId - 学籍番号
   * @param {Object} registrationData - 履修データ
   * @returns {{ success: boolean, message?: string, timestamp?: string, error?: string }}
   */
  adminSubmitRegistration(studentId, registrationData) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success || !Auth.isTeacherOrAbove(userInfo.user.role)) {
        throw new Error('教員権限が必要です');
      }
      const restriction = SettingsService.checkRestriction(userInfo.user.role, null);
      if (!restriction.allowed) {
        throw new Error(restriction.reason);
      }

      // 生徒の学年から次年度を算出
      const studentInfo = StudentService.getStudentBasicInfo(studentId);
      const adminCurrentGrade = Validator.safeParseGrade(studentInfo?.学年, 1);
      const registrationYear = this.getNextGradeForStudent(studentId, adminCurrentGrade);

      // データ整合性チェック
      const validation = this.validateRegistrationData(registrationData, registrationYear);
      if (!validation.valid) {
        throw new Error(`データ整合性エラー: ${validation.errors.join(', ')}`);
      }

      this.adminSaveToSubmissionSheet(studentId, registrationData, CONFIG.STATUSES.PROVISIONAL);

      Logger_.log(`教職員による仮登録完了: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '履修データを仮登録しました');

    } catch (error) {
      return Logger_.handleError(error, '教職員による仮登録');
    }
  },

  /**
   * 管理者による登録データの直接編集保存
   * ステータスは変更せず、マークのみ更新する
   * @param {string} studentId - 学籍番号
   * @param {Object} registrationData - 履修データ
   * @param {string|null} newStatus - 新ステータス（nullなら現状維持）
   * @param {Object|null} basicFields - 基本情報フィールド
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  adminUpdateRegistrationMarks(studentId, registrationData, newStatus, basicFields) {
    try {
      Auth.requireAdmin('登録データ管理');

      // 現在のステータスを取得して保持
      const allData = SheetAccess.getSubmissionData();
      const headers = allData.length > 0 ? allData[0] : [];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const idIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

      let currentStatus = CONFIG.STATUSES.DRAFT;
      const found = SheetUtils.findRow(allData, headerMap, '学籍番号', studentId);
      if (found) {
        currentStatus = found.rowData[statusIdx] || CONFIG.STATUSES.DRAFT;
      }

      // newStatusが指定されていればそれを使用、なければ現在のステータスを保持
      const statusToSave = (newStatus != null) ? newStatus : currentStatus;

      // 基本情報の更新（科目マークとは別にシートに直接書き込み）
      if (basicFields && typeof basicFields === 'object' && found) {
        const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
        const targetRow = found.rowIndex;

        const allowedFields = ['学籍番号', '学年', '組', '番号', 'メールアドレス', '来年度学年', '認証コード'];
        allowedFields.forEach(field => {
          if (field in basicFields) {
            const colIdx = SheetUtils.getColIndex(headerMap, field);
            if (colIdx !== -1) {
              submissionSheet.getRange(targetRow, colIdx + 1).setValue(basicFields[field]);
            }
          }
        });
      }

      // バリデーションはスキップ（管理者による自由編集）
      this.adminSaveToSubmissionSheet(studentId, registrationData, statusToSave);

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log(`管理者による登録データ直接編集: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '登録データを更新しました');
    } catch (error) {
      return Logger_.handleError(error, '登録データ管理');
    }
  },

  /**
   * 履修データの検証
   * @param {Object} registrationData - 履修データ
   * @param {number} registrationYear - 登録対象の学年
   * @returns {{ valid: boolean, errors: Array<string> }}
   */
  validateRegistrationData(registrationData, registrationYear) {
    const errors = [];

    try {
      if (!registrationData || typeof registrationData !== 'object') {
        errors.push('履修データが無効です');
        return { valid: false, errors: errors };
      }

      // 検証1: 選択科目の値チェック
      const VALID_MARKS = new Set(['○', '〇', '●', '○1', '○2', '○3', '〇1', '〇2', '〇3', '●1', '●2', '●3', '○1N', '○2N', '○3N', '〇1N', '〇2N', '〇3N']);
      const METADATA_KEYS = new Set(CONFIG.BASIC_COLUMNS.concat(['提出日時', '更新日時']));

      Object.entries(registrationData).forEach(([key, value]) => {
        if (METADATA_KEYS.has(key)) return;
        if (value === '' || value === null || value === undefined) return;
        if (!VALID_MARKS.has(value)) {
          errors.push('不正な値: ' + key + '=' + value);
        }
      });

      // 検証2: 単位数上限チェック
      const totalUnits = this.calculateTotalUnits(registrationData, registrationYear);
      if (totalUnits > CONFIG.UNIT_LIMIT) {
        errors.push('単位数上限超過: ' + totalUnits + '単位（上限' + CONFIG.UNIT_LIMIT + '単位）');
      }

      return { valid: errors.length === 0, errors: errors };

    } catch (error) {
      return { valid: false, errors: ['検証処理エラー: ' + error.toString()] };
    }
  },

  /**
   * 登録単位数の計算（次年度の登録科目のみカウント）
   * @param {Object} registrationData - 履修データ
   * @param {number} registrationYear - 登録対象の学年
   * @returns {number} 総単位数
   */
  calculateTotalUnits(registrationData, registrationYear) {
    let totalUnits = 0;

    try {
      const courseData = CourseService.getCourseData();
      if (!courseData.success) {
        throw new Error('科目データの取得に失敗');
      }

      const targetMarks = new Set(['○', '〇']);
      if (registrationYear) {
        targetMarks.add('○' + registrationYear);
        targetMarks.add('〇' + registrationYear);
        targetMarks.add('○' + registrationYear + 'N');
        targetMarks.add('〇' + registrationYear + 'N');
      }

      Object.entries(registrationData).forEach(([courseName, value]) => {
        if (!targetMarks.has(value)) return;

        Object.values(courseData.data).forEach(gradeData => {
          ['required', 'elective'].forEach(category => {
            if (gradeData[category]) {
              const course = gradeData[category].find(c => c['科目名'] === courseName);
              if (course) {
                totalUnits += parseInt(course['単位数']) || 0;
              }
            }
          });
        });
      });

      return totalUnits;

    } catch (error) {
      Logger_.log(`単位数計算エラー: ${error.toString()}`, 'ERROR');
      return 0;
    }
  },

  /**
   * 登録データシートの「来年度学年」列から次年度学年を取得
   * @param {string} studentId - 学籍番号
   * @param {number} currentGrade - 現在の学年
   * @returns {number} 次年度の学年
   */
  getNextGradeForStudent(studentId, currentGrade) {
    const data = SheetAccess.getSubmissionData();
    if (data.length < 2) return Math.min(currentGrade + 1, 3);

    const headerMap = SheetUtils.buildHeaderMap(data[0]);
    const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
    if (found) {
      const nextGradeIdx = SheetUtils.getColIndex(headerMap, '来年度学年');
      if (nextGradeIdx !== -1) {
        const val = parseInt(found.rowData[nextGradeIdx]);
        if (!isNaN(val) && val >= 1 && val <= 3) return val;
      }
    }
    return Math.min(currentGrade + 1, 3);
  },

  /**
   * 生徒用: 履修データをシートに保存（メールアドレスで行検索）
   * @param {string} studentId - 学籍番号
   * @param {Object} data - 履修データ
   * @param {string} status - ステータス
   * @param {string|null} email - メールアドレス
   * @returns {boolean} 保存成功
   */
  saveToSubmissionSheet(studentId, data, status, email) {
    return this._saveToSubmissionSheetInternal(studentId, data, status, email, false);
  },

  /**
   * 教務用: 履修データをシートに保存（学籍番号で行検索）
   * @param {string} studentId - 学籍番号
   * @param {Object} data - 履修データ
   * @param {string} status - ステータス
   * @returns {boolean} 保存成功
   */
  adminSaveToSubmissionSheet(studentId, data, status) {
    return this._saveToSubmissionSheetInternal(studentId, data, status, null, true);
  },

  /**
   * 履修データを登録データシートに保存する（内部実装）
   * @param {string} studentId - 学籍番号
   * @param {Object} data - 履修データ
   * @param {string} status - 提出状態
   * @param {string|null} email - メールアドレス
   * @param {boolean} useStudentIdSearch - trueなら学籍番号で行検索
   * @returns {boolean} 保存成功
   * @private
   */
  _saveToSubmissionSheetInternal(studentId, data, status, email, useStudentIdSearch) {
    try {
      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const allData = SheetAccess.getSubmissionData();
      const headers = allData.length > 0 ? allData[0] : submissionSheet.getRange(1, 1, 1, submissionSheet.getLastColumn()).getValues()[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      // 必須列チェック
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');

      if (studentIdIdx === -1 || statusIdx === -1) {
        throw new Error(`基本列が見つかりません: 学籍番号(${studentIdIdx}), ステータス(${statusIdx})`);
      }

      // 既存データの検索
      let targetRow = -1;

      if (useStudentIdSearch) {
        const found = SheetUtils.findRow(allData, headerMap, '学籍番号', studentId);
        if (found) targetRow = found.rowIndex;
      } else {
        const searchEmail = email || Session.getActiveUser().getEmail();
        if (emailIdx !== -1 && searchEmail) {
          const found = SheetUtils.findRow(allData, headerMap, 'メールアドレス', searchEmail);
          if (found) targetRow = found.rowIndex;
        }
      }

      // 行データの構築
      const lastCol = headers.length;
      let rowData;
      const isNewRow = (targetRow === -1);

      if (isNewRow) {
        targetRow = submissionSheet.getLastRow() + 1;
        rowData = new Array(lastCol).fill('');

        // 学生基本情報を設定
        const studentInfo = StudentService.getStudentBasicInfo(studentId);
        rowData[studentIdIdx] = studentId;

        if (studentInfo) {
          const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
          const classIdx = SheetUtils.getColIndex(headerMap, '組');
          const numberIdx = SheetUtils.getColIndex(headerMap, '番号');
          const nameIdx = SheetUtils.getColIndex(headerMap, '名前');

          if (gradeIdx !== -1) rowData[gradeIdx] = studentInfo.学年;
          if (classIdx !== -1) rowData[classIdx] = studentInfo.組;
          if (numberIdx !== -1) rowData[numberIdx] = studentInfo.番号;
          if (nameIdx !== -1) rowData[nameIdx] = studentInfo.名前;
          const eIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
          if (eIdx !== -1) rowData[eIdx] = studentInfo.メールアドレス || '';
        } else {
          Logger_.log(`学生基本情報が見つかりません: ${studentId}`, 'WARN');
        }
      } else {
        rowData = submissionSheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
      }

      // 科目列を特定（基本列以外）
      const courseColumns = headers.filter(header => header && !CONFIG.BASIC_COLUMNS.includes(header));

      // 科目データをすべてクリア
      courseColumns.forEach(courseName => {
        const idx = SheetUtils.getColIndex(headerMap, courseName);
        if (idx !== -1) rowData[idx] = '';
      });

      // 新しい科目データを設定
      let savedSubjects = 0;
      Object.keys(data).forEach(courseName => {
        if (courseName === 'タイムスタンプ') return;

        const idx = SheetUtils.getColIndex(headerMap, courseName);
        if (idx !== -1) {
          rowData[idx] = SheetUtils.sanitizeCellValue(data[courseName]);
          savedSubjects++;
        } else {
          Logger_.log(`科目「${courseName}」の列が見つかりません`, 'WARN');
        }
      });

      // ステータスを設定
      rowData[statusIdx] = status;

      // タイムスタンプを設定
      if (timestampIdx !== -1) {
        rowData[timestampIdx] = Logger_.getJSTTimestamp();
      }

      // メールアドレスを設定
      if (emailIdx !== -1) {
        if (email) {
          rowData[emailIdx] = email;
        } else if (!useStudentIdSearch) {
          try {
            const userEmail = Session.getActiveUser().getEmail();
            if (userEmail && userEmail.trim()) {
              rowData[emailIdx] = userEmail;
            }
          } catch (emailError) {
            Logger_.log(`メールアドレス取得エラー: ${emailError}`, 'ERROR');
          }
        }
      }

      // 一括で書き込み
      submissionSheet.getRange(targetRow, 1, 1, lastCol).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      return true;

    } catch (error) {
      Logger_.log(`_saveToSubmissionSheetInternal エラー: ${error.toString()}`, 'ERROR');
      throw error;
    }
  },

  /**
   * 履修データを読み込み（内部用・認可チェックなし）
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, header?: Array, row?: Array, data?: Object, timestamp?: string, error?: string }}
   */
  loadRegistrationDataInternal(studentId) {
    try {
      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) {
        return { success: true, data: {} };
      }

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      if (timestampIdx === -1) {
        return { success: true, data: {} };
      }

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        return { success: true, data: {} };
      }

      // rowの中のDateオブジェクトを文字列に変換
      const serializedRow = found.rowData.map(cell =>
        cell instanceof Date ? cell.toISOString() : cell
      );

      const timestamp = found.rowData[timestampIdx];
      return {
        success: true,
        header: headers,
        row: serializedRow,
        timestamp: timestamp instanceof Date ? timestamp.toISOString() : timestamp
      };

    } catch (error) {
      Logger_.log('[データ読み込み] エラー発生: ' + error.toString(), 'ERROR');
      return Logger_.handleError(error, '履修データ読み込み');
    }
  },

  /**
   * 履修データを読み込み（公開用・認可チェック付き）
   * @param {string} studentId - 学籍番号
   * @returns {Object} 履修データ
   */
  loadRegistrationData(studentId) {
    Auth.requireSelfOrTeacher(studentId, '履修データ読み込み');
    const result = this.loadRegistrationDataInternal(studentId);
    if (Auth.shouldHideClass()) {
      return Auth.stripClassInfoFromHeaderRow(result);
    }
    return result;
  },

  /**
   * 履修登録を承認（教職員チェック）
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  approveRegistration(studentId) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success || !Auth.isTeacherOrAbove(userInfo.user.role)) {
        throw new Error('教員権限が必要です');
      }
      const restriction = SettingsService.checkRestriction(userInfo.user.role, null);
      if (!restriction.allowed) {
        throw new Error(restriction.reason);
      }

      return this.approveStudent(studentId);

    } catch (error) {
      return Logger_.handleError(error, '履修登録承認');
    }
  },

  /**
   * 生徒の承認処理
   * ステータスを仮登録→教職員チェック済みに更新、教職員チェック列にタイムスタンプ設定
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  approveStudent(studentId) {
    try {
      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['学籍番号', 'ステータス', 'タイムスタンプ', '教職員チェック']);

      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');
      const teacherCheckIdx = SheetUtils.getColIndex(headerMap, '教職員チェック');

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        throw new Error('指定された学籍番号の生徒が見つかりません');
      }

      const currentStatus = found.rowData[statusIdx];
      if (currentStatus !== CONFIG.STATUSES.PROVISIONAL) {
        throw new Error('仮登録状態の生徒のみ承認できます（現在のステータス: ' + currentStatus + '）');
      }

      const rowData = found.rowData.slice();
      rowData[statusIdx] = '確認済';
      rowData[timestampIdx] = Logger_.getJSTTimestamp();
      rowData[teacherCheckIdx] = Logger_.getJSTTimestamp();

      submissionSheet.getRange(found.rowIndex, 1, 1, rowData.length).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log(`履修登録承認完了: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '履修登録を承認しました');

    } catch (error) {
      return Logger_.handleError(error, '生徒承認');
    }
  },

  /**
   * 生徒の履修登録を差戻し
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  revertSubmission(studentId) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success || !Auth.isTeacherOrAbove(userInfo.user.role)) {
        throw new Error('教員権限が必要です');
      }
      const restriction = SettingsService.checkRestriction(userInfo.user.role, null);
      if (!restriction.allowed) {
        throw new Error(restriction.reason);
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['学籍番号', 'ステータス', '認証コード', 'タイムスタンプ']);

      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const authCodeIdx = SheetUtils.getColIndex(headerMap, '認証コード');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        throw new Error('指定された学籍番号の生徒が見つかりません');
      }

      const rowData = found.rowData.slice();
      rowData[statusIdx] = CONFIG.STATUSES.REVERTED;
      rowData[authCodeIdx] = '';
      rowData[timestampIdx] = Logger_.getJSTTimestamp();
      submissionSheet.getRange(found.rowIndex, 1, 1, rowData.length).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log(`履修登録差戻し完了: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '履修登録を差戻しました');

    } catch (error) {
      return Logger_.handleError(error, '履修登録差戻し');
    }
  },

  /**
   * 確認済の生徒を仮登録に戻す
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  revertToProvisional(studentId) {
    try {
      Auth.requireTeacherOrAdmin('承認取消');

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['学籍番号', 'ステータス', 'タイムスタンプ']);

      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        throw new Error('指定された学籍番号の生徒が見つかりません');
      }

      const currentStatus = found.rowData[statusIdx];
      if (currentStatus !== '確認済') {
        throw new Error('この生徒は確認済ではありません');
      }

      const rowData = found.rowData.slice();
      rowData[statusIdx] = CONFIG.STATUSES.PROVISIONAL;
      rowData[timestampIdx] = Logger_.getJSTTimestamp();
      submissionSheet.getRange(found.rowIndex, 1, 1, rowData.length).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log(`承認取消完了: 学籍番号 ${studentId} を仮登録に戻しました`);

      return Logger_.successResponse({}, '承認を取り消して仮登録に戻しました');

    } catch (error) {
      Logger_.log(`承認取消エラー: ${error.message}`, 'ERROR');
      return { success: false, error: error.message };
    }
  },

  /**
   * 管理者が自分の利用停止ステータスを解除する
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  clearOwnSuspension() {
    try {
      const email = Session.getActiveUser().getEmail();
      if (!email) {
        throw new Error('メールアドレスが取得できません');
      }

      // 教職員データシートから直接ロール判定（getUserInfo経由しない）
      const roleResult = Auth.determineUserRole(email);
      if (!roleResult || roleResult.role !== CONFIG.ROLES.ADMIN) {
        throw new Error('管理者権限が必要です');
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['メールアドレス', 'ステータス', 'タイムスタンプ']);

      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      const normalizedEmail = email.trim().toLowerCase();

      for (let i = 1; i < data.length; i++) {
        const rowEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
        if (rowEmail === normalizedEmail) {
          if (data[i][statusIdx] !== CONFIG.STATUSES.SUSPENDED) {
            throw new Error('現在のステータスは利用停止ではありません');
          }

          const rowData = data[i].slice();
          rowData[statusIdx] = '';
          rowData[timestampIdx] = Logger_.getJSTTimestamp();
          submissionSheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);

          // キャッシュクリア
          SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

          Logger_.log(`管理者利用停止解除: ${email}`);

          return Logger_.successResponse({}, '利用停止を解除しました');
        }
      }

      throw new Error('登録データシートにアカウント情報が見つかりません');

    } catch (error) {
      return Logger_.handleError(error, '利用停止解除');
    }
  },

  /**
   * 生徒の本登録許可を設定
   * @param {string} studentId - 学籍番号
   * @param {string} authCode - 認証コード
   * @returns {{ success: boolean, message?: string, authCode?: string, error?: string }}
   */
  setFinalApproval(studentId, authCode) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success || !Auth.isTeacherOrAbove(userInfo.user.role)) {
        throw new Error('教員以上の権限が必要です');
      }
      const restriction = SettingsService.checkRestriction(userInfo.user.role, CONFIG.SETTING_KEYS.FINAL_REG_ALLOWED);
      if (!restriction.allowed) {
        throw new Error(restriction.reason);
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['学籍番号', '認証コード', 'ステータス', 'タイムスタンプ']);

      const authCodeIdx = SheetUtils.getColIndex(headerMap, '認証コード');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        throw new Error('指定された学籍番号の生徒が見つかりません');
      }

      if (found.rowData[authCodeIdx]) {
        throw new Error('既に本登録許可されています');
      }

      const rowData = found.rowData.slice();
      rowData[authCodeIdx] = authCode;
      rowData[statusIdx] = CONFIG.STATUSES.APPROVED;
      rowData[timestampIdx] = Logger_.getJSTTimestamp();
      submissionSheet.getRange(found.rowIndex, 1, 1, rowData.length).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log(`本登録許可設定完了: 学籍番号 ${studentId}, 認証コード ${authCode}`);

      return Logger_.successResponse({ authCode: authCode }, '本登録許可を設定しました');

    } catch (error) {
      return Logger_.handleError(error, '本登録許可設定');
    }
  },

  /**
   * 本登録許可を取り消す
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, message?: string, error?: string }}
   */
  cancelFinalApproval(studentId) {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success || !Auth.isTeacherOrAbove(userInfo.user.role)) {
        throw new Error('教員以上の権限が必要です');
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['学籍番号', '認証コード', 'ステータス', 'タイムスタンプ']);

      const authCodeIdx = SheetUtils.getColIndex(headerMap, '認証コード');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        throw new Error('指定された学籍番号の生徒が見つかりません');
      }

      const rowData = found.rowData.slice();
      rowData[authCodeIdx] = '';
      rowData[statusIdx] = CONFIG.STATUSES.PROVISIONAL;
      rowData[timestampIdx] = Logger_.getJSTTimestamp();
      submissionSheet.getRange(found.rowIndex, 1, 1, rowData.length).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log(`本登録許可取消完了: 学籍番号 ${studentId}`);

      return Logger_.successResponse({}, '本登録許可を取り消しました');

    } catch (error) {
      return Logger_.handleError(error, '本登録許可取消');
    }
  },

  /**
   * 認証コードを検証して学生データを取得
   * @param {string} authCode - 認証コード
   * @returns {{ success: boolean, studentId?: string, studentData?: Object, registrationData?: Object, status?: string, error?: string }}
   */
  validateAuthCode(authCode) {
    try {
      if (!authCode) {
        throw new Error('認証コードが指定されていません');
      }

      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) {
        throw new Error('登録データシートが見つかりません');
      }

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['認証コード', '学籍番号', 'ステータス']);

      const authCodeIdx = SheetUtils.getColIndex(headerMap, '認証コード');
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');

      // 認証コードで検索
      const found = SheetUtils.findRow(data, headerMap, '認証コード', authCode);
      if (!found) {
        throw new Error('無効な認証コードです');
      }

      const studentId = found.rowData[studentIdIdx];
      const status = found.rowData[statusIdx];

      // セッション検証
      Auth.requireSelfOrTeacher(studentId, '認証コード検証');

      // 学生データ取得
      let studentData = StudentService.getStudentDataByIdInternal(studentId);
      let registrationData = this.loadRegistrationDataInternal(studentId);

      // クラス非表示モード
      if (Auth.shouldHideClass()) {
        studentData = Auth.stripClassInfoFromStudentData(studentData);
        registrationData = Auth.stripClassInfoFromHeaderRow(registrationData);
      }

      Logger_.log(`認証コード検証成功: ${authCode}, 学籍番号: ${studentId}`);

      return {
        success: true,
        studentId: studentId,
        studentData: studentData,
        registrationData: registrationData,
        status: status
      };

    } catch (error) {
      return Logger_.handleError(error, '認証コード検証');
    }
  },

  /**
   * 本登録を実行
   * @param {string} authCode - 認証コード
   * @returns {{ success: boolean, message?: string, studentId?: string, error?: string }}
   */
  finalRegister(authCode) {
    try {
      if (!authCode) {
        throw new Error('認証コードが指定されていません');
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const data = SheetAccess.getSubmissionData();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['認証コード', '学籍番号', 'ステータス', 'タイムスタンプ']);

      const authCodeIdx = SheetUtils.getColIndex(headerMap, '認証コード');
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      const timestampIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');

      // 認証コードで検索
      const found = SheetUtils.findRow(data, headerMap, '認証コード', authCode);
      if (!found) {
        throw new Error('無効な認証コードです');
      }

      const studentId = found.rowData[studentIdIdx];

      // セッション検証
      Auth.requireSelfOrTeacher(studentId, '本登録実行');

      // メンテナンスモード・機能設定チェック
      const callerInfo = Auth.getUserInfo();
      if (callerInfo.success) {
        const restriction = SettingsService.checkRestriction(callerInfo.user.role, CONFIG.SETTING_KEYS.FINAL_REG_ALLOWED);
        if (!restriction.allowed) {
          throw new Error(restriction.reason);
        }
      }

      const rowData = found.rowData.slice();
      rowData[statusIdx] = CONFIG.STATUSES.FINAL;
      rowData[timestampIdx] = Logger_.getJSTTimestamp();
      rowData[authCodeIdx] = ''; // 認証コード削除（セキュリティ対策）

      submissionSheet.getRange(found.rowIndex, 1, 1, rowData.length).setValues([rowData]);

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log(`本登録完了: 学籍番号 ${studentId}, 認証コード ${authCode}`);

      return Logger_.successResponse({ studentId: studentId }, '本登録が完了しました');

    } catch (error) {
      return Logger_.handleError(error, '本登録実行');
    }
  },
};

// --- 後方互換ラッパー ---
function saveRegistrationData(registrationData) { return RegistrationService.saveRegistrationData(registrationData); }
function submitRegistrationData(registrationData) { return RegistrationService.submitRegistrationData(registrationData); }
function adminSubmitRegistration(studentId, registrationData) { return RegistrationService.adminSubmitRegistration(studentId, registrationData); }
function adminUpdateRegistrationMarks(studentId, registrationData, newStatus, basicFields) { return RegistrationService.adminUpdateRegistrationMarks(studentId, registrationData, newStatus, basicFields); }
function validateRegistrationData(registrationData, registrationYear) { return RegistrationService.validateRegistrationData(registrationData, registrationYear); }
function calculateTotalUnits(registrationData, registrationYear) { return RegistrationService.calculateTotalUnits(registrationData, registrationYear); }
function getNextGradeForStudent(studentId, currentGrade) { return RegistrationService.getNextGradeForStudent(studentId, currentGrade); }
function saveToSubmissionSheet(studentId, data, status, email) { return RegistrationService.saveToSubmissionSheet(studentId, data, status, email); }
function adminSaveToSubmissionSheet(studentId, data, status) { return RegistrationService.adminSaveToSubmissionSheet(studentId, data, status); }
function loadRegistrationDataInternal(studentId) { return RegistrationService.loadRegistrationDataInternal(studentId); }
function loadRegistrationData(studentId) { return RegistrationService.loadRegistrationData(studentId); }
function approveRegistration(studentId) { return RegistrationService.approveRegistration(studentId); }
function approveStudent(studentId) { return RegistrationService.approveStudent(studentId); }
function revertSubmission(studentId) { return RegistrationService.revertSubmission(studentId); }
function revertToProvisional(studentId) { return RegistrationService.revertToProvisional(studentId); }
function clearOwnSuspension() { return RegistrationService.clearOwnSuspension(); }
function setFinalApproval(studentId, authCode) { return RegistrationService.setFinalApproval(studentId, authCode); }
function cancelFinalApproval(studentId) { return RegistrationService.cancelFinalApproval(studentId); }
function validateAuthCode(authCode) { return RegistrationService.validateAuthCode(authCode); }
function finalRegister(authCode) { return RegistrationService.finalRegister(authCode); }
