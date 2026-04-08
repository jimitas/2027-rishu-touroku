/**
 * 06_StudentService.gs - 学生データサービス
 *
 * 学生データの取得・検索・ステータス管理を提供する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs, 03_Auth.gs
 */

const StudentService = {

  /**
   * 学籍番号から生徒データを取得（内部用・認可チェックなし）
   * 生徒情報シートから検索する
   * @param {string} studentId - 学籍番号
   * @returns {{ studentId: string, name: string, grade: number, class: string, number: number }|null}
   */
  getStudentDataByIdInternal(studentId) {
    try {
      const data = SheetAccess.getRosterData();
      if (!data || data.length < 2) return null;

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) return null;

      const row = found.rowData;
      return {
        studentId: studentId,
        name: row[SheetUtils.getColIndex(headerMap, '氏名')] || '',
        grade: Validator.safeParseGrade(row[SheetUtils.getColIndex(headerMap, '学年')], 1),
        class: row[SheetUtils.getColIndex(headerMap, '組')] != null
          ? String(row[SheetUtils.getColIndex(headerMap, '組')]).trim() : '',
        number: parseInt(row[SheetUtils.getColIndex(headerMap, '番号')]) || 1
      };
    } catch (error) {
      Logger_.log(`生徒データ取得エラー: ${error.toString()}`, 'ERROR');
      return null;
    }
  },

  /**
   * 学籍番号から生徒データを取得（公開用・認可チェック付き）
   * @param {string} studentId - 学籍番号
   * @returns {{ studentId: string, name: string, grade: number, class: string, number: number }|null}
   */
  getStudentDataById(studentId) {
    Auth.requireSelfOrTeacher(studentId, '生徒データ取得');
    const result = this.getStudentDataByIdInternal(studentId);
    if (Auth.shouldHideClass()) {
      return Auth.stripClassInfoFromStudentData(result);
    }
    return result;
  },

  /**
   * バッチ処理用の生徒データルックアップMapを構築
   * @returns {Map<string, Object>} 学籍番号→生徒データのMap
   */
  buildStudentLookupMap() {
    const studentMap = new Map();

    try {
      const data = SheetAccess.getRosterData();
      if (!data || data.length < 2) return studentMap;

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const nameIdx = SheetUtils.getColIndex(headerMap, '氏名');
      const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
      const classIdx = SheetUtils.getColIndex(headerMap, '組');
      const numberIdx = SheetUtils.getColIndex(headerMap, '番号');

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const sid = row[studentIdIdx];
        if (sid) {
          studentMap.set(String(sid), {
            studentId: String(sid),
            name: nameIdx !== -1 ? row[nameIdx] : '',
            grade: Validator.safeParseGrade(gradeIdx !== -1 ? row[gradeIdx] : null, 1),
            class: classIdx !== -1 && row[classIdx] != null ? String(row[classIdx]).trim() : '',
            number: numberIdx !== -1 ? (parseInt(row[numberIdx]) || 1) : 1
          });
        }
      }

      return studentMap;
    } catch (error) {
      Logger_.log(`生徒Mapビルドエラー: ${error.toString()}`, 'ERROR');
      return studentMap;
    }
  },

  /**
   * 学生基本情報を生徒情報シートから取得
   * @param {string} studentId - 学籍番号
   * @returns {{ 学年: *, 組: *, 番号: *, 名前: *, メールアドレス: string }|null}
   */
  getStudentBasicInfo(studentId) {
    try {
      const data = SheetAccess.getRosterData();
      if (!data || data.length < 2) return null;

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      if (!headerMap.has('学籍番号')) {
        Logger_.log('生徒情報シートに学籍番号列が見つかりません', 'WARN');
        return null;
      }

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        Logger_.log(`学籍番号${studentId}の学生情報が見つかりません`, 'WARN');
        return null;
      }

      const row = found.rowData;
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');

      return {
        学年: row[SheetUtils.getColIndex(headerMap, '学年')] || '',
        組: row[SheetUtils.getColIndex(headerMap, '組')] || '',
        番号: row[SheetUtils.getColIndex(headerMap, '番号')] || '',
        名前: row[SheetUtils.getColIndex(headerMap, '名前')] || '',
        メールアドレス: emailIdx !== -1 ? (row[emailIdx] || '') : ''
      };

    } catch (error) {
      Logger_.log(`学生基本情報取得エラー: ${error.toString()}`, 'ERROR');
      return null;
    }
  },

  /**
   * メールアドレスから学籍番号を取得（TextFinder優先、フォールバック付き）
   * @param {string} email - メールアドレス
   * @returns {string|null} 学籍番号
   */
  getStudentIdFromEmail(email) {
    try {
      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        Logger_.log('登録データシートが見つかりません', 'ERROR');
        return null;
      }

      // 高速検索：TextFinderを使用
      const finder = submissionSheet.createTextFinder(email);
      const foundCell = finder.findNext();

      if (foundCell) {
        const foundRow = foundCell.getRow();
        const cachedData = SheetAccess.getSubmissionData();
        const headers = cachedData[0];
        const headerMap = SheetUtils.buildHeaderMap(headers);
        const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

        if (studentIdIdx === -1) {
          Logger_.log('登録データシートに学籍番号列が見つかりません', 'ERROR');
          return this.getSubmissionLinearSearchFallback(email);
        }

        const studentId = submissionSheet.getRange(foundRow, studentIdIdx + 1).getValue();
        if (studentId) return studentId;
      }

      return this.getSubmissionLinearSearchFallback(email);

    } catch (error) {
      Logger_.log(`getStudentIdFromEmail エラー: ${error.toString()}`, 'ERROR');
      return this.getSubmissionLinearSearchFallback(email);
    }
  },

  /**
   * 登録データシート線形検索フォールバック
   * @param {string} email - メールアドレス
   * @returns {string|null} 学籍番号
   */
  getSubmissionLinearSearchFallback(email) {
    try {
      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) return null;

      const headerMap = SheetUtils.buildHeaderMap(data[0]);
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

      if (emailIdx === -1 || studentIdIdx === -1) {
        Logger_.log('登録データシートにメールアドレス列または学籍番号列が見つかりません', 'ERROR');
        return null;
      }

      for (let i = 1; i < data.length; i++) {
        if (data[i][emailIdx] === email && data[i][studentIdIdx]) {
          return data[i][studentIdIdx];
        }
      }

      return null;

    } catch (error) {
      Logger_.log(`登録データシート線形検索フォールバックエラー: ${error.toString()}`, 'ERROR');
      return null;
    }
  },

  /**
   * 提出状態を確認（内部用・認可チェックなし）
   * @param {string} studentId - 学籍番号
   * @returns {{ isSubmitted: boolean, status: string|null, timestamp: string|null }}
   */
  checkSubmissionStatusInternal(studentId) {
    try {
      const submissionData = SheetAccess.getSubmissionData();
      if (!submissionData || submissionData.length < 2) {
        return { isSubmitted: false, status: null, timestamp: null };
      }

      const header = submissionData[0];
      const headerMap = SheetUtils.buildHeaderMap(header);
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');

      if (studentIdIdx === -1 || statusIdx === -1) {
        return { isSubmitted: false, status: null, timestamp: null };
      }

      const dataRows = submissionData.slice(1);

      // 最新の提出データを検索
      const submittedRows = dataRows
        .filter(row => row[studentIdIdx] === studentId &&
          (row[statusIdx] === CONFIG.STATUSES.PROVISIONAL || row[statusIdx] === '一時保存'))
        .sort((a, b) => new Date(b[0]) - new Date(a[0]));

      if (submittedRows.length > 0) {
        const latestRow = submittedRows[0];
        const timestamp = latestRow[0];
        return {
          isSubmitted: latestRow[statusIdx] === CONFIG.STATUSES.PROVISIONAL,
          status: latestRow[statusIdx],
          timestamp: timestamp instanceof Date
            ? Utilities.formatDate(timestamp, CONFIG.DEFAULTS.TIMEZONE, CONFIG.DEFAULTS.DATE_FORMAT)
            : String(timestamp)
        };
      }

      return { isSubmitted: false, status: null, timestamp: null };

    } catch (e) {
      Logger_.log(`提出状態チェックエラー: ${e.toString()}`, 'ERROR');
      return { isSubmitted: false, status: null, timestamp: null };
    }
  },

  /**
   * 提出状態を確認（公開用・認可チェック付き）
   * @param {string} studentId - 学籍番号
   * @returns {{ isSubmitted: boolean, status: string|null, timestamp: string|null }}
   */
  checkSubmissionStatus(studentId) {
    Auth.requireSelfOrTeacher(studentId, '提出状態確認');
    return this.checkSubmissionStatusInternal(studentId);
  },

  /**
   * 履修済み科目データを取得
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, completedCourses: Array, retakenCourses: Array }}
   */
  getCompletedCourses(studentId) {
    try {
      Auth.requireSelfOrTeacher(studentId, '履修済み科目取得');

      // 履修済み科目シートは未実装のためデフォルトを返す
      return { success: true, completedCourses: [], retakenCourses: [] };

    } catch (error) {
      return Logger_.handleError(error, '履修済み科目データ取得');
    }
  },

  /**
   * 再履修科目の学年マッピングを生成
   * @param {Array} retakenCourses - 再履修科目リスト
   * @returns {Object} 再履修科目の学年マッピング
   */
  generateRetakenYearMap(retakenCourses) {
    const retakenYearMap = {};

    try {
      const courseData = CourseService.getCourseData();
      if (!courseData.success) return retakenYearMap;

      retakenCourses.forEach(courseName => {
        Object.keys(courseData.data).forEach(gradeKey => {
          const gradeData = courseData.data[gradeKey];
          const gradeNum = parseInt(gradeKey.replace('year', ''));

          ['required', 'elective'].forEach(category => {
            if (gradeData[category]) {
              const course = gradeData[category].find(c => c['科目名'] === courseName);
              if (course) {
                retakenYearMap[courseName] = gradeNum;
              }
            }
          });
        });
      });

      return retakenYearMap;

    } catch (error) {
      Logger_.log(`再履修学年マップ生成エラー: ${error.toString()}`, 'ERROR');
      return retakenYearMap;
    }
  },

  /**
   * 全生徒の履修データを取得（内部用）
   * @param {boolean} includeTeachers - 教職員データも含めるか
   * @returns {{ success: boolean, students: Array, timestamp?: string }}
   */
  getAllStudentsDataInternal(includeTeachers) {
    try {
      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) {
        return { success: true, students: [] };
      }

      const headers = data[0];
      if (!headers || headers.length === 0) {
        return { success: true, students: [] };
      }

      const headerMap = SheetUtils.buildHeaderMap(headers);
      const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

      // 除外すべき列のSet
      const excludeColumns = new Set(['タイムスタンプ', '更新日時', '提出日時', '教職員チェック', '教務チェック']);

      // 教職員メールSetを構築
      const teacherEmails = new Set();
      if (!includeTeachers) {
        const teacherData = SheetAccess.getTeacherData();
        if (teacherData.length > 1) {
          const tHeaderMap = SheetUtils.buildHeaderMap(teacherData[0]);
          const tEmailIdx = SheetUtils.getColIndex(tHeaderMap, 'メールアドレス');
          if (tEmailIdx !== -1) {
            for (let t = 1; t < teacherData.length; t++) {
              const te = String(teacherData[t][tEmailIdx] || '').trim().toLowerCase();
              if (te) teacherEmails.add(te);
            }
          }
        }
      }

      const students = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        // 学籍番号が空白の行をスキップ
        if (studentIdIdx >= 0 && !row[studentIdIdx]) continue;
        // 教職員データを除外
        if (!includeTeachers && emailIdx >= 0) {
          const rowEmail = String(row[emailIdx] || '').trim().toLowerCase();
          if (teacherEmails.has(rowEmail)) continue;
        }

        const studentData = {};
        headers.forEach((header, index) => {
          if (!excludeColumns.has(header)) {
            studentData[header] = row[index];
          }
        });

        // 履修データをパース
        try {
          if (studentData['履修データ']) {
            studentData['履修データ'] = JSON.parse(studentData['履修データ']);
          }
        } catch (parseError) {
          studentData['履修データ'] = {};
        }

        students.push(studentData);
      }

      return Logger_.successResponse({ students: students });

    } catch (error) {
      return Logger_.handleError(error, '全生徒データ取得');
    }
  },

  /**
   * 全生徒の履修データを取得（公開用・認可チェック付き）
   * @param {boolean} includeTeachers - 教職員データも含めるか
   * @returns {{ success: boolean, students: Array }}
   */
  getAllStudentsData(includeTeachers) {
    Auth.requireTeacherOrAdmin('全生徒データ取得');
    return this.getAllStudentsDataInternal(includeTeachers);
  },

  /**
   * 軽量ステータス取得（バッジ自動更新用）
   * @returns {{ success: boolean, statuses: Object }}
   */
  getStudentStatuses() {
    Auth.requireTeacherOrAdmin('ステータス取得');
    try {
      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) return { success: true, statuses: {} };

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const idIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
      if (idIdx === -1 || statusIdx === -1) return { success: true, statuses: {} };

      const statuses = {};
      for (let i = 1; i < data.length; i++) {
        const id = data[i][idIdx];
        if (id) statuses[String(id)] = data[i][statusIdx] || '';
      }
      return { success: true, statuses: statuses };
    } catch (error) {
      return Logger_.handleError(error, 'ステータス取得');
    }
  },

  /**
   * 個別生徒の認証コードを取得（教員/admin専用）
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, authCode?: string, error?: string }}
   */
  getAuthCode(studentId) {
    try {
      Auth.requireTeacherOrAdmin('認証コード取得');

      if (!studentId) {
        throw new Error('学籍番号が指定されていません');
      }

      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) {
        throw new Error('登録データが見つかりません');
      }

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      SheetUtils.requireColumns(headerMap, ['認証コード', '学籍番号']);

      const authCodeIdx = SheetUtils.getColIndex(headerMap, '認証コード');

      const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);
      if (!found) {
        throw new Error('該当する生徒が見つかりません');
      }

      return { success: true, authCode: found.rowData[authCodeIdx] || '' };

    } catch (error) {
      return Logger_.handleError(error, '認証コード取得');
    }
  },

  /**
   * クラス一覧（学年・組のユニークペア）を取得
   * @returns {{ success: boolean, classList: Array, timestamp?: string }}
   */
  getClassList() {
    try {
      Auth.requireTeacherOrAdmin('クラス一覧取得');

      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) {
        return { success: true, classList: [] };
      }

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
      const classIdx = SheetUtils.getColIndex(headerMap, '組');

      if (gradeIdx === -1 || classIdx === -1) {
        return { success: true, classList: [] };
      }

      const classSet = new Set();
      for (let i = 1; i < data.length; i++) {
        const grade = parseInt(data[i][gradeIdx]);
        const classVal = data[i][classIdx] != null ? String(data[i][classIdx]).trim() : '';
        if (!isNaN(grade) && classVal !== '') {
          classSet.add(grade + '-' + classVal);
        }
      }

      return Logger_.successResponse({ classList: Array.from(classSet) });

    } catch (error) {
      return Logger_.handleError(error, 'クラス一覧取得');
    }
  },

  /**
   * 提出済み生徒データを取得（教員用）
   * @returns {{ success: boolean, students: Array, timestamp?: string }}
   */
  getSubmittedStudents() {
    try {
      const userInfo = Auth.getUserInfo();
      if (!userInfo.success || !Auth.isTeacherOrAbove(userInfo.user.role)) {
        throw new Error('教員権限が必要です');
      }

      const data = SheetAccess.getSubmissionData();
      if (!data || data.length < 2) {
        return { success: true, students: [] };
      }

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');

      // バッチ処理用Mapを事前構築
      const studentMap = this.buildStudentLookupMap();

      const submittedStudents = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const sid = row[studentIdIdx];
        const status = row[statusIdx];

        if (status === CONFIG.STATUSES.PROVISIONAL) {
          const studentData = studentMap.get(String(sid));
          if (studentData) {
            submittedStudents.push({
              studentId: sid,
              name: studentData.name,
              grade: studentData.grade,
              class: studentData.class,
              number: studentData.number,
              status: status
            });
          }
        }
      }

      return Logger_.successResponse({ students: submittedStudents });

    } catch (error) {
      return Logger_.handleError(error, '提出済み生徒データ取得');
    }
  },

  /**
   * 教員用生徒データ取得（学籍番号ベース）
   * @param {string} studentId - 学籍番号
   * @returns {{ success: boolean, student?: Object, registrationData?: Object, completedCourses?: Array, retakenCourses?: Array, error?: string }}
   */
  getStudentDataForView(studentId) {
    try {
      Auth.requireTeacherOrAdmin('教員用生徒データ取得');

      const studentData = this.getStudentDataByIdInternal(studentId);
      if (!studentData) {
        throw new Error('生徒データが見つかりません');
      }

      const registrationData = RegistrationService.loadRegistrationDataInternal(studentId);

      const completedCourses = [];
      const retakenCourses = [];

      if (registrationData.success && registrationData.data) {
        Object.keys(registrationData.data).forEach(courseName => {
          const mark = registrationData.data[courseName];
          if (mark === '●' || /^●[123]$/.test(mark)) {
            completedCourses.push(courseName);
          } else if (mark === '2' || mark === '3' || mark === '○2' || mark === '○3' || mark === '○2N' || mark === '○3N') {
            retakenCourses.push(courseName);
          }
        });
      }

      return Logger_.successResponse({
        student: studentData,
        registrationData: registrationData.data,
        completedCourses: completedCourses,
        retakenCourses: retakenCourses
      });

    } catch (error) {
      return Logger_.handleError(error, '教員用生徒データ取得');
    }
  },

  /**
   * 教科書データ生成用：生徒情報シートから学籍番号で学生情報を取得
   * @param {string} studentId - 学籍番号
   * @returns {{ header: Array, row: Array }|null}
   */
  getStudentInfoForTextbook(studentId) {
    try {
      const data = SheetAccess.getRosterData();
      if (!data || data.length < 2) {
        Logger_.log('生徒情報シートにデータがありません', 'ERROR');
        return null;
      }

      const header = data[0];
      const headerMap = SheetUtils.buildHeaderMap(header);

      if (!headerMap.has('学籍番号')) {
        Logger_.log('生徒情報シートに学籍番号列が見つかりません', 'ERROR');
        return null;
      }

      // findRowはヘッダー行をスキップするが、元コードではfindで全行を検索（ヘッダー含む可能性あり）
      // 元コードの挙動に合わせ、ヘッダーを含む全行をチェック
      const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const studentRow = data.find(row => row[studentIdIdx] === studentId);
      if (studentRow) {
        return { header: header, row: studentRow };
      }

      return null;

    } catch (error) {
      Logger_.log(`教科書データ生成用学生情報取得エラー: ${error.toString()}`, 'ERROR');
      return null;
    }
  },
};

// --- 後方互換ラッパー ---
function getStudentDataByIdInternal(studentId) { return StudentService.getStudentDataByIdInternal(studentId); }
function getStudentDataById(studentId) { return StudentService.getStudentDataById(studentId); }
function buildStudentLookupMap() { return StudentService.buildStudentLookupMap(); }
function getStudentBasicInfo(studentId) { return StudentService.getStudentBasicInfo(studentId); }
function getStudentIdFromEmail(email) { return StudentService.getStudentIdFromEmail(email); }
function getSubmissionLinearSearchFallback(email) { return StudentService.getSubmissionLinearSearchFallback(email); }
function checkSubmissionStatusInternal(studentId) { return StudentService.checkSubmissionStatusInternal(studentId); }
function checkSubmissionStatus(studentId) { return StudentService.checkSubmissionStatus(studentId); }
function getCompletedCourses(studentId) { return StudentService.getCompletedCourses(studentId); }
function generateRetakenYearMap(retakenCourses) { return StudentService.generateRetakenYearMap(retakenCourses); }
function getAllStudentsDataInternal(includeTeachers) { return StudentService.getAllStudentsDataInternal(includeTeachers); }
function getAllStudentsData(includeTeachers) { return StudentService.getAllStudentsData(includeTeachers); }
function getStudentStatuses() { return StudentService.getStudentStatuses(); }
function getAuthCode(studentId) { return StudentService.getAuthCode(studentId); }
function getClassList() { return StudentService.getClassList(); }
function getSubmittedStudents() { return StudentService.getSubmittedStudents(); }
function getStudentDataForView(studentId) { return StudentService.getStudentDataForView(studentId); }
function getStudentInfoForTextbook(studentId) { return StudentService.getStudentInfoForTextbook(studentId); }
