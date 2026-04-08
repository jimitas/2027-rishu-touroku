/**
 * 04_CourseService.gs - 科目データ管理サービス
 *
 * 科目データの取得・パース・検索・CRUD操作を提供する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs, 03_Auth.gs
 */

const CourseService = {

  /**
   * 科目データを取得（認証チェック付き）
   * @returns {{ success: boolean, data?: Object, timestamp?: string, error?: string }}
   */
  getCourseData() {
    try {
      Auth.getUserInfo();

      const { valid, missing } = SheetAccess.validateSheets([
        CONFIG.SHEETS.COURSE, CONFIG.SHEETS.TEACHER, CONFIG.SHEETS.SUBMISSION
      ]);
      if (!valid) {
        throw new Error('必要なシートが見つかりません: ' + missing.join(', '));
      }

      const courseData = this.parseCourseSheet();

      return Logger_.successResponse({ data: courseData });

    } catch (error) {
      return Logger_.handleError(error, '科目データ取得');
    }
  },

  /**
   * 科目シートを解析し、学年・区分別にグルーピングする
   * @returns {Object} { year1: { required: [...], elective: [...] }, ... }
   */
  parseCourseSheet() {
    const data = SheetAccess.getCourseData();
    if (!data || data.length < 2) return {};

    const headers = data[0];
    const coursesByGrade = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const course = {};

      headers.forEach((header, index) => {
        if (header) course[header] = row[index];
      });

      const grade = `year${course['学年'] || 1}`;
      if (!coursesByGrade[grade]) {
        coursesByGrade[grade] = { required: [], elective: [] };
      }

      const category = course['区分'] === '必修' ? 'required' : 'elective';
      coursesByGrade[grade][category].push(course);
    }

    return coursesByGrade;
  },

  /**
   * 科目名から科目データを検索
   * @param {string} courseName - 科目名
   * @param {Object} courseData - parseCourseSheet() の戻り値
   * @returns {Object|null} 科目データ
   */
  findCourseByName(courseName, courseData) {
    for (const gradeKey of Object.keys(courseData)) {
      const gradeData = courseData[gradeKey];
      for (const category of ['required', 'elective']) {
        if (gradeData[category]) {
          const course = gradeData[category].find(c => c['科目名'] === courseName);
          if (course) return course;
        }
      }
    }
    return null;
  },

  /**
   * 科目一覧を取得（管理画面用）
   * @returns {{ success: boolean, courses?: Array, headers?: Array, error?: string }}
   */
  getCourseList() {
    Auth.requireAdmin('科目一覧取得');
    try {
      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      if (!sheet) {
        throw new Error('科目データシートが見つかりません');
      }
      const data = SheetAccess.getDataWithRetry(sheet, '科目一覧');
      if (data.length === 0) {
        return { success: true, courses: [], headers: [] };
      }
      const headers = data[0];
      const courses = SheetUtils.toObjects(data).filter(c => c['科目名']);
      return { success: true, courses: courses, headers: headers };
    } catch (error) {
      return Logger_.handleError(error, '科目一覧取得');
    }
  },

  /**
   * 科目を追加（科目データシート + 登録データシートに列追加）
   * @param {Object} courseData - 科目データ
   * @returns {{ success: boolean, error?: string }}
   */
  addCourse(courseData) {
    Auth.requireAdmin('科目追加');
    try {
      if (!courseData || !courseData['科目名']) {
        throw new Error('科目名は必須です');
      }
      const courseName = String(courseData['科目名']).trim();
      if (!courseName) {
        throw new Error('科目名は必須です');
      }

      const courseSheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      if (!courseSheet) {
        throw new Error('科目データシートが見つかりません');
      }
      const courseAllData = SheetAccess.getDataWithRetry(courseSheet, '科目追加');
      const headers = courseAllData[0];

      // 重複チェック
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const nameIdx = SheetUtils.getColIndex(headerMap, '科目名');
      if (nameIdx === -1) {
        throw new Error('科目データシートに科目名列が見つかりません');
      }
      for (let i = 1; i < courseAllData.length; i++) {
        if (String(courseAllData[i][nameIdx]).trim() === courseName) {
          return { success: false, error: '同名の科目が既に存在します: ' + courseName };
        }
      }

      // ヘッダー順に新しい行データを構築
      const newRow = headers.map(header =>
        courseData.hasOwnProperty(header) ? courseData[header] : ''
      );
      courseSheet.appendRow(newRow);

      // 登録データシートの末尾に列追加（ヘッダー = 科目名）
      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (submissionSheet) {
        const lastCol = submissionSheet.getLastColumn();
        submissionSheet.getRange(1, lastCol + 1).setValue(courseName);
      }

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('科目追加: ' + courseName);
      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '科目追加');
    }
  },

  /**
   * 科目を更新（科目名変更時は登録データシートのヘッダーもリネーム）
   * @param {string} originalName - 元の科目名
   * @param {Object} courseData - 更新後の科目データ
   * @returns {{ success: boolean, error?: string }}
   */
  updateCourse(originalName, courseData) {
    Auth.requireAdmin('科目更新');
    try {
      if (!originalName || !courseData || !courseData['科目名']) {
        throw new Error('科目名は必須です');
      }
      const newName = String(courseData['科目名']).trim();
      if (!newName) {
        throw new Error('科目名は必須です');
      }

      const courseSheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      if (!courseSheet) {
        throw new Error('科目データシートが見つかりません');
      }
      const courseAllData = SheetAccess.getDataWithRetry(courseSheet, '科目更新');
      const headers = courseAllData[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const nameIdx = SheetUtils.getColIndex(headerMap, '科目名');
      if (nameIdx === -1) {
        throw new Error('科目データシートに科目名列が見つかりません');
      }

      // 該当行を検索
      const found = SheetUtils.findRow(courseAllData, headerMap, '科目名', String(originalName).trim());
      if (!found) {
        throw new Error('科目が見つかりません: ' + originalName);
      }
      const targetRow = found.rowIndex;

      // 科目名変更の場合は新名の重複チェック
      if (newName !== String(originalName).trim()) {
        for (let i = 1; i < courseAllData.length; i++) {
          if (i + 1 === targetRow) continue;
          if (String(courseAllData[i][nameIdx]).trim() === newName) {
            return { success: false, error: '同名の科目が既に存在します: ' + newName };
          }
        }

        // 登録データシートのヘッダーもリネーム
        const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
        if (submissionSheet) {
          const subHeaders = submissionSheet.getRange(1, 1, 1, submissionSheet.getLastColumn()).getValues()[0];
          for (let j = 0; j < subHeaders.length; j++) {
            if (String(subHeaders[j]).trim() === String(originalName).trim()) {
              submissionSheet.getRange(1, j + 1).setValue(newName);
              break;
            }
          }
        }
      }

      // 各列の値を更新
      headers.forEach((header, colIndex) => {
        if (courseData.hasOwnProperty(header)) {
          courseSheet.getRange(targetRow, colIndex + 1).setValue(courseData[header]);
        }
      });

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('科目更新: ' + originalName + (newName !== originalName ? ' → ' + newName : ''));
      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '科目更新');
    }
  },

  /**
   * 科目を削除（履修データが存在する場合は拒否）
   * @param {string} courseName - 科目名
   * @returns {{ success: boolean, error?: string }}
   */
  deleteCourse(courseName) {
    Auth.requireAdmin('科目削除');
    try {
      if (!courseName) {
        throw new Error('科目名が指定されていません');
      }
      courseName = String(courseName).trim();

      // 登録データシートで該当科目の列を検索
      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }
      const subData = submissionSheet.getDataRange().getValues();
      const subHeaders = subData[0];
      let subColIdx = -1;
      for (let j = 0; j < subHeaders.length; j++) {
        if (String(subHeaders[j]).trim() === courseName) {
          subColIdx = j;
          break;
        }
      }

      // 安全チェック: 登録データシートに列がある場合、履修データを走査
      if (subColIdx !== -1) {
        let dataCount = 0;
        for (let i = 1; i < subData.length; i++) {
          const val = String(subData[i][subColIdx] || '').trim();
          if (/[●○⚪]/.test(val)) {
            dataCount++;
          }
        }
        if (dataCount > 0) {
          return {
            success: false,
            error: '履修データが存在するため削除できません（' + dataCount + '名分のデータがあります）'
          };
        }

        // 登録データシートの列を削除
        submissionSheet.deleteColumn(subColIdx + 1);
      }

      // 科目データシートの行を削除
      const courseSheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      if (!courseSheet) {
        throw new Error('科目データシートが見つかりません');
      }
      const courseAllData = courseSheet.getDataRange().getValues();
      const courseHeaders = courseAllData[0];
      const courseHeaderMap = SheetUtils.buildHeaderMap(courseHeaders);
      const found = SheetUtils.findRow(courseAllData, courseHeaderMap, '科目名', courseName);
      if (!found) {
        throw new Error('科目データシートに該当科目が見つかりません: ' + courseName);
      }
      courseSheet.deleteRow(found.rowIndex);

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('科目削除: ' + courseName);
      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '科目削除');
    }
  },

  /**
   * 科目の学籍番号列を更新
   * @param {string} courseName - 科目名
   * @param {string} columnName - 列名（'抽選学籍番号','非表示学籍番号','表示学籍番号'）
   * @param {string} studentIds - カンマ区切りの学籍番号文字列
   * @returns {{ success: boolean, error?: string }}
   */
  updateCourseStudentIds(courseName, columnName, studentIds) {
    Auth.requireAdmin('科目学籍番号更新');
    try {
      const allowedColumns = ['抽選学籍番号', '非表示学籍番号', '表示学籍番号'];
      if (!allowedColumns.includes(columnName)) {
        throw new Error('無効な列名: ' + columnName);
      }
      if (!courseName) {
        throw new Error('科目名が指定されていません');
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      const colIdx = SheetUtils.getColIndex(headerMap, columnName);
      if (colIdx === -1) {
        throw new Error('列が見つかりません: ' + columnName);
      }

      const found = SheetUtils.findRow(data, headerMap, '科目名', String(courseName).trim());
      if (!found) {
        throw new Error('科目が見つかりません: ' + courseName);
      }

      sheet.getRange(found.rowIndex, colIdx + 1).setValue(studentIds || '');

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.COURSE);

      Logger_.log('科目学籍番号更新: ' + courseName + ' / ' + columnName);

      return { success: true };
    } catch (error) {
      return Logger_.handleError(error, '科目学籍番号更新');
    }
  },

  /**
   * 科目別学籍番号を一括更新（Excel取込用）
   * @param {Array} updates - [{ courseName: string, studentIds: string }, ...]
   * @param {string} columnName - 列名
   * @returns {{ success: boolean, updatedCount?: number, error?: string }}
   */
  batchUpdateCourseStudentIds(updates, columnName) {
    Auth.requireAdmin('科目学籍番号一括更新');
    try {
      const allowedColumns = ['抽選学籍番号', '非表示学籍番号', '表示学籍番号'];
      if (!allowedColumns.includes(columnName)) {
        throw new Error('無効な列名: ' + columnName);
      }
      if (!updates || !Array.isArray(updates) || updates.length === 0) {
        throw new Error('更新データが空です');
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      const colIdx = SheetUtils.getColIndex(headerMap, columnName);
      if (colIdx === -1) {
        throw new Error('列が見つかりません: ' + columnName);
      }
      const nameIdx = SheetUtils.getColIndex(headerMap, '科目名');
      if (nameIdx === -1) {
        throw new Error('科目名列が見つかりません');
      }

      // 科目名→行番号マップを構築
      const courseRowMap = {};
      for (let i = 1; i < data.length; i++) {
        const name = String(data[i][nameIdx]).trim();
        if (name) courseRowMap[name] = i + 1;
      }

      let updatedCount = 0;
      updates.forEach(update => {
        const cn = String(update.courseName).trim();
        const sids = update.studentIds || '';
        const targetRow = courseRowMap[cn];
        if (!targetRow) {
          Logger_.log('科目が見つかりません（スキップ）: ' + cn, 'WARN');
          return;
        }
        sheet.getRange(targetRow, colIdx + 1).setValue(sids);
        updatedCount++;
      });

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.COURSE);

      Logger_.log('科目学籍番号一括更新: ' + columnName + ' / ' + updatedCount + '科目');

      return { success: true, updatedCount: updatedCount };
    } catch (error) {
      return Logger_.handleError(error, '科目学籍番号一括更新');
    }
  },
};

// --- 後方互換ラッパー ---
function getCourseData() { return CourseService.getCourseData(); }
function parseCourseSheet() { return CourseService.parseCourseSheet(); }
function findCourseByName(courseName, courseData) { return CourseService.findCourseByName(courseName, courseData); }
function getCourseList() { return CourseService.getCourseList(); }
function addCourse(courseData) { return CourseService.addCourse(courseData); }
function updateCourse(originalName, courseData) { return CourseService.updateCourse(originalName, courseData); }
function deleteCourse(courseName) { return CourseService.deleteCourse(courseName); }
function updateCourseStudentIds(courseName, columnName, studentIds) { return CourseService.updateCourseStudentIds(courseName, columnName, studentIds); }
function batchUpdateCourseStudentIds(updates, columnName) { return CourseService.batchUpdateCourseStudentIds(updates, columnName); }
