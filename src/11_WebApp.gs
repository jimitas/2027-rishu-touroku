/**
 * 11_WebApp.gs - WebAppエントリーポイント
 *
 * GAS WebApp の doGet, include, 初期データ取得などを提供する。
 * doGet, include, getWebAppUrl はGAS要件によりグローバル関数として定義する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs, 03_Auth.gs,
 *        04_CourseService.gs, 06_StudentService.gs, 08_SettingsService.gs
 */

const WebApp = {

  /**
   * 初期データを取得（キャッシュ対応版 - 行データをヘッダー付きで返す）
   * フロントエンドの初回ロードで呼ばれるメインAPI
   * @returns {Object} 初期データ（ユーザー情報、科目データ、登録状況等）
   */
  getInitialData() {
    try {
      // シートの存在確認
      const { valid, missing } = SheetAccess.validateSheets([
        CONFIG.SHEETS.COURSE, CONFIG.SHEETS.TEACHER, CONFIG.SHEETS.SUBMISSION
      ]);
      if (!valid) {
        throw new Error('必要なシートが見つかりません。シート名を確認してください。');
      }

      // 全キャッシュを事前に準備
      SheetAccess.getTeacherData();
      SheetAccess.getSubmissionData();
      SheetAccess.getCourseData();
      SheetAccess.getRosterData();

      const userEmail = Session.getActiveUser().getEmail();
      if (!userEmail) {
        throw new Error('メールアドレスが取得できません。Googleアカウントでログインしてください。');
      }

      // 1. 権限判定
      const userRoleResult = Auth.determineUserRole(userEmail);
      const userRole = userRoleResult ? userRoleResult.role : CONFIG.ROLES.STUDENT;

      // メンテナンスモード判定
      const settings = SettingsService.getSettings();
      const restriction = SettingsService.checkRestriction(userRole, null);
      if (!restriction.allowed) {
        return {
          maintenanceMode: true,
          settings: settings,
          gasUserEmail: userEmail
        };
      }

      // 2. 登録データから学籍番号と学生データを同時取得
      let studentId;
      let studentSubmissionData = null;

      {
        const submissionAllData = SheetAccess.getSubmissionData();
        if (submissionAllData.length > 1) {
          const submissionHeader = submissionAllData[0];
          const headerMap = SheetUtils.buildHeaderMap(submissionHeader);
          const emailIdx = SheetUtils.getColIndex(headerMap, 'メールアドレス');
          const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

          if (emailIdx !== -1 && studentIdIdx !== -1) {
            const found = SheetUtils.findRow(submissionAllData, headerMap, 'メールアドレス', userEmail);
            if (found) {
              studentId = found.rowData[studentIdIdx];
              studentSubmissionData = {
                header: submissionHeader,
                row: found.rowData
              };
            }
          }
        }

        // 教員判定
        if (!studentId) {
          if (Auth.isTeacherOrAbove(userRole)) {
            const teacherData = userRoleResult && userRoleResult.name
              ? { name: userRoleResult.name, クラス: userRoleResult.assignedClasses }
              : { name: '教職員' };

            const courseAllData = SheetAccess.getCourseData();
            const teacherCourseData = courseAllData.length > 1
              ? { header: courseAllData[0], rows: courseAllData.slice(1) }
              : { header: [], rows: [] };

            return {
              userRole: userRole,
              teacherData: teacherData,
              isTeacher: true,
              gasUserEmail: userEmail,
              message: '教員としてログインしました。',
              settings: settings,
              webAppUrl: ScriptApp.getService().getUrl(),
              courseData: teacherCourseData
            };
          }
          throw new Error(`認証エラー: メールアドレス「${userEmail}」が登録データシートに見つかりません。`);
        }
      }

      // 利用停止チェック
      if (studentSubmissionData) {
        const suspHeaderMap = SheetUtils.buildHeaderMap(studentSubmissionData.header);
        const suspStatusIdx = SheetUtils.getColIndex(suspHeaderMap, 'ステータス');
        if (suspStatusIdx !== -1 && studentSubmissionData.row[suspStatusIdx] === CONFIG.STATUSES.SUSPENDED) {
          return { suspended: true, settings: settings, gasUserEmail: userEmail, userRole: userRole };
        }
      }

      // 5. 科目マスタ
      let courseData;
      const courseAllData = SheetAccess.getCourseData();
      if (courseAllData.length > 1) {
        courseData = {
          header: courseAllData[0],
          rows: courseAllData.slice(1)
        };
      }

      // 提出状態チェック
      const submissionStatus = StudentService.checkSubmissionStatusInternal(studentId);

      // teacherData取得
      let teacherData;
      if (Auth.isTeacherOrAbove(userRole)) {
        teacherData = userRoleResult && userRoleResult.name
          ? { name: userRoleResult.name, クラス: userRoleResult.assignedClasses }
          : { name: '教職員' };
      } else {
        teacherData = { name: '教職員' };
      }

      // studentSubmissionDataをプリミティブ値に変換
      let serializedStudentData = null;
      let nextGradeOverride = null;
      if (studentSubmissionData) {
        serializedStudentData = {
          header: studentSubmissionData.header
            ? Array.from(studentSubmissionData.header).map(v => v != null ? String(v) : '')
            : [],
          row: studentSubmissionData.row
            ? Array.from(studentSubmissionData.row).map(v => v != null ? String(v) : '')
            : []
        };
        // 来年度学年列の値を抽出
        const nextGradeHeaderMap = SheetUtils.buildHeaderMap(studentSubmissionData.header);
        const nextGradeColIdx = SheetUtils.getColIndex(nextGradeHeaderMap, '来年度学年');
        if (nextGradeColIdx !== -1) {
          const val = parseInt(studentSubmissionData.row[nextGradeColIdx]);
          if (!isNaN(val) && val >= 1 && val <= 3) {
            nextGradeOverride = val;
          }
        }
      }

      // クラス非表示モード
      if (Auth.shouldHideClass() && serializedStudentData) {
        serializedStudentData = Auth.stripClassInfoFromHeaderRow(serializedStudentData);
      }

      return {
        userRole: userRole,
        studentId: studentId,
        gasUserEmail: userEmail,
        studentRosterData: serializedStudentData,
        studentSubmissionData: serializedStudentData,
        courseData: courseData,
        submissionStatus: submissionStatus,
        nextGradeOverride: nextGradeOverride,
        currentYear: new Date().getFullYear(),
        isTeacher: Auth.isTeacherOrAbove(userRole),
        teacherData: teacherData,
        message: Auth.isTeacherOrAbove(userRole)
          ? '教員+生徒データ両方読み込み完了（直接認証版）'
          : '直接認証完了',
        settings: settings,
        webAppUrl: ScriptApp.getService().getUrl()
      };

    } catch (e) {
      Logger_.log('getInitialData エラー: ' + e.message, 'ERROR');
      return {
        error: e.message,
        userInfo: {},
        allCourses: [],
        completedCourses: [],
        inProgressCourses: [],
        retakenYearMap: {},
        currentYear: new Date().getFullYear(),
        submissionStatus: { isSubmitted: false, status: null, timestamp: null }
      };
    }
  },

  /**
   * 認証付き初期データ取得
   * @returns {Object} ユーザー情報、科目マスタ、登録状況を含む初期データ
   */
  getInitialDataWithAuth() {
    try {
      // 全キャッシュを事前に準備
      SheetAccess.getTeacherData();
      SheetAccess.getCourseData();
      SheetAccess.getSubmissionData();
      SheetAccess.getRosterData();

      const email = Session.getActiveUser().getEmail();
      const userInfo = Auth.parseEmailForUserInfo(email);

      // 科目マスタを整形
      const courseData = SheetAccess.getCourseData();
      const courses = SheetUtils.toObjects(courseData).filter(c => c['科目名']);

      // 生徒の場合は登録状況も取得
      let registrationStatus = null;
      let completedCourses = [];
      let registeredCourses = [];
      let registrationData = {};

      if (userInfo.role === CONFIG.ROLES.STUDENT && userInfo.studentId) {
        const submissionData = SheetAccess.getSubmissionData();
        if (submissionData.length > 1) {
          const headers = submissionData[0];
          const headerMap = SheetUtils.buildHeaderMap(headers);
          const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
          const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');

          const found = SheetUtils.findRow(submissionData, headerMap, '学籍番号', userInfo.studentId);
          if (found) {
            registrationStatus = {
              status: found.rowData[statusIdx] || '未提出',
              completedCourses: [],
              registeredCourses: []
            };

            const basicColumnsSet = new Set(CONFIG.BASIC_COLUMNS);
            headers.forEach((header, index) => {
              if (!header) return;
              const value = found.rowData[index];
              if (!value) return;
              if (basicColumnsSet.has(header)) return;

              registrationData[header] = value;

              if (value === '●' || /^●[123]$/.test(value)) {
                completedCourses.push(header);
              } else if (value === '○' || /^○[123]N?$/.test(value)) {
                registeredCourses.push(header);
              }
            });

            registrationStatus.completedCourses = completedCourses;
            registrationStatus.registeredCourses = registeredCourses;
          }
        }
      }

      return {
        success: true,
        user: userInfo,
        courses: courses,
        registrationStatus: registrationStatus,
        registrationData: registrationData,
        completedCourses: completedCourses,
        timestamp: Logger_.getJSTTimestamp()
      };

    } catch (error) {
      Logger_.log('getInitialDataWithAuth エラー: ' + error.toString(), 'ERROR');
      return Logger_.handleError(error, '初期データ取得');
    }
  },

  /**
   * キャッシュウォームアップ（トリガーで定期実行用）
   */
  warmupCache() {
    try {
      SheetAccess.getTeacherData();
      SheetAccess.getCourseData();
      SheetAccess.getSubmissionData();
      SheetAccess.getRosterData();
    } catch (error) {
      Logger_.log('キャッシュウォームアップエラー: ' + error.toString(), 'ERROR');
    }
  },
};

// === GAS要件によりグローバル関数として定義 ===

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - リクエストパラメータ
 * @returns {HtmlOutput} HTMLページ
 */
function doGet(e) {
  try {
    const token = e.parameter.token || '';
    let html = HtmlService.createHtmlOutputFromFile('index').getContent();
    html = html.replace('%%TOKEN%%', token.replace(/[^a-zA-Z0-9\-_]/g, ''));

    return HtmlService.createHtmlOutput(html)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('履修登録システム - 京都芸術大学附属高等学校')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <html>
        <head><title>システムエラー</title></head>
        <body style="text-align:center;padding:40px;font-family:sans-serif;">
          <h2 style="color:#dc3545;">システムエラー</h2>
          <p>履修登録システムの起動に失敗しました</p>
          <button onclick="location.reload()">再読み込み</button>
        </body>
      </html>
    `);
  }
}

/**
 * HTMLファイル内でincludeするためのヘルパー関数
 * @param {string} filename - ファイル名
 * @returns {string} ファイル内容
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger_.log(`include エラー (${filename}): ${error.toString()}`, 'ERROR');
    return `<!-- Error loading ${filename}: ${error.toString()} -->`;
  }
}

/**
 * ウェブアプリのURLを取得
 * @returns {string} ウェブアプリURL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    Logger_.log(`ウェブアプリURL取得エラー: ${error.toString()}`, 'ERROR');
    return '';
  }
}

// --- 後方互換ラッパー ---
function getInitialData() { return WebApp.getInitialData(); }
function getInitialDataWithAuth() { return WebApp.getInitialDataWithAuth(); }
function warmupCache() { return WebApp.warmupCache(); }
