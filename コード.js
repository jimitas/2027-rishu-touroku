/**
 * 履修登録システム - Google Apps Script
 * 京都芸術大学附属高等学校 普通科
 * 
 * 構文エラーを解消し、モジュラー設計で再構築
 */

// ===============================
// スクリプトレベル インメモリキャッシュ
// （GASでは各実行が独立V8インスタンスのため古いデータのリスクなし）
// ===============================
let _memCache = {};

// ===============================
// Webアプリケーション エントリーポイント
// ===============================

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {object} e - リクエストパラメータ
 * @returns {HtmlOutput} HTMLページ
 */
function doGet(e) {
  try {
    const token = e.parameter.token || '';
    let html = HtmlService.createHtmlOutputFromFile('index').getContent();
    // tokenのみプレースホルダー置換（GAS iframeでURLパラメータ取得不可のため）
    html = html.replace('%%TOKEN%%', token.replace(/[^a-zA-Z0-9\-_]/g, ''));

    return HtmlService.createHtmlOutput(html)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('履修登録システム - 京都芸術大学附属高等学校')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    // 軽量なエラーページ
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
    logMessage(`include エラー (${filename}): ${error.toString()}`, 'ERROR');
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
    logMessage(`ウェブアプリURL取得エラー: ${error.toString()}`, 'ERROR');
    return '';
  }
}

// ===============================
// 定数・設定
// ===============================

// スプレッドシート参照（遅延ローディング用変数）
let SPREADSHEET = null;

// シート参照（遅延ローディング用変数）
let COURSE_MASTER_SHEET = null;
let TEACHER_DATA_SHEET = null;
let STUDENT_ROSTER_SHEET = null;
let SUBMISSION_SHEET = null;
let SETTINGS_SHEET = null;
let SM_SETTINGS_SHEET = null;

// キャッシュ設定
const CACHE_EXPIRY_MS = 5 * 60 * 1000; // 5分
const UNIT_LIMIT = 30;


// ===============================
// ユーティリティ関数
// ===============================

/**
 * スプレッドシートの遅延取得
 * @returns {SpreadsheetApp.Spreadsheet} スプレッドシート
 */
function getSpreadsheet() {
  if (!SPREADSHEET) {
    SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  }
  return SPREADSHEET;
}

/**
 * 科目データシートの遅延取得
 * @returns {SpreadsheetApp.Sheet} 科目データシート
 */
function getCourseSheet() {
  if (!COURSE_MASTER_SHEET) {
    COURSE_MASTER_SHEET = getSpreadsheet().getSheetByName('科目データ');
  }
  return COURSE_MASTER_SHEET;
}

/**
 * 教職員データシートの遅延取得
 * @returns {SpreadsheetApp.Sheet} 教職員データシート
 */
function getTeacherSheet() {
  if (!TEACHER_DATA_SHEET) {
    TEACHER_DATA_SHEET = getSpreadsheet().getSheetByName('教職員データ');
  }
  return TEACHER_DATA_SHEET;
}

/**
 * 生徒情報シートの遅延取得
 * @returns {SpreadsheetApp.Sheet} 生徒情報シート
 */
function getStudentRosterSheet() {
  if (!STUDENT_ROSTER_SHEET) {
    STUDENT_ROSTER_SHEET = getSpreadsheet().getSheetByName('生徒情報');
  }
  return STUDENT_ROSTER_SHEET;
}

/**
 * 登録データシートの遅延取得
 * @returns {SpreadsheetApp.Sheet} 登録データシート
 */
function getSubmissionSheet() {
  if (!SUBMISSION_SHEET) {
    SUBMISSION_SHEET = getSpreadsheet().getSheetByName('登録データ');
  }
  return SUBMISSION_SHEET;
}

/**
 * 設定シートの遅延取得
 * @returns {SpreadsheetApp.Sheet} 設定シート
 */
function getSettingsSheet() {
  if (!SETTINGS_SHEET) {
    SETTINGS_SHEET = getSpreadsheet().getSheetByName('設定');
    if (!SETTINGS_SHEET) {
      // シートが存在しない場合は作成
      SETTINGS_SHEET = getSpreadsheet().insertSheet('設定');
      console.log('設定シートを作成しました');

      // 初期データを設定
      SETTINGS_SHEET.getRange(1, 1, 1, 2).setValues([['設定名', '内容']]);
      SETTINGS_SHEET.getRange(2, 1, 7, 2).setValues([
        ['学校名', '京都芸術大学附属高等学校　普通科'],
        ['本登録許可', '可'],
        ['仮登録', '可'],
        ['1年最低単位数', '17'],
        ['2年最低単位数', '44'],
        ['3年最低単位数', '74'],
        ['生徒用ドメイン', '']
      ]);
      console.log('設定シートに初期データを追加しました');
    }
  }
  return SETTINGS_SHEET;
}

/**
 * 日本時間のタイムスタンプを取得
 * @returns {string} 日本時間のタイムスタンプ（yyyy-MM-dd HH:mm:ss形式）
 */
function getJSTTimestamp() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
}

/**
 * ログ出力用ヘルパー
 * @param {string} message - ログメッセージ
 * @param {string} level - ログレベル (INFO, ERROR, WARN)
 */
function logMessage(message, level = 'INFO') {
  const timestamp = getJSTTimestamp();
  const logMsg = `[${timestamp}] ${level}: ${message}`;
  console.log(logMsg);
}

/**
 * エラーハンドリング用ヘルパー
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラー発生コンテキスト
 * @returns {object} エラーレスポンス
 */
function handleError(error, context) {
  const errorMsg = `${context}: ${error.toString()}`;
  logMessage(errorMsg, 'ERROR');
  return {
    success: false,
    error: errorMsg,
    timestamp: getJSTTimestamp()
  };
}

/**
 * シートの存在確認
 * @returns {boolean} 全必要シートが存在するかどうか
 */
function validateSheets() {
  const requiredSheets = [
    { name: '科目データ', sheet: getCourseSheet() },
    { name: '教職員データ', sheet: getTeacherSheet() },
    { name: '登録データ', sheet: getSubmissionSheet() }
  ];
  
  let allSheetsFound = true;
  
  requiredSheets.forEach(({ name, sheet }) => {
    if (!sheet) {
      logMessage(`${name} シートが見つかりません`, 'ERROR');
      allSheetsFound = false;
    }
  });
  
  return allSheetsFound;
}

// ===============================
// 認証・ユーザー管理
// ===============================

/**
 * 現在のユーザー情報を取得
 * @returns {object} ユーザー情報
 */
function getUserInfo() {
  try {
    const email = Session.getActiveUser().getEmail();
    const userInfo = parseEmailForUserInfo(email);
    
    logMessage(`ユーザー認証成功: ${JSON.stringify(userInfo)}`);
    
    return {
      success: true,
      user: userInfo,
      timestamp: getJSTTimestamp()
    };
    
  } catch (error) {
    return handleError(error, 'ユーザー認証');
  }
}

/**
 * 教員/管理者権限を要求するガード関数
 * @param {string} context - エラーメッセージに含めるコンテキスト
 * @returns {object} userInfo
 */
function requireTeacherOrAdmin(context) {
  const userInfo = getUserInfo();
  if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
    throw new Error(`${context}には教員権限が必要です`);
  }
  return userInfo;
}

/**
 * 管理者権限を要求するガード関数
 * @param {string} context - エラーメッセージに含めるコンテキスト
 * @returns {object} userInfo
 */
function requireAdmin(context) {
  const userInfo = getUserInfo();
  if (!userInfo.success || userInfo.user.role !== 'admin') {
    throw new Error(`権限エラー: ${context}には管理者権限が必要です`);
  }
  return userInfo;
}

/**
 * 認証済みユーザー自身のstudentIdか教員/管理者であることを要求
 * @param {string} requestedStudentId - リクエストされた学籍番号
 * @param {string} context - エラーメッセージに含めるコンテキスト
 * @returns {object} userInfo
 */
function requireSelfOrTeacher(requestedStudentId, context) {
  const userInfo = getUserInfo();
  if (!userInfo.success) {
    throw new Error('ユーザー認証に失敗しました');
  }
  const user = userInfo.user;
  if (isTeacherOrAbove(user.role)) {
    return userInfo;
  }
  if (user.role === 'student' && String(user.studentId) === String(requestedStudentId)) {
    return userInfo;
  }
  throw new Error(`${context}への権限がありません`);
}

/**
 * クラス情報を非表示にすべきかを判定
 * 設定「生徒への表示」が「学籍番号」の場合、トップページではクラス情報を返さない
 * トップページは生徒・教員共通画面のため、ロールに関係なく同じ表示にする
 * @returns {boolean} クラス情報を非表示にすべきか
 */
function shouldHideClass() {
  const settings = getSettings();
  return settings['生徒への表示'] === '学籍番号';
}

/**
 * {header, row} 形式のデータから組・番号の値をクリア
 * @param {object} data - {header: [...], row: [...]} 形式のデータ
 * @returns {object} 組・番号をクリアしたデータ
 */
function stripClassInfoFromHeaderRow(data) {
  if (!data || !data.header || !data.row) return data;
  const classIdx = data.header.indexOf('組');
  const numberIdx = data.header.indexOf('番号');
  const newRow = [...data.row];
  if (classIdx !== -1) newRow[classIdx] = '';
  if (numberIdx !== -1) newRow[numberIdx] = '';
  return { ...data, row: newRow };
}

/**
 * {studentId, name, grade, class, number} 形式からclass/numberをクリア
 * @param {object} data - 生徒データオブジェクト
 * @returns {object} class/numberをクリアしたデータ
 */
function stripClassInfoFromStudentData(data) {
  if (!data) return data;
  return { ...data, class: '', number: '' };
}

/**
 * 【高速化】初期データ一括取得関数
 * フロントエンドの初回API呼び出しで、必要なすべてのデータを一括取得する
 * これにより複数回のAPI呼び出しを1回に削減
 * @returns {object} ユーザー情報、科目マスタ、登録状況を含む初期データ
 */
function getInitialDataWithAuth() {
  try {
    // 1. 全キャッシュを事前に準備（これにより後続の処理が高速化）
    const teacherData = getTeacherDataCached();
    const courseData = getCourseDataCached();
    const submissionData = getSubmissionSheetDataCached();
    const rosterData = getStudentRosterDataCached();

    // 2. ユーザー情報を取得（キャッシュ済みデータを使用するため高速）
    const email = Session.getActiveUser().getEmail();
    const userInfo = parseEmailForUserInfo(email);

    // 3. 科目マスタを整形
    const courses = [];
    if (courseData.length > 1) {
      const headers = courseData[0];
      for (let i = 1; i < courseData.length; i++) {
        const row = courseData[i];
        const course = {};
        headers.forEach((header, index) => {
          if (header) course[header] = row[index];
        });
        if (course['科目名']) courses.push(course);
      }
    }

    // 4. 生徒の場合は登録状況も取得
    let registrationStatus = null;
    let completedCourses = [];
    let registeredCourses = [];
    let registrationData = {};

    if (userInfo.role === 'student' && userInfo.studentId) {
      // 登録データから該当生徒の情報を取得
      if (submissionData.length > 1) {
        const headers = submissionData[0];
        const studentIdIndex = headers.indexOf('学籍番号');
        const statusIndex = headers.indexOf('ステータス');

        for (let i = 1; i < submissionData.length; i++) {
          const row = submissionData[i];
          if (row[studentIdIndex] == userInfo.studentId) {
            // ステータス
            registrationStatus = {
              status: row[statusIndex] || '未提出',
              completedCourses: [],
              registeredCourses: []
            };

            // 各科目の登録状況を抽出
            const basicColumnsSet = new Set(['学籍番号', 'ステータス', 'タイムスタンプ', '学年', '組', '番号', '名前', 'メールアドレス', '教職員チェック', '認証コード', '来年度学年']);
            headers.forEach((header, index) => {
              if (!header) return;
              const value = row[index];
              if (!value) return;

              // 基本列は除外
              if (basicColumnsSet.has(header)) return;

              // 科目データとして登録
              registrationData[header] = value;

              if (value === '●' || /^●[123]$/.test(value)) {
                completedCourses.push(header);
              } else if (value === '○' || /^○[123]N?$/.test(value)) {
                registeredCourses.push(header);
              }
            });

            registrationStatus.completedCourses = completedCourses;
            registrationStatus.registeredCourses = registeredCourses;
            break;
          }
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
      timestamp: getJSTTimestamp()
    };

  } catch (error) {
    console.error('getInitialDataWithAuth エラー:', error);
    return handleError(error, '初期データ取得');
  }
}

/**
 * 【高速化】キャッシュウォームアップ関数
 * トリガーで定期実行し、キャッシュを事前に構築しておく
 */
function warmupCache() {
  try {
    getTeacherDataCached();
    getCourseDataCached();
    getSubmissionSheetDataCached();
    getStudentRosterDataCached();
  } catch (error) {
    console.error('キャッシュウォームアップエラー:', error);
  }
}

/**
 * メールアドレスから学籍番号を取得（登録データシート直接検索版）
 * @param {string} email - メールアドレス
 * @returns {string|null} 学籍番号
 */
function getStudentIdFromEmail(email) {
  try {
    if (!getSubmissionSheet()) {
      console.error('登録データシートが見つかりません');
      return null;
    }

    // 高速検索：登録データシートでTextFinderを使用
    const submissionSheet = getSubmissionSheet();
    const finder = submissionSheet.createTextFinder(email);
    const foundCell = finder.findNext();

    if (foundCell) {
      const foundRow = foundCell.getRow();
      // キャッシュからヘッダーと学籍番号列を特定
      const cachedData = getSubmissionSheetDataCached();
      const headers = cachedData[0];
      const studentIdColumnIndex = headers.indexOf('学籍番号');

      if (studentIdColumnIndex === -1) {
        console.error('登録データシートに学籍番号列が見つかりません');
        return getSubmissionLinearSearchFallback(email);
      }

      // 学籍番号列のみ取得（行全体の取得は不要）
      const studentId = submissionSheet.getRange(foundRow, studentIdColumnIndex + 1).getValue();

      if (studentId) {
        return studentId;
      }
    }

    return getSubmissionLinearSearchFallback(email);

  } catch (error) {
    console.error('getStudentIdFromEmail エラー:', error);
    return getSubmissionLinearSearchFallback(email);
  }
}

/**
 * 登録データシート線形検索フォールバック（TextFinder失敗時のみ）
 * @param {string} email - メールアドレス
 * @returns {string|null} 学籍番号
 */
function getSubmissionLinearSearchFallback(email) {
  try {
    if (!getSubmissionSheet()) {
      return null;
    }

    const submissionData = getSubmissionSheetDataCached();
    if (submissionData.length <= 1) {
      return null;
    }

    const header = submissionData[0];
    const emailIndex = header.indexOf('メールアドレス');
    const studentIdIndex = header.indexOf('学籍番号');

    if (emailIndex === -1 || studentIdIndex === -1) {
      console.error('登録データシートにメールアドレス列または学籍番号列が見つかりません');
      return null;
    }

    // 簡潔な線形検索
    for (let i = 1; i < submissionData.length; i++) {
      const row = submissionData[i];
      if (row[emailIndex] === email && row[studentIdIndex]) {
        return row[studentIdIndex];
      }
    }

    return null;

  } catch (error) {
    console.error('登録データシート線形検索フォールバックエラー:', error);
    return null;
  }
}

/**
 * 教科書データ生成用：生徒情報シートから学籍番号で学生情報を取得
 * @param {string} studentId - 学籍番号
 * @returns {Object|null} 学生情報（ヘッダー付き行データ）
 */
function getStudentInfoForTextbook(studentId) {
  try {
    if (!getStudentRosterSheet()) {
      console.error('生徒情報シートが見つかりません（教科書データ生成用）');
      return null;
    }

    const rosterData = getStudentRosterDataCached();
    if (rosterData.length <= 1) {
      console.error('生徒情報シートにデータがありません');
      return null;
    }

    const header = rosterData[0];
    const studentIdIndex = header.indexOf('学籍番号');

    if (studentIdIndex === -1) {
      console.error('生徒情報シートに学籍番号列が見つかりません');
      return null;
    }

    // 学籍番号で検索
    const studentRow = rosterData.find(row => row[studentIdIndex] === studentId);
    if (studentRow) {
      return {
        header: header,
        row: studentRow
      };
    }

    return null;

  } catch (error) {
    console.error('教科書データ生成用学生情報取得エラー:', error);
    return null;
  }
}


/**
 * 提出状態を確認（内部用・認可チェックなし）
 * @param {string} studentId - 学籍番号
 * @returns {object} 提出状態
 */
function checkSubmissionStatusInternal(studentId) {
  try {
    if (!getSubmissionSheet()) {
      return { isSubmitted: false, status: null, timestamp: null };
    }

    // キャッシュからデータを取得
    const submissionData = getSubmissionSheetDataCached();
    const header = submissionData[0];
    const dataRows = submissionData.slice(1);

    const studentIdIndex = header.indexOf('学籍番号');
    const statusIndex = header.indexOf('ステータス');

    if (studentIdIndex === -1 || statusIndex === -1) {
      return { isSubmitted: false, status: null, timestamp: null };
    }

    // 最新の提出データを検索（仮提出または保存）
    const submittedRows = dataRows
      .filter(row => row[studentIdIndex] === studentId && (row[statusIndex] === '仮登録' || row[statusIndex] === '一時保存'))
      .sort((a, b) => new Date(b[0]) - new Date(a[0])); // タイムスタンプで降順ソート

    if (submittedRows.length > 0) {
      const latestRow = submittedRows[0];
      const timestamp = latestRow[0];
      return {
        isSubmitted: latestRow[statusIndex] === '仮登録',
        status: latestRow[statusIndex],
        timestamp: timestamp instanceof Date ? Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss") : String(timestamp)
      };
    }

    return { isSubmitted: false, status: null, timestamp: null };

  } catch (e) {
    console.error('提出状態チェックエラー:', e);
    return { isSubmitted: false, status: null, timestamp: null };
  }
}

/**
 * 提出状態を確認（公開用・認可チェック付き）
 * @param {string} studentId - 学籍番号
 * @returns {object} 提出状態
 */
function checkSubmissionStatus(studentId) {
  requireSelfOrTeacher(studentId, '提出状態確認');
  return checkSubmissionStatusInternal(studentId);
}

/**
 * メールアドレスから学籍番号を解析
 * @param {string} email - メールアドレス
 * @returns {object} ユーザー情報
 */
function parseEmailForUserInfo(email) {
  logMessage(`メールアドレス解析開始: ${email}`);

  // 1. 教職員データシートで検索（優先・1回の走査でロール+名前を取得）
  const teacherResult = determineUserRole(email);
  if (teacherResult !== null && isTeacherOrAbove(teacherResult.role)) {
    const teacherName = teacherResult.name || '教職員';
    logMessage(`教職員として認証: ${teacherName} (${teacherResult.role})`);
    return {
      email: email,
      role: teacherResult.role,
      name: teacherName
    };
  }

  // 2. 生徒データで検索
  const studentId = getStudentIdFromEmail(email);
  if (studentId) {
    const studentData = getStudentDataByIdInternal(studentId);
    logMessage(`生徒として認証: ${studentId}`);
    return {
      studentId: studentId,
      name: studentData?.name || `学生${studentId}`,
      grade: studentData?.grade ?? 1,
      class: studentData?.class || 1,
      number: studentData?.number || 1,
      role: 'student',
      email: email
    };
  }

  // 3. どちらにも該当しない場合はアクセス拒否
  logMessage(`認証失敗: どのデータにも該当しません: ${email}`, 'ERROR');
  throw new Error('このアカウントはシステムに登録されていません。管理者に連絡してください。');
}

/** ロールが教員以上か判定 */
function isTeacherOrAbove(role) {
  return role === 'teacher' || role === 'kyomu' || role === 'admin';
}

/** ロールが教務以上か判定 */
function isKyomuOrAbove(role) {
  return role === 'kyomu' || role === 'admin';
}

/**
 * ユーザーの役割を判定（ロール・名前・担当クラスを一括取得）
 * @param {string} email - メールアドレス
 * @returns {{ role: string, name: string|null, assignedClasses: string|null }} ユーザー役割情報
 */
function determineUserRole(email) {
  try {
    // 1. 教職員データシートでの教員・管理者判定を優先
    if (getTeacherSheet()) {
      const teacherData = getTeacherDataCached();
      if (teacherData.length > 1) {
        const teacherHeaders = teacherData[0];
        const emailIndex = teacherHeaders.indexOf('メールアドレス');
        const roleIndex = teacherHeaders.indexOf('権限');
        let nameIndex = teacherHeaders.indexOf('氏名');
        if (nameIndex === -1) nameIndex = teacherHeaders.indexOf('名前');
        const classesIndex = teacherHeaders.indexOf('クラス');

        if (emailIndex !== -1) {
          for (let i = 1; i < teacherData.length; i++) {
            const row = teacherData[i];
            if (String(row[emailIndex]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
              let role = 'teacher';
              if (roleIndex !== -1) {
                if (row[roleIndex] === '管理者') role = 'admin';
                else if (row[roleIndex] === '教務') role = 'kyomu';
              }
              const result = {
                role: role,
                name: nameIndex !== -1 ? row[nameIndex] : null,
                assignedClasses: classesIndex !== -1 ? row[classesIndex] : null
              };
              return result;
            }
          }
        }
      }
    }

    // 2. 教職員データシートに該当しない場合は全て生徒として扱う
    return { role: 'student', name: null, assignedClasses: null };

  } catch (error) {
    logMessage(`権限判定エラー: ${error.toString()}`, 'ERROR');
    return null;
  }
}

/**
 * 学籍番号から生徒データを取得（内部用・認可チェックなし）
 * @param {string} studentId - 学籍番号
 * @returns {object|null} 生徒データ
 */
// 学年値の安全なパース（0を許容）
function safeParseGradeValue(value, defaultValue) {
  if (value == null || value === '') return defaultValue;
  var parsed = parseInt(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

function getStudentDataByIdInternal(studentId) {
  if (!getStudentRosterSheet()) {
    return null;
  }

  try {
    const data = getStudentRosterDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const nameIndex = headers.indexOf('氏名');
    const gradeIndex = headers.indexOf('学年');
    const classIndex = headers.indexOf('組');
    const numberIndex = headers.indexOf('番号');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[studentIdIndex] == studentId) {
        return {
          studentId: studentId,
          name: row[nameIndex],
          grade: safeParseGradeValue(row[gradeIndex], 1),
          class: row[classIndex] != null ? String(row[classIndex]).trim() : '',
          number: parseInt(row[numberIndex]) || 1
        };
      }
    }

    return null;
  } catch (error) {
    logMessage(`生徒データ取得エラー: ${error.toString()}`, 'ERROR');
    return null;
  }
}

/**
 * 学籍番号から生徒データを取得（公開用・認可チェック付き）
 * @param {string} studentId - 学籍番号
 * @returns {object|null} 生徒データ
 */
function getStudentDataById(studentId) {
  const userInfo = requireSelfOrTeacher(studentId, '生徒データ取得');
  const result = getStudentDataByIdInternal(studentId);
  if (shouldHideClass()) {
    return stripClassInfoFromStudentData(result);
  }
  return result;
}

/**
 * 【高速化】バッチ処理用の生徒データルックアップMapを構築
 * ループ内でgetStudentDataById()を繰り返し呼ぶ代わりにこのMapを使用
 * @returns {Map} 学籍番号→生徒データのMap
 */
function buildStudentLookupMap() {
  const studentMap = new Map();

  if (!getStudentRosterSheet()) {
    return studentMap;
  }

  try {
    const data = getStudentRosterDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const nameIndex = headers.indexOf('氏名');
    const gradeIndex = headers.indexOf('学年');
    const classIndex = headers.indexOf('組');
    const numberIndex = headers.indexOf('番号');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentId = row[studentIdIndex];
      if (studentId) {
        studentMap.set(String(studentId), {
          studentId: String(studentId),
          name: row[nameIndex],
          grade: safeParseGradeValue(row[gradeIndex], 1),
          class: row[classIndex] != null ? String(row[classIndex]).trim() : '',
          number: parseInt(row[numberIndex]) || 1
        });
      }
    }

    return studentMap;
  } catch (error) {
    logMessage(`生徒Mapビルドエラー: ${error.toString()}`, 'ERROR');
    return studentMap;
  }
}

// ===============================
// データ取得・管理
// ===============================

/**
 * 科目データを取得
 * @returns {object} 科目データ
 */
function getCourseData() {
  try {
    // 認証チェック: 認証済みユーザーのみアクセス可能
    getUserInfo();

    if (!validateSheets()) {
      throw new Error('必要なシートが見つかりません');
    }

    const courseData = parseCourseSheet();
    
    return {
      success: true,
      data: courseData,
      timestamp: getJSTTimestamp()
    };
    
  } catch (error) {
    return handleError(error, '科目データ取得');
  }
}

/**
 * 科目シートを解析
 * @returns {object} 学年・科目別データ
 */
function parseCourseSheet() {
  const data = getCourseDataCached();
  const headers = data[0];
  
  const coursesByGrade = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const course = {};
    
    headers.forEach((header, index) => {
      course[header] = row[index];
    });
    
    const grade = `year${course['学年'] || 1}`;
    if (!coursesByGrade[grade]) {
      coursesByGrade[grade] = {
        required: [],
        elective: []
      };
    }
    
    const category = course['区分'] === '必修' ? 'required' : 'elective';
    coursesByGrade[grade][category].push(course);
  }
  
  return coursesByGrade;
}

/**
 * 履修データを保存
 * @param {object} registrationData - 履修データ
 * @returns {object} 保存結果
 */
function saveRegistrationData(registrationData) {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success) {
      throw new Error('ユーザー認証に失敗');
    }
    checkServerSideRestrictions(userInfo.user.role, '一時保存');

    const studentId = userInfo.user.studentId;
    // ステータスは常に「保存」に強制（クライアント指定を無視し、ステータス注入を防止）
    const status = '保存';
    const result = saveToSubmissionSheet(studentId, registrationData, status, userInfo.user.email);
    
    logMessage(`履修データ保存完了: 学籍番号 ${studentId}`);
    
    return {
      success: true,
      message: '履修データを保存しました',
      timestamp: getJSTTimestamp()
    };
    
  } catch (error) {
    return handleError(error, '履修データ保存');
  }
}

/**
 * 履修データを仮登録
 * @param {object} registrationData - 履修データ
 * @returns {object} 提出結果
 */
function submitRegistrationData(registrationData) {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success) {
      throw new Error('ユーザー認証に失敗');
    }
    checkServerSideRestrictions(userInfo.user.role, '仮登録');

    const studentId = userInfo.user.studentId;

    // 次年度の学年を算出（来年度学年列優先、なければ現在の学年+1、最大3）
    const currentGrade = safeParseGradeValue(userInfo.user.grade, 1);
    const registrationYear = getNextGradeForStudent(studentId, currentGrade);

    // データ整合性チェック
    const validation = validateRegistrationData(registrationData, registrationYear);
    if (!validation.valid) {
      throw new Error(`データ整合性エラー: ${validation.errors.join(', ')}`);
    }

    const result = saveToSubmissionSheet(studentId, registrationData, '仮登録', userInfo.user.email);
    
    logMessage(`履修データ仮登録完了: 学籍番号 ${studentId}`);
    
    return {
      success: true,
      message: '履修データを仮登録しました',
      timestamp: getJSTTimestamp()
    };
    
  } catch (error) {
    return handleError(error, '履修データ仮登録');
  }
}

/**
 * 教務が生徒の代わりに履修データを仮登録
 * @param {string} studentId - 学籍番号
 * @param {object} registrationData - 履修データ
 * @returns {object} 提出結果
 */
function adminSubmitRegistration(studentId, registrationData) {
  try {
    // 教員権限チェック
    const userInfo = getUserInfo();
    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員権限が必要です');
    }
    checkServerSideRestrictions(userInfo.user.role, null);

    // 生徒の学年から次年度を算出（来年度学年列優先）
    const studentInfo = getStudentBasicInfo(studentId);
    const adminCurrentGrade = safeParseGradeValue(studentInfo?.学年, 1);
    const registrationYear = getNextGradeForStudent(studentId, adminCurrentGrade);

    // データ整合性チェック
    const validation = validateRegistrationData(registrationData, registrationYear);
    if (!validation.valid) {
      throw new Error(`データ整合性エラー: ${validation.errors.join(', ')}`);
    }

    const result = adminSaveToSubmissionSheet(studentId, registrationData, '仮登録');

    logMessage(`教職員による仮登録完了: 学籍番号 ${studentId}`);

    return {
      success: true,
      message: '履修データを仮登録しました',
      timestamp: getJSTTimestamp()
    };

  } catch (error) {
    return handleError(error, '教職員による仮登録');
  }
}

/**
 * 管理者による登録データの直接編集保存
 * ステータスは変更せず、マークのみ更新する
 */
function adminUpdateRegistrationMarks(studentId, registrationData, newStatus, basicFields) {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success || userInfo.user.role !== 'admin') {
      throw new Error('管理者権限が必要です');
    }

    // 現在のステータスを取得して保持
    const allData = getSubmissionSheetDataCached();
    const headers = allData.length > 0 ? allData[0] : [];
    const statusIdx = headers.indexOf('ステータス');
    const idIdx = headers.indexOf('学籍番号');
    let currentStatus = '保存';
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][idIdx]) === String(studentId)) {
        currentStatus = allData[i][statusIdx] || '保存';
        break;
      }
    }

    // newStatusが指定されていればそれを使用、なければ現在のステータスを保持
    const statusToSave = (newStatus != null) ? newStatus : currentStatus;

    // 基本情報の更新（科目マークとは別にシートに直接書き込み）
    if (basicFields && typeof basicFields === 'object') {
      const submissionSheet = getSubmissionSheet();
      const headerMap = new Map(headers.map(function(h, i) { return [h, i]; }));

      var targetRow = -1;
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][idIdx]) === String(studentId)) {
          targetRow = i + 1; // シートは1-indexed、ヘッダー行分+1
          break;
        }
      }

      if (targetRow !== -1) {
        var allowedFields = ['学籍番号', '学年', '組', '番号', 'メールアドレス', '来年度学年', '認証コード'];
        allowedFields.forEach(function(field) {
          if (field in basicFields) {
            var colIdx = headerMap.get(field);
            if (colIdx !== undefined) {
              submissionSheet.getRange(targetRow, colIdx + 1).setValue(basicFields[field]);
            }
          }
        });
      }
    }

    // バリデーションはスキップ（管理者による自由編集）
    adminSaveToSubmissionSheet(studentId, registrationData, statusToSave);

    // キャッシュクリア（次回読み込み時に最新データを取得するため）
    clearDataCache();

    logMessage(`管理者による登録データ直接編集: 学籍番号 ${studentId}`);

    return {
      success: true,
      message: '登録データを更新しました',
      timestamp: getJSTTimestamp()
    };
  } catch (error) {
    return handleError(error, '登録データ管理');
  }
}

/**
 * 履修データの検証
 * @param {object} registrationData - 履修データ
 * @returns {object} 検証結果
 */
function validateRegistrationData(registrationData, registrationYear) {
  const errors = [];

  try {
    // 基本的なデータ構造チェック
    if (!registrationData || typeof registrationData !== 'object') {
      errors.push('履修データが無効です');
      return { valid: false, errors: errors };
    }

    // 検証1: 選択科目の値チェック（全ての正当な履修マーカーを許可）
    const VALID_MARKS = new Set(['○', '〇', '●', '○1', '○2', '○3', '〇1', '〇2', '〇3', '●1', '●2', '●3', '○1N', '○2N', '○3N', '〇1N', '〇2N', '〇3N']);
    // メタデータキー: saveToSubmissionSheet()のbasicColumnsと統一 + フロントエンド追加フィールド
    const METADATA_KEYS = new Set([
      '学籍番号', 'ステータス', 'タイムスタンプ', '学年', '組', '番号',
      '名前', 'メールアドレス', '教職員チェック', '認証コード',
      '提出日時', '更新日時', '来年度学年'
    ]);
    Object.entries(registrationData).forEach(function(entry) {
      var key = entry[0];
      var value = entry[1];
      if (METADATA_KEYS.has(key)) return;
      if (value === '' || value === null || value === undefined) return;
      if (!VALID_MARKS.has(value)) {
        errors.push('不正な値: ' + key + '=' + value);
      }
    });

    // 検証2: 30単位制限（次年度の登録科目のみ）
    var totalUnits = calculateTotalUnits(registrationData, registrationYear);
    if (totalUnits > 30) {
      errors.push('単位数上限超過: ' + totalUnits + '単位（上限30単位）');
    }

    return {
      valid: errors.length === 0,
      errors: errors
    };

  } catch (error) {
    return {
      valid: false,
      errors: ['検証処理エラー: ' + error.toString()]
    };
  }
}

/**
 * 登録単位数の計算（次年度の登録科目のみカウント）
 * @param {object} registrationData - 履修データ
 * @param {number} registrationYear - 登録対象の学年（例: 現在2年なら3）
 * @returns {number} 総単位数
 */
function calculateTotalUnits(registrationData, registrationYear) {
  let totalUnits = 0;

  try {
    const courseData = getCourseData();
    if (!courseData.success) {
      throw new Error('科目データの取得に失敗');
    }

    // 次年度の登録科目のみカウント対象とするマーク
    // ○/〇（年番号なし＝新規選択）と ○N/〇N（N == registrationYear）のみ
    const targetMarks = new Set(['○', '〇']);
    if (registrationYear) {
      targetMarks.add('○' + registrationYear);
      targetMarks.add('〇' + registrationYear);
      targetMarks.add('○' + registrationYear + 'N');  // 留年生用Nサフィックス
      targetMarks.add('〇' + registrationYear + 'N');
    }

    // 登録された科目の単位数を合計（次年度分のみ）
    Object.entries(registrationData).forEach(([courseName, value]) => {
      if (!targetMarks.has(value)) return; // ●（履修済み）や他年度の○Nはスキップ

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
    logMessage(`単位数計算エラー: ${error.toString()}`, 'ERROR');
    return 0;
  }
}


/**
 * 登録データシートの「来年度学年」列から次年度学年を取得
 * 値が未設定の場合は currentGrade+1（最大3）を返す
 * @param {string} studentId - 学籍番号
 * @param {number} currentGrade - 現在の学年
 * @returns {number} 次年度の学年
 */
function getNextGradeForStudent(studentId, currentGrade) {
  const data = getSubmissionSheetDataCached();
  if (data.length < 2) return Math.min(currentGrade + 1, 3);
  const headers = data[0];
  const nextGradeColIdx = headers.indexOf('来年度学年');
  const studentIdColIdx = headers.indexOf('学籍番号');
  if (nextGradeColIdx === -1 || studentIdColIdx === -1) return Math.min(currentGrade + 1, 3);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][studentIdColIdx]).trim() === String(studentId).trim()) {
      const val = parseInt(data[i][nextGradeColIdx]);
      if (!isNaN(val) && val >= 1 && val <= 3) return val;
      break;
    }
  }
  return Math.min(currentGrade + 1, 3);
}

/**
 * 学生基本情報を生徒情報シートから取得
 * @param {string} studentId - 学籍番号
 * @returns {object|null} 学生基本情報
 */
function getStudentBasicInfo(studentId) {
  try {
    if (!getStudentRosterSheet()) {
      console.warn('生徒情報シートが見つかりません');
      return null;
    }

    const data = getStudentRosterDataCached();
    if (data.length === 0) return null;

    const headers = data[0];
    const studentIdIndex = headers.indexOf('学籍番号');

    if (studentIdIndex === -1) {
      console.warn('生徒情報シートに学籍番号列が見つかりません');
      return null;
    }

    // 【高速化】ヘッダーインデックスをループ外で事前計算
    const gradeIndex = headers.indexOf('学年');
    const classIndex = headers.indexOf('組');
    const numberIndex = headers.indexOf('番号');
    const nameIndex = headers.indexOf('名前');

    // 該当学生の行を検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][studentIdIndex] == studentId) {
        const studentRow = data[i];
        const emailIdx = headers.indexOf('メールアドレス');
        return {
          学年: studentRow[gradeIndex] || '',
          組: studentRow[classIndex] || '',
          番号: studentRow[numberIndex] || '',
          名前: studentRow[nameIndex] || '',
          メールアドレス: emailIdx !== -1 ? (studentRow[emailIdx] || '') : ''
        };
      }
    }
    
    console.warn(`学籍番号${studentId}の学生情報が見つかりません`);
    return null;
    
  } catch (error) {
    console.error(`学生基本情報取得エラー: ${error}`);
    return null;
  }
}

/**
 * 生徒用: 履修データをシートに保存（メールアドレスで行検索）
 */
function saveToSubmissionSheet(studentId, data, status = '保存', email = null) {
  return _saveToSubmissionSheetInternal(studentId, data, status, email, false);
}

/**
 * 教務用: 履修データをシートに保存（学籍番号で行検索）
 */
function adminSaveToSubmissionSheet(studentId, data, status) {
  return _saveToSubmissionSheetInternal(studentId, data, status, null, true);
}

/**
 * 履修データを登録データシートに保存する（内部実装）
 * @param {string} studentId - 学籍番号
 * @param {object} data - 履修データ
 * @param {string} status - 提出状態
 * @param {string} email - メールアドレス（生徒用検索時に使用）
 * @param {boolean} useStudentIdSearch - trueなら学籍番号で行検索（教務用）
 * @returns {boolean} 保存成功フラグ
 */
function _saveToSubmissionSheetInternal(studentId, data, status = '保存', email = null, useStudentIdSearch = false) {
  try {
    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const submissionSheet = getSubmissionSheet();
    // キャッシュからヘッダーを取得（シートへの個別読み取りを省略）
    const allData = getSubmissionSheetDataCached();
    const headers = allData.length > 0 ? allData[0] : submissionSheet.getRange(1, 1, 1, submissionSheet.getLastColumn()).getValues()[0];
    // ヘッダーMapを構築（O(1)ルックアップ）
    const headerMap = new Map(headers.map((h, i) => [h, i]));
    // 安全な列番号取得関数
    function getColumnIndex(columnName) {
      const index = headerMap.get(columnName);
      if (index === undefined) {
        console.error(`❌ 列が見つかりません: "${columnName}"`);
        return null;
      }
      return index + 1; // 1ベースの列番号
    }

    // 基本列のインデックス取得
    const studentIdIndex = getColumnIndex('学籍番号');
    const statusIndex = getColumnIndex('ステータス');
    const timestampIndex = getColumnIndex('タイムスタンプ');

    // 必要な基本列が見つからない場合はエラー
    if (!studentIdIndex || !statusIndex) {
      throw new Error(`基本列が見つかりません: 学籍番号(${studentIdIndex}), ステータス(${statusIndex})`);
    }

    // 既存データの検索
    const emailIndex = getColumnIndex('メールアドレス');
    let targetRow = -1;

    if (useStudentIdSearch) {
      // 教務用: 学籍番号で行検索
      const studentIdColIdx = headerMap.get('学籍番号');
      if (studentIdColIdx !== undefined) {
        for (let i = 1; i < allData.length; i++) {
          if (String(allData[i][studentIdColIdx]) === String(studentId)) {
            targetRow = i + 1;
            break;
          }
        }
      }
    } else {
      // 生徒用: メールアドレスで行検索
      const searchEmail = email || Session.getActiveUser().getEmail();
      if (emailIndex && searchEmail) {
        for (let i = 1; i < allData.length; i++) {
          if (allData[i][emailIndex - 1] == searchEmail) {
            targetRow = i + 1;
            break;
          }
        }
      }
    }
    
    // === 最適化: 行データをメモリ上で構築し、一括で書き込み ===
    const lastCol = headers.length;
    let rowData;
    const isNewRow = (targetRow === -1);

    if (isNewRow) {
      targetRow = submissionSheet.getLastRow() + 1;
      rowData = new Array(lastCol).fill('');

      // 学生基本情報を設定（メモリ上で）
      const studentInfo = getStudentBasicInfo(studentId);
      rowData[studentIdIndex - 1] = studentId;

      if (studentInfo) {
        const gradeIdx = headerMap.get('学年');
        const classIdx = headerMap.get('組');
        const numberIdx = headerMap.get('番号');
        const nameIdx = headerMap.get('名前');

        if (gradeIdx !== undefined) rowData[gradeIdx] = studentInfo.学年;
        if (classIdx !== undefined) rowData[classIdx] = studentInfo.組;
        if (numberIdx !== undefined) rowData[numberIdx] = studentInfo.番号;
        if (nameIdx !== undefined) rowData[nameIdx] = studentInfo.名前;
        const emailIdx = headerMap.get('メールアドレス');
        if (emailIdx !== undefined) rowData[emailIdx] = studentInfo.メールアドレス || '';

      } else {
        console.warn(`⚠️ 学生基本情報が見つかりません: ${studentId}`);
      }
    } else {
      // 既存行のデータを取得（1回のgetValues）
      rowData = submissionSheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
    }

    // 科目列を特定するため、ヘッダーから科目名を抽出（基本列以外）
    const basicColumns = ['学籍番号', 'ステータス', 'タイムスタンプ', '学年', '組', '番号', '名前', 'メールアドレス', '教職員チェック', '認証コード', '来年度学年'];
    const courseColumns = headers.filter(header => header && !basicColumns.includes(header));

    // 科目データをすべてクリア（メモリ上で）
    courseColumns.forEach(courseName => {
      const idx = headerMap.get(courseName);
      if (idx !== undefined) rowData[idx] = '';
    });

    // 科目セルに書き込む値をサニタイズ（数式インジェクション防止）
    function sanitizeCellValue(value) {
      if (typeof value !== 'string') return '';
      // 先頭が = + - @ のセルは数式インジェクションの危険があるため除去
      if (/^[=+\-@]/.test(value)) return '';
      return value;
    }

    // 新しい科目データを設定（メモリ上で）
    let savedSubjects = 0;
    Object.keys(data).forEach(courseName => {
      if (courseName === 'タイムスタンプ') return;

      const idx = headerMap.get(courseName);
      if (idx !== undefined) {
        rowData[idx] = sanitizeCellValue(data[courseName]);
        savedSubjects++;
      } else {
        console.warn(`⚠️ 科目「${courseName}」の列が見つかりません`);
      }
    });

    // ステータスを設定（メモリ上で）
    rowData[statusIndex - 1] = status;

    // タイムスタンプを設定（メモリ上で）
    if (timestampIndex) {
      rowData[timestampIndex - 1] = getJSTTimestamp();
    }

    // メールアドレスを設定（メモリ上で）
    if (emailIndex) {
      if (email) {
        // 明示的にメールが渡された場合はそれを使用
        rowData[emailIndex - 1] = email;
      } else if (!useStudentIdSearch) {
        // 生徒用: Sessionから取得
        try {
          const userEmail = Session.getActiveUser().getEmail();
          if (userEmail && userEmail.trim()) {
            rowData[emailIndex - 1] = userEmail;
          }
        } catch (emailError) {
          console.error(`❌ [メールアドレス設定] メールアドレス取得エラー: ${emailError}`);
        }
      }
      // 教務用 + email未指定: 既存行なら既存値を保持、新規行ならstudentInfoから設定済み
    }

    // 一括で書き込み（1回のsetValues）
    submissionSheet.getRange(targetRow, 1, 1, lastCol).setValues([rowData]);

    // 登録データのキャッシュをクリア
    try {
      clearSubmissionSheetCache();
    } catch (cacheError) {
      console.warn('登録データキャッシュクリアに失敗:', cacheError);
      // キャッシュクリアエラーは保存処理に影響させない
    }

    return true;
    
  } catch (error) {
    console.error(`❌ _saveToSubmissionSheetInternal エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 履修データを読み込み（内部用・認可チェックなし）
 * @param {string} studentId - 学籍番号
 * @returns {object} 履修データ
 */
function loadRegistrationDataInternal(studentId) {
  try {
    if (!getSubmissionSheet()) {
      return { success: true, data: {} };
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    if (studentIdIndex === -1 || timestampIndex === -1) {
      return { success: true, data: {} };
    }

    // 該当学生の行を検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowStudentId = row[studentIdIndex];

      // 文字列として統一比較
      if (String(rowStudentId).trim() === String(studentId).trim()) {
        // rowの中のDateオブジェクトを文字列に変換（google.script.runでシリアライズ可能にするため）
        const serializedRow = row.map(cell => {
          if (cell instanceof Date) {
            return cell.toISOString();
          }
          return cell;
        });

        const timestamp = row[timestampIndex];
        return {
          success: true,
          header: headers,
          row: serializedRow,
          timestamp: timestamp instanceof Date ? timestamp.toISOString() : timestamp
        };
      }
    }

    return { success: true, data: {} };

  } catch (error) {
    console.error('[データ読み込み] エラー発生:', error);
    return handleError(error, '履修データ読み込み');
  }
}

/**
 * 履修データを読み込み（公開用・認可チェック付き）
 * @param {string} studentId - 学籍番号
 * @returns {object} 履修データ
 */
function loadRegistrationData(studentId) {
  const userInfo = requireSelfOrTeacher(studentId, '履修データ読み込み');
  const result = loadRegistrationDataInternal(studentId);
  if (shouldHideClass()) {
    return stripClassInfoFromHeaderRow(result);
  }
  return result;
}

// ===============================
// 履修済み・再履修データ処理
// ===============================

/**
 * 履修済み科目データを取得
 * @param {string} studentId - 学籍番号
 * @returns {object} 履修済み科目データ
 */
function getCompletedCourses(studentId) {
  try {
    requireSelfOrTeacher(studentId, '履修済み科目取得');
    // 履修済み科目シートからデータを取得
    if (!COMPLETED_COURSES_SHEET) {
      return { success: true, completedCourses: [], retakenCourses: [] };
    }
    
    const data = getDataWithRetry(COMPLETED_COURSES_SHEET, 'loadCompletedCourses');
    const headers = data[0];
    
    const studentIdIndex = headers.indexOf('学籍番号');
    const courseNameIndex = headers.indexOf('科目名');
    const statusIndex = headers.indexOf('履修状況');
    const gradeIndex = headers.indexOf('成績');
    
    const completedCourses = [];
    const retakenCourses = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[studentIdIndex] == studentId) {
        const courseName = row[courseNameIndex];
        const status = row[statusIndex];
        const grade = row[gradeIndex];
        
        if (status === '履修済み') {
          completedCourses.push(courseName);
        } else if (status === '再履修' || grade === '不可' || grade === 'F') {
          retakenCourses.push(courseName);
        }
      }
    }
    
    return {
      success: true,
      completedCourses: completedCourses,
      retakenCourses: retakenCourses,
      timestamp: getJSTTimestamp()
    };
    
  } catch (error) {
    return handleError(error, '履修済み科目データ取得');
  }
}

/**
 * 再履修科目の学年マッピングを生成
 * @param {Array} retakenCourses - 再履修科目リスト
 * @returns {object} 再履修科目の学年マッピング
 */
function generateRetakenYearMap(retakenCourses) {
  const retakenYearMap = {};
  
  try {
    const courseData = getCourseData();
    if (!courseData.success) return retakenYearMap;
    
    // 各再履修科目の学年を特定
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
    logMessage(`再履修学年マップ生成エラー: ${error.toString()}`, 'ERROR');
    return retakenYearMap;
  }
}

// ===============================
// 教員・管理機能
// ===============================

/**
 * 全生徒の履修データを取得（教員用）
 * @returns {object} 全生徒の履修データ
 */
function getAllStudentsDataInternal(includeTeachers) {
  try {
    if (!getSubmissionSheet()) {
      return { success: true, students: [] };
    }

    const data = getSubmissionSheetDataCached();
    if (data.length === 0) {
      return { success: true, students: [] };
    }
    const headers = data[0];
    if (!headers || headers.length === 0) {
      return { success: true, students: [] };
    }

    let students = [];

    // 除外すべき列のSet（システム管理用の列、O(1)ルックアップ）
    const excludeColumns = new Set(['タイムスタンプ', '更新日時', '提出日時', '教職員チェック', '教務チェック']);

    // メールアドレス・学籍番号が空の行（空行）をスキップするためのインデックス
    const emailIndex = headers.indexOf('メールアドレス');
    const studentIdIndex = headers.indexOf('学籍番号');

    // 教職員メールSetを構築（教職員データを除外するため）
    const teacherEmails = new Set();
    const teacherData = getTeacherDataCached();
    if (teacherData.length > 1) {
      const tHeaders = teacherData[0];
      const tEmailIdx = tHeaders.indexOf('メールアドレス');
      if (tEmailIdx !== -1) {
        for (let t = 1; t < teacherData.length; t++) {
          const email = String(teacherData[t][tEmailIdx] || '').trim().toLowerCase();
          if (email) teacherEmails.add(email);
        }
      }
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 学籍番号が空白の行のみスキップ（メールアドレスが空でも一覧に表示する）
      if (!row[studentIdIndex] && studentIdIndex >= 0) continue;
      // 教職員データを除外（includeTeachers=true の場合はスキップ）
      if (!includeTeachers && emailIndex >= 0) {
        const email = String(row[emailIndex] || '').trim().toLowerCase();
        if (teacherEmails.has(email)) continue;
      }
      const studentData = {};

      headers.forEach((header, index) => {
        // 除外列Setに含まれない場合のみ追加
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

    return {
      success: true,
      students: students,
      timestamp: getJSTTimestamp()
    };
    
  } catch (error) {
    return handleError(error, '全生徒データ取得');
  }
}

/**
 * 全生徒の履修データを取得（公開用・認可チェック付き）
 * @returns {object} 全生徒の履修データ
 */
function getAllStudentsData(includeTeachers) {
  requireTeacherOrAdmin('全生徒データ取得');
  return getAllStudentsDataInternal(includeTeachers);
}

/**
 * 軽量ステータス取得（バッジ自動更新用）
 * 学籍番号とステータスのみ返却し通信量を最小化
 */
function getStudentStatuses() {
  requireTeacherOrAdmin('ステータス取得');
  try {
    const sheet = getSubmissionSheet();
    if (!sheet) return { success: true, statuses: {} };

    const data = getSubmissionSheetDataCached();
    if (data.length === 0) return { success: true, statuses: {} };

    const headers = data[0];
    const idIdx = headers.indexOf('学籍番号');
    const statusIdx = headers.indexOf('ステータス');
    if (idIdx === -1 || statusIdx === -1) return { success: true, statuses: {} };

    const statuses = {};
    for (let i = 1; i < data.length; i++) {
      const id = data[i][idIdx];
      if (id) statuses[String(id)] = data[i][statusIdx] || '';
    }
    return { success: true, statuses: statuses };
  } catch (error) {
    return handleError(error, 'ステータス取得');
  }
}

/**
 * 個別生徒の認証コードを取得（教員/admin専用）
 * @param {string} studentId - 学籍番号
 * @returns {object} 認証コード
 */
function getAuthCode(studentId) {
  try {
    requireTeacherOrAdmin('認証コード取得');

    if (!studentId) {
      throw new Error('学籍番号が指定されていません');
    }

    const data = getSubmissionSheetDataCached();
    if (data.length === 0) {
      throw new Error('登録データが見つかりません');
    }

    const headers = data[0];
    const authCodeIndex = headers.indexOf('認証コード');
    const studentIdIndex = headers.indexOf('学籍番号');

    if (authCodeIndex === -1 || studentIdIndex === -1) {
      throw new Error('必要な列が見つかりません');
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][studentIdIndex]) === String(studentId)) {
        return {
          success: true,
          authCode: data[i][authCodeIndex] || ''
        };
      }
    }

    throw new Error('該当する生徒が見つかりません');

  } catch (error) {
    return handleError(error, '認証コード取得');
  }
}

/**
 * クラス一覧（学年・組のユニークペア）を取得
 * @returns {object} クラスリスト
 */
function getClassList() {
  try {
    requireTeacherOrAdmin('クラス一覧取得');
    if (!getSubmissionSheet()) {
      return { success: true, classList: [] };
    }

    const data = getSubmissionSheetDataCached();
    if (data.length === 0) {
      return { success: true, classList: [] };
    }

    const headers = data[0];
    if (!headers || headers.length === 0) {
      return { success: true, classList: [] };
    }

    const gradeIndex = headers.indexOf('学年');
    const classIndex = headers.indexOf('組');

    if (gradeIndex === -1 || classIndex === -1) {
      return { success: true, classList: [] };
    }

    const classSet = new Set();
    for (let i = 1; i < data.length; i++) {
      const grade = parseInt(data[i][gradeIndex]);
      const classVal = data[i][classIndex] != null ? String(data[i][classIndex]).trim() : '';
      if (!isNaN(grade) && classVal !== '') {
        classSet.add(grade + '-' + classVal);
      }
    }

    return {
      success: true,
      classList: Array.from(classSet),
      timestamp: getJSTTimestamp()
    };

  } catch (error) {
    return handleError(error, 'クラス一覧取得');
  }
}

/**
 * 生徒の履修状態を承認
 * @param {string} studentId - 学籍番号
 * @returns {object} 承認結果
 */

/**
 * 確認済の生徒を仮登録に戻す
 * @param {string} studentId - 学籍番号
 * @returns {object} 処理結果
 */
function revertToProvisional(studentId) {
  try {
    requireTeacherOrAdmin('承認取消');
    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const statusIndex = headers.indexOf('ステータス');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    // インデックスチェック
    if (studentIdIndex === -1) {
      throw new Error('ヘッダーに「学籍番号」列が見つかりません');
    }
    if (statusIndex === -1) {
      throw new Error('ヘッダーに「ステータス」列が見つかりません');
    }
    if (timestampIndex === -1) {
      throw new Error('ヘッダーに「タイムスタンプ」列が見つかりません');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][studentIdIndex] == studentId) {
        const currentStatus = data[i][statusIndex];
        if (currentStatus !== '確認済') {
          throw new Error('この生徒は確認済ではありません');
        }

        const targetRow = i + 1;
        // 最適化: キャッシュデータから直接更新
        const rowData = data[i].slice();
        rowData[statusIndex] = '仮登録';
        rowData[timestampIndex] = getJSTTimestamp();
        getSubmissionSheet().getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

        // キャッシュクリア
        try {
          clearSubmissionSheetCache();
        } catch (cacheError) {
          console.warn('登録データキャッシュクリアに失敗:', cacheError);
        }

        logMessage(`承認取消完了: 学籍番号 ${studentId} を仮登録に戻しました`);

        return {
          success: true,
          message: '承認を取り消して仮登録に戻しました',
          timestamp: getJSTTimestamp()
        };
      }
    }

    throw new Error('指定された学籍番号の生徒が見つかりません');

  } catch (error) {
    logMessage(`承認取消エラー: ${error.message}`, 'ERROR');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 科目名から科目データを検索
 * @param {string} courseName - 科目名
 * @param {object} courseData - 科目データ
 * @returns {object} 科目データ
 */
function findCourseByName(courseName, courseData) {
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
}

// ===============================
// 初期データ取得・生徒管理機能
// ===============================

// ===============================
// リトライ機能
// ===============================

/**
 * スプレッドシート読み込みのリトライ機能
 * @param {SpreadsheetApp.Sheet} sheet - 対象シート
 * @param {string} functionName - 関数名（ログ用）
 * @returns {Array} シートデータ
 */
function getDataWithRetry(sheet, functionName) {
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const data = sheet.getDataRange().getValues();
      return data;
    } catch (error) {
      const isLastAttempt = attempt === maxRetries;
      const delay = Math.pow(2, attempt - 1) * 1000; // 1秒、2秒、4秒

      console.warn(`[${functionName}] 読み込み失敗 (${attempt}/${maxRetries}): ${error.toString()}`);

      if (isLastAttempt) {
        console.error(`[${functionName}] 最大リトライ回数に達しました。エラーを再スロー`);
        throw error;
      }

      Utilities.sleep(delay);
    }
  }
}

// ===============================
// キャッシュ機能
// ===============================

/**
 * キャッシュをクリアする関数
 * データ更新時に呼び出してキャッシュを無効化
 */
/**
 * 登録データのキャッシュのみをクリア
 */
function clearSubmissionSheetCache() {
  _memCache.submission = null;
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('submission_sheet_data');
    return { success: true };
  } catch (error) {
    console.error('登録データキャッシュクリアエラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 全データキャッシュをクリア
 */
function clearDataCache() {
  _memCache = {};
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('student_roster_data');
    cache.remove('course_master_data');
    cache.remove('submission_sheet_data');
    cache.remove('teacher_data');
    cache.remove('settings_data');
    return { success: true };
  } catch (error) {
    console.error('キャッシュクリアエラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 設定データをキャッシュ対応で取得
 * @returns {Object} 設定データのオブジェクト（例: {学校名: '京都芸術大学附属高等学校　普通科'}）
 */
function getSettings() {
  try {
    const cache = CacheService.getScriptCache();

    // キャッシュから取得を試行
    const cached = cache.get('settings_data');
    if (cached) {
      return JSON.parse(cached);
    }

    // キャッシュがない場合はシートから読み込み
    const sheet = getSettingsSheet();
    const data = sheet.getDataRange().getValues();

    // ヘッダー行をスキップして設定をオブジェクトに変換
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      const settingName = data[i][0];
      const settingValue = data[i][1];
      if (settingName) {
        settings[settingName] = settingValue;
      }
    }

    // キャッシュに保存
    cache.put('settings_data', JSON.stringify(settings), CACHE_EXPIRY_MS / 1000);
    return settings;
  } catch (error) {
    console.error('設定データ取得エラー:', error);
    // エラー時はデフォルト値を返す
    return {
      '学校名': '京都芸術大学附属高等学校　普通科'
    };
  }
}

/**
 * 設定を更新
 * @param {object} settingsData - 更新する設定データ（キー: 設定名、値: 内容）
 * @returns {object} 成功/失敗の結果
 */
function updateSettings(settingsData) {
  try {
    // 管理者権限チェック
    const userInfo = getUserInfo();
    if (!userInfo.success || userInfo.user.role !== 'admin') {
      return {
        success: false,
        error: '管理者権限が必要です'
      };
    }

    const sheet = getSettingsSheet();
    const data = sheet.getDataRange().getValues();

    // 各設定項目を更新
    for (let i = 1; i < data.length; i++) {
      const settingName = data[i][0];
      if (settingName && settingsData.hasOwnProperty(settingName)) {
        sheet.getRange(i + 1, 2).setValue(settingsData[settingName]);
      }
    }

    // 新規設定項目の追加（シートに存在しない設定を追記）
    const existingKeys = new Set(data.map(row => row[0]));
    Object.keys(settingsData).forEach(key => {
        if (!existingKeys.has(key)) {
            sheet.appendRow([key, settingsData[key]]);
        }
    });

    // キャッシュをクリア
    const cache = CacheService.getScriptCache();
    cache.remove('settings_data');
    return {
      success: true,
      message: '設定を更新しました'
    };

  } catch (error) {
    console.error('設定更新エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 登録データシートのデータをキャッシュ対応で取得
 * @returns {Array} 登録データシートの全データ（ヘッダー含む）
 */
function getSubmissionSheetDataCached() {
  if (_memCache.submission) return _memCache.submission;
  try {
    const cache = CacheService.getScriptCache();

    // キャッシュから取得を試行
    let cachedData = cache.get('submission_sheet_data');
    if (cachedData) {
      const parsed = JSON.parse(cachedData);
      _memCache.submission = parsed;
      return parsed;
    }

    // キャッシュにない場合は直接シートから取得
    if (!getSubmissionSheet()) {
      console.warn('登録データシートが見つかりません');
      return [];
    }

    const submissionData = getDataWithRetry(getSubmissionSheet(), 'getSubmissionSheetDataCached');

    // 5分間キャッシュに保存（データが大きすぎる場合はスキップ）
    try {
      const jsonData = JSON.stringify(submissionData);
      // CacheServiceの制限は約100KB
      if (jsonData.length < 100000) {
        cache.put('submission_sheet_data', jsonData, 300);
      }
    } catch (cacheError) {
    }

    _memCache.submission = submissionData;
    return submissionData;

  } catch (error) {
    console.error('登録データ取得エラー:', error);
    // エラー時は直接シートアクセスにフォールバック
    const submissionSheet = getSubmissionSheet();
    if (submissionSheet) {
      try {
        return getDataWithRetry(submissionSheet, 'getSubmissionSheetDataCached-fallback');
      } catch (fallbackError) {
        console.error('フォールバック処理も失敗:', fallbackError);
        return [];
      }
    }
    return [];
  }
}

/**
 * 科目マスタデータをキャッシュ対応で取得
 * @returns {Array} 科目マスタシートの全データ（ヘッダー含む）
 */
function getCourseDataCached() {
  if (_memCache.course) return _memCache.course;
  try {
    const cache = CacheService.getScriptCache();

    // キャッシュから取得を試行
    let cachedData = cache.get('course_master_data');
    if (cachedData) {
      const parsed = JSON.parse(cachedData);
      _memCache.course = parsed;
      return parsed;
    }

    // キャッシュにない場合は直接シートから取得
    if (!getCourseSheet()) {
      console.warn('科目マスタシートが見つかりません');
      return [];
    }

    const courseData = getDataWithRetry(getCourseSheet(), 'getCourseDataCached');

    // 1時間キャッシュに保存
    cache.put('course_master_data', JSON.stringify(courseData), 3600);
    _memCache.course = courseData;
    return courseData;

  } catch (error) {
    console.error('科目マスタ取得エラー:', error);
    // エラー時は直接シートアクセスにフォールバック
    const courseSheet = getCourseSheet();
    if (courseSheet) {
      try {
        return getDataWithRetry(courseSheet, 'getCourseDataCached-fallback');
      } catch (fallbackError) {
        console.error('科目マスタフォールバック処理も失敗:', fallbackError);
        return [];
      }
    }
    return [];
  }
}

/**
 * 教職員データをキャッシュ付きで取得
 * @returns {Array} 教職員データの配列
 */
function getTeacherDataCached() {
  if (_memCache.teacher) return _memCache.teacher;
  try {
    const cache = CacheService.getScriptCache();
    let cachedData = cache.get('teacher_data');

    if (cachedData) {
      const parsed = JSON.parse(cachedData);
      _memCache.teacher = parsed;
      return parsed;
    }

    if (!getTeacherSheet()) {
      return [];
    }

    const teacherData = getDataWithRetry(getTeacherSheet(), 'getTeacherDataCached');

    // 30分（1800秒）キャッシュに保存
    cache.put('teacher_data', JSON.stringify(teacherData), 1800);
    _memCache.teacher = teacherData;
    return teacherData;

  } catch (error) {
    console.error('教職員データ取得エラー:', error);
    // エラー時は直接シートアクセスにフォールバック
    const teacherSheet = getTeacherSheet();
    if (teacherSheet) {
      try {
        return getDataWithRetry(teacherSheet, 'getTeacherDataCached-fallback');
      } catch (fallbackError) {
        console.error('教職員データフォールバック処理も失敗:', fallbackError);
        return [];
      }
    }
    return [];
  }
}

/**
 * 教職員メールアドレス一覧を取得
 * @returns {string[]} メールアドレスの配列
 */
function getTeacherEmails() {
  requireTeacherOrAdmin('教職員メール取得');
  const data = getTeacherDataCached();
  if (data.length < 2) return [];
  const headers = data[0];
  const emailIndex = headers.indexOf('メールアドレス');
  if (emailIndex === -1) return [];
  const emails = [];
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][emailIndex] || '').trim().toLowerCase();
    if (email) emails.push(email);
  }
  return emails;
}

/**
 * 生徒名簿データをキャッシュ付きで取得
 * @returns {Array} 生徒名簿データの配列
 */
function getStudentRosterDataCached() {
  if (_memCache.roster) return _memCache.roster;
  try {
    const cache = CacheService.getScriptCache();
    let cachedData = cache.get('student_roster_data');

    if (cachedData) {
      const parsed = JSON.parse(cachedData);
      _memCache.roster = parsed;
      return parsed;
    }

    if (!getStudentRosterSheet()) {
      return [];
    }

    const rosterData = getDataWithRetry(getStudentRosterSheet(), 'getStudentRosterDataCached');

    // 30分（1800秒）キャッシュに保存
    cache.put('student_roster_data', JSON.stringify(rosterData), 1800);
    _memCache.roster = rosterData;
    return rosterData;

  } catch (error) {
    console.error('生徒名簿データ取得エラー:', error);
    // エラー時は直接シートアクセスにフォールバック
    const studentRosterSheet = getStudentRosterSheet();
    if (studentRosterSheet) {
      try {
        return getDataWithRetry(studentRosterSheet, 'getStudentRosterDataCached-fallback');
      } catch (fallbackError) {
        console.error('生徒名簿データフォールバック処理も失敗:', fallbackError);
        return [];
      }
    }
    return [];
  }
}

/**
 * メンテナンスモード・機能設定のサーバーサイドチェック
 * @param {string} userRole - ユーザーの権限
 * @param {string} featureKey - チェックする機能設定キー（'仮登録', '一時保存', '本登録許可'、nullでスキップ）
 * @throws {Error} ブロック時
 */
function checkServerSideRestrictions(userRole, featureKey) {
  const settings = getSettings();

  // メンテナンスモードチェック（getInitialDataと同一ロジック）
  const maintenanceMode = settings['メンテナンスモード'] || '無効';
  const isBlocked =
      (maintenanceMode === '管理者のみ' || maintenanceMode === '有効') && userRole !== 'admin' ||
      maintenanceMode === '教員以上' && !isTeacherOrAbove(userRole);
  if (isBlocked) {
    throw new Error('メンテナンスモード中のため操作できません');
  }

  // 機能設定チェック（管理者は免除）
  if (featureKey && userRole !== 'admin') {
    const featureValue = settings[featureKey];
    if (featureValue && featureValue !== '可') {
      throw new Error('現在' + featureKey + 'は許可されていません');
    }
  }
}

/**
 * 初期データを取得（キャッシュ対応版 - 行データをヘッダー付きで返す）
 * @returns {object} 生データ（ヘッダー付き行データ）
 */
function getInitialData() {
  try {
    // キャッシュサービスの初期化
    const cache = CacheService.getScriptCache();

    // シートの存在確認
    if (!validateSheets()) {
      throw new Error('必要なシートが見つかりません。シート名を確認してください。');
    }

    // 【高速化】全キャッシュを事前に準備（後続の処理で個別読み込みを防ぐ）
    getTeacherDataCached();
    getSubmissionSheetDataCached();
    getCourseDataCached();
    getStudentRosterDataCached();
    
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error("メールアドレスが取得できません。Googleアカウントでログインしてください。");
    }
    
    // 1. 権限判定（ロール+名前を一括取得）
    const userRoleResult = determineUserRole(userEmail);
    const userRole = userRoleResult ? userRoleResult.role : 'student';

    // メンテナンスモード判定: 3段階（無効/教員以上/管理者のみ）
    // 後方互換性: 旧値'有効'は'管理者のみ'として扱う
    const settings = getSettings();
    const maintenanceMode = settings['メンテナンスモード'] || '無効';
    const isBlocked =
        (maintenanceMode === '管理者のみ' || maintenanceMode === '有効') && userRole !== 'admin' ||
        maintenanceMode === '教員以上' && !isTeacherOrAbove(userRole);
    if (isBlocked) {
        return {
            maintenanceMode: true,
            settings: settings,
            gasUserEmail: userEmail
        };
    }

    // 2. 登録データから学籍番号と学生データを同時取得（一括処理で高速化）
    let studentId;
    let studentSubmissionData = null;

    {
      // 登録データシートから一度の取得でメールアドレス該当行を検索
      const submissionAllData = getSubmissionSheetDataCached();

      if (submissionAllData.length > 1) {
        const submissionHeader = submissionAllData[0];
        const emailIndex = submissionHeader.indexOf('メールアドレス');
        const studentIdIndex = submissionHeader.indexOf('学籍番号');

        if (emailIndex !== -1 && studentIdIndex !== -1) {
          // メールアドレスで該当行を検索して学籍番号を取得
          const submissionRow = submissionAllData.find(row => row[emailIndex] === userEmail);

          if (submissionRow) {
            studentId = submissionRow[studentIdIndex];
            studentSubmissionData = {
              header: submissionHeader,
              row: submissionRow
            };
          } else {
          }
        } else {
          console.warn(`必要な列が見つかりません: メールアドレス列=${emailIndex}, 学籍番号列=${studentIdIndex}`);
        }
      }

      // 学籍番号が取得できない場合の教員判定処理
      if (!studentId) {
        if (isTeacherOrAbove(userRole)) {
          // determineUserRoleの結果を再利用（重複走査を回避）
          const teacherData = userRoleResult && userRoleResult.name
            ? { name: userRoleResult.name, クラス: userRoleResult.assignedClasses }
            : { name: '教職員' };
          // 科目データを取得（教員画面・管理画面で必要）
          const courseAllData = getCourseDataCached();
          const teacherCourseData = courseAllData.length > 1
            ? { header: courseAllData[0], rows: courseAllData.slice(1) }
            : { header: [], rows: [] };

          const teacherResult = {
            userRole: userRole,
            teacherData: teacherData,
            isTeacher: true,
            gasUserEmail: userEmail,
            message: '教員としてログインしました。',
            settings: getSettings(),
            webAppUrl: ScriptApp.getService().getUrl(),
            courseData: teacherCourseData
          };

          return teacherResult;
        }
        throw new Error(`認証エラー: メールアドレス「${userEmail}」が登録データシートに見つかりません。`);
      }
    }

    // 利用停止チェック（全ロール共通）
    if (studentSubmissionData) {
      const suspendedStatusIdx = studentSubmissionData.header.indexOf('ステータス');
      if (suspendedStatusIdx !== -1 && studentSubmissionData.row[suspendedStatusIdx] === '利用停止') {
        return { suspended: true, settings: settings, gasUserEmail: userEmail, userRole: userRole };
      }
    }

    // 5. 科目マスタ全体を取得（ヘッダー付き）
    let courseData;
    if (getCourseSheet()) {
      // キャッシュから科目データを取得
      const courseAllData = getCourseDataCached();
      courseData = {
        header: courseAllData[0],
        rows: courseAllData.slice(1)
      };
    }
    // 提出状態をチェック（内部関数使用：認証済みユーザーの自身のstudentId）
    const submissionStatus = checkSubmissionStatusInternal(studentId);
    
    // 結果を返す（登録データから統合取得）
    // teacherData取得（determineUserRoleの結果を再利用）
    let teacherData;
    if (isTeacherOrAbove(userRole)) {
      teacherData = userRoleResult && userRoleResult.name
        ? { name: userRoleResult.name, クラス: userRoleResult.assignedClasses }
        : { name: '教職員' };
    } else {
      teacherData = { name: '教職員' };
    }

    // studentSubmissionDataをプリミティブ値に変換（JSON化可能にする）
    let serializedStudentData = null;
    let nextGradeOverride = null;
    if (studentSubmissionData) {
      serializedStudentData = {
        header: studentSubmissionData.header ? Array.from(studentSubmissionData.header).map(v => v != null ? String(v) : '') : [],
        row: studentSubmissionData.row ? Array.from(studentSubmissionData.row).map(v => v != null ? String(v) : '') : []
      };
      // 来年度学年列の値を抽出
      const nextGradeColIdx = studentSubmissionData.header.indexOf('来年度学年');
      if (nextGradeColIdx !== -1) {
        const val = parseInt(studentSubmissionData.row[nextGradeColIdx]);
        if (!isNaN(val) && val >= 1 && val <= 3) {
          nextGradeOverride = val;
        }
      }
    }

    // クラス非表示モード: トップページでは組・番号を返さない
    if (shouldHideClass() && serializedStudentData) {
      serializedStudentData = stripClassInfoFromHeaderRow(serializedStudentData);
    }

    const result = {
      userRole: userRole,
      studentId: studentId,
      gasUserEmail: userEmail,
      studentRosterData: serializedStudentData, // 登録データシートから学生基本情報も取得
      studentSubmissionData: serializedStudentData,
      courseData: courseData,
      submissionStatus: submissionStatus,
      nextGradeOverride: nextGradeOverride,
      currentYear: new Date().getFullYear(),
      isTeacher: isTeacherOrAbove(userRole),
      teacherData: teacherData,
      message: isTeacherOrAbove(userRole) ? '教員+生徒データ両方読み込み完了（直接認証版）' : '直接認証完了',
      settings: getSettings(),
      webAppUrl: ScriptApp.getService().getUrl()
    };

    return result;

  } catch (e) {
    console.error('=== getInitialData エラー発生 ===');
    console.error('エラーメッセージ:', e.message);
    console.error('エラースタック:', e.stack);
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
}


/**
 * 提出済み生徒データを取得（教員用）
 * @returns {object} 提出済み生徒データ
 */
function getSubmittedStudents() {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員権限が必要です');
    }

    if (!getSubmissionSheet()) {
      return { success: true, students: [] };
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const statusIndex = headers.indexOf('ステータス');

    // 【高速化】バッチ処理用Mapを事前構築（O(n²)→O(n)）
    const studentMap = buildStudentLookupMap();

    const submittedStudents = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentId = row[studentIdIndex];
      const status = row[statusIndex];

      if (status === '仮登録') {
        const studentData = studentMap.get(String(studentId));
        if (studentData) {
          submittedStudents.push({
            studentId: studentId,
            name: studentData.name,
            grade: studentData.grade,
            class: studentData.class,
            number: studentData.number,
            status: status
          });
        }
      }
    }

    return {
      success: true,
      students: submittedStudents,
      timestamp: getJSTTimestamp()
    };

  } catch (error) {
    return handleError(error, '提出済み生徒データ取得');
  }
}

/**
 * 履修登録を承認
 * @param {string} studentId - 学籍番号
 * @returns {object} 承認結果
 */
function approveRegistration(studentId) {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員権限が必要です');
    }
    checkServerSideRestrictions(userInfo.user.role, null);

    return approveStudent(studentId);

  } catch (error) {
    return handleError(error, '履修登録承認');
  }
}

/**
 * 生徒の履修登録を差戻し
 * @param {string} studentId - 学籍番号
 * @returns {object} 差戻し結果
 */
function revertSubmission(studentId) {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員権限が必要です');
    }
    checkServerSideRestrictions(userInfo.user.role, null);

    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const statusIndex = headers.indexOf('ステータス');
    const authCodeIndex = headers.indexOf('認証コード');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    // インデックスチェック
    if (studentIdIndex === -1) {
      throw new Error('ヘッダーに「学籍番号」列が見つかりません');
    }
    if (statusIndex === -1) {
      throw new Error('ヘッダーに「ステータス」列が見つかりません');
    }
    if (authCodeIndex === -1) {
      throw new Error('ヘッダーに「認証コード」列が見つかりません');
    }
    if (timestampIndex === -1) {
      throw new Error('ヘッダーに「タイムスタンプ」列が見つかりません');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][studentIdIndex] == studentId) {
        const targetRow = i + 1;
        // 最適化: キャッシュデータから直接更新
        const rowData = data[i].slice();
        rowData[statusIndex] = '差戻し';
        rowData[authCodeIndex] = '';
        rowData[timestampIndex] = getJSTTimestamp();
        getSubmissionSheet().getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

        // キャッシュクリア
        try {
          clearSubmissionSheetCache();
        } catch (cacheError) {
          console.warn('登録データキャッシュクリアに失敗:', cacheError);
        }

        logMessage(`履修登録差戻し完了: 学籍番号 ${studentId}`);
        
        return {
          success: true,
          message: '履修登録を差戻しました',
          timestamp: getJSTTimestamp()
        };
      }
    }
    
    throw new Error('指定された学籍番号の生徒が見つかりません');

  } catch (error) {
    return handleError(error, '履修登録差戻し');
  }
}

/**
 * 管理者が自分の利用停止ステータスを解除する
 * getUserInfo() は利用停止ユーザーでは循環するため、determineUserRole() を直接使用
 * @returns {object} 解除結果
 */
function clearOwnSuspension() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      throw new Error('メールアドレスが取得できません');
    }

    // 教職員データシートから直接ロール判定（getUserInfo経由しない）
    const roleResult = determineUserRole(email);
    if (!roleResult || roleResult.role !== 'admin') {
      throw new Error('管理者権限が必要です');
    }

    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const emailIndex = headers.indexOf('メールアドレス');
    const statusIndex = headers.indexOf('ステータス');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    if (emailIndex === -1) throw new Error('ヘッダーに「メールアドレス」列が見つかりません');
    if (statusIndex === -1) throw new Error('ヘッダーに「ステータス」列が見つかりません');
    if (timestampIndex === -1) throw new Error('ヘッダーに「タイムスタンプ」列が見つかりません');

    const normalizedEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailIndex] || '').trim().toLowerCase();
      if (rowEmail === normalizedEmail) {
        if (data[i][statusIndex] !== '利用停止') {
          throw new Error('現在のステータスは利用停止ではありません');
        }

        const targetRow = i + 1;
        const rowData = data[i].slice();
        rowData[statusIndex] = '';
        rowData[timestampIndex] = getJSTTimestamp();
        getSubmissionSheet().getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

        // キャッシュクリア
        try {
          clearSubmissionSheetCache();
        } catch (cacheError) {
          console.warn('登録データキャッシュクリアに失敗:', cacheError);
        }

        logMessage(`管理者利用停止解除: ${email}`);

        return {
          success: true,
          message: '利用停止を解除しました',
          timestamp: getJSTTimestamp()
        };
      }
    }

    throw new Error('登録データシートにアカウント情報が見つかりません');

  } catch (error) {
    return handleError(error, '利用停止解除');
  }
}

/**
 * 生徒の本登録許可を設定
 * @param {string} studentId - 学籍番号
 * @returns {object} 設定結果
 */
function setFinalApproval(studentId, authCode) {
  try {
    const userInfo = getUserInfo();

    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員以上の権限が必要です');
    }
    checkServerSideRestrictions(userInfo.user.role, '本登録許可');

    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const authCodeIndex = headers.indexOf('認証コード');
    const statusIndex = headers.indexOf('ステータス');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    // インデックスチェック
    if (studentIdIndex === -1) {
      throw new Error('ヘッダーに「学籍番号」列が見つかりません');
    }
    if (authCodeIndex === -1) {
      throw new Error('ヘッダーに「認証コード」列が見つかりません');
    }
    if (statusIndex === -1) {
      throw new Error('ヘッダーに「ステータス」列が見つかりません');
    }
    if (timestampIndex === -1) {
      throw new Error('ヘッダーに「タイムスタンプ」列が見つかりません');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][studentIdIndex] == studentId) {
        // 既に認証コードが設定されている場合はエラー
        if (data[i][authCodeIndex]) {
          throw new Error('既に本登録許可されています');
        }

        const targetRow = i + 1;
        // 最適化: キャッシュデータから直接更新
        const rowData = data[i].slice();
        rowData[authCodeIndex] = authCode;
        rowData[statusIndex] = '本登録許可';
        rowData[timestampIndex] = getJSTTimestamp();
        getSubmissionSheet().getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

        // キャッシュクリア
        try {
          clearSubmissionSheetCache();
        } catch (cacheError) {
          console.warn('登録データキャッシュクリアに失敗:', cacheError);
        }

        logMessage(`本登録許可設定完了: 学籍番号 ${studentId}, 認証コード ${authCode}`);

        return {
          success: true,
          message: '本登録許可を設定しました',
          authCode: authCode,
          timestamp: getJSTTimestamp()
        };
      }
    }

    throw new Error('指定された学籍番号の生徒が見つかりません');

  } catch (error) {
    return handleError(error, '本登録許可設定');
  }
}

/**
 * 本登録許可を取り消す
 * @param {string} studentId - 学籍番号
 * @returns {object} 処理結果
 */
function cancelFinalApproval(studentId) {
  try {
    const userInfo = getUserInfo();

    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員以上の権限が必要です');
    }

    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const studentIdIndex = headers.indexOf('学籍番号');
    const authCodeIndex = headers.indexOf('認証コード');
    const statusIndex = headers.indexOf('ステータス');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    // インデックスチェック
    if (studentIdIndex === -1) {
      throw new Error('ヘッダーに「学籍番号」列が見つかりません');
    }
    if (authCodeIndex === -1) {
      throw new Error('ヘッダーに「認証コード」列が見つかりません');
    }
    if (statusIndex === -1) {
      throw new Error('ヘッダーに「ステータス」列が見つかりません');
    }
    if (timestampIndex === -1) {
      throw new Error('ヘッダーに「タイムスタンプ」列が見つかりません');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][studentIdIndex] == studentId) {
        const targetRow = i + 1;
        // 最適化: キャッシュデータから直接更新
        const rowData = data[i].slice();
        rowData[authCodeIndex] = '';
        rowData[statusIndex] = '仮登録';
        rowData[timestampIndex] = getJSTTimestamp();
        getSubmissionSheet().getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

        // キャッシュクリア
        try {
          clearSubmissionSheetCache();
        } catch (cacheError) {
          console.warn('登録データキャッシュクリアに失敗:', cacheError);
        }

        logMessage(`本登録許可取消完了: 学籍番号 ${studentId}`);

        return {
          success: true,
          message: '本登録許可を取り消しました',
          timestamp: getJSTTimestamp()
        };
      }
    }

    throw new Error('指定された学籍番号の生徒が見つかりません');

  } catch (error) {
    return handleError(error, '本登録許可取消');
  }
}

/**
 * 認証コードを検証して学生データを取得
 * @param {string} authCode - 認証コード
 * @returns {object} 検証結果と学生データ
 */
function validateAuthCode(authCode) {
  try {
    if (!authCode) {
      throw new Error('認証コードが指定されていません');
    }

    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const authCodeIndex = headers.indexOf('認証コード');
    const studentIdIndex = headers.indexOf('学籍番号');
    const statusIndex = headers.indexOf('ステータス');

    // インデックスチェック
    if (authCodeIndex === -1) {
      throw new Error('ヘッダーに「認証コード」列が見つかりません');
    }
    if (studentIdIndex === -1) {
      throw new Error('ヘッダーに「学籍番号」列が見つかりません');
    }
    if (statusIndex === -1) {
      throw new Error('ヘッダーに「ステータス」列が見つかりません');
    }

    // 認証コードで検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][authCodeIndex] === authCode) {
        const studentId = data[i][studentIdIndex];
        const status = data[i][statusIndex];

        // セッション検証: 生徒本人 OR 教員/adminであることを検証
        const callerInfo = requireSelfOrTeacher(studentId, '認証コード検証');

        // 学生データ取得（内部関数使用：認証コード検証済みのためスキップ）
        let studentData = getStudentDataByIdInternal(studentId);
        let registrationData = loadRegistrationDataInternal(studentId);

        // クラス非表示モード: トップページでは組・番号を返さない
        if (shouldHideClass()) {
          studentData = stripClassInfoFromStudentData(studentData);
          registrationData = stripClassInfoFromHeaderRow(registrationData);
        }

        logMessage(`認証コード検証成功: ${authCode}, 学籍番号: ${studentId}`);

        return {
          success: true,
          studentId: studentId,
          studentData: studentData,
          registrationData: registrationData,
          status: status
        };
      }
    }

    throw new Error('無効な認証コードです');

  } catch (error) {
    return handleError(error, '認証コード検証');
  }
}

/**
 * 本登録を実行
 * @param {string} authCode - 認証コード
 * @returns {object} 実行結果
 */
function finalRegister(authCode) {
  try {
    if (!authCode) {
      throw new Error('認証コードが指定されていません');
    }

    if (!getSubmissionSheet()) {
      throw new Error('登録データシートが見つかりません');
    }

    const data = getSubmissionSheetDataCached();
    const headers = data[0];

    const authCodeIndex = headers.indexOf('認証コード');
    const studentIdIndex = headers.indexOf('学籍番号');
    const statusIndex = headers.indexOf('ステータス');
    const timestampIndex = headers.indexOf('タイムスタンプ');

    // インデックスチェック
    if (authCodeIndex === -1) {
      throw new Error('ヘッダーに「認証コード」列が見つかりません');
    }
    if (studentIdIndex === -1) {
      throw new Error('ヘッダーに「学籍番号」列が見つかりません');
    }
    if (statusIndex === -1) {
      throw new Error('ヘッダーに「ステータス」列が見つかりません');
    }
    if (timestampIndex === -1) {
      throw new Error('ヘッダーに「タイムスタンプ」列が見つかりません');
    }

    // 認証コードで検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][authCodeIndex] === authCode) {
        const studentId = data[i][studentIdIndex];

        // セッション検証: 生徒本人 OR 教員/adminであることを検証
        requireSelfOrTeacher(studentId, '本登録実行');

        // メンテナンスモード・機能設定チェック
        const callerInfo = getUserInfo();
        if (callerInfo.success) {
          checkServerSideRestrictions(callerInfo.user.role, '本登録許可');
        }

        const targetRow = i + 1;

        // 行データをコピーして更新内容を反映
        const rowData = data[i].slice();
        const currentTimestamp = getJSTTimestamp();
        rowData[statusIndex] = '本登録';
        rowData[timestampIndex] = currentTimestamp;
        rowData[authCodeIndex] = ''; // 認証コード削除（セキュリティ対策）

        // 仮登録シートを更新
        getSubmissionSheet().getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

        // キャッシュクリア
        try {
          clearSubmissionSheetCache();
        } catch (cacheError) {
          console.warn('登録データキャッシュクリアに失敗:', cacheError);
        }

        logMessage(`本登録完了: 学籍番号 ${studentId}, 認証コード ${authCode}`);

        return {
          success: true,
          message: '本登録が完了しました',
          studentId: studentId,
          timestamp: getJSTTimestamp()
        };
      }
    }

    throw new Error('無効な認証コードです');

  } catch (error) {
    return handleError(error, '本登録実行');
  }
}

/**
 * 教職員データからテーチャー名を取得
 * @param {string} email - メールアドレス
 * @returns {string} 教員名
 */
function getTeacherData(email) {
  try {
    if (!getTeacherSheet()) {
      return null;
    }

    const data = getTeacherDataCached();
    const headers = data[0];

    const emailIndex = headers.indexOf('メールアドレス');
    // 「氏名」または「名前」のどちらかを探す
    let nameIndex = headers.indexOf('氏名');
    if (nameIndex === -1) {
      nameIndex = headers.indexOf('名前');
    }
    const gradeIndex = headers.indexOf('学年');
    const classIndex = headers.indexOf('組');

    if (emailIndex === -1 || nameIndex === -1) {
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sheetEmail = row[emailIndex];
      // 大文字小文字を無視して比較、前後の空白も除去
      if (String(sheetEmail).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        return {
          name: row[nameIndex],
          学年: gradeIndex !== -1 ? row[gradeIndex] : undefined,
          組: classIndex !== -1 ? row[classIndex] : undefined
        };
      }
    }

    return null;

  } catch (error) {
    logMessage(`教員データ取得エラー: ${error.toString()}`, 'ERROR');
    return null;
  }
}

// 下位互換性のため残す
function getTeacherName(email) {
  const teacherData = getTeacherData(email);
  return teacherData ? teacherData.name : null;
}

/**
 * 教員用生徒データ取得（学籍番号ベース）
 * @param {string} studentId - 学籍番号
 * @returns {object} 生徒データ
 */
function getStudentDataForView(studentId) {
  try {
    const userInfo = getUserInfo();
    if (!userInfo.success || !isTeacherOrAbove(userInfo.user.role)) {
      throw new Error('教員権限が必要です');
    }

    const studentData = getStudentDataByIdInternal(studentId);
    if (!studentData) {
      throw new Error('生徒データが見つかりません');
    }

    const registrationData = loadRegistrationDataInternal(studentId);

    // 履修済み科目も登録データから取得（統合済み）
    const completedCourses = [];
    const retakenCourses = [];

    if (registrationData.success && registrationData.data) {
      // 登録データから●印（履修済み）と年度マーク（再履修）を抽出
      Object.keys(registrationData.data).forEach(courseName => {
        const mark = registrationData.data[courseName];
        if (mark === '●' || /^●[123]$/.test(mark)) {
          completedCourses.push(courseName);
        } else if (mark === '2' || mark === '3' || mark === '○2' || mark === '○3' || mark === '○2N' || mark === '○3N') {
          retakenCourses.push(courseName);
        }
      });
    }

    return {
      success: true,
      student: studentData,
      registrationData: registrationData.data,
      completedCourses: completedCourses,
      retakenCourses: retakenCourses,
      timestamp: getJSTTimestamp()
    };

  } catch (error) {
    return handleError(error, '教員用生徒データ取得');
  }
}

/**
 * テスト結果Excelデータを仮登録シートに取り込み
 * @param {object} importData - { 学籍番号: { 科目名: 学年(数値), ... }, ... }
 * @returns {object} 処理結果サマリー
 */
function importTestResults(importData) {
  try {
    // 教務権限チェック
    const userInfo = getUserInfo();
    if (!userInfo.success || userInfo.user.role !== 'admin') {
      throw new Error('教務権限が必要です');
    }

    if (!importData || typeof importData !== 'object') {
      throw new Error('取込データが無効です');
    }

    const submissionSheet = getSubmissionSheet();
    if (!submissionSheet) {
      throw new Error('登録データシートが見つかりません');
    }

    // 全データを一括読み込み（キャッシュではなくシートから直接読む）
    const allData = submissionSheet.getDataRange().getValues();
    const headers = allData[0];
    const lastCol = headers.length;

    const studentIdColIdx = headers.indexOf('学籍番号');
    if (studentIdColIdx === -1) {
      throw new Error('登録データシートに学籍番号列が見つかりません');
    }

    // 基本列（科目以外の列）
    const basicColumns = ['学籍番号', 'ステータス', 'タイムスタンプ', '学年', '組', '番号', '名前', 'メールアドレス', '教職員チェック', '認証コード', '来年度学年'];

    // 学籍番号→allDataインデックスのマッピング構築
    const studentRowMap = {};
    for (let i = 1; i < allData.length; i++) {
      const sid = String(allData[i][studentIdColIdx]).trim();
      if (sid) {
        studentRowMap[sid] = i; // allDataの0ベースインデックス
      }
    }

    const studentIds = Object.keys(importData);
    let updatedStudents = 0;
    let updatedCells = 0;
    let skippedCells = 0;
    const notFoundStudents = [];
    const updatedRowIndices = new Set();

    studentIds.forEach(studentId => {
      const sid = String(studentId).trim();
      let dataIndex = studentRowMap[sid];

      if (dataIndex === undefined) {
        // 登録データシートにない学籍番号はスキップ
        notFoundStudents.push(sid);
        return;
      }

      // メモリ上のデータを直接参照
      const rowData = allData[dataIndex];
      let rowUpdated = false;

      const subjects = importData[studentId];
      Object.keys(subjects).forEach(subjectName => {
        const newGrade = parseInt(subjects[subjectName]);
        if (!newGrade || newGrade < 1 || newGrade > 3) return;

        const colIdx = headers.indexOf(subjectName);
        if (colIdx === -1 || basicColumns.includes(subjectName)) return;

        const existingMark = String(rowData[colIdx] || '').trim();

        // 上書き保護: 既存マークの学年を抽出
        let existingGrade = 0;
        if (existingMark.includes('●')) {
          const match = existingMark.match(/●(\d)/);
          if (match) {
            existingGrade = parseInt(match[1]);
          } else if (existingMark === '●') {
            // ●のみの場合は学年不明 → 上書きしない（安全側）
            existingGrade = 99;
          }
        }

        if (existingGrade >= newGrade) {
          skippedCells++;
          return;
        }

        // ●{学年} で書き込み（allDataを直接更新）
        rowData[colIdx] = '●' + newGrade;
        rowUpdated = true;
        updatedCells++;
      });

      if (rowUpdated) {
        updatedStudents++;
        updatedRowIndices.add(dataIndex);
      }
    });

    // 更新行のみ書き戻し（タイムスタンプ等の未変更行への影響を防止）
    updatedRowIndices.forEach(function(rowIndex) {
      submissionSheet.getRange(rowIndex + 1, 1, 1, lastCol).setValues([allData[rowIndex]]);
    });

    // キャッシュクリア
    try {
      clearSubmissionSheetCache();
    } catch (cacheError) {
      console.warn('キャッシュクリア失敗:', cacheError);
    }

    const result = {
      success: true,
      updatedStudents: updatedStudents,
      updatedCells: updatedCells,
      skippedCells: skippedCells,
      notFoundStudents: notFoundStudents,
      timestamp: getJSTTimestamp()
    };

    logMessage('テストデータ取込完了: ' + updatedStudents + '名, ' + updatedCells + '件更新');

    return result;

  } catch (error) {
    return handleError(error, 'テストデータ取込');
  }
}

/**
 * 管理者向け: 学籍移動による生徒行の物理削除
 * @param {string[]} deleteStudentIds - 削除する学籍番号の配列
 * @returns {object} 処理結果
 */
function deleteStudentRows(deleteStudentIds) {
  try {
    // 管理者権限チェック
    const userInfo = getUserInfo();
    if (!userInfo.success || userInfo.user.role !== 'admin') {
      throw new Error('管理者権限が必要です');
    }

    if (!deleteStudentIds || !Array.isArray(deleteStudentIds) || deleteStudentIds.length === 0) {
      throw new Error('削除対象の学籍番号が指定されていません');
    }

    const submissionSheet = getSubmissionSheet();
    if (!submissionSheet) {
      throw new Error('登録データシートが見つかりません');
    }

    const allData = submissionSheet.getDataRange().getValues();
    const headers = allData[0];
    const studentIdColIdx = headers.indexOf('学籍番号');

    if (studentIdColIdx === -1) {
      throw new Error('登録データシートに学籍番号列が見つかりません');
    }

    // 学籍番号→行インデックスのマッピング構築
    const studentRowMap = {};
    for (let i = 1; i < allData.length; i++) {
      const sid = String(allData[i][studentIdColIdx]).trim();
      if (sid) studentRowMap[sid] = i;
    }

    // 削除対象の行番号を収集
    const rowsToDelete = [];
    deleteStudentIds.forEach(function(sid) {
      var idx = studentRowMap[String(sid).trim()];
      if (idx !== undefined) {
        rowsToDelete.push(idx + 1); // シートの1ベース行番号
      }
    });

    // 下から削除（行番号ずれ防止）
    rowsToDelete.sort(function(a, b) { return b - a; });
    rowsToDelete.forEach(function(sheetRow) {
      submissionSheet.deleteRow(sheetRow);
    });

    // キャッシュクリア
    try {
      clearSubmissionSheetCache();
    } catch (cacheError) {
      console.warn('キャッシュクリア失敗:', cacheError);
    }

    const deletedStudents = rowsToDelete.length;
    logMessage('生徒行削除完了: ' + deletedStudents + '名削除（学籍移動）');

    return {
      success: true,
      deletedStudents: deletedStudents,
      timestamp: getJSTTimestamp()
    };
  } catch (error) {
    return handleError(error, '生徒行削除');
  }
}

/**
 * 管理者向け: 生徒名簿データを取得
 * @returns {object} { success: boolean, students: Array }
 */
function getStudentRosterForAdmin() {
  requireTeacherOrAdmin('生徒名簿取得');
  try {
    const data = getStudentRosterDataCached();
    if (!data || data.length < 2) return { success: true, students: [] };
    const headers = data[0];
    const idIdx = headers.indexOf('学籍番号');
    const nameIdx = headers.indexOf('氏名');
    const gradeIdx = headers.indexOf('学年');
    const classIdx = headers.indexOf('組');
    const numIdx = headers.indexOf('番号');
    const students = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idIdx]) continue;
      students.push({
        studentId: String(row[idIdx]),
        name: row[nameIdx] || '',
        grade: row[gradeIdx] || '',
        class: row[classIdx] || '',
        number: row[numIdx] || ''
      });
    }
    return { success: true, students: students };
  } catch (error) {
    return handleError(error, '生徒名簿取得');
  }
}

/**
 * 管理者向け: 科目の学籍番号列を更新
 * @param {string} courseName - 科目名
 * @param {string} columnName - 列名（'抽選学籍番号','非表示学籍番号','表示学籍番号'のいずれか）
 * @param {string} studentIds - カンマ区切りの学籍番号文字列
 * @returns {object} { success: boolean }
 */
function updateCourseStudentIds(courseName, columnName, studentIds) {
  requireAdmin('科目学籍番号更新');
  try {
    // 許可された列名のバリデーション
    const allowedColumns = ['抽選学籍番号', '非表示学籍番号', '表示学籍番号'];
    if (!allowedColumns.includes(columnName)) {
      throw new Error('無効な列名: ' + columnName);
    }
    if (!courseName) {
      throw new Error('科目名が指定されていません');
    }

    const sheet = getCourseSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 列インデックスを特定
    const colIdx = headers.indexOf(columnName);
    if (colIdx === -1) {
      throw new Error('列が見つかりません: ' + columnName);
    }

    // 科目名列のインデックス
    const nameIdx = headers.indexOf('科目名');
    if (nameIdx === -1) {
      throw new Error('科目名列が見つかりません');
    }

    // 該当科目の行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameIdx]).trim() === String(courseName).trim()) {
        targetRow = i + 1; // シートの行番号（1始まり）
        break;
      }
    }
    if (targetRow === -1) {
      throw new Error('科目が見つかりません: ' + courseName);
    }

    // セルを更新
    sheet.getRange(targetRow, colIdx + 1).setValue(studentIds || '');

    // 科目データキャッシュをクリア
    _memCache.course = null;
    try {
      CacheService.getScriptCache().remove('course_master_data');
    } catch (e) {
      console.warn('キャッシュクリア失敗:', e);
    }

    logMessage('科目学籍番号更新: ' + courseName + ' / ' + columnName);

    return { success: true };
  } catch (error) {
    return handleError(error, '科目学籍番号更新');
  }
}

/**
 * 科目別学籍番号を一括更新（Excel取込用）
 * @param {Array} updates - [{courseName: string, studentIds: string}, ...]
 * @param {string} columnName - 列名（'抽選学籍番号'/'非表示学籍番号'/'表示学籍番号'）
 * @returns {object} { success: true, updatedCount: N }
 */
function batchUpdateCourseStudentIds(updates, columnName) {
  requireAdmin('科目学籍番号一括更新');
  try {
    // 列名バリデーション
    const allowedColumns = ['抽選学籍番号', '非表示学籍番号', '表示学籍番号'];
    if (!allowedColumns.includes(columnName)) {
      throw new Error('無効な列名: ' + columnName);
    }
    if (!updates || !Array.isArray(updates) || updates.length === 0) {
      throw new Error('更新データが空です');
    }

    const sheet = getCourseSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 列インデックスを特定
    const colIdx = headers.indexOf(columnName);
    if (colIdx === -1) {
      throw new Error('列が見つかりません: ' + columnName);
    }
    const nameIdx = headers.indexOf('科目名');
    if (nameIdx === -1) {
      throw new Error('科目名列が見つかりません');
    }

    // 科目名→行番号マップを構築（1回のシート読み込みで済ませる）
    const courseRowMap = {};
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][nameIdx]).trim();
      if (name) courseRowMap[name] = i + 1; // シートの行番号（1始まり）
    }

    // 各updateを処理
    let updatedCount = 0;
    updates.forEach(function(update) {
      const courseName = String(update.courseName).trim();
      const studentIds = update.studentIds || '';
      const targetRow = courseRowMap[courseName];
      if (!targetRow) {
        console.warn('科目が見つかりません（スキップ）: ' + courseName);
        return;
      }
      sheet.getRange(targetRow, colIdx + 1).setValue(studentIds);
      updatedCount++;
    });

    // 科目データキャッシュをクリア
    _memCache.course = null;
    try {
      CacheService.getScriptCache().remove('course_master_data');
    } catch (e) {
      console.warn('キャッシュクリア失敗:', e);
    }

    logMessage('科目学籍番号一括更新: ' + columnName + ' / ' + updatedCount + '科目');

    return { success: true, updatedCount: updatedCount };
  } catch (error) {
    return handleError(error, '科目学籍番号一括更新');
  }
}

/**
 * 科目データの学籍番号列（抽選/非表示/表示）を一括置換
 * 生徒マスターや登録データ管理で学籍番号を変更した際の連動更新用
 * @param {Array} replacements - [{oldId: '旧学籍番号', newId: '新学籍番号'}, ...]
 * @returns {object} {success: true, updatedCount: N}
 */
function replaceStudentIdInCourseSettings(replacements) {
  requireAdmin('科目設定学籍番号置換');
  try {
    if (!Array.isArray(replacements) || replacements.length === 0) {
      return { success: true, updatedCount: 0 };
    }

    const sheet = getCourseSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 対象列のインデックスを取得
    const targetColumns = ['抽選学籍番号', '非表示学籍番号', '表示学籍番号'];
    const colIndices = [];
    for (var c = 0; c < targetColumns.length; c++) {
      var idx = headers.indexOf(targetColumns[c]);
      if (idx !== -1) colIndices.push(idx);
    }
    if (colIndices.length === 0) {
      return { success: true, updatedCount: 0 };
    }

    // 置換マップを構築（文字列トリム済み）
    var replaceMap = {};
    for (var r = 0; r < replacements.length; r++) {
      var oldId = String(replacements[r].oldId).trim();
      var newId = String(replacements[r].newId).trim();
      if (oldId && newId && oldId !== newId) {
        replaceMap[oldId] = newId;
      }
    }
    if (Object.keys(replaceMap).length === 0) {
      return { success: true, updatedCount: 0 };
    }

    // 全行を走査して置換
    var updatedCount = 0;
    for (var i = 1; i < data.length; i++) {
      for (var j = 0; j < colIndices.length; j++) {
        var ci = colIndices[j];
        var cellValue = String(data[i][ci] || '').trim();
        if (!cellValue) continue;

        var ids = cellValue.split(',').map(function(s) { return s.trim(); });
        var changed = false;
        for (var k = 0; k < ids.length; k++) {
          if (replaceMap[ids[k]]) {
            ids[k] = replaceMap[ids[k]];
            changed = true;
          }
        }
        if (changed) {
          sheet.getRange(i + 1, ci + 1).setValue(ids.join(','));
          updatedCount++;
        }
      }
    }

    // 科目データキャッシュをクリア
    _memCache.course = null;
    try {
      CacheService.getScriptCache().remove('course_master_data');
    } catch (e) {
      console.warn('キャッシュクリア失敗:', e);
    }

    logMessage('科目設定学籍番号置換: ' + updatedCount + '件更新');

    return { success: true, updatedCount: updatedCount };
  } catch (error) {
    return handleError(error, '科目設定学籍番号置換');
  }
}

// ===== 科目データ管理（admin専用） =====

/**
 * 科目一覧を取得（管理画面用）
 * @returns {object} { success: true, courses: [...], headers: [...] }
 */
function getCourseList() {
  requireAdmin('科目一覧取得');
  try {
    const sheet = getCourseSheet();
    if (!sheet) {
      throw new Error('科目データシートが見つかりません');
    }
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      return { success: true, courses: [], headers: [] };
    }
    const headers = data[0];
    const courses = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const course = {};
      headers.forEach(function(header, index) {
        course[header] = row[index];
      });
      if (course['科目名']) {
        courses.push(course);
      }
    }
    return { success: true, courses: courses, headers: headers };
  } catch (error) {
    return handleError(error, '科目一覧取得');
  }
}

/**
 * 科目を追加（科目データシート + 登録データシートに列追加）
 * @param {object} courseData - 科目データ
 * @returns {object} { success: boolean }
 */
function addCourse(courseData) {
  requireAdmin('科目追加');
  try {
    if (!courseData || !courseData['科目名']) {
      throw new Error('科目名は必須です');
    }
    const courseName = String(courseData['科目名']).trim();
    if (!courseName) {
      throw new Error('科目名は必須です');
    }

    // 科目データシートに追加
    const courseSheet = getCourseSheet();
    if (!courseSheet) {
      throw new Error('科目データシートが見つかりません');
    }
    const courseAllData = courseSheet.getDataRange().getValues();
    const headers = courseAllData[0];

    // 重複チェック
    const nameIdx = headers.indexOf('科目名');
    if (nameIdx === -1) {
      throw new Error('科目データシートに科目名列が見つかりません');
    }
    for (let i = 1; i < courseAllData.length; i++) {
      if (String(courseAllData[i][nameIdx]).trim() === courseName) {
        return { success: false, error: '同名の科目が既に存在します: ' + courseName };
      }
    }

    // ヘッダー順に新しい行データを構築
    const newRow = headers.map(function(header) {
      if (courseData.hasOwnProperty(header)) {
        return courseData[header];
      }
      return '';
    });
    courseSheet.appendRow(newRow);

    // 登録データシートの末尾に列追加（ヘッダー = 科目名）
    const submissionSheet = getSubmissionSheet();
    if (submissionSheet) {
      const lastCol = submissionSheet.getLastColumn();
      submissionSheet.getRange(1, lastCol + 1).setValue(courseName);
    }

    // キャッシュクリア
    clearDataCache();

    logMessage('科目追加: ' + courseName);
    return { success: true };
  } catch (error) {
    return handleError(error, '科目追加');
  }
}

/**
 * 科目を更新（科目名変更時は登録データシートのヘッダーもリネーム）
 * @param {string} originalName - 元の科目名
 * @param {object} courseData - 更新後の科目データ
 * @returns {object} { success: boolean }
 */
function updateCourse(originalName, courseData) {
  requireAdmin('科目更新');
  try {
    if (!originalName || !courseData || !courseData['科目名']) {
      throw new Error('科目名は必須です');
    }
    const newName = String(courseData['科目名']).trim();
    if (!newName) {
      throw new Error('科目名は必須です');
    }

    const courseSheet = getCourseSheet();
    if (!courseSheet) {
      throw new Error('科目データシートが見つかりません');
    }
    const courseAllData = courseSheet.getDataRange().getValues();
    const headers = courseAllData[0];
    const nameIdx = headers.indexOf('科目名');
    if (nameIdx === -1) {
      throw new Error('科目データシートに科目名列が見つかりません');
    }

    // 該当行を検索
    var targetRow = -1;
    for (let i = 1; i < courseAllData.length; i++) {
      if (String(courseAllData[i][nameIdx]).trim() === String(originalName).trim()) {
        targetRow = i + 1; // シートの行番号（1始まり）
        break;
      }
    }
    if (targetRow === -1) {
      throw new Error('科目が見つかりません: ' + originalName);
    }

    // 科目名変更の場合は新名の重複チェック
    if (newName !== String(originalName).trim()) {
      for (let i = 1; i < courseAllData.length; i++) {
        if (i + 1 === targetRow) continue; // 自分自身はスキップ
        if (String(courseAllData[i][nameIdx]).trim() === newName) {
          return { success: false, error: '同名の科目が既に存在します: ' + newName };
        }
      }

      // 登録データシートのヘッダーもリネーム
      const submissionSheet = getSubmissionSheet();
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
    headers.forEach(function(header, colIndex) {
      if (courseData.hasOwnProperty(header)) {
        courseSheet.getRange(targetRow, colIndex + 1).setValue(courseData[header]);
      }
    });

    // キャッシュクリア
    clearDataCache();

    logMessage('科目更新: ' + originalName + (newName !== originalName ? ' → ' + newName : ''));
    return { success: true };
  } catch (error) {
    return handleError(error, '科目更新');
  }
}

/**
 * 科目を削除（履修データが存在する場合は拒否）
 * @param {string} courseName - 科目名
 * @returns {object} { success: boolean }
 */
function deleteCourse(courseName) {
  requireAdmin('科目削除');
  try {
    if (!courseName) {
      throw new Error('科目名が指定されていません');
    }
    courseName = String(courseName).trim();

    // 登録データシートで該当科目の列を検索
    const submissionSheet = getSubmissionSheet();
    if (!submissionSheet) {
      throw new Error('登録データシートが見つかりません');
    }
    const subData = submissionSheet.getDataRange().getValues();
    const subHeaders = subData[0];
    var subColIdx = -1;
    for (let j = 0; j < subHeaders.length; j++) {
      if (String(subHeaders[j]).trim() === courseName) {
        subColIdx = j;
        break;
      }
    }

    // 安全チェック: 登録データシートに列がある場合、履修データを走査
    if (subColIdx !== -1) {
      var dataCount = 0;
      for (let i = 1; i < subData.length; i++) {
        var val = String(subData[i][subColIdx] || '').trim();
        // ●/○（半角○含む）が1つでもあれば拒否
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
    const courseSheet = getCourseSheet();
    if (!courseSheet) {
      throw new Error('科目データシートが見つかりません');
    }
    const courseAllData = courseSheet.getDataRange().getValues();
    const courseHeaders = courseAllData[0];
    const nameIdx = courseHeaders.indexOf('科目名');
    if (nameIdx === -1) {
      throw new Error('科目データシートに科目名列が見つかりません');
    }

    var courseRow = -1;
    for (let i = 1; i < courseAllData.length; i++) {
      if (String(courseAllData[i][nameIdx]).trim() === courseName) {
        courseRow = i + 1; // シートの行番号（1始まり）
        break;
      }
    }
    if (courseRow === -1) {
      throw new Error('科目データシートに該当科目が見つかりません: ' + courseName);
    }
    courseSheet.deleteRow(courseRow);

    // キャッシュクリア
    clearDataCache();

    logMessage('科目削除: ' + courseName);
    return { success: true };
  } catch (error) {
    return handleError(error, '科目削除');
  }
}

// ============================================================
// 教職員データ管理（管理者専用）
// ============================================================

/**
 * 教職員一覧を取得（管理画面用）
 * @returns {object} { success: true, staff: [...], headers: [...] }
 */
function getStaffList() {
  requireAdmin('教職員一覧取得');
  try {
    const sheet = getTeacherSheet();
    if (!sheet) {
      throw new Error('教職員データシートが見つかりません');
    }
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      return { success: true, staff: [], headers: [] };
    }
    const headers = data[0];
    const staff = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const member = {};
      headers.forEach(function(header, index) {
        member[header] = row[index];
      });
      // メールアドレスがある行のみ
      if (member['メールアドレス']) {
        staff.push(member);
      }
    }
    return { success: true, staff: staff, headers: headers };
  } catch (error) {
    return handleError(error, '教職員一覧取得');
  }
}

/**
 * 教職員を追加
 * @param {object} staffData - 教職員データ
 * @returns {object} { success: boolean }
 */
function addStaff(staffData) {
  requireAdmin('教職員追加');
  try {
    if (!staffData || !staffData['メールアドレス']) {
      throw new Error('メールアドレスは必須です');
    }
    const email = String(staffData['メールアドレス']).trim().toLowerCase();
    if (!email) {
      throw new Error('メールアドレスは必須です');
    }

    const sheet = getTeacherSheet();
    if (!sheet) {
      throw new Error('教職員データシートが見つかりません');
    }
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    // メールアドレス重複チェック（大文字小文字不問）
    const emailIdx = headers.indexOf('メールアドレス');
    if (emailIdx === -1) {
      throw new Error('教職員データシートにメールアドレス列が見つかりません');
    }
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][emailIdx]).trim().toLowerCase() === email) {
        return { success: false, error: '同じメールアドレスの教職員が既に存在します: ' + email };
      }
    }

    // ヘッダー順に新しい行データを構築
    const newRow = headers.map(function(header) {
      if (staffData.hasOwnProperty(header)) {
        return staffData[header];
      }
      return '';
    });
    sheet.appendRow(newRow);

    // キャッシュクリア
    clearDataCache();

    logMessage('教職員追加: ' + email);
    return { success: true };
  } catch (error) {
    return handleError(error, '教職員追加');
  }
}

/**
 * 教職員を更新
 * @param {string} originalEmail - 元のメールアドレス
 * @param {object} staffData - 更新後の教職員データ
 * @returns {object} { success: boolean }
 */
function updateStaff(originalEmail, staffData) {
  requireAdmin('教職員更新');
  try {
    if (!originalEmail || !staffData || !staffData['メールアドレス']) {
      throw new Error('メールアドレスは必須です');
    }
    const newEmail = String(staffData['メールアドレス']).trim().toLowerCase();
    if (!newEmail) {
      throw new Error('メールアドレスは必須です');
    }

    const sheet = getTeacherSheet();
    if (!sheet) {
      throw new Error('教職員データシートが見つかりません');
    }
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const emailIdx = headers.indexOf('メールアドレス');
    if (emailIdx === -1) {
      throw new Error('教職員データシートにメールアドレス列が見つかりません');
    }

    // 該当行を検索
    var targetRow = -1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][emailIdx]).trim().toLowerCase() === String(originalEmail).trim().toLowerCase()) {
        targetRow = i + 1; // シートの行番号（1始まり）
        break;
      }
    }
    if (targetRow === -1) {
      throw new Error('教職員が見つかりません: ' + originalEmail);
    }

    // メールアドレス変更の場合は新メールアドレスの重複チェック
    if (newEmail !== String(originalEmail).trim().toLowerCase()) {
      for (let i = 1; i < allData.length; i++) {
        if (i + 1 === targetRow) continue; // 自分自身はスキップ
        if (String(allData[i][emailIdx]).trim().toLowerCase() === newEmail) {
          return { success: false, error: '同じメールアドレスの教職員が既に存在します: ' + newEmail };
        }
      }
    }

    // 各列の値を更新
    headers.forEach(function(header, colIndex) {
      if (staffData.hasOwnProperty(header)) {
        sheet.getRange(targetRow, colIndex + 1).setValue(staffData[header]);
      }
    });

    // キャッシュクリア
    clearDataCache();

    logMessage('教職員更新: ' + originalEmail + (newEmail !== String(originalEmail).trim().toLowerCase() ? ' → ' + newEmail : ''));
    return { success: true };
  } catch (error) {
    return handleError(error, '教職員更新');
  }
}

/**
 * 教職員を削除
 * @param {string} email - メールアドレス
 * @returns {object} { success: boolean }
 */
function deleteStaff(email) {
  requireAdmin('教職員削除');
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

    const sheet = getTeacherSheet();
    if (!sheet) {
      throw new Error('教職員データシートが見つかりません');
    }
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const emailIdx = headers.indexOf('メールアドレス');
    if (emailIdx === -1) {
      throw new Error('教職員データシートにメールアドレス列が見つかりません');
    }

    var targetRow = -1;
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
    clearDataCache();

    logMessage('教職員削除: ' + email);
    return { success: true };
  } catch (error) {
    return handleError(error, '教職員削除');
  }
}

/**
 * 教職員データを一括置換
 * @param {Array<Object>} staffArray - 教職員データの配列
 * @returns {object} { success: boolean, count: number }
 */
function replaceAllStaff(staffArray) {
  requireAdmin('教職員データ一括保存');
  try {
    // バリデーション
    if (!Array.isArray(staffArray) || staffArray.length === 0) {
      throw new Error('教職員データが空です');
    }

    // 自分自身のメールが含まれているか確認（自己削除防止）
    var currentEmail = Session.getActiveUser().getEmail().trim().toLowerCase();
    var found = false;
    var emailSet = {};
    for (var i = 0; i < staffArray.length; i++) {
      var email = String(staffArray[i]['メールアドレス'] || '').trim().toLowerCase();
      if (!email) {
        throw new Error((i + 1) + '行目: メールアドレスが空です');
      }
      // メールアドレス重複チェック
      if (emailSet[email]) {
        throw new Error('メールアドレスが重複しています: ' + email);
      }
      emailSet[email] = true;
      if (email === currentEmail) {
        found = true;
      }
    }
    if (!found) {
      throw new Error('自分自身のメールアドレスが含まれていません。自分を削除することはできません。');
    }

    // シート取得・ヘッダー行を取得
    var sheet = getTeacherSheet();
    if (!sheet) {
      throw new Error('教職員データシートが見つかりません');
    }
    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];

    // 既存データ行をすべて削除（ヘッダー行は残す）
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    // 新データをヘッダー順にマッピングして一括書き込み
    var rows = staffArray.map(function(staff) {
      return headers.map(function(header) {
        if (staff.hasOwnProperty(header)) {
          return staff[header];
        }
        return '';
      });
    });
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    // キャッシュクリア
    clearDataCache();

    logMessage('教職員データ一括保存: ' + rows.length + '件');
    return { success: true, count: rows.length };
  } catch (error) {
    return handleError(error, '教職員データ一括保存');
  }
}

/**
 * 生徒一覧を取得（登録データシートから基本列のみ読み取り）
 * @returns {object} { success: boolean, students: Array, headers: Array }
 */
function getStudentList() {
  requireAdmin('生徒一覧取得');
  try {
    var submissionData = getSubmissionSheetDataCached();
    if (!submissionData || submissionData.length === 0) {
      return { success: true, students: [], headers: [] };
    }
    var subHeaders = submissionData[0];

    // 基本フィールド
    var fields = ['学籍番号', '学年', '組', '番号', '名前', 'メールアドレス', '来年度学年'];
    var fieldIndices = {};
    for (var f = 0; f < fields.length; f++) {
      var idx = subHeaders.indexOf(fields[f]);
      if (idx !== -1) {
        fieldIndices[fields[f]] = idx;
      }
    }

    var idIdx = fieldIndices['学籍番号'];
    if (idIdx === undefined) {
      throw new Error('登録データシートに学籍番号列が見つかりません');
    }

    // 各行から生徒情報を抽出（基本列のみ）
    var students = [];
    for (var i = 1; i < submissionData.length; i++) {
      var row = submissionData[i];
      var studentId = String(row[idIdx] || '').trim();
      if (!studentId) continue;

      var student = {};
      for (var key in fieldIndices) {
        var val = row[fieldIndices[key]];
        student[key] = val !== undefined && val !== null && val !== '' ? String(val) : '';
      }
      students.push(student);
    }

    return { success: true, students: students, headers: fields };
  } catch (error) {
    return handleError(error, '生徒一覧取得');
  }
}

/**
 * 登録データシートのデータ行を学年→組→番号で昇順ソート
 * 組は数値のみの値を前、非数値を含む値を後に並べる
 * @param {Sheet} sheet - 登録データシート
 * @param {Array} headers - ヘッダー行の配列
 */
function _sortSubmissionSheet(sheet, headers) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // ヘッダーのみ

  var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var data = dataRange.getValues();

  // タイムスタンプ列のDate→String変換によるずれを防止
  var tsIdx = headers.indexOf('タイムスタンプ');
  if (tsIdx !== -1) {
    var displayValues = dataRange.getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      data[i][tsIdx] = displayValues[i][tsIdx];
    }
  }

  var gradeIdx = headers.indexOf('学年');
  var classIdx = headers.indexOf('組');
  var numIdx = headers.indexOf('番号');
  if (gradeIdx === -1 || classIdx === -1 || numIdx === -1) return;

  data.sort(function(a, b) {
    // 非数値組（例: 1K）は全体の末尾に配置
    var aClass = String(a[classIdx] || '').trim();
    var bClass = String(b[classIdx] || '').trim();
    var aIsNum = !aClass || /^\d+$/.test(aClass);
    var bIsNum = !bClass || /^\d+$/.test(bClass);
    if (aIsNum !== bIsNum) return aIsNum ? -1 : 1;

    // 学年で昇順
    var aGrade = safeParseGradeValue(a[gradeIdx], 999);
    var bGrade = safeParseGradeValue(b[gradeIdx], 999);
    if (aGrade !== bGrade) return aGrade - bGrade;

    // 組で昇順
    if (!aClass && bClass) return 1;
    if (aClass && !bClass) return -1;
    if (aIsNum && bIsNum) {
      var classDiff = parseInt(aClass) - parseInt(bClass);
      if (classDiff !== 0) return classDiff;
    } else {
      var classCmp = aClass.localeCompare(bClass);
      if (classCmp !== 0) return classCmp;
    }

    // 番号で昇順
    var aNum = parseInt(a[numIdx]) || 999;
    var bNum = parseInt(b[numIdx]) || 999;
    return aNum - bNum;
  });

  dataRange.setValues(data);
}

/**
 * 生徒データを一括保存（登録データシートの基本列を差分更新・新規追加・削除）
 * @param {Array} studentArray - 生徒データの配列（基本列のみ）
 * @returns {object} { success: boolean, count: number }
 */
function replaceAllStudents(studentArray) {
  requireAdmin('生徒データ一括保存');
  try {
    // バリデーション
    if (!Array.isArray(studentArray) || studentArray.length === 0) {
      throw new Error('生徒データが空です');
    }

    var idSet = {};
    for (var i = 0; i < studentArray.length; i++) {
      var studentId = String(studentArray[i]['学籍番号'] || '').trim();
      if (!studentId) {
        throw new Error((i + 1) + '行目: 学籍番号が空です');
      }
      if (idSet[studentId]) {
        throw new Error('学籍番号が重複しています: ' + studentId);
      }
      idSet[studentId] = true;
      // 学年の値域チェック
      var grade = String(studentArray[i]['学年'] || '').trim();
      if (grade && grade !== '0' && grade !== '1' && grade !== '2' && grade !== '3') {
        throw new Error((i + 1) + '行目: 学年は0/1/2/3のいずれかで入力してください（' + grade + '）');
      }
      // 来年度学年の値域チェック
      var nextGrade = String(studentArray[i]['来年度学年'] || '').trim();
      if (nextGrade && nextGrade !== '1' && nextGrade !== '2' && nextGrade !== '3') {
        throw new Error((i + 1) + '行目: 来年度学年は1/2/3のいずれかで入力してください（' + nextGrade + '）');
      }
    }

    // --- 登録データシートの基本列を差分更新・新規追加・削除 ---
    var submissionSheet = getSubmissionSheet();
    if (!submissionSheet) {
      throw new Error('登録データシートが見つかりません');
    }
    var subData = submissionSheet.getDataRange().getValues();
    if (subData.length === 0) {
      throw new Error('登録データシートにヘッダー行がありません');
    }
    var subHeaders = subData[0];
    var subIdIdx = subHeaders.indexOf('学籍番号');
    if (subIdIdx === -1) {
      throw new Error('登録データシートに学籍番号列が見つかりません');
    }

    // 更新対象の基本列
    var syncColumns = ['学年', '組', '番号', '名前', 'メールアドレス', '来年度学年'];
    var syncColIndices = {};
    for (var c = 0; c < syncColumns.length; c++) {
      var idx = subHeaders.indexOf(syncColumns[c]);
      if (idx !== -1) {
        syncColIndices[syncColumns[c]] = idx;
      }
    }

    // 学籍番号→生徒データのマップを構築
    var newStudentMap = {};
    for (var j = 0; j < studentArray.length; j++) {
      var sid = String(studentArray[j]['学籍番号'] || '').trim();
      if (sid) newStudentMap[sid] = studentArray[j];
    }

    // 既存行の学籍番号セットを構築
    var existingIdSet = {};
    for (var k = 1; k < subData.length; k++) {
      var existingId = String(subData[k][subIdIdx] || '').trim();
      if (existingId) existingIdSet[existingId] = true;
    }

    // 1. 既存行の基本列を差分更新
    for (var k = 1; k < subData.length; k++) {
      var existingId = String(subData[k][subIdIdx] || '').trim();
      if (existingId && newStudentMap.hasOwnProperty(existingId)) {
        var student = newStudentMap[existingId];
        var rowChanged = false;
        var row = subData[k].slice(); // 行データのコピー
        // 基本列の更新
        for (var col in syncColIndices) {
          var colIdx = syncColIndices[col];
          var currentVal = String(row[colIdx] || '').trim();
          var newVal = String(student[col] || '').trim();
          if (currentVal !== newVal) {
            // 数値列（学年, 組, 番号, 来年度学年）は数値に変換（非数値文字列はそのまま）
            if ((col === '学年' || col === '組' || col === '番号' || col === '来年度学年') && newVal !== '') {
              row[colIdx] = /^\d+$/.test(newVal) ? Number(newVal) : newVal;
            } else {
              row[colIdx] = newVal;
            }
            rowChanged = true;
          }
        }
        if (rowChanged) {
          submissionSheet.getRange(k + 1, 1, 1, row.length).setValues([row]);
        }
      }
    }

    // 2. 新規生徒の行追加（studentArrayにあるが登録データシートにない学籍番号）
    var newRows = [];
    for (var n = 0; n < studentArray.length; n++) {
      var newId = String(studentArray[n]['学籍番号'] || '').trim();
      if (newId && !existingIdSet[newId]) {
        // ヘッダー順にマッピングして行を作成（科目列は空）
        var newRow = subHeaders.map(function(header) {
          if (studentArray[n].hasOwnProperty(header)) {
            var val = String(studentArray[n][header] || '').trim();
            // 数値列は数値に変換（非数値文字列はそのまま）
            if ((header === '学年' || header === '組' || header === '番号' || header === '来年度学年') && val !== '') {
              return /^\d+$/.test(val) ? Number(val) : val;
            }
            return val;
          }
          return '';
        });
        newRows.push(newRow);
      }
    }
    if (newRows.length > 0) {
      var lastRow = submissionSheet.getLastRow();
      submissionSheet.getRange(lastRow + 1, 1, newRows.length, subHeaders.length).setValues(newRows);
      console.log('新規生徒追加: ' + newRows.length + '件');
    }

    // 3. 削除された生徒の行削除（登録データシートにあるがstudentArrayにない学籍番号）
    // 下から上へ削除してインデックスずれを防止
    var deletedCount = 0;
    for (var d = subData.length - 1; d >= 1; d--) {
      var delId = String(subData[d][subIdIdx] || '').trim();
      if (delId && !newStudentMap.hasOwnProperty(delId)) {
        submissionSheet.deleteRow(d + 1); // シートは1ベース
        deletedCount++;
      }
    }
    if (deletedCount > 0) {
      console.log('生徒削除: ' + deletedCount + '件');
    }

    // 4. 登録データシートを学年→組→番号でソート
    _sortSubmissionSheet(submissionSheet, subHeaders);

    // キャッシュクリア
    clearDataCache();

    logMessage('生徒データ一括保存: ' + studentArray.length + '件（新規: ' + newRows.length + '件、削除: ' + deletedCount + '件）');
    return { success: true, count: studentArray.length };
  } catch (error) {
    return handleError(error, '生徒データ一括保存');
  }
}

/**
 * 新入生データを登録データシートに追加（admin専用）
 * @param {Array} studentArray - [{学籍番号, 名前, 組?, 番号?, 学年?}, ...]
 * @param {boolean} autoMark - ○1を自動付与するか
 * @param {boolean} finalRegistration - 本登録ステータスで登録するか
 * @returns {object} { success, addedCount, skippedDuplicates }
 */
function importNewStudents(studentArray, autoMark, finalRegistration) {
  requireAdmin('新入生データ取込');
  try {
    // バリデーション
    if (!Array.isArray(studentArray) || studentArray.length === 0) {
      throw new Error('新入生データが空です');
    }

    // 学籍番号の必須チェック・重複チェック・学年チェック
    var idSet = {};
    for (var i = 0; i < studentArray.length; i++) {
      var studentId = String(studentArray[i]['学籍番号'] || '').trim();
      if (!studentId) {
        throw new Error((i + 1) + '行目: 学籍番号が空です');
      }
      if (!String(studentArray[i]['名前'] || '').trim()) {
        throw new Error((i + 1) + '行目: 名前が空です（学籍番号: ' + studentId + '）');
      }
      if (idSet[studentId]) {
        throw new Error('Excel内で学籍番号が重複しています: ' + studentId);
      }
      idSet[studentId] = true;
      var grade = String(studentArray[i]['学年'] || '0').trim();
      if (grade !== '0' && grade !== '1') {
        throw new Error((i + 1) + '行目: 学年は0または1のみ指定可能です（' + grade + '）');
      }
    }

    // 登録データシート読み込み
    var submissionSheet = getSubmissionSheet();
    if (!submissionSheet) {
      throw new Error('登録データシートが見つかりません');
    }
    var subData = submissionSheet.getDataRange().getValues();
    if (subData.length === 0) {
      throw new Error('登録データシートにヘッダー行がありません');
    }
    var subHeaders = subData[0];
    var subIdIdx = subHeaders.indexOf('学籍番号');
    if (subIdIdx === -1) {
      throw new Error('登録データシートに学籍番号列が見つかりません');
    }

    // 既存学籍番号のセットを構築
    var existingIdSet = {};
    for (var k = 1; k < subData.length; k++) {
      var existingId = String(subData[k][subIdIdx] || '').trim();
      if (existingId) existingIdSet[existingId] = true;
    }

    // 新規行を構築（重複はスキップ）
    var newRows = [];
    var skippedDuplicates = [];
    var statusColIdx = subHeaders.indexOf('ステータス');

    for (var n = 0; n < studentArray.length; n++) {
      var newId = String(studentArray[n]['学籍番号'] || '').trim();
      if (existingIdSet[newId]) {
        skippedDuplicates.push(newId);
        continue;
      }
      // ヘッダー順にマッピングして行を作成（科目列は空）
      var newRow = subHeaders.map(function(header) {
        if (header === 'ステータス' && finalRegistration) {
          return '本登録';
        }
        if (studentArray[n].hasOwnProperty(header)) {
          var val = String(studentArray[n][header] || '').trim();
          // 数値列は数値に変換
          if ((header === '学年' || header === '組' || header === '番号') && val !== '') {
            return /^\d+$/.test(val) ? Number(val) : val;
          }
          return val;
        }
        return '';
      });
      newRows.push(newRow);
    }

    // ○1自動付与
    if (autoMark && newRows.length > 0) {
      var courseData = getCourseDataCached();
      if (courseData && courseData.length > 1) {
        var courseHeaders = courseData[0];
        var courseNameIdx = courseHeaders.indexOf('科目名');
        var courseGradeIdx = courseHeaders.indexOf('学年');
        var courseCategoryIdx = courseHeaders.indexOf('区分');
        var courseNoOpenIdx = courseHeaders.indexOf('開講なし');
        var courseVisShowIdx = courseHeaders.indexOf('表示学籍番号');
        var courseVisHideIdx = courseHeaders.indexOf('非表示学籍番号');

        // ○1対象の科目リストを構築（表示/非表示情報付き）
        var autoMarkCourses = [];
        for (var ci = 1; ci < courseData.length; ci++) {
          var cRow = courseData[ci];
          var cGrade = String(cRow[courseGradeIdx] ?? '').trim();
          var cCategory = String(cRow[courseCategoryIdx] ?? '').trim();
          var cNoOpen = String(cRow[courseNoOpenIdx] ?? '').trim();
          var cName = String(cRow[courseNameIdx] ?? '').trim();
          if (cGrade === '1' && cName && !cNoOpen &&
              (cCategory === '必修' || cCategory === '履修指定' || cCategory === '選択必修-履修指定')) {
            autoMarkCourses.push({
              name: cName,
              showIds: courseVisShowIdx !== -1 ? String(cRow[courseVisShowIdx] ?? '').trim() : '',
              hideIds: courseVisHideIdx !== -1 ? String(cRow[courseVisHideIdx] ?? '').trim() : ''
            });
          }
        }

        if (autoMarkCourses.length > 0) {
          // 登録データシートヘッダーの科目名→列インデックスマップ
          var courseColMap = {};
          for (var hi = 0; hi < subHeaders.length; hi++) {
            courseColMap[subHeaders[hi]] = hi;
          }
          var gradeColIdx = subHeaders.indexOf('学年');
          var idColIdx = subHeaders.indexOf('学籍番号');

          // 新規行の中で学年0/1の行に○1を設定（表示/非表示チェック付き）
          for (var ri = 0; ri < newRows.length; ri++) {
            var rowGrade = String(newRows[ri][gradeColIdx] ?? '');
            if (rowGrade === '0' || rowGrade === '1' || newRows[ri][gradeColIdx] === 0 || newRows[ri][gradeColIdx] === 1) {
              var studentIdForMark = String(newRows[ri][idColIdx] ?? '').trim();
              for (var mi = 0; mi < autoMarkCourses.length; mi++) {
                var course = autoMarkCourses[mi];
                // 非表示学籍番号チェック（ブラックリスト）
                if (course.hideIds) {
                  var excludedIds = course.hideIds.split(',').map(function(id) { return id.trim(); });
                  if (excludedIds.indexOf(studentIdForMark) !== -1) continue;
                }
                // 表示学籍番号チェック（ホワイトリスト）
                if (course.showIds) {
                  var targetIds = course.showIds.split(',').map(function(id) { return id.trim(); });
                  if (targetIds.indexOf(studentIdForMark) === -1) continue;
                }
                var colIdx = courseColMap[course.name];
                if (colIdx !== undefined) {
                  newRows[ri][colIdx] = '○1';
                }
              }
            }
          }
        }
      }
    }

    // 行追加
    if (newRows.length > 0) {
      var lastRow = submissionSheet.getLastRow();
      submissionSheet.getRange(lastRow + 1, 1, newRows.length, subHeaders.length).setValues(newRows);
    }

    // ソート
    _sortSubmissionSheet(submissionSheet, subHeaders);

    // キャッシュクリア
    clearDataCache();

    logMessage('新入生データ取込: 追加' + newRows.length + '件、重複スキップ' + skippedDuplicates.length + '件');
    return {
      success: true,
      addedCount: newRows.length,
      skippedDuplicates: skippedDuplicates
    };
  } catch (error) {
    return handleError(error, '新入生データ取込');
  }
}

/**
 * 教科書シートからデータを取得（教員以上）
 * @returns {object} { success: true, data: [ { courseName, codes } ] }
 */
function getTextbookData() {
  requireTeacherOrAdmin('教科書データ取得');
  try {
    var sheet = getSpreadsheet().getSheetByName('教科書');
    if (!sheet) {
      return { success: false, error: '「教科書」シートが見つかりません。', timestamp: getJSTTimestamp() };
    }

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return { success: true, data: [], timestamp: getJSTTimestamp() };
    }

    var header = data[0];
    var courseNameIdx = header.indexOf('科目名');
    var codeIdx = header.indexOf('商品コード');

    if (courseNameIdx === -1 || codeIdx === -1) {
      return { success: false, error: '教科書シートに「科目名」または「商品コード」列が見つかりません。', timestamp: getJSTTimestamp() };
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var courseName = String(data[i][courseNameIdx] || '').trim();
      var rawCode = String(data[i][codeIdx] || '').trim();
      if (!courseName || !rawCode) continue;

      // カンマ区切りで分割し、空文字を除外
      var codes = rawCode.split(/[,，]/).map(function(c) { return c.trim(); }).filter(function(c) { return c !== ''; });
      if (codes.length > 0) {
        result.push({ courseName: courseName, codes: codes });
      }
    }

    return { success: true, data: result, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, '教科書データ取得');
  }
}

/**
 * SM設定シートを取得（なければ作成）
 */
function getSMSettingsSheet() {
  if (!SM_SETTINGS_SHEET) {
    SM_SETTINGS_SHEET = getSpreadsheet().getSheetByName('SM設定');
    if (!SM_SETTINGS_SHEET) {
      SM_SETTINGS_SHEET = getSpreadsheet().insertSheet('SM設定');
      SM_SETTINGS_SHEET.getRange(1, 1, 1, 3).setValues([['科目名', 'グループ番号', 'クラス一覧']]);
    }
  }
  return SM_SETTINGS_SHEET;
}

/**
 * SM設定データを取得（教員以上）
 * @returns {object} { success: true, data: [{courseName, groupNumber, classes}] }
 */
function getSMSettings() {
  requireTeacherOrAdmin('SM設定取得');
  try {
    var sheet = getSMSettingsSheet();
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return { success: true, data: [], timestamp: getJSTTimestamp() };
    }

    var header = data[0];
    var courseNameIdx = header.indexOf('科目名');
    var groupNumberIdx = header.indexOf('グループ番号');
    var classesIdx = header.indexOf('クラス一覧');

    if (courseNameIdx === -1 || groupNumberIdx === -1 || classesIdx === -1) {
      return { success: false, error: 'SM設定シートのヘッダーが不正です。', timestamp: getJSTTimestamp() };
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var courseName = String(data[i][courseNameIdx] || '').trim();
      var groupNumber = parseInt(data[i][groupNumberIdx] || '0', 10);
      var classes = String(data[i][classesIdx] || '').trim();
      if (!courseName) continue;
      result.push({ courseName: courseName, groupNumber: groupNumber, classes: classes });
    }

    return { success: true, data: result, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, 'SM設定取得');
  }
}

/**
 * SM設定データを保存（教務以上）
 * @param {Array} data [{courseName, groupNumber, classes}]
 * @returns {object} { success: true }
 */
function saveSMSettings(data) {
  var userInfo = getUserInfo();
  if (!userInfo.success || !isKyomuOrAbove(userInfo.user.role)) {
    return { success: false, error: '教務以上の権限が必要です。', timestamp: getJSTTimestamp() };
  }
  try {
    var sheet = getSMSettingsSheet();
    // ヘッダー行以外を全削除
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    // ヘッダーから列位置を動的に検出
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var courseNameIdx = header.indexOf('科目名');
    var groupNumberIdx = header.indexOf('グループ番号');
    var classesIdx = header.indexOf('クラス一覧');
    if (courseNameIdx === -1 || groupNumberIdx === -1 || classesIdx === -1) {
      return { success: false, error: 'SM設定シートに必要な列（科目名/グループ番号/クラス一覧）が見つかりません。', timestamp: getJSTTimestamp() };
    }
    var numCols = header.length;
    // 新データを一括書き込み
    if (data && data.length > 0) {
      var rows = data.map(function(item) {
        var row = new Array(numCols).fill('');
        row[courseNameIdx] = item.courseName || '';
        row[groupNumberIdx] = item.groupNumber || 0;
        row[classesIdx] = item.classes || '';
        return row;
      });
      sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
    }
    return { success: true, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, 'SM設定保存');
  }
}

/**
 * レポートシートからデータを取得（教員以上）
 * @returns {object} { success: true, data: [ { courseName, codes } ] }
 */
function getReportData() {
  requireTeacherOrAdmin('レポートデータ取得');
  try {
    var sheet = getSpreadsheet().getSheetByName('レポート');
    if (!sheet) {
      return { success: false, error: '「レポート」シートが見つかりません。', timestamp: getJSTTimestamp() };
    }

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return { success: true, data: [], timestamp: getJSTTimestamp() };
    }

    var header = data[0];
    var courseNameIdx = header.indexOf('科目名');
    var codeIdx = header.indexOf('科目コード');

    if (courseNameIdx === -1 || codeIdx === -1) {
      return { success: false, error: 'レポートシートに「科目名」または「科目コード」列が見つかりません。', timestamp: getJSTTimestamp() };
    }

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var courseName = String(data[i][courseNameIdx] || '').trim();
      var rawCode = String(data[i][codeIdx] || '').trim();
      if (!courseName || !rawCode) continue;

      // カンマ区切りで分割し、空文字を除外
      var codes = rawCode.split(/[,，]/).map(function(c) { return c.trim(); }).filter(function(c) { return c !== ''; });
      if (codes.length > 0) {
        result.push({ courseName: courseName, codes: codes });
      }
    }

    return { success: true, data: result, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, 'レポートデータ取得');
  }
}

/**
 * 教科書データを保存（教務以上）
 * @param {Array} data [{courseName: '科目名', codes: 'コード文字列'}, ...]
 * @returns {object} { success: true }
 */
function saveTextbookData(data) {
  var userInfo = getUserInfo();
  if (!userInfo.success || !isKyomuOrAbove(userInfo.user.role)) {
    return { success: false, error: '教務以上の権限が必要です。', timestamp: getJSTTimestamp() };
  }
  try {
    var sheet = getSpreadsheet().getSheetByName('教科書');
    if (!sheet) {
      return { success: false, error: '「教科書」シートが見つかりません。', timestamp: getJSTTimestamp() };
    }
    // ヘッダー行以外を全削除
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    // ヘッダーから列位置を動的に検出
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var courseNameIdx = header.indexOf('科目名');
    var codeIdx = header.indexOf('商品コード');
    if (courseNameIdx === -1 || codeIdx === -1) {
      return { success: false, error: '教科書シートに「科目名」または「商品コード」列が見つかりません。', timestamp: getJSTTimestamp() };
    }
    var numCols = header.length;
    // codesが空でない行のみ一括書き込み
    if (data && data.length > 0) {
      var rows = data.filter(function(item) {
        return item.codes && String(item.codes).trim() !== '';
      }).map(function(item) {
        var row = new Array(numCols).fill('');
        row[courseNameIdx] = item.courseName || '';
        row[codeIdx] = String(item.codes || '').trim();
        return row;
      });
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
      }
    }
    return { success: true, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, '教科書データ保存');
  }
}

/**
 * レポートデータを保存（教務以上）
 * @param {Array} data [{courseName: '科目名', codes: 'コード文字列'}, ...]
 * @returns {object} { success: true }
 */
function saveReportData(data) {
  var userInfo = getUserInfo();
  if (!userInfo.success || !isKyomuOrAbove(userInfo.user.role)) {
    return { success: false, error: '教務以上の権限が必要です。', timestamp: getJSTTimestamp() };
  }
  try {
    var sheet = getSpreadsheet().getSheetByName('レポート');
    if (!sheet) {
      return { success: false, error: '「レポート」シートが見つかりません。', timestamp: getJSTTimestamp() };
    }
    // ヘッダー行以外を全削除
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    // ヘッダーから列位置を動的に検出
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var courseNameIdx = header.indexOf('科目名');
    var codeIdx = header.indexOf('科目コード');
    if (courseNameIdx === -1 || codeIdx === -1) {
      return { success: false, error: 'レポートシートに「科目名」または「科目コード」列が見つかりません。', timestamp: getJSTTimestamp() };
    }
    var numCols = header.length;
    // codesが空でない行のみ一括書き込み
    if (data && data.length > 0) {
      var rows = data.filter(function(item) {
        return item.codes && String(item.codes).trim() !== '';
      }).map(function(item) {
        var row = new Array(numCols).fill('');
        row[courseNameIdx] = item.courseName || '';
        row[codeIdx] = String(item.codes || '').trim();
        return row;
      });
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
      }
    }
    return { success: true, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, 'レポートデータ保存');
  }
}

// ===== 次年度引き継ぎ機能 =====

/**
 * 次年度引き継ぎ処理を一括実行する（管理者専用）
 * @param {object} options - 実行する処理のフラグ
 * @param {boolean} options.upgradeGrade - 学年を繰り上げる
 * @param {boolean} options.resetTimestamp - タイムスタンプをリセット
 * @param {boolean} options.resetStatus - ステータスをリセット（利用停止は維持）
 * @param {boolean} options.resetNextGrade - 来年度学年をリセット
 * @param {boolean} options.resetLotteryIds - 抽選科目の学籍番号をリセット
 * @returns {object} { success: boolean, summary: object }
 */
function executeYearTransition(options) {
  requireAdmin('次年度引き継ぎ');
  try {
    if (!options || typeof options !== 'object') {
      throw new Error('オプションが指定されていません');
    }

    const summary = {};

    // --- 登録データシート処理 ---
    const needsSubmissionSheet = options.upgradeGrade || options.resetTimestamp ||
                                  options.resetStatus || options.resetNextGrade;
    if (needsSubmissionSheet) {
      const sheet = getSubmissionSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      // 列インデックス取得
      const gradeIdx = headers.indexOf('学年');
      const classIdx = headers.indexOf('組');
      const numberIdx = headers.indexOf('番号');
      const tsIdx = headers.indexOf('タイムスタンプ');
      const statusIdx = headers.indexOf('ステータス');
      const nextGradeIdx = headers.indexOf('来年度学年');

      // 必要な列が存在するかチェック
      if (options.upgradeGrade && (gradeIdx === -1 || classIdx === -1 || numberIdx === -1 || nextGradeIdx === -1)) {
        throw new Error('登録データシートに必要な列（学年/組/番号/来年度学年）が見つかりません');
      }
      if (options.resetTimestamp && tsIdx === -1) {
        throw new Error('登録データシートにタイムスタンプ列が見つかりません');
      }
      if (options.resetStatus && statusIdx === -1) {
        throw new Error('登録データシートにステータス列が見つかりません');
      }
      if (options.resetNextGrade && nextGradeIdx === -1) {
        throw new Error('登録データシートに来年度学年列が見つかりません');
      }

      let upgradedCount = 0;
      let timestampResetCount = 0;
      let statusResetCount = 0;
      let nextGradeResetCount = 0;

      // 行ごとに処理（ヘッダー行をスキップ）
      for (let i = 1; i < data.length; i++) {
        // 空行スキップ
        if (!data[i][headers.indexOf('学籍番号')]) continue;

        // 1. 学年繰り上げ（来年度学年リセットより先に処理）
        if (options.upgradeGrade) {
          const nextGrade = data[i][nextGradeIdx];
          if (nextGrade && [1, 2, 3].includes(Number(nextGrade))) {
            data[i][gradeIdx] = Number(nextGrade);
          } else {
            const current = Number(data[i][gradeIdx]);
            if (current >= 1 && current <= 3) {
              data[i][gradeIdx] = Math.min(current + 1, 3);
            }
          }
          data[i][classIdx] = '';
          data[i][numberIdx] = '';
          upgradedCount++;
        }

        // 2. タイムスタンプリセット
        if (options.resetTimestamp) {
          data[i][tsIdx] = '';
          timestampResetCount++;
        }

        // 3. ステータスリセット（利用停止はスキップ）
        if (options.resetStatus) {
          if (data[i][statusIdx] !== '利用停止') {
            data[i][statusIdx] = '';
            statusResetCount++;
          }
        }

        // 4. 来年度学年リセット
        if (options.resetNextGrade) {
          data[i][nextGradeIdx] = '';
          nextGradeResetCount++;
        }
      }

      // 一括書き込み
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

      if (options.upgradeGrade) summary.upgradedCount = upgradedCount;
      if (options.resetTimestamp) summary.timestampResetCount = timestampResetCount;
      if (options.resetStatus) summary.statusResetCount = statusResetCount;
      if (options.resetNextGrade) summary.nextGradeResetCount = nextGradeResetCount;
    }

    // --- 5. 抽選科目の学籍番号リセット（別シート） ---
    if (options.resetLotteryIds) {
      const courseSheet = getCourseSheet();
      const courseData = courseSheet.getDataRange().getValues();
      const courseHeaders = courseData[0];
      const lotteryColIdx = courseHeaders.indexOf('抽選学籍番号');

      if (lotteryColIdx === -1) {
        throw new Error('科目データシートに抽選学籍番号列が見つかりません');
      }

      let lotteryResetCount = 0;
      for (let i = 1; i < courseData.length; i++) {
        if (courseData[i][lotteryColIdx]) {
          courseData[i][lotteryColIdx] = '';
          lotteryResetCount++;
        }
      }

      // 一括書き込み
      courseSheet.getRange(1, 1, courseData.length, courseData[0].length).setValues(courseData);
      summary.lotteryResetCount = lotteryResetCount;

      // 科目データキャッシュクリア
      _memCache.course = null;
      try {
        CacheService.getScriptCache().remove('course_master_data');
      } catch (e) {
        console.warn('科目キャッシュクリア失敗:', e);
      }
    }

    // 登録データキャッシュクリア
    if (needsSubmissionSheet) {
      clearDataCache();
    }

    logMessage('次年度引き継ぎ実行: ' + JSON.stringify(options) + ' / 結果: ' + JSON.stringify(summary));

    return { success: true, summary: summary, timestamp: getJSTTimestamp() };
  } catch (error) {
    return handleError(error, '次年度引き継ぎ');
  }
}
