/**
 * 00_Config.gs - 定数・設定値の一元管理
 *
 * システム全体で使用する定数、シート名、ステータス文字列、
 * ロール名、基本カラム名等を一元管理する。
 * 学校固有のデフォルト値もここに集約し、転用時の変更箇所を明確にする。
 */

const CONFIG = Object.freeze({
  // --- キャッシュ設定 ---
  CACHE_EXPIRY_SEC: 300,        // 5分
  CACHE_EXPIRY_MS: 5 * 60 * 1000,
  CACHE_MAX_SIZE: 95000,        // CacheService制限（100KB）に余裕を持たせる
  CACHE_KEYS: Object.freeze({
    SUBMISSION: 'submission_data',
    COURSE:     'course_data',
    TEACHER:    'teacher_data',
    ROSTER:     'roster_data',
    SETTINGS:   'settings_data',
  }),

  // --- 単位制限 ---
  UNIT_LIMIT: 30,

  // --- シート名 ---
  SHEETS: Object.freeze({
    COURSE:     '科目データ',
    SUBMISSION: '登録データ',
    TEACHER:    '教職員データ',
    ROSTER:     '生徒情報',
    SETTINGS:   '設定',
    // Phase 2: 縦持ちシート
    REG_HEADER: '登録ヘッダー',
    REG_DETAIL: '登録明細',
    // 後回し（コア機能外）
    // TEXTBOOK:   '教科書',
    // REPORT:     'レポート',
    // SM_SETTINGS:'SM設定',
  }),

  // --- 登録データの基本カラム（科目カラムとの区別に使用）---
  BASIC_COLUMNS: Object.freeze([
    'タイムスタンプ', 'ステータス', '認証コード',
    '学籍番号', '学年', '組', '番号', '名前',
    'メールアドレス', '教職員チェック', '来年度学年'
  ]),

  // --- ステータス文字列 ---
  STATUSES: Object.freeze({
    DRAFT:      '保存',
    PROVISIONAL:'仮登録',
    APPROVED:   '本登録許可',
    FINAL:      '本登録',
    REVERTED:   '差戻し',
    SUSPENDED:  '利用停止',
  }),

  // --- ユーザーロール ---
  ROLES: Object.freeze({
    STUDENT: 'student',
    TEACHER: 'teacher',
    KYOMU:   'kyomu',
    ADMIN:   'admin',
  }),

  // --- 設定シートのキー名 ---
  SETTING_KEYS: Object.freeze({
    SCHOOL_NAME:        '学校名',
    FINAL_REG_ALLOWED:  '本登録許可',
    PROVISIONAL_REG:    '仮登録',
    TEMP_SAVE:          '一時保存',
    MAINTENANCE:        'メンテナンスモード',
    STUDENT_DOMAIN:     '生徒用ドメイン',
    DISPLAY_MODE:       '生徒への表示',
    CIRCLE_DISPLAY:     '○マーク履修中表示',
    MIN_UNITS_Y1:       '1年最低単位数',
    MIN_UNITS_Y2:       '2年最低単位数',
    MIN_UNITS_Y3:       '3年最低単位数',
    GRAD_REQ_SCIENCE:   '卒業要件_理科',
  }),

  // --- 学校固有デフォルト値（転用時にここを変更）---
  DEFAULTS: Object.freeze({
    SCHOOL_NAME: '学校名を設定してください',
    TIMEZONE:    'Asia/Tokyo',
    DATE_FORMAT: 'yyyy-MM-dd HH:mm:ss',
  }),

  // --- 初期設定データ（設定シート自動作成時）---
  INITIAL_SETTINGS: Object.freeze([
    ['設定名', '内容'],
    ['学校名', ''],
    ['本登録許可', '不可'],
    ['仮登録', '不可'],
    ['1年最低単位数', '17'],
    ['2年最低単位数', '44'],
    ['3年最低単位数', '74'],
    ['生徒用ドメイン', ''],
    ['メンテナンスモード', '無効'],
    ['一時保存', '不可'],
    ['○マーク履修中表示', '有効'],
    ['生徒への表示', 'クラス'],
  ]),
});
