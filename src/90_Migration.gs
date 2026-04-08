/**
 * 90_Migration.gs - データマイグレーションスクリプト
 *
 * スプレッドシートのスキーマ変更を行うワンタイムスクリプト群。
 * GASエディタから手動実行する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs
 */

// =============================================
// Phase 2-1: 科目IDの導入
// =============================================

/**
 * 科目データシートに「科目ID」列を追加し、既存全科目にIDを自動採番する。
 * 非破壊的: 既存データには影響しない。新しい列が先頭に追加される。
 *
 * ID形式: C001, C002, ... C999
 *
 * 実行方法: GASエディタでこの関数を選択して実行
 */
function migrateAddCourseIds() {
  const sheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
  if (!sheet) {
    console.error('科目データシートが見つかりません');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 既に科目ID列が存在するかチェック
  if (headers[0] === '科目ID' || headers.includes('科目ID')) {
    console.log('科目ID列は既に存在します。スキップします。');
    return;
  }

  // A列に列を挿入
  sheet.insertColumnBefore(1);

  // ヘッダーを設定
  sheet.getRange(1, 1).setValue('科目ID');

  // 各行にIDを採番
  let idCounter = 1;
  for (let i = 1; i < data.length; i++) {
    // 科目名が空の行はスキップ
    const courseNameIdx = headers.indexOf('科目名');
    if (courseNameIdx !== -1 && data[i][courseNameIdx]) {
      const courseId = 'C' + String(idCounter).padStart(3, '0');
      sheet.getRange(i + 1, 1).setValue(courseId);
      idCounter++;
    }
  }

  // キャッシュクリア
  SheetAccess.clearCache(CONFIG.CACHE_KEYS.COURSE);

  console.log(`科目IDを ${idCounter - 1} 件の科目に付与しました。`);
}

/**
 * 科目ID→科目名のマッピングを取得
 * @returns {Map<string, string>} 科目ID → 科目名
 */
function getCourseIdToNameMap() {
  const data = SheetAccess.getCourseData();
  if (data.length <= 1) return new Map();

  const headerMap = SheetUtils.buildHeaderMap(data[0]);
  const idIdx = SheetUtils.getColIndex(headerMap, '科目ID');
  const nameIdx = SheetUtils.getColIndex(headerMap, '科目名');

  if (idIdx === -1 || nameIdx === -1) return new Map();

  const map = new Map();
  for (let i = 1; i < data.length; i++) {
    const id = data[i][idIdx];
    const name = data[i][nameIdx];
    if (id && name) map.set(String(id), String(name));
  }
  return map;
}

/**
 * 科目名→科目IDのマッピングを取得
 * @returns {Map<string, string>} 科目名 → 科目ID
 */
function getCourseNameToIdMap() {
  const idToName = getCourseIdToNameMap();
  const nameToId = new Map();
  idToName.forEach((name, id) => nameToId.set(name, id));
  return nameToId;
}

// =============================================
// Phase 2-2: 登録データの縦持ち化
// =============================================

/**
 * 登録ヘッダーシートと登録明細シートを新規作成する。
 * 既存の横持ち「登録データ」シートは変更しない（並行運用期間）。
 *
 * 実行方法: GASエディタでこの関数を選択して実行
 */
function migrateCreateVerticalSheets() {
  const ss = SheetAccess.getSpreadsheet();

  // 登録ヘッダーシート
  let headerSheet = ss.getSheetByName('登録ヘッダー');
  if (!headerSheet) {
    headerSheet = ss.insertSheet('登録ヘッダー');
    headerSheet.getRange(1, 1, 1, 11).setValues([[
      '学籍番号', '学年', '組', '番号', '名前',
      'メールアドレス', 'ステータス', 'タイムスタンプ',
      '認証コード', '来年度学年', '教職員チェック'
    ]]);
    console.log('登録ヘッダーシートを作成しました');
  } else {
    console.log('登録ヘッダーシートは既に存在します');
  }

  // 登録明細シート
  let detailSheet = ss.getSheetByName('登録明細');
  if (!detailSheet) {
    detailSheet = ss.insertSheet('登録明細');
    detailSheet.getRange(1, 1, 1, 6).setValues([[
      '学籍番号', '科目ID', '科目名', 'マーク', '学年', '更新日時'
    ]]);
    console.log('登録明細シートを作成しました');
  } else {
    console.log('登録明細シートは既に存在します');
  }
}

/**
 * 既存の横持ち「登録データ」シートから縦持ちシートにデータを変換・移行する。
 *
 * 前提: migrateAddCourseIds() と migrateCreateVerticalSheets() が実行済み
 *
 * 実行方法: GASエディタでこの関数を選択して実行
 */
function migrateHorizontalToVertical() {
  // 科目名→IDマップ取得
  const courseNameToId = getCourseNameToIdMap();
  if (courseNameToId.size === 0) {
    console.error('科目IDが設定されていません。先にmigrateAddCourseIds()を実行してください。');
    return;
  }

  // 横持ちデータ読み込み
  const submissionData = SheetAccess.getSubmissionData();
  if (submissionData.length <= 1) {
    console.log('登録データが空です。');
    return;
  }

  const headers = submissionData[0];
  const headerMap = SheetUtils.buildHeaderMap(headers);
  const basicColumnsSet = new Set(CONFIG.BASIC_COLUMNS);

  // ヘッダーシート用データ
  const headerRows = [];
  // 明細シート用データ
  const detailRows = [];

  const timestamp = Logger_.getJSTTimestamp();

  for (let i = 1; i < submissionData.length; i++) {
    const row = submissionData[i];
    const studentId = row[SheetUtils.getColIndex(headerMap, '学籍番号')];
    if (!studentId) continue;

    // ヘッダー行
    headerRows.push([
      studentId,
      row[SheetUtils.getColIndex(headerMap, '学年')] || '',
      row[SheetUtils.getColIndex(headerMap, '組')] || '',
      row[SheetUtils.getColIndex(headerMap, '番号')] || '',
      row[SheetUtils.getColIndex(headerMap, '名前')] || '',
      row[SheetUtils.getColIndex(headerMap, 'メールアドレス')] || '',
      row[SheetUtils.getColIndex(headerMap, 'ステータス')] || '',
      row[SheetUtils.getColIndex(headerMap, 'タイムスタンプ')] || '',
      row[SheetUtils.getColIndex(headerMap, '認証コード')] || '',
      row[SheetUtils.getColIndex(headerMap, '来年度学年')] || '',
      row[SheetUtils.getColIndex(headerMap, '教職員チェック')] || '',
    ]);

    // 明細行（科目カラムを走査）
    headers.forEach((colName, colIdx) => {
      if (!colName || basicColumnsSet.has(colName)) return;
      const value = row[colIdx];
      if (!value || value === '') return;

      // マークと学年を分離: ○1 → マーク=○, 学年=1
      const match = String(value).match(/^([○〇●△])(\d)?(N)?$/);
      if (!match) return;

      const mark = match[1];
      const grade = match[2] || '';
      const courseId = courseNameToId.get(colName) || '';

      detailRows.push([
        studentId,
        courseId,
        colName,
        mark + (match[3] || ''),  // Nサフィックスはマークに含める
        grade,
        timestamp,
      ]);
    });
  }

  // 書き込み
  const ss = SheetAccess.getSpreadsheet();

  if (headerRows.length > 0) {
    const headerSheet = ss.getSheetByName('登録ヘッダー');
    headerSheet.getRange(2, 1, headerRows.length, headerRows[0].length).setValues(headerRows);
    console.log(`登録ヘッダー: ${headerRows.length} 件の生徒データを移行しました`);
  }

  if (detailRows.length > 0) {
    const detailSheet = ss.getSheetByName('登録明細');
    detailSheet.getRange(2, 1, detailRows.length, detailRows[0].length).setValues(detailRows);
    console.log(`登録明細: ${detailRows.length} 件の科目登録データを移行しました`);
  }

  console.log('横持ち→縦持ちマイグレーション完了');
}
