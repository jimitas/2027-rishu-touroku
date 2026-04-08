/**
 * 10_DataTransformer.gs - データ変換レイヤー
 *
 * 縦持ち（登録ヘッダー + 登録明細）と横持ち（マトリクス表示）の
 * 相互変換を行う。
 *
 * フロントエンドのマトリクス表示は横持ちのまま維持しつつ、
 * データストアは縦持ちに移行するためのアダプター層。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs
 */

const DataTransformer = {

  // --- シート名定数（Phase 2 完了後にCONFIG.SHEETSに移動予定）---
  HEADER_SHEET: '登録ヘッダー',
  DETAIL_SHEET: '登録明細',

  // --- 縦持ち → 横持ち変換 ---

  /**
   * 特定の生徒の縦持ちデータを横持ちオブジェクトに変換
   * フロントエンドの科目マトリクス表示用
   *
   * @param {string} studentId - 学籍番号
   * @returns {Object} { '現代の国語': '○1', '言語文化': '●1', ... }
   */
  toHorizontal(studentId) {
    const detailData = this._getDetailData();
    if (detailData.length <= 1) return {};

    const headerMap = SheetUtils.buildHeaderMap(detailData[0]);
    const sidIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
    const nameIdx = SheetUtils.getColIndex(headerMap, '科目名');
    const markIdx = SheetUtils.getColIndex(headerMap, 'マーク');
    const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');

    if (sidIdx === -1 || nameIdx === -1 || markIdx === -1) return {};

    const result = {};
    const searchStr = String(studentId);

    for (let i = 1; i < detailData.length; i++) {
      if (String(detailData[i][sidIdx]) !== searchStr) continue;

      const courseName = detailData[i][nameIdx];
      const mark = detailData[i][markIdx] || '';
      const grade = gradeIdx !== -1 ? (detailData[i][gradeIdx] || '') : '';

      if (courseName && mark) {
        // 旧形式互換: ○ + 1 → ○1
        result[String(courseName)] = mark + grade;
      }
    }
    return result;
  },

  /**
   * 横持ちデータを縦持ちの行配列に変換（保存時）
   *
   * @param {string} studentId - 学籍番号
   * @param {Object} horizontalData - { '現代の国語': '○1', ... }
   * @returns {Array<Array>} 明細行の配列（ヘッダーなし）
   */
  toVertical(studentId, horizontalData) {
    const courseNameToId = getCourseNameToIdMap();
    const timestamp = Logger_.getJSTTimestamp();
    const rows = [];

    Object.entries(horizontalData).forEach(([courseName, value]) => {
      if (!value || value === '') return;

      // マーク+学年を分離: ○1N → mark=○N, grade=1
      const match = String(value).match(/^([○〇●△])([\d])?(N)?$/);
      if (!match) return;

      const mark = match[1] + (match[3] || ''); // ○ or ○N
      const grade = match[2] || '';
      const courseId = courseNameToId.get(courseName) || '';

      rows.push([
        studentId,
        courseId,
        courseName,
        mark,
        grade,
        timestamp,
      ]);
    });

    return rows;
  },

  // --- 全生徒マトリクス（教員画面用）---

  /**
   * 全生徒の横持ちマトリクスを生成
   * 教員画面の一覧表示用
   *
   * @param {Array<Object>} courseList - 科目オブジェクト配列
   * @returns {{ headers: Array<string>, students: Array<Object> }}
   */
  buildMatrix(courseList) {
    const headerData = this._getHeaderData();
    const detailData = this._getDetailData();
    if (headerData.length <= 1) return { headers: [], students: [] };

    const courseNames = courseList.map(c => c['科目名']).filter(Boolean);
    const matrixHeaders = ['学籍番号', '名前', '学年', '組', '番号', 'ステータス', ...courseNames];

    // ヘッダーシートから生徒基本情報を構築
    const hdrMap = SheetUtils.buildHeaderMap(headerData[0]);
    const studentMap = new Map();

    for (let i = 1; i < headerData.length; i++) {
      const row = headerData[i];
      const sid = String(row[SheetUtils.getColIndex(hdrMap, '学籍番号')] || '');
      if (!sid) continue;

      studentMap.set(sid, {
        '学籍番号': sid,
        '名前': row[SheetUtils.getColIndex(hdrMap, '名前')] || '',
        '学年': row[SheetUtils.getColIndex(hdrMap, '学年')] || '',
        '組': row[SheetUtils.getColIndex(hdrMap, '組')] || '',
        '番号': row[SheetUtils.getColIndex(hdrMap, '番号')] || '',
        'ステータス': row[SheetUtils.getColIndex(hdrMap, 'ステータス')] || '',
      });
    }

    // 明細データを各生徒に集約
    if (detailData.length > 1) {
      const detMap = SheetUtils.buildHeaderMap(detailData[0]);
      const dSidIdx = SheetUtils.getColIndex(detMap, '学籍番号');
      const dNameIdx = SheetUtils.getColIndex(detMap, '科目名');
      const dMarkIdx = SheetUtils.getColIndex(detMap, 'マーク');
      const dGradeIdx = SheetUtils.getColIndex(detMap, '学年');

      for (let i = 1; i < detailData.length; i++) {
        const row = detailData[i];
        const sid = String(row[dSidIdx] || '');
        const courseName = row[dNameIdx];
        const mark = row[dMarkIdx] || '';
        const grade = dGradeIdx !== -1 ? (row[dGradeIdx] || '') : '';

        if (studentMap.has(sid) && courseName) {
          studentMap.get(sid)[String(courseName)] = mark + grade;
        }
      }
    }

    return { headers: matrixHeaders, students: Array.from(studentMap.values()) };
  },

  // --- 縦持ちシートへの書き込み ---

  /**
   * 生徒のヘッダー情報を保存/更新
   *
   * @param {string} studentId - 学籍番号
   * @param {Object} basicInfo - { 学年, 組, 番号, 名前, メール, ステータス, ... }
   */
  saveHeader(studentId, basicInfo) {
    const ss = SheetAccess.getSpreadsheet();
    const sheet = ss.getSheetByName(this.HEADER_SHEET);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const headerMap = SheetUtils.buildHeaderMap(data[0]);
    const found = SheetUtils.findRow(data, headerMap, '学籍番号', studentId);

    const row = [
      studentId,
      basicInfo['学年'] || '',
      basicInfo['組'] || '',
      basicInfo['番号'] || '',
      basicInfo['名前'] || '',
      basicInfo['メールアドレス'] || '',
      basicInfo['ステータス'] || '',
      Logger_.getJSTTimestamp(),
      basicInfo['認証コード'] || '',
      basicInfo['来年度学年'] || '',
      basicInfo['教職員チェック'] || '',
    ];

    if (found) {
      sheet.getRange(found.rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  },

  /**
   * 生徒の科目明細を保存（全置換方式）
   *
   * @param {string} studentId - 学籍番号
   * @param {Object} horizontalData - { '科目名': 'マーク学年', ... }
   */
  saveDetail(studentId, horizontalData) {
    const ss = SheetAccess.getSpreadsheet();
    const sheet = ss.getSheetByName(this.DETAIL_SHEET);
    if (!sheet) return;

    // 既存の当該生徒の明細を削除
    const data = sheet.getDataRange().getValues();
    const headerMap = SheetUtils.buildHeaderMap(data[0]);
    const sidIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

    if (sidIdx !== -1) {
      // 下から走査して削除（行番号のずれを防止）
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][sidIdx]) === String(studentId)) {
          sheet.deleteRow(i + 1);
        }
      }
    }

    // 新しい明細を追加
    const rows = this.toVertical(studentId, horizontalData);
    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  },

  // --- 内部ヘルパー ---

  /**
   * 登録ヘッダーシートのデータを取得
   * @private
   */
  _getHeaderData() {
    const ss = SheetAccess.getSpreadsheet();
    const sheet = ss.getSheetByName(this.HEADER_SHEET);
    if (!sheet) return [];
    return sheet.getDataRange().getValues();
  },

  /**
   * 登録明細シートのデータを取得
   * @private
   */
  _getDetailData() {
    const ss = SheetAccess.getSpreadsheet();
    const sheet = ss.getSheetByName(this.DETAIL_SHEET);
    if (!sheet) return [];
    return sheet.getDataRange().getValues();
  },
};
