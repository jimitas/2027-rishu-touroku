/**
 * 09_AdminService.gs - 管理者向けサービス
 *
 * 生徒データの一括管理、テスト結果取込、次年度引き継ぎなど、
 * 管理者専用の操作を提供する。
 *
 * 依存: 00_Config.gs, 01_Utils.gs, 02_SheetAccess.gs, 03_Auth.gs,
 *        04_CourseService.gs, 06_StudentService.gs
 */

const AdminService = {

  /**
   * 学籍移動による生徒行の物理削除
   * @param {Array<string>} deleteStudentIds - 削除する学籍番号の配列
   * @returns {{ success: boolean, deletedStudents?: number, error?: string }}
   */
  deleteStudentRows(deleteStudentIds) {
    try {
      Auth.requireAdmin('生徒行削除');

      if (!deleteStudentIds || !Array.isArray(deleteStudentIds) || deleteStudentIds.length === 0) {
        throw new Error('削除対象の学籍番号が指定されていません');
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      const allData = submissionSheet.getDataRange().getValues();
      const headers = allData[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const studentIdColIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

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
      deleteStudentIds.forEach(sid => {
        const idx = studentRowMap[String(sid).trim()];
        if (idx !== undefined) {
          rowsToDelete.push(idx + 1); // シートの1ベース行番号
        }
      });

      // 下から削除（行番号ずれ防止）
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach(sheetRow => {
        submissionSheet.deleteRow(sheetRow);
      });

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log('生徒行削除完了: ' + rowsToDelete.length + '名削除（学籍移動）');

      return Logger_.successResponse({ deletedStudents: rowsToDelete.length });
    } catch (error) {
      return Logger_.handleError(error, '生徒行削除');
    }
  },

  /**
   * 生徒名簿データを取得
   * @returns {{ success: boolean, students: Array }}
   */
  getStudentRosterForAdmin() {
    Auth.requireTeacherOrAdmin('生徒名簿取得');
    try {
      const data = SheetAccess.getRosterData();
      if (!data || data.length < 2) return { success: true, students: [] };

      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);
      const idIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      const nameIdx = SheetUtils.getColIndex(headerMap, '氏名');
      const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
      const classIdx = SheetUtils.getColIndex(headerMap, '組');
      const numIdx = SheetUtils.getColIndex(headerMap, '番号');

      const students = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (idIdx === -1 || !row[idIdx]) continue;
        students.push({
          studentId: String(row[idIdx]),
          name: nameIdx !== -1 ? (row[nameIdx] || '') : '',
          grade: gradeIdx !== -1 ? (row[gradeIdx] || '') : '',
          class: classIdx !== -1 ? (row[classIdx] || '') : '',
          number: numIdx !== -1 ? (row[numIdx] || '') : ''
        });
      }
      return { success: true, students: students };
    } catch (error) {
      return Logger_.handleError(error, '生徒名簿取得');
    }
  },

  /**
   * 生徒一覧を取得（登録データシートから基本列のみ）
   * @returns {{ success: boolean, students: Array, headers: Array }}
   */
  getStudentList() {
    Auth.requireAdmin('生徒一覧取得');
    try {
      const submissionData = SheetAccess.getSubmissionData();
      if (!submissionData || submissionData.length === 0) {
        return { success: true, students: [], headers: [] };
      }
      const subHeaders = submissionData[0];
      const headerMap = SheetUtils.buildHeaderMap(subHeaders);

      const fields = ['学籍番号', '学年', '組', '番号', '名前', 'メールアドレス', '来年度学年'];
      const fieldIndices = {};
      fields.forEach(f => {
        const idx = SheetUtils.getColIndex(headerMap, f);
        if (idx !== -1) fieldIndices[f] = idx;
      });

      const idIdx = fieldIndices['学籍番号'];
      if (idIdx === undefined) {
        throw new Error('登録データシートに学籍番号列が見つかりません');
      }

      const students = [];
      for (let i = 1; i < submissionData.length; i++) {
        const row = submissionData[i];
        const studentId = String(row[idIdx] || '').trim();
        if (!studentId) continue;

        const student = {};
        for (const key in fieldIndices) {
          const val = row[fieldIndices[key]];
          student[key] = val !== undefined && val !== null && val !== '' ? String(val) : '';
        }
        students.push(student);
      }

      return { success: true, students: students, headers: fields };
    } catch (error) {
      return Logger_.handleError(error, '生徒一覧取得');
    }
  },

  /**
   * 登録データシートのデータ行を学年→組→番号で昇順ソート
   * @param {SpreadsheetApp.Sheet} sheet - 登録データシート
   * @param {Array} headers - ヘッダー行の配列
   * @private
   */
  _sortSubmissionSheet(sheet, headers) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    const data = dataRange.getValues();

    // タイムスタンプ列のDate→String変換によるずれを防止
    const headerMap = SheetUtils.buildHeaderMap(headers);
    const tsIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');
    if (tsIdx !== -1) {
      const displayValues = dataRange.getDisplayValues();
      for (let i = 0; i < data.length; i++) {
        data[i][tsIdx] = displayValues[i][tsIdx];
      }
    }

    const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
    const classIdx = SheetUtils.getColIndex(headerMap, '組');
    const numIdx = SheetUtils.getColIndex(headerMap, '番号');
    if (gradeIdx === -1 || classIdx === -1 || numIdx === -1) return;

    data.sort((a, b) => {
      const aClass = String(a[classIdx] || '').trim();
      const bClass = String(b[classIdx] || '').trim();
      const aIsNum = !aClass || /^\d+$/.test(aClass);
      const bIsNum = !bClass || /^\d+$/.test(bClass);
      if (aIsNum !== bIsNum) return aIsNum ? -1 : 1;

      const aGrade = Validator.safeParseGrade(a[gradeIdx], 999);
      const bGrade = Validator.safeParseGrade(b[gradeIdx], 999);
      if (aGrade !== bGrade) return aGrade - bGrade;

      if (!aClass && bClass) return 1;
      if (aClass && !bClass) return -1;
      if (aIsNum && bIsNum) {
        const classDiff = parseInt(aClass) - parseInt(bClass);
        if (classDiff !== 0) return classDiff;
      } else {
        const classCmp = aClass.localeCompare(bClass);
        if (classCmp !== 0) return classCmp;
      }

      const aNum = parseInt(a[numIdx]) || 999;
      const bNum = parseInt(b[numIdx]) || 999;
      return aNum - bNum;
    });

    dataRange.setValues(data);
  },

  /**
   * 生徒データを一括保存（登録データシートの基本列を差分更新・新規追加・削除）
   * @param {Array<Object>} studentArray - 生徒データの配列
   * @returns {{ success: boolean, count?: number, error?: string }}
   */
  replaceAllStudents(studentArray) {
    Auth.requireAdmin('生徒データ一括保存');
    try {
      if (!Array.isArray(studentArray) || studentArray.length === 0) {
        throw new Error('生徒データが空です');
      }

      const idSet = {};
      for (let i = 0; i < studentArray.length; i++) {
        const studentId = String(studentArray[i]['学籍番号'] || '').trim();
        if (!studentId) {
          throw new Error((i + 1) + '行目: 学籍番号が空です');
        }
        if (idSet[studentId]) {
          throw new Error('学籍番号が重複しています: ' + studentId);
        }
        idSet[studentId] = true;
        const grade = String(studentArray[i]['学年'] || '').trim();
        if (grade && grade !== '0' && grade !== '1' && grade !== '2' && grade !== '3') {
          throw new Error((i + 1) + '行目: 学年は0/1/2/3のいずれかで入力してください（' + grade + '）');
        }
        const nextGrade = String(studentArray[i]['来年度学年'] || '').trim();
        if (nextGrade && nextGrade !== '1' && nextGrade !== '2' && nextGrade !== '3') {
          throw new Error((i + 1) + '行目: 来年度学年は1/2/3のいずれかで入力してください（' + nextGrade + '）');
        }
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }
      const subData = submissionSheet.getDataRange().getValues();
      if (subData.length === 0) {
        throw new Error('登録データシートにヘッダー行がありません');
      }
      const subHeaders = subData[0];
      const headerMap = SheetUtils.buildHeaderMap(subHeaders);
      const subIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      if (subIdIdx === -1) {
        throw new Error('登録データシートに学籍番号列が見つかりません');
      }

      const syncColumns = ['学年', '組', '番号', '名前', 'メールアドレス', '来年度学年'];
      const syncColIndices = {};
      syncColumns.forEach(col => {
        const idx = SheetUtils.getColIndex(headerMap, col);
        if (idx !== -1) syncColIndices[col] = idx;
      });

      // 学籍番号→生徒データのマップ
      const newStudentMap = {};
      studentArray.forEach(s => {
        const sid = String(s['学籍番号'] || '').trim();
        if (sid) newStudentMap[sid] = s;
      });

      // 既存行の学籍番号セット
      const existingIdSet = {};
      for (let k = 1; k < subData.length; k++) {
        const existingId = String(subData[k][subIdIdx] || '').trim();
        if (existingId) existingIdSet[existingId] = true;
      }

      // 1. 既存行の基本列を差分更新
      for (let k = 1; k < subData.length; k++) {
        const existingId = String(subData[k][subIdIdx] || '').trim();
        if (existingId && newStudentMap.hasOwnProperty(existingId)) {
          const student = newStudentMap[existingId];
          let rowChanged = false;
          const row = subData[k].slice();
          for (const col in syncColIndices) {
            const colIdx = syncColIndices[col];
            const currentVal = String(row[colIdx] || '').trim();
            const newVal = String(student[col] || '').trim();
            if (currentVal !== newVal) {
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

      // 2. 新規生徒の行追加
      const newRows = [];
      studentArray.forEach(s => {
        const newId = String(s['学籍番号'] || '').trim();
        if (newId && !existingIdSet[newId]) {
          const newRow = subHeaders.map(header => {
            if (s.hasOwnProperty(header)) {
              const val = String(s[header] || '').trim();
              if ((header === '学年' || header === '組' || header === '番号' || header === '来年度学年') && val !== '') {
                return /^\d+$/.test(val) ? Number(val) : val;
              }
              return val;
            }
            return '';
          });
          newRows.push(newRow);
        }
      });
      if (newRows.length > 0) {
        const lastRow = submissionSheet.getLastRow();
        submissionSheet.getRange(lastRow + 1, 1, newRows.length, subHeaders.length).setValues(newRows);
      }

      // 3. 削除された生徒の行削除（下から上へ）
      let deletedCount = 0;
      for (let d = subData.length - 1; d >= 1; d--) {
        const delId = String(subData[d][subIdIdx] || '').trim();
        if (delId && !newStudentMap.hasOwnProperty(delId)) {
          submissionSheet.deleteRow(d + 1);
          deletedCount++;
        }
      }

      // 4. ソート
      this._sortSubmissionSheet(submissionSheet, subHeaders);

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('生徒データ一括保存: ' + studentArray.length + '件（新規: ' + newRows.length + '件、削除: ' + deletedCount + '件）');
      return { success: true, count: studentArray.length };
    } catch (error) {
      return Logger_.handleError(error, '生徒データ一括保存');
    }
  },

  /**
   * 新入生データを登録データシートに追加
   * @param {Array<Object>} studentArray - 新入生データ
   * @param {boolean} autoMark - ○1を自動付与するか
   * @param {boolean} finalRegistration - 本登録ステータスで登録するか
   * @returns {{ success: boolean, addedCount?: number, skippedDuplicates?: Array, error?: string }}
   */
  importNewStudents(studentArray, autoMark, finalRegistration) {
    Auth.requireAdmin('新入生データ取込');
    try {
      if (!Array.isArray(studentArray) || studentArray.length === 0) {
        throw new Error('新入生データが空です');
      }

      const idSet = {};
      for (let i = 0; i < studentArray.length; i++) {
        const studentId = String(studentArray[i]['学籍番号'] || '').trim();
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
        const grade = String(studentArray[i]['学年'] || '0').trim();
        if (grade !== '0' && grade !== '1') {
          throw new Error((i + 1) + '行目: 学年は0または1のみ指定可能です（' + grade + '）');
        }
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }
      const subData = submissionSheet.getDataRange().getValues();
      if (subData.length === 0) {
        throw new Error('登録データシートにヘッダー行がありません');
      }
      const subHeaders = subData[0];
      const headerMap = SheetUtils.buildHeaderMap(subHeaders);
      const subIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      if (subIdIdx === -1) {
        throw new Error('登録データシートに学籍番号列が見つかりません');
      }

      // 既存学籍番号セット
      const existingIdSet = {};
      for (let k = 1; k < subData.length; k++) {
        const existingId = String(subData[k][subIdIdx] || '').trim();
        if (existingId) existingIdSet[existingId] = true;
      }

      // 新規行を構築
      const newRows = [];
      const skippedDuplicates = [];
      const statusColIdx = SheetUtils.getColIndex(headerMap, 'ステータス');

      for (let n = 0; n < studentArray.length; n++) {
        const newId = String(studentArray[n]['学籍番号'] || '').trim();
        if (existingIdSet[newId]) {
          skippedDuplicates.push(newId);
          continue;
        }
        const newRow = subHeaders.map(header => {
          if (header === 'ステータス' && finalRegistration) {
            return CONFIG.STATUSES.FINAL;
          }
          if (studentArray[n].hasOwnProperty(header)) {
            const val = String(studentArray[n][header] || '').trim();
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
        const courseData = SheetAccess.getCourseData();
        if (courseData && courseData.length > 1) {
          const courseHeaders = courseData[0];
          const courseHeaderMap = SheetUtils.buildHeaderMap(courseHeaders);
          const courseNameIdx = SheetUtils.getColIndex(courseHeaderMap, '科目名');
          const courseGradeIdx = SheetUtils.getColIndex(courseHeaderMap, '学年');
          const courseCategoryIdx = SheetUtils.getColIndex(courseHeaderMap, '区分');
          const courseNoOpenIdx = SheetUtils.getColIndex(courseHeaderMap, '開講なし');
          const courseVisShowIdx = SheetUtils.getColIndex(courseHeaderMap, '表示学籍番号');
          const courseVisHideIdx = SheetUtils.getColIndex(courseHeaderMap, '非表示学籍番号');

          // ○1対象の科目リストを構築
          const autoMarkCourses = [];
          for (let ci = 1; ci < courseData.length; ci++) {
            const cRow = courseData[ci];
            const cGrade = courseGradeIdx !== -1 ? String(cRow[courseGradeIdx] ?? '').trim() : '';
            const cCategory = courseCategoryIdx !== -1 ? String(cRow[courseCategoryIdx] ?? '').trim() : '';
            const cNoOpen = courseNoOpenIdx !== -1 ? String(cRow[courseNoOpenIdx] ?? '').trim() : '';
            const cName = courseNameIdx !== -1 ? String(cRow[courseNameIdx] ?? '').trim() : '';
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
            const courseColMap = {};
            subHeaders.forEach((h, i) => { courseColMap[h] = i; });
            const gradeColIdx = SheetUtils.getColIndex(headerMap, '学年');
            const idColIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

            for (let ri = 0; ri < newRows.length; ri++) {
              const rowGrade = String(newRows[ri][gradeColIdx] ?? '');
              if (rowGrade === '0' || rowGrade === '1' || newRows[ri][gradeColIdx] === 0 || newRows[ri][gradeColIdx] === 1) {
                const studentIdForMark = String(newRows[ri][idColIdx] ?? '').trim();
                for (let mi = 0; mi < autoMarkCourses.length; mi++) {
                  const course = autoMarkCourses[mi];
                  if (course.hideIds) {
                    const excludedIds = Validator.splitCSV(course.hideIds);
                    if (excludedIds.indexOf(studentIdForMark) !== -1) continue;
                  }
                  if (course.showIds) {
                    const targetIds = Validator.splitCSV(course.showIds);
                    if (targetIds.indexOf(studentIdForMark) === -1) continue;
                  }
                  const colIdx = courseColMap[course.name];
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
        const lastRow = submissionSheet.getLastRow();
        submissionSheet.getRange(lastRow + 1, 1, newRows.length, subHeaders.length).setValues(newRows);
      }

      // ソート
      this._sortSubmissionSheet(submissionSheet, subHeaders);

      // キャッシュクリア
      SheetAccess.clearAllCaches();

      Logger_.log('新入生データ取込: 追加' + newRows.length + '件、重複スキップ' + skippedDuplicates.length + '件');
      return {
        success: true,
        addedCount: newRows.length,
        skippedDuplicates: skippedDuplicates
      };
    } catch (error) {
      return Logger_.handleError(error, '新入生データ取込');
    }
  },

  /**
   * テスト結果Excelデータを登録データシートに取り込み
   * @param {Object} importData - { 学籍番号: { 科目名: 学年(数値), ... }, ... }
   * @returns {{ success: boolean, updatedStudents?: number, updatedCells?: number, skippedCells?: number, notFoundStudents?: Array, error?: string }}
   */
  importTestResults(importData) {
    try {
      Auth.requireAdmin('テストデータ取込');

      if (!importData || typeof importData !== 'object') {
        throw new Error('取込データが無効です');
      }

      const submissionSheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
      if (!submissionSheet) {
        throw new Error('登録データシートが見つかりません');
      }

      // シートから直接読む（キャッシュではなく最新データ）
      const allData = submissionSheet.getDataRange().getValues();
      const headers = allData[0];
      const lastCol = headers.length;
      const headerMap = SheetUtils.buildHeaderMap(headers);

      const studentIdColIdx = SheetUtils.getColIndex(headerMap, '学籍番号');
      if (studentIdColIdx === -1) {
        throw new Error('登録データシートに学籍番号列が見つかりません');
      }

      // 学籍番号→allDataインデックスのマッピング構築
      const studentRowMap = SheetUtils.buildLookupMap(allData, headerMap, '学籍番号');

      const studentIds = Object.keys(importData);
      let updatedStudents = 0;
      let updatedCells = 0;
      let skippedCells = 0;
      const notFoundStudents = [];
      const updatedRowIndices = new Set();

      // 行インデックスをstudentRowMapから再構築（buildLookupMapは行データを返すため）
      const studentIdxMap = {};
      for (let i = 1; i < allData.length; i++) {
        const sid = String(allData[i][studentIdColIdx]).trim();
        if (sid) studentIdxMap[sid] = i;
      }

      studentIds.forEach(studentId => {
        const sid = String(studentId).trim();
        const dataIndex = studentIdxMap[sid];

        if (dataIndex === undefined) {
          notFoundStudents.push(sid);
          return;
        }

        const rowData = allData[dataIndex];
        let rowUpdated = false;

        const subjects = importData[studentId];
        Object.keys(subjects).forEach(subjectName => {
          const newGrade = parseInt(subjects[subjectName]);
          if (!newGrade || newGrade < 1 || newGrade > 3) return;

          const colIdx = SheetUtils.getColIndex(headerMap, subjectName);
          if (colIdx === -1 || CONFIG.BASIC_COLUMNS.includes(subjectName)) return;

          const existingMark = String(rowData[colIdx] || '').trim();

          let existingGrade = 0;
          if (existingMark.includes('●')) {
            const match = existingMark.match(/●(\d)/);
            if (match) {
              existingGrade = parseInt(match[1]);
            } else if (existingMark === '●') {
              existingGrade = 99;
            }
          }

          if (existingGrade >= newGrade) {
            skippedCells++;
            return;
          }

          rowData[colIdx] = '●' + newGrade;
          rowUpdated = true;
          updatedCells++;
        });

        if (rowUpdated) {
          updatedStudents++;
          updatedRowIndices.add(dataIndex);
        }
      });

      // 更新行のみ書き戻し
      updatedRowIndices.forEach(rowIndex => {
        submissionSheet.getRange(rowIndex + 1, 1, 1, lastCol).setValues([allData[rowIndex]]);
      });

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.SUBMISSION);

      Logger_.log('テストデータ取込完了: ' + updatedStudents + '名, ' + updatedCells + '件更新');

      return Logger_.successResponse({
        updatedStudents: updatedStudents,
        updatedCells: updatedCells,
        skippedCells: skippedCells,
        notFoundStudents: notFoundStudents
      });

    } catch (error) {
      return Logger_.handleError(error, 'テストデータ取込');
    }
  },

  /**
   * 科目データの学籍番号列（抽選/非表示/表示）を一括置換
   * @param {Array<{oldId: string, newId: string}>} replacements - 置換情報
   * @returns {{ success: boolean, updatedCount?: number, error?: string }}
   */
  replaceStudentIdInCourseSettings(replacements) {
    Auth.requireAdmin('科目設定学籍番号置換');
    try {
      if (!Array.isArray(replacements) || replacements.length === 0) {
        return { success: true, updatedCount: 0 };
      }

      const sheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const headerMap = SheetUtils.buildHeaderMap(headers);

      // 対象列のインデックスを取得
      const targetColumns = ['抽選学籍番号', '非表示学籍番号', '表示学籍番号'];
      const colIndices = [];
      targetColumns.forEach(col => {
        const idx = SheetUtils.getColIndex(headerMap, col);
        if (idx !== -1) colIndices.push(idx);
      });
      if (colIndices.length === 0) {
        return { success: true, updatedCount: 0 };
      }

      // 置換マップを構築
      const replaceMap = {};
      replacements.forEach(r => {
        const oldId = String(r.oldId).trim();
        const newId = String(r.newId).trim();
        if (oldId && newId && oldId !== newId) {
          replaceMap[oldId] = newId;
        }
      });
      if (Object.keys(replaceMap).length === 0) {
        return { success: true, updatedCount: 0 };
      }

      // 全行を走査して置換
      let updatedCount = 0;
      for (let i = 1; i < data.length; i++) {
        for (let j = 0; j < colIndices.length; j++) {
          const ci = colIndices[j];
          const cellValue = String(data[i][ci] || '').trim();
          if (!cellValue) continue;

          const ids = Validator.splitCSV(cellValue);
          let changed = false;
          for (let k = 0; k < ids.length; k++) {
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

      // キャッシュクリア
      SheetAccess.clearCache(CONFIG.CACHE_KEYS.COURSE);

      Logger_.log('科目設定学籍番号置換: ' + updatedCount + '件更新');

      return { success: true, updatedCount: updatedCount };
    } catch (error) {
      return Logger_.handleError(error, '科目設定学籍番号置換');
    }
  },

  /**
   * 次年度引き継ぎ処理を一括実行する
   * @param {Object} options - 実行する処理のフラグ
   * @returns {{ success: boolean, summary?: Object, error?: string }}
   */
  executeYearTransition(options) {
    Auth.requireAdmin('次年度引き継ぎ');
    try {
      if (!options || typeof options !== 'object') {
        throw new Error('オプションが指定されていません');
      }

      const summary = {};

      // --- 登録データシート処理 ---
      const needsSubmissionSheet = options.upgradeGrade || options.resetTimestamp ||
                                    options.resetStatus || options.resetNextGrade;
      if (needsSubmissionSheet) {
        const sheet = SheetAccess.getSheet(CONFIG.SHEETS.SUBMISSION);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const headerMap = SheetUtils.buildHeaderMap(headers);

        const gradeIdx = SheetUtils.getColIndex(headerMap, '学年');
        const classIdx = SheetUtils.getColIndex(headerMap, '組');
        const numberIdx = SheetUtils.getColIndex(headerMap, '番号');
        const tsIdx = SheetUtils.getColIndex(headerMap, 'タイムスタンプ');
        const statusIdx = SheetUtils.getColIndex(headerMap, 'ステータス');
        const nextGradeIdx = SheetUtils.getColIndex(headerMap, '来年度学年');
        const studentIdIdx = SheetUtils.getColIndex(headerMap, '学籍番号');

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

        for (let i = 1; i < data.length; i++) {
          if (studentIdIdx !== -1 && !data[i][studentIdIdx]) continue;

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

          if (options.resetTimestamp) {
            data[i][tsIdx] = '';
            timestampResetCount++;
          }

          if (options.resetStatus) {
            if (data[i][statusIdx] !== CONFIG.STATUSES.SUSPENDED) {
              data[i][statusIdx] = '';
              statusResetCount++;
            }
          }

          if (options.resetNextGrade) {
            data[i][nextGradeIdx] = '';
            nextGradeResetCount++;
          }
        }

        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

        if (options.upgradeGrade) summary.upgradedCount = upgradedCount;
        if (options.resetTimestamp) summary.timestampResetCount = timestampResetCount;
        if (options.resetStatus) summary.statusResetCount = statusResetCount;
        if (options.resetNextGrade) summary.nextGradeResetCount = nextGradeResetCount;
      }

      // --- 抽選科目の学籍番号リセット ---
      if (options.resetLotteryIds) {
        const courseSheet = SheetAccess.getSheet(CONFIG.SHEETS.COURSE);
        const courseData = courseSheet.getDataRange().getValues();
        const courseHeaders = courseData[0];
        const courseHeaderMap = SheetUtils.buildHeaderMap(courseHeaders);
        const lotteryColIdx = SheetUtils.getColIndex(courseHeaderMap, '抽選学籍番号');

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

        courseSheet.getRange(1, 1, courseData.length, courseData[0].length).setValues(courseData);
        summary.lotteryResetCount = lotteryResetCount;

        SheetAccess.clearCache(CONFIG.CACHE_KEYS.COURSE);
      }

      // キャッシュクリア
      if (needsSubmissionSheet) {
        SheetAccess.clearAllCaches();
      }

      Logger_.log('次年度引き継ぎ実行: ' + JSON.stringify(options) + ' / 結果: ' + JSON.stringify(summary));

      return Logger_.successResponse({ summary: summary });
    } catch (error) {
      return Logger_.handleError(error, '次年度引き継ぎ');
    }
  },

  // ===== 教科書/レポート/SM設定関連 =====
  // 後回し: コア機能外のため、Phase 2以降で実装予定

  /** @returns {Object} 教科書データ */
  getTextbookData() {
    // TODO: 後回し - 教科書シートからデータを取得
    return { success: false, error: '未実装（後回し）' };
  },

  /** @param {Array} data @returns {Object} */
  saveTextbookData(data) {
    // TODO: 後回し - 教科書データを保存
    return { success: false, error: '未実装（後回し）' };
  },

  /** @returns {Object} レポートデータ */
  getReportData() {
    // TODO: 後回し - レポートシートからデータを取得
    return { success: false, error: '未実装（後回し）' };
  },

  /** @param {Array} data @returns {Object} */
  saveReportData(data) {
    // TODO: 後回し - レポートデータを保存
    return { success: false, error: '未実装（後回し）' };
  },

  /** @returns {Object} SM設定データ */
  getSMSettings() {
    // TODO: 後回し - SM設定シートからデータを取得
    return { success: false, error: '未実装（後回し）' };
  },

  /** @param {Array} data @returns {Object} */
  saveSMSettings(data) {
    // TODO: 後回し - SM設定データを保存
    return { success: false, error: '未実装（後回し）' };
  },
};

// --- 後方互換ラッパー ---
function deleteStudentRows(deleteStudentIds) { return AdminService.deleteStudentRows(deleteStudentIds); }
function getStudentRosterForAdmin() { return AdminService.getStudentRosterForAdmin(); }
function getStudentList() { return AdminService.getStudentList(); }
function replaceAllStudents(studentArray) { return AdminService.replaceAllStudents(studentArray); }
function importNewStudents(studentArray, autoMark, finalRegistration) { return AdminService.importNewStudents(studentArray, autoMark, finalRegistration); }
function importTestResults(importData) { return AdminService.importTestResults(importData); }
function replaceStudentIdInCourseSettings(replacements) { return AdminService.replaceStudentIdInCourseSettings(replacements); }
function executeYearTransition(options) { return AdminService.executeYearTransition(options); }
