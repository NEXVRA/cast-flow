/**
 * 指定月の本人シフトを取得
 *
 * @param {number} year
 * @param {number} month 1〜12
 * @return {Object}
 */
function getMonthlyShiftData(year, month) {
  try {
    validateYearMonth_(year, month);

    const sheet = getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
    );

    const user = getCurrentShiftUser_();
    const targetMonth = createTargetMonth_(year, month);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        entries: [],
        submitted: isMonthSubmitted_(
          user.castId,
          targetMonth
        )
      };
    }

    const values = sheet
      .getRange(2, 1, lastRow - 1, 13)
      .getValues();

    const entries = values
      .filter((row) => {
        return isTargetShiftRow_(
          row,
          user.castId,
          targetMonth
        );
      })
      .map((row) => {
        return {
          shiftId: String(row[0] || ""),
          targetMonth: normalizeTargetMonth_(row[1]),
          date: formatDateValue_(row[2]),
          weekday: String(row[3] || ""),
          castId: String(row[4] || "").trim(),
          castName: String(row[5] || ""),
          status: convertStatusToClient_(row[6]),
          startTime: formatTimeValue_(row[7]),
          endTime: formatTimeValue_(row[8]),
          transport: convertTransportToBoolean_(row[9]),
          note: String(row[10] || "")
        };
      })
      .filter((entry) => entry.date);

    return {
      success: true,
      entries: entries,
      submitted: isMonthSubmitted_(
        user.castId,
        targetMonth
      )
    };

  } catch (error) {
    writeErrorLog_(
      "getMonthlyShiftData",
      error,
      {
        year: year,
        month: month
      }
    );

    throw new Error(
      "シフトの読み込みに失敗しました：" +
      error.message
    );
  }
}


/**
 * 1日分のシフトを保存
 *
 * 同じキャスト・同じ日付が存在する場合は上書き。
 *
 * @param {Object} entry
 * @return {Object}
 */
function saveShiftEntry(entry) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    validateShiftEntry_(entry);

    const sheet = getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
    );

    const user = getCurrentShiftUser_();
    const workDate = parseDateString_(entry.date);

    const targetMonth = Utilities.formatDate(
      workDate,
      CONFIG.TIMEZONE,
      "yyyy-MM"
    );

    const weekday = getJapaneseWeekday_(workDate);
    const now = new Date();

    const currentData = findCurrentShiftRow_(
      sheet,
      user.castId,
      entry.date
    );

    const shiftId = currentData
      ? String(currentData.values[0])
      : createId("SHIFT");

    const status = convertStatusToSheet_(entry.status);

    const startTime =
      entry.status === "working"
        ? String(entry.startTime || "")
        : "";

    const endTime =
      entry.status === "working"
        ? String(entry.endTime || "")
        : "";

    const transport =
      entry.status === "working"
        ? (entry.transport ? "あり" : "なし")
        : "";

    const newRow = [
      shiftId,
      targetMonth,
      workDate,
      weekday,
      user.castId,
      user.castName,
      status,
      startTime,
      endTime,
      transport,
      String(entry.note || ""),
      now,
      user.castName
    ];

    if (currentData) {
      writeShiftChangeLogs_(
        shiftId,
        user,
        entry.date,
        currentData.values,
        newRow
      );

      sheet
        .getRange(currentData.rowNumber, 1, 1, 13)
        .setValues([newRow]);

    } else {
      sheet.appendRow(newRow);
    }

    return {
      success: true,
      entry: {
        shiftId: shiftId,
        targetMonth: targetMonth,
        date: entry.date,
        weekday: weekday,
        castId: user.castId,
        castName: user.castName,
        status: entry.status,
        startTime: startTime,
        endTime: endTime,
        transport: Boolean(entry.transport),
        note: String(entry.note || "")
      }
    };

  } catch (error) {
    writeErrorLog_(
      "saveShiftEntry",
      error,
      entry
    );

    throw new Error(
      "シフトの保存に失敗しました：" +
      error.message
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * 1日分のシフトを削除
 *
 * @param {string} dateString YYYY-MM-DD
 * @return {Object}
 */
function deleteShiftEntry(dateString) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    parseDateString_(dateString);

    const sheet = getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
    );

    const user = getCurrentShiftUser_();

    const currentData = findCurrentShiftRow_(
      sheet,
      user.castId,
      dateString
    );

    if (!currentData) {
      return {
        success: true,
        deleted: false
      };
    }

    writeDeletionChangeLog_(
      currentData.values,
      user,
      dateString
    );

    sheet.deleteRow(currentData.rowNumber);

    return {
      success: true,
      deleted: true
    };

  } catch (error) {
    writeErrorLog_(
      "deleteShiftEntry",
      error,
      {
        date: dateString
      }
    );

    throw new Error(
      "入力の削除に失敗しました：" +
      error.message
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * 1か月分を正式提出
 *
 * 未入力の日があっても提出可能。
 * 1日以上入力されていれば提出できる。
 *
 * @param {number} year
 * @param {number} month 1〜12
 * @return {Object}
 */
function submitMonthlyShift(year, month) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    validateYearMonth_(year, month);

    const currentSheet = getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
    );

    const inputSheet = getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_INPUT
    );

    const weeklySheet = getRequiredSheet_(
      CONFIG.SHEETS.WEEKLY
    );

    const user = getCurrentShiftUser_();
    const targetMonth = createTargetMonth_(year, month);

    const entries = getCurrentMonthRows_(
      currentSheet,
      user.castId,
      targetMonth
    );

    if (entries.length === 0) {
      throw new Error(
        "シフトを1日以上入力してください"
      );
    }

    const validEntries = entries.filter((item) => {
      return Boolean(
        formatDateValue_(item.values[2])
      );
    });

    if (validEntries.length === 0) {
      throw new Error(
        "有効な勤務日が見つかりません"
      );
    }

    const uniqueDates = new Set(
      validEntries.map((item) => {
        return formatDateValue_(
          item.values[2]
        );
      })
    );

    if (uniqueDates.size !== validEntries.length) {
      throw new Error(
        "同じ日付のシフトが重複しています"
      );
    }

    const submittedAt = new Date();
    const receptionId = createId("RECEPTION");

    const inputRows = validEntries.map((item) => {
      const row = item.values;

      return [
        receptionId,
        submittedAt,
        targetMonth,
        row[2],
        user.castId,
        user.castName,
        row[6],
        row[7],
        row[8],
        row[9],
        row[10],
        "WEBアプリ",
        ""
      ];
    });

    inputSheet
      .getRange(
        inputSheet.getLastRow() + 1,
        1,
        inputRows.length,
        13
      )
      .setValues(inputRows);

    upsertMonthlySubmission_(
      weeklySheet,
      user,
      targetMonth,
      year,
      month,
      validEntries.length,
      submittedAt
    );

    return {
      success: true,
      receptionId: receptionId,
      submittedAt: submittedAt.toISOString(),
      entryCount: validEntries.length
    };

  } catch (error) {
    writeErrorLog_(
      "submitMonthlyShift",
      error,
      {
        year: year,
        month: month
      }
    );

    throw new Error(
      "月間シフトの提出に失敗しました：" +
      error.message
    );

  } finally {
    lock.releaseLock();
  }
}

/**
 * 曜日指定で1か月分のシフトを一括保存する。
 *
 * 上書きOFF：
 * すでに入力済みの日は変更しない。
 *
 * 上書きON：
 * すでに入力済みの日も新しい内容へ変更する。
 *
 * @param {Object} data
 * @return {Object}
 */
function saveBulkShifts(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    validateBulkShiftData_(data);

    const sheet = getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
    );

    const user = getCurrentShiftUser_();

    const targetMonth =
      normalizeTargetMonth_(data.targetMonth);

    const monthParts =
      targetMonth.split("-").map(Number);

    const year = monthParts[0];
    const month = monthParts[1];

    validateYearMonth_(year, month);

    const selectedWeekdays =
      data.weekdays.map(Number);

    const overwrite =
      Boolean(data.overwrite);

    const status =
      String(data.status || "");

    const sheetStatus =
      convertStatusToSheet_(status);

    const startTime =
      status === "working"
        ? String(data.start || "")
        : "";

    const endTime =
      status === "working"
        ? String(data.end || "")
        : "";

    const transport =
      status === "working"
        ? (data.transport ? "あり" : "なし")
        : "";

    const note =
      String(data.note || "");

    const now = new Date();

    const lastRow =
      sheet.getLastRow();

    const existingValues =
      lastRow >= 2
        ? sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              13
            )
            .getValues()
        : [];

    /*
     * 本人の既存シフトを
     * 日付ごとに検索できる形へ変換
     */
    const existingByDate = {};

    existingValues.forEach(
      (row, index) => {
        const rowCastId =
          String(row[4] || "").trim();

        if (
          rowCastId !==
          String(user.castId).trim()
        ) {
          return;
        }

        const dateString =
          formatDateValue_(row[2]);

        if (!dateString) {
          return;
        }

        existingByDate[dateString] = {
          rowNumber: index + 2,
          values: row
        };
      }
    );

    const lastDay =
      new Date(year, month, 0).getDate();

    const rowsToAppend = [];
    const rowsToUpdate = [];

    let matchedCount = 0;
    let savedCount = 0;
    let skippedCount = 0;

    for (
      let day = 1;
      day <= lastDay;
      day++
    ) {
      const workDate =
        new Date(year, month - 1, day);

      const weekdayNumber =
        workDate.getDay();

      if (
        !selectedWeekdays.includes(
          weekdayNumber
        )
      ) {
        continue;
      }

      matchedCount++;

      const dateString =
        Utilities.formatDate(
          workDate,
          CONFIG.TIMEZONE,
          "yyyy-MM-dd"
        );

      const currentData =
        existingByDate[dateString];

      /*
       * 上書きOFFで既存入力がある場合は
       * 何も変更しない
       */
      if (
        currentData &&
        !overwrite
      ) {
        skippedCount++;
        continue;
      }

      const shiftId =
        currentData
          ? String(
              currentData.values[0] || ""
            ) || createId("SHIFT")
          : createId("SHIFT");

      const weekday =
        getJapaneseWeekday_(workDate);

      const newRow = [
        shiftId,
        targetMonth,
        workDate,
        weekday,
        user.castId,
        user.castName,
        sheetStatus,
        startTime,
        endTime,
        transport,
        note,
        now,
        user.castName
      ];

      if (currentData) {
        rowsToUpdate.push({
          rowNumber:
            currentData.rowNumber,

          oldRow:
            currentData.values,

          newRow:
            newRow,

          shiftId:
            shiftId,

          dateString:
            dateString
        });

      } else {
        rowsToAppend.push(newRow);
      }

      savedCount++;
    }

    /*
     * 既存行を更新
     */
    rowsToUpdate.forEach(
      (item) => {
        writeShiftChangeLogs_(
          item.shiftId,
          user,
          item.dateString,
          item.oldRow,
          item.newRow
        );

        sheet
          .getRange(
            item.rowNumber,
            1,
            1,
            13
          )
          .setValues([
            item.newRow
          ]);
      }
    );

    /*
     * 新規行をまとめて追加
     */
    if (rowsToAppend.length > 0) {
      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          rowsToAppend.length,
          13
        )
        .setValues(rowsToAppend);
    }

    return {
      success: true,
      targetMonth: targetMonth,
      matchedCount: matchedCount,
      savedCount: savedCount,
      skippedCount: skippedCount,
      overwrite: overwrite
    };

  } catch (error) {
    writeErrorLog_(
      "saveBulkShifts",
      error,
      data
    );

    throw new Error(
      "曜日一括入力に失敗しました：" +
      error.message
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * 曜日一括入力データの確認
 *
 * @param {Object} data
 */
function validateBulkShiftData_(data) {
  if (
    !data ||
    typeof data !== "object"
  ) {
    throw new Error(
      "一括入力データがありません"
    );
  }

  const targetMonth =
    normalizeTargetMonth_(
      data.targetMonth
    );

  if (
    !/^\d{4}-\d{2}$/.test(
      targetMonth
    )
  ) {
    throw new Error(
      "対象月が正しくありません"
    );
  }

  if (
    !Array.isArray(data.weekdays) ||
    data.weekdays.length === 0
  ) {
    throw new Error(
      "曜日を1つ以上選択してください"
    );
  }

  const weekdays =
    data.weekdays.map(Number);

  const invalidWeekday =
    weekdays.some((weekday) => {
      return (
        !Number.isInteger(weekday) ||
        weekday < 0 ||
        weekday > 6
      );
    });

  if (invalidWeekday) {
    throw new Error(
      "曜日データが正しくありません"
    );
  }

  const allowedStatuses = [
    "working",
    "holiday",
    "undecided"
  ];

  if (
    !allowedStatuses.includes(
      data.status
    )
  ) {
    throw new Error(
      "勤務区分が正しくありません"
    );
  }

  if (
    data.status === "working"
  ) {
    if (!data.start) {
      throw new Error(
        "出勤時間を選択してください"
      );
    }

    if (!data.end) {
      throw new Error(
        "退勤時間を選択してください"
      );
    }
  }

  if (
    data.note &&
    String(data.note).length > 100
  ) {
    throw new Error(
      "備考は100文字以内で入力してください"
    );
  }
}

/* =====================================================
 * 現在利用者・CAST情報
 * ===================================================== */


/**
 * 現在利用しているキャスト情報を取得する。
 *
 * 現在：
 * CONFIG.TEST_USER.CAST_IDを使用
 *
 * 将来：
 * LINEユーザーIDからCAST_IDを特定する。
 *
 * @return {Object}
 */
function getCurrentShiftUser_() {

  const castId =
    String(
      CONFIG.TEST_USER.CAST_ID || ""
    ).trim();


  if (!castId) {

    throw new Error(
      "利用者のキャストIDが設定されていません"
    );

  }


  /*
   * CASTシートから本人情報を取得
   */
  const cast =
    getCastById_(
      castId
    );


  /*
   * CASTに登録がある場合
   */
  if (cast) {

    return {

      castId:
        cast.castId,

      castName:
        cast.castName

    };

  }


  /*
   * CASTに見つからない場合は
   * 仮ユーザー情報へフォールバック
   */
  return {

    castId:
      castId,

    castName:
      String(
        CONFIG.TEST_USER.CAST_NAME || ""
      ).trim()

  };

}


/* =====================================================
 * CAST検索
 * ===================================================== */

/**
 * CAST_IDからCASTシートを検索する。
 *
 * @param {string} castId
 * @return {Object|null}
 */
function getCastById_(
  castId
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.CAST
    );


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {

    return null;

  }


  /*
   * ヘッダー取得
   */
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(
        (value) =>
          String(value || "").trim()
      );


  /*
   * 必要な列を探す
   */
  const castIdIndex =
    findCastHeaderIndex_(
      headers,
      [
        "キャストID",
        "CAST_ID"
      ]
    );


  const castNameIndex =
    findCastHeaderIndex_(
      headers,
      [
        "キャスト名",
        "CAST_NAME"
      ]
    );


  const startTimeIndex =
    findCastHeaderIndex_(
      headers,
      [
        "DEFAULT_START_TIME"
      ]
    );


  const endTimeIndex =
    findCastHeaderIndex_(
      headers,
      [
        "DEFAULT_END_TIME"
      ]
    );


  if (castIdIndex < 0) {

    throw new Error(
      "CASTシートにキャストID列がありません"
    );

  }


  /*
   * CASTデータ取得
   *
   * getValues()を使うことで、
   * 時刻セルがDate型でも取得できる。
   */
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();


  const normalizedCastId =
    String(
      castId || ""
    ).trim();


  for (
    let index = 0;
    index < values.length;
    index++
  ) {

    const row =
      values[index];


    const rowCastId =
      String(
        row[castIdIndex] || ""
      ).trim();


    if (
      rowCastId !==
      normalizedCastId
    ) {

      continue;

    }


    return {

      castId:
        rowCastId,


      castName:
        castNameIndex >= 0

          ? String(
              row[castNameIndex] || ""
            ).trim()

          : "",


      defaultStartTime:
        startTimeIndex >= 0

          ? formatTimeValue_(
              row[startTimeIndex]
            )

          : "",


      defaultEndTime:
        endTimeIndex >= 0

          ? formatTimeValue_(
              row[endTimeIndex]
            )

          : ""

    };

  }


  return null;

}


/* =====================================================
 * CASTヘッダー検索
 * ===================================================== */

/**
 * 複数候補からヘッダー列を探す。
 */
function findCastHeaderIndex_(
  headers,
  candidates
) {

  for (
    let index = 0;
    index < candidates.length;
    index++
  ) {

    const candidate =
      candidates[index];


    const foundIndex =
      headers.indexOf(
        candidate
      );


    if (
      foundIndex >= 0
    ) {

      return foundIndex;

    }

  }


  return -1;

}


/* =====================================================
 * 現在利用者・デフォルト勤務時間
 * ===================================================== */

/**
 * 現在のキャストの
 * デフォルト勤務時間を取得する。
 *
 * 優先順位
 *
 * 1. CASTシート
 * 2. SETTINGS
 *
 * @return {Object}
 */
function getCurrentCastDefaultTimes() {

  try {

    const castId =
      String(
        CONFIG.TEST_USER.CAST_ID || ""
      ).trim();


    const cast =
      getCastById_(
        castId
      );


    /*
     * SETTINGS側の共通デフォルト
     */
    const settings =
      getShiftSettings();


    const settingsStartTime =
      String(
        settings.defaultStartHour
      ).padStart(
        2,
        "0"
      ) +
      ":" +
      String(
        settings.defaultStartMinute
      ).padStart(
        2,
        "0"
      );


    const settingsEndTime =
      String(
        settings.defaultEndHour
      ).padStart(
        2,
        "0"
      ) +
      ":" +
      String(
        settings.defaultEndMinute
      ).padStart(
        2,
        "0"
      );


    /*
     * CASTに時間があれば優先。
     * 空欄ならSETTINGSを使用。
     */
    return {

      success:
        true,


      castId:
        cast
          ? cast.castId
          : castId,


      castName:
        cast
          ? cast.castName
          : String(
              CONFIG.TEST_USER.CAST_NAME || ""
            ).trim(),


      defaultStartTime:
        cast &&
        cast.defaultStartTime

          ? cast.defaultStartTime

          : settingsStartTime,


      defaultEndTime:
        cast &&
        cast.defaultEndTime

          ? cast.defaultEndTime

          : settingsEndTime

    };


  } catch (error) {

    writeErrorLog_(
      "getCurrentCastDefaultTimes",
      error,
      {}
    );


    throw new Error(
      "契約時間の取得に失敗しました：" +
      error.message
    );

  }

}


/* =====================================================
 * シート取得
 * ===================================================== */


function getRequiredSheet_(sheetName) {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      sheetName + "シートが見つかりません"
    );
  }

  return sheet;
}


/* =====================================================
 * 入力チェック
 * ===================================================== */


function validateYearMonth_(year, month) {
  const parsedYear = Number(year);
  const parsedMonth = Number(month);

  if (
    !Number.isInteger(parsedYear) ||
    parsedYear < 2000 ||
    parsedYear > 2100
  ) {
    throw new Error("年が正しくありません");
  }

  if (
    !Number.isInteger(parsedMonth) ||
    parsedMonth < 1 ||
    parsedMonth > 12
  ) {
    throw new Error("月が正しくありません");
  }
}


function validateShiftEntry_(entry) {
  if (!entry || typeof entry !== "object") {
    throw new Error(
      "シフトデータがありません"
    );
  }

  parseDateString_(entry.date);

  const allowedStatuses = [
    "working",
    "holiday",
    "undecided"
  ];

  if (!allowedStatuses.includes(entry.status)) {
    throw new Error(
      "勤務区分が正しくありません"
    );
  }

  if (entry.status === "working") {
    if (!entry.startTime) {
      throw new Error(
        "出勤時間を選択してください"
      );
    }

    if (!entry.endTime) {
      throw new Error(
        "退勤時間を選択してください"
      );
    }
  }

  if (
    entry.note &&
    String(entry.note).length > 100
  ) {
    throw new Error(
      "備考は100文字以内で入力してください"
    );
  }
}


/* =====================================================
 * 日付・対象月の変換
 * ===================================================== */


function parseDateString_(dateString) {
  if (
    typeof dateString !== "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(dateString)
  ) {
    throw new Error(
      "日付形式が正しくありません"
    );
  }

  const parts = dateString
    .split("-")
    .map(Number);

  const date = new Date(
    parts[0],
    parts[1] - 1,
    parts[2]
  );

  if (
    date.getFullYear() !== parts[0] ||
    date.getMonth() !== parts[1] - 1 ||
    date.getDate() !== parts[2]
  ) {
    throw new Error(
      "存在しない日付です"
    );
  }

  return date;
}


function createTargetMonth_(year, month) {
  return (
    String(year) +
    "-" +
    String(month).padStart(2, "0")
  );
}


/**
 * 対象月を必ずyyyy-MM形式へ変換する。
 *
 * 「2026-08」
 * 「2026/08」
 * 日付型
 * すべて対応。
 */
function normalizeTargetMonth_(value) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }

  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {
    if (isNaN(value.getTime())) {
      return "";
    }

    return Utilities.formatDate(
      value,
      CONFIG.TIMEZONE,
      "yyyy-MM"
    );
  }

  const text = String(value).trim();

  const monthOnlyMatch = text.match(
    /^(\d{4})[-\/](\d{1,2})$/
  );

  if (monthOnlyMatch) {
    return (
      monthOnlyMatch[1] +
      "-" +
      String(monthOnlyMatch[2]).padStart(2, "0")
    );
  }

  const fullDateMatch = text.match(
    /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/
  );

  if (fullDateMatch) {
    return (
      fullDateMatch[1] +
      "-" +
      String(fullDateMatch[2]).padStart(2, "0")
    );
  }

  const parsedDate = new Date(value);

  if (!isNaN(parsedDate.getTime())) {
    return Utilities.formatDate(
      parsedDate,
      CONFIG.TIMEZONE,
      "yyyy-MM"
    );
  }

  return text;
}


/**
 * 勤務日から対象月を取得する。
 */
function getTargetMonthFromWorkDate_(value) {
  const dateString = formatDateValue_(value);

  if (!dateString) {
    return "";
  }

  return dateString.substring(0, 7);
}


function getJapaneseWeekday_(date) {
  const weekdays = [
    "日",
    "月",
    "火",
    "水",
    "木",
    "金",
    "土"
  ];

  return weekdays[date.getDay()];
}


function formatDateValue_(value) {
  if (!value) {
    return "";
  }

  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {
    if (isNaN(value.getTime())) {
      return "";
    }

    return Utilities.formatDate(
      value,
      CONFIG.TIMEZONE,
      "yyyy-MM-dd"
    );
  }

  const text = String(value).trim();

  const exactMatch = text.match(
    /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/
  );

  if (exactMatch) {
    return (
      exactMatch[1] +
      "-" +
      String(exactMatch[2]).padStart(2, "0") +
      "-" +
      String(exactMatch[3]).padStart(2, "0")
    );
  }

  const parsed = new Date(value);

  if (isNaN(parsed.getTime())) {
    return "";
  }

  return Utilities.formatDate(
    parsed,
    CONFIG.TIMEZONE,
    "yyyy-MM-dd"
  );
}


function formatTimeValue_(value) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }

  if (
    String(value).trim().toUpperCase() ===
    "LAST"
  ) {
    return "LAST";
  }

  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {
    return Utilities.formatDate(
      value,
      CONFIG.TIMEZONE,
      "HH:mm"
    );
  }

  const text = String(value).trim();

  const timeMatch = text.match(
    /^(\d{1,2}):(\d{1,2})/
  );

  if (timeMatch) {
    return (
      String(timeMatch[1]).padStart(2, "0") +
      ":" +
      String(timeMatch[2]).padStart(2, "0")
    );
  }

  return text;
}


/* =====================================================
 * ステータス変換
 * ===================================================== */


function convertStatusToSheet_(status) {
  const map = {
    working: "出勤",
    holiday: "休み",
    undecided: "未定"
  };

  return map[status] || "";
}


function convertStatusToClient_(status) {
  const map = {
    出勤: "working",
    休み: "holiday",
    未定: "undecided"
  };

  return map[String(status || "").trim()] ||
    "undecided";
}


function convertTransportToBoolean_(value) {
  return (
    value === true ||
    String(value || "").trim() === "あり" ||
    String(value || "").trim().toUpperCase() ===
      "TRUE"
  );
}


/* =====================================================
 * シフト行判定
 * ===================================================== */


/**
 * 指定キャスト・指定月の行か判定する。
 *
 * B列の対象月だけでなく、
 * C列の勤務日からも対象月を判定する。
 */
function isTargetShiftRow_(
  row,
  castId,
  targetMonth
) {
  const normalizedCastId =
    String(castId || "").trim();

  const rowCastId =
    String(row[4] || "").trim();

  if (rowCastId !== normalizedCastId) {
    return false;
  }

  const normalizedTargetMonth =
    normalizeTargetMonth_(targetMonth);

  const rowTargetMonth =
    normalizeTargetMonth_(row[1]);

  const workDateTargetMonth =
    getTargetMonthFromWorkDate_(row[2]);

  return (
    rowTargetMonth === normalizedTargetMonth ||
    workDateTargetMonth === normalizedTargetMonth
  );
}


/**
 * 指定キャスト・指定日付の行を検索する。
 */
function findCurrentShiftRow_(
  sheet,
  castId,
  dateString
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 13)
    .getValues();

  const normalizedCastId =
    String(castId || "").trim();

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    const row = values[index];

    const rowCastId =
      String(row[4] || "").trim();

    const rowDate =
      formatDateValue_(row[2]);

    if (
      rowCastId === normalizedCastId &&
      rowDate === dateString
    ) {
      return {
        rowNumber: index + 2,
        values: row
      };
    }
  }

  return null;
}


/**
 * 指定キャスト・指定月の全シフトを取得する。
 */
function getCurrentMonthRows_(
  sheet,
  castId,
  targetMonth
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 13)
    .getValues();

  return values
    .map((row, index) => {
      return {
        rowNumber: index + 2,
        values: row
      };
    })
    .filter((item) => {
      return isTargetShiftRow_(
        item.values,
        castId,
        targetMonth
      );
    });
}


/* =====================================================
 * 変更履歴
 * ===================================================== */


function writeShiftChangeLogs_(
  shiftId,
  user,
  dateString,
  oldRow,
  newRow
) {
  const fields = [
    {
      label: "勤務区分",
      index: 6
    },
    {
      label: "出勤時間",
      index: 7
    },
    {
      label: "退勤時間",
      index: 8
    },
    {
      label: "送り",
      index: 9
    },
    {
      label: "備考",
      index: 10
    }
  ];

  const changeRows = [];

  fields.forEach((field) => {
    const oldValue =
      formatChangeLogValue_(
        oldRow[field.index],
        field.label
      );

    const newValue =
      formatChangeLogValue_(
        newRow[field.index],
        field.label
      );

    if (oldValue !== newValue) {
      changeRows.push([
        createId("CHANGE"),
        new Date(),
        shiftId,
        parseDateString_(dateString),
        user.castId,
        user.castName,
        field.label,
        oldValue,
        newValue,
        user.castName,
        "本人による変更"
      ]);
    }
  });

  if (changeRows.length === 0) {
    return;
  }

  const changeSheet = getRequiredSheet_(
    CONFIG.SHEETS.SHIFT_CHANGES
  );

  changeSheet
    .getRange(
      changeSheet.getLastRow() + 1,
      1,
      changeRows.length,
      11
    )
    .setValues(changeRows);
}


function formatChangeLogValue_(value, fieldLabel) {
  if (
    fieldLabel === "出勤時間" ||
    fieldLabel === "退勤時間"
  ) {
    return formatTimeValue_(value);
  }

  return String(value || "").trim();
}


function writeDeletionChangeLog_(
  oldRow,
  user,
  dateString
) {
  const changeSheet = getRequiredSheet_(
    CONFIG.SHEETS.SHIFT_CHANGES
  );

  changeSheet.appendRow([
    createId("CHANGE"),
    new Date(),
    String(oldRow[0] || ""),
    parseDateString_(dateString),
    user.castId,
    user.castName,
    "シフト",
    "登録済み",
    "削除",
    user.castName,
    "本人による入力削除"
  ]);
}


/* =====================================================
 * 月間提出状態
 * ===================================================== */


function isMonthSubmitted_(castId, targetMonth) {
  const sheet = getRequiredSheet_(
    CONFIG.SHEETS.WEEKLY
  );

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 7)
    .getValues();

  const normalizedCastId =
    String(castId || "").trim();

  const normalizedTargetMonth =
    normalizeTargetMonth_(targetMonth);

  return values.some((row) => {
    const rowTargetMonth =
      normalizeTargetMonth_(row[0]);

    const rowCastId =
      String(row[3] || "").trim();

    const rowStatus =
      String(row[4] || "").trim();

    return (
      rowTargetMonth === normalizedTargetMonth &&
      rowCastId === normalizedCastId &&
      rowStatus === "提出済み"
    );
  });
}


function upsertMonthlySubmission_(
  sheet,
  user,
  targetMonth,
  year,
  month,
  entryCount,
  submittedAt
) {
  const lastRow = sheet.getLastRow();

  const startDate = new Date(
    year,
    month - 1,
    1
  );

  const endDate = new Date(
    year,
    month,
    0
  );

  const normalizedTargetMonth =
    normalizeTargetMonth_(targetMonth);

  const rowValues = [
    normalizedTargetMonth,
    startDate,
    endDate,
    user.castId,
    "提出済み",
    entryCount,
    submittedAt
  ];

  if (lastRow < 2) {
    sheet.appendRow(rowValues);
    return;
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 7)
    .getValues();

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    const row = values[index];

    const rowTargetMonth =
      normalizeTargetMonth_(row[0]);

    const rowCastId =
      String(row[3] || "").trim();

    if (
      rowTargetMonth === normalizedTargetMonth &&
      rowCastId === String(user.castId).trim()
    ) {
      sheet
        .getRange(index + 2, 1, 1, 7)
        .setValues([rowValues]);

      return;
    }
  }

  sheet.appendRow(rowValues);
}


/* =====================================================
 * エラーログ
 * ===================================================== */


function writeErrorLog_(
  processName,
  error,
  targetData
) {
  try {
    const sheet = getRequiredSheet_(
      CONFIG.SHEETS.ERROR_LOG
    );

    sheet.appendRow([
      createId("ERROR"),
      new Date(),
      processName,
      error && error.stack
        ? error.stack
        : String(error),
      JSON.stringify(targetData || {}),
      "未対応"
    ]);

  } catch (logError) {
    console.error(
      "エラーログの記録に失敗しました",
      logError
    );
  }
}

/* =====================================================
 * LINE認証対応：シフト利用者取得
 * ===================================================== */

/**
 * LINE IDトークンから
 * シフトを操作する本人を取得する。
 *
 * @param {string} idToken
 * @return {Object}
 */
function getShiftUserByIdToken_(
  idToken
) {

  /*
   * LINE公式APIでIDトークン検証
   */
  const verified =
    verifyLineIdToken_(
      idToken
    );


  /*
   * LINEユーザーIDからCAST検索
   */
  const cast =
    getCastByLineUserId_(
      verified.lineUserId
    );


  if (!cast) {

    throw new Error(
      "CAST登録が完了していません"
    );

  }


  if (
    !cast.castId
  ) {

    throw new Error(
      "キャストIDを取得できません"
    );

  }


  if (
    !cast.castName
  ) {

    throw new Error(
      "キャスト名を取得できません"
    );

  }


  /*
   * 退店・無効ユーザー対策
   *
   * 在籍状況が空欄の場合は
   * 移行期間を考慮して許可する。
   */
  const activeStatus =
    String(
      cast.activeStatus || ""
    ).trim();


  if (
    activeStatus &&
    activeStatus !== "在籍"
  ) {

    throw new Error(
      "現在このアカウントではシフトを操作できません"
    );

  }


  return {

    castId:
      cast.castId,

    castName:
      cast.castName,

    lineUserId:
      verified.lineUserId,

    lineDisplayName:
      cast.lineDisplayName ||
      verified.displayName,

    isAdmin:
      Boolean(
        cast.isAdmin
      )

  };

}


/* =====================================================
 * LINE認証対応：本人確認テスト
 * ===================================================== */

/**
 * API接続テスト用。
 *
 * LIFFのIDトークンを使って
 * 誰として認識されるか確認する。
 *
 * @param {string} idToken
 * @return {Object}
 */
function getShiftUserAuthTest_(
  idToken
) {

  const user =
    getShiftUserByIdToken_(
      idToken
    );


  return {

    success:
      true,

    user: {

      castId:
        user.castId,

      castName:
        user.castName,

      lineDisplayName:
        user.lineDisplayName,

      isAdmin:
        user.isAdmin

    }

  };

}

/* =====================================================
 * LINE認証対応：月間シフト取得
 * ===================================================== */

/**
 * LINE認証済み本人の月間シフトを取得する。
 *
 * @param {string} idToken
 * @param {number} year
 * @param {number} month 1〜12
 * @return {Object}
 */
function getMonthlyShiftDataByLine(
  idToken,
  year,
  month
) {

  try {

    validateYearMonth_(
      year,
      month
    );


    /*
     * LINE IDトークンから本人を確定
     */
    const user =
      getShiftUserByIdToken_(
        idToken
      );


    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const targetMonth =
      createTargetMonth_(
        year,
        month
      );


    const lastRow =
      sheet.getLastRow();


    if (
      lastRow < 2
    ) {

      return {

        success:
          true,

        cast: {

          castId:
            user.castId,

          castName:
            user.castName

        },

        entries:
          [],

        submitted:
          isMonthSubmitted_(
            user.castId,
            targetMonth
          )

      };

    }


    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          13
        )
        .getValues();


    const entries =
      values

        .filter(
          (row) => {

            return isTargetShiftRow_(
              row,
              user.castId,
              targetMonth
            );

          }
        )

        .map(
          (row) => {

            return {

              shiftId:
                String(
                  row[0] || ""
                ),

              targetMonth:
                normalizeTargetMonth_(
                  row[1]
                ),

              date:
                formatDateValue_(
                  row[2]
                ),

              weekday:
                String(
                  row[3] || ""
                ),

              castId:
                String(
                  row[4] || ""
                ).trim(),

              castName:
                String(
                  row[5] || ""
                ),

              status:
                convertStatusToClient_(
                  row[6]
                ),

              startTime:
                formatTimeValue_(
                  row[7]
                ),

              endTime:
                formatTimeValue_(
                  row[8]
                ),

              transport:
                convertTransportToBoolean_(
                  row[9]
                ),

              note:
                String(
                  row[10] || ""
                )

            };

          }
        )

        .filter(
          (entry) =>
            entry.date
        );


    return {

      success:
        true,

      cast: {

        castId:
          user.castId,

        castName:
          user.castName

      },

      entries:
        entries,

      submitted:
        isMonthSubmitted_(
          user.castId,
          targetMonth
        )

    };


  } catch (error) {

    writeErrorLog_(
      "getMonthlyShiftDataByLine",
      error,
      {
        year:
          year,

        month:
          month
      }
    );


    throw new Error(
      "シフトの読み込みに失敗しました：" +
      error.message
    );

  }

}

/* =====================================================
 * LINE認証対応：1日シフト保存
 * ===================================================== */

function saveShiftEntryByLine(
  idToken,
  entry
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    validateShiftEntry_(
      entry
    );


    const user =
      getShiftUserByIdToken_(
        idToken
      );


    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const workDate =
      parseDateString_(
        entry.date
      );


    const targetMonth =
      Utilities.formatDate(
        workDate,
        CONFIG.TIMEZONE,
        "yyyy-MM"
      );


    const weekday =
      getJapaneseWeekday_(
        workDate
      );


    const now =
      new Date();


    const currentData =
      findCurrentShiftRow_(
        sheet,
        user.castId,
        entry.date
      );


    const shiftId =
      currentData

        ? String(
            currentData.values[0]
          )

        : createId(
            "SHIFT"
          );


    const status =
      convertStatusToSheet_(
        entry.status
      );


    const startTime =
      entry.status === "working"

        ? String(
            entry.startTime || ""
          )

        : "";


    const endTime =
      entry.status === "working"

        ? String(
            entry.endTime || ""
          )

        : "";


    const transport =
      entry.status === "working"

        ? (
            entry.transport
              ? "あり"
              : "なし"
          )

        : "";


    const newRow = [

      shiftId,

      targetMonth,

      workDate,

      weekday,

      user.castId,

      user.castName,

      status,

      startTime,

      endTime,

      transport,

      String(
        entry.note || ""
      ),

      now,

      user.castName

    ];


    if (currentData) {

      writeShiftChangeLogs_(
        shiftId,
        user,
        entry.date,
        currentData.values,
        newRow
      );


      sheet
        .getRange(
          currentData.rowNumber,
          1,
          1,
          13
        )
        .setValues([
          newRow
        ]);

    } else {

      sheet.appendRow(
        newRow
      );

    }


    return {

      success:
        true,

      entry: {

        shiftId:
          shiftId,

        targetMonth:
          targetMonth,

        date:
          entry.date,

        weekday:
          weekday,

        castId:
          user.castId,

        castName:
          user.castName,

        status:
          entry.status,

        startTime:
          startTime,

        endTime:
          endTime,

        transport:
          Boolean(
            entry.transport
          ),

        note:
          String(
            entry.note || ""
          )

      }

    };


  } catch (error) {

    writeErrorLog_(
      "saveShiftEntryByLine",
      error,
      entry || {}
    );


    throw new Error(
      "シフトの保存に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * LINE認証対応：1日シフト削除
 * ===================================================== */

function deleteShiftEntryByLine(
  idToken,
  dateString
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    parseDateString_(
      dateString
    );


    const user =
      getShiftUserByIdToken_(
        idToken
      );


    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const currentData =
      findCurrentShiftRow_(
        sheet,
        user.castId,
        dateString
      );


    if (!currentData) {

      return {

        success:
          true,

        deleted:
          false

      };

    }


    writeDeletionChangeLog_(
      currentData.values,
      user,
      dateString
    );


    sheet.deleteRow(
      currentData.rowNumber
    );


    return {

      success:
        true,

      deleted:
        true

    };


  } catch (error) {

    writeErrorLog_(
      "deleteShiftEntryByLine",
      error,
      {
        date:
          dateString
      }
    );


    throw new Error(
      "入力の削除に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * LINE認証対応：月間シフト提出
 * ===================================================== */

function submitMonthlyShiftByLine(
  idToken,
  year,
  month
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    validateYearMonth_(
      year,
      month
    );


    /*
     * LINE本人認証
     */
    const user =
      getShiftUserByIdToken_(
        idToken
      );


    const currentSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const inputSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_INPUT
      );


    const weeklySheet =
      getRequiredSheet_(
        CONFIG.SHEETS.WEEKLY
      );


    const targetMonth =
      createTargetMonth_(
        year,
        month
      );


    /*
     * すでに提出済みか確認
     *
     * 再提出の場合は
     * 管理者通知にも表示する
     */
    const wasSubmitted =
      isMonthSubmitted_(
        user.castId,
        targetMonth
      );


    const entries =
      getCurrentMonthRows_(
        currentSheet,
        user.castId,
        targetMonth
      );


    if (
      entries.length === 0
    ) {

      throw new Error(
        "シフトを1日以上入力してください"
      );

    }


    const validEntries =
      entries.filter(
        item => {

          return Boolean(
            formatDateValue_(
              item.values[2]
            )
          );

        }
      );


    if (
      validEntries.length === 0
    ) {

      throw new Error(
        "有効な勤務日が見つかりません"
      );

    }


    /*
     * 同じ日付の重複チェック
     */
    const uniqueDates =
      new Set(

        validEntries.map(
          item => {

            return formatDateValue_(
              item.values[2]
            );

          }
        )

      );


    if (
      uniqueDates.size !==
      validEntries.length
    ) {

      throw new Error(
        "同じ日付のシフトが重複しています"
      );

    }


    const submittedAt =
      new Date();


    const receptionId =
      createId(
        "RECEPTION"
      );


    /*
     * 提出履歴作成
     */
    const inputRows =
      validEntries.map(
        item => {

          const row =
            item.values;


          return [

            receptionId,

            submittedAt,

            targetMonth,

            row[2],

            user.castId,

            user.castName,

            row[6],

            row[7],

            row[8],

            row[9],

            row[10],

            "LINEアプリ",

            ""

          ];

        }
      );


    /*
     * SHIFT_INPUTへ提出内容を保存
     */
    inputSheet
      .getRange(
        inputSheet.getLastRow() + 1,
        1,
        inputRows.length,
        13
      )
      .setValues(
        inputRows
      );


    /*
     * WEEKLYへ提出済み状態を保存
     */
    upsertMonthlySubmission_(
      weeklySheet,
      user,
      targetMonth,
      year,
      month,
      validEntries.length,
      submittedAt
    );


    /*
     * -------------------------------------------------
     * 管理者へLINE通知
     *
     * 通知に失敗しても
     * シフト提出そのものは成功扱いにする。
     * -------------------------------------------------
     */
    try {

      sendMonthlyShiftSubmittedAdminNotification_(
        user,
        targetMonth,
        validEntries,
        submittedAt,
        wasSubmitted
      );

    } catch (
      notificationError
    ) {

      /*
       * LINE通知だけ失敗した場合は
       * ERROR_LOGへ残す
       */
      writeErrorLog_(
        "sendMonthlyShiftSubmittedAdminNotification_",
        notificationError,
        {

          castId:
            user.castId,

          castName:
            user.castName,

          targetMonth:
            targetMonth,

          receptionId:
            receptionId

        }
      );

    }


    return {

      success:
        true,

      receptionId:
        receptionId,

      submittedAt:
        submittedAt.toISOString(),

      entryCount:
        validEntries.length,

      resubmission:
        wasSubmitted

    };


  } catch (error) {

    writeErrorLog_(
      "submitMonthlyShiftByLine",
      error,
      {

        year:
          year,

        month:
          month

      }
    );


    throw new Error(
      "月間シフトの提出に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * LINE認証対応：曜日一括入力
 * ===================================================== */

function saveBulkShiftsByLine(
  idToken,
  data
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    validateBulkShiftData_(
      data
    );


    const user =
      getShiftUserByIdToken_(
        idToken
      );


    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const targetMonth =
      normalizeTargetMonth_(
        data.targetMonth
      );


    const monthParts =
      targetMonth
        .split("-")
        .map(Number);


    const year =
      monthParts[0];


    const month =
      monthParts[1];


    validateYearMonth_(
      year,
      month
    );


    const selectedWeekdays =
      data.weekdays.map(
        Number
      );


    const overwrite =
      Boolean(
        data.overwrite
      );


    const status =
      String(
        data.status || ""
      );


    const sheetStatus =
      convertStatusToSheet_(
        status
      );


    const startTime =
      status === "working"

        ? String(
            data.start || ""
          )

        : "";


    const endTime =
      status === "working"

        ? String(
            data.end || ""
          )

        : "";


    const transport =
      status === "working"

        ? (
            data.transport
              ? "あり"
              : "なし"
          )

        : "";


    const note =
      String(
        data.note || ""
      );


    const now =
      new Date();


    const lastRow =
      sheet.getLastRow();


    const existingValues =
      lastRow >= 2

        ? sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              13
            )
            .getValues()

        : [];


    const existingByDate = {};


    existingValues.forEach(
      (row, index) => {

        const rowCastId =
          String(
            row[4] || ""
          ).trim();


        if (
          rowCastId !==
          String(
            user.castId
          ).trim()
        ) {

          return;

        }


        const dateString =
          formatDateValue_(
            row[2]
          );


        if (!dateString) {

          return;

        }


        existingByDate[
          dateString
        ] = {

          rowNumber:
            index + 2,

          values:
            row

        };

      }
    );


    const lastDay =
      new Date(
        year,
        month,
        0
      ).getDate();


    const rowsToAppend = [];

    const rowsToUpdate = [];


    let matchedCount = 0;

    let savedCount = 0;

    let skippedCount = 0;


    for (
      let day = 1;
      day <= lastDay;
      day++
    ) {

      const workDate =
        new Date(
          year,
          month - 1,
          day
        );


      const weekdayNumber =
        workDate.getDay();


      if (
        !selectedWeekdays.includes(
          weekdayNumber
        )
      ) {

        continue;

      }


      matchedCount++;


      const dateString =
        Utilities.formatDate(
          workDate,
          CONFIG.TIMEZONE,
          "yyyy-MM-dd"
        );


      const currentData =
        existingByDate[
          dateString
        ];


      if (
        currentData &&
        !overwrite
      ) {

        skippedCount++;
        continue;

      }


      const shiftId =
        currentData

          ? (
              String(
                currentData.values[0] || ""
              ) ||
              createId(
                "SHIFT"
              )
            )

          : createId(
              "SHIFT"
            );


      const weekday =
        getJapaneseWeekday_(
          workDate
        );


      const newRow = [

        shiftId,

        targetMonth,

        workDate,

        weekday,

        user.castId,

        user.castName,

        sheetStatus,

        startTime,

        endTime,

        transport,

        note,

        now,

        user.castName

      ];


      if (currentData) {

        rowsToUpdate.push({

          rowNumber:
            currentData.rowNumber,

          oldRow:
            currentData.values,

          newRow:
            newRow,

          shiftId:
            shiftId,

          dateString:
            dateString

        });

      } else {

        rowsToAppend.push(
          newRow
        );

      }


      savedCount++;

    }


    rowsToUpdate.forEach(
      (item) => {

        writeShiftChangeLogs_(
          item.shiftId,
          user,
          item.dateString,
          item.oldRow,
          item.newRow
        );


        sheet
          .getRange(
            item.rowNumber,
            1,
            1,
            13
          )
          .setValues([
            item.newRow
          ]);

      }
    );


    if (
      rowsToAppend.length > 0
    ) {

      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          rowsToAppend.length,
          13
        )
        .setValues(
          rowsToAppend
        );

    }


    return {

      success:
        true,

      targetMonth:
        targetMonth,

      matchedCount:
        matchedCount,

      savedCount:
        savedCount,

      skippedCount:
        skippedCount,

      overwrite:
        overwrite

    };


  } catch (error) {

    writeErrorLog_(
      "saveBulkShiftsByLine",
      error,
      data || {}
    );


    throw new Error(
      "曜日一括入力に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * LINE認証対応：契約時間取得
 * ===================================================== */

function getCurrentCastDefaultTimesByLine(
  idToken
) {

  try {

    const user =
      getShiftUserByIdToken_(
        idToken
      );


    const cast =
      getCastById_(
        user.castId
      );


    const settings =
      getShiftSettings();


    const settingsStartTime =
      String(
        settings.defaultStartHour
      ).padStart(
        2,
        "0"
      ) +
      ":" +
      String(
        settings.defaultStartMinute
      ).padStart(
        2,
        "0"
      );


    const settingsEndTime =
      String(
        settings.defaultEndHour
      ).padStart(
        2,
        "0"
      ) +
      ":" +
      String(
        settings.defaultEndMinute
      ).padStart(
        2,
        "0"
      );


    return {

      success:
        true,

      castId:
        user.castId,

      castName:
        user.castName,

      defaultStartTime:
        cast &&
        cast.defaultStartTime

          ? cast.defaultStartTime

          : settingsStartTime,

      defaultEndTime:
        cast &&
        cast.defaultEndTime

          ? cast.defaultEndTime

          : settingsEndTime

    };


  } catch (error) {

    writeErrorLog_(
      "getCurrentCastDefaultTimesByLine",
      error,
      {}
    );


    throw new Error(
      "契約時間の取得に失敗しました：" +
      error.message
    );

  }

}

/* =====================================================
 * LINE認証対応：選択日一括入力
 * ===================================================== */

function saveMultiShiftsByLine(
  idToken,
  data
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    /*
     * 本人確認
     */
    const user =
      getShiftUserByIdToken_(
        idToken
      );


    /*
     * 基本チェック
     */
    if (
      !data ||
      !Array.isArray(
        data.dates
      ) ||
      data.dates.length === 0
    ) {

      throw new Error(
        "日付を1日以上選択してください"
      );

    }


    const status =
      String(
        data.status || ""
      );


    if (
      !isAllowedShiftStatus_(
        status
      )
    ) {

      throw new Error(
        "勤務区分が不正です"
      );

    }


    const dates =
      [
        ...new Set(
          data.dates.map(
            (date) =>
              String(
                date || ""
              ).trim()
          )
        )
      ];


    /*
     * 全日付を検証
     */
    dates.forEach(
      (dateString) => {

        parseDateString_(
          dateString
        );

      }
    );


    /*
     * 出勤の場合のみ時間を使用
     */
    const startTime =
      status === "working"

        ? String(
            data.startTime || ""
          ).trim()

        : "";


    const endTime =
      status === "working"

        ? String(
            data.endTime || ""
          ).trim()

        : "";


    if (
      status === "working" &&
      (
        !startTime ||
        !endTime
      )
    ) {

      throw new Error(
        "出勤時間と退勤時間を入力してください"
      );

    }


    const note =
      String(
        data.note || ""
      );


    const overwrite =
      Boolean(
        data.overwrite
      );


    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const lastRow =
      sheet.getLastRow();


    /*
     * 現在のSHIFT_CURRENTを取得
     */
    const existingValues =
      lastRow >= 2

        ? sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              13
            )
            .getValues()

        : [];


    /*
     * 本人分を
     * 日付 → 行データ
     * の形にする
     */
    const existingByDate = {};


    existingValues.forEach(
      (row, index) => {

        const rowCastId =
          String(
            row[4] || ""
          ).trim();


        if (
          rowCastId !==
          String(
            user.castId
          ).trim()
        ) {

          return;

        }


        const dateString =
          formatDateValue_(
            row[2]
          );


        if (!dateString) {
          return;
        }


        existingByDate[
          dateString
        ] = {

          rowNumber:
            index + 2,

          values:
            row

        };

      }
    );


    const rowsToAppend = [];

    const rowsToUpdate = [];

    const savedEntries = [];


    let savedCount = 0;
    let skippedCount = 0;


    const now =
      new Date();


    const sheetStatus =
      convertStatusToSheet_(
        status
      );


    /*
     * 選択日ごとに保存データ生成
     */
    dates.forEach(
      (dateString) => {

        const workDate =
          parseDateString_(
            dateString
          );


        const targetMonth =
          Utilities.formatDate(
            workDate,
            CONFIG.TIMEZONE,
            "yyyy-MM"
          );


        const weekday =
          getJapaneseWeekday_(
            workDate
          );


        const currentData =
          existingByDate[
            dateString
          ];


        /*
         * 入力済み +
         * 上書きOFFならスキップ
         */
        if (
          currentData &&
          !overwrite
        ) {

          skippedCount++;

          return;

        }


        const shiftId =
          currentData

            ? (
                String(
                  currentData.values[0] || ""
                ) ||
                createId(
                  "SHIFT"
                )
              )

            : createId(
                "SHIFT"
              );


        const newRow = [

          shiftId,

          targetMonth,

          workDate,

          weekday,

          user.castId,

          user.castName,

          sheetStatus,

          startTime,

          endTime,

          "",

          note,

          now,

          user.castName

        ];


        if (currentData) {

          rowsToUpdate.push({

            rowNumber:
              currentData.rowNumber,

            oldRow:
              currentData.values,

            newRow:
              newRow,

            shiftId:
              shiftId,

            dateString:
              dateString

          });

        } else {

          rowsToAppend.push(
            newRow
          );

        }


        /*
         * フロントへ返すデータ
         */
        savedEntries.push({

          shiftId:
            shiftId,

          targetMonth:
            targetMonth,

          date:
            dateString,

          weekday:
            weekday,

          castId:
            user.castId,

          castName:
            user.castName,

          status:
            status,

          startTime:
            startTime,

          endTime:
            endTime,

          note:
            note

        });


        savedCount++;

      }
    );


    /*
     * 既存行更新
     */
    rowsToUpdate.forEach(
      (item) => {

        writeShiftChangeLogs_(
          item.shiftId,
          user,
          item.dateString,
          item.oldRow,
          item.newRow
        );


        sheet
          .getRange(
            item.rowNumber,
            1,
            1,
            13
          )
          .setValues([
            item.newRow
          ]);

      }
    );


    /*
     * 新規行追加
     */
    if (
      rowsToAppend.length > 0
    ) {

      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          rowsToAppend.length,
          13
        )
        .setValues(
          rowsToAppend
        );

    }


    return {

      success:
        true,

      savedCount:
        savedCount,

      skippedCount:
        skippedCount,

      overwrite:
        overwrite,

      entries:
        savedEntries

    };


  } catch (error) {

    writeErrorLog_(
      "saveMultiShiftsByLine",
      error,
      data || {}
    );


    throw new Error(
      "選択日の一括登録に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * LINE認証対応：勤務区分チェック
 * ===================================================== */

function isAllowedShiftStatus_(
  status
) {

  return [

    "working",
    "holiday",
    "undecided"

  ].includes(
    String(
      status || ""
    )
  );

}


/* =====================================================
 * LINE認証対応：当日出勤OK
 * ===================================================== */

/**
 * LIFFのIDトークンから本人認証を行い、
 * 当日の出勤確認を「出勤OK」に更新する。
 *
 * 既存LIFF用。
 *
 * @param {string} idToken
 * @param {string} dateString YYYY-MM-DD
 * @return {Object}
 */
function confirmAttendanceOkByLine(
  idToken,
  dateString
) {

  const user =
    getShiftUserByIdToken_(
      idToken
    );

  return confirmAttendanceOk_(
    user,
    dateString,
    "LIFF"
  );

}


/**
 * LINE Webhookから受け取った
 * LINEユーザーIDで本人を特定し、
 * 当日の出勤確認を「出勤OK」に更新する。
 *
 * Cloudflare Worker経由専用。
 *
 * @param {string} lineUserId
 * @param {string} dateString YYYY-MM-DD
 * @return {Object}
 */
function confirmAttendanceOkByWebhook(
  lineUserId,
  dateString
) {

  const normalizedLineUserId =
    String(
      lineUserId || ""
    ).trim();


  if (!normalizedLineUserId) {

    throw new Error(
      "LINEユーザーIDがありません"
    );

  }


  /*
   * LINEユーザーIDからCASTを取得
   */
  const cast =
    getCastByLineUserId_(
      normalizedLineUserId
    );


  if (!cast) {

    throw new Error(
      "CAST登録が完了していません"
    );

  }


  if (!cast.castId) {

    throw new Error(
      "キャストIDを取得できません"
    );

  }


  if (!cast.castName) {

    throw new Error(
      "キャスト名を取得できません"
    );

  }


  /*
   * 退店・無効ユーザー対策
   */
  const activeStatus =
    String(
      cast.activeStatus || ""
    ).trim();


  if (
    activeStatus &&
    activeStatus !== "在籍"
  ) {

    throw new Error(
      "現在このアカウントでは出勤確認できません"
    );

  }


  const user = {

    castId:
      cast.castId,

    castName:
      cast.castName,

    lineUserId:
      normalizedLineUserId,

    lineDisplayName:
      cast.lineDisplayName || "",

    isAdmin:
      Boolean(
        cast.isAdmin
      )

  };


  return confirmAttendanceOk_(
    user,
    dateString,
    "WEBHOOK"
  );

}


/**
 * 出勤OKの共通更新処理。
 *
 * LIFF・Webhookのどちらからでも
 * 最終的にはこの関数を通す。
 *
 * @param {Object} user
 * @param {string} dateString YYYY-MM-DD
 * @param {string} source
 * @return {Object}
 */
function confirmAttendanceOk_(
  user,
  dateString,
  source
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    /*
     * 日付チェック
     */
    parseDateString_(
      dateString
    );


    if (
      !user ||
      !user.castId ||
      !user.castName
    ) {

      throw new Error(
        "キャスト情報が正しくありません"
      );

    }


    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.ATTENDANCE_CONFIRMATIONS
      );


    const confirmation =
      findAttendanceConfirmation_(
        sheet,
        dateString,
        user.castId
      );


    if (!confirmation) {

      throw new Error(
        "出勤確認データが見つかりません"
      );

    }


    const now =
      new Date();


    /*
     * F 状態
     * G 確認日時
     * H 更新日時
     */
    sheet
      .getRange(
        confirmation.rowNumber,
        6,
        1,
        3
      )
      .setValues([
        [
          "出勤OK",
          now,
          now
        ]
      ]);


    return {

      success:
        true,

      status:
        "出勤OK",

      date:
        dateString,

      castId:
        user.castId,

      castName:
        user.castName,

      source:
        String(
          source || ""
        )

    };


  } catch (error) {

    writeErrorLog_(
      "confirmAttendanceOk_",
      error,
      {

        date:
          dateString,

        castId:
          user &&
          user.castId
            ? user.castId
            : "",

        source:
          String(
            source || ""
          )

      }
    );


    throw new Error(
      "出勤確認に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * LINE認証対応：当日出勤変更あり
 * ===================================================== */

/**
 * LINE本人認証を行い、
 * 当日の出勤確認を「変更あり」に更新する。
 *
 * 更新後、
 * ・本人へ変更受付LINE
 * ・管理者へ即時LINE
 * を送信する。
 *
 * LINE通知に失敗しても、
 * 変更ありへの更新自体は成功扱いにする。
 *
 * @param {string} idToken
 * @param {string} dateString YYYY-MM-DD
 * @return {Object}
 */
function confirmAttendanceChangeByLine(
  idToken,
  dateString
) {

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(
      30000
    );


    /*
     * =========================================
     * LINE本人認証
     * =========================================
     */

    const user =
      getShiftUserByIdToken_(
        idToken
      );


    /*
     * =========================================
     * 日付チェック
     * =========================================
     */

    parseDateString_(
      dateString
    );


    /*
     * =========================================
     * 出勤確認データ取得
     * =========================================
     */

    const confirmationSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.ATTENDANCE_CONFIRMATIONS
      );


    const confirmation =
      findAttendanceConfirmation_(
        confirmationSheet,
        dateString,
        user.castId
      );


    if (!confirmation) {

      throw new Error(
        "出勤確認データが見つかりません"
      );

    }


    const now =
      new Date();


    /*
     * =========================================
     * ATTENDANCE_CONFIRMATIONS更新
     *
     * F 状態
     * G 確認日時
     * H 更新日時
     * =========================================
     */

    confirmationSheet
      .getRange(
        confirmation.rowNumber,
        6,
        1,
        3
      )
      .setValues([
        [
          "変更あり",
          now,
          now
        ]
      ]);


    /*
     * =========================================
     * 最新シフト取得
     * =========================================
     */

    const shiftSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.SHIFT_CURRENT
      );


    const shiftData =
      findCurrentShiftRow_(
        shiftSheet,
        user.castId,
        dateString
      );


    let shift = null;


    if (shiftData) {

      const row =
        shiftData.values;


      shift = {

        status:
          convertStatusToClient_(
            row[6]
          ),

        startTime:
          formatTimeValue_(
            row[7]
          ),

        endTime:
          formatTimeValue_(
            row[8]
          ),

        note:
          String(
            row[10] || ""
          ).trim()

      };

    }


    /*
     * =========================================
     * 最新の同伴・来店予定取得
     * =========================================
     */

    const customerVisits =
      getCustomerVisitsForCastDate_(
        user.castId,
        dateString
      );


    /*
     * =========================================
     * 本人・管理者へLINE通知
     *
     * 通知失敗で保存結果まで失敗にしない。
     * =========================================
     */

    let notificationSent =
      true;


    try {

      sendAttendanceChangeAcceptedNotifications_(
        user,
        dateString,
        shift,
        customerVisits
      );


    } catch (
      notificationError
    ) {

      notificationSent =
        false;


      writeErrorLog_(
        "sendAttendanceChangeAcceptedNotifications_",
        notificationError,
        {

          date:
            dateString,

          castId:
            user.castId,

          castName:
            user.castName

        }
      );

    }


    /*
     * =========================================
     * 結果
     * =========================================
     */

    return {

      success:
        true,

      status:
        "変更あり",

      date:
        dateString,

      castId:
        user.castId,

      castName:
        user.castName,

      notificationSent:
        notificationSent

    };


  } catch (error) {

    writeErrorLog_(
      "confirmAttendanceChangeByLine",
      error,
      {
        date:
          dateString
      }
    );


    throw new Error(
      "出勤変更確認に失敗しました：" +
      error.message
    );


  } finally {

    lock.releaseLock();

  }

}