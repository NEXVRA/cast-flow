/* =====================================================
 * ROYAL Shift System
 * Settings.gs
 *
 * Version : 1.1.0
 *
 * CONFIG.SHEETS.SETTINGS で指定された
 * 設定シートからシステム設定を取得する
 * ===================================================== */


/* =====================================================
 * 設定取得
 * ===================================================== */

/**
 * システム設定シートの内容を取得
 *
 * @returns {Object}
 */
function getShiftSettings() {

  const sheetName =
    CONFIG &&
    CONFIG.SHEETS &&
    CONFIG.SHEETS.SETTINGS
      ? CONFIG.SHEETS.SETTINGS
      : "SETTINGS";


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        sheetName
      );


  if (!sheet) {
    throw new Error(
      sheetName +
      "シートが見つかりません。"
    );
  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {
    throw new Error(
      sheetName +
      "シートに設定値がありません。"
    );
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  const settings = {};


  values.forEach(
    ([key, value]) => {

      const normalizedKey =
        String(
          key || ""
        ).trim();


      if (
        !normalizedKey
      ) {
        return;
      }


      settings[
        normalizedKey
      ] = value;

    }
  );


  return createSettingsObject(
    settings
  );

}


/* =====================================================
 * 設定オブジェクト生成
 * ===================================================== */

/**
 * システム設定シートの値を
 * JavaScriptで使いやすい形へ変換
 *
 * @param {Object} settings
 * @returns {Object}
 */
function createSettingsObject(
  settings
) {

  const startHourMin =
    Number(
      settings.START_HOUR_MIN
    );


  const startHourMax =
    Number(
      settings.START_HOUR_MAX
    );


  const endHours =
    String(
      settings.END_HOURS || ""
    )
      .split(",")
      .map(
        hour =>
          hour.trim()
      )
      .filter(Boolean);


  const minuteStep =
    Number(
      settings.MINUTE_STEP
    );


  return {

    /* ---------- 出勤 ---------- */

    startHourMin:
      startHourMin,

    startHourMax:
      startHourMax,

    startHourList:
      createStartHourList(
        startHourMin,
        startHourMax
      ),


    /* ---------- 退勤 ---------- */

    endHours:
      endHours,

    endHourList:
      createEndHourList(
        endHours
      ),


    /* ---------- 共通 ---------- */

    minuteStep:
      minuteStep,

    minuteList:
      createMinuteList(
        minuteStep
      ),

    enableLast:
      String(
        settings.ENABLE_LAST
      ).toUpperCase() === "TRUE",


    /* ---------- デフォルト ---------- */

    defaultStartHour:
      String(
        settings.DEFAULT_START_HOUR
      ).padStart(2, "0"),

    defaultStartMinute:
      String(
        settings.DEFAULT_START_MINUTE
      ).padStart(2, "0"),

    defaultEndHour:
      String(
        settings.DEFAULT_END_HOUR
      ).padStart(2, "0"),

    defaultEndMinute:
      String(
        settings.DEFAULT_END_MINUTE
      ).padStart(2, "0")

  };

}


/* =====================================================
 * 時間リスト生成
 * ===================================================== */

/**
 * 出勤時間の「時」リストを生成
 *
 * 例:
 * START_HOUR_MIN = 20
 * START_HOUR_MAX = 22
 *
 * => ["20", "21", "22"]
 *
 * @param {number} minHour
 * @param {number} maxHour
 * @returns {string[]}
 */
function createStartHourList(
  minHour,
  maxHour
) {

  const result = [];

  for (
    let hour = minHour;
    hour <= maxHour;
    hour++
  ) {

    result.push(
      String(hour).padStart(2, "0")
    );

  }

  return result;
}


/**
 * 退勤時間の「時」リストを整形
 *
 * 例:
 * ["23", "00", "01", "02", "03"]
 *
 * @param {string[]} endHours
 * @returns {string[]}
 */
function createEndHourList(
  endHours
) {

  return endHours
    .map(hour => {

      const number =
        Number(hour);

      if (
        !Number.isInteger(number) ||
        number < 0 ||
        number > 23
      ) {
        return null;
      }

      return String(number)
        .padStart(2, "0");

    })
    .filter(Boolean);

}


/**
 * 分リストを生成
 *
 * 例:
 * MINUTE_STEP = 15
 *
 * => ["00", "15", "30", "45"]
 *
 * @param {number} step
 * @returns {string[]}
 */
function createMinuteList(
  step
) {

  if (
    !Number.isInteger(step) ||
    step <= 0 ||
    step > 60
  ) {
    throw new Error(
      "MINUTE_STEPが正しくありません。"
    );
  }

  const result = [];

  for (
    let minute = 0;
    minute < 60;
    minute += step
  ) {

    result.push(
      String(minute)
        .padStart(2, "0")
    );

  }

  return result;
}
