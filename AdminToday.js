/* =====================================================
 * CAST FLOW
 * AdminToday.gs
 *
 * 管理者用：本日の出勤確認
 * ===================================================== */


/* =====================================================
 * 01. 本日の出勤情報取得
 * ===================================================== */

/**
 * 管理者用に本日の出勤予定・確認状態・
 * 同伴/来店予定をまとめて返す。
 *
 * @param {string} idToken
 * @return {Object}
 */
function getAdminTodayAttendance_(
  idToken
) {

  requireAdmin_(
    idToken
  );


  const today =
    new Date();


  const dateString =
    Utilities.formatDate(
      today,
      CONFIG.TIMEZONE,
      "yyyy-MM-dd"
    );


  const displayDate =
    Utilities.formatDate(
      today,
      CONFIG.TIMEZONE,
      "yyyy年M月d日"
    ) +
    "（" +
    getJapaneseWeekday_(
      today
    ) +
    "）";


  const shiftSheet =
    getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
    );


  const confirmationMap =
    getAdminTodayConfirmationMap_(
      dateString
    );


  const lastRow =
    shiftSheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return {

      success:
        true,

      date:
        dateString,

      displayDate:
        displayDate,

      items:
        []

    };

  }


  const values =
    shiftSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        13
      )
      .getValues();


  const items = [];


  values.forEach(
    row => {

      const rowDate =
        formatDateValue_(
          row[2]
        );


      const shiftStatus =
        String(
          row[6] || ""
        ).trim();


      if (
        rowDate !==
          dateString ||
        shiftStatus !==
          "出勤"
      ) {

        return;

      }


      const castId =
        String(
          row[4] || ""
        ).trim();


      const castName =
        String(
          row[5] || ""
        ).trim();


      if (
        !castId
      ) {

        return;

      }


      const confirmation =
        confirmationMap[
          castId
        ] || null;


      const visits =
        getCustomerVisitsForCastDate_(
          castId,
          dateString
        );


      items.push({

        shiftId:
          String(
            row[0] || ""
          ).trim(),

        date:
          dateString,

        castId:
          castId,

        castName:
          castName,

        startTime:
          formatTimeValue_(
            row[7]
          ),

        endTime:
          formatTimeValue_(
            row[8]
          ),

        transport:
          String(
            row[9] || ""
          ).trim(),

        note:
          String(
            row[10] || ""
          ).trim(),

        attendanceStatus:
          confirmation
            ? confirmation.status
            : "未送信",

        confirmedAt:
          confirmation
            ? formatAdminTodayDateTime_(
                confirmation.confirmedAt
              )
            : "",

        confirmationUpdatedAt:
          confirmation
            ? formatAdminTodayDateTime_(
                confirmation.updatedAt
              )
            : "",

        visits:
          visits

      });

    }
  );


  items.sort(
    (
      a,
      b
    ) => {

      const timeDiff =
        adminTodayTimeSortValue_(
          a.startTime
        ) -
        adminTodayTimeSortValue_(
          b.startTime
        );


      if (
        timeDiff !== 0
      ) {

        return timeDiff;

      }


      return String(
        a.castName || ""
      ).localeCompare(
        String(
          b.castName || ""
        ),
        "ja"
      );

    }
  );


  return {

    success:
      true,

    date:
      dateString,

    displayDate:
      displayDate,

    items:
      items

  };

}


/* =====================================================
 * 02. 出勤確認状態取得
 * ===================================================== */

function getAdminTodayConfirmationMap_(
  dateString
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.ATTENDANCE_CONFIRMATIONS
    );


  const lastRow =
    sheet.getLastRow();


  const result = {};


  if (
    lastRow < 2
  ) {

    return result;

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        9
      )
      .getValues();


  values.forEach(
    row => {

      const rowDate =
        formatDateValue_(
          row[1]
        );


      if (
        rowDate !==
        dateString
      ) {

        return;

      }


      const castId =
        String(
          row[2] || ""
        ).trim();


      if (
        !castId
      ) {

        return;

      }


      const updatedAt =
        normalizeAdminTodayDate_(
          row[7]
        );


      const existing =
        result[
          castId
        ];


      if (
        existing &&
        existing.updatedAtDate &&
        updatedAt &&
        existing.updatedAtDate.getTime() >
          updatedAt.getTime()
      ) {

        return;

      }


      result[
        castId
      ] = {

        status:
          String(
            row[5] ||
            "確認待ち"
          ).trim(),

        confirmedAt:
          row[6] || null,

        updatedAt:
          row[7] || null,

        updatedAtDate:
          updatedAt,

        note:
          String(
            row[8] || ""
          ).trim()

      };

    }
  );


  return result;

}


/* =====================================================
 * 03. 時刻ソート
 * ===================================================== */

function adminTodayTimeSortValue_(
  value
) {

  const text =
    String(
      value || ""
    ).trim();


  const match =
    text.match(
      /^(\d{1,2}):(\d{2})$/
    );


  if (
    !match
  ) {

    return 99999;

  }


  let hour =
    Number(
      match[1]
    );


  const minute =
    Number(
      match[2]
    );


  /*
   * 夜営業なので0〜11時は
   * 24時以降として扱う。
   */
  if (
    hour < 12
  ) {

    hour += 24;

  }


  return (
    hour * 60 +
    minute
  );

}


/* =====================================================
 * 04. 日時表示
 * ===================================================== */

function normalizeAdminTodayDate_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return null;

  }


  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]"
  ) {

    return isNaN(
      value.getTime()
    )
      ? null
      : value;

  }


  const date =
    new Date(
      value
    );


  return isNaN(
    date.getTime()
  )
    ? null
    : date;

}


function formatAdminTodayDateTime_(
  value
) {

  const date =
    normalizeAdminTodayDate_(
      value
    );


  if (
    !date
  ) {

    return "";

  }


  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE ||
      "Asia/Tokyo",
    "yyyy/MM/dd HH:mm"
  );

}
