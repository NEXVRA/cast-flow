/* =====================================================
 * CAST FLOW
 * AdminShift.gs
 *
 * 管理者用：月間シフト提出状況
 * ===================================================== */


/* =====================================================
 * 01. 月間提出状況取得
 * ===================================================== */

/**
 * 管理者用に指定月の提出状況を返す。
 *
 * @param {string} idToken
 * @param {number} year
 * @param {number} month 1〜12
 * @return {Object}
 */
function getMonthlySubmissionStatus_(
  idToken,
  year,
  month
) {

  requireAdmin_(
    idToken
  );

  validateYearMonth_(
    year,
    month
  );

  const targetMonth =
    createTargetMonth_(
      Number(year),
      Number(month)
    );

  const casts =
    getAdminMonthlySubmissionCasts_();

  const submissionMap =
    getAdminMonthlySubmissionMap_(
      targetMonth
    );

  const shiftSummaryMap =
    getAdminMonthlyShiftSummaryMap_(
      targetMonth
    );

  const items =
    casts.map(
      cast => {

        const submission =
          submissionMap[
            cast.castId
          ] || null;

        const summary =
          shiftSummaryMap[
            cast.castId
          ] || {
            workingDays: 0,
            undecidedDays: 0,
            holidayDays: 0,
            updatedAt: null
          };

        return {

          castId:
            cast.castId,

          castName:
            cast.castName,

          submitted:
            Boolean(
              submission &&
              submission.submitted
            ),

          workingDays:
            summary.workingDays,

          undecidedDays:
            summary.undecidedDays,

          holidayDays:
            summary.holidayDays,

          submittedAt:
            submission
              ? formatAdminMonthlyDateTime_(
                  submission.submittedAt
                )
              : "",

          updatedAt:
            formatAdminMonthlyDateTime_(
              summary.updatedAt
            )

        };

      }
    );

  return {

    success:
      true,

    targetMonth:
      targetMonth,

    year:
      Number(year),

    month:
      Number(month),

    items:
      items

  };

}


/* =====================================================
 * 02. 対象CAST取得
 * ===================================================== */

function getAdminMonthlySubmissionCasts_() {

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

    return [];

  }

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
        value =>
          String(
            value || ""
          ).trim()
      );

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

  const activeStatusIndex =
    findCastHeaderIndex_(
      headers,
      [
        "在籍状況"
      ]
    );

  if (
    castIdIndex < 0 ||
    castNameIndex < 0
  ) {

    throw new Error(
      "CASTシートに必要な列がありません：キャストID / キャスト名"
    );

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();

  const result = [];

  values.forEach(
    row => {

      const castId =
        String(
          row[
            castIdIndex
          ] || ""
        ).trim();

      const castName =
        String(
          row[
            castNameIndex
          ] || ""
        ).trim();

      if (
        !castId ||
        !castName
      ) {

        return;

      }

      if (
        activeStatusIndex >= 0
      ) {

        const activeStatus =
          String(
            row[
              activeStatusIndex
            ] || ""
          ).trim();

        if (
          activeStatus &&
          activeStatus !== "在籍"
        ) {

          return;

        }

      }

      result.push({

        castId:
          castId,

        castName:
          castName

      });

    }
  );

  return result;

}


/* =====================================================
 * 03. WEEKLYから提出状態取得
 * ===================================================== */

/**
 * 現行WEEKLYの実データ構成
 *
 * A 対象月
 * B 開始日
 * C 終了日
 * D CAST_ID
 * E 状態
 * F 入力日数
 * G 提出日時
 *
 * 現在のWEEKLY見出しは旧用途の表記が残っているため、
 * submitMonthlyShiftByLine / upsertMonthlySubmission_ と同じ
 * 7列構成を基準に読む。
 */
function getAdminMonthlySubmissionMap_(
  targetMonth
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.WEEKLY
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
        7
      )
      .getValues();

  const normalizedTargetMonth =
    normalizeTargetMonth_(
      targetMonth
    );

  values.forEach(
    row => {

      const rowTargetMonth =
        normalizeTargetMonth_(
          row[0]
        );

      if (
        rowTargetMonth !==
        normalizedTargetMonth
      ) {

        return;

      }

      const castId =
        String(
          row[3] || ""
        ).trim();

      if (
        !castId
      ) {

        return;

      }

      const status =
        String(
          row[4] || ""
        ).trim();

      result[
        castId
      ] = {

        submitted:
          status ===
          "提出済み",

        entryCount:
          Number(
            row[5] || 0
          ),

        submittedAt:
          row[6] || null

      };

    }
  );

  return result;

}


/* =====================================================
 * 04. SHIFT_CURRENT集計
 * ===================================================== */

function getAdminMonthlyShiftSummaryMap_(
  targetMonth
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.SHIFT_CURRENT
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
        13
      )
      .getValues();

  const normalizedTargetMonth =
    normalizeTargetMonth_(
      targetMonth
    );

  values.forEach(
    row => {

      const rowTargetMonth =
        normalizeTargetMonth_(
          row[1]
        );

      const workDateTargetMonth =
        getTargetMonthFromWorkDate_(
          row[2]
        );

      if (
        rowTargetMonth !==
          normalizedTargetMonth &&
        workDateTargetMonth !==
          normalizedTargetMonth
      ) {

        return;

      }

      const castId =
        String(
          row[4] || ""
        ).trim();

      if (
        !castId
      ) {

        return;

      }

      if (
        !result[
          castId
        ]
      ) {

        result[
          castId
        ] = {

          workingDays:
            0,

          undecidedDays:
            0,

          holidayDays:
            0,

          updatedAt:
            null

        };

      }

      const status =
        convertStatusToClient_(
          row[6]
        );

      if (
        status === "working"
      ) {

        result[
          castId
        ].workingDays++;

      } else if (
        status === "holiday"
      ) {

        result[
          castId
        ].holidayDays++;

      } else {

        result[
          castId
        ].undecidedDays++;

      }

      const updatedAt =
        normalizeAdminMonthlyDate_(
          row[11]
        );

      if (
        updatedAt &&
        (
          !result[
            castId
          ].updatedAt ||
          updatedAt.getTime() >
            result[
              castId
            ].updatedAt.getTime()
        )
      ) {

        result[
          castId
        ].updatedAt =
          updatedAt;

      }

    }
  );

  return result;

}


/* =====================================================
 * 05. 日時処理
 * ===================================================== */

function normalizeAdminMonthlyDate_(
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


function formatAdminMonthlyDateTime_(
  value
) {

  const date =
    normalizeAdminMonthlyDate_(
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
