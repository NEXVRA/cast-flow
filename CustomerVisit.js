/* =====================================================
 * CustomerVisit.gs
 * CAST FLOW
 * 同伴・来店予定管理
 * ===================================================== */


/* =====================================================
 * 01. 共通
 * ===================================================== */

/**
 * CUSTOMER_VISITS シートを取得
 */
function getCustomerVisitsSheet_() {

  return getRequiredSheet_(
    CONFIG.SHEETS.CUSTOMER_VISITS
  );

}


/**
 * CUSTOMER_VISITS のヘッダー情報を取得
 */
function getCustomerVisitHeaderMap_() {

  const sheet =
    getCustomerVisitsSheet_();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastColumn < 1
  ) {

    throw new Error(
      "CUSTOMER_VISITSシートにヘッダーがありません。"
    );

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


  const requiredHeaders = [

    "予定ID",
    "勤務日",
    "キャストID",
    "キャスト名",
    "種別",
    "顧客名",
    "予定時間",
    "備考",
    "更新日時",
    "更新者"

  ];


  requiredHeaders.forEach(
    header => {

      if (
        headers.indexOf(
          header
        ) < 0
      ) {

        throw new Error(
          "CUSTOMER_VISITSシートに「" +
          header +
          "」列がありません。"
        );

      }

    }
  );


  const map = {};


  headers.forEach(
    (
      header,
      index
    ) => {

      map[header] =
        index;

    }
  );


  return {
    headers:
      headers,
    map:
      map
  };

}


/**
 * 日付を yyyy-MM-dd に統一
 */
function normalizeCustomerVisitDate_(
  value
) {

  if (!value) {
    return "";
  }


  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]"
  ) {

    return Utilities.formatDate(
      value,
      CONFIG.TIMEZONE,
      "yyyy-MM-dd"
    );

  }


  const text =
    String(
      value
    ).trim();


  if (!text) {
    return "";
  }


  /*
   * yyyy/MM/dd
   */
  const slashMatch =
    text.match(
      /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/
    );


  if (
    slashMatch
  ) {

    return [
      slashMatch[1],
      String(
        slashMatch[2]
      ).padStart(
        2,
        "0"
      ),
      String(
        slashMatch[3]
      ).padStart(
        2,
        "0"
      )
    ].join(
      "-"
    );

  }


  /*
   * yyyy-MM-dd
   */
  const hyphenMatch =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    );


  if (
    hyphenMatch
  ) {

    return [
      hyphenMatch[1],
      String(
        hyphenMatch[2]
      ).padStart(
        2,
        "0"
      ),
      String(
        hyphenMatch[3]
      ).padStart(
        2,
        "0"
      )
    ].join(
      "-"
    );

  }


  throw new Error(
    "勤務日の形式が不正です: " +
    text
  );

}


/**
 * CUSTOMER_VISITS保存用のDateを生成
 */
function customerVisitDateToDate_(
  dateString
) {

  const normalized =
    normalizeCustomerVisitDate_(
      dateString
    );


  const parts =
    normalized
      .split(
        "-"
      );


  return new Date(
    Number(
      parts[0]
    ),
    Number(
      parts[1]
    ) - 1,
    Number(
      parts[2]
    )
  );

}


/**
 * 予定IDを生成
 */
function createCustomerVisitId_() {

  return (
    "VISIT_" +
    Utilities.getUuid()
  );

}


/**
 * 種別チェック
 */
function normalizeCustomerVisitType_(
  value
) {

  const type =
    String(
      value || ""
    ).trim();


  if (
    type !== "同伴" &&
    type !== "来店"
  ) {

    throw new Error(
      "種別は「同伴」または「来店」で指定してください。"
    );

  }


  return type;

}


/**
 * 予定時間チェック
 *
 * 空欄は許可
 * HH:mm 文字列 / スプレッドシート時刻(Date) / 数値時刻に対応
 */
function normalizeCustomerVisitTime_(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return "";

  }


  /*
   * Googleスプレッドシートの時刻セルは
   * Dateとして返ることがある
   */
  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]"
  ) {

    return Utilities.formatDate(
      value,
      CONFIG.TIMEZONE,
      "HH:mm"
    );

  }


  /*
   * 0.5 = 12:00 のような
   * 数値シリアル値にも対応
   */
  if (
    typeof value === "number" &&
    isFinite(value)
  ) {

    const totalMinutes =
      Math.round(
        value * 24 * 60
      );


    const hour =
      Math.floor(
        totalMinutes / 60
      ) % 24;


    const minute =
      totalMinutes % 60;


    return (
      String(
        hour
      ).padStart(
        2,
        "0"
      ) +
      ":" +
      String(
        minute
      ).padStart(
        2,
        "0"
      )
    );

  }


  const text =
    String(
      value
    ).trim();


  if (!text) {
    return "";
  }


  const match =
    text.match(
      /^(\d{1,2}):(\d{1,2})$/
    );


  if (
    !match
  ) {

    throw new Error(
      "予定時間は HH:mm 形式で入力してください。"
    );

  }


  const hour =
    Number(
      match[1]
    );


  const minute =
    Number(
      match[2]
    );


  if (
    hour < 0 ||
    hour > 29 ||
    minute < 0 ||
    minute > 59
  ) {

    throw new Error(
      "予定時間が不正です: " +
      text
    );

  }


  return (
    String(
      hour
    ).padStart(
      2,
      "0"
    ) +
    ":" +
    String(
      minute
    ).padStart(
      2,
      "0"
    )
  );

}


/* =====================================================
 * 02. 取得
 * ===================================================== */

/**
 * 指定キャスト・指定日の
 * 同伴 / 来店予定を取得
 */
function getCustomerVisitsForCastDate_(
  castId,
  dateString
) {

  const normalizedDate =
    normalizeCustomerVisitDate_(
      dateString
    );


  const sheet =
    getCustomerVisitsSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return [];

  }


  const headerInfo =
    getCustomerVisitHeaderMap_();


  const headers =
    headerInfo.headers;


  const map =
    headerInfo.map;


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const result = [];


  values.forEach(
    row => {

      const rowCastId =
        String(
          row[
            map["キャストID"]
          ] || ""
        ).trim();


      if (
        rowCastId !==
        String(
          castId || ""
        ).trim()
      ) {

        return;

      }


      const rowDate =
        normalizeCustomerVisitDate_(
          row[
            map["勤務日"]
          ]
        );


      if (
        rowDate !==
        normalizedDate
      ) {

        return;

      }


      result.push({

        visitId:
          String(
            row[
              map["予定ID"]
            ] || ""
          ).trim(),

        date:
          rowDate,

        castId:
          rowCastId,

        castName:
          String(
            row[
              map["キャスト名"]
            ] || ""
          ).trim(),

        type:
          String(
            row[
              map["種別"]
            ] || ""
          ).trim(),

        customerName:
          String(
            row[
              map["顧客名"]
            ] || ""
          ).trim(),

        time:
          normalizeCustomerVisitTime_(
            row[
              map["予定時間"]
            ]
          ),

        note:
          String(
            row[
              map["備考"]
            ] || ""
          ).trim()

      });

    }
  );


  /*
   * 同伴 → 来店
   * 同種別なら時間順
   */
  result.sort(
    (
      a,
      b
    ) => {

      const typeOrder = {
        "同伴":
          0,
        "来店":
          1
      };


      const typeDiff =
        (
          typeOrder[
            a.type
          ] ?? 99
        ) -
        (
          typeOrder[
            b.type
          ] ?? 99
        );


      if (
        typeDiff !== 0
      ) {

        return typeDiff;

      }


      return String(
        a.time || ""
      ).localeCompare(
        String(
          b.time || ""
        )
      );

    }
  );


  return result;

}


/**
 * LIFF本人用
 * 指定日の予定取得
 */
function getMyCustomerVisitsByLine(
  idToken,
  dateString
) {

  const user =
    getShiftUserByIdToken_(
      idToken
    );


  return {

    success:
      true,

    date:
      normalizeCustomerVisitDate_(
        dateString
      ),

    castId:
      user.castId,

    castName:
      user.castName,

    visits:
      getCustomerVisitsForCastDate_(
        user.castId,
        dateString
      )

  };

}


/* =====================================================
 * 03. 保存
 * ===================================================== */

/**
 * 指定キャスト・指定日の予定を
 * すべて置き換える
 *
 * visits = [
 *   {
 *     type: "同伴",
 *     customerName: "山田様",
 *     time: "20:30",
 *     note: ""
 *   }
 * ]
 */
function replaceCustomerVisitsForCastDate_(
  user,
  dateString,
  visits
) {

  const normalizedDate =
    normalizeCustomerVisitDate_(
      dateString
    );


  if (
    !Array.isArray(
      visits
    )
  ) {

    visits = [];

  }


  const normalizedVisits =
    visits.map(
      item => {

        const type =
          normalizeCustomerVisitType_(
            item.type
          );


        const customerName =
          String(
            item.customerName || ""
          ).trim();


        if (
          !customerName
        ) {

          throw new Error(
            type +
            "の顧客名を入力してください。"
          );

        }


        return {

          type:
            type,

          customerName:
            customerName,

          time:
            normalizeCustomerVisitTime_(
              item.time
            ),

          note:
            String(
              item.note || ""
            ).trim()

        };

      }
    );


  const sheet =
    getCustomerVisitsSheet_();


  const headerInfo =
    getCustomerVisitHeaderMap_();


  const headers =
    headerInfo.headers;


  const map =
    headerInfo.map;


  const lastRow =
    sheet.getLastRow();


  /*
   * 既存データ削除
   * 下から消す
   */
  if (
    lastRow >= 2
  ) {

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          headers.length
        )
        .getValues();


    const deleteRows = [];


    values.forEach(
      (
        row,
        index
      ) => {

        const rowCastId =
          String(
            row[
              map["キャストID"]
            ] || ""
          ).trim();


        if (
          rowCastId !==
          String(
            user.castId || ""
          ).trim()
        ) {

          return;

        }


        const rowDate =
          normalizeCustomerVisitDate_(
            row[
              map["勤務日"]
            ]
          );


        if (
          rowDate ===
          normalizedDate
        ) {

          deleteRows.push(
            index + 2
          );

        }

      }
    );


    deleteRows
      .sort(
        (
          a,
          b
        ) =>
          b - a
      )
      .forEach(
        rowNumber => {

          sheet.deleteRow(
            rowNumber
          );

        }
      );

  }


  /*
   * 新規データ保存
   */
  if (
    normalizedVisits.length === 0
  ) {

    return {

      success:
        true,

      date:
        normalizedDate,

      visits:
        []

    };

  }


  const now =
    new Date();


  const rows =
    normalizedVisits.map(
      item => {

        const row =
          new Array(
            headers.length
          ).fill(
            ""
          );


        row[
          map["予定ID"]
        ] =
          createCustomerVisitId_();


        row[
          map["勤務日"]
        ] =
          customerVisitDateToDate_(
            normalizedDate
          );


        row[
          map["キャストID"]
        ] =
          user.castId;


        row[
          map["キャスト名"]
        ] =
          user.castName;


        row[
          map["種別"]
        ] =
          item.type;


        row[
          map["顧客名"]
        ] =
          item.customerName;


        row[
          map["予定時間"]
        ] =
          item.time;


        row[
          map["備考"]
        ] =
          item.note;


        row[
          map["更新日時"]
        ] =
          now;


        row[
          map["更新者"]
        ] =
          user.castName;


        return row;

      }
    );


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      headers.length
    )
    .setValues(
      rows
    );


  return {

    success:
      true,

    date:
      normalizedDate,

    visits:
      getCustomerVisitsForCastDate_(
        user.castId,
        normalizedDate
      )

  };

}


/**
 * LIFF本人用
 * 同伴・来店予定保存
 */
function saveMyCustomerVisitsByLine(
  idToken,
  dateString,
  visits
) {

  const user =
    getShiftUserByIdToken_(
      idToken
    );


  return replaceCustomerVisitsForCastDate_(
    user,
    dateString,
    visits
  );

}


/* =====================================================
 * 04. 表示用
 * ===================================================== */

/**
 * LINE等で表示する予定一覧テキスト
 */
function buildCustomerVisitsDisplayText_(
  visits
) {

  if (
    !Array.isArray(
      visits
    ) ||
    visits.length === 0
  ) {

    return [
      "同伴：なし",
      "来店：なし"
    ].join(
      "\n"
    );

  }


  const companions =
    visits.filter(
      item =>
        item.type ===
        "同伴"
    );


  const visitsOnly =
    visits.filter(
      item =>
        item.type ===
        "来店"
    );


  const lines = [];


  if (
    companions.length === 0
  ) {

    lines.push(
      "同伴：なし"
    );

  } else {

    lines.push(
      "同伴："
    );


    companions.forEach(
      item => {

        lines.push(
          "・" +
          item.customerName +
          (
            item.time
              ? " " +
                item.time
              : ""
          )
        );

      }
    );

  }


  if (
    visitsOnly.length === 0
  ) {

    lines.push(
      "来店：なし"
    );

  } else {

    lines.push(
      "来店："
    );


    visitsOnly.forEach(
      item => {

        lines.push(
          "・" +
          item.customerName +
          (
            item.time
              ? " " +
                item.time
              : ""
          )
        );

      }
    );

  }


  return lines.join(
    "\n"
  );

}