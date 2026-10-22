/* =====================================================
 * CAST FLOW / Payroll.gs
 * 01. 給与用スプレッドシートへの接続確認
 *
 * データの変更・給与計算はまだ行わない
 * ===================================================== */

const CF_PAYROLL_CONNECTION = Object.freeze({
  spreadsheetId: "1jO2aeTTdvrTT6ZOpv-sZQZUBVopZ9yCBgAJ0i33UAU4"
});


/* =====================================================
 * 02. 給与用スプレッドシート取得
 * ===================================================== */

function cfPayrollOpenSpreadsheet_() {

  const payrollId =
    CF_PAYROLL_CONNECTION.spreadsheetId;

  const shiftSpreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!shiftSpreadsheet) {
    throw new Error(
      "シフト用スプレッドシートに紐づいたGASで実行してください。"
    );
  }

  if (shiftSpreadsheet.getId() === payrollId) {
    throw new Error(
      "シフト用と給与用の保存先が同じです。設定を確認してください。"
    );
  }

  return SpreadsheetApp.openById(payrollId);

}


/* =====================================================
 * 03. 接続テスト
 * GASエディタから手動実行する
 * ===================================================== */

function testPayrollConnection() {

  const spreadsheet =
    cfPayrollOpenSpreadsheet_();

  const result = {
    success: true,

    spreadsheetName:
      spreadsheet.getName(),

    sheetNames:
      spreadsheet.getSheets().map(function(sheet) {
        return sheet.getName();
      })
  };

  Logger.log(
    "給与用接続OK: " + JSON.stringify(result)
  );

  return result;

}


/* =====================================================
 * 04. 給与用タブの初期作成・構成更新
 *
 * ・新規タブは現在の最新構成で作成
 * ・旧構成のPAY_SETTINGS / PAY_RULESは
 *   既存データを残したまま「表示名」列を追加
 * ・不明な列構成の場合は変更せず停止
 * ===================================================== */

function setupPayrollSheets() {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const book = cfPayrollOpenSpreadsheet_();


    /* =================================================
     * シート定義
     * ================================================= */

    const definitions = [

      {
        name: "PAY_SETTINGS",

        headers: [
          "設定キー",
          "表示名",
          "設定値",
          "説明"
        ],

        legacyHeaders: [
          "設定キー",
          "設定値",
          "説明"
        ],

        migrate: function(sheet) {

          // 設定キーの右隣へ「表示名」を追加
          sheet.insertColumnAfter(1);

        }
      },


      {
        name: "PAY_CAST_MAP",

        headers: [
          "元システム",
          "元キャストID",
          "元キャスト名",
          "CAST_ID",
          "確認状態",
          "更新日時"
        ]
      },


      {
        name: "PAY_RULES",

        headers: [
          "ルールID",
          "表示名",
          "適用開始日",
          "適用終了日",
          "対象CAST_ID",
          "項目キー",
          "設定値",
          "単位",
          "確認状態",
          "備考"
        ],

        legacyHeaders: [
          "ルールID",
          "適用開始日",
          "適用終了日",
          "対象CAST_ID",
          "項目キー",
          "設定値",
          "単位",
          "確認状態",
          "備考"
        ],

        migrate: function(sheet) {

          // ルールIDの右隣へ「表示名」を追加
          sheet.insertColumnAfter(1);

        }
      },


      {
        name: "PAY_IMPORT_LOG",

        headers: [
          "取込ID",
          "ファイルID",
          "ファイル名",
          "内容ハッシュ",
          "データ種別",
          "対象開始日",
          "対象終了日",
          "処理状態",
          "読込行数",
          "反映行数",
          "処理日時",
          "エラー内容"
        ]
      }

    ];


    const initialized = [];
    const migrated = [];
    const preserved = [];


    /* =================================================
     * 各シート処理
     * ================================================= */

    definitions.forEach(function(def) {

      let sheet =
        book.getSheetByName(def.name);


      /* -----------------------------------------------
       * シートが存在しない場合
       * ----------------------------------------------- */

      if (!sheet) {

        sheet =
          book.insertSheet(def.name);

      }


      /* -----------------------------------------------
       * 完全に空の場合
       * ----------------------------------------------- */

      if (sheet.getLastRow() === 0) {

        cfPayrollEnsureColumns_(
          sheet,
          def.headers.length
        );

        cfPayrollWriteHeader_(
          sheet,
          def.headers
        );

        initialized.push(def.name);

        return;
      }


      /* -----------------------------------------------
       * 現行構成ならそのまま保持
       * ----------------------------------------------- */

      if (
        cfPayrollHeaderMatches_(
          sheet,
          def.headers
        )
      ) {

        cfPayrollFormatHeader_(
          sheet,
          def.headers.length
        );

        preserved.push(def.name);

        return;
      }


      /* -----------------------------------------------
       * 旧構成なら安全に移行
       * ----------------------------------------------- */

      if (
        def.legacyHeaders &&
        def.migrate &&
        cfPayrollHeaderMatches_(
          sheet,
          def.legacyHeaders
        )
      ) {

        def.migrate(sheet);


        cfPayrollEnsureColumns_(
          sheet,
          def.headers.length
        );


        cfPayrollWriteHeader_(
          sheet,
          def.headers
        );


        migrated.push(def.name);

        return;
      }


      /* -----------------------------------------------
       * 未知の構成
       * ----------------------------------------------- */

      throw new Error(
        def.name +
        " の列構成が想定と異なります。" +
        "データ保護のため変更せず停止しました。"
      );

    });


    SpreadsheetApp.flush();


    const result = {

      success: true,

      spreadsheetName:
        book.getName(),

      initialized:
        initialized,

      migrated:
        migrated,

      preserved:
        preserved

    };


    Logger.log(
      "給与用タブ設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 04-A. 必要列数を確保
 * ===================================================== */

function cfPayrollEnsureColumns_(
  sheet,
  requiredColumns
) {

  const currentColumns =
    sheet.getMaxColumns();


  if (currentColumns >= requiredColumns) {
    return;
  }


  sheet.insertColumnsAfter(
    currentColumns,
    requiredColumns - currentColumns
  );

}


/* =====================================================
 * 04-B. ヘッダー一致確認
 * ===================================================== */

function cfPayrollHeaderMatches_(
  sheet,
  headers
) {

  if (
    sheet.getLastColumn() !==
    headers.length
  ) {
    return false;
  }


  const range =
    sheet.getRange(
      1,
      1,
      1,
      headers.length
    );


  const values =
    range.getValues()[0];


  const formulas =
    range.getFormulas()[0];


  return headers.every(
    function(header, index) {

      return (
        values[index] === header &&
        formulas[index] === ""
      );

    }
  );

}


/* =====================================================
 * 04-C. ヘッダー書込
 * ===================================================== */

function cfPayrollWriteHeader_(
  sheet,
  headers
) {

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);


  cfPayrollFormatHeader_(
    sheet,
    headers.length
  );

}


/* =====================================================
 * 04-D. ヘッダー装飾
 * ===================================================== */

function cfPayrollFormatHeader_(
  sheet,
  columnCount
) {

  sheet
    .getRange(
      1,
      1,
      1,
      columnCount
    )
    .setBackground("#243247")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setWrap(true);


  sheet.setFrozenRows(1);

  sheet.setColumnWidths(
    1,
    columnCount,
    160
  );

  sheet.setRowHeight(
    1,
    40
  );

}


/* =====================================================
 * 05. 給与基本設定・初期ルール登録
 *
 * setupPayrollSheets() 実行後に1回実行する。
 *
 * ・既存の設定値は上書きしない
 * ・未登録キーのみ追加
 * ・英字キー + 日本語表示名で管理
 * ===================================================== */

function setupPayrollDefaults() {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const book = cfPayrollOpenSpreadsheet_();

    const settingsSheet =
      book.getSheetByName("PAY_SETTINGS");

    const rulesSheet =
      book.getSheetByName("PAY_RULES");


    if (!settingsSheet || !rulesSheet) {
      throw new Error(
        "PAY_SETTINGS または PAY_RULES がありません。" +
        "先に setupPayrollSheets() を実行してください。"
      );
    }


    /* =================================================
     * 最新ヘッダー構成を確認
     * ================================================= */

    const settingsHeaders = [
      "設定キー",
      "表示名",
      "設定値",
      "説明"
    ];

    const rulesHeaders = [
      "ルールID",
      "表示名",
      "適用開始日",
      "適用終了日",
      "対象CAST_ID",
      "項目キー",
      "設定値",
      "単位",
      "確認状態",
      "備考"
    ];


    if (
      !cfPayrollHeaderMatches_(
        settingsSheet,
        settingsHeaders
      )
    ) {
      throw new Error(
        "PAY_SETTINGS の列構成が異なります。"
      );
    }


    if (
      !cfPayrollHeaderMatches_(
        rulesSheet,
        rulesHeaders
      )
    ) {
      throw new Error(
        "PAY_RULES の列構成が異なります。"
      );
    }


    /* =================================================
     * PAY_SETTINGS
     *
     * ここは店舗全体の設定
     * ================================================= */

    const defaultSettings = [

      [
        "TIMEZONE",
        "タイムゾーン",
        "Asia/Tokyo",
        "給与計算で使用するタイムゾーン"
      ],

      [
        "WORK_UNIT_MINUTES",
        "勤務時間計算単位",
        15,
        "勤務時間を何分単位で計算するか"
      ],

      [
        "WORK_ROUND_MODE",
        "勤務時間端数処理",
        "CEIL",
        "CEIL = 15分単位で切り上げ"
      ],

      [
        "TAX_SERVICE_RATE",
        "税サ控除率",
        0.30,
        "売上計算時に使用する税サ率。30%なら0.30"
      ],

      [
        "PAYROLL_COMPARE_MODE",
        "給与計算方式",
        "MAX",
        "MAX = 時給制と売上制を比較して高い方を採用"
      ],

      [
        "SALES_MODE_BACK_ENABLED",
        "売上制採用時の通常バック",
        false,
        "false = 売上制採用時は通常バックを付けない"
      ],

      [
        "WEEK_START_DAY",
        "給与週の開始曜日",
        "",
        "例：SATURDAY。運用確定後に設定"
      ],

      [
        "WEEK_END_DAY",
        "給与週の締め曜日",
        "",
        "例：FRIDAY。運用確定後に設定"
      ],

      [
        "PAYDAY_DAY",
        "給与支給曜日",
        "",
        "支給曜日。運用確定後に設定"
      ]

    ];


    /* =================================================
     * PAY_RULES
     *
     * 対象CAST_IDが空欄
     * → 全キャスト共通ルール
     * ================================================= */

    const defaultRules = [

      /* ---------------------------------------------
       * 指名バック
       * --------------------------------------------- */

      [
        "R_BACK_MAIN_NOMINATION",
        "本指名バック",
        "",
        "",
        "",
        "BACK_MAIN_NOMINATION",
        1000,
        "円/件",
        "確認済",
        "全キャスト共通の基本値"
      ],

      [
        "R_BACK_VIP_NOMINATION",
        "VIP指名バック",
        "",
        "",
        "",
        "BACK_VIP_NOMINATION",
        1500,
        "円/件",
        "確認済",
        "全キャスト共通の基本値"
      ],

      [
        "R_BACK_INHOUSE_NOMINATION",
        "場内指名バック",
        "",
        "",
        "",
        "BACK_INHOUSE_NOMINATION",
        500,
        "円/件",
        "確認済",
        "全キャスト共通の基本値"
      ],


      /* ---------------------------------------------
       * ドリンク系
       * 金額は後で設定
       * --------------------------------------------- */

      [
        "R_BACK_DRINK",
        "ドリンクバック",
        "",
        "",
        "",
        "BACK_DRINK",
        "",
        "円/杯",
        "要設定",
        "金額確定後に設定値を入力"
      ],

      [
        "R_BACK_SHOT",
        "ショットバック",
        "",
        "",
        "",
        "BACK_SHOT",
        "",
        "円/杯",
        "要設定",
        "金額確定後に設定値を入力"
      ],


      /* ---------------------------------------------
       * 遅刻・欠勤
       * --------------------------------------------- */

      [
        "R_PENALTY_LATE_15",
        "遅刻罰金（15分）",
        "",
        "",
        "",
        "PENALTY_LATE_15",
        "",
        "円/15分",
        "要設定",
        "15分単位の遅刻罰金"
      ],

      [
        "R_PENALTY_SAME_DAY_ABSENCE",
        "当日欠勤罰金",
        "",
        "",
        "",
        "PENALTY_SAME_DAY_ABSENCE",
        "",
        "円",
        "要設定",
        "当日欠勤時の罰金"
      ],

      [
        "R_PENALTY_NO_SHOW",
        "無断欠勤罰金",
        "",
        "",
        "",
        "PENALTY_NO_SHOW",
        "",
        "円",
        "要設定",
        "無断欠勤時の罰金"
      ],


      /* ---------------------------------------------
       * その他バック
       * --------------------------------------------- */

      [
        "R_BACK_BLOG",
        "ブログバック",
        "",
        "",
        "",
        "BACK_BLOG",
        500,
        "円/件",
        "確認済",
        "ブログ1件あたり"
      ]

    ];


    /* =================================================
     * 既存設定キー取得
     * ================================================= */

    const existingSettings =
      cfPayrollGetExistingKeys_(
        settingsSheet,
        1
      );


    const existingRules =
      cfPayrollGetExistingKeys_(
        rulesSheet,
        1
      );


    /* =================================================
     * 未登録設定だけ追加
     * ================================================= */

    const newSettings =
      defaultSettings.filter(function(row) {

        return !existingSettings[row[0]];

      });


    if (newSettings.length > 0) {

      settingsSheet
        .getRange(
          settingsSheet.getLastRow() + 1,
          1,
          newSettings.length,
          4
        )
        .setValues(newSettings);

    }


    /* =================================================
     * 未登録ルールだけ追加
     * ================================================= */

    const newRules =
      defaultRules.filter(function(row) {

        return !existingRules[row[0]];

      });


    if (newRules.length > 0) {

      rulesSheet
        .getRange(
          rulesSheet.getLastRow() + 1,
          1,
          newRules.length,
          10
        )
        .setValues(newRules);

    }


    SpreadsheetApp.flush();


    const result = {

      success: true,

      settingsAdded:
        newSettings.length,

      rulesAdded:
        newRules.length,

      settingsTotal:
        settingsSheet.getLastRow() - 1,

      rulesTotal:
        rulesSheet.getLastRow() - 1

    };


    Logger.log(
      "給与初期設定登録OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 05-A. 既存キー一覧取得
 * ===================================================== */

function cfPayrollGetExistingKeys_(
  sheet,
  keyColumn
) {

  const result = {};

  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {
    return result;
  }


  const values =
    sheet
      .getRange(
        2,
        keyColumn,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  values.forEach(function(row) {

    const key =
      String(row[0] || "").trim();


    if (key) {
      result[key] = true;
    }

  });


  return result;

}


/* =====================================================
 * 06. シフト側キャスト → 給与側PAY_CAST_MAP同期
 *
 * ・CAST_MASTERの列位置には依存しない
 * ・CAST_ID / キャストID の両方に対応
 * ・CAST_NAME / キャスト名 の両方に対応
 *
 * 新規キャスト：
 *   元キャストID = シフト側CAST_ID
 *   CAST_ID     = 同じIDを初期値として登録
 *
 * 既存キャスト：
 *   名前は最新状態へ更新
 *   給与側CAST_IDは上書きしない
 *
 * 削除済み・退店済みキャスト：
 *   給与履歴保護のためPAY_CAST_MAPから自動削除しない
 * ===================================================== */

function syncPayrollCastMap() {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const shiftBook =
      SpreadsheetApp.getActiveSpreadsheet();

    const payrollBook =
      cfPayrollOpenSpreadsheet_();


    if (!shiftBook) {
      throw new Error(
        "シフト用スプレッドシートに紐づいたGASで実行してください。"
      );
    }


    /* =================================================
     * シフト側CAST_MASTER取得
     * ================================================= */

    const castSheetName =
      (
        typeof CONFIG !== "undefined" &&
        CONFIG.SHEETS &&
        CONFIG.SHEETS.CAST
      )
        ? CONFIG.SHEETS.CAST
        : "CAST_MASTER";


    const castSheet =
      shiftBook.getSheetByName(
        castSheetName
      );


    if (!castSheet) {
      throw new Error(
        "シフト側に " +
        castSheetName +
        " が見つかりません。"
      );
    }


    /* =================================================
     * 給与側PAY_CAST_MAP取得
     * ================================================= */

    const mapSheet =
      payrollBook.getSheetByName(
        "PAY_CAST_MAP"
      );


    if (!mapSheet) {
      throw new Error(
        "給与側にPAY_CAST_MAPがありません。" +
        "先にsetupPayrollSheets()を実行してください。"
      );
    }


    const expectedMapHeaders = [
      "元システム",
      "元キャストID",
      "元キャスト名",
      "CAST_ID",
      "確認状態",
      "更新日時"
    ];


    if (
      !cfPayrollHeaderMatches_(
        mapSheet,
        expectedMapHeaders
      )
    ) {
      throw new Error(
        "PAY_CAST_MAPの列構成が異なります。"
      );
    }


    /* =================================================
     * CAST_MASTERヘッダー取得
     * ================================================= */

    if (castSheet.getLastRow() < 1) {
      throw new Error(
        castSheetName +
        " に見出し行がありません。"
      );
    }


    const sourceLastColumn =
      castSheet.getLastColumn();


    const sourceHeaders =
      castSheet
        .getRange(
          1,
          1,
          1,
          sourceLastColumn
        )
        .getDisplayValues()[0]
        .map(function(value) {
          return String(value || "").trim();
        });


    /* =================================================
     * 必要列を見出し名から検索
     * ================================================= */

    const castIdColumn =
      cfPayrollFindHeaderIndex_(
        sourceHeaders,
        [
          "キャストID",
          "CAST_ID"
        ]
      );


    const castNameColumn =
      cfPayrollFindHeaderIndex_(
        sourceHeaders,
        [
          "キャスト名",
          "CAST_NAME"
        ]
      );


    if (castIdColumn === -1) {
      throw new Error(
        castSheetName +
        " に「キャストID」または「CAST_ID」がありません。"
      );
    }


    if (castNameColumn === -1) {
      throw new Error(
        castSheetName +
        " に「キャスト名」または「CAST_NAME」がありません。"
      );
    }


    /* =================================================
     * シフト側キャスト取得
     * ================================================= */

    const sourceLastRow =
      castSheet.getLastRow();


    const sourceRows =
      sourceLastRow <= 1
        ? []
        : castSheet
            .getRange(
              2,
              1,
              sourceLastRow - 1,
              sourceLastColumn
            )
            .getValues();


    const sourceCasts = [];

    const sourceIdCheck = {};


    sourceRows.forEach(function(row, index) {

      const castId =
        String(
          row[castIdColumn] || ""
        ).trim();


      const castName =
        String(
          row[castNameColumn] || ""
        ).trim();


      /* CAST_ID空欄行は無視 */
      if (!castId) {
        return;
      }


      /* CAST_MASTER内の重複ID防止 */
      if (sourceIdCheck[castId]) {

        throw new Error(
          castSheetName +
          " にCAST_IDの重複があります：" +
          castId +
          "（行 " +
          (index + 2) +
          "）"
        );

      }


      sourceIdCheck[castId] = true;


      sourceCasts.push({

        sourceCastId:
          castId,

        sourceCastName:
          castName

      });

    });


    /* =================================================
     * PAY_CAST_MAP既存データ取得
     * ================================================= */

    const mapLastRow =
      mapSheet.getLastRow();


    const mapRows =
      mapLastRow <= 1
        ? []
        : mapSheet
            .getRange(
              2,
              1,
              mapLastRow - 1,
              6
            )
            .getValues();


    const SOURCE_SYSTEM =
      "CAST_FLOW";


    const existingMap = {};


    mapRows.forEach(function(row, index) {

      const sourceSystem =
        String(row[0] || "").trim();


      const sourceCastId =
        String(row[1] || "").trim();


      if (!sourceSystem || !sourceCastId) {
        return;
      }


      const key =
        sourceSystem +
        "::" +
        sourceCastId;


      if (existingMap[key]) {

        throw new Error(
          "PAY_CAST_MAPに重複があります：" +
          sourceSystem +
          " / " +
          sourceCastId +
          "（行 " +
          (index + 2) +
          "）"
        );

      }


      existingMap[key] = {

        sheetRow:
          index + 2,

        sourceCastName:
          String(row[2] || "").trim(),

        payrollCastId:
          String(row[3] || "").trim(),

        status:
          String(row[4] || "").trim(),

        updatedAt:
          row[5]

      };

    });


    /* =================================================
     * 同期処理
     * ================================================= */

    const now =
      new Date();


    const newRows = [];

    const updateRows = [];

    let unchangedCount = 0;


    sourceCasts.forEach(function(cast) {

      const key =
        SOURCE_SYSTEM +
        "::" +
        cast.sourceCastId;


      const existing =
        existingMap[key];


      /* -----------------------------------------------
       * 新規キャスト
       * ----------------------------------------------- */

      if (!existing) {

        newRows.push([
          SOURCE_SYSTEM,
          cast.sourceCastId,
          cast.sourceCastName,

          // 初期値はシフト側IDと同じ
          cast.sourceCastId,

          "自動登録",
          now
        ]);

        return;

      }


      /* -----------------------------------------------
       * 既存キャスト
       *
       * PAY_CAST_MAP側CAST_IDは絶対に上書きしない
       * ----------------------------------------------- */

      const newPayrollCastId =
        existing.payrollCastId ||
        cast.sourceCastId;


      const nameChanged =
        existing.sourceCastName !==
        cast.sourceCastName;


      const payrollIdWasEmpty =
        !existing.payrollCastId;


      if (
        !nameChanged &&
        !payrollIdWasEmpty
      ) {

        unchangedCount++;

        return;

      }


      updateRows.push({

        sheetRow:
          existing.sheetRow,

        values: [
          SOURCE_SYSTEM,
          cast.sourceCastId,
          cast.sourceCastName,
          newPayrollCastId,

          existing.status ||
          "自動登録",

          now
        ]

      });

    });


    /* =================================================
     * 既存行更新
     * ================================================= */

    updateRows.forEach(function(item) {

      mapSheet
        .getRange(
          item.sheetRow,
          1,
          1,
          6
        )
        .setValues([
          item.values
        ]);

    });


    /* =================================================
     * 新規行追加
     * ================================================= */

    if (newRows.length > 0) {

      mapSheet
        .getRange(
          mapSheet.getLastRow() + 1,
          1,
          newRows.length,
          6
        )
        .setValues(newRows);

    }


    /* =================================================
     * 更新日時表示
     * ================================================= */

    if (mapSheet.getLastRow() > 1) {

      mapSheet
        .getRange(
          2,
          6,
          mapSheet.getLastRow() - 1,
          1
        )
        .setNumberFormat(
          "yyyy/mm/dd hh:mm:ss"
        );

    }


    SpreadsheetApp.flush();


    const result = {

      success: true,

      sourceSheet:
        castSheetName,

      sourceCount:
        sourceCasts.length,

      added:
        newRows.length,

      updated:
        updateRows.length,

      unchanged:
        unchangedCount

    };


    Logger.log(
      "給与キャスト同期OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 06-A. 複数候補からヘッダー位置を取得
 *
 * 戻り値は配列INDEX
 * A列 = 0
 * B列 = 1
 * ===================================================== */

function cfPayrollFindHeaderIndex_(
  headers,
  candidates
) {

  for (
    let i = 0;
    i < candidates.length;
    i++
  ) {

    const index =
      headers.indexOf(
        candidates[i]
      );


    if (index !== -1) {
      return index;
    }

  }


  return -1;

}


/* =====================================================
 * 07. 給与実績データ用タブ作成
 *
 * PAY_FACTS
 *   → Excel等から取り込んだ元実績を保存
 *
 * PAY_ADJUSTMENTS
 *   → 手動修正・追加分を保存
 *
 * 計算はまだ行わない
 * ===================================================== */

function setupPayrollDataSheets() {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const definitions = [

      /* =================================================
       * 取込実績
       * ================================================= */

      {
        name: "PAY_FACTS",

        headers: [
          "実績ID",
          "取込ID",
          "日付",
          "元システム",
          "元キャストID",
          "CAST_ID",
          "項目キー",
          "表示名",
          "値",
          "単位",
          "元ファイルID",
          "元ファイル名",
          "元行番号",
          "内容ハッシュ",
          "登録日時"
        ]
      },


      /* =================================================
       * 手動調整
       * ================================================= */

      {
        name: "PAY_ADJUSTMENTS",

        headers: [
          "調整ID",
          "対象日",
          "CAST_ID",
          "項目キー",
          "表示名",
          "調整値",
          "単位",
          "理由",
          "承認状態",
          "登録者",
          "登録日時",
          "更新日時"
        ]
      }

    ];


    const initialized = [];
    const preserved = [];


    definitions.forEach(function(def) {

      let sheet =
        book.getSheetByName(
          def.name
        );


      /* -----------------------------------------------
       * シートが無ければ作成
       * ----------------------------------------------- */

      if (!sheet) {

        sheet =
          book.insertSheet(
            def.name
          );

      }


      /* -----------------------------------------------
       * 空シート
       * ----------------------------------------------- */

      if (sheet.getLastRow() === 0) {

        cfPayrollEnsureColumns_(
          sheet,
          def.headers.length
        );


        cfPayrollWriteHeader_(
          sheet,
          def.headers
        );


        initialized.push(
          def.name
        );

        return;

      }


      /* -----------------------------------------------
       * 既存シート
       * ----------------------------------------------- */

      if (
        !cfPayrollHeaderMatches_(
          sheet,
          def.headers
        )
      ) {

        throw new Error(
          def.name +
          " の列構成が異なります。" +
          "データ保護のため変更せず停止しました。"
        );

      }


      cfPayrollFormatHeader_(
        sheet,
        def.headers.length
      );


      preserved.push(
        def.name
      );

    });


    /* =================================================
     * PAY_FACTS 表示設定
     * ================================================= */

    const factsSheet =
      book.getSheetByName(
        "PAY_FACTS"
      );


    if (factsSheet) {

      /* 日付 */
      factsSheet
        .getRange(
          "C:C"
        )
        .setNumberFormat(
          "yyyy/mm/dd"
        );


      /* 値 */
      factsSheet
        .getRange(
          "I:I"
        )
        .setNumberFormat(
          "0.########"
        );


      /* 登録日時 */
      factsSheet
        .getRange(
          "O:O"
        )
        .setNumberFormat(
          "yyyy/mm/dd hh:mm:ss"
        );

    }


    /* =================================================
     * PAY_ADJUSTMENTS 表示設定
     * ================================================= */

    const adjustmentSheet =
      book.getSheetByName(
        "PAY_ADJUSTMENTS"
      );


    if (adjustmentSheet) {

      /* 対象日 */
      adjustmentSheet
        .getRange(
          "B:B"
        )
        .setNumberFormat(
          "yyyy/mm/dd"
        );


      /* 調整値 */
      adjustmentSheet
        .getRange(
          "F:F"
        )
        .setNumberFormat(
          "0.########"
        );


      /* 登録日時 */
      adjustmentSheet
        .getRange(
          "K:L"
        )
        .setNumberFormat(
          "yyyy/mm/dd hh:mm:ss"
        );

    }


    SpreadsheetApp.flush();


    const result = {

      success: true,

      initialized:
        initialized,

      preserved:
        preserved

    };


    Logger.log(
      "給与実績データ用タブ設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 08. 給与取込項目マッピング
 *
 * PAY_IMPORT_MAP
 *
 * Excel等の元項目名を
 * PAY_FACTSの項目キーへ変換するための設定表。
 *
 * GASに列名を固定しない。
 * マッピング内容はシート上で変更可能。
 * ===================================================== */

function setupPayrollImportMap() {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const sheetName =
      "PAY_IMPORT_MAP";


    const headers = [
      "マッピングID",
      "データ種別",
      "元項目名",
      "項目キー",
      "表示名",
      "単位",
      "値種別",
      "倍率",
      "有効",
      "備考"
    ];


    let sheet =
      book.getSheetByName(
        sheetName
      );


    /* =================================================
     * シート作成
     * ================================================= */

    if (!sheet) {

      sheet =
        book.insertSheet(
          sheetName
        );

    }


    /* =================================================
     * 空シートならヘッダー作成
     * ================================================= */

    if (sheet.getLastRow() === 0) {

      cfPayrollEnsureColumns_(
        sheet,
        headers.length
      );


      cfPayrollWriteHeader_(
        sheet,
        headers
      );

    } else {

      /* -----------------------------------------------
       * 既存シートの構成確認
       * ----------------------------------------------- */

      if (
        !cfPayrollHeaderMatches_(
          sheet,
          headers
        )
      ) {

        throw new Error(
          "PAY_IMPORT_MAPの列構成が異なります。" +
          "データ保護のため変更せず停止しました。"
        );

      }

    }


    /* =================================================
     * 初期マッピング
     *
     * 既に存在するマッピングIDは上書きしない。
     * ================================================= */

    const defaults = [

      [
        "MAP_MAIN_NOMINATION",
        "ANY",
        "本指名",
        "BACK_MAIN_NOMINATION",
        "本指名",
        "件",
        "NUMBER",
        1,
        true,
        "本指名件数"
      ],

      [
        "MAP_VIP_NOMINATION",
        "ANY",
        "VIP指名",
        "BACK_VIP_NOMINATION",
        "VIP指名",
        "件",
        "NUMBER",
        1,
        true,
        "VIP指名件数"
      ],

      [
        "MAP_INHOUSE_NOMINATION",
        "ANY",
        "場内",
        "BACK_INHOUSE_NOMINATION",
        "場内指名",
        "件",
        "NUMBER",
        1,
        true,
        "場内指名件数"
      ],

      [
        "MAP_DRINK",
        "ANY",
        "ドリンク",
        "BACK_DRINK",
        "ドリンク",
        "杯",
        "NUMBER",
        1,
        true,
        "ドリンク杯数"
      ],

      [
        "MAP_SHOT",
        "ANY",
        "ショット",
        "BACK_SHOT",
        "ショット",
        "杯",
        "NUMBER",
        1,
        true,
        "ショット杯数"
      ],

      [
        "MAP_BLOG",
        "ANY",
        "ブログ",
        "BACK_BLOG",
        "ブログ",
        "件",
        "NUMBER",
        1,
        true,
        "ブログ件数"
      ],

      [
        "MAP_SALES_TOTAL",
        "ANY",
        "売上",
        "SALES_TOTAL",
        "売上",
        "円",
        "NUMBER",
        1,
        true,
        "給与計算対象売上"
      ],

      [
        "MAP_BOTTLE_SALES",
        "ANY",
        "ボトル売上",
        "BOTTLE_SALES",
        "ボトル売上",
        "円",
        "NUMBER",
        1,
        true,
        "ボトルバック計算用"
      ],

      [
        "MAP_TIP",
        "ANY",
        "チップ",
        "TIP_BACK",
        "チップ",
        "円",
        "NUMBER",
        1,
        true,
        "チップバック"
      ],

      [
        "MAP_WORK_MINUTES",
        "ANY",
        "勤務時間",
        "WORK_MINUTES",
        "勤務時間",
        "分",
        "TIME",
        1,
        true,
        "勤務時間。取込時に分へ変換"
      ]

    ];


    /* =================================================
     * 既存マッピングID取得
     * ================================================= */

    const existing =
      cfPayrollGetExistingKeys_(
        sheet,
        1
      );


    const newRows =
      defaults.filter(function(row) {

        return !existing[row[0]];

      });


    /* =================================================
     * 未登録分だけ追加
     * ================================================= */

    if (newRows.length > 0) {

      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          newRows.length,
          headers.length
        )
        .setValues(
          newRows
        );

    }


    /* =================================================
     * 表示設定
     * ================================================= */

    if (sheet.getLastRow() > 1) {

      /* 倍率 */
      sheet
        .getRange(
          2,
          8,
          sheet.getLastRow() - 1,
          1
        )
        .setNumberFormat(
          "0.########"
        );


      /* 有効 */
      sheet
        .getRange(
          2,
          9,
          sheet.getLastRow() - 1,
          1
        )
        .insertCheckboxes();

    }


    /* =================================================
     * 値種別の入力候補
     * ================================================= */

    const valueTypeRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            "NUMBER",
            "TIME",
            "TEXT",
            "BOOLEAN"
          ],
          true
        )
        .setAllowInvalid(false)
        .build();


    if (sheet.getMaxRows() > 1) {

      sheet
        .getRange(
          2,
          7,
          sheet.getMaxRows() - 1,
          1
        )
        .setDataValidation(
          valueTypeRule
        );

    }


    SpreadsheetApp.flush();


    const result = {

      success: true,

      mappingsAdded:
        newRows.length,

      mappingsTotal:
        sheet.getLastRow() - 1

    };


    Logger.log(
      "給与取込マッピング設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 09. 給与設定読込エンジン
 *
 * PAY_SETTINGS
 * PAY_RULES
 * PAY_IMPORT_MAP
 *
 * をコード側から安全に読み取る。
 *
 * 給与計算・Excel取込では、
 * 原則として固定値を使わず
 * この設定読込処理を経由する。
 * ===================================================== */


/* =====================================================
 * 09-A. PAY_SETTINGS 一括取得
 *
 * 戻り値例：
 *
 * {
 *   TAX_SERVICE_RATE: 0.30,
 *   WORK_UNIT_MINUTES: 15
 * }
 * ===================================================== */

function cfPayrollGetSettings_() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_SETTINGS"
    );


  if (!sheet) {
    throw new Error(
      "PAY_SETTINGSがありません。"
    );
  }


  const expectedHeaders = [
    "設定キー",
    "表示名",
    "設定値",
    "説明"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      expectedHeaders
    )
  ) {
    throw new Error(
      "PAY_SETTINGSの列構成が異なります。"
    );
  }


  const lastRow =
    sheet.getLastRow();


  const result = {};


  if (lastRow <= 1) {
    return result;
  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .getValues();


  rows.forEach(function(row, index) {

    const key =
      String(row[0] || "")
        .trim();


    if (!key) {
      return;
    }


    if (
      Object.prototype
        .hasOwnProperty.call(
          result,
          key
        )
    ) {

      throw new Error(
        "PAY_SETTINGSに設定キーの重複があります：" +
        key +
        "（行 " +
        (index + 2) +
        "）"
      );

    }


    result[key] =
      row[2];

  });


  return result;

}


/* =====================================================
 * 09-B. 指定設定値取得
 *
 * required = true
 *   → 未設定ならエラー
 *
 * required = false
 *   → 未設定ならdefaultValue
 * ===================================================== */

function cfPayrollGetSetting_(
  key,
  defaultValue,
  required
) {

  const settings =
    cfPayrollGetSettings_();


  if (
    Object.prototype
      .hasOwnProperty.call(
        settings,
        key
      )
  ) {

    const value =
      settings[key];


    if (
      value !== "" &&
      value !== null
    ) {
      return value;
    }

  }


  if (required) {

    throw new Error(
      "給与設定が未設定です：" +
      key
    );

  }


  return defaultValue;

}


/* =====================================================
 * 09-C. PAY_IMPORT_MAP取得
 *
 * dataType
 *   例：
 *   SALES
 *   ATTENDANCE
 *   BACK
 *
 * ANYは全データ種別共通
 *
 * 特定データ種別とANYが同じ元項目名を持つ場合、
 * 特定データ種別を優先する。
 * ===================================================== */

function cfPayrollGetImportMappings_(
  dataType
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_MAP"
    );


  if (!sheet) {
    throw new Error(
      "PAY_IMPORT_MAPがありません。"
    );
  }


  const expectedHeaders = [
    "マッピングID",
    "データ種別",
    "元項目名",
    "項目キー",
    "表示名",
    "単位",
    "値種別",
    "倍率",
    "有効",
    "備考"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      expectedHeaders
    )
  ) {
    throw new Error(
      "PAY_IMPORT_MAPの列構成が異なります。"
    );
  }


  const targetType =
    String(
      dataType || "ANY"
    )
      .trim()
      .toUpperCase();


  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {
    return {};
  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        10
      )
      .getValues();


  const candidates = [];


  rows.forEach(function(row, index) {

    const mappingId =
      String(row[0] || "")
        .trim();


    const rowDataType =
      String(row[1] || "ANY")
        .trim()
        .toUpperCase();


    const sourceName =
      String(row[2] || "")
        .trim();


    const itemKey =
      String(row[3] || "")
        .trim();


    const displayName =
      String(row[4] || "")
        .trim();


    const unit =
      String(row[5] || "")
        .trim();


    const valueType =
      String(row[6] || "NUMBER")
        .trim()
        .toUpperCase();


    const multiplierRaw =
      row[7];


    const enabled =
      row[8] === true;


    const note =
      String(row[9] || "")
        .trim();


    /* 無効設定は無視 */
    if (!enabled) {
      return;
    }


    /* 必須値が不足 */
    if (
      !mappingId ||
      !sourceName ||
      !itemKey
    ) {

      throw new Error(
        "PAY_IMPORT_MAPの必須項目が不足しています。" +
        " 行：" +
        (index + 2)
      );

    }


    /* 対象データ種別以外は無視 */
    if (
      rowDataType !== "ANY" &&
      rowDataType !== targetType
    ) {
      return;
    }


    let multiplier = 1;


    if (
      multiplierRaw !== "" &&
      multiplierRaw !== null
    ) {

      multiplier =
        Number(multiplierRaw);


      if (
        !Number.isFinite(
          multiplier
        )
      ) {

        throw new Error(
          "PAY_IMPORT_MAPの倍率が数値ではありません：" +
          mappingId
        );

      }

    }


    candidates.push({

      mappingId:
        mappingId,

      dataType:
        rowDataType,

      sourceName:
        sourceName,

      itemKey:
        itemKey,

      displayName:
        displayName,

      unit:
        unit,

      valueType:
        valueType,

      multiplier:
        multiplier,

      note:
        note,

      sheetRow:
        index + 2

    });

  });


  /* =================================================
   * ANYを先に入れ、
   * 個別データ種別で上書き
   * ================================================= */

  candidates.sort(function(a, b) {

    const priorityA =
      a.dataType === "ANY"
        ? 0
        : 1;


    const priorityB =
      b.dataType === "ANY"
        ? 0
        : 1;


    return (
      priorityA -
      priorityB
    );

  });


  const result = {};


  candidates.forEach(function(item) {

    const normalizedName =
      cfPayrollNormalizeHeaderName_(
        item.sourceName
      );


    result[normalizedName] =
      item;

  });


  return result;

}


/* =====================================================
 * 09-D. 元項目名の正規化
 *
 * 全角スペース・半角スペース等による
 * 微妙な表記揺れを吸収する。
 * ===================================================== */

function cfPayrollNormalizeHeaderName_(
  value
) {

  return String(
    value || ""
  )
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, "")
    .trim()
    .toLowerCase();

}


/* =====================================================
 * 09-E. 取込値変換
 *
 * PAY_IMPORT_MAPの
 *
 * 値種別
 * 倍率
 *
 * を反映する。
 * ===================================================== */

function cfPayrollConvertImportValue_(
  rawValue,
  mapping
) {

  if (!mapping) {
    throw new Error(
      "マッピング情報がありません。"
    );
  }


  const valueType =
    String(
      mapping.valueType ||
      "NUMBER"
    )
      .trim()
      .toUpperCase();


  const multiplier =
    Number(
      mapping.multiplier === "" ||
      mapping.multiplier === null
        ? 1
        : mapping.multiplier
    );


  /* =================================================
   * NUMBER
   * ================================================= */

  if (valueType === "NUMBER") {

    if (
      rawValue === "" ||
      rawValue === null
    ) {
      return 0;
    }


    let value =
      rawValue;


    /*
     * "12,345"
     * "¥12,345"
     * なども数値化
     */
    if (
      typeof value === "string"
    ) {

      value =
        value
          .replace(/,/g, "")
          .replace(/[￥¥]/g, "")
          .trim();

    }


    const number =
      Number(value);


    if (
      !Number.isFinite(
        number
      )
    ) {

      throw new Error(
        "数値として読み取れません：" +
        rawValue
      );

    }


    return (
      number *
      multiplier
    );

  }


  /* =================================================
   * TIME
   *
   * 戻り値は「分」
   * ================================================= */

  if (valueType === "TIME") {

    let minutes = 0;


    /*
     * Google Sheets上の時刻
     * Date型の場合
     */
    if (
      rawValue instanceof Date
    ) {

      minutes =
        rawValue.getHours() * 60 +
        rawValue.getMinutes();

    }


    /*
     * 3:30
     * → 210分
     */
    else if (
      typeof rawValue === "string" &&
      rawValue.indexOf(":") !== -1
    ) {

      const parts =
        rawValue
          .trim()
          .split(":");


      const hours =
        Number(parts[0] || 0);


      const mins =
        Number(parts[1] || 0);


      if (
        !Number.isFinite(hours) ||
        !Number.isFinite(mins)
      ) {

        throw new Error(
          "時間として読み取れません：" +
          rawValue
        );

      }


      minutes =
        hours * 60 +
        mins;

    }


    /*
     * Sheetsの時間シリアル値
     * 0.5 = 12時間
     */
    else if (
      typeof rawValue === "number"
    ) {

      minutes =
        rawValue <= 1
          ? rawValue * 24 * 60
          : rawValue;

    }


    else if (
      rawValue === "" ||
      rawValue === null
    ) {

      minutes = 0;

    }


    else {

      throw new Error(
        "時間として読み取れません：" +
        rawValue
      );

    }


    return (
      minutes *
      multiplier
    );

  }


  /* =================================================
   * BOOLEAN
   * ================================================= */

  if (valueType === "BOOLEAN") {

    if (
      rawValue === true ||
      rawValue === 1 ||
      rawValue === "1"
    ) {
      return true;
    }


    const text =
      String(rawValue || "")
        .trim()
        .toLowerCase();


    return [
      "true",
      "yes",
      "on",
      "あり",
      "有",
      "○",
      "〇"
    ].indexOf(text) !== -1;

  }


  /* =================================================
   * TEXT
   * ================================================= */

  if (valueType === "TEXT") {

    return String(
      rawValue === null ||
      rawValue === undefined
        ? ""
        : rawValue
    );

  }


  throw new Error(
    "未対応の値種別です：" +
    valueType
  );

}


/* =====================================================
 * 09-F. 設定読込テスト
 *
 * GASから手動実行する。
 * ===================================================== */

function testPayrollConfigReader() {

  const settings =
    cfPayrollGetSettings_();


  const mappings =
    cfPayrollGetImportMappings_(
      "ANY"
    );


  const result = {

    success: true,

    settingsCount:
      Object.keys(
        settings
      ).length,

    mappingsCount:
      Object.keys(
        mappings
      ).length,

    taxServiceRate:
      settings.TAX_SERVICE_RATE,

    workUnitMinutes:
      settings.WORK_UNIT_MINUTES

  };


  Logger.log(
    "給与設定読込OK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 10. 給与取込元設定 V2
 *
 * 実ファイル対応版
 *
 * 追加機能：
 * ・日付取得方法
 * ・ファイル名日付形式
 * ・解析方式
 *
 * 旧11列構成から14列構成へ
 * 既存データを残したまま自動移行する。
 * ===================================================== */

function setupPayrollImportSources() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const sheetName =
      "PAY_IMPORT_SOURCES";


    const headers = [
      "取込設定ID",
      "表示名",
      "データ種別",
      "ファイル名キーワード",
      "対象シート名",
      "日付取得方法",
      "日付列名",
      "ファイル名日付形式",
      "キャストID列名",
      "キャスト名列名",
      "開始行",
      "解析方式",
      "有効",
      "備考"
    ];


    const legacyHeaders = [
      "取込設定ID",
      "表示名",
      "データ種別",
      "ファイル名キーワード",
      "対象シート名",
      "日付列名",
      "キャストID列名",
      "キャスト名列名",
      "開始行",
      "有効",
      "備考"
    ];


    let sheet =
      book.getSheetByName(
        sheetName
      );


    if (!sheet) {

      sheet =
        book.insertSheet(
          sheetName
        );

    }


    /* =================================================
     * 新規シート
     * ================================================= */

    if (
      sheet.getLastRow() === 0
    ) {

      cfPayrollEnsureColumns_(
        sheet,
        headers.length
      );


      cfPayrollWriteHeader_(
        sheet,
        headers
      );

    }


    /* =================================================
     * 旧構成 → V2へ移行
     * ================================================= */

    else if (
      cfPayrollHeaderMatches_(
        sheet,
        legacyHeaders
      )
    ) {

      /*
       * 旧：
       *
       * A ID
       * B 表示名
       * C データ種別
       * D ファイル名
       * E シート
       * F 日付列
       * G CAST ID
       * H CAST名
       * I 開始行
       * J 有効
       * K 備考
       */


      /*
       * F列へ「日付取得方法」
       */
      sheet.insertColumnBefore(
        6
      );


      /*
       * H列へ
       * 「ファイル名日付形式」
       */
      sheet.insertColumnBefore(
        8
      );


      /*
       * L列へ「解析方式」
       */
      sheet.insertColumnBefore(
        12
      );


      cfPayrollEnsureColumns_(
        sheet,
        headers.length
      );


      cfPayrollWriteHeader_(
        sheet,
        headers
      );


      /* ---------------------------------------------
       * 旧データへ初期値設定
       * --------------------------------------------- */

      const lastRow =
        sheet.getLastRow();


      if (lastRow > 1) {

        const rows =
          sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              headers.length
            )
            .getValues();


        rows.forEach(function(
          row
        ) {

          /*
           * 日付列名が設定済みなら
           * COLUMN扱い
           */
          if (!row[5]) {

            row[5] =
              row[6]
                ? "COLUMN"
                : "NONE";

          }


          /*
           * 旧設定は通常解析
           */
          if (!row[11]) {

            row[11] =
              "GENERIC";

          }

        });


        sheet
          .getRange(
            2,
            1,
            rows.length,
            headers.length
          )
          .setValues(
            rows
          );

      }

    }


    /* =================================================
     * V2構成確認
     * ================================================= */

    else if (
      !cfPayrollHeaderMatches_(
        sheet,
        headers
      )
    ) {

      throw new Error(
        "PAY_IMPORT_SOURCESの列構成が想定と異なります。" +
        "データ保護のため変更せず停止しました。"
      );

    }


    /* =================================================
     * 実ファイル5種類
     *
     * ※ 最初は全て無効
     * ※ 各解析エンジン完成後にONにする
     * ================================================= */

    const defaults = [


      /* ---------------------------------------------
       * オーダー売上照会
       * --------------------------------------------- */

      [
        "SRC_ORDER_SALES",
        "オーダー売上照会",
        "ORDER",
        "オーダー売上照会",
        "Sheet1",

        "FILENAME",
        "",
        "YYYYMMDD",

        "キャストコード",
        "キャスト名",

        2,

        "PRODUCT_ROWS",

        false,

        "ドリンク・ショット・クライナー等の商品別実績"
      ],


      /* ---------------------------------------------
       * キャスト勤怠照会
       * --------------------------------------------- */

      [
        "SRC_CAST_ATTENDANCE",
        "キャスト勤怠照会",
        "ATTENDANCE",
        "キャスト勤怠照会",
        "Sheet1",

        "FILENAME",
        "",
        "YYYYMMDD",

        "コード",
        "キャスト名",

        2,

        "ATTENDANCE_ROWS",

        false,

        "出勤予定・退勤予定・実出勤・実退勤・欠勤"
      ],


      /* ---------------------------------------------
       * キャスト成績一覧照会
       * --------------------------------------------- */

      [
        "SRC_CAST_PERFORMANCE",
        "キャスト成績一覧照会",
        "PERFORMANCE",
        "キャスト成績一覧照会",
        "Sheet1",

        "FILENAME",
        "",
        "YYYYMMDD",

        "コード",
        "キャスト名",

        2,

        "PERFORMANCE_ROWS",

        false,

        "指名・同伴・場内・送り・前払金・請求実績等"
      ],


      /* ---------------------------------------------
       * 精算入金一覧照会
       * --------------------------------------------- */

      [
        "SRC_SETTLEMENT",
        "精算入金一覧照会",
        "SETTLEMENT",
        "精算入金一覧照会",
        "Sheet1",

        "FILENAME",
        "",
        "YYYYMMDD",

        "",
        "キャスト",

        2,

        "SETTLEMENT_DETAIL",

        false,

        "伝票・入金・注文明細。主に検算と追跡用"
      ],


      /* ---------------------------------------------
       * 伝票一覧照会
       * --------------------------------------------- */

      [
        "SRC_SLIP_LIST",
        "伝票一覧照会",
        "SLIP",
        "伝票一覧照会",
        "Sheet1",

        "FILENAME",
        "",
        "YYYYMMDD",

        "コード",
        "キャスト名",

        2,

        "SLIP_ROWS",

        false,

        "純売実績・ボトル実績・F&D実績等"
      ]

    ];


    /* =================================================
     * 既存ID取得
     * ================================================= */

    const existing =
      cfPayrollGetExistingKeys_(
        sheet,
        1
      );


    const newRows =
      defaults.filter(function(
        row
      ) {

        return !existing[
          row[0]
        ];

      });


    /* =================================================
     * 新規分だけ追加
     * ================================================= */

    if (
      newRows.length > 0
    ) {

      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          newRows.length,
          headers.length
        )
        .setValues(
          newRows
        );

    }


    /* =================================================
     * 日付取得方法
     * ================================================= */

    const dateMethodRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            "COLUMN",
            "FILENAME",
            "NONE"
          ],
          true
        )
        .setAllowInvalid(
          false
        )
        .build();


    sheet
      .getRange(
        2,
        6,
        sheet.getMaxRows() - 1,
        1
      )
      .setDataValidation(
        dateMethodRule
      );


    /* =================================================
     * 解析方式
     * ================================================= */

    const parserRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            "GENERIC",
            "PRODUCT_ROWS",
            "ATTENDANCE_ROWS",
            "PERFORMANCE_ROWS",
            "SETTLEMENT_DETAIL",
            "SLIP_ROWS"
          ],
          true
        )
        .setAllowInvalid(
          false
        )
        .build();


    sheet
      .getRange(
        2,
        12,
        sheet.getMaxRows() - 1,
        1
      )
      .setDataValidation(
        parserRule
      );


    /* =================================================
     * 有効
     * ================================================= */

    sheet
      .getRange(
        2,
        13,
        sheet.getMaxRows() - 1,
        1
      )
      .insertCheckboxes();


    /* =================================================
     * 開始行
     * ================================================= */

    sheet
      .getRange(
        2,
        11,
        sheet.getMaxRows() - 1,
        1
      )
      .setNumberFormat(
        "0"
      );


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      sourcesAdded:
        newRows.length,

      sourcesTotal:
        sheet.getLastRow() - 1

    };


    Logger.log(
      "給与取込元設定V2 OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 10-A. 有効な取込設定一覧取得 V2
 * ===================================================== */

function cfPayrollGetImportSources_() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_SOURCES"
    );


  if (!sheet) {

    throw new Error(
      "PAY_IMPORT_SOURCESがありません。"
    );

  }


  const headers = [
    "取込設定ID",
    "表示名",
    "データ種別",
    "ファイル名キーワード",
    "対象シート名",
    "日付取得方法",
    "日付列名",
    "ファイル名日付形式",
    "キャストID列名",
    "キャスト名列名",
    "開始行",
    "解析方式",
    "有効",
    "備考"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      headers
    )
  ) {

    throw new Error(
      "PAY_IMPORT_SOURCESの列構成が異なります。"
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return [];

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const result = [];


  rows.forEach(function(
    row,
    index
  ) {

    const sourceId =
      String(
        row[0] || ""
      ).trim();


    const displayName =
      String(
        row[1] || ""
      ).trim();


    const dataType =
      String(
        row[2] || ""
      )
        .trim()
        .toUpperCase();


    const fileKeyword =
      String(
        row[3] || ""
      ).trim();


    const sheetName =
      String(
        row[4] || ""
      ).trim();


    let dateMethod =
      String(
        row[5] || ""
      )
        .trim()
        .toUpperCase();


    const dateHeader =
      String(
        row[6] || ""
      ).trim();


    const filenameDateFormat =
      String(
        row[7] || ""
      )
        .trim()
        .toUpperCase();


    const castIdHeader =
      String(
        row[8] || ""
      ).trim();


    const castNameHeader =
      String(
        row[9] || ""
      ).trim();


    const startRow =
      Number(
        row[10] || 2
      );


    const parserMode =
      String(
        row[11] ||
        "GENERIC"
      )
        .trim()
        .toUpperCase();


    const enabled =
      row[12] === true;


    const note =
      String(
        row[13] || ""
      ).trim();


    /*
     * 無効設定は使用しない
     */
    if (!enabled) {
      return;
    }


    if (!sourceId) {

      throw new Error(
        "PAY_IMPORT_SOURCESの取込設定IDが空です。" +
        " 行：" +
        (index + 2)
      );

    }


    if (!dataType) {

      throw new Error(
        sourceId +
        " のデータ種別が未設定です。"
      );

    }


    if (
      !Number.isInteger(
        startRow
      ) ||
      startRow < 2
    ) {

      throw new Error(
        sourceId +
        " の開始行が不正です。"
      );

    }


    /* ---------------------------------------------
     * 日付取得方法
     * --------------------------------------------- */

    if (!dateMethod) {

      dateMethod =
        dateHeader
          ? "COLUMN"
          : "NONE";

    }


    if (
      [
        "COLUMN",
        "FILENAME",
        "NONE"
      ].indexOf(
        dateMethod
      ) === -1
    ) {

      throw new Error(
        sourceId +
        " の日付取得方法が不正です：" +
        dateMethod
      );

    }


    if (
      dateMethod ===
        "COLUMN" &&
      !dateHeader
    ) {

      throw new Error(
        sourceId +
        " はCOLUMN方式ですが、日付列名が未設定です。"
      );

    }


    if (
      dateMethod ===
        "FILENAME" &&
      !filenameDateFormat
    ) {

      throw new Error(
        sourceId +
        " はFILENAME方式ですが、" +
        "ファイル名日付形式が未設定です。"
      );

    }


    result.push({

      sourceId:
        sourceId,

      displayName:
        displayName,

      dataType:
        dataType,

      fileKeyword:
        fileKeyword,

      sheetName:
        sheetName,

      dateMethod:
        dateMethod,

      dateHeader:
        dateHeader,

      filenameDateFormat:
        filenameDateFormat,

      castIdHeader:
        castIdHeader,

      castNameHeader:
        castNameHeader,

      startRow:
        startRow,

      parserMode:
        parserMode,

      note:
        note,

      sheetRow:
        index + 2

    });

  });


  return result;

}


/* =====================================================
 * 10-B. 取込設定読込テスト
 * ===================================================== */

function testPayrollImportSources() {

  const sources =
    cfPayrollGetImportSources_();


  const result = {

    success:
      true,

    enabledSources:
      sources.length,

    sources:
      sources.map(function(
        source
      ) {

        return {

          sourceId:
            source.sourceId,

          displayName:
            source.displayName,

          dataType:
            source.dataType,

          dateMethod:
            source.dateMethod,

          parserMode:
            source.parserMode

        };

      })

  };


  Logger.log(
    "給与取込元読込OK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 11. 給与取込プレビュー
 *
 * 実データをPAY_FACTSへ書き込まず、
 * ファイルの読み取り結果だけ確認する。
 *
 * 現段階：
 * ・Googleスプレッドシート対応
 * ・PAY_IMPORT_SOURCESの設定を使用
 * ・PAY_IMPORT_MAPの設定を使用
 * ・CAST_ID / キャスト名の認識確認
 * ・実績項目の変換確認
 *
 * 本登録はまだ行わない。
 * ===================================================== */


/* =====================================================
 * 11-A. ファイルID + 取込設定IDでプレビュー
 *
 * Google Sheets / Excel / CSV 共通入口。
 *
 * 解析方式に応じて
 * GENERIC / PRODUCT_ROWS 等を自動振り分け。
 *
 * Section 19の日付取得処理も反映する。
 * ===================================================== */

function cfPayrollPreviewImport_(
  fileId,
  sourceId
) {

  if (!fileId) {

    throw new Error(
      "取込対象のファイルIDがありません。"
    );

  }


  if (!sourceId) {

    throw new Error(
      "取込設定IDがありません。"
    );

  }


  const source =
    cfPayrollGetImportSourceById_(
      sourceId
    );


  if (!source) {

    throw new Error(
      "有効な取込設定が見つかりません：" +
      sourceId
    );

  }


  return cfPayrollWithReadableSpreadsheet_(

    fileId,

    function(
      sourceBook,
      prepared
    ) {

      /*
       * 解析方式ごとに振り分け
       */
      const preview =
        cfPayrollAnalyzeImportSpreadsheetV2_(
          sourceBook,
          prepared,
          source
        );


      /*
       * COLUMN / FILENAME / NONE
       * の日付取得方式を反映
       */
      return cfPayrollApplyImportDateStrategy_(
        preview,
        prepared,
        source
      );

    }

  );

}


/* =====================================================
 * 11-B. 取込設定ID検索
 * ===================================================== */

function cfPayrollGetImportSourceById_(
  sourceId
) {

  const sources =
    cfPayrollGetImportSources_();


  const target =
    String(
      sourceId || ""
    ).trim();


  for (
    let i = 0;
    i < sources.length;
    i++
  ) {

    if (
      sources[i].sourceId ===
      target
    ) {

      return sources[i];

    }

  }


  return null;

}


/* =====================================================
 * 11-C. 設定された列名を探す
 *
 * 空欄設定なら -1
 * ===================================================== */

function cfPayrollFindConfiguredHeaderIndex_(
  headers,
  configuredHeader
) {

  const target =
    cfPayrollNormalizeHeaderName_(
      configuredHeader
    );


  if (!target) {
    return -1;
  }


  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    if (
      cfPayrollNormalizeHeaderName_(
        headers[i]
      ) === target
    ) {

      return i;

    }

  }


  return -1;

}


/* =====================================================
 * 11-D. PAY_CAST_MAP検索辞書
 *
 * 安全設計：
 *
 * キャストコードだけでは人を確定しない。
 *
 * 例：
 *
 * 0030 + さやか
 * 0030 + マジン
 *
 * は別人物候補として扱う。
 *
 * 優先：
 * 1. 元ID + 名前
 * 2. 名前
 * 3. 名前が無い場合のみ元ID
 * ===================================================== */

function cfPayrollBuildCastResolver_() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_CAST_MAP"
    );


  if (!sheet) {

    throw new Error(
      "PAY_CAST_MAPがありません。"
    );

  }


  const expectedHeaders = [
    "元システム",
    "元キャストID",
    "元キャスト名",
    "CAST_ID",
    "確認状態",
    "更新日時"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      expectedHeaders
    )
  ) {

    throw new Error(
      "PAY_CAST_MAPの列構成が異なります。"
    );

  }


  const result = {

    /*
     * 0030::さやか
     */
    byIdName: {},

    /*
     * IDしか取れないデータ用。
     *
     * 同一IDが別CAST_IDへ繋がった場合は
     * nullにして使用禁止。
     */
    bySourceId: {},

    /*
     * 名前が一意な場合用。
     */
    byName: {}

  };


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return result;

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  rows.forEach(function(
    row
  ) {

    const sourceCastId =
      cfPayrollNormalizeExternalCastId_(
        row[1]
      );


    const sourceCastName =
      String(
        row[2] || ""
      ).trim();


    const castId =
      String(
        row[3] || ""
      ).trim();


    /*
     * 正式CAST_IDが無い行は
     * 解決辞書には入れない。
     */
    if (!castId) {

      return;

    }


    const normalizedName =
      cfPayrollNormalizeCastName_(
        sourceCastName
      );


    /* =================================================
     * 元ID + 名前
     * ================================================= */

    if (
      sourceCastId &&
      normalizedName
    ) {

      const key =
        sourceCastId +
        "::" +
        normalizedName;


      if (
        Object.prototype
          .hasOwnProperty.call(
            result.byIdName,
            key
          )
      ) {

        if (
          result.byIdName[
            key
          ] !== castId
        ) {

          /*
           * 同じID+名前が
           * 複数CAST_IDへ繋がっている。
           *
           * 自動判定禁止。
           */
          result.byIdName[
            key
          ] =
            null;

        }

      } else {

        result.byIdName[
          key
        ] =
          castId;

      }

    }


    /* =================================================
     * 名前
     * ================================================= */

    if (
      normalizedName
    ) {

      if (
        Object.prototype
          .hasOwnProperty.call(
            result.byName,
            normalizedName
          )
      ) {

        if (
          result.byName[
            normalizedName
          ] !== castId
        ) {

          /*
           * 同名別人
           */
          result.byName[
            normalizedName
          ] =
            null;

        }

      } else {

        result.byName[
          normalizedName
        ] =
          castId;

      }

    }


    /* =================================================
     * IDのみ
     *
     * 名前が無いデータでのみ後から使用。
     * ================================================= */

    if (
      sourceCastId
    ) {

      if (
        Object.prototype
          .hasOwnProperty.call(
            result.bySourceId,
            sourceCastId
          )
      ) {

        if (
          result.bySourceId[
            sourceCastId
          ] !== castId
        ) {

          result.bySourceId[
            sourceCastId
          ] =
            null;

        }

      } else {

        result.bySourceId[
          sourceCastId
        ] =
          castId;

      }

    }

  });


  return result;

}


/* =====================================================
 * 11-E. CAST_ID解決
 *
 * 安全優先順位：
 *
 * 1. 元ID + 名前
 * 2. 名前
 * 3. 名前が無い場合だけ元ID
 *
 * 同じコードで別の名前が存在しても
 * コードだけで勝手に紐付けない。
 * ===================================================== */

function cfPayrollResolveCastId_(
  resolver,
  sourceCastId,
  sourceCastName
) {

  const normalizedSourceId =
    cfPayrollNormalizeExternalCastId_(
      sourceCastId
    );


  const normalizedName =
    cfPayrollNormalizeCastName_(
      sourceCastName
    );


  /* =================================================
   * ID + 名前
   * ================================================= */

  if (
    normalizedSourceId &&
    normalizedName
  ) {

    const key =
      normalizedSourceId +
      "::" +
      normalizedName;


    if (
      Object.prototype
        .hasOwnProperty.call(
          resolver.byIdName,
          key
        )
    ) {

      /*
       * nullなら競合しているため
       * 自動判定しない。
       */
      return (
        resolver.byIdName[
          key
        ] ||
        ""
      );

    }

  }


  /* =================================================
   * 名前
   *
   * 名前が一意に登録済みなら
   * CAST_IDへ接続。
   * ================================================= */

  if (
    normalizedName
  ) {

    if (
      Object.prototype
        .hasOwnProperty.call(
          resolver.byName,
          normalizedName
        )
    ) {

      return (
        resolver.byName[
          normalizedName
        ] ||
        ""
      );

    }


    /*
     * 名前があるのに一致しない場合、
     * IDだけで勝手に判定しない。
     */
    return "";

  }


  /* =================================================
   * 名前が取得できないデータのみ
   * ID単独判定を許可。
   * ================================================= */

  if (
    normalizedSourceId &&
    Object.prototype
      .hasOwnProperty.call(
        resolver.bySourceId,
        normalizedSourceId
      )
  ) {

    return (
      resolver.bySourceId[
        normalizedSourceId
      ] ||
      ""
    );

  }


  return "";

}


/* =====================================================
 * 11-F. キャスト名正規化
 * ===================================================== */

function cfPayrollNormalizeCastName_(
  value
) {

  return String(
    value || ""
  )
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, "")
    .trim()
    .toLowerCase();

}


/* =====================================================
 * 11-G. 日付正規化
 *
 * 戻り値：
 * Date型（時刻00:00:00）
 * ===================================================== */

function cfPayrollNormalizeDate_(
  rawValue
) {

  if (
    rawValue === "" ||
    rawValue === null
  ) {
    return "";
  }


  let date;


  if (
    rawValue instanceof Date
  ) {

    date =
      new Date(
        rawValue.getTime()
      );

  } else {

    date =
      new Date(
        rawValue
      );

  }


  if (
    Number.isNaN(
      date.getTime()
    )
  ) {

    throw new Error(
      "日付として読み取れません：" +
      rawValue
    );

  }


  date.setHours(
    0,
    0,
    0,
    0
  );


  return date;

}


/* =====================================================
 * 12. 給与取込フォルダ基盤
 *
 * Google Drive上に
 *
 * CAST FLOW PAYROLL
 * ├─ 01_取込待ち
 * ├─ 02_処理済み
 * └─ 03_エラー
 *
 * を作成・管理する。
 *
 * フォルダIDはPAY_SETTINGSへ保存する。
 *
 * この段階ではファイルの取込処理は行わない。
 * ===================================================== */


/* =====================================================
 * 12-A. 給与取込フォルダ初期設定
 * ===================================================== */

function setupPayrollImportFolders() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const payrollBook =
      cfPayrollOpenSpreadsheet_();


    /* =================================================
     * ルートフォルダ
     * ================================================= */

    const rootFolder =
      cfPayrollResolveRootFolder_(
        payrollBook
      );


    /* =================================================
     * 子フォルダ
     * ================================================= */

    const inboxFolder =
      cfPayrollGetOrCreateChildFolder_(
        rootFolder,
        "01_取込待ち"
      );


    const processedFolder =
      cfPayrollGetOrCreateChildFolder_(
        rootFolder,
        "02_処理済み"
      );


    const errorFolder =
      cfPayrollGetOrCreateChildFolder_(
        rootFolder,
        "03_エラー"
      );


    /* =================================================
     * PAY_SETTINGSへID保存
     * ================================================= */

    cfPayrollSaveSystemSetting_(
      "PAYROLL_IMPORT_ROOT_FOLDER_ID",
      "給与取込ルートフォルダID",
      rootFolder.getId(),
      "システム管理用。通常は変更しない"
    );


    cfPayrollSaveSystemSetting_(
      "PAYROLL_IMPORT_INBOX_FOLDER_ID",
      "取込待ちフォルダID",
      inboxFolder.getId(),
      "01_取込待ち。システム管理用"
    );


    cfPayrollSaveSystemSetting_(
      "PAYROLL_IMPORT_PROCESSED_FOLDER_ID",
      "処理済みフォルダID",
      processedFolder.getId(),
      "02_処理済み。システム管理用"
    );


    cfPayrollSaveSystemSetting_(
      "PAYROLL_IMPORT_ERROR_FOLDER_ID",
      "エラーフォルダID",
      errorFolder.getId(),
      "03_エラー。システム管理用"
    );


    SpreadsheetApp.flush();


    const result = {

      success: true,

      rootFolder: {
        id: rootFolder.getId(),
        name: rootFolder.getName()
      },

      inboxFolder: {
        id: inboxFolder.getId(),
        name: inboxFolder.getName()
      },

      processedFolder: {
        id: processedFolder.getId(),
        name: processedFolder.getName()
      },

      errorFolder: {
        id: errorFolder.getId(),
        name: errorFolder.getName()
      }

    };


    Logger.log(
      "給与取込フォルダ設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 12-B. ルートフォルダ取得・作成
 *
 * PAY_SETTINGSに既存IDがあればそれを使用。
 *
 * 未設定の場合：
 * 給与スプレッドシートと同じ親フォルダへ
 * CAST FLOW PAYROLL を作成。
 * ===================================================== */

function cfPayrollResolveRootFolder_(
  payrollBook
) {

  const settings =
    cfPayrollGetSettings_();


  const existingId =
    String(
      settings
        .PAYROLL_IMPORT_ROOT_FOLDER_ID ||
      ""
    ).trim();


  /* =================================================
   * 既存設定あり
   * ================================================= */

  if (existingId) {

    try {

      return DriveApp.getFolderById(
        existingId
      );

    } catch (error) {

      throw new Error(
        "PAYROLL_IMPORT_ROOT_FOLDER_IDで設定された" +
        "フォルダを取得できません。" +
        " 削除または権限変更されていないか確認してください。" +
        " ID：" +
        existingId
      );

    }

  }


  /* =================================================
   * 給与スプレッドシートの親フォルダ取得
   * ================================================= */

  const payrollFile =
    DriveApp.getFileById(
      payrollBook.getId()
    );


  const parents =
    payrollFile.getParents();


  let parentFolder;


  if (parents.hasNext()) {

    parentFolder =
      parents.next();

  } else {

    parentFolder =
      DriveApp.getRootFolder();

  }


  /* =================================================
   * 既に同名フォルダがあれば再利用
   * ================================================= */

  const existingFolders =
    parentFolder.getFoldersByName(
      "CAST FLOW PAYROLL"
    );


  if (existingFolders.hasNext()) {

    return existingFolders.next();

  }


  /* =================================================
   * 新規作成
   * ================================================= */

  return parentFolder.createFolder(
    "CAST FLOW PAYROLL"
  );

}


/* =====================================================
 * 12-C. 子フォルダ取得・作成
 * ===================================================== */

function cfPayrollGetOrCreateChildFolder_(
  parentFolder,
  folderName
) {

  const folders =
    parentFolder.getFoldersByName(
      folderName
    );


  if (folders.hasNext()) {

    return folders.next();

  }


  return parentFolder.createFolder(
    folderName
  );

}


/* =====================================================
 * 12-D. システム設定保存
 *
 * ・キーが無ければ追加
 * ・値が空なら設定
 * ・既に値がある場合は上書きしない
 *
 * フォルダIDを誤って再生成しないための安全策。
 * ===================================================== */

function cfPayrollSaveSystemSetting_(
  key,
  displayName,
  value,
  description
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_SETTINGS"
    );


  if (!sheet) {

    throw new Error(
      "PAY_SETTINGSがありません。"
    );

  }


  const headers = [
    "設定キー",
    "表示名",
    "設定値",
    "説明"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      headers
    )
  ) {

    throw new Error(
      "PAY_SETTINGSの列構成が異なります。"
    );

  }


  const lastRow =
    sheet.getLastRow();


  /* =================================================
   * 既存キー検索
   * ================================================= */

  if (lastRow > 1) {

    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .getValues();


    for (
      let i = 0;
      i < rows.length;
      i++
    ) {

      const existingKey =
        String(
          rows[i][0] || ""
        ).trim();


      if (
        existingKey !== key
      ) {
        continue;
      }


      const existingValue =
        String(
          rows[i][2] || ""
        ).trim();


      /*
       * 既に値がある場合は保護
       */
      if (existingValue) {

        return existingValue;

      }


      /*
       * キーはあるが値が空
       */
      sheet
        .getRange(
          i + 2,
          2,
          1,
          3
        )
        .setValues([[
          displayName,
          value,
          description
        ]]);


      return value;

    }

  }


  /* =================================================
   * 新規設定追加
   * ================================================= */

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      4
    )
    .setValues([[
      key,
      displayName,
      value,
      description
    ]]);


  return value;

}


/* =====================================================
 * 12-E. 取込待ちフォルダ取得
 * ===================================================== */

function cfPayrollGetInboxFolder_() {

  const folderId =
    String(
      cfPayrollGetSetting_(
        "PAYROLL_IMPORT_INBOX_FOLDER_ID",
        "",
        true
      )
    ).trim();


  try {

    return DriveApp.getFolderById(
      folderId
    );

  } catch (error) {

    throw new Error(
      "取込待ちフォルダを取得できません。" +
      " setupPayrollImportFolders() を確認してください。"
    );

  }

}


/* =====================================================
 * 12-F. 処理済みフォルダ取得
 * ===================================================== */

function cfPayrollGetProcessedFolder_() {

  const folderId =
    String(
      cfPayrollGetSetting_(
        "PAYROLL_IMPORT_PROCESSED_FOLDER_ID",
        "",
        true
      )
    ).trim();


  try {

    return DriveApp.getFolderById(
      folderId
    );

  } catch (error) {

    throw new Error(
      "処理済みフォルダを取得できません。"
    );

  }

}


/* =====================================================
 * 12-G. エラーフォルダ取得
 * ===================================================== */

function cfPayrollGetErrorFolder_() {

  const folderId =
    String(
      cfPayrollGetSetting_(
        "PAYROLL_IMPORT_ERROR_FOLDER_ID",
        "",
        true
      )
    ).trim();


  try {

    return DriveApp.getFolderById(
      folderId
    );

  } catch (error) {

    throw new Error(
      "エラーフォルダを取得できません。"
    );

  }

}


/* =====================================================
 * 12-H. 取込待ちファイル一覧確認
 *
 * 読み取りのみ。
 * ファイル移動・給与反映はしない。
 * ===================================================== */

function testPayrollInboxFiles() {

  const folder =
    cfPayrollGetInboxFolder_();


  const files =
    folder.getFiles();


  const result = [];


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    result.push({

      id:
        file.getId(),

      name:
        file.getName(),

      mimeType:
        file.getMimeType(),

      size:
        file.getSize(),

      lastUpdated:
        file.getLastUpdated()

    });

  }


  Logger.log(
    "給与取込待ちファイル: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 13. 取込待ちファイル自動判定
 *
 * 01_取込待ち に入っているファイルを確認し、
 * PAY_IMPORT_SOURCES の
 * 「ファイル名キーワード」から取込設定を判定する。
 *
 * この段階では
 * ・ファイル移動しない
 * ・PAY_FACTSへ書き込まない
 * ・ログ登録しない
 *
 * 完全な読み取り専用。
 * ===================================================== */


/* =====================================================
 * 13-A. 取込待ちファイル判定プレビュー
 * ===================================================== */

function previewPayrollInboxDetection() {

  const inboxFolder =
    cfPayrollGetInboxFolder_();


  const sources =
    cfPayrollGetImportSources_();


  const files =
    inboxFolder.getFiles();


  const results = [];


  let matchedCount = 0;
  let unmatchedCount = 0;
  let ambiguousCount = 0;
  let unsupportedCount = 0;


  while (files.hasNext()) {

    const file =
      files.next();


    const fileInfo =
      cfPayrollBuildImportFileInfo_(
        file
      );


    const detection =
      cfPayrollDetectImportSourceForFile_(
        fileInfo,
        sources
      );


    if (
      detection.status ===
      "MATCHED"
    ) {

      matchedCount++;

    } else if (
      detection.status ===
      "UNMATCHED"
    ) {

      unmatchedCount++;

    } else if (
      detection.status ===
      "AMBIGUOUS"
    ) {

      ambiguousCount++;

    } else if (
      detection.status ===
      "UNSUPPORTED"
    ) {

      unsupportedCount++;

    }


    results.push({

      file:
        fileInfo,

      detection:
        detection

    });

  }


  /* ファイル名順 */
  results.sort(function(a, b) {

    return a.file.name
      .localeCompare(
        b.file.name,
        "ja"
      );

  });


  const result = {

    success: true,

    mode:
      "DETECTION_ONLY",

    summary: {

      totalFiles:
        results.length,

      matched:
        matchedCount,

      unmatched:
        unmatchedCount,

      ambiguous:
        ambiguousCount,

      unsupported:
        unsupportedCount

    },

    files:
      results

  };


  Logger.log(
    "給与取込ファイル判定: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 13-B. ファイル基本情報生成
 * ===================================================== */

function cfPayrollBuildImportFileInfo_(
  file
) {

  const mimeType =
    String(
      file.getMimeType() || ""
    );


  return {

    id:
      file.getId(),

    name:
      file.getName(),

    mimeType:
      mimeType,

    fileType:
      cfPayrollDetectFileType_(
        mimeType,
        file.getName()
      ),

    size:
      file.getSize(),

    lastUpdated:
      file.getLastUpdated()

  };

}


/* =====================================================
 * 13-C. ファイル種類判定
 *
 * GOOGLE_SHEETS
 * XLSX
 * XLS
 * CSV
 * OTHER
 * ===================================================== */

function cfPayrollDetectFileType_(
  mimeType,
  fileName
) {

  const mime =
    String(
      mimeType || ""
    ).toLowerCase();


  const name =
    String(
      fileName || ""
    ).toLowerCase();


  if (
    mime ===
    "application/vnd.google-apps.spreadsheet"
  ) {

    return "GOOGLE_SHEETS";

  }


  if (
    mime ===
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
    name.endsWith(".xlsx")
  ) {

    return "XLSX";

  }


  if (
    mime ===
      "application/vnd.ms-excel" ||
    name.endsWith(".xls")
  ) {

    return "XLS";

  }


  if (
    mime === "text/csv" ||
    name.endsWith(".csv")
  ) {

    return "CSV";

  }


  return "OTHER";

}


/* =====================================================
 * 13-D. ファイル → 取込設定判定
 * ===================================================== */

function cfPayrollDetectImportSourceForFile_(
  fileInfo,
  sources
) {

  /* =================================================
   * 未対応ファイル
   * ================================================= */

  const supportedTypes = [
    "GOOGLE_SHEETS",
    "XLSX",
    "XLS",
    "CSV"
  ];


  if (
    supportedTypes.indexOf(
      fileInfo.fileType
    ) === -1
  ) {

    return {

      status:
        "UNSUPPORTED",

      sourceId:
        "",

      displayName:
        "",

      reason:
        "未対応のファイル形式です。"

    };

  }


  const normalizedFileName =
    cfPayrollNormalizeFileName_(
      fileInfo.name
    );


  const matches = [];


  sources.forEach(function(source) {

    const keyword =
      cfPayrollNormalizeFileName_(
        source.fileKeyword
      );


    /*
     * キーワード空欄は
     * 自動判定には使用しない
     */
    if (!keyword) {
      return;
    }


    if (
      normalizedFileName.indexOf(
        keyword
      ) !== -1
    ) {

      matches.push({

        sourceId:
          source.sourceId,

        displayName:
          source.displayName,

        dataType:
          source.dataType,

        fileKeyword:
          source.fileKeyword,

        sheetName:
          source.sheetName

      });

    }

  });


  /* =================================================
   * 一致なし
   * ================================================= */

  if (matches.length === 0) {

    return {

      status:
        "UNMATCHED",

      sourceId:
        "",

      displayName:
        "",

      reason:
        "ファイル名キーワードに一致する有効な取込設定がありません。"

    };

  }


  /* =================================================
   * 複数一致
   *
   * 勝手に選ばず停止。
   * 設定ミスを検出する。
   * ================================================= */

  if (matches.length > 1) {

    return {

      status:
        "AMBIGUOUS",

      sourceId:
        "",

      displayName:
        "",

      reason:
        "複数の取込設定に一致しました。",

      matches:
        matches

    };

  }


  /* =================================================
   * 1件だけ一致
   * ================================================= */

  return {

    status:
      "MATCHED",

    sourceId:
      matches[0].sourceId,

    displayName:
      matches[0].displayName,

    dataType:
      matches[0].dataType,

    fileKeyword:
      matches[0].fileKeyword,

    sheetName:
      matches[0].sheetName,

    reason:
      "取込設定を自動判定しました。"

  };

}


/* =====================================================
 * 13-E. ファイル名正規化
 *
 * 大文字小文字
 * 全角半角
 * スペース
 *
 * の表記揺れを減らす。
 * ===================================================== */

function cfPayrollNormalizeFileName_(
  value
) {

  let text =
    String(
      value || ""
    );


  /*
   * Unicode正規化
   * 例：
   * 全角英数字 → 半角寄りへ
   */
  if (
    typeof text.normalize ===
    "function"
  ) {

    text =
      text.normalize(
        "NFKC"
      );

  }


  return text
    .replace(/\u3000/g, "")
    .replace(/\s+/g, "")
    .trim()
    .toLowerCase();

}


/* =====================================================
 * 14. Excel / CSV 読込基盤
 *
 * 対応形式：
 * ・Googleスプレッドシート
 * ・.xlsx
 * ・.xls
 * ・.csv
 *
 * Excel / CSVの場合：
 *
 * 元ファイル
 *   ↓
 * 04_一時変換 にGoogle Sheets版を一時作成
 *   ↓
 * 読み取り
 *   ↓
 * 一時ファイルだけゴミ箱へ
 *
 * 元ファイルには一切変更を加えない。
 *
 * ※ Apps Scriptの拡張サービス
 *    「Drive API」を有効にすること
 * ===================================================== */


/* =====================================================
 * 14-A. 一時変換フォルダ設定
 * ===================================================== */

function setupPayrollConversionFolder() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const rootFolderId =
      String(
        cfPayrollGetSetting_(
          "PAYROLL_IMPORT_ROOT_FOLDER_ID",
          "",
          true
        )
      ).trim();


    let rootFolder;


    try {

      rootFolder =
        DriveApp.getFolderById(
          rootFolderId
        );

    } catch (error) {

      throw new Error(
        "給与取込ルートフォルダを取得できません。" +
        "先に setupPayrollImportFolders() を実行してください。"
      );

    }


    const tempFolder =
      cfPayrollGetOrCreateChildFolder_(
        rootFolder,
        "04_一時変換"
      );


    cfPayrollSaveSystemSetting_(
      "PAYROLL_IMPORT_TEMP_FOLDER_ID",
      "一時変換フォルダID",
      tempFolder.getId(),
      "Excel等を一時的にGoogle Sheetsへ変換する場所。通常変更禁止"
    );


    SpreadsheetApp.flush();


    const result = {

      success: true,

      folderId:
        tempFolder.getId(),

      folderName:
        tempFolder.getName()

    };


    Logger.log(
      "給与一時変換フォルダ設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 14-B. 一時変換フォルダ取得
 * ===================================================== */

function cfPayrollGetTempFolder_() {

  const folderId =
    String(
      cfPayrollGetSetting_(
        "PAYROLL_IMPORT_TEMP_FOLDER_ID",
        "",
        true
      )
    ).trim();


  try {

    return DriveApp.getFolderById(
      folderId
    );

  } catch (error) {

    throw new Error(
      "一時変換フォルダを取得できません。" +
      " setupPayrollConversionFolder() を確認してください。"
    );

  }

}


/* =====================================================
 * 14-C. 読取可能なSpreadsheetを準備
 *
 * 戻り値：
 *
 * {
 *   originalFileId,
 *   spreadsheetId,
 *   temporary,
 *   originalFileType,
 *   originalFileName
 * }
 *
 * GOOGLE_SHEETSならそのまま使用。
 * Excel / CSVなら一時変換。
 * ===================================================== */

function cfPayrollPrepareReadableSpreadsheet_(
  fileId
) {

  if (!fileId) {

    throw new Error(
      "ファイルIDがありません。"
    );

  }


  let file;


  try {

    file =
      DriveApp.getFileById(
        fileId
      );

  } catch (error) {

    throw new Error(
      "取込ファイルを取得できません。" +
      " ファイルID：" +
      fileId
    );

  }


  const mimeType =
    file.getMimeType();


  const fileName =
    file.getName();


  const fileType =
    cfPayrollDetectFileType_(
      mimeType,
      fileName
    );


  /* =================================================
   * Google Sheets
   *
   * 変換不要
   * ================================================= */

  if (
    fileType ===
    "GOOGLE_SHEETS"
  ) {

    return {

      originalFileId:
        fileId,

      spreadsheetId:
        fileId,

      temporary:
        false,

      originalFileType:
        fileType,

      originalFileName:
        fileName

    };

  }


  /* =================================================
   * 変換対象
   * ================================================= */

  const convertibleTypes = [
    "XLSX",
    "XLS",
    "CSV"
  ];


  if (
    convertibleTypes.indexOf(
      fileType
    ) === -1
  ) {

    throw new Error(
      "Google Sheetsへ変換できないファイル形式です。" +
      " ファイル名：" +
      fileName +
      " / 種類：" +
      fileType
    );

  }


  const tempFolder =
    cfPayrollGetTempFolder_();


  const tempName =
    cfPayrollBuildTemporaryFileName_(
      fileName
    );


  /* =================================================
   * 元ファイルBlob
   * ================================================= */

  const blob =
    file.getBlob();


  /* =================================================
   * Drive APIでGoogle Sheetsへ変換
   *
   * 元Excelは変更しない。
   * 新しいGoogle Sheetsファイルを作る。
   * ================================================= */

  let createdFile;


  try {

    createdFile =
      Drive.Files.create(
        {
          name:
            tempName,

          mimeType:
            "application/vnd.google-apps.spreadsheet",

          parents: [
            tempFolder.getId()
          ]
        },

        blob,

        {
          fields:
            "id,name,mimeType",

          supportsAllDrives:
            true
        }
      );

  } catch (error) {

    throw new Error(
      "Google Sheetsへの一時変換に失敗しました。" +
      " ファイル：" +
      fileName +
      " / " +
      error.message
    );

  }


  if (
    !createdFile ||
    !createdFile.id
  ) {

    throw new Error(
      "一時変換後のファイルIDを取得できません。"
    );

  }


  /* =================================================
   * Spreadsheetとして開けるまで待機
   * ================================================= */

  cfPayrollWaitForSpreadsheetReady_(
    createdFile.id
  );


  return {

    originalFileId:
      fileId,

    spreadsheetId:
      createdFile.id,

    temporary:
      true,

    originalFileType:
      fileType,

    originalFileName:
      fileName

  };

}


/* =====================================================
 * 14-D. Spreadsheetオープン待機
 *
 * Drive変換直後は一瞬だけ
 * SpreadsheetAppから開けない場合があるため、
 * 少し待ちながら再試行する。
 * ===================================================== */

function cfPayrollWaitForSpreadsheetReady_(
  spreadsheetId
) {

  const maxAttempts =
    6;


  let lastError;


  for (
    let attempt = 1;
    attempt <= maxAttempts;
    attempt++
  ) {

    try {

      const book =
        SpreadsheetApp.openById(
          spreadsheetId
        );


      /*
       * 実際にSheetsとして認識できているか確認
       */
      book.getSheets();


      return true;

    } catch (error) {

      lastError =
        error;


      if (
        attempt <
        maxAttempts
      ) {

        Utilities.sleep(
          500
        );

      }

    }

  }


  throw new Error(
    "変換後ファイルをGoogle Sheetsとして開けません。" +
    " ID：" +
    spreadsheetId +
    " / " +
    (
      lastError
        ? lastError.message
        : ""
    )
  );

}


/* =====================================================
 * 14-E. 一時ファイル名生成
 * ===================================================== */

function cfPayrollBuildTemporaryFileName_(
  originalName
) {

  const baseName =
    String(
      originalName || "import"
    )
      .replace(
        /\.(xlsx|xls|csv)$/i,
        ""
      );


  const timestamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() ||
        "Asia/Tokyo",
      "yyyyMMdd_HHmmss_SSS"
    );


  return (
    "__CF_TEMP__" +
    timestamp +
    "_" +
    baseName
  );

}


/* =====================================================
 * 14-F. 一時変換ファイル削除
 *
 * 完全削除ではなくゴミ箱へ移動。
 * 元Excelには影響しない。
 * ===================================================== */

function cfPayrollCleanupTemporarySpreadsheet_(
  prepared
) {

  if (!prepared) {
    return;
  }


  if (
    prepared.temporary !== true
  ) {

    return;
  }


  if (
    !prepared.spreadsheetId
  ) {

    return;
  }


  try {

    DriveApp
      .getFileById(
        prepared.spreadsheetId
      )
      .setTrashed(
        true
      );

  } catch (error) {

    /*
     * 給与処理自体を失敗させない。
     * 削除失敗だけログへ残す。
     */

    console.warn(
      "一時変換ファイルのゴミ箱移動に失敗：" +
      prepared.spreadsheetId +
      " / " +
      error.message
    );

  }

}


/* =====================================================
 * 14-G. 安全なSpreadsheet読込ラッパー
 *
 * callbackの処理が成功しても失敗しても、
 * 一時変換ファイルは最後に片付ける。
 *
 * 使用例：
 *
 * cfPayrollWithReadableSpreadsheet_(
 *   fileId,
 *   function(book, prepared) {
 *     // bookを読む
 *   }
 * );
 * ===================================================== */

function cfPayrollWithReadableSpreadsheet_(
  fileId,
  callback
) {

  if (
    typeof callback !==
    "function"
  ) {

    throw new Error(
      "読込処理callbackが指定されていません。"
    );

  }


  let prepared;


  try {

    prepared =
      cfPayrollPrepareReadableSpreadsheet_(
        fileId
      );


    const book =
      SpreadsheetApp.openById(
        prepared.spreadsheetId
      );


    return callback(
      book,
      prepared
    );


  } finally {

    /*
     * Excel/CSVを変換していた場合のみ
     * 一時ファイルを片付ける
     */
    if (prepared) {

      cfPayrollCleanupTemporarySpreadsheet_(
        prepared
      );

    }

  }

}


/* =====================================================
 * 14-H. 読取テスト
 *
 * ファイル内容は変更しない。
 * 給与データにも反映しない。
 *
 * 今はまだ実行しなくてOK。
 * ===================================================== */

function testPayrollReadableFile(
  fileId
) {

  return cfPayrollWithReadableSpreadsheet_(
    fileId,

    function(
      book,
      prepared
    ) {

      const result = {

        success:
          true,

        originalFile:
          prepared.originalFileName,

        originalType:
          prepared.originalFileType,

        temporaryConversion:
          prepared.temporary,

        spreadsheetName:
          book.getName(),

        sheets:
          book
            .getSheets()
            .map(function(sheet) {

              return {

                name:
                  sheet.getName(),

                lastRow:
                  sheet.getLastRow(),

                lastColumn:
                  sheet.getLastColumn()

              };

            })

      };


      Logger.log(
        "給与ファイル読取OK: " +
        JSON.stringify(result)
      );


      return result;

    }
  );

}


/* =====================================================
 * 15. Excel対応 給与取込プレビュー統合
 *
 * Section 11
 *   取込内容解析
 *
 * Section 13
 *   ファイル自動判定
 *
 * Section 14
 *   Excel / CSV 読込
 *
 * を統合する。
 *
 * このSectionではまだ
 *
 * ・PAY_FACTSへ書き込まない
 * ・PAY_IMPORT_LOGへ書き込まない
 * ・ファイルを移動しない
 * ・給与計算しない
 *
 * 完全なプレビュー専用。
 * ===================================================== */


/* =====================================================
 * 15-A. Spreadsheet内容解析
 *
 * sourceBook：
 *   読取可能状態になったSpreadsheet
 *
 * prepared：
 *   元ファイル情報
 *
 * source：
 *   PAY_IMPORT_SOURCESの設定
 * ===================================================== */

function cfPayrollAnalyzeImportSpreadsheet_(
  sourceBook,
  prepared,
  source
) {

  /* =================================================
   * 対象シート決定
   * ================================================= */

  let sourceSheet;


  if (source.sheetName) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / ファイル：" +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (sheets.length === 0) {

      throw new Error(
        "取込ファイルにシートがありません。"
      );

    }


    /*
     * 対象シート名が空欄なら
     * 一番左のシートを使用
     */
    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow < source.startRow ||
    lastColumn < 1
  ) {

    throw new Error(
      "取込対象データがありません。" +
      " シート：" +
      sourceSheet.getName()
    );

  }


  /* =================================================
   * ヘッダー取得
   *
   * 開始行の1つ上を
   * ヘッダーとして扱う。
   * ================================================= */

  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(value) {

        return String(
          value || ""
        ).trim();

      });


  /* =================================================
   * 基本列検索
   * ================================================= */

  const dateIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.dateHeader
    );


  const castIdIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castIdHeader
    );


  const castNameIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castNameHeader
    );


  /* =================================================
   * 設定されている列が
   * 実ファイルに存在するか確認
   * ================================================= */

  if (
    source.dateHeader &&
    dateIndex === -1
  ) {

    throw new Error(
      "日付列が見つかりません：" +
      source.dateHeader
    );

  }


  if (
    source.castIdHeader &&
    castIdIndex === -1
  ) {

    throw new Error(
      "キャストID列が見つかりません：" +
      source.castIdHeader
    );

  }


  if (
    source.castNameHeader &&
    castNameIndex === -1
  ) {

    throw new Error(
      "キャスト名列が見つかりません：" +
      source.castNameHeader
    );

  }


  if (
    castIdIndex === -1 &&
    castNameIndex === -1
  ) {

    throw new Error(
      "キャストを識別できません。" +
      " PAY_IMPORT_SOURCESで" +
      "CAST_ID列またはキャスト名列を設定してください。"
    );

  }


  /* =================================================
   * PAY_IMPORT_MAP取得
   * ================================================= */

  const mappings =
    cfPayrollGetImportMappings_(
      source.dataType
    );


  const mappedColumns = [];


  headers.forEach(function(
    header,
    index
  ) {

    const normalized =
      cfPayrollNormalizeHeaderName_(
        header
      );


    const mapping =
      mappings[
        normalized
      ];


    if (!mapping) {
      return;
    }


    mappedColumns.push({

      columnIndex:
        index,

      originalHeader:
        header,

      mapping:
        mapping

    });

  });


  if (
    mappedColumns.length === 0
  ) {

    throw new Error(
      "PAY_IMPORT_MAPに一致する項目がありません。" +
      " ファイル：" +
      prepared.originalFileName
    );

  }


  /* =================================================
   * データ取得
   * ================================================= */

  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getValues();


  /* =================================================
   * キャスト解決辞書
   * ================================================= */

  const castResolver =
    cfPayrollBuildCastResolver_();


  /* =================================================
   * 解析
   * ================================================= */

  const previewRows = [];

  const unknownCasts = {};

  const itemTotals = {};

  const recognizedCastIds = {};

  let skippedEmptyRows = 0;
  let blankDateRows = 0;


  values.forEach(function(
    row,
    index
  ) {

    const sourceRowNumber =
      source.startRow +
      index;


    /* -----------------------------------------------
     * 完全空行
     * ----------------------------------------------- */

    const hasAnyValue =
      row.some(function(value) {

        return (
          value !== "" &&
          value !== null
        );

      });


    if (!hasAnyValue) {

      skippedEmptyRows++;

      return;

    }


    /* -----------------------------------------------
     * 日付
     * ----------------------------------------------- */

    let workDate = "";


    if (dateIndex !== -1) {

      const rawDate =
        row[
          dateIndex
        ];


      if (
        rawDate === "" ||
        rawDate === null
      ) {

        blankDateRows++;

      } else {

        workDate =
          cfPayrollNormalizeDate_(
            rawDate
          );

      }

    }


    /* -----------------------------------------------
     * 元キャスト情報
     * ----------------------------------------------- */

    const sourceCastId =
      castIdIndex === -1
        ? ""
        : String(
            row[
              castIdIndex
            ] || ""
          ).trim();


    const sourceCastName =
      castNameIndex === -1
        ? ""
        : String(
            row[
              castNameIndex
            ] || ""
          ).trim();


    /* -----------------------------------------------
     * 給与CAST_IDへ変換
     * ----------------------------------------------- */

    const payrollCastId =
      cfPayrollResolveCastId_(
        castResolver,
        sourceCastId,
        sourceCastName
      );


    if (payrollCastId) {

      recognizedCastIds[
        payrollCastId
      ] = true;

    } else {

      const unknownKey =
        sourceCastId +
        "::" +
        sourceCastName;


      if (
        !unknownCasts[
          unknownKey
        ]
      ) {

        unknownCasts[
          unknownKey
        ] = {

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          rows:
            []

        };

      }


      unknownCasts[
        unknownKey
      ].rows.push(
        sourceRowNumber
      );

    }


    /* -----------------------------------------------
     * 実績項目
     * ----------------------------------------------- */

    const facts = [];


    mappedColumns.forEach(function(
      column
    ) {

      const rawValue =
        row[
          column.columnIndex
        ];


      /*
       * 完全空欄は
       * プレビュー項目から除外。
       *
       * 0は有効値なので除外しない。
       */
      if (
        rawValue === "" ||
        rawValue === null
      ) {

        return;

      }


      const convertedValue =
        cfPayrollConvertImportValue_(
          rawValue,
          column.mapping
        );


      facts.push({

        itemKey:
          column.mapping.itemKey,

        displayName:
          column.mapping.displayName,

        rawValue:
          rawValue,

        value:
          convertedValue,

        unit:
          column.mapping.unit,

        multiplier:
          column.mapping.multiplier,

        sourceHeader:
          column.originalHeader

      });


      /* ---------------------------------------------
       * 数値項目合計
       * --------------------------------------------- */

      if (
        typeof convertedValue ===
        "number"
      ) {

        const itemKey =
          column.mapping.itemKey;


        if (
          !Object.prototype
            .hasOwnProperty.call(
              itemTotals,
              itemKey
            )
        ) {

          itemTotals[
            itemKey
          ] = {

            displayName:
              column.mapping.displayName,

            unit:
              column.mapping.unit,

            total:
              0

          };

        }


        itemTotals[
          itemKey
        ].total +=
          convertedValue;

      }

    });


    previewRows.push({

      sourceRow:
        sourceRowNumber,

      date:
        workDate,

      sourceCastId:
        sourceCastId,

      sourceCastName:
        sourceCastName,

      castId:
        payrollCastId,

      facts:
        facts

    });

  });


  /* =================================================
   * マッピングされなかった列
   * ================================================= */

  const mappedColumnIndexes = {};


  mappedColumns.forEach(function(
    column
  ) {

    mappedColumnIndexes[
      column.columnIndex
    ] = true;

  });


  const unmappedHeaders =
    headers.filter(function(
      header,
      index
    ) {

      if (!header) {
        return false;
      }


      /*
       * 日付・CAST_ID・名前は
       * 実績マッピング対象外でも正常
       */
      if (
        index === dateIndex ||
        index === castIdIndex ||
        index === castNameIndex
      ) {

        return false;

      }


      return (
        !mappedColumnIndexes[
          index
        ]
      );

    });


  /* =================================================
   * 結果
   * ================================================= */

  const result = {

    success:
      true,

    mode:
      "PREVIEW_ONLY",

    file: {

      id:
        prepared.originalFileId,

      name:
        prepared.originalFileName,

      type:
        prepared.originalFileType,

      temporaryConversion:
        prepared.temporary,

      sheetName:
        sourceSheet.getName()

    },

    importSource: {

      sourceId:
        source.sourceId,

      displayName:
        source.displayName,

      dataType:
        source.dataType

    },

    summary: {

      sourceRows:
        values.length,

      recognizedRows:
        previewRows.length,

      skippedEmptyRows:
        skippedEmptyRows,

      blankDateRows:
        blankDateRows,

      mappedColumns:
        mappedColumns.length,

      unmappedColumns:
        unmappedHeaders.length,

      recognizedCasts:
        Object.keys(
          recognizedCastIds
        ).length,

      unknownCastCount:
        Object.keys(
          unknownCasts
        ).length

    },

    mappedColumns:
      mappedColumns.map(function(
        column
      ) {

        return {

          sourceHeader:
            column.originalHeader,

          itemKey:
            column.mapping.itemKey,

          displayName:
            column.mapping.displayName,

          unit:
            column.mapping.unit,

          multiplier:
            column.mapping.multiplier

        };

      }),

    unmappedHeaders:
      unmappedHeaders,

    itemTotals:
      itemTotals,

    unknownCasts:
      Object.keys(
        unknownCasts
      ).map(function(key) {

        return unknownCasts[
          key
        ];

      }),

    rows:
      previewRows

  };


  Logger.log(
    "給与取込プレビューOK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 15-B. 取込待ちフォルダを一括プレビュー
 *
 * 01_取込待ち内の全ファイルについて
 *
 * 1. ファイル種類確認
 * 2. 取込設定自動判定
 * 3. Excel等を必要に応じ一時変換
 * 4. 内容解析
 *
 * まで行う。
 *
 * 書込・移動は一切しない。
 * ===================================================== */

function previewPayrollInboxImports() {

  const inboxFolder =
    cfPayrollGetInboxFolder_();


  const sources =
    cfPayrollGetImportSources_();


  const files =
    inboxFolder.getFiles();


  const results = [];


  const summary = {

    totalFiles:
      0,

    previewSuccess:
      0,

    unmatched:
      0,

    ambiguous:
      0,

    unsupported:
      0,

    previewError:
      0

  };


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    summary.totalFiles++;


    const fileInfo =
      cfPayrollBuildImportFileInfo_(
        file
      );


    const detection =
      cfPayrollDetectImportSourceForFile_(
        fileInfo,
        sources
      );


    /* =================================================
     * 自動判定できなかったファイル
     * ================================================= */

    if (
      detection.status !==
      "MATCHED"
    ) {

      if (
        detection.status ===
        "UNMATCHED"
      ) {

        summary.unmatched++;

      } else if (
        detection.status ===
        "AMBIGUOUS"
      ) {

        summary.ambiguous++;

      } else if (
        detection.status ===
        "UNSUPPORTED"
      ) {

        summary.unsupported++;

      }


      results.push({

        file:
          fileInfo,

        detection:
          detection,

        preview:
          null

      });


      continue;

    }


    /* =================================================
     * 中身プレビュー
     * ================================================= */

    try {

      const preview =
        cfPayrollPreviewImport_(
          fileInfo.id,
          detection.sourceId
        );


      summary.previewSuccess++;


      results.push({

        file:
          fileInfo,

        detection:
          detection,

        preview:
          preview

      });


    } catch (error) {

      summary.previewError++;


      results.push({

        file:
          fileInfo,

        detection:
          detection,

        preview: {

          success:
            false,

          error:
            error.message

        }

      });

    }

  }


  /* =================================================
   * ファイル名順
   * ================================================= */

  results.sort(function(
    a,
    b
  ) {

    return a.file.name.localeCompare(
      b.file.name,
      "ja"
    );

  });


  const result = {

    success:
      true,

    mode:
      "INBOX_PREVIEW_ONLY",

    summary:
      summary,

    files:
      results

  };


  Logger.log(
    "給与取込フォルダ一括プレビュー: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 16. 給与取込 二重取込防止エンジン
 *
 * ・取込ID発行
 * ・SHA-256内容ハッシュ生成
 * ・ファイルID重複確認
 * ・内容重複確認
 * ・PAY_IMPORT_LOG登録
 * ・成功 / エラー状態更新
 *
 * 同名・別名に関係なく、
 * 同じ内容を二重反映しない。
 * ===================================================== */


/* =====================================================
 * 16-A. PAY_IMPORT_LOG初期確認・表示設定
 * ===================================================== */

function setupPayrollImportProtection() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_LOG"
    );


  if (!sheet) {

    throw new Error(
      "PAY_IMPORT_LOGがありません。"
    );

  }


  const headers = [
    "取込ID",
    "ファイルID",
    "ファイル名",
    "内容ハッシュ",
    "データ種別",
    "対象開始日",
    "対象終了日",
    "処理状態",
    "読込行数",
    "反映行数",
    "処理日時",
    "エラー内容"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      headers
    )
  ) {

    throw new Error(
      "PAY_IMPORT_LOGの列構成が異なります。"
    );

  }


  /*
   * シートに現在存在する行だけ書式設定。
   *
   * "F:G" のような列全体指定は
   * 非常に大きな範囲になるため使用しない。
   */
  const dataRowCount =
    Math.max(
      sheet.getMaxRows() - 1,
      1
    );


  /* 対象開始日・終了日 */
  sheet
    .getRange(
      2,
      6,
      dataRowCount,
      2
    )
    .setNumberFormat(
      "yyyy/mm/dd"
    );


  /* 読込行数・反映行数 */
  sheet
    .getRange(
      2,
      9,
      dataRowCount,
      2
    )
    .setNumberFormat(
      "0"
    );


  /* 処理日時 */
  sheet
    .getRange(
      2,
      11,
      dataRowCount,
      1
    )
    .setNumberFormat(
      "yyyy/mm/dd hh:mm:ss"
    );


  SpreadsheetApp.flush();


  const result = {

    success:
      true,

    formattedRows:
      dataRowCount

  };


  Logger.log(
    "給与取込保護設定OK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 17. PAY_FACTS 本取込エンジン
 *
 * プレビュー済みの内容をPAY_FACTSへ保存する。
 *
 * 安全設計：
 * ・二重取込禁止
 * ・未登録キャストがいれば停止
 * ・日付欠落を検出
 * ・書込途中で失敗した場合はロールバック
 * ・成功後のみ取込完了扱い
 *
 * このSection単体では自動実行しない。
 * ===================================================== */


/* =====================================================
 * 17-A. 本取込設定
 *
 * PAY_SETTINGSへ安全設定を追加。
 * 既存値は上書きしない。
 * ===================================================== */

function setupPayrollImportExecution() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    cfPayrollSaveSystemSetting_(
      "IMPORT_BLOCK_UNKNOWN_CASTS",
      "未登録キャスト時に取込停止",
      true,
      "true = PAY_CAST_MAPで判別できないキャストが1人でもいれば本取込しない"
    );


    cfPayrollSaveSystemSetting_(
      "IMPORT_BLOCK_BLANK_DATES",
      "日付空欄時に取込停止",
      true,
      "true = 日付列が設定されているデータで日付空欄があれば本取込しない"
    );


    cfPayrollSaveSystemSetting_(
      "IMPORT_STORE_ZERO_VALUES",
      "0件・0円実績を保存",
      false,
      "false = 値が0の実績はPAY_FACTSへ保存しない"
    );


    SpreadsheetApp.flush();


    const result = {
      success: true
    };


    Logger.log(
      "給与本取込設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 17-B. 1ファイル本取込
 *
 * 内部処理用。
 *
 * fileId：
 *   元ファイルID
 *
 * sourceId：
 *   SRC_ORDER_SALES等
 *
 * 安全処理：
 * ・二重取込防止
 * ・プレビュー解析
 * ・未登録キャスト確認
 * ・未登録商品確認
 * ・PAY_FACTS書込
 * ・エラー時ロールバック
 * ===================================================== */

function cfPayrollImportFile_(
  fileId,
  sourceId
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  let importSession =
    null;


  let factsWritten =
    false;


  try {

    /* =================================================
     * 取込設定
     * ================================================= */

    const source =
      cfPayrollGetImportSourceById_(
        sourceId
      );


    if (!source) {

      throw new Error(
        "有効な取込設定が見つかりません：" +
        sourceId
      );

    }


    /* =================================================
     * 二重取込確認 + 取込ログ開始
     * ================================================= */

    importSession =
      cfPayrollBeginImport_(
        fileId,
        source.dataType
      );


    if (
      importSession.duplicate ===
      true
    ) {

      return {

        success:
          true,

        imported:
          false,

        duplicate:
          true,

        reason:
          importSession.reason,

        previousImportId:
          importSession.previousImportId,

        fileName:
          importSession.fileName

      };

    }


    /* =================================================
     * 内容プレビュー
     *
     * Section 11-A →
     * Section 21等の解析方式へ振り分け。
     * ================================================= */

    const preview =
      cfPayrollPreviewImport_(
        fileId,
        sourceId
      );


    /* =================================================
     * 安全チェック V2
     *
     * ・未登録キャスト
     * ・日付
     * ・未登録商品
     * ================================================= */

    cfPayrollValidateImportPreviewV2_(
      preview,
      source
    );


    /* =================================================
     * PAY_FACTS用行生成
     * ================================================= */

    const factRows =
      cfPayrollBuildFactRows_(
        preview,
        importSession
      );


    /* =================================================
     * PAY_FACTS書込
     * ================================================= */

    if (
      factRows.length > 0
    ) {

      cfPayrollAppendFactRows_(
        factRows
      );


      factsWritten =
        true;

    }


    /* =================================================
     * 対象期間
     * ================================================= */

    const period =
      cfPayrollGetPreviewPeriod_(
        preview
      );


    /* =================================================
     * 取込完了
     * ================================================= */

    cfPayrollCompleteImport_(
      importSession.importId,
      {

        startDate:
          period.startDate,

        endDate:
          period.endDate,

        readRows:
          preview.summary
            .recognizedRows,

        reflectedRows:
          factRows.length

      }
    );


    /* =================================================
     * 元ファイルを処理済みへ
     * ================================================= */

    try {

      cfPayrollMoveOriginalFile_(
        fileId,
        cfPayrollGetProcessedFolder_()
      );

    } catch (moveError) {

      console.warn(
        "給与取込は成功しましたが、" +
        "処理済みフォルダへの移動に失敗しました：" +
        moveError.message
      );

    }


    const result = {

      success:
        true,

      imported:
        true,

      duplicate:
        false,

      importId:
        importSession.importId,

      fileName:
        importSession.fileName,

      dataType:
        source.dataType,

      parserMode:
        source.parserMode,

      readRows:
        preview.summary
          .recognizedRows,

      factRows:
        factRows.length,

      startDate:
        period.startDate,

      endDate:
        period.endDate

    };


    Logger.log(
      "給与本取込OK: " +
      JSON.stringify(result)
    );


    return result;


  } catch (error) {

    /* =================================================
     * PAY_FACTSロールバック
     * ================================================= */

    if (
      factsWritten &&
      importSession &&
      importSession.importId
    ) {

      try {

        cfPayrollRollbackFactsByImportId_(
          importSession.importId
        );

      } catch (
        rollbackError
      ) {

        console.error(
          "PAY_FACTSロールバック失敗：" +
          rollbackError.message
        );

      }

    }


    /* =================================================
     * 取込ログをエラーへ
     * ================================================= */

    if (
      importSession &&
      importSession.importId
    ) {

      try {

        cfPayrollFailImport_(
          importSession.importId,
          error
        );

      } catch (
        logError
      ) {

        console.error(
          "取込エラーログ更新失敗：" +
          logError.message
        );

      }

    }


    /* =================================================
     * 元ファイルをエラーフォルダへ
     * ================================================= */

    if (
      importSession &&
      importSession.importId
    ) {

      try {

        cfPayrollMoveOriginalFile_(
          fileId,
          cfPayrollGetErrorFolder_()
        );

      } catch (
        moveError
      ) {

        console.warn(
          "エラーフォルダへの移動失敗：" +
          moveError.message
        );

      }

    }


    throw error;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 17-C. プレビュー安全チェック
 *
 * ・未登録キャスト
 * ・営業日未確定
 * ・取込可能データなし
 *
 * 営業日は
 * dateHeaderの有無ではなく
 * dateMethodで判定する。
 * ===================================================== */

function cfPayrollValidateImportPreview_(
  preview,
  source
) {

  if (
    !preview ||
    preview.success !== true
  ) {

    throw new Error(
      "取込プレビューに失敗しています。"
    );

  }


  const blockUnknown =
    cfPayrollToBoolean_(
      cfPayrollGetSetting_(
        "IMPORT_BLOCK_UNKNOWN_CASTS",
        true,
        false
      )
    );


  const blockBlankDates =
    cfPayrollToBoolean_(
      cfPayrollGetSetting_(
        "IMPORT_BLOCK_BLANK_DATES",
        true,
        false
      )
    );


  /* =================================================
   * 未登録キャスト
   * ================================================= */

  const unknownCastCount =
    Number(
      preview.summary &&
      preview.summary.unknownCastCount
        ? preview.summary.unknownCastCount
        : 0
    );


  if (
    blockUnknown &&
    unknownCastCount > 0
  ) {

    const names =
      (
        preview.unknownCasts ||
        []
      )
        .map(function(
          item
        ) {

          return (
            item.sourceCastName ||
            item.sourceCastId ||
            "名称不明"
          );

        })
        .join(", ");


    throw new Error(
      "未登録キャストが存在するため取込を停止しました：" +
      names
    );

  }


  /* =================================================
   * 営業日
   *
   * FILENAME / COLUMN 等、
   * 日付を使用する取込は
   * 営業日未確定なら本取込禁止。
   *
   * PAY_IMPORT_QUEUEで確定してから取込する。
   * ================================================= */

  const dateMethod =
    String(
      source.dateMethod ||
      "NONE"
    )
      .trim()
      .toUpperCase();


  const blankDateRows =
    Number(
      preview.summary &&
      preview.summary.blankDateRows
        ? preview.summary.blankDateRows
        : 0
    );


  if (
    blockBlankDates &&
    dateMethod !==
      "NONE" &&
    blankDateRows > 0
  ) {

    throw new Error(
      "営業日が未確定の行が " +
      blankDateRows +
      " 件あるため取込を停止しました。" +
      "PAY_IMPORT_QUEUEで営業日を確定してください。"
    );

  }


  /* =================================================
   * データなし
   * ================================================= */

  const recognizedRows =
    Number(
      preview.summary &&
      preview.summary.recognizedRows
        ? preview.summary.recognizedRows
        : 0
    );


  if (
    recognizedRows <= 0
  ) {

    throw new Error(
      "取込可能なデータ行がありません。"
    );

  }


  return true;

}


/* =====================================================
 * 17-D. PAY_FACTS用データ生成
 * ===================================================== */

function cfPayrollBuildFactRows_(
  preview,
  importSession
) {

  const rows = [];


  const storeZeroValues =
    cfPayrollToBoolean_(
      cfPayrollGetSetting_(
        "IMPORT_STORE_ZERO_VALUES",
        false,
        false
      )
    );


  const now =
    new Date();


  let sequence = 0;


  preview.rows.forEach(function(
    sourceRow
  ) {

    sourceRow.facts.forEach(function(
      fact
    ) {

      /* ---------------------------------------------
       * 0を保存しない設定
       * --------------------------------------------- */

      if (
        !storeZeroValues &&
        typeof fact.value ===
          "number" &&
        fact.value === 0
      ) {

        return;

      }


      sequence++;


      const factId =
        cfPayrollCreateFactId_(
          importSession.importId,
          sourceRow.sourceRow,
          sequence
        );


      rows.push([

        /* A 実績ID */
        factId,

        /* B 取込ID */
        importSession.importId,

        /* C 日付 */
        sourceRow.date || "",

        /* D 元システム */
        "IMPORT:" +
          preview.importSource.sourceId,

        /* E 元キャストID */
        sourceRow.sourceCastId || "",

        /* F CAST_ID */
        sourceRow.castId || "",

        /* G 項目キー */
        fact.itemKey,

        /* H 表示名 */
        fact.displayName,

        /* I 値 */
        fact.value,

        /* J 単位 */
        fact.unit,

        /* K 元ファイルID */
        importSession.fileId,

        /* L 元ファイル名 */
        importSession.fileName,

        /* M 元行番号 */
        sourceRow.sourceRow,

        /* N 内容ハッシュ */
        importSession.contentHash,

        /* O 登録日時 */
        now

      ]);

    });

  });


  return rows;

}


/* =====================================================
 * 17-E. 実績ID発行
 * ===================================================== */

function cfPayrollCreateFactId_(
  importId,
  sourceRow,
  sequence
) {

  return (
    "FACT_" +
    importId +
    "_" +
    sourceRow +
    "_" +
    sequence
  );

}


/* =====================================================
 * 17-F. PAY_FACTSへ一括追加
 * ===================================================== */

function cfPayrollAppendFactRows_(
  rows
) {

  if (
    !rows ||
    rows.length === 0
  ) {

    return 0;

  }


  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_FACTS"
    );


  if (!sheet) {

    throw new Error(
      "PAY_FACTSがありません。"
    );

  }


  const headers = [
    "実績ID",
    "取込ID",
    "日付",
    "元システム",
    "元キャストID",
    "CAST_ID",
    "項目キー",
    "表示名",
    "値",
    "単位",
    "元ファイルID",
    "元ファイル名",
    "元行番号",
    "内容ハッシュ",
    "登録日時"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      headers
    )
  ) {

    throw new Error(
      "PAY_FACTSの列構成が異なります。"
    );

  }


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


  SpreadsheetApp.flush();


  return rows.length;

}


/* =====================================================
 * 17-G. 取込ID単位ロールバック
 *
 * エラー時のみ使用。
 *
 * PAY_FACTSの中から
 * 指定取込IDの行だけ削除する。
 * ===================================================== */

function cfPayrollRollbackFactsByImportId_(
  importId
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_FACTS"
    );


  if (!sheet) {
    return 0;
  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1
  ) {
    return 0;
  }


  const importIds =
    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  let deleted =
    0;


  /*
   * 下から削除。
   * 行番号ズレ防止。
   */
  for (
    let i =
      importIds.length - 1;

    i >= 0;

    i--
  ) {

    if (
      String(
        importIds[i][0] || ""
      ).trim() !==
      importId
    ) {

      continue;

    }


    sheet.deleteRow(
      i + 2
    );


    deleted++;

  }


  SpreadsheetApp.flush();


  Logger.log(
    "PAY_FACTSロールバック: " +
    importId +
    " / " +
    deleted +
    "行"
  );


  return deleted;

}


/* =====================================================
 * 17-H. プレビュー内の対象期間取得
 * ===================================================== */

function cfPayrollGetPreviewPeriod_(
  preview
) {

  const dates = [];


  preview.rows.forEach(function(row) {

    if (
      row.date instanceof Date &&
      !Number.isNaN(
        row.date.getTime()
      )
    ) {

      dates.push(
        row.date
      );

    }

  });


  if (
    dates.length === 0
  ) {

    return {

      startDate:
        "",

      endDate:
        ""

    };

  }


  dates.sort(function(
    a,
    b
  ) {

    return (
      a.getTime() -
      b.getTime()
    );

  });


  return {

    startDate:
      dates[0],

    endDate:
      dates[
        dates.length - 1
      ]

  };

}


/* =====================================================
 * 17-I. ファイル移動
 * ===================================================== */

function cfPayrollMoveOriginalFile_(
  fileId,
  destinationFolder
) {

  if (!destinationFolder) {

    throw new Error(
      "移動先フォルダがありません。"
    );

  }


  const file =
    DriveApp.getFileById(
      fileId
    );


  file.moveTo(
    destinationFolder
  );


  return true;

}


/* =====================================================
 * 17-J. Boolean変換
 *
 * Sheetsで
 * TRUE / FALSE
 * true / false
 * 1 / 0
 * 等になっても扱えるようにする。
 * ===================================================== */

function cfPayrollToBoolean_(
  value
) {

  if (
    value === true ||
    value === 1
  ) {

    return true;

  }


  if (
    value === false ||
    value === 0 ||
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return false;

  }


  const text =
    String(
      value
    )
      .trim()
      .toLowerCase();


  return [
    "true",
    "1",
    "yes",
    "on",
    "有効",
    "はい",
    "○",
    "〇"
  ].indexOf(text) !== -1;

}

/* =====================================================
 * 16-B. 取込ID発行
 *
 * 例：
 * IMP_20260914_234501_ab12cd34
 * ===================================================== */

function cfPayrollCreateImportId_() {

  const timezone =
    String(
      cfPayrollGetSetting_(
        "TIMEZONE",
        "Asia/Tokyo",
        false
      )
    );


  const timestamp =
    Utilities.formatDate(
      new Date(),
      timezone,
      "yyyyMMdd_HHmmss"
    );


  const uuid =
    Utilities
      .getUuid()
      .replace(/-/g, "")
      .substring(0, 8);


  return (
    "IMP_" +
    timestamp +
    "_" +
    uuid
  );

}


/* =====================================================
 * 16-C. ファイル内容ハッシュ生成
 *
 * SHA-256
 *
 * Excel / CSV：
 *   ファイル本体のバイト列
 *
 * Google Sheets：
 *   シート名・値・数式を基に生成
 * ===================================================== */

function cfPayrollComputeImportHash_(
  fileId
) {

  const file =
    DriveApp.getFileById(
      fileId
    );


  const fileType =
    cfPayrollDetectFileType_(
      file.getMimeType(),
      file.getName()
    );


  /* =================================================
   * Google Sheets
   * ================================================= */

  if (
    fileType ===
    "GOOGLE_SHEETS"
  ) {

    return cfPayrollHashGoogleSpreadsheet_(
      fileId
    );

  }


  /* =================================================
   * Excel / CSV
   * ================================================= */

  if (
    [
      "XLSX",
      "XLS",
      "CSV"
    ].indexOf(
      fileType
    ) !== -1
  ) {

    const bytes =
      file
        .getBlob()
        .getBytes();


    const digest =
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        bytes
      );


    return cfPayrollDigestToHex_(
      digest
    );

  }


  throw new Error(
    "内容ハッシュを作成できない形式です：" +
    file.getName()
  );

}


/* =====================================================
 * 16-D. Google Sheets内容ハッシュ
 * ===================================================== */

function cfPayrollHashGoogleSpreadsheet_(
  spreadsheetId
) {

  const book =
    SpreadsheetApp.openById(
      spreadsheetId
    );


  const parts = [];


  book
    .getSheets()
    .forEach(function(sheet) {

      const lastRow =
        sheet.getLastRow();


      const lastColumn =
        sheet.getLastColumn();


      parts.push(
        "SHEET:" +
        sheet.getName()
      );


      parts.push(
        "SIZE:" +
        lastRow +
        "x" +
        lastColumn
      );


      if (
        lastRow === 0 ||
        lastColumn === 0
      ) {
        return;
      }


      const range =
        sheet.getRange(
          1,
          1,
          lastRow,
          lastColumn
        );


      const values =
        range
          .getValues()
          .map(function(row) {

            return row.map(
              cfPayrollNormalizeHashValue_
            );

          });


      const formulas =
        range.getFormulas();


      parts.push(
        JSON.stringify(
          values
        )
      );


      parts.push(
        JSON.stringify(
          formulas
        )
      );

    });


  const text =
    parts.join(
      "\n"
    );


  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      text,
      Utilities.Charset.UTF_8
    );


  return cfPayrollDigestToHex_(
    digest
  );

}


/* =====================================================
 * 16-E. ハッシュ用値正規化
 * ===================================================== */

function cfPayrollNormalizeHashValue_(
  value
) {

  if (
    value instanceof Date
  ) {

    return (
      "DATE:" +
      value.getTime()
    );

  }


  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return value;

}


/* =====================================================
 * 16-F. SHA-256結果を16進数へ変換
 * ===================================================== */

function cfPayrollDigestToHex_(
  digest
) {

  return digest
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;


      return (
        "0" +
        value.toString(16)
      ).slice(-2);

    })
    .join("");

}


/* =====================================================
 * 16-G. 既存の取込完了データを検索
 *
 * ファイルID または 内容ハッシュ
 * が一致した場合は重複扱い。
 *
 * エラー履歴は再取込可能。
 * ===================================================== */

function cfPayrollFindDuplicateImport_(
  fileId,
  contentHash
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_LOG"
    );


  if (!sheet) {

    throw new Error(
      "PAY_IMPORT_LOGがありません。"
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {
    return null;
  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        12
      )
      .getValues();


  for (
    let i = 0;
    i < rows.length;
    i++
  ) {

    const row =
      rows[i];


    const existingImportId =
      String(
        row[0] || ""
      ).trim();


    const existingFileId =
      String(
        row[1] || ""
      ).trim();


    const existingHash =
      String(
        row[3] || ""
      ).trim();


    const status =
      String(
        row[7] || ""
      ).trim();


    /*
     * 正常に取り込み済みのものだけ
     * 二重取込防止対象。
     */
    if (
      status !==
      "取込完了"
    ) {
      continue;
    }


    if (
      fileId &&
      existingFileId === fileId
    ) {

      return {

        duplicate:
          true,

        reason:
          "FILE_ID",

        importId:
          existingImportId,

        sheetRow:
          i + 2

      };

    }


    if (
      contentHash &&
      existingHash === contentHash
    ) {

      return {

        duplicate:
          true,

        reason:
          "CONTENT_HASH",

        importId:
          existingImportId,

        sheetRow:
          i + 2

      };

    }

  }


  return null;

}


/* =====================================================
 * 16-H. 取込開始予約
 *
 * Section 17の本取込開始時に使用。
 *
 * 戻り値：
 *
 * duplicate:true
 *   → 取込禁止
 *
 * duplicate:false
 *   → importIdを使って処理開始
 * ===================================================== */

function cfPayrollBeginImport_(
  fileId,
  dataType
) {

  const file =
    DriveApp.getFileById(
      fileId
    );


  const contentHash =
    cfPayrollComputeImportHash_(
      fileId
    );


  const duplicate =
    cfPayrollFindDuplicateImport_(
      fileId,
      contentHash
    );


  if (duplicate) {

    return {

      duplicate:
        true,

      reason:
        duplicate.reason,

      previousImportId:
        duplicate.importId,

      fileId:
        fileId,

      fileName:
        file.getName(),

      contentHash:
        contentHash

    };

  }


  const importId =
    cfPayrollCreateImportId_();


  const book =
    cfPayrollOpenSpreadsheet_();


  const logSheet =
    book.getSheetByName(
      "PAY_IMPORT_LOG"
    );


  logSheet
    .getRange(
      logSheet.getLastRow() + 1,
      1,
      1,
      12
    )
    .setValues([[
      importId,
      fileId,
      file.getName(),
      contentHash,
      String(
        dataType || ""
      ).toUpperCase(),
      "",
      "",
      "処理中",
      0,
      0,
      new Date(),
      ""
    ]]);


  SpreadsheetApp.flush();


  return {

    duplicate:
      false,

    importId:
      importId,

    fileId:
      fileId,

    fileName:
      file.getName(),

    contentHash:
      contentHash

  };

}


/* =====================================================
 * 16-I. 取込成功
 * ===================================================== */

function cfPayrollCompleteImport_(
  importId,
  options
) {

  options =
    options || {};


  cfPayrollUpdateImportLog_(
    importId,
    {

      startDate:
        options.startDate || "",

      endDate:
        options.endDate || "",

      status:
        "取込完了",

      readRows:
        Number(
          options.readRows || 0
        ),

      reflectedRows:
        Number(
          options.reflectedRows || 0
        ),

      error:
        ""

    }
  );

}


/* =====================================================
 * 16-J. 取込失敗
 * ===================================================== */

function cfPayrollFailImport_(
  importId,
  error
) {

  let message = "";


  if (
    error instanceof Error
  ) {

    message =
      error.message;

  } else {

    message =
      String(
        error || ""
      );

  }


  cfPayrollUpdateImportLog_(
    importId,
    {

      status:
        "エラー",

      error:
        message

    }
  );

}


/* =====================================================
 * 16-K. PAY_IMPORT_LOG更新
 * ===================================================== */

function cfPayrollUpdateImportLog_(
  importId,
  options
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_LOG"
    );


  if (!sheet) {

    throw new Error(
      "PAY_IMPORT_LOGがありません。"
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {

    throw new Error(
      "取込IDが見つかりません：" +
      importId
    );

  }


  const importIds =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  let targetRow =
    -1;


  for (
    let i = 0;
    i < importIds.length;
    i++
  ) {

    if (
      String(
        importIds[i][0] || ""
      ).trim() === importId
    ) {

      targetRow =
        i + 2;

      break;

    }

  }


  if (
    targetRow === -1
  ) {

    throw new Error(
      "取込IDが見つかりません：" +
      importId
    );

  }


  const current =
    sheet
      .getRange(
        targetRow,
        1,
        1,
        12
      )
      .getValues()[0];


  /*
   * F：対象開始日
   * G：対象終了日
   * H：処理状態
   * I：読込行数
   * J：反映行数
   * K：処理日時
   * L：エラー内容
   */

  if (
    Object.prototype
      .hasOwnProperty.call(
        options,
        "startDate"
      )
  ) {

    current[5] =
      options.startDate;

  }


  if (
    Object.prototype
      .hasOwnProperty.call(
        options,
        "endDate"
      )
  ) {

    current[6] =
      options.endDate;

  }


  if (
    Object.prototype
      .hasOwnProperty.call(
        options,
        "status"
      )
  ) {

    current[7] =
      options.status;

  }


  if (
    Object.prototype
      .hasOwnProperty.call(
        options,
        "readRows"
      )
  ) {

    current[8] =
      options.readRows;

  }


  if (
    Object.prototype
      .hasOwnProperty.call(
        options,
        "reflectedRows"
      )
  ) {

    current[9] =
      options.reflectedRows;

  }


  current[10] =
    new Date();


  if (
    Object.prototype
      .hasOwnProperty.call(
        options,
        "error"
      )
  ) {

    current[11] =
      options.error;

  }


  sheet
    .getRange(
      targetRow,
      1,
      1,
      12
    )
    .setValues([
      current
    ]);


  SpreadsheetApp.flush();

}


/* =====================================================
 * 16-L. 重複判定プレビュー
 *
 * PAY_IMPORT_LOGには書き込まない。
 * 後でテスト用に使用。
 * ===================================================== */

function previewPayrollDuplicateCheck(
  fileId
) {

  const file =
    DriveApp.getFileById(
      fileId
    );


  const hash =
    cfPayrollComputeImportHash_(
      fileId
    );


  const duplicate =
    cfPayrollFindDuplicateImport_(
      fileId,
      hash
    );


  const result = {

    success:
      true,

    fileId:
      fileId,

    fileName:
      file.getName(),

    contentHash:
      hash,

    duplicate:
      duplicate !== null,

    duplicateInfo:
      duplicate

  };


  Logger.log(
    "給与重複チェック: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 18. 給与取込待ち 一括本取込
 *
 * 01_取込待ち 内のファイルを
 *
 * 1. ファイル形式判定
 * 2. 取込設定自動判定
 * 3. 内容解析
 * 4. 二重取込確認
 * 5. PAY_FACTS本登録
 * 6. 処理済み / エラーへ移動
 *
 * の順に処理する。
 *
 * 実行時間超過防止のため、
 * 1回あたりの最大処理数を設定可能。
 * ===================================================== */


/* =====================================================
 * 18-A. 一括取込設定
 * ===================================================== */

function setupPayrollInboxRunner() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    cfPayrollSaveSystemSetting_(
      "IMPORT_MAX_FILES_PER_RUN",
      "1回の最大取込ファイル数",
      5,
      "Apps Scriptの実行時間超過防止。通常は5程度を推奨"
    );


    cfPayrollSaveSystemSetting_(
      "IMPORT_MOVE_DUPLICATES_TO_PROCESSED",
      "重複ファイルを処理済みへ移動",
      true,
      "true = 既に取り込み済みの同一ファイルを02_処理済みへ移動"
    );


    cfPayrollSaveSystemSetting_(
      "IMPORT_MOVE_REJECTED_TO_ERROR",
      "判定不能ファイルをエラーへ移動",
      true,
      "true = 未対応・判定不能・複数一致ファイルを03_エラーへ移動"
    );


    SpreadsheetApp.flush();


    const result = {

      success:
        true

    };


    Logger.log(
      "給与一括取込設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 18-B. 取込待ち一括本処理
 *
 * ★ 本番入口
 *
 * 将来的には時間トリガーや
 * 管理画面のボタンからこの関数を呼ぶ。
 * ===================================================== */

function runPayrollInboxImport() {

  const inboxFolder =
    cfPayrollGetInboxFolder_();


  const processedFolder =
    cfPayrollGetProcessedFolder_();


  const errorFolder =
    cfPayrollGetErrorFolder_();


  const sources =
    cfPayrollGetImportSources_();


  const maxFiles =
    cfPayrollGetPositiveIntegerSetting_(
      "IMPORT_MAX_FILES_PER_RUN",
      5
    );


  const moveDuplicates =
    cfPayrollToBoolean_(
      cfPayrollGetSetting_(
        "IMPORT_MOVE_DUPLICATES_TO_PROCESSED",
        true,
        false
      )
    );


  const moveRejected =
    cfPayrollToBoolean_(
      cfPayrollGetSetting_(
        "IMPORT_MOVE_REJECTED_TO_ERROR",
        true,
        false
      )
    );


  /* =================================================
   * 取込開始時点のファイル一覧を固定
   *
   * 処理中にファイルを移動しても
   * Iteratorが崩れないようにする。
   * ================================================= */

  const iterator =
    inboxFolder.getFiles();


  const queue = [];


  while (
    iterator.hasNext()
  ) {

    const file =
      iterator.next();


    queue.push({

      id:
        file.getId(),

      name:
        file.getName(),

      lastUpdated:
        file.getLastUpdated()

    });

  }


  /* =================================================
   * ファイル名順で処理
   * ================================================= */

  queue.sort(function(
    a,
    b
  ) {

    return a.name.localeCompare(
      b.name,
      "ja"
    );

  });


  /*
   * 1回の最大件数だけ処理
   */
  const targets =
    queue.slice(
      0,
      maxFiles
    );


  const results = [];


  const summary = {

    queued:
      queue.length,

    attempted:
      targets.length,

    imported:
      0,

    duplicate:
      0,

    rejected:
      0,

    error:
      0,

    remaining:
      Math.max(
        queue.length -
        targets.length,
        0
      )

  };


  /* =================================================
   * 1ファイルずつ処理
   * ================================================= */

  targets.forEach(function(
    queuedFile
  ) {

    let file;


    try {

      file =
        DriveApp.getFileById(
          queuedFile.id
        );

    } catch (error) {

      summary.error++;


      results.push({

        fileName:
          queuedFile.name,

        fileId:
          queuedFile.id,

        status:
          "ERROR",

        message:
          "ファイルを取得できません：" +
          error.message

      });


      return;

    }


    const fileInfo =
      cfPayrollBuildImportFileInfo_(
        file
      );


    /* =================================================
     * 取込設定自動判定
     * ================================================= */

    const detection =
      cfPayrollDetectImportSourceForFile_(
        fileInfo,
        sources
      );


    /* =================================================
     * 判定不能
     * ================================================= */

    if (
      detection.status !==
      "MATCHED"
    ) {

      summary.rejected++;


      const rejectMessage =
        cfPayrollBuildDetectionErrorMessage_(
          detection
        );


      /* ---------------------------------------------
       * ログ保存
       * --------------------------------------------- */

      try {

        cfPayrollLogRejectedImport_(
          fileInfo,
          detection,
          rejectMessage
        );

      } catch (logError) {

        console.warn(
          "判定エラーログ保存失敗：" +
          logError.message
        );

      }


      /* ---------------------------------------------
       * エラーフォルダへ
       * --------------------------------------------- */

      if (moveRejected) {

        try {

          cfPayrollMoveFileIfNeeded_(
            fileInfo.id,
            errorFolder
          );

        } catch (moveError) {

          console.warn(
            "判定不能ファイルの移動失敗：" +
            moveError.message
          );

        }

      }


      results.push({

        fileName:
          fileInfo.name,

        fileId:
          fileInfo.id,

        status:
          detection.status,

        message:
          rejectMessage

      });


      return;

    }


    /* =================================================
     * 本取込
     * ================================================= */

    try {

      const importResult =
        cfPayrollImportFile_(
          fileInfo.id,
          detection.sourceId
        );


      /* ---------------------------------------------
       * 重複
       * --------------------------------------------- */

      if (
        importResult.duplicate ===
        true
      ) {

        summary.duplicate++;


        if (moveDuplicates) {

          try {

            cfPayrollMoveFileIfNeeded_(
              fileInfo.id,
              processedFolder
            );

          } catch (moveError) {

            console.warn(
              "重複ファイルの移動失敗：" +
              moveError.message
            );

          }

        }


        results.push({

          fileName:
            fileInfo.name,

          fileId:
            fileInfo.id,

          status:
            "DUPLICATE",

          previousImportId:
            importResult.previousImportId,

          reason:
            importResult.reason

        });


        return;

      }


      /* ---------------------------------------------
       * 正常取込
       * --------------------------------------------- */

      summary.imported++;


      results.push({

        fileName:
          fileInfo.name,

        fileId:
          fileInfo.id,

        status:
          "IMPORTED",

        importId:
          importResult.importId,

        dataType:
          importResult.dataType,

        readRows:
          importResult.readRows,

        factRows:
          importResult.factRows

      });


    } catch (error) {

      summary.error++;


      /*
       * cfPayrollImportFile_ 側でも
       * エラーフォルダ移動を試しているが、
       * その前段で落ちた場合に備えて再確認。
       */
      try {

        cfPayrollMoveFileIfNeeded_(
          fileInfo.id,
          errorFolder
        );

      } catch (moveError) {

        console.warn(
          "エラーファイル移動失敗：" +
          moveError.message
        );

      }


      results.push({

        fileName:
          fileInfo.name,

        fileId:
          fileInfo.id,

        status:
          "ERROR",

        message:
          error.message

      });

    }

  });


  /* =================================================
   * 実行結果
   * ================================================= */

  const result = {

    success:
      true,

    summary:
      summary,

    files:
      results

  };


  Logger.log(
    "給与取込待ち一括処理完了: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 18-C. 判定不能理由生成
 * ===================================================== */

function cfPayrollBuildDetectionErrorMessage_(
  detection
) {

  if (!detection) {

    return (
      "ファイルの取込設定を判定できませんでした。"
    );

  }


  if (
    detection.status ===
    "UNSUPPORTED"
  ) {

    return (
      detection.reason ||
      "未対応のファイル形式です。"
    );

  }


  if (
    detection.status ===
    "UNMATCHED"
  ) {

    return (
      detection.reason ||
      "一致する取込設定がありません。"
    );

  }


  if (
    detection.status ===
    "AMBIGUOUS"
  ) {

    const names =
      (
        detection.matches ||
        []
      )
        .map(function(item) {

          return (
            item.displayName ||
            item.sourceId
          );

        })
        .join(", ");


    return (
      "複数の取込設定に一致しました。" +
      (
        names
          ? " 候補：" + names
          : ""
      )
    );

  }


  return (
    detection.reason ||
    "取込設定を判定できませんでした。"
  );

}


/* =====================================================
 * 18-D. 判定不能ファイルをログ保存
 *
 * PAY_FACTSには一切反映しない。
 * ===================================================== */

function cfPayrollLogRejectedImport_(
  fileInfo,
  detection,
  message
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_LOG"
    );


  if (!sheet) {

    throw new Error(
      "PAY_IMPORT_LOGがありません。"
    );

  }


  const importId =
    cfPayrollCreateImportId_();


  let dataType = "";


  if (
    detection &&
    detection.dataType
  ) {

    dataType =
      detection.dataType;

  }


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      12
    )
    .setValues([[
      importId,
      fileInfo.id,
      fileInfo.name,

      /* 内容ハッシュは本取込前なので空欄 */
      "",

      dataType,

      "",
      "",

      "判定エラー",

      0,
      0,

      new Date(),

      message
    ]]);


  SpreadsheetApp.flush();


  return importId;

}


/* =====================================================
 * 18-E. 必要な時だけファイル移動
 *
 * 既に目的フォルダ内なら何もしない。
 * ===================================================== */

function cfPayrollMoveFileIfNeeded_(
  fileId,
  destinationFolder
) {

  if (!destinationFolder) {

    throw new Error(
      "移動先フォルダがありません。"
    );

  }


  const file =
    DriveApp.getFileById(
      fileId
    );


  const destinationId =
    destinationFolder.getId();


  const parents =
    file.getParents();


  while (
    parents.hasNext()
  ) {

    const parent =
      parents.next();


    if (
      parent.getId() ===
      destinationId
    ) {

      return false;

    }

  }


  file.moveTo(
    destinationFolder
  );


  return true;

}


/* =====================================================
 * 18-F. 正の整数設定取得
 * ===================================================== */

function cfPayrollGetPositiveIntegerSetting_(
  key,
  defaultValue
) {

  const rawValue =
    cfPayrollGetSetting_(
      key,
      defaultValue,
      false
    );


  const number =
    Number(
      rawValue
    );


  if (
    !Number.isInteger(number) ||
    number < 1
  ) {

    throw new Error(
      "設定値は1以上の整数にしてください：" +
      key
    );

  }


  return number;

  }

/* =====================================================
 * 19. 営業日安全判定 V2
 *
 * ダウンロード日時を営業日として使用しない。
 *
 * 優先順位：
 * 1. 元データの日付列（COLUMN）
 * 2. ファイル名の「営業日YYYYMMDD」
 * 3. PAY_IMPORT_QUEUEの内容ハッシュ履歴
 * 4. 未確定 → PAY_IMPORT_QUEUEへ保留
 *
 * 日付未確定のまま本取込すると、
 * Section 17のblankDateチェックで停止する。
 * ===================================================== */

function cfPayrollApplyImportDateStrategy_(
  preview,
  prepared,
  source
) {

  if (!preview || !source) {
    throw new Error(
      "営業日判定に必要な情報がありません。"
    );
  }


  const method =
    String(
      source.dateMethod ||
      "NONE"
    )
      .trim()
      .toUpperCase();


  if (!preview.importSource) {
    preview.importSource = {};
  }


  preview.importSource.dateMethod =
    method;


  preview.importSource.filenameDateFormat =
    source.filenameDateFormat || "";


  /* =================================================
   * COLUMN
   *
   * 元データに営業日列がある場合は
   * その日付を最優先。
   * ================================================= */

  if (
    method ===
    "COLUMN"
  ) {

    const blankDateRows =
      (preview.rows || [])
        .filter(function(
          row
        ) {

          return !row.date;

        })
        .length;


    if (
      preview.summary
    ) {

      preview.summary.blankDateRows =
        blankDateRows;

    }


    preview.importSource.dateResolutionMethod =
      "COLUMN";


    return preview;

  }


  /* =================================================
   * NONE
   * ================================================= */

  if (
    method ===
    "NONE"
  ) {

    (preview.rows || [])
      .forEach(function(
        row
      ) {

        row.date = "";

      });


    if (
      preview.summary
    ) {

      preview.summary.blankDateRows =
        (preview.rows || []).length;

    }


    preview.importSource.resolvedDate =
      "";


    preview.importSource.dateResolutionMethod =
      "NONE";


    return preview;

  }


  /* =================================================
   * FILENAME
   *
   * 旧仕様のように単なるYYYYMMDDは使わない。
   *
   * 「営業日YYYYMMDD」
   * のみ明示営業日として扱う。
   * ================================================= */

  if (
    method ===
    "FILENAME"
  ) {

    const resolution =
      cfPayrollResolveBusinessDate_(
        prepared,
        source
      );


    if (
      resolution.date
    ) {

      cfPayrollApplyBusinessDateToPreview_(
        preview,
        resolution.date,
        resolution.method,
        resolution.queueId || ""
      );


      return preview;

    }


    /* ---------------------------------------------
     * 未確定
     *
     * プレビュー自体は成功。
     * 営業日は空欄。
     *
     * 本取込はblankDateチェックで停止。
     * --------------------------------------------- */

    (preview.rows || [])
      .forEach(function(
        row
      ) {

        row.date = "";

      });


    if (
      preview.summary
    ) {

      preview.summary.blankDateRows =
        (preview.rows || []).length;

    }


    preview.importSource.resolvedDate =
      "";


    preview.importSource.dateResolutionMethod =
      "PENDING_QUEUE";


    preview.importSource.dateQueueId =
      resolution.queueId || "";


    return preview;

  }


  throw new Error(
    "未対応の日付取得方法です：" +
    method
  );

}


/* =====================================================
 * 19-A. 営業日解決
 * ===================================================== */

function cfPayrollResolveBusinessDate_(
  prepared,
  source
) {

  const fileName =
    String(
      prepared.originalFileName ||
      ""
    );


  /* =================================================
   * 1. 明示ファイル名
   *
   * 例：
   *
   * オーダー売上照会_
   * 営業日20260901_
   * 20260914_233104.xlsx
   * ================================================= */

  const explicitDate =
    cfPayrollExtractExplicitBusinessDateFromFileName_(
      fileName
    );


  if (
    explicitDate
  ) {

    const hash =
      cfPayrollComputeImportHash_(
        prepared.originalFileId
      );


    const queueInfo =
      cfPayrollUpsertImportDateQueue_(
        prepared,
        source,
        hash,
        explicitDate,
        "EXPLICIT_FILENAME"
      );


    return {

      date:
        explicitDate,

      method:
        "EXPLICIT_FILENAME",

      queueId:
        queueInfo.queueId

    };

  }


  /* =================================================
   * 2. 内容ハッシュ
   * ================================================= */

  const hash =
    cfPayrollComputeImportHash_(
      prepared.originalFileId
    );


  const history =
    cfPayrollFindImportDateByHash_(
      hash,
      source.sourceId
    );


  if (
    history &&
    history.businessDate
  ) {

    cfPayrollUpsertImportDateQueue_(
      prepared,
      source,
      hash,
      history.businessDate,
      "HASH_HISTORY"
    );


    return {

      date:
        history.businessDate,

      method:
        "HASH_HISTORY",

      queueId:
        history.queueId

    };

  }


  /* =================================================
   * 3. 初見ファイル
   *
   * 営業日不明としてキューへ。
   * ================================================= */

  const queueInfo =
    cfPayrollUpsertImportDateQueue_(
      prepared,
      source,
      hash,
      "",
      "PENDING"
    );


  /* =================================================
   * 4. キューで営業日入力済み
   * ================================================= */

  if (
    queueInfo.businessDate
  ) {

    return {

      date:
        queueInfo.businessDate,

      method:
        queueInfo.resolutionMethod ||
        "MANUAL_QUEUE",

      queueId:
        queueInfo.queueId

    };

  }


  return {

    date:
      null,

    method:
      "PENDING_QUEUE",

    queueId:
      queueInfo.queueId

  };

}


/* =====================================================
 * 19-B. ファイル名から明示営業日取得
 *
 * 単なる20260914は無視。
 *
 * 「営業日」の文字が必須。
 *
 * 対応：
 *
 * 営業日20260901
 * 営業日_20260901
 * 営業日2026-09-01
 * 営業日_2026_09_01
 * ===================================================== */

function cfPayrollExtractExplicitBusinessDateFromFileName_(
  fileName
) {

  let text =
    String(
      fileName ||
      ""
    );


  try {

    text =
      text.normalize(
        "NFKC"
      );

  } catch (
    error
  ) {

    /*
     * normalize非対応でも続行
     */

  }


  const match =
    text.match(
      /営業日[\s_-]*(20\d{2})[\/_-]?(\d{2})[\/_-]?(\d{2})/
    );


  if (
    !match
  ) {

    return null;

  }


  return cfPayrollBuildBusinessDate_(
    Number(
      match[1]
    ),
    Number(
      match[2]
    ),
    Number(
      match[3]
    )
  );

}


/* =====================================================
 * 19-C. 営業日Date生成
 * ===================================================== */

function cfPayrollBuildBusinessDate_(
  year,
  month,
  day
) {

  const date =
    new Date(
      year,
      month - 1,
      day,
      0,
      0,
      0,
      0
    );


  if (
    date.getFullYear() !==
      year ||
    date.getMonth() !==
      month - 1 ||
    date.getDate() !==
      day
  ) {

    return null;

  }


  return date;

}


/* =====================================================
 * 19-D. プレビューへ営業日反映
 * ===================================================== */

function cfPayrollApplyBusinessDateToPreview_(
  preview,
  businessDate,
  resolutionMethod,
  queueId
) {

  (preview.rows || [])
    .forEach(function(
      row
    ) {

      row.date =
        businessDate;

    });


  if (
    preview.summary
  ) {

    preview.summary.blankDateRows =
      0;

  }


  if (
    !preview.importSource
  ) {

    preview.importSource =
      {};

  }


  preview.importSource.resolvedDate =
    businessDate;


  preview.importSource.dateResolutionMethod =
    resolutionMethod;


  preview.importSource.dateQueueId =
    queueId || "";

}


/* =====================================================
 * 19-E. 旧関数互換
 *
 * 旧関数名は残す。
 *
 * ただし現在は
 * 「営業日」明示付きのみ認識。
 * ===================================================== */

function cfPayrollExtractDateFromFileName_(
  fileName
) {

  return cfPayrollExtractExplicitBusinessDateFromFileName_(
    fileName
  );

}


/* =====================================================
 * 19-F. 営業日判定テスト
 * ===================================================== */

function testPayrollFilenameDate() {

  const cases = [

    {

      fileName:
        "オーダー売上照会_営業日20260901_20260914_233104.xlsx",

      expected:
        "2026/09/01"

    },

    {

      fileName:
        "オーダー売上照会_営業日2026-09-02_20260914_233119.xlsx",

      expected:
        "2026/09/02"

    },

    {

      /*
       * ダウンロード日は営業日にしない
       */
      fileName:
        "オーダー売上照会_20260914_233131.xlsx",

      expected:
        ""

    }

  ];


  const timezone =
    cfPayrollGetSetting_(
      "TIMEZONE",
      "Asia/Tokyo",
      false
    );


  const results =
    cases.map(function(
      item
    ) {

      const date =
        cfPayrollExtractExplicitBusinessDateFromFileName_(
          item.fileName
        );


      const actual =
        date
          ? Utilities.formatDate(
              date,
              timezone,
              "yyyy/MM/dd"
            )
          : "";


      return {

        fileName:
          item.fileName,

        expected:
          item.expected,

        actual:
          actual,

        success:
          actual ===
          item.expected

      };

    });


  const success =
    results.every(function(
      item
    ) {

      return item.success;

    });


  const result = {

    success:
      success,

    cases:
      results

  };


  Logger.log(
    "営業日ファイル名判定テスト: " +
    JSON.stringify(
      result
    )
  );


  if (
    !success
  ) {

    throw new Error(
      "営業日ファイル名判定テストに失敗しました。"
    );

  }


  return result;

}


/* =====================================================
 * 20. 商品マッピング
 *
 * PAY_PRODUCT_MAP
 *
 * オーダー売上照会の商品を
 * 給与システム上の項目キーへ変換する。
 *
 * 商品コードを最優先で照合。
 * コードが無い場合のみ商品名で照合する。
 *
 * 値取得元：
 *
 * QUANTITY
 *   → 数量
 *
 * AMOUNT
 *   → 金額
 *
 * POINTS
 *   → ポイント
 *
 * BACK
 *   → 元システムのバック値
 *
 * 基本はQUANTITYを使用し、
 * 実際のバック金額はPAY_RULES側で計算する。
 * ===================================================== */


/* =====================================================
 * 20-A. 商品マッピング初期設定
 * ===================================================== */

function setupPayrollProductMap() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const sheetName =
      "PAY_PRODUCT_MAP";


    const headers = [
      "商品マッピングID",
      "商品コード",
      "商品名",
      "項目キー",
      "表示名",
      "値取得元",
      "倍率",
      "単位",
      "有効",
      "確認状態",
      "備考"
    ];


    let sheet =
      book.getSheetByName(
        sheetName
      );


    /* =================================================
     * シート作成
     * ================================================= */

    if (!sheet) {

      sheet =
        book.insertSheet(
          sheetName
        );

    }


    /* =================================================
     * 空シート
     * ================================================= */

    if (
      sheet.getLastRow() === 0
    ) {

      cfPayrollEnsureColumns_(
        sheet,
        headers.length
      );


      cfPayrollWriteHeader_(
        sheet,
        headers
      );

    }


    /* =================================================
     * 既存シート構成確認
     * ================================================= */

    else if (
      !cfPayrollHeaderMatches_(
        sheet,
        headers
      )
    ) {

      throw new Error(
        "PAY_PRODUCT_MAPの列構成が異なります。" +
        "データ保護のため変更せず停止しました。"
      );

    }


    /* =================================================
     * 実ファイルから確認できた商品
     *
     * アネホ・クライナー・誕生日カクテルは
     * 独立キーで保持。
     *
     * 後からPAY_RULESで金額を決めるため、
     * 勝手に「ショット」と同額にはしない。
     * ================================================= */

    const defaults = [

      [
        "PROD_0010",
        "0010",
        "ショット（各種）",
        "BACK_SHOT",
        "ショット",
        "QUANTITY",
        1,
        "杯",
        true,
        "確認済",
        "オーダー売上照会の商品コード0010"
      ],


      [
        "PROD_0040",
        "0040",
        "お誕生日カクテル☆彡",
        "BACK_BIRTHDAY_COCKTAIL",
        "お誕生日カクテル",
        "QUANTITY",
        1,
        "杯",
        true,
        "要設定",
        "給与バック金額はPAY_RULESで後から設定"
      ],


      [
        "PROD_0111",
        "0111",
        "キャストドリンク",
        "BACK_DRINK",
        "キャストドリンク",
        "QUANTITY",
        1,
        "杯",
        true,
        "確認済",
        "オーダー売上照会の商品コード0111"
      ],


      [
        "PROD_0116",
        "0116",
        "アネホ",
        "BACK_ANEJO",
        "アネホ",
        "QUANTITY",
        1,
        "杯",
        true,
        "要設定",
        "ショット扱いにするか独立バックにするか後から設定可能"
      ],


      [
        "PROD_0129",
        "0129",
        "クライナー",
        "BACK_KLEINER",
        "クライナー",
        "QUANTITY",
        1,
        "本",
        true,
        "要設定",
        "給与バック金額はPAY_RULESで後から設定"
      ]

    ];


    /* =================================================
     * 既存マッピングID取得
     * ================================================= */

    const existing =
      cfPayrollGetExistingKeys_(
        sheet,
        1
      );


    const newRows =
      defaults.filter(function(
        row
      ) {

        return !existing[
          row[0]
        ];

      });


    /* =================================================
     * 未登録分だけ追加
     * ================================================= */

    if (
      newRows.length > 0
    ) {

      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          newRows.length,
          headers.length
        )
        .setValues(
          newRows
        );

    }


    /* =================================================
     * 値取得元プルダウン
     * ================================================= */

    const valueSourceRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            "QUANTITY",
            "AMOUNT",
            "POINTS",
            "BACK"
          ],
          true
        )
        .setAllowInvalid(
          false
        )
        .build();


    sheet
      .getRange(
        2,
        6,
        sheet.getMaxRows() - 1,
        1
      )
      .setDataValidation(
        valueSourceRule
      );


    /* =================================================
     * 有効チェックボックス
     * ================================================= */

    sheet
      .getRange(
        2,
        9,
        sheet.getMaxRows() - 1,
        1
      )
      .insertCheckboxes();


    /* =================================================
     * 確認状態
     * ================================================= */

    const statusRule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          [
            "確認済",
            "要設定",
            "無効"
          ],
          true
        )
        .setAllowInvalid(
          false
        )
        .build();


    sheet
      .getRange(
        2,
        10,
        sheet.getMaxRows() - 1,
        1
      )
      .setDataValidation(
        statusRule
      );


    /* =================================================
     * 倍率表示
     * ================================================= */

    sheet
      .getRange(
        2,
        7,
        sheet.getMaxRows() - 1,
        1
      )
      .setNumberFormat(
        "0.########"
      );


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      productsAdded:
        newRows.length,

      productsTotal:
        sheet.getLastRow() - 1

    };


    Logger.log(
      "給与商品マッピング設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 20-B. 商品マッピング取得
 *
 * 戻り値：
 *
 * {
 *   byCode: {...},
 *   byName: {...}
 * }
 * ===================================================== */

function cfPayrollGetProductMappings_() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_PRODUCT_MAP"
    );


  if (!sheet) {

    throw new Error(
      "PAY_PRODUCT_MAPがありません。" +
      "setupPayrollProductMap()を実行してください。"
    );

  }


  const headers = [
    "商品マッピングID",
    "商品コード",
    "商品名",
    "項目キー",
    "表示名",
    "値取得元",
    "倍率",
    "単位",
    "有効",
    "確認状態",
    "備考"
  ];


  if (
    !cfPayrollHeaderMatches_(
      sheet,
      headers
    )
  ) {

    throw new Error(
      "PAY_PRODUCT_MAPの列構成が異なります。"
    );

  }


  const result = {

    byCode: {},
    byName: {}

  };


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return result;

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  rows.forEach(function(
    row,
    index
  ) {

    const mappingId =
      String(
        row[0] || ""
      ).trim();


    /*
     * getValues()だと
     * "0010"が数値化される可能性もあるため、
     * 4桁コードへ正規化。
     */
    const productCode =
      cfPayrollNormalizeProductCode_(
        row[1]
      );


    const productName =
      String(
        row[2] || ""
      ).trim();


    const itemKey =
      String(
        row[3] || ""
      ).trim();


    const displayName =
      String(
        row[4] || ""
      ).trim();


    const valueSource =
      String(
        row[5] || "QUANTITY"
      )
        .trim()
        .toUpperCase();


    const multiplierRaw =
      row[6];


    const unit =
      String(
        row[7] || ""
      ).trim();


    const enabled =
      row[8] === true;


    const status =
      String(
        row[9] || ""
      ).trim();


    const note =
      String(
        row[10] || ""
      ).trim();


    if (!enabled) {
      return;
    }


    if (
      !mappingId ||
      !itemKey
    ) {

      throw new Error(
        "PAY_PRODUCT_MAPの必須項目が不足しています。" +
        " 行：" +
        (index + 2)
      );

    }


    if (
      !productCode &&
      !productName
    ) {

      throw new Error(
        mappingId +
        " は商品コードまたは商品名が必要です。"
      );

    }


    if (
      [
        "QUANTITY",
        "AMOUNT",
        "POINTS",
        "BACK"
      ].indexOf(
        valueSource
      ) === -1
    ) {

      throw new Error(
        mappingId +
        " の値取得元が不正です：" +
        valueSource
      );

    }


    let multiplier =
      1;


    if (
      multiplierRaw !== "" &&
      multiplierRaw !== null
    ) {

      multiplier =
        Number(
          multiplierRaw
        );


      if (
        !Number.isFinite(
          multiplier
        )
      ) {

        throw new Error(
          mappingId +
          " の倍率が数値ではありません。"
        );

      }

    }


    const mapping = {

      mappingId:
        mappingId,

      productCode:
        productCode,

      productName:
        productName,

      itemKey:
        itemKey,

      displayName:
        displayName,

      valueSource:
        valueSource,

      multiplier:
        multiplier,

      unit:
        unit,

      status:
        status,

      note:
        note,

      sheetRow:
        index + 2

    };


    /* =================================================
     * 商品コード辞書
     * ================================================= */

    if (productCode) {

      if (
        result.byCode[
          productCode
        ]
      ) {

        throw new Error(
          "PAY_PRODUCT_MAPの商品コードが重複しています：" +
          productCode
        );

      }


      result.byCode[
        productCode
      ] =
        mapping;

    }


    /* =================================================
     * 商品名辞書
     * ================================================= */

    if (productName) {

      const normalizedName =
        cfPayrollNormalizeProductName_(
          productName
        );


      /*
       * 同名で別設定がある場合は
       * 商品名だけで自動判定しない。
       */
      if (
        Object.prototype
          .hasOwnProperty.call(
            result.byName,
            normalizedName
          )
      ) {

        result.byName[
          normalizedName
        ] =
          null;

      } else {

        result.byName[
          normalizedName
        ] =
          mapping;

      }

    }

  });


  return result;

}


/* =====================================================
 * 20-C. 商品マッピング解決
 *
 * 優先順位：
 *
 * 1. 商品コード
 * 2. 商品名
 * ===================================================== */

function cfPayrollResolveProductMapping_(
  mappings,
  productCode,
  productName
) {

  const normalizedCode =
    cfPayrollNormalizeProductCode_(
      productCode
    );


  if (
    normalizedCode &&
    mappings.byCode[
      normalizedCode
    ]
  ) {

    return mappings.byCode[
      normalizedCode
    ];

  }


  const normalizedName =
    cfPayrollNormalizeProductName_(
      productName
    );


  if (
    normalizedName &&
    mappings.byName[
      normalizedName
    ]
  ) {

    return mappings.byName[
      normalizedName
    ];

  }


  return null;

}


/* =====================================================
 * 20-D. 商品コード正規化
 *
 * 10
 * "10"
 * "0010"
 *
 * を全て
 *
 * "0010"
 *
 * として扱う。
 * ===================================================== */

function cfPayrollNormalizeProductCode_(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return "";

  }


  let text =
    String(
      value
    ).trim();


  if (!text) {
    return "";
  }


  /*
   * Excelから数値として来た場合
   */
  if (
    /^\d+$/.test(
      text
    )
  ) {

    return text.padStart(
      4,
      "0"
    );

  }


  return text;

}


/* =====================================================
 * 20-E. 商品名正規化
 * ===================================================== */

function cfPayrollNormalizeProductName_(
  value
) {

  let text =
    String(
      value || ""
    );


  if (
    typeof text.normalize ===
    "function"
  ) {

    text =
      text.normalize(
        "NFKC"
      );

  }


  return text
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, "")
    .trim()
    .toLowerCase();

}


/* =====================================================
 * 20-F. 商品の値取得
 *
 * オーダー売上照会の
 *
 * 数量
 * 金額
 * ポイント
 * バック
 *
 * からマッピング設定に応じて選択。
 * ===================================================== */

function cfPayrollGetMappedProductValue_(
  mapping,
  values
) {

  if (!mapping) {

    throw new Error(
      "商品マッピングがありません。"
    );

  }


  values =
    values || {};


  let rawValue;


  if (
    mapping.valueSource ===
    "QUANTITY"
  ) {

    rawValue =
      values.quantity;

  }


  else if (
    mapping.valueSource ===
    "AMOUNT"
  ) {

    rawValue =
      values.amount;

  }


  else if (
    mapping.valueSource ===
    "POINTS"
  ) {

    rawValue =
      values.points;

  }


  else if (
    mapping.valueSource ===
    "BACK"
  ) {

    rawValue =
      values.back;

  }


  else {

    throw new Error(
      "未対応の商品値取得元です：" +
      mapping.valueSource
    );

  }


  if (
    rawValue === "" ||
    rawValue === null ||
    rawValue === undefined
  ) {

    return 0;

  }


  const number =
    Number(
      rawValue
    );


  if (
    !Number.isFinite(
      number
    )
  ) {

    throw new Error(
      mapping.displayName +
      " の値を数値として読み取れません：" +
      rawValue
    );

  }


  return (
    number *
    mapping.multiplier
  );

}


/* =====================================================
 * 20-G. 商品マッピング読込テスト
 *
 * 給与データには触れない。
 * ===================================================== */

function testPayrollProductMappings() {

  const mappings =
    cfPayrollGetProductMappings_();


  const result = {

    success:
      true,

    codeCount:
      Object.keys(
        mappings.byCode
      ).length,

    nameCount:
      Object.keys(
        mappings.byName
      ).length,

    products:
      Object.keys(
        mappings.byCode
      ).map(function(
        code
      ) {

        const mapping =
          mappings.byCode[
            code
          ];


        return {

          productCode:
            code,

          productName:
            mapping.productName,

          itemKey:
            mapping.itemKey,

          valueSource:
            mapping.valueSource

        };

      })

  };


  Logger.log(
    "給与商品マッピング読込OK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 21. オーダー売上照会 PRODUCT_ROWS解析
 *
 * 実ファイル：
 *
 * キャストコード
 * キャスト名
 * コード
 * 商品名
 * 数量
 * 金額
 * ポイント
 * バック
 *
 * PAY_PRODUCT_MAPと照合し
 * PAY_FACTS用の実績へ変換する。
 *
 * キャスト情報が完全空欄の商品行は
 * 給与対象外としてスキップする。
 * ===================================================== */


/* =====================================================
 * 21-A. 商品解析安全設定
 * ===================================================== */

function setupPayrollProductParser() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    cfPayrollSaveSystemSetting_(
      "IMPORT_BLOCK_UNKNOWN_PRODUCTS",
      "未登録商品時に取込停止",
      true,
      "true = PAY_PRODUCT_MAPに無い商品があればオーダー売上の本取込を停止"
    );


    SpreadsheetApp.flush();


    const result = {

      success:
        true

    };


    Logger.log(
      "商品行解析設定OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 21-B. 解析方式振り分け
 *
 * PRODUCT_ROWS
 *   → オーダー売上
 *
 * ATTENDANCE_ROWS
 *   → キャスト勤怠
 *
 * PERFORMANCE_ROWS
 *   → キャスト成績一覧
 *
 * SLIP_ROWS
 *   → 伝票一覧
 *
 * SETTLEMENT_DETAIL
 *   → 精算入金監査
 *
 * その他
 *   → 従来GENERIC解析
 * ===================================================== */

function cfPayrollAnalyzeImportSpreadsheetV2_(
  sourceBook,
  prepared,
  source
) {

  const parserMode =
    String(
      source.parserMode ||
      "GENERIC"
    )
      .trim()
      .toUpperCase();


  if (
    parserMode ===
    "PRODUCT_ROWS"
  ) {

    return cfPayrollAnalyzeProductRows_(
      sourceBook,
      prepared,
      source
    );

  }


  if (
    parserMode ===
    "ATTENDANCE_ROWS"
  ) {

    return cfPayrollAnalyzeAttendanceRows_(
      sourceBook,
      prepared,
      source
    );

  }


  if (
    parserMode ===
    "PERFORMANCE_ROWS"
  ) {

    return cfPayrollAnalyzePerformanceRows_(
      sourceBook,
      prepared,
      source
    );

  }


  if (
    parserMode ===
    "SLIP_ROWS"
  ) {

    return cfPayrollAnalyzeSlipRows_(
      sourceBook,
      prepared,
      source
    );

  }


  if (
    parserMode ===
    "SETTLEMENT_DETAIL"
  ) {

    return cfPayrollAnalyzeSettlementDetail_(
      sourceBook,
      prepared,
      source
    );

  }


  return cfPayrollAnalyzeImportSpreadsheet_(
    sourceBook,
    prepared,
    source
  );

}


/* =====================================================
 * 21-C. オーダー売上照会解析
 * ===================================================== */

function cfPayrollAnalyzeProductRows_(
  sourceBook,
  prepared,
  source
) {

  /* =================================================
   * 対象シート
   * ================================================= */

  let sourceSheet;


  if (
    source.sheetName
  ) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / ファイル：" +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (
      sheets.length === 0
    ) {

      throw new Error(
        "取込ファイルにシートがありません。"
      );

    }


    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow <
      source.startRow ||
    lastColumn < 1
  ) {

    throw new Error(
      "取込対象データがありません。" +
      " シート：" +
      sourceSheet.getName()
    );

  }


  /* =================================================
   * ヘッダー
   * ================================================= */

  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  /* =================================================
   * 列位置
   * ================================================= */

  const castIdIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castIdHeader ||
      "キャストコード"
    );


  const castNameIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castNameHeader ||
      "キャスト名"
    );


  const productCodeIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "コード"
    );


  const productNameIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "商品名"
    );


  const quantityIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "数量"
    );


  const amountIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "金額"
    );


  const pointsIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "ポイント"
    );


  const backIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "バック"
    );


  /* =================================================
   * 必須列確認
   * ================================================= */

  const required = [

    [
      castIdIndex,
      source.castIdHeader ||
      "キャストコード"
    ],

    [
      castNameIndex,
      source.castNameHeader ||
      "キャスト名"
    ],

    [
      productCodeIndex,
      "コード"
    ],

    [
      productNameIndex,
      "商品名"
    ],

    [
      quantityIndex,
      "数量"
    ],

    [
      amountIndex,
      "金額"
    ],

    [
      pointsIndex,
      "ポイント"
    ],

    [
      backIndex,
      "バック"
    ]

  ];


  required.forEach(function(
    item
  ) {

    if (
      item[0] === -1
    ) {

      throw new Error(
        "オーダー売上照会に必要な列がありません：" +
        item[1]
      );

    }

  });


  /* =================================================
   * データ取得
   * ================================================= */

  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getValues();


  /* =================================================
   * キャスト / 商品辞書
   * ================================================= */

  const castResolver =
    cfPayrollBuildCastResolver_();


  const productMappings =
    cfPayrollGetProductMappings_();


  /* =================================================
   * 解析結果
   * ================================================= */

  const previewRows =
    [];


  const unknownCasts =
    {};


  const unknownProducts =
    {};


  const recognizedCastIds =
    {};


  const itemTotals =
    {};


  let skippedEmptyRows =
    0;


  let skippedNoCastRows =
    0;


  let mappedProductRows =
    0;


  /* =================================================
   * 各行解析
   * ================================================= */

  values.forEach(function(
    row,
    index
  ) {

    const sourceRowNumber =
      source.startRow +
      index;


    /* ---------------------------------------------
     * 完全空行
     * --------------------------------------------- */

    const hasAnyValue =
      row.some(function(
        value
      ) {

        return (
          value !== "" &&
          value !== null
        );

      });


    if (!hasAnyValue) {

      skippedEmptyRows++;

      return;

    }


    /* ---------------------------------------------
     * キャスト
     * --------------------------------------------- */

    const sourceCastId =
      String(
        row[
          castIdIndex
        ] || ""
      ).trim();


    const sourceCastName =
      String(
        row[
          castNameIndex
        ] || ""
      ).trim();


    /*
     * キャストが完全空欄
     *
     * → 客注文等と判断
     * → 給与対象外
     */
    if (
      !sourceCastId &&
      !sourceCastName
    ) {

      skippedNoCastRows++;

      return;

    }


    /* ---------------------------------------------
     * CAST_ID解決
     * --------------------------------------------- */

    const payrollCastId =
      cfPayrollResolveCastId_(
        castResolver,
        sourceCastId,
        sourceCastName
      );


    if (
      payrollCastId
    ) {

      recognizedCastIds[
        payrollCastId
      ] =
        true;

    } else {

      const unknownCastKey =
        sourceCastId +
        "::" +
        sourceCastName;


      if (
        !unknownCasts[
          unknownCastKey
        ]
      ) {

        unknownCasts[
          unknownCastKey
        ] = {

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          rows:
            []

        };

      }


      unknownCasts[
        unknownCastKey
      ].rows.push(
        sourceRowNumber
      );

    }


    /* ---------------------------------------------
     * 商品情報
     * --------------------------------------------- */

    const rawProductCode =
      row[
        productCodeIndex
      ];


    const productCode =
      cfPayrollNormalizeProductCode_(
        rawProductCode
      );


    const productName =
      String(
        row[
          productNameIndex
        ] || ""
      ).trim();


    /* ---------------------------------------------
     * 商品マッピング解決
     * --------------------------------------------- */

    const mapping =
      cfPayrollResolveProductMapping_(
        productMappings,
        productCode,
        productName
      );


    const facts =
      [];


    /* ---------------------------------------------
     * 未登録商品
     * --------------------------------------------- */

    if (!mapping) {

      const unknownProductKey =
        productCode +
        "::" +
        cfPayrollNormalizeProductName_(
          productName
        );


      if (
        !unknownProducts[
          unknownProductKey
        ]
      ) {

        unknownProducts[
          unknownProductKey
        ] = {

          productCode:
            productCode,

          productName:
            productName,

          rows:
            []

        };

      }


      unknownProducts[
        unknownProductKey
      ].rows.push(
        sourceRowNumber
      );

    }


    /* ---------------------------------------------
     * 登録済商品
     * --------------------------------------------- */

    else {

      const sourceValues = {

        quantity:
          row[
            quantityIndex
          ],

        amount:
          row[
            amountIndex
          ],

        points:
          row[
            pointsIndex
          ],

        back:
          row[
            backIndex
          ]

      };


      const convertedValue =
        cfPayrollGetMappedProductValue_(
          mapping,
          sourceValues
        );


      const rawValue =
        cfPayrollGetProductRawValue_(
          mapping,
          sourceValues
        );


      facts.push({

        itemKey:
          mapping.itemKey,

        displayName:
          mapping.displayName,

        rawValue:
          rawValue,

        value:
          convertedValue,

        unit:
          mapping.unit,

        multiplier:
          mapping.multiplier,

        sourceHeader:
          "商品:" +
          (
            productName ||
            productCode
          ),

        productCode:
          productCode,

        productName:
          productName

      });


      mappedProductRows++;


      /* -------------------------------------------
       * 項目別合計
       * ------------------------------------------- */

      if (
        typeof convertedValue ===
        "number"
      ) {

        if (
          !Object.prototype
            .hasOwnProperty.call(
              itemTotals,
              mapping.itemKey
            )
        ) {

          itemTotals[
            mapping.itemKey
          ] = {

            displayName:
              mapping.displayName,

            unit:
              mapping.unit,

            total:
              0

          };

        }


        itemTotals[
          mapping.itemKey
        ].total +=
          convertedValue;

      }

    }


    /* ---------------------------------------------
     * プレビュー行
     * --------------------------------------------- */

    previewRows.push({

      sourceRow:
        sourceRowNumber,

      /*
       * Section 19で
       * ファイル名から営業日を入れる
       */
      date:
        "",

      sourceCastId:
        sourceCastId,

      sourceCastName:
        sourceCastName,

      castId:
        payrollCastId,

      productCode:
        productCode,

      productName:
        productName,

      facts:
        facts

    });

  });


  /* =================================================
   * 想定外ヘッダー
   * ================================================= */

  const expectedHeaderIndexes =
    {};


  [
    castIdIndex,
    castNameIndex,
    productCodeIndex,
    productNameIndex,
    quantityIndex,
    amountIndex,
    pointsIndex,
    backIndex
  ].forEach(function(
    index
  ) {

    expectedHeaderIndexes[
      index
    ] =
      true;

  });


  const unmappedHeaders =
    headers.filter(function(
      header,
      index
    ) {

      return (
        header &&
        !expectedHeaderIndexes[
          index
        ]
      );

    });


  /* =================================================
   * 結果
   * ================================================= */

  const result = {

    success:
      true,

    mode:
      "PREVIEW_ONLY",

    file: {

      id:
        prepared.originalFileId,

      name:
        prepared.originalFileName,

      type:
        prepared.originalFileType,

      temporaryConversion:
        prepared.temporary,

      sheetName:
        sourceSheet.getName()

    },


    importSource: {

      sourceId:
        source.sourceId,

      displayName:
        source.displayName,

      dataType:
        source.dataType,

      parserMode:
        "PRODUCT_ROWS"

    },


    summary: {

      sourceRows:
        values.length,

      recognizedRows:
        previewRows.length,

      skippedEmptyRows:
        skippedEmptyRows,

      skippedNoCastRows:
        skippedNoCastRows,

      blankDateRows:
        0,

      mappedColumns:
        8,

      unmappedColumns:
        unmappedHeaders.length,

      mappedProductRows:
        mappedProductRows,

      recognizedCasts:
        Object.keys(
          recognizedCastIds
        ).length,

      unknownCastCount:
        Object.keys(
          unknownCasts
        ).length,

      unknownProductCount:
        Object.keys(
          unknownProducts
        ).length

    },


    mappedColumns: [

      {
        sourceHeader:
          "キャストコード",

        role:
          "CAST_ID"
      },

      {
        sourceHeader:
          "キャスト名",

        role:
          "CAST_NAME"
      },

      {
        sourceHeader:
          "コード",

        role:
          "PRODUCT_CODE"
      },

      {
        sourceHeader:
          "商品名",

        role:
          "PRODUCT_NAME"
      },

      {
        sourceHeader:
          "数量",

        role:
          "QUANTITY"
      },

      {
        sourceHeader:
          "金額",

        role:
          "AMOUNT"
      },

      {
        sourceHeader:
          "ポイント",

        role:
          "POINTS"
      },

      {
        sourceHeader:
          "バック",

        role:
          "BACK"
      }

    ],


    unmappedHeaders:
      unmappedHeaders,


    itemTotals:
      itemTotals,


    unknownCasts:
      Object.keys(
        unknownCasts
      ).map(function(
        key
      ) {

        return unknownCasts[
          key
        ];

      }),


    unknownProducts:
      Object.keys(
        unknownProducts
      ).map(function(
        key
      ) {

        return unknownProducts[
          key
        ];

      }),


    rows:
      previewRows

  };


  Logger.log(
    "オーダー売上解析OK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 21-D. 商品マッピングの元値取得
 * ===================================================== */

function cfPayrollGetProductRawValue_(
  mapping,
  values
) {

  if (
    mapping.valueSource ===
    "QUANTITY"
  ) {

    return values.quantity;

  }


  if (
    mapping.valueSource ===
    "AMOUNT"
  ) {

    return values.amount;

  }


  if (
    mapping.valueSource ===
    "POINTS"
  ) {

    return values.points;

  }


  if (
    mapping.valueSource ===
    "BACK"
  ) {

    return values.back;

  }


  throw new Error(
    "未対応の商品値取得元です：" +
    mapping.valueSource
  );

}


/* =====================================================
 * 21-E. 本取込安全チェック V2
 *
 * 従来チェック
 * +
 * PRODUCT_ROWSの場合は
 * 未登録商品も確認する。
 * ===================================================== */

function cfPayrollValidateImportPreviewV2_(
  preview,
  source
) {

  /* =================================================
   * 従来安全チェック
   * ================================================= */

  cfPayrollValidateImportPreview_(
    preview,
    source
  );


  const parserMode =
    String(
      source.parserMode ||
      ""
    )
      .trim()
      .toUpperCase();


  /*
   * PRODUCT_ROWS以外は
   * 追加チェック不要
   */
  if (
    parserMode !==
    "PRODUCT_ROWS"
  ) {

    return true;

  }


  /* =================================================
   * 未登録商品停止設定
   * ================================================= */

  const blockUnknownProducts =
    cfPayrollToBoolean_(
      cfPayrollGetSetting_(
        "IMPORT_BLOCK_UNKNOWN_PRODUCTS",
        true,
        false
      )
    );


  const unknownProductCount =
    Number(
      preview &&
      preview.summary
        ? preview.summary
            .unknownProductCount || 0
        : 0
    );


  if (
    blockUnknownProducts &&
    unknownProductCount > 0
  ) {

    const products =
      (
        preview.unknownProducts ||
        []
      )
        .map(function(
          item
        ) {

          if (
            item.productCode &&
            item.productName
          ) {

            return (
              item.productCode +
              " " +
              item.productName
            );

          }


          return (
            item.productName ||
            item.productCode ||
            "名称不明"
          );

        })
        .join(", ");


    throw new Error(
      "PAY_PRODUCT_MAPに未登録の商品があるため取込を停止しました：" +
      products
    );

  }


  return true;

}


/* =====================================================
 * 22. ROYAL POSキャストコード同期
 *
 * 01_取込待ち内の実ファイルから
 *
 * キャストコード
 * キャスト名
 *
 * を読み取り、
 *
 * PAY_CAST_MAPへ
 *
 * ROYAL_POS
 * 0020
 * まのあ
 * ↓
 * CAST FLOW側CAST_ID
 *
 * の対応を自動登録する。
 *
 * 名前が一意に一致した場合のみ
 * 自動でCAST_IDを設定する。
 *
 * 一致しない場合は
 * CAST_ID空欄 + 要確認
 * として安全に残す。
 * ===================================================== */


/* =====================================================
 * 22-A. 取込待ちファイルからPOSキャスト同期
 *
 * 安全版：
 *
 * 元キャストID + 元キャスト名
 * の組み合わせを1人候補として扱う。
 *
 * 同じコードでも名前が違えば
 * 別行としてPAY_CAST_MAPへ保存する。
 *
 * ファイル移動・PAY_FACTS反映は行わない。
 * ===================================================== */

function syncPayrollPosCastMapFromInbox() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(
    30000
  );


  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const mapSheet =
      book.getSheetByName(
        "PAY_CAST_MAP"
      );


    if (!mapSheet) {

      throw new Error(
        "PAY_CAST_MAPがありません。"
      );

    }


    const headers = [
      "元システム",
      "元キャストID",
      "元キャスト名",
      "CAST_ID",
      "確認状態",
      "更新日時"
    ];


    if (
      !cfPayrollHeaderMatches_(
        mapSheet,
        headers
      )
    ) {

      throw new Error(
        "PAY_CAST_MAPの列構成が異なります。"
      );

    }


    const SOURCE_SYSTEM =
      "ROYAL_POS";


    const castFlowNameMap =
      cfPayrollBuildCastFlowNameMap_(
        mapSheet
      );


    const existingPosMap =
      cfPayrollBuildExistingPosCastMap_(
        mapSheet,
        SOURCE_SYSTEM
      );


    const inboxFolder =
      cfPayrollGetInboxFolder_();


    const sources =
      cfPayrollGetImportSources_();


    const files =
      inboxFolder.getFiles();


    const discovered =
      {};


    let scannedFiles =
      0;


    let matchedFiles =
      0;


    while (
      files.hasNext()
    ) {

      const file =
        files.next();


      scannedFiles++;


      const fileInfo =
        cfPayrollBuildImportFileInfo_(
          file
        );


      const detection =
        cfPayrollDetectImportSourceForFile_(
          fileInfo,
          sources
        );


      if (
        detection.status !==
        "MATCHED"
      ) {

        continue;

      }


      const source =
        cfPayrollGetImportSourceById_(
          detection.sourceId
        );


      if (!source) {

        continue;

      }


      if (
        !source.castIdHeader &&
        !source.castNameHeader
      ) {

        continue;

      }


      matchedFiles++;


      cfPayrollWithReadableSpreadsheet_(

        fileInfo.id,

        function(
          sourceBook,
          prepared
        ) {

          const aliases =
            cfPayrollExtractCastAliasesFromSpreadsheet_(
              sourceBook,
              prepared,
              source
            );


          aliases.forEach(function(
            alias
          ) {

            const sourceCastId =
              cfPayrollNormalizeExternalCastId_(
                alias.sourceCastId
              );


            const sourceCastName =
              String(
                alias.sourceCastName ||
                ""
              ).trim();


            if (!sourceCastId) {

              return;

            }


            const normalizedName =
              cfPayrollNormalizeCastName_(
                sourceCastName
              );


            /*
             * ID + 名前で別候補として管理。
             */
            const key =
              sourceCastId +
              "::" +
              normalizedName;


            if (
              !discovered[
                key
              ]
            ) {

              discovered[
                key
              ] = {

                sourceCastId:
                  sourceCastId,

                sourceCastName:
                  sourceCastName,

                files:
                  []

              };

            }


            discovered[
              key
            ].files.push(
              prepared.originalFileName
            );

          });

        }

      );

    }


    const newRows =
      [];


    const updateRows =
      [];


    let autoMatched =
      0;


    let needsReview =
      0;


    let unchanged =
      0;


    const now =
      new Date();


    Object.keys(
      discovered
    ).forEach(function(
      key
    ) {

      const alias =
        discovered[
          key
        ];


      const normalizedName =
        cfPayrollNormalizeCastName_(
          alias.sourceCastName
        );


      let matchedCastId =
        "";


      if (
        normalizedName &&
        castFlowNameMap[
          normalizedName
        ]
      ) {

        matchedCastId =
          castFlowNameMap[
            normalizedName
          ];

      }


      const existing =
        existingPosMap[
          key
        ];


      /* =================================================
       * 同じID + 名前が既に存在
       * ================================================= */

      if (
        existing
      ) {

        const finalCastId =
          existing.castId ||
          matchedCastId;


        const finalStatus =
          finalCastId
            ? (
                existing.castId
                  ? (
                      existing.status ||
                      "確認済"
                    )
                  : "自動照合"
              )
            : "要確認";


        const castIdChanged =
          existing.castId !==
          finalCastId;


        const statusChanged =
          existing.status !==
          finalStatus;


        if (
          !castIdChanged &&
          !statusChanged
        ) {

          unchanged++;

          return;

        }


        if (
          !existing.castId &&
          matchedCastId
        ) {

          autoMatched++;

        }


        if (
          !finalCastId
        ) {

          needsReview++;

        }


        updateRows.push({

          sheetRow:
            existing.sheetRow,

          values: [
            SOURCE_SYSTEM,
            alias.sourceCastId,
            alias.sourceCastName,
            finalCastId,
            finalStatus,
            now
          ]

        });


        return;

      }


      /* =================================================
       * 新しいID + 名前の組み合わせ
       * ================================================= */

      const status =
        matchedCastId
          ? "自動照合"
          : "要確認";


      if (
        matchedCastId
      ) {

        autoMatched++;

      } else {

        needsReview++;

      }


      newRows.push([
        SOURCE_SYSTEM,
        alias.sourceCastId,
        alias.sourceCastName,
        matchedCastId,
        status,
        now
      ]);

    });


    updateRows.forEach(function(
      item
    ) {

      mapSheet
        .getRange(
          item.sheetRow,
          1,
          1,
          6
        )
        .setValues([
          item.values
        ]);

    });


    if (
      newRows.length > 0
    ) {

      mapSheet
        .getRange(
          mapSheet.getLastRow() + 1,
          1,
          newRows.length,
          6
        )
        .setValues(
          newRows
        );

    }


    if (
      mapSheet.getLastRow() > 1
    ) {

      mapSheet
        .getRange(
          2,
          6,
          mapSheet.getLastRow() - 1,
          1
        )
        .setNumberFormat(
          "yyyy/mm/dd hh:mm:ss"
        );

    }


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      scannedFiles:
        scannedFiles,

      matchedFiles:
        matchedFiles,

      discoveredCasts:
        Object.keys(
          discovered
        ).length,

      added:
        newRows.length,

      updated:
        updateRows.length,

      autoMatched:
        autoMatched,

      needsReview:
        needsReview,

      unchanged:
        unchanged

    };


    Logger.log(
      "ROYAL POSキャスト同期OK: " +
      JSON.stringify(
        result
      )
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 22-B. CAST FLOW名前辞書
 *
 * 同名で別CAST_IDが存在する場合は
 * 自動照合しない。
 * ===================================================== */

function cfPayrollBuildCastFlowNameMap_(
  mapSheet
) {

  const result =
    {};


  const lastRow =
    mapSheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return result;

  }


  const rows =
    mapSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  rows.forEach(function(
    row
  ) {

    const sourceSystem =
      String(
        row[0] || ""
      ).trim();


    if (
      sourceSystem !==
      "CAST_FLOW"
    ) {

      return;

    }


    const sourceCastName =
      String(
        row[2] || ""
      ).trim();


    const castId =
      String(
        row[3] || ""
      ).trim();


    if (
      !sourceCastName ||
      !castId
    ) {

      return;

    }


    const normalizedName =
      cfPayrollNormalizeCastName_(
        sourceCastName
      );


    if (
      Object.prototype
        .hasOwnProperty.call(
          result,
          normalizedName
        ) &&
      result[
        normalizedName
      ] !== castId
    ) {

      result[
        normalizedName
      ] =
        null;

      return;

    }


    result[
      normalizedName
    ] =
      castId;

  });


  return result;

}


/* =====================================================
 * 22-C. 既存ROYAL_POSマッピング取得
 *
 * ID + 名前で管理。
 *
 * 0030 さやか
 * 0030 マジン
 *
 * は別エントリとして扱う。
 * ===================================================== */

function cfPayrollBuildExistingPosCastMap_(
  mapSheet,
  sourceSystem
) {

  const result =
    {};


  const lastRow =
    mapSheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return result;

  }


  const rows =
    mapSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  rows.forEach(function(
    row,
    index
  ) {

    const rowSourceSystem =
      String(
        row[0] || ""
      ).trim();


    if (
      rowSourceSystem !==
      sourceSystem
    ) {

      return;

    }


    const sourceCastId =
      cfPayrollNormalizeExternalCastId_(
        row[1]
      );


    const sourceCastName =
      String(
        row[2] || ""
      ).trim();


    if (
      !sourceCastId
    ) {

      return;

    }


    const normalizedName =
      cfPayrollNormalizeCastName_(
        sourceCastName
      );


    const key =
      sourceCastId +
      "::" +
      normalizedName;


    /*
     * 同じID + 同じ名前が
     * 二重登録されている場合だけエラー。
     *
     * 同じID + 別名は許可する。
     */
    if (
      result[
        key
      ]
    ) {

      throw new Error(
        "PAY_CAST_MAPに同一のROYAL_POSキャストが重複しています：" +
        sourceCastId +
        " / " +
        sourceCastName
      );

    }


    result[
      key
    ] = {

      sheetRow:
        index + 2,

      sourceCastName:
        sourceCastName,

      castId:
        String(
          row[3] || ""
        ).trim(),

      status:
        String(
          row[4] || ""
        ).trim()

    };

  });


  return result;

}


/* =====================================================
 * 22-D. Spreadsheetからキャスト一覧抽出
 *
 * parserModeごとに
 * 実際に意味のあるデータが存在する行だけ対象。
 *
 * PRODUCT_ROWS
 *   → キャスト付き商品行
 *
 * ATTENDANCE_ROWS
 *   → 勤怠情報あり
 *
 * PERFORMANCE_ROWS
 *   → 出勤または成績情報あり
 * ===================================================== */

function cfPayrollExtractCastAliasesFromSpreadsheet_(
  sourceBook,
  prepared,
  source
) {

  let sourceSheet;


  if (
    source.sheetName
  ) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / " +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (
      sheets.length === 0
    ) {

      return [];

    }


    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow <
      source.startRow ||
    lastColumn < 1
  ) {

    return [];

  }


  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  const castIdIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castIdHeader
    );


  const castNameIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castNameHeader
    );


  if (
    castIdIndex === -1 &&
    castNameIndex === -1
  ) {

    return [];

  }


  const parserMode =
    String(
      source.parserMode ||
      "GENERIC"
    )
      .trim()
      .toUpperCase();


  /* =================================================
   * 勤怠判定列
   * ================================================= */

  const attendanceIndexes =
    parserMode ===
    "ATTENDANCE_ROWS"
      ? [
          "出勤予定",
          "退勤予定",
          "欠区",
          "欠勤",
          "出勤",
          "退勤"
        ]
          .map(function(
            header
          ) {

            return cfPayrollFindConfiguredHeaderIndex_(
              headers,
              header
            );

          })
          .filter(function(
            index
          ) {

            return index !== -1;

          })
      : [];


  /* =================================================
   * 成績判定列
   * ================================================= */

  const performanceIndexes =
    parserMode ===
    "PERFORMANCE_ROWS"
      ? [
          "出勤数",
          "送り",
          "店服",
          "前払金",
          "指名数",
          "指名通",
          "指名優",
          "指名Ｐ",
          "延長数",
          "延長Ｐ",
          "同伴数",
          "同伴通",
          "同伴優",
          "同伴Ｐ",
          "ご新規Ｐ",
          "新規延長Ｐ",
          "場内数",
          "ボトルＰ",
          "合計Ｐ",
          "請求実績",
          "男子Ｍ",
          "男子Ｐ"
        ]
          .map(function(
            header
          ) {

            return cfPayrollFindConfiguredHeaderIndex_(
              headers,
              header
            );

          })
          .filter(function(
            index
          ) {

            return index !== -1;

          })
      : [];


  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getDisplayValues();


  const result =
    [];


  values.forEach(function(
    row
  ) {

    const sourceCastId =
      castIdIndex === -1
        ? ""
        : cfPayrollNormalizeExternalCastId_(
            row[
              castIdIndex
            ]
          );


    const sourceCastName =
      castNameIndex === -1
        ? ""
        : String(
            row[
              castNameIndex
            ] || ""
          ).trim();


    if (
      !sourceCastId &&
      !sourceCastName
    ) {

      return;

    }


    /* =================================================
     * ATTENDANCE_ROWS
     * ================================================= */

    if (
      parserMode ===
      "ATTENDANCE_ROWS"
    ) {

      const hasAttendanceData =
        attendanceIndexes.some(function(
          index
        ) {

          return (
            String(
              row[index] || ""
            ).trim() !== ""
          );

        });


      if (
        !hasAttendanceData
      ) {

        return;

      }

    }


    /* =================================================
     * PERFORMANCE_ROWS
     *
     * 0だけの在籍キャストは除外。
     * ================================================= */

    if (
      parserMode ===
      "PERFORMANCE_ROWS"
    ) {

      const hasPerformanceData =
        performanceIndexes.some(function(
          index
        ) {

          const value =
            String(
              row[index] || ""
            ).trim();


          if (!value) {

            return false;

          }


          const number =
            Number(
              value.replace(
                /,/g,
                ""
              )
            );


          if (
            Number.isFinite(
              number
            )
          ) {

            return number !== 0;

          }


          return true;

        });


      if (
        !hasPerformanceData
      ) {

        return;

      }

    }


    result.push({

      sourceCastId:
        sourceCastId,

      sourceCastName:
        sourceCastName

    });

  });


  return result;

}


/* =====================================================
 * 22-E. 外部キャストID正規化
 *
 * POSコード：
 *
 * 20
 * "20"
 * "0020"
 *
 * ↓
 *
 * "0020"
 *
 * 英数字ID：
 *
 * S001
 *
 * ↓
 *
 * "S001"
 * ===================================================== */

function cfPayrollNormalizeExternalCastId_(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return "";

  }


  let text =
    String(
      value
    ).trim();


  if (!text) {

    return "";

  }


  /*
   * Excel変換で
   * 20.0 のようになった場合
   */
  if (
    /^\d+\.0+$/.test(
      text
    )
  ) {

    text =
      String(
        parseInt(
          text,
          10
        )
      );

  }


  /*
   * 数字だけなら
   * POSコードとして4桁へ。
   */
  if (
    /^\d+$/.test(
      text
    )
  ) {

    return text.padStart(
      4,
      "0"
    );

  }


  return text;

}


/* =====================================================
 * 23. LINE登録後 CAST_ID自動紐付け
 *
 * PAY_CAST_MAP内の
 *
 * ROYAL_POS等の未紐付けキャスト
 *
 *   ↓
 *
 * CAST_FLOW側で正式CAST_IDが発行された後、
 * 名前が一意に一致したものだけ自動接続する。
 *
 * 安全設計：
 *
 * ・既存CAST_IDは上書きしない
 * ・同名キャストは自動接続しない
 * ・名前不一致は空欄のまま
 * ・手動確認が必要なものは「要確認」
 * ===================================================== */


/* =====================================================
 * 23-A. キャストID同期 一括入口
 *
 * 将来的には
 * LINE登録完了 / 管理者承認後に
 * この関数を呼べばよい。
 *
 * 処理順：
 *
 * 1. CAST FLOW → PAY_CAST_MAP同期
 * 2. 未紐付け外部キャストを正式CAST_IDへ接続
 * ===================================================== */

function syncPayrollCastIdentity() {

  const castFlowResult =
    syncPayrollCastMap();


  const linkResult =
    linkPayrollPendingCastMappings();


  const result = {

    success:
      true,

    castFlowSync:
      castFlowResult,

    externalLink:
      linkResult

  };


  Logger.log(
    "給与キャストID一括同期OK: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 23-B. 未紐付け外部キャストを自動接続
 *
 * CAST_FLOW以外で
 * CAST_IDが空欄の行を対象にする。
 *
 * 名前がCAST_FLOW側で一意に一致した場合のみ
 * CAST_IDを設定する。
 * ===================================================== */

function linkPayrollPendingCastMappings() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(
    30000
  );


  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const sheet =
      book.getSheetByName(
        "PAY_CAST_MAP"
      );


    if (!sheet) {

      throw new Error(
        "PAY_CAST_MAPがありません。"
      );

    }


    const headers = [
      "元システム",
      "元キャストID",
      "元キャスト名",
      "CAST_ID",
      "確認状態",
      "更新日時"
    ];


    if (
      !cfPayrollHeaderMatches_(
        sheet,
        headers
      )
    ) {

      throw new Error(
        "PAY_CAST_MAPの列構成が異なります。"
      );

    }


    /* =================================================
     * CAST FLOW側
     * 名前 → CAST_ID辞書
     * ================================================= */

    const castFlowResolver =
      cfPayrollBuildRegisteredCastFlowResolver_(
        sheet
      );


    const lastRow =
      sheet.getLastRow();


    if (
      lastRow <= 1
    ) {

      return {

        success:
          true,

        registeredCastFlow:
          0,

        pending:
          0,

        linked:
          0,

        ambiguous:
          0,

        unmatched:
          0

      };

    }


    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          6
        )
        .getValues();


    const now =
      new Date();


    let pending =
      0;


    let linked =
      0;


    let ambiguous =
      0;


    let unmatched =
      0;


    const linkedItems =
      [];


    const pendingItems =
      [];


    /* =================================================
     * 外部システム行を確認
     * ================================================= */

    rows.forEach(function(
      row,
      index
    ) {

      const sheetRow =
        index + 2;


      const sourceSystem =
        String(
          row[0] || ""
        ).trim();


      const sourceCastId =
        String(
          row[1] || ""
        ).trim();


      const sourceCastName =
        String(
          row[2] || ""
        ).trim();


      const existingCastId =
        String(
          row[3] || ""
        ).trim();


      const currentStatus =
        String(
          row[4] || ""
        ).trim();


      /* ---------------------------------------------
       * CAST_FLOW自身は対象外
       * --------------------------------------------- */

      if (
        sourceSystem ===
        "CAST_FLOW"
      ) {

        return;

      }


      /* ---------------------------------------------
       * CAST_ID設定済みは保護
       * --------------------------------------------- */

      if (
        existingCastId
      ) {

        return;

      }


      /*
       * 外部キャスト情報が
       * 完全空欄なら無視。
       */
      if (
        !sourceCastId &&
        !sourceCastName
      ) {

        return;

      }


      pending++;


      /* ---------------------------------------------
       * 名前なし
       * --------------------------------------------- */

      if (
        !sourceCastName
      ) {

        unmatched++;


        cfPayrollUpdatePendingCastStatus_(
          sheet,
          sheetRow,
          "要確認",
          now
        );


        pendingItems.push({

          sourceSystem:
            sourceSystem,

          sourceCastId:
            sourceCastId,

          sourceCastName:
            "",

          reason:
            "NAME_EMPTY"

        });


        return;

      }


      const normalizedName =
        cfPayrollNormalizeCastName_(
          sourceCastName
        );


      /* ---------------------------------------------
       * CAST FLOWに名前なし
       * --------------------------------------------- */

      if (
        !Object.prototype
          .hasOwnProperty.call(
            castFlowResolver.byName,
            normalizedName
          )
      ) {

        unmatched++;


        if (
          currentStatus !==
          "要確認"
        ) {

          cfPayrollUpdatePendingCastStatus_(
            sheet,
            sheetRow,
            "要確認",
            now
          );

        }


        pendingItems.push({

          sourceSystem:
            sourceSystem,

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          reason:
            "NOT_REGISTERED"

        });


        return;

      }


      /* ---------------------------------------------
       * 同名重複
       * --------------------------------------------- */

      if (
        castFlowResolver.byName[
          normalizedName
        ] === null
      ) {

        ambiguous++;


        cfPayrollUpdatePendingCastStatus_(
          sheet,
          sheetRow,
          "要確認",
          now
        );


        pendingItems.push({

          sourceSystem:
            sourceSystem,

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          reason:
            "AMBIGUOUS_NAME"

        });


        return;

      }


      /* ---------------------------------------------
       * 一意一致
       * --------------------------------------------- */

      const matchedCastId =
        castFlowResolver.byName[
          normalizedName
        ];


      sheet
        .getRange(
          sheetRow,
          4,
          1,
          3
        )
        .setValues([[
          matchedCastId,
          "自動照合",
          now
        ]]);


      linked++;


      linkedItems.push({

        sourceSystem:
          sourceSystem,

        sourceCastId:
          sourceCastId,

        sourceCastName:
          sourceCastName,

        castId:
          matchedCastId

      });

    });


    SpreadsheetApp.flush();


    const remainingPending =
      pending -
      linked;


    const result = {

      success:
        true,

      registeredCastFlow:
        castFlowResolver.registeredCount,

      pending:
        pending,

      linked:
        linked,

      remainingPending:
        remainingPending,

      ambiguous:
        ambiguous,

      unmatched:
        unmatched,

      linkedItems:
        linkedItems,

      pendingItems:
        pendingItems

    };


    Logger.log(
      "未紐付けキャスト自動照合OK: " +
      JSON.stringify(result)
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 23-C. 登録済みCAST FLOW辞書
 *
 * CAST_IDが存在するCAST_FLOW行だけ対象。
 *
 * 名前が同じでCAST_IDが違う場合は
 * nullとして自動照合禁止。
 * ===================================================== */

function cfPayrollBuildRegisteredCastFlowResolver_(
  sheet
) {

  const result = {

    byName: {},

    registeredCount:
      0

  };


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= 1
  ) {

    return result;

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  rows.forEach(function(
    row
  ) {

    const sourceSystem =
      String(
        row[0] || ""
      ).trim();


    if (
      sourceSystem !==
      "CAST_FLOW"
    ) {

      return;

    }


    const sourceCastName =
      String(
        row[2] || ""
      ).trim();


    const castId =
      String(
        row[3] || ""
      ).trim();


    /*
     * LINE登録等で正式CAST_IDが
     * 発行済みのキャストだけ対象。
     */
    if (
      !sourceCastName ||
      !castId
    ) {

      return;

    }


    result.registeredCount++;


    const normalizedName =
      cfPayrollNormalizeCastName_(
        sourceCastName
      );


    if (
      Object.prototype
        .hasOwnProperty.call(
          result.byName,
          normalizedName
        )
    ) {

      if (
        result.byName[
          normalizedName
        ] !== castId
      ) {

        /*
         * 同名別CAST_ID
         *
         * 自動紐付け禁止。
         */
        result.byName[
          normalizedName
        ] =
          null;

      }


      return;

    }


    result.byName[
      normalizedName
    ] =
      castId;

  });


  return result;

}


/* =====================================================
 * 23-D. 未紐付け行ステータス更新
 *
 * CAST_IDには触らない。
 * ===================================================== */

function cfPayrollUpdatePendingCastStatus_(
  sheet,
  sheetRow,
  status,
  updatedAt
) {

  sheet
    .getRange(
      sheetRow,
      5,
      1,
      2
    )
    .setValues([[
      status,
      updatedAt
    ]]);

}


/* =====================================================
 * 23-E. 未紐付けキャスト確認
 *
 * 読み取り専用。
 *
 * 現在まだCAST_IDが無い
 * 外部キャスト一覧を確認できる。
 * ===================================================== */

function previewPayrollPendingCastMappings() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_CAST_MAP"
    );


  if (!sheet) {

    throw new Error(
      "PAY_CAST_MAPがありません。"
    );

  }


  const lastRow =
    sheet.getLastRow();


  const pending =
    [];


  if (
    lastRow > 1
  ) {

    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          6
        )
        .getDisplayValues();


    rows.forEach(function(
      row,
      index
    ) {

      const sourceSystem =
        String(
          row[0] || ""
        ).trim();


      const sourceCastId =
        String(
          row[1] || ""
        ).trim();


      const sourceCastName =
        String(
          row[2] || ""
        ).trim();


      const castId =
        String(
          row[3] || ""
        ).trim();


      const status =
        String(
          row[4] || ""
        ).trim();


      if (
        sourceSystem ===
        "CAST_FLOW"
      ) {

        return;

      }


      if (
        castId
      ) {

        return;

      }


      if (
        !sourceCastId &&
        !sourceCastName
      ) {

        return;

      }


      pending.push({

        sheetRow:
          index + 2,

        sourceSystem:
          sourceSystem,

        sourceCastId:
          sourceCastId,

        sourceCastName:
          sourceCastName,

        status:
          status

      });

    });

  }


  const result = {

    success:
      true,

    pendingCount:
      pending.length,

    pending:
      pending

  };


  Logger.log(
    "未紐付け給与キャスト一覧: " +
    JSON.stringify(result)
  );


  return result;

}


/* =====================================================
 * 24. キャスト勤怠照会解析
 *
 * 実ファイル列：
 *
 * index
 * コード
 * キャスト名
 * カナ
 * 課
 * 出勤予定
 * 退勤予定
 * 欠区
 * 欠勤
 * 出勤
 * 退勤
 *
 * 給与計算に必要な
 *
 * ・実勤務時間
 * ・遅刻時間
 * ・早退時間
 * ・欠勤
 * ・当欠
 * ・無欠
 *
 * をPAY_FACTS用データへ変換する。
 *
 * 15分丸めはここでは行わない。
 * 生の勤務分数を保存する。
 * ===================================================== */


/* =====================================================
 * 24-A. キャスト勤怠照会解析
 * ===================================================== */

function cfPayrollAnalyzeAttendanceRows_(
  sourceBook,
  prepared,
  source
) {

  let sourceSheet;


  /* =================================================
   * 対象シート
   * ================================================= */

  if (
    source.sheetName
  ) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / ファイル：" +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (
      sheets.length === 0
    ) {

      throw new Error(
        "取込ファイルにシートがありません。"
      );

    }


    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow <
      source.startRow ||
    lastColumn < 1
  ) {

    throw new Error(
      "取込対象データがありません。" +
      " シート：" +
      sourceSheet.getName()
    );

  }


  /* =================================================
   * ヘッダー取得
   * ================================================= */

  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  /* =================================================
   * 列位置
   * ================================================= */

  const castIdIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castIdHeader ||
      "コード"
    );


  const castNameIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      source.castNameHeader ||
      "キャスト名"
    );


  const scheduledStartIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "出勤予定"
    );


  const scheduledEndIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "退勤予定"
    );


  const absenceCodeIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "欠区"
    );


  const absenceTypeIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "欠勤"
    );


  const actualStartIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "出勤"
    );


  const actualEndIndex =
    cfPayrollFindConfiguredHeaderIndex_(
      headers,
      "退勤"
    );


  /* =================================================
   * 必須列確認
   * ================================================= */

  const required = [

    [
      castIdIndex,
      source.castIdHeader ||
      "コード"
    ],

    [
      castNameIndex,
      source.castNameHeader ||
      "キャスト名"
    ],

    [
      scheduledStartIndex,
      "出勤予定"
    ],

    [
      scheduledEndIndex,
      "退勤予定"
    ],

    [
      absenceCodeIndex,
      "欠区"
    ],

    [
      absenceTypeIndex,
      "欠勤"
    ],

    [
      actualStartIndex,
      "出勤"
    ],

    [
      actualEndIndex,
      "退勤"
    ]

  ];


  required.forEach(function(
    item
  ) {

    if (
      item[0] === -1
    ) {

      throw new Error(
        "キャスト勤怠照会に必要な列がありません：" +
        item[1]
      );

    }

  });


  /* =================================================
   * データ取得
   *
   * displayValuesを使用して
   * 時刻を HH:mm の文字列として取得。
   * ================================================= */

  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getDisplayValues();


  const castResolver =
    cfPayrollBuildCastResolver_();


  const previewRows =
    [];


  const unknownCasts =
    {};


  const recognizedCastIds =
    {};


  const itemTotals =
    {};


  let skippedEmptyRows =
    0;


  let skippedNoAttendanceRows =
    0;


  let incompleteActualRows =
    0;


  let absenceRows =
    0;


  let workedRows =
    0;


  /* =================================================
   * 行解析
   * ================================================= */

  values.forEach(function(
    row,
    index
  ) {

    const sourceRowNumber =
      source.startRow +
      index;


    /* ---------------------------------------------
     * 完全空行
     * --------------------------------------------- */

    const hasAnyValue =
      row.some(function(
        value
      ) {

        return (
          String(
            value || ""
          ).trim() !== ""
        );

      });


    if (!hasAnyValue) {

      skippedEmptyRows++;

      return;

    }


    /* ---------------------------------------------
     * キャスト
     * --------------------------------------------- */

    const sourceCastId =
      cfPayrollNormalizeExternalCastId_(
        row[
          castIdIndex
        ]
      );


    const sourceCastName =
      String(
        row[
          castNameIndex
        ] || ""
      ).trim();


    /* ---------------------------------------------
     * 勤怠値
     * --------------------------------------------- */

    const scheduledStartRaw =
      String(
        row[
          scheduledStartIndex
        ] || ""
      ).trim();


    const scheduledEndRaw =
      String(
        row[
          scheduledEndIndex
        ] || ""
      ).trim();


    const absenceCodeRaw =
      String(
        row[
          absenceCodeIndex
        ] || ""
      ).trim();


    const absenceTypeRaw =
      String(
        row[
          absenceTypeIndex
        ] || ""
      ).trim();


    const actualStartRaw =
      String(
        row[
          actualStartIndex
        ] || ""
      ).trim();


    const actualEndRaw =
      String(
        row[
          actualEndIndex
        ] || ""
      ).trim();


    /* ---------------------------------------------
     * 勤怠情報なし
     *
     * 在籍しているだけで
     * 当日勤務していないキャストは除外。
     * --------------------------------------------- */

    const hasAttendanceData =
      !!(
        scheduledStartRaw ||
        scheduledEndRaw ||
        absenceCodeRaw ||
        absenceTypeRaw ||
        actualStartRaw ||
        actualEndRaw
      );


    if (
      !hasAttendanceData
    ) {

      skippedNoAttendanceRows++;

      return;

    }


    /* ---------------------------------------------
     * CAST_ID解決
     * --------------------------------------------- */

    const payrollCastId =
      cfPayrollResolveCastId_(
        castResolver,
        sourceCastId,
        sourceCastName
      );


    if (
      payrollCastId
    ) {

      recognizedCastIds[
        payrollCastId
      ] =
        true;

    } else {

      const unknownCastKey =
        sourceCastId +
        "::" +
        sourceCastName;


      if (
        !unknownCasts[
          unknownCastKey
        ]
      ) {

        unknownCasts[
          unknownCastKey
        ] = {

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          rows:
            []

        };

      }


      unknownCasts[
        unknownCastKey
      ].rows.push(
        sourceRowNumber
      );

    }


    /* ---------------------------------------------
     * 時刻 → 0:00からの分数
     * --------------------------------------------- */

    const scheduledStartMinutes =
      cfPayrollParseTimeMinutes_(
        scheduledStartRaw
      );


    const scheduledEndMinutes =
      cfPayrollParseTimeMinutes_(
        scheduledEndRaw
      );


    const actualStartMinutes =
      cfPayrollParseTimeMinutes_(
        actualStartRaw
      );


    const actualEndMinutes =
      cfPayrollParseTimeMinutes_(
        actualEndRaw
      );


    const facts =
      [];


    /* =================================================
     * 予定時刻
     * ================================================= */

    cfPayrollPushAttendanceFact_(
      facts,
      itemTotals,
      "SCHEDULED_START_MINUTES",
      "出勤予定",
      scheduledStartMinutes,
      "分",
      "出勤予定",
      false
    );


    cfPayrollPushAttendanceFact_(
      facts,
      itemTotals,
      "SCHEDULED_END_MINUTES",
      "退勤予定",
      scheduledEndMinutes,
      "分",
      "退勤予定",
      false
    );


    /* =================================================
     * 実時刻
     * ================================================= */

    cfPayrollPushAttendanceFact_(
      facts,
      itemTotals,
      "ACTUAL_START_MINUTES",
      "実出勤",
      actualStartMinutes,
      "分",
      "出勤",
      false
    );


    cfPayrollPushAttendanceFact_(
      facts,
      itemTotals,
      "ACTUAL_END_MINUTES",
      "実退勤",
      actualEndMinutes,
      "分",
      "退勤",
      false
    );


    /* =================================================
     * 実勤務時間
     *
     * 例：
     *
     * 09:48 ～ 13:00
     * ↓
     * 192分
     *
     * 日跨ぎにも対応。
     * ================================================= */

    if (
      actualStartMinutes !== null &&
      actualEndMinutes !== null
    ) {

      const workMinutes =
        cfPayrollDurationMinutes_(
          actualStartMinutes,
          actualEndMinutes
        );


      cfPayrollPushAttendanceFact_(
        facts,
        itemTotals,
        "WORK_MINUTES",
        "勤務時間",
        workMinutes,
        "分",
        "出勤-退勤"
      );


      workedRows++;

    }

    /*
     * 出勤か退勤の片方だけある場合
     */
    else if (
      actualStartRaw ||
      actualEndRaw
    ) {

      incompleteActualRows++;

    }


    /* =================================================
     * 遅刻時間
     *
     * 金額計算はしない。
     * 生の遅刻分数のみ保存。
     * ================================================= */

    if (
      scheduledStartMinutes !== null &&
      actualStartMinutes !== null
    ) {

      const lateMinutes =
        Math.max(
          0,
          cfPayrollSignedTimeDifference_(
            actualStartMinutes,
            scheduledStartMinutes
          )
        );


      cfPayrollPushAttendanceFact_(
        facts,
        itemTotals,
        "LATE_MINUTES",
        "遅刻時間",
        lateMinutes,
        "分",
        "出勤予定→実出勤"
      );

    }


    /* =================================================
     * 早退時間
     * ================================================= */

    if (
      scheduledEndMinutes !== null &&
      actualEndMinutes !== null
    ) {

      const earlyLeaveMinutes =
        Math.max(
          0,
          cfPayrollSignedTimeDifference_(
            scheduledEndMinutes,
            actualEndMinutes
          )
        );


      cfPayrollPushAttendanceFact_(
        facts,
        itemTotals,
        "EARLY_LEAVE_MINUTES",
        "早退時間",
        earlyLeaveMinutes,
        "分",
        "実退勤→退勤予定"
      );

    }


    /* =================================================
     * 欠勤
     * ================================================= */

    if (
      absenceCodeRaw ||
      absenceTypeRaw
    ) {

      absenceRows++;


      /* ---------------------------------------------
       * 欠勤1件
       * --------------------------------------------- */

      cfPayrollPushAttendanceFact_(
        facts,
        itemTotals,
        "ABSENCE_COUNT",
        "欠勤",
        1,
        "件",
        "欠勤"
      );


      /* ---------------------------------------------
       * 欠区
       * --------------------------------------------- */

      if (
        absenceCodeRaw
      ) {

        const absenceCodeNumber =
          Number(
            absenceCodeRaw
          );


        facts.push({

          itemKey:
            "ABSENCE_CODE",

          displayName:
            "欠勤区分コード",

          rawValue:
            absenceCodeRaw,

          value:
            Number.isFinite(
              absenceCodeNumber
            )
              ? absenceCodeNumber
              : absenceCodeRaw,

          unit:
            "コード",

          multiplier:
            1,

          sourceHeader:
            "欠区"

        });

      }


      /* ---------------------------------------------
       * 欠勤種別
       * --------------------------------------------- */

      if (
        absenceTypeRaw
      ) {

        facts.push({

          itemKey:
            "ABSENCE_TYPE",

          displayName:
            "欠勤区分",

          rawValue:
            absenceTypeRaw,

          value:
            absenceTypeRaw,

          unit:
            "区分",

          multiplier:
            1,

          sourceHeader:
            "欠勤"

        });


        /* -------------------------------------------
         * 当欠
         *
         * PAY_RULES
         * PENALTY_SAME_DAY_ABSENCE
         * と後で接続。
         * ------------------------------------------- */

        if (
          absenceTypeRaw ===
          "当欠"
        ) {

          cfPayrollPushAttendanceFact_(
            facts,
            itemTotals,
            "PENALTY_SAME_DAY_ABSENCE",
            "当日欠勤",
            1,
            "件",
            "欠勤"
          );

        }


        /* -------------------------------------------
         * 無欠
         *
         * 将来ファイルに出た場合にも対応。
         * ------------------------------------------- */

        if (
          absenceTypeRaw ===
          "無欠"
        ) {

          cfPayrollPushAttendanceFact_(
            facts,
            itemTotals,
            "PENALTY_NO_SHOW",
            "無断欠勤",
            1,
            "件",
            "欠勤"
          );

        }

      }

    }


    /* =================================================
     * プレビュー行
     * ================================================= */

    previewRows.push({

      sourceRow:
        sourceRowNumber,

      /*
       * Section 19で
       * ファイル名から営業日を設定
       */
      date:
        "",

      sourceCastId:
        sourceCastId,

      sourceCastName:
        sourceCastName,

      castId:
        payrollCastId,

      scheduledStart:
        scheduledStartRaw,

      scheduledEnd:
        scheduledEndRaw,

      actualStart:
        actualStartRaw,

      actualEnd:
        actualEndRaw,

      absenceCode:
        absenceCodeRaw,

      absenceType:
        absenceTypeRaw,

      facts:
        facts

    });

  });


  /* =================================================
   * 想定ヘッダー
   *
   * index / カナ / 課は
   * 給与計算には使わないが既知列として扱う。
   * ================================================= */

  const knownHeaders = {

    "index":
      true,

    "コード":
      true,

    "キャスト名":
      true,

    "カナ":
      true,

    "課":
      true,

    "出勤予定":
      true,

    "退勤予定":
      true,

    "欠区":
      true,

    "欠勤":
      true,

    "出勤":
      true,

    "退勤":
      true

  };


  const unmappedHeaders =
    headers.filter(function(
      header
    ) {

      return (
        header &&
        !knownHeaders[
          header
        ]
      );

    });


  /* =================================================
   * 結果
   * ================================================= */

  const result = {

    success:
      true,

    mode:
      "PREVIEW_ONLY",


    file: {

      id:
        prepared.originalFileId,

      name:
        prepared.originalFileName,

      type:
        prepared.originalFileType,

      temporaryConversion:
        prepared.temporary,

      sheetName:
        sourceSheet.getName()

    },


    importSource: {

      sourceId:
        source.sourceId,

      displayName:
        source.displayName,

      dataType:
        source.dataType,

      parserMode:
        "ATTENDANCE_ROWS"

    },


    summary: {

      sourceRows:
        values.length,

      recognizedRows:
        previewRows.length,

      skippedEmptyRows:
        skippedEmptyRows,

      skippedNoAttendanceRows:
        skippedNoAttendanceRows,

      blankDateRows:
        0,

      mappedColumns:
        8,

      unmappedColumns:
        unmappedHeaders.length,

      recognizedCasts:
        Object.keys(
          recognizedCastIds
        ).length,

      unknownCastCount:
        Object.keys(
          unknownCasts
        ).length,

      workedRows:
        workedRows,

      absenceRows:
        absenceRows,

      incompleteActualRows:
        incompleteActualRows

    },


    mappedColumns: [

      {
        sourceHeader:
          "コード",

        role:
          "CAST_ID"
      },

      {
        sourceHeader:
          "キャスト名",

        role:
          "CAST_NAME"
      },

      {
        sourceHeader:
          "出勤予定",

        role:
          "SCHEDULED_START"
      },

      {
        sourceHeader:
          "退勤予定",

        role:
          "SCHEDULED_END"
      },

      {
        sourceHeader:
          "欠区",

        role:
          "ABSENCE_CODE"
      },

      {
        sourceHeader:
          "欠勤",

        role:
          "ABSENCE_TYPE"
      },

      {
        sourceHeader:
          "出勤",

        role:
          "ACTUAL_START"
      },

      {
        sourceHeader:
          "退勤",

        role:
          "ACTUAL_END"
      }

    ],


    ignoredHeaders: [
      "index",
      "カナ",
      "課"
    ],


    unmappedHeaders:
      unmappedHeaders,


    itemTotals:
      itemTotals,


    unknownCasts:
      Object.keys(
        unknownCasts
      ).map(function(
        key
      ) {

        return unknownCasts[
          key
        ];

      }),


    rows:
      previewRows

  };


  Logger.log(
    "キャスト勤怠解析OK: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 24-B. 時刻 → 分数
 *
 * 09:48
 * ↓
 * 588
 *
 * 8:00にも対応。
 * ===================================================== */

function cfPayrollParseTimeMinutes_(
  value
) {

  const text =
    String(
      value || ""
    ).trim();


  if (
    !text
  ) {

    return null;

  }


  const match =
    text.match(
      /^(\d{1,2}):(\d{2})$/
    );


  if (
    !match
  ) {

    return null;

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
    !Number.isInteger(
      hour
    ) ||
    !Number.isInteger(
      minute
    ) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {

    return null;

  }


  return (
    hour * 60 +
    minute
  );

}


/* =====================================================
 * 24-C. 勤務分数
 *
 * 日跨ぎ対応。
 *
 * 23:00 → 02:00
 * ↓
 * 180分
 * ===================================================== */

function cfPayrollDurationMinutes_(
  startMinutes,
  endMinutes
) {

  if (
    startMinutes === null ||
    endMinutes === null
  ) {

    return null;

  }


  let duration =
    endMinutes -
    startMinutes;


  if (
    duration < 0
  ) {

    duration +=
      1440;

  }


  return duration;

}


/* =====================================================
 * 24-D. 時刻差
 *
 * 深夜跨ぎでも
 * 遅刻・早退を正しく比較する。
 * ===================================================== */

function cfPayrollSignedTimeDifference_(
  leftMinutes,
  rightMinutes
) {

  let difference =
    leftMinutes -
    rightMinutes;


  while (
    difference > 720
  ) {

    difference -=
      1440;

  }


  while (
    difference < -720
  ) {

    difference +=
      1440;

  }


  return difference;

}


/* =====================================================
 * 24-E. 勤怠fact追加
 *
 * aggregate=falseの場合、
 * itemTotalsには合計しない。
 *
 * 出勤時刻などは合計しても意味がないため
 * falseを使用する。
 * ===================================================== */

function cfPayrollPushAttendanceFact_(
  facts,
  itemTotals,
  itemKey,
  displayName,
  value,
  unit,
  sourceHeader,
  aggregate
) {

  if (
    value === null ||
    value === "" ||
    value === undefined
  ) {

    return;

  }


  facts.push({

    itemKey:
      itemKey,

    displayName:
      displayName,

    rawValue:
      value,

    value:
      value,

    unit:
      unit,

    multiplier:
      1,

    sourceHeader:
      sourceHeader

  });


  if (
    aggregate === false ||
    typeof value !==
      "number"
  ) {

    return;

  }


  if (
    !Object.prototype
      .hasOwnProperty.call(
        itemTotals,
        itemKey
      )
  ) {

    itemTotals[
      itemKey
    ] = {

      displayName:
        displayName,

      unit:
        unit,

      total:
        0

    };

  }


  itemTotals[
    itemKey
  ].total +=
    value;

}


/* =====================================================
 * 24-F. 時刻解析テスト
 * ===================================================== */

function testPayrollAttendanceTimeParser() {

  const cases = [

    [
      "09:48",
      588
    ],

    [
      "8:00",
      480
    ],

    [
      "23:45",
      1425
    ],

    [
      "",
      null
    ],

    [
      "LAST",
      null
    ]

  ];


  const results =
    cases.map(function(
      item
    ) {

      const actual =
        cfPayrollParseTimeMinutes_(
          item[0]
        );


      return {

        input:
          item[0],

        expected:
          item[1],

        actual:
          actual,

        success:
          actual ===
          item[1]

      };

    });


  const success =
    results.every(function(
      item
    ) {

      return item.success;

    });


  const result = {

    success:
      success,

    cases:
      results

  };


  Logger.log(
    "勤怠時間解析テスト: " +
    JSON.stringify(
      result
    )
  );


  if (
    !success
  ) {

    throw new Error(
      "勤怠時間解析テストに失敗しました。"
    );

  }


  return result;

}


/* =====================================================
 * 22-F. 誤登録された直近90行の安全復旧
 *
 * Section 22-D修正前に
 * 勤怠ファイルの在籍者全員を拾ってしまい、
 * 90行追加されたケース専用。
 *
 * 安全条件：
 *
 * ・末尾90行すべてROYAL_POS
 * ・CAST_IDが空欄
 * ・確認状態が「要確認」
 *
 * 条件が1件でも違えば
 * 一切削除せずエラー停止。
 *
 * 一度だけ実行する。
 * ===================================================== */

function cleanupPayrollAccidentalPosSync() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(
    30000
  );


  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const sheet =
      book.getSheetByName(
        "PAY_CAST_MAP"
      );


    if (!sheet) {

      throw new Error(
        "PAY_CAST_MAPがありません。"
      );

    }


    const ROWS_TO_CLEAN =
      90;


    const lastRow =
      sheet.getLastRow();


    if (
      lastRow <=
      ROWS_TO_CLEAN
    ) {

      throw new Error(
        "PAY_CAST_MAPの行数が不足しています。" +
        " 復旧処理は実行していません。"
      );

    }


    const startRow =
      lastRow -
      ROWS_TO_CLEAN +
      1;


    const rows =
      sheet
        .getRange(
          startRow,
          1,
          ROWS_TO_CLEAN,
          6
        )
        .getValues();


    /* =================================================
     * 全90行を安全確認
     * ================================================= */

    rows.forEach(function(
      row,
      index
    ) {

      const sourceSystem =
        String(
          row[0] || ""
        ).trim();


      const castId =
        String(
          row[3] || ""
        ).trim();


      const status =
        String(
          row[4] || ""
        ).trim();


      if (
        sourceSystem !==
        "ROYAL_POS"
      ) {

        throw new Error(
          "復旧中止：" +
          (startRow + index) +
          "行目がROYAL_POSではありません。"
        );

      }


      if (
        castId
      ) {

        throw new Error(
          "復旧中止：" +
          (startRow + index) +
          "行目にはCAST_IDが入っています。"
        );

      }


      if (
        status !==
        "要確認"
      ) {

        throw new Error(
          "復旧中止：" +
          (startRow + index) +
          "行目の確認状態が要確認ではありません。"
        );

      }

    });


    /* =================================================
     * 条件一致した場合のみクリア
     *
     * 行そのものは削除せず
     * データだけ消す。
     * ================================================= */

    sheet
      .getRange(
        startRow,
        1,
        ROWS_TO_CLEAN,
        6
      )
      .clearContent();


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      cleanedRows:
        ROWS_TO_CLEAN,

      startRow:
        startRow,

      endRow:
        lastRow

    };


    Logger.log(
      "誤登録POSキャスト復旧OK: " +
      JSON.stringify(
        result
      )
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 25. キャスト成績一覧照会解析
 *
 * 主に以下を取得：
 *
 * ・送り
 * ・店服
 * ・前払金
 * ・指名実績
 * ・延長実績
 * ・同伴実績
 * ・場内実績
 * ・ポイント
 * ・請求実績
 *
 * 注意：
 *
 * VIP指名はこのファイルだけでは確定しない。
 * 後の伝票一覧解析で席種を確認して処理する。
 *
 * 遅刻・欠勤等は勤怠照会を正とするため、
 * このファイルからは給与factへ重複登録しない。
 * ===================================================== */


/* =====================================================
 * 25-A. キャスト成績一覧解析
 * ===================================================== */

function cfPayrollAnalyzePerformanceRows_(
  sourceBook,
  prepared,
  source
) {

  let sourceSheet;


  if (
    source.sheetName
  ) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / ファイル：" +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (
      sheets.length === 0
    ) {

      throw new Error(
        "取込ファイルにシートがありません。"
      );

    }


    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow <
      source.startRow ||
    lastColumn < 1
  ) {

    throw new Error(
      "取込対象データがありません。"
    );

  }


  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  /* =================================================
   * 必要列
   * ================================================= */

  const requiredHeaders = [

    source.castIdHeader ||
      "コード",

    source.castNameHeader ||
      "キャスト名",

    "出勤数",
    "送り",
    "店服",
    "前払金",

    "指名数",
    "指名通",
    "指名優",
    "指名Ｐ",

    "延長数",
    "延長Ｐ",

    "同伴数",
    "同伴通",
    "同伴優",
    "同伴Ｐ",

    "ご新規Ｐ",
    "新規延長Ｐ",

    "場内数",

    "ボトルＰ",
    "合計Ｐ",

    "請求実績",

    "男子Ｍ",
    "男子Ｐ"

  ];


  const indexes =
    {};


  requiredHeaders.forEach(function(
    header
  ) {

    const index =
      cfPayrollFindConfiguredHeaderIndex_(
        headers,
        header
      );


    if (
      index === -1
    ) {

      throw new Error(
        "キャスト成績一覧照会に必要な列がありません：" +
        header
      );

    }


    indexes[
      header
    ] =
      index;

  });


  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getValues();


  const castResolver =
    cfPayrollBuildCastResolver_();


  const previewRows =
    [];


  const unknownCasts =
    {};


  const recognizedCastIds =
    {};


  const itemTotals =
    {};


  let skippedEmptyRows =
    0;


  let skippedNoPerformanceRows =
    0;


  /* =================================================
   * 給与用項目定義
   * ================================================= */

  const metricDefinitions = [

    {
      header:
        "送り",

      itemKey:
        "SEND_AMOUNT",

      displayName:
        "送り",

      unit:
        "円"
    },

    {
      header:
        "店服",

      itemKey:
        "DRESS_RENTAL_AMOUNT",

      displayName:
        "店服",

      unit:
        "円"
    },

    {
      header:
        "前払金",

      itemKey:
        "ADVANCE_PAYMENT",

      displayName:
        "前払金",

      unit:
        "円"
    },

    {
      header:
        "指名数",

      itemKey:
        "PERFORMANCE_NOMINATION_COUNT",

      displayName:
        "指名数",

      unit:
        "件"
    },

    {
      header:
        "指名通",

      itemKey:
        "PERFORMANCE_NOMINATION_NORMAL_COUNT",

      displayName:
        "指名通",

      unit:
        "件"
    },

    {
      header:
        "指名優",

      itemKey:
        "PERFORMANCE_NOMINATION_PRIORITY_COUNT",

      displayName:
        "指名優",

      unit:
        "件"
    },

    {
      header:
        "指名Ｐ",

      itemKey:
        "PERFORMANCE_NOMINATION_POINTS",

      displayName:
        "指名P",

      unit:
        "P"
    },

    {
      header:
        "延長数",

      itemKey:
        "PERFORMANCE_EXTENSION_COUNT",

      displayName:
        "延長数",

      unit:
        "件"
    },

    {
      header:
        "延長Ｐ",

      itemKey:
        "PERFORMANCE_EXTENSION_POINTS",

      displayName:
        "延長P",

      unit:
        "P"
    },

    {
      header:
        "同伴数",

      itemKey:
        "PERFORMANCE_ACCOMPANIMENT_COUNT",

      displayName:
        "同伴数",

      unit:
        "件"
    },

    {
      header:
        "同伴通",

      itemKey:
        "PERFORMANCE_ACCOMPANIMENT_NORMAL_COUNT",

      displayName:
        "同伴通",

      unit:
        "件"
    },

    {
      header:
        "同伴優",

      itemKey:
        "PERFORMANCE_ACCOMPANIMENT_PRIORITY_COUNT",

      displayName:
        "同伴優",

      unit:
        "件"
    },

    {
      header:
        "同伴Ｐ",

      itemKey:
        "PERFORMANCE_ACCOMPANIMENT_POINTS",

      displayName:
        "同伴P",

      unit:
        "P"
    },

    {
      header:
        "ご新規Ｐ",

      itemKey:
        "PERFORMANCE_NEW_CUSTOMER_POINTS",

      displayName:
        "ご新規P",

      unit:
        "P"
    },

    {
      header:
        "新規延長Ｐ",

      itemKey:
        "PERFORMANCE_NEW_EXTENSION_POINTS",

      displayName:
        "新規延長P",

      unit:
        "P"
    },

    {
      header:
        "場内数",

      itemKey:
        "PERFORMANCE_INHOUSE_COUNT",

      displayName:
        "場内数",

      unit:
        "件"
    },

    {
      header:
        "ボトルＰ",

      itemKey:
        "PERFORMANCE_BOTTLE_POINTS",

      displayName:
        "ボトルP",

      unit:
        "P"
    },

    {
      header:
        "合計Ｐ",

      itemKey:
        "PERFORMANCE_TOTAL_POINTS",

      displayName:
        "合計P",

      unit:
        "P"
    },

    {
      header:
        "請求実績",

      itemKey:
        "SALES_TOTAL",

      displayName:
        "請求実績",

      unit:
        "円"
    },

    {
      header:
        "男子Ｍ",

      itemKey:
        "PERFORMANCE_MALE_M",

      displayName:
        "男子M",

      unit:
        "件"
    },

    {
      header:
        "男子Ｐ",

      itemKey:
        "PERFORMANCE_MALE_POINTS",

      displayName:
        "男子P",

      unit:
        "P"
    }

  ];


  values.forEach(function(
    row,
    index
  ) {

    const sourceRowNumber =
      source.startRow +
      index;


    const hasAnyValue =
      row.some(function(
        value
      ) {

        return (
          value !== "" &&
          value !== null
        );

      });


    if (
      !hasAnyValue
    ) {

      skippedEmptyRows++;

      return;

    }


    const sourceCastId =
      cfPayrollNormalizeExternalCastId_(
        row[
          indexes[
            source.castIdHeader ||
            "コード"
          ]
        ]
      );


    const sourceCastName =
      String(
        row[
          indexes[
            source.castNameHeader ||
            "キャスト名"
          ]
        ] || ""
      ).trim();


    /* =================================================
     * 実績有無
     *
     * 出勤数または給与関連数値が
     * 1つでも0以外なら対象。
     * ================================================= */

    const attendanceCount =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "出勤数"
          ]
        ]
      );


    const hasMetric =
      metricDefinitions.some(function(
        definition
      ) {

        const value =
          cfPayrollPerformanceNumber_(
            row[
              indexes[
                definition.header
              ]
            ]
          );


        return (
          value !== null &&
          value !== 0
        );

      });


    if (
      (!attendanceCount ||
       attendanceCount === 0) &&
      !hasMetric
    ) {

      skippedNoPerformanceRows++;

      return;

    }


    /* =================================================
     * CAST_ID
     * ================================================= */

    const payrollCastId =
      cfPayrollResolveCastId_(
        castResolver,
        sourceCastId,
        sourceCastName
      );


    if (
      payrollCastId
    ) {

      recognizedCastIds[
        payrollCastId
      ] =
        true;

    } else {

      const key =
        sourceCastId +
        "::" +
        sourceCastName;


      if (
        !unknownCasts[
          key
        ]
      ) {

        unknownCasts[
          key
        ] = {

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          rows:
            []

        };

      }


      unknownCasts[
        key
      ].rows.push(
        sourceRowNumber
      );

    }


    const facts =
      [];


    /* =================================================
     * 出勤数
     *
     * 勤務時間自体は勤怠照会を正とする。
     * ここでは監査用出勤回数のみ。
     * ================================================= */

    if (
      attendanceCount !== null &&
      attendanceCount !== 0
    ) {

      cfPayrollPushPerformanceFact_(
        facts,
        itemTotals,
        "PERFORMANCE_ATTENDANCE_COUNT",
        "出勤数",
        attendanceCount,
        "日",
        "出勤数"
      );

    }


    /* =================================================
     * 各成績
     * ================================================= */

    metricDefinitions.forEach(function(
      definition
    ) {

      const value =
        cfPayrollPerformanceNumber_(
          row[
            indexes[
              definition.header
            ]
          ]
        );


      /*
       * 0は保存しない。
       * PAY_FACTSを無駄に膨らませない。
       */
      if (
        value === null ||
        value === 0
      ) {

        return;

      }


      cfPayrollPushPerformanceFact_(
        facts,
        itemTotals,
        definition.itemKey,
        definition.displayName,
        value,
        definition.unit,
        definition.header
      );

    });


    previewRows.push({

      sourceRow:
        sourceRowNumber,

      date:
        "",

      sourceCastId:
        sourceCastId,

      sourceCastName:
        sourceCastName,

      castId:
        payrollCastId,

      facts:
        facts

    });

  });


  const result = {

    success:
      true,

    mode:
      "PREVIEW_ONLY",


    file: {

      id:
        prepared.originalFileId,

      name:
        prepared.originalFileName,

      type:
        prepared.originalFileType,

      temporaryConversion:
        prepared.temporary,

      sheetName:
        sourceSheet.getName()

    },


    importSource: {

      sourceId:
        source.sourceId,

      displayName:
        source.displayName,

      dataType:
        source.dataType,

      parserMode:
        "PERFORMANCE_ROWS"

    },


    summary: {

      sourceRows:
        values.length,

      recognizedRows:
        previewRows.length,

      skippedEmptyRows:
        skippedEmptyRows,

      skippedNoPerformanceRows:
        skippedNoPerformanceRows,

      blankDateRows:
        0,

      mappedColumns:
        requiredHeaders.length,

      unmappedColumns:
        0,

      recognizedCasts:
        Object.keys(
          recognizedCastIds
        ).length,

      unknownCastCount:
        Object.keys(
          unknownCasts
        ).length

    },


    itemTotals:
      itemTotals,


    unknownCasts:
      Object.keys(
        unknownCasts
      ).map(function(
        key
      ) {

        return unknownCasts[
          key
        ];

      }),


    rows:
      previewRows

  };


  Logger.log(
    "キャスト成績解析OK: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 25-B. 成績数値変換
 * ===================================================== */

function cfPayrollPerformanceNumber_(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return null;

  }


  if (
    typeof value ===
    "number"
  ) {

    return Number.isFinite(
      value
    )
      ? value
      : null;

  }


  const text =
    String(
      value
    )
      .trim()
      .replace(
        /,/g,
        ""
      );


  if (!text) {

    return null;

  }


  const number =
    Number(
      text
    );


  return Number.isFinite(
    number
  )
    ? number
    : null;

}


/* =====================================================
 * 25-C. 成績fact追加
 * ===================================================== */

function cfPayrollPushPerformanceFact_(
  facts,
  itemTotals,
  itemKey,
  displayName,
  value,
  unit,
  sourceHeader
) {

  facts.push({

    itemKey:
      itemKey,

    displayName:
      displayName,

    rawValue:
      value,

    value:
      value,

    unit:
      unit,

    multiplier:
      1,

    sourceHeader:
      sourceHeader

  });


  if (
    typeof value !==
    "number"
  ) {

    return;

  }


  if (
    !Object.prototype
      .hasOwnProperty.call(
        itemTotals,
        itemKey
      )
  ) {

    itemTotals[
      itemKey
    ] = {

      displayName:
        displayName,

      unit:
        unit,

      total:
        0

    };

  }


  itemTotals[
    itemKey
  ].total +=
    value;

}


/* =====================================================
 * 26. 伝票一覧照会解析
 *
 * このファイルでは主に
 *
 * ・VIP指名数
 * ・ボトル実績
 *
 * を給与計算用として取得。
 *
 * さらに検算用として
 *
 * ・請求実績
 * ・純売実績
 * ・F&D実績
 *
 * も保存する。
 *
 * 注意：
 *
 * 場内数はキャスト成績一覧照会を正とする。
 * 通常指名数もキャスト成績一覧照会を正とする。
 *
 * 伝票一覧の指名数は
 * 「席種がVIPかどうか」の判定に使用する。
 * ===================================================== */


/* =====================================================
 * 26-A. 伝票一覧解析
 * ===================================================== */

function cfPayrollAnalyzeSlipRows_(
  sourceBook,
  prepared,
  source
) {

  let sourceSheet;


  /* =================================================
   * 対象シート
   * ================================================= */

  if (
    source.sheetName
  ) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / ファイル：" +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (
      sheets.length === 0
    ) {

      throw new Error(
        "取込ファイルにシートがありません。"
      );

    }


    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow <
      source.startRow ||
    lastColumn < 1
  ) {

    throw new Error(
      "取込対象データがありません。"
    );

  }


  /* =================================================
   * ヘッダー
   * ================================================= */

  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  /* =================================================
   * 使用列
   * ================================================= */

  const requiredHeaders = [

    source.castIdHeader ||
      "コード",

    source.castNameHeader ||
      "キャスト名",

    "伝票",
    "行",
    "伝票状況",
    "席種",

    "指名数",

    "請求実績",
    "純売実績",
    "ボトル実績",
    "Ｆ＆Ｄ実績",

    "分配数"

  ];


  const indexes =
    {};


  requiredHeaders.forEach(function(
    header
  ) {

    const index =
      cfPayrollFindConfiguredHeaderIndex_(
        headers,
        header
      );


    if (
      index === -1
    ) {

      throw new Error(
        "伝票一覧照会に必要な列がありません：" +
        header
      );

    }


    indexes[
      header
    ] =
      index;

  });


  /* =================================================
   * データ
   * ================================================= */

  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getValues();


  const castResolver =
    cfPayrollBuildCastResolver_();


  const previewRows =
    [];


  const unknownCasts =
    {};


  const recognizedCastIds =
    {};


  const itemTotals =
    {};


  let skippedEmptyRows =
    0;


  let skippedNoCastRows =
    0;


  let vipRows =
    0;


  let bottleSalesRows =
    0;


  /* =================================================
   * 行解析
   * ================================================= */

  values.forEach(function(
    row,
    index
  ) {

    const sourceRowNumber =
      source.startRow +
      index;


    /* ---------------------------------------------
     * 完全空行
     * --------------------------------------------- */

    const hasAnyValue =
      row.some(function(
        value
      ) {

        return (
          value !== "" &&
          value !== null
        );

      });


    if (
      !hasAnyValue
    ) {

      skippedEmptyRows++;

      return;

    }


    /* ---------------------------------------------
     * キャスト
     * --------------------------------------------- */

    const sourceCastId =
      cfPayrollNormalizeExternalCastId_(
        row[
          indexes[
            source.castIdHeader ||
            "コード"
          ]
        ]
      );


    const sourceCastName =
      String(
        row[
          indexes[
            source.castNameHeader ||
            "キャスト名"
          ]
        ] || ""
      ).trim();


    /*
     * 伝票にはキャストなしの追加行もある。
     *
     * その行は給与対象にしない。
     */
    if (
      !sourceCastId &&
      !sourceCastName
    ) {

      skippedNoCastRows++;

      return;

    }


    /* ---------------------------------------------
     * CAST_ID
     * --------------------------------------------- */

    const payrollCastId =
      cfPayrollResolveCastId_(
        castResolver,
        sourceCastId,
        sourceCastName
      );


    if (
      payrollCastId
    ) {

      recognizedCastIds[
        payrollCastId
      ] =
        true;

    } else {

      const unknownKey =
        sourceCastId +
        "::" +
        sourceCastName;


      if (
        !unknownCasts[
          unknownKey
        ]
      ) {

        unknownCasts[
          unknownKey
        ] = {

          sourceCastId:
            sourceCastId,

          sourceCastName:
            sourceCastName,

          rows:
            []

        };

      }


      unknownCasts[
        unknownKey
      ].rows.push(
        sourceRowNumber
      );

    }


    /* =================================================
     * 伝票情報
     * ================================================= */

    const slipNumber =
      String(
        row[
          indexes[
            "伝票"
          ]
        ] || ""
      ).trim();


    const slipLine =
      String(
        row[
          indexes[
            "行"
          ]
        ] || ""
      ).trim();


    const slipStatus =
      String(
        row[
          indexes[
            "伝票状況"
          ]
        ] || ""
      ).trim();


    const seatTypeRaw =
      String(
        row[
          indexes[
            "席種"
          ]
        ] || ""
      ).trim();


    const seatType =
      cfPayrollNormalizeSeatType_(
        seatTypeRaw
      );


    const nominationCount =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "指名数"
          ]
        ]
      );


    const billingSales =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "請求実績"
          ]
        ]
      );


    const netSales =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "純売実績"
          ]
        ]
      );


    const bottleSales =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "ボトル実績"
          ]
        ]
      );


    const foodDrinkSales =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "Ｆ＆Ｄ実績"
          ]
        ]
      );


    const distributionCount =
      cfPayrollPerformanceNumber_(
        row[
          indexes[
            "分配数"
          ]
        ]
      );


    const facts =
      [];


    /* =================================================
     * VIP指名
     *
     * 席種がVIPで、
     * 指名数が1以上の場合のみ。
     *
     * 通常指名は成績一覧を正とするため
     * ここでは登録しない。
     * ================================================= */

    if (
      seatType ===
        "VIP" &&
      nominationCount !== null &&
      nominationCount > 0
    ) {

      cfPayrollPushSlipFact_(
        facts,
        itemTotals,
        "VIP_NOMINATION_COUNT",
        "VIP指名数",
        nominationCount,
        "件",
        "席種+指名数"
      );


      vipRows++;

    }


    /* =================================================
     * ボトル実績
     *
     * 売上制給与計算用の正式キー。
     * ================================================= */

    if (
      bottleSales !== null &&
      bottleSales !== 0
    ) {

      cfPayrollPushSlipFact_(
        facts,
        itemTotals,
        "BOTTLE_SALES",
        "ボトル売上",
        bottleSales,
        "円",
        "ボトル実績"
      );


      bottleSalesRows++;

    }


    /* =================================================
     * 検算用
     *
     * PERFORMANCE側のSALES_TOTALとは
     * 別キーにすることで二重計上を防ぐ。
     * ================================================= */

    if (
      billingSales !== null &&
      billingSales !== 0
    ) {

      cfPayrollPushSlipFact_(
        facts,
        itemTotals,
        "SLIP_BILLING_SALES",
        "伝票請求実績",
        billingSales,
        "円",
        "請求実績"
      );

    }


    if (
      netSales !== null &&
      netSales !== 0
    ) {

      cfPayrollPushSlipFact_(
        facts,
        itemTotals,
        "SLIP_NET_SALES",
        "伝票純売実績",
        netSales,
        "円",
        "純売実績"
      );

    }


    if (
      foodDrinkSales !== null &&
      foodDrinkSales !== 0
    ) {

      cfPayrollPushSlipFact_(
        facts,
        itemTotals,
        "SLIP_FD_SALES",
        "伝票F&D実績",
        foodDrinkSales,
        "円",
        "Ｆ＆Ｄ実績"
      );

    }


    if (
      distributionCount !== null &&
      distributionCount !== 0
    ) {

      cfPayrollPushSlipFact_(
        facts,
        itemTotals,
        "SLIP_DISTRIBUTION_COUNT",
        "伝票分配数",
        distributionCount,
        "件",
        "分配数"
      );

    }


    /* =================================================
     * プレビュー行
     * ================================================= */

    previewRows.push({

      sourceRow:
        sourceRowNumber,

      date:
        "",

      sourceCastId:
        sourceCastId,

      sourceCastName:
        sourceCastName,

      castId:
        payrollCastId,

      slipNumber:
        slipNumber,

      slipLine:
        slipLine,

      slipStatus:
        slipStatus,

      seatType:
        seatTypeRaw,

      normalizedSeatType:
        seatType,

      facts:
        facts

    });

  });


  /* =================================================
   * 結果
   * ================================================= */

  const result = {

    success:
      true,

    mode:
      "PREVIEW_ONLY",


    file: {

      id:
        prepared.originalFileId,

      name:
        prepared.originalFileName,

      type:
        prepared.originalFileType,

      temporaryConversion:
        prepared.temporary,

      sheetName:
        sourceSheet.getName()

    },


    importSource: {

      sourceId:
        source.sourceId,

      displayName:
        source.displayName,

      dataType:
        source.dataType,

      parserMode:
        "SLIP_ROWS"

    },


    summary: {

      sourceRows:
        values.length,

      recognizedRows:
        previewRows.length,

      skippedEmptyRows:
        skippedEmptyRows,

      skippedNoCastRows:
        skippedNoCastRows,

      blankDateRows:
        0,

      mappedColumns:
        requiredHeaders.length,

      unmappedColumns:
        0,

      recognizedCasts:
        Object.keys(
          recognizedCastIds
        ).length,

      unknownCastCount:
        Object.keys(
          unknownCasts
        ).length,

      vipRows:
        vipRows,

      bottleSalesRows:
        bottleSalesRows

    },


    itemTotals:
      itemTotals,


    unknownCasts:
      Object.keys(
        unknownCasts
      ).map(function(
        key
      ) {

        return unknownCasts[
          key
        ];

      }),


    rows:
      previewRows

  };


  Logger.log(
    "伝票一覧解析OK: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 26-B. 席種正規化
 *
 * ＶＩＰ
 * VIP
 * vip
 *
 * ↓
 *
 * VIP
 * ===================================================== */

function cfPayrollNormalizeSeatType_(
  value
) {

  let text =
    String(
      value || ""
    ).trim();


  if (!text) {

    return "";

  }


  /*
   * 全角英数字 → 半角
   */
  try {

    text =
      text.normalize(
        "NFKC"
      );

  } catch (
    error
  ) {

    /*
     * normalize非対応時も
     * 解析自体は継続。
     */

  }


  return text
    .replace(
      /\s+/g,
      ""
    )
    .toUpperCase();

}


/* =====================================================
 * 26-C. 伝票fact追加
 * ===================================================== */

function cfPayrollPushSlipFact_(
  facts,
  itemTotals,
  itemKey,
  displayName,
  value,
  unit,
  sourceHeader
) {

  facts.push({

    itemKey:
      itemKey,

    displayName:
      displayName,

    rawValue:
      value,

    value:
      value,

    unit:
      unit,

    multiplier:
      1,

    sourceHeader:
      sourceHeader

  });


  if (
    typeof value !==
    "number"
  ) {

    return;

  }


  if (
    !Object.prototype
      .hasOwnProperty.call(
        itemTotals,
        itemKey
      )
  ) {

    itemTotals[
      itemKey
    ] = {

      displayName:
        displayName,

      unit:
        unit,

      total:
        0

    };

  }


  itemTotals[
    itemKey
  ].total +=
    value;

}


/* =====================================================
 * 27. 営業日キュー / 履歴管理
 *
 * PAY_IMPORT_QUEUEを
 *
 * ・営業日未確定キュー
 * ・内容ハッシュ → 営業日の履歴
 *
 * として使用する。
 *
 * 過去データを再ダウンロードして
 * ファイル名が変わっても、
 *
 * 中身が同じなら営業日を復元できる。
 * ===================================================== */


/* =====================================================
 * 27-A. PAY_IMPORT_QUEUE作成
 * ===================================================== */

function setupPayrollImportDateQueue() {

  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    30000
  );


  try {

    const book =
      cfPayrollOpenSpreadsheet_();


    const sheetName =
      "PAY_IMPORT_QUEUE";


    const headers = [

      "キューID",
      "ファイルID",
      "ファイル名",
      "内容ハッシュ",
      "データ種別",
      "取込設定ID",
      "ダウンロード日時",
      "営業日候補",
      "営業日",
      "判定方法",
      "状態",
      "確認状態",
      "登録日時",
      "更新日時",
      "備考"

    ];


    let sheet =
      book.getSheetByName(
        sheetName
      );


    if (
      !sheet
    ) {

      sheet =
        book.insertSheet(
          sheetName
        );

    }


    cfPayrollEnsureColumns_(
      sheet,
      headers.length
    );


    if (
      sheet.getLastRow() ===
      0
    ) {

      cfPayrollWriteHeader_(
        sheet,
        headers
      );

    } else if (
      !cfPayrollHeaderMatches_(
        sheet,
        headers
      )
    ) {

      throw new Error(
        "PAY_IMPORT_QUEUEの列構成が異なります。" +
        "既存データを保護するため自動変更しません。"
      );

    }


    cfPayrollFormatHeader_(
      sheet,
      headers.length
    );


    if (
      sheet.getMaxRows() >
      1
    ) {

      sheet
        .getRange(
          2,
          7,
          sheet.getMaxRows() - 1,
          1
        )
        .setNumberFormat(
          "yyyy/mm/dd hh:mm:ss"
        );


      sheet
        .getRange(
          2,
          8,
          sheet.getMaxRows() - 1,
          2
        )
        .setNumberFormat(
          "yyyy/mm/dd"
        );


      sheet
        .getRange(
          2,
          13,
          sheet.getMaxRows() - 1,
          2
        )
        .setNumberFormat(
          "yyyy/mm/dd hh:mm:ss"
        );

    }


    /* =================================================
     * 初回過去データ一括設定用
     * ================================================= */

    cfPayrollSaveSystemSetting_(
      "IMPORT_SEQUENCE_START_DATE",
      "過去データ連番開始営業日",
      "",
      "例 2026/09/01。過去ファイルをダウンロード順に営業日候補へ並べる時だけ使用"
    );


    cfPayrollSaveSystemSetting_(
      "IMPORT_SEQUENCE_SOURCE_ID",
      "過去データ連番対象取込設定ID",
      "",
      "例 SRC_ORDER_SALES。空欄のままでは連番候補を作成しない"
    );


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      sheetName:
        sheetName

    };


    Logger.log(
      "営業日キュー作成OK: " +
      JSON.stringify(
        result
      )
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 27-B. 内容ハッシュから営業日履歴検索
 * ===================================================== */

function cfPayrollFindImportDateByHash_(
  contentHash,
  sourceId
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_QUEUE"
    );


  if (
    !sheet
  ) {

    return null;

  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <=
    1
  ) {

    return null;

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        15
      )
      .getValues();


  let found =
    null;


  rows.forEach(function(
    row,
    index
  ) {

    const rowHash =
      String(
        row[3] || ""
      ).trim();


    const rowSourceId =
      String(
        row[5] || ""
      ).trim();


    if (
      rowHash !==
        contentHash ||
      rowSourceId !==
        sourceId
    ) {

      return;

    }


    const businessDate =
      cfPayrollNormalizeBusinessDateValue_(
        row[8]
      );


    if (
      !businessDate
    ) {

      return;

    }


    if (
      found &&
      !cfPayrollSameBusinessDate_(
        found.businessDate,
        businessDate
      )
    ) {

      throw new Error(
        "同じ内容ハッシュに異なる営業日が登録されています。" +
        " PAY_IMPORT_QUEUEを確認してください。"
      );

    }


    found = {

      queueId:
        String(
          row[0] || ""
        ).trim(),

      sheetRow:
        index + 2,

      businessDate:
        businessDate,

      resolutionMethod:
        String(
          row[9] ||
          "HASH_HISTORY"
        ).trim()

    };

  });


  return found;

}


/* =====================================================
 * 27-C. キュー登録 / 更新
 * ===================================================== */

function cfPayrollUpsertImportDateQueue_(
  prepared,
  source,
  contentHash,
  businessDate,
  resolutionMethod
) {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_QUEUE"
    );


  if (
    !sheet
  ) {

    throw new Error(
      "PAY_IMPORT_QUEUEがありません。" +
      "setupPayrollImportDateQueue()を先に実行してください。"
    );

  }


  const normalizedBusinessDate =
    cfPayrollNormalizeBusinessDateValue_(
      businessDate
    );


  const lastRow =
    sheet.getLastRow();


  let existing =
    null;


  if (
    lastRow >
    1
  ) {

    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          15
        )
        .getValues();


    rows.some(function(
      row,
      index
    ) {

      const rowHash =
        String(
          row[3] || ""
        ).trim();


      const rowSourceId =
        String(
          row[5] || ""
        ).trim();


      if (
        rowHash ===
          contentHash &&
        rowSourceId ===
          source.sourceId
      ) {

        existing = {

          sheetRow:
            index + 2,

          queueId:
            String(
              row[0] || ""
            ).trim(),

          businessDate:
            cfPayrollNormalizeBusinessDateValue_(
              row[8]
            ),

          candidateDate:
            cfPayrollNormalizeBusinessDateValue_(
              row[7]
            ),

          resolutionMethod:
            String(
              row[9] || ""
            ).trim(),

          status:
            String(
              row[10] || ""
            ).trim(),

          confirmation:
            String(
              row[11] || ""
            ).trim()

        };


        return true;

      }


      return false;

    });

  }


  const now =
    new Date();


  const downloadDateTime =
    cfPayrollExtractDownloadTimestamp_(
      prepared.originalFileName
    );


  /* =================================================
   * 既存
   * ================================================= */

  if (
    existing
  ) {

    let finalBusinessDate =
      existing.businessDate;


    let finalMethod =
      existing.resolutionMethod;


    if (
      normalizedBusinessDate
    ) {

      if (
        finalBusinessDate &&
        !cfPayrollSameBusinessDate_(
          finalBusinessDate,
          normalizedBusinessDate
        )
      ) {

        throw new Error(
          "同一ファイル内容に異なる営業日を設定しようとしました。" +
          " PAY_IMPORT_QUEUEを確認してください。"
        );

      }


      finalBusinessDate =
        normalizedBusinessDate;


      finalMethod =
        resolutionMethod ||
        finalMethod ||
        "MANUAL_QUEUE";

    }


    const finalStatus =
      finalBusinessDate
        ? "確定"
        : (
            existing.candidateDate
              ? "候補確認待ち"
              : "営業日未設定"
          );


    const finalConfirmation =
      finalBusinessDate
        ? "確認済"
        : "要確認";


    const registeredAt =
      sheet
        .getRange(
          existing.sheetRow,
          13
        )
        .getValue() ||
      now;


    sheet
      .getRange(
        existing.sheetRow,
        2,
        1,
        14
      )
      .setValues([[
        prepared.originalFileId,
        prepared.originalFileName,
        contentHash,
        source.dataType,
        source.sourceId,
        downloadDateTime || "",
        existing.candidateDate || "",
        finalBusinessDate || "",
        finalMethod || "PENDING",
        finalStatus,
        finalConfirmation,
        registeredAt,
        now,
        ""
      ]]);


    return {

      queueId:
        existing.queueId,

      businessDate:
        finalBusinessDate,

      resolutionMethod:
        finalMethod

    };

  }


  /* =================================================
   * 新規
   * ================================================= */

  const queueId =
    "Q_" +
    Utilities.getUuid();


  const status =
    normalizedBusinessDate
      ? "確定"
      : "営業日未設定";


  const confirmation =
    normalizedBusinessDate
      ? "確認済"
      : "要確認";


  const method =
    normalizedBusinessDate
      ? (
          resolutionMethod ||
          "MANUAL_QUEUE"
        )
      : "PENDING";


  sheet
    .appendRow([
      queueId,
      prepared.originalFileId,
      prepared.originalFileName,
      contentHash,
      source.dataType,
      source.sourceId,
      downloadDateTime || "",
      "",
      normalizedBusinessDate || "",
      method,
      status,
      confirmation,
      now,
      now,
      ""
    ]);


  return {

    queueId:
      queueId,

    businessDate:
      normalizedBusinessDate,

    resolutionMethod:
      method

  };

}


/* =====================================================
 * 27-D. 営業日値正規化
 *
 * 対応：
 *
 * Date
 * yyyy/mm/dd
 * yyyy-mm-dd
 * yyyymmdd
 * ===================================================== */

function cfPayrollNormalizeBusinessDateValue_(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return null;

  }


  if (
    Object.prototype
      .toString.call(
        value
      ) ===
    "[object Date]"
  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return null;

    }


    return cfPayrollBuildBusinessDate_(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate()
    );

  }


  const text =
    String(
      value
    ).trim();


  const match =
    text.match(
      /^(20\d{2})[\/_-]?(\d{2})[\/_-]?(\d{2})$/
    );


  if (
    !match
  ) {

    return null;

  }


  return cfPayrollBuildBusinessDate_(
    Number(
      match[1]
    ),
    Number(
      match[2]
    ),
    Number(
      match[3]
    )
  );

}


/* =====================================================
 * 27-E. 営業日比較
 * ===================================================== */

function cfPayrollSameBusinessDate_(
  left,
  right
) {

  if (
    !left ||
    !right
  ) {

    return false;

  }


  return (
    left.getFullYear() ===
      right.getFullYear() &&
    left.getMonth() ===
      right.getMonth() &&
    left.getDate() ===
      right.getDate()
  );

}


/* =====================================================
 * 27-F. ファイル名からDL日時取得
 *
 * 例：
 *
 * _20260914_233104.xlsx
 * ===================================================== */

function cfPayrollExtractDownloadTimestamp_(
  fileName
) {

  const text =
    String(
      fileName || ""
    );


  const match =
    text.match(
      /(?:^|_)(20\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})(?=\D|$)/
    );


  if (
    !match
  ) {

    return null;

  }


  const year =
    Number(
      match[1]
    );


  const month =
    Number(
      match[2]
    );


  const day =
    Number(
      match[3]
    );


  const hour =
    Number(
      match[4]
    );


  const minute =
    Number(
      match[5]
    );


  const second =
    Number(
      match[6]
    );


  const date =
    new Date(
      year,
      month - 1,
      day,
      hour,
      minute,
      second,
      0
    );


  if (
    date.getFullYear() !==
      year ||
    date.getMonth() !==
      month - 1 ||
    date.getDate() !==
      day ||
    date.getHours() !==
      hour ||
    date.getMinutes() !==
      minute ||
    date.getSeconds() !==
      second
  ) {

    return null;

  }


  return date;

}


/* =====================================================
 * 27-G. 取込待ちを営業日キューへ登録
 *
 * 給与解析・PAY_FACTS書込はしない。
 *
 * 営業日管理のみ。
 * ===================================================== */

function scanPayrollInboxBusinessDates() {

  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    30000
  );


  try {

    const inbox =
      cfPayrollGetInboxFolder_();


    const sources =
      cfPayrollGetImportSources_();


    const files =
      inbox.getFiles();


    let totalFiles =
      0;


    let matched =
      0;


    let queued =
      0;


    let resolved =
      0;


    let unmatched =
      0;


    const results =
      [];


    while (
      files.hasNext()
    ) {

      const file =
        files.next();


      totalFiles++;


      const fileInfo =
        cfPayrollBuildImportFileInfo_(
          file
        );


      const detection =
        cfPayrollDetectImportSourceForFile_(
          fileInfo,
          sources
        );


      if (
        detection.status !==
        "MATCHED"
      ) {

        unmatched++;


        results.push({

          fileName:
            fileInfo.name,

          status:
            detection.status

        });


        continue;

      }


      matched++;


      const source =
        cfPayrollGetImportSourceById_(
          detection.sourceId
        );


      if (
        !source
      ) {

        continue;

      }


      const prepared = {

        originalFileId:
          fileInfo.id,

        originalFileName:
          fileInfo.name

      };


      const resolution =
        cfPayrollResolveBusinessDate_(
          prepared,
          source
        );


      if (
        resolution.date
      ) {

        resolved++;

      } else {

        queued++;

      }


      results.push({

        fileName:
          fileInfo.name,

        sourceId:
          source.sourceId,

        resolved:
          !!resolution.date,

        method:
          resolution.method,

        queueId:
          resolution.queueId

      });

    }


    const result = {

      success:
        true,

      totalFiles:
        totalFiles,

      matched:
        matched,

      resolved:
        resolved,

      queued:
        queued,

      unmatched:
        unmatched,

      files:
        results

    };


    Logger.log(
      "営業日キュースキャンOK: " +
      JSON.stringify(
        result
      )
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 27-H. 未確定営業日一覧
 * ===================================================== */

function previewPayrollImportDateQueue() {

  const book =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    book.getSheetByName(
      "PAY_IMPORT_QUEUE"
    );


  if (
    !sheet
  ) {

    throw new Error(
      "PAY_IMPORT_QUEUEがありません。"
    );

  }


  const lastRow =
    sheet.getLastRow();


  const pending =
    [];


  if (
    lastRow >
    1
  ) {

    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          15
        )
        .getDisplayValues();


    rows.forEach(function(
      row,
      index
    ) {

      const businessDate =
        String(
          row[8] || ""
        ).trim();


      if (
        businessDate
      ) {

        return;

      }


      pending.push({

        sheetRow:
          index + 2,

        queueId:
          row[0],

        fileName:
          row[2],

        dataType:
          row[4],

        sourceId:
          row[5],

        downloadDateTime:
          row[6],

        candidateDate:
          row[7],

        status:
          row[10]

      });

    });

  }


  const result = {

    success:
      true,

    pendingCount:
      pending.length,

    pending:
      pending

  };


  Logger.log(
    "営業日未確定一覧: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 27-I. 過去データ連番営業日候補
 *
 * PAY_SETTINGS：
 *
 * 過去データ連番開始営業日
 * 過去データ連番対象取込設定ID
 *
 * を使う。
 *
 * DL日時の古い順に
 * 1日ずつ営業日候補を付ける。
 *
 * まだ営業日として確定はしない。
 * ===================================================== */

function buildPayrollSequentialDateCandidates() {

  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    30000
  );


  try {

    const startValue =
      cfPayrollGetSetting_(
        "IMPORT_SEQUENCE_START_DATE",
        "",
        false
      );


    const startDate =
      cfPayrollNormalizeBusinessDateValue_(
        startValue
      );


    const targetSourceId =
      String(
        cfPayrollGetSetting_(
          "IMPORT_SEQUENCE_SOURCE_ID",
          "",
          false
        ) ||
        ""
      ).trim();


    if (
      !startDate
    ) {

      throw new Error(
        "PAY_SETTINGSの「過去データ連番開始営業日」を" +
        "先に入力してください。例：2026/09/01"
      );

    }


    if (
      !targetSourceId
    ) {

      throw new Error(
        "PAY_SETTINGSの「過去データ連番対象取込設定ID」を" +
        "先に入力してください。例：SRC_ORDER_SALES"
      );

    }


    const book =
      cfPayrollOpenSpreadsheet_();


    const sheet =
      book.getSheetByName(
        "PAY_IMPORT_QUEUE"
      );


    if (
      !sheet
    ) {

      throw new Error(
        "PAY_IMPORT_QUEUEがありません。"
      );

    }


    const lastRow =
      sheet.getLastRow();


    if (
      lastRow <=
      1
    ) {

      return {

        success:
          true,

        candidateCount:
          0

      };

    }


    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          15
        )
        .getValues();


    const targets =
      [];


    rows.forEach(function(
      row,
      index
    ) {

      const sourceId =
        String(
          row[5] || ""
        ).trim();


      if (
        sourceId !==
        targetSourceId
      ) {

        return;

      }


      const businessDate =
        cfPayrollNormalizeBusinessDateValue_(
          row[8]
        );


      if (
        businessDate
      ) {

        return;

      }


      const downloadDate =
        cfPayrollNormalizeDateTimeValue_(
          row[6]
        ) ||
        cfPayrollNormalizeDateTimeValue_(
          row[12]
        );


      targets.push({

        sheetRow:
          index + 2,

        downloadDate:
          downloadDate,

        fileName:
          String(
            row[2] || ""
          )

      });

    });


    targets.sort(function(
      left,
      right
    ) {

      const leftTime =
        left.downloadDate
          ? left.downloadDate.getTime()
          : Number.MAX_SAFE_INTEGER;


      const rightTime =
        right.downloadDate
          ? right.downloadDate.getTime()
          : Number.MAX_SAFE_INTEGER;


      if (
        leftTime !==
        rightTime
      ) {

        return (
          leftTime -
          rightTime
        );

      }


      return left.fileName
        .localeCompare(
          right.fileName,
          "ja"
        );

    });


    const now =
      new Date();


    targets.forEach(function(
      target,
      index
    ) {

      const candidate =
        new Date(
          startDate.getFullYear(),
          startDate.getMonth(),
          startDate.getDate() + index,
          0,
          0,
          0,
          0
        );


      const registeredAt =
        sheet
          .getRange(
            target.sheetRow,
            13
          )
          .getValue() ||
        now;


      sheet
        .getRange(
          target.sheetRow,
          8,
          1,
          7
        )
        .setValues([[
          candidate,
          "",
          "SEQUENCE_CANDIDATE",
          "候補確認待ち",
          "要確認",
          registeredAt,
          now
        ]]);

    });


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      sourceId:
        targetSourceId,

      startDate:
        startDate,

      candidateCount:
        targets.length

    };


    Logger.log(
      "営業日連番候補作成OK: " +
      JSON.stringify(
        result
      )
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 27-J. 連番候補を営業日確定
 *
 * PAY_IMPORT_QUEUEを目視確認した後だけ実行。
 *
 * 候補確認待ちのみ対象。
 * ===================================================== */

function confirmPayrollSequentialDateCandidates() {

  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    30000
  );


  try {

    const targetSourceId =
      String(
        cfPayrollGetSetting_(
          "IMPORT_SEQUENCE_SOURCE_ID",
          "",
          false
        ) ||
        ""
      ).trim();


    if (
      !targetSourceId
    ) {

      throw new Error(
        "PAY_SETTINGSの「過去データ連番対象取込設定ID」を" +
        "先に入力してください。例：SRC_ORDER_SALES"
      );

    }


    const book =
      cfPayrollOpenSpreadsheet_();


    const sheet =
      book.getSheetByName(
        "PAY_IMPORT_QUEUE"
      );


    if (
      !sheet
    ) {

      throw new Error(
        "PAY_IMPORT_QUEUEがありません。"
      );

    }


    const lastRow =
      sheet.getLastRow();


    if (
      lastRow <=
      1
    ) {

      return {

        success:
          true,

        confirmed:
          0

      };

    }


    const rows =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          15
        )
        .getValues();


    const now =
      new Date();


    let confirmed =
      0;


    rows.forEach(function(
      row,
      index
    ) {

      const sourceId =
        String(
          row[5] || ""
        ).trim();


      if (
        sourceId !==
        targetSourceId
      ) {

        return;

      }


      const status =
        String(
          row[10] || ""
        ).trim();


      const candidateDate =
        cfPayrollNormalizeBusinessDateValue_(
          row[7]
        );


      const existingBusinessDate =
        cfPayrollNormalizeBusinessDateValue_(
          row[8]
        );


      if (
        status !==
          "候補確認待ち" ||
        !candidateDate ||
        existingBusinessDate
      ) {

        return;

      }


      const sheetRow =
        index + 2;


      const registeredAt =
        row[12] ||
        now;


      sheet
        .getRange(
          sheetRow,
          9,
          1,
          6
        )
        .setValues([[
          candidateDate,
          "SEQUENCE_CONFIRMED",
          "確定",
          "確認済",
          registeredAt,
          now
        ]]);


      confirmed++;

    });


    SpreadsheetApp.flush();


    const result = {

      success:
        true,

      sourceId:
        targetSourceId,

      confirmed:
        confirmed

    };


    Logger.log(
      "営業日連番候補確定OK: " +
      JSON.stringify(
        result
      )
    );


    return result;


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
 * 27-K. DateTime正規化
 * ===================================================== */

function cfPayrollNormalizeDateTimeValue_(
  value
) {

  if (
    Object.prototype
      .toString.call(
        value
      ) ===
      "[object Date]" &&
    !isNaN(
      value.getTime()
    )
  ) {

    return value;

  }


  return null;

}


/* =====================================================
 * 28. 精算入金一覧照会
 *     監査・照合専用Parser
 *
 * 給与バックの直接計算には使用しない。
 *
 * 全項目キーを
 * AUDIT_
 * で開始する。
 *
 * 主な用途：
 *
 * ・料金合計
 * ・現金入金
 * ・クレジット入金
 * ・差引
 * ・売上伝票件数
 * ・注文明細
 * ・場内明細
 * ・値引調整
 *
 * の検算・追跡。
 *
 * 「ドリンク」列は金額なので
 * BACK_DRINKには絶対変換しない。
 * ===================================================== */


/* =====================================================
 * 28-A. 精算入金一覧解析 V2
 *
 * 実ファイル構造対応：
 *
 * 「入金明細」の2列左
 *   → 明細種別
 *      入金 / 注文 / 場内 / 値引調整
 *
 * 「入金明細」の1列左
 *   → 入金方法
 *      現金 / VISA / JCB 等
 *
 * 「入金明細」
 *   → 入金額
 *
 * 給与計算には使わず
 * AUDIT_項目のみ生成する。
 * ===================================================== */

function cfPayrollAnalyzeSettlementDetail_(
  sourceBook,
  prepared,
  source
) {

  let sourceSheet;


  /* =================================================
   * 対象シート
   * ================================================= */

  if (
    source.sheetName
  ) {

    sourceSheet =
      sourceBook.getSheetByName(
        source.sheetName
      );


    if (!sourceSheet) {

      throw new Error(
        "対象シートが見つかりません：" +
        source.sheetName +
        " / ファイル：" +
        prepared.originalFileName
      );

    }

  } else {

    const sheets =
      sourceBook.getSheets();


    if (
      sheets.length === 0
    ) {

      throw new Error(
        "取込ファイルにシートがありません。"
      );

    }


    sourceSheet =
      sheets[0];

  }


  const lastRow =
    sourceSheet.getLastRow();


  const lastColumn =
    sourceSheet.getLastColumn();


  if (
    lastRow <
      source.startRow ||
    lastColumn < 1
  ) {

    throw new Error(
      "取込対象データがありません。"
    );

  }


  /* =================================================
   * ヘッダー
   * ================================================= */

  const headerRow =
    source.startRow - 1;


  const headers =
    sourceSheet
      .getRange(
        headerRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  /* =================================================
   * 必須列
   * ================================================= */

  const requiredHeaders = [

    "伝票種類",
    "伝票",
    "行",
    "キャスト",

    "料金合計",
    "内消費税",
    "差引",

    "現金入金",
    "クレ入金",

    "入金明細",
    "他売明細",

    "Ｃ",
    "名前",
    "数",
    "単価",
    "金額"

  ];


  const indexes =
    {};


  requiredHeaders.forEach(function(
    header
  ) {

    const columnIndex =
      cfPayrollFindConfiguredHeaderIndex_(
        headers,
        header
      );


    if (
      columnIndex === -1
    ) {

      throw new Error(
        "精算入金一覧照会に必要な列がありません：" +
        header
      );

    }


    indexes[
      header
    ] =
      columnIndex;

  });


  /* =================================================
   * 精算入金の特殊明細列
   *
   * 実ファイル上：
   *
   * [無名]     ← 入金 / 注文 / 場内 / 値引調整
   * [無名]     ← 現金 / VISA / JCB
   * 入金明細   ← 入金額
   * 他売明細
   *
   * 「入金明細」から相対位置で判定する。
   * ================================================= */

  const detailTypeIndex =
    indexes[
      "入金明細"
    ] - 2;


  const paymentMethodIndex =
    indexes[
      "入金明細"
    ] - 1;


  const paymentDetailAmountIndex =
    indexes[
      "入金明細"
    ];


  if (
    detailTypeIndex < 0 ||
    paymentMethodIndex < 0
  ) {

    throw new Error(
      "精算入金一覧の明細列位置を判定できません。"
    );

  }


  /* =================================================
   * 任意列
   * ================================================= */

  const optionalHeaders = [

    "入金",
    "精算",
    "卓",

    "基本料金",
    "延長料金",
    "指名料金",
    "場内料金",

    "ドリンク",
    "フード",
    "他注文",
    "ボトル",

    "値引調整",

    "他売上",
    "他入金"

  ];


  optionalHeaders.forEach(function(
    header
  ) {

    indexes[
      header
    ] =
      cfPayrollFindConfiguredHeaderIndex_(
        headers,
        header
      );

  });


  /* =================================================
   * データ
   * ================================================= */

  const rowCount =
    lastRow -
    source.startRow +
    1;


  const values =
    sourceSheet
      .getRange(
        source.startRow,
        1,
        rowCount,
        lastColumn
      )
      .getValues();


  const castResolver =
    cfPayrollBuildCastResolver_();


  const previewRows =
    [];


  const itemTotals =
    {};


  const resolvedCastIds =
    {};


  const unlinkedCastNames =
    {};


  let skippedEmptyRows =
    0;


  let settlementRows =
    0;


  let salesSlipRows =
    0;


  let paymentDetailRows =
    0;


  let orderDetailRows =
    0;


  let inhouseDetailRows =
    0;


  let adjustmentDetailRows =
    0;


  let otherDetailRows =
    0;


  let ignoredRows =
    0;


  let paymentMismatchRows =
    0;


  /* =================================================
   * 精算集計項目
   * ================================================= */

  const summaryDefinitions = [

    [
      "基本料金",
      "AUDIT_BASE_FEE",
      "基本料金"
    ],

    [
      "延長料金",
      "AUDIT_EXTENSION_FEE",
      "延長料金"
    ],

    [
      "指名料金",
      "AUDIT_NOMINATION_FEE",
      "指名料金"
    ],

    [
      "場内料金",
      "AUDIT_INHOUSE_FEE",
      "場内料金"
    ],

    [
      "ドリンク",
      "AUDIT_DRINK_SALES",
      "ドリンク売上"
    ],

    [
      "フード",
      "AUDIT_FOOD_SALES",
      "フード売上"
    ],

    [
      "他注文",
      "AUDIT_OTHER_ORDER_SALES",
      "その他注文売上"
    ],

    [
      "ボトル",
      "AUDIT_BOTTLE_SALES",
      "ボトル売上監査"
    ],

    [
      "値引調整",
      "AUDIT_DISCOUNT_ADJUSTMENT",
      "値引調整"
    ],

    [
      "料金合計",
      "AUDIT_TOTAL_CHARGE",
      "料金合計"
    ],

    [
      "内消費税",
      "AUDIT_INCLUDED_TAX",
      "内消費税"
    ],

    [
      "他売上",
      "AUDIT_OTHER_SALES",
      "その他売上"
    ],

    [
      "他入金",
      "AUDIT_OTHER_PAYMENT",
      "その他入金"
    ],

    [
      "差引",
      "AUDIT_BALANCE",
      "差引"
    ],

    [
      "現金入金",
      "AUDIT_CASH_PAYMENT",
      "現金入金"
    ],

    [
      "クレ入金",
      "AUDIT_CREDIT_PAYMENT",
      "クレジット入金"
    ]

  ];


  /* =================================================
   * 行解析
   * ================================================= */

  values.forEach(function(
    row,
    index
  ) {

    const sourceRowNumber =
      source.startRow +
      index;


    const hasAnyValue =
      row.some(function(
        value
      ) {

        return (
          value !== "" &&
          value !== null
        );

      });


    if (
      !hasAnyValue
    ) {

      skippedEmptyRows++;

      return;

    }


    /* =================================================
     * 行情報
     * ================================================= */

    const slipType =
      cfPayrollNormalizeSettlementText_(
        row[
          indexes[
            "伝票種類"
          ]
        ]
      );


    const detailType =
      cfPayrollNormalizeSettlementText_(
        row[
          detailTypeIndex
        ]
      );


    const paymentMethod =
      String(
        row[
          paymentMethodIndex
        ] || ""
      ).trim();


    const paymentDetailAmount =
      cfPayrollAuditNumber_(
        row[
          paymentDetailAmountIndex
        ]
      );


    const castName =
      String(
        row[
          indexes[
            "キャスト"
          ]
        ] || ""
      ).trim();


    let castId =
      "";


    if (
      castName
    ) {

      castId =
        cfPayrollResolveCastId_(
          castResolver,
          "",
          castName
        );


      if (
        castId
      ) {

        resolvedCastIds[
          castId
        ] =
          true;

      } else {

        unlinkedCastNames[
          cfPayrollNormalizeCastName_(
            castName
          )
        ] =
          castName;

      }

    }


    const facts =
      [];


    let recordType =
      "";


    /* =================================================
     * 精算集計
     * ================================================= */

    if (
      slipType ===
        "【精算入金】" ||
      slipType ===
        "精算入金"
    ) {

      recordType =
        "SETTLEMENT_SUMMARY";


      settlementRows++;


      cfPayrollPushAuditFact_(
        facts,
        itemTotals,
        "AUDIT_SETTLEMENT_COUNT",
        "精算件数",
        1,
        "件",
        "伝票種類"
      );


      summaryDefinitions.forEach(function(
        definition
      ) {

        const columnIndex =
          indexes[
            definition[0]
          ];


        if (
          columnIndex ===
            undefined ||
          columnIndex === -1
        ) {

          return;

        }


        const value =
          cfPayrollAuditNumber_(
            row[
              columnIndex
            ]
          );


        if (
          value === null ||
          value === 0
        ) {

          return;

        }


        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          definition[1],
          definition[2],
          value,
          "円",
          definition[0]
        );

      });


      /* ---------------------------------------------
       * 入金差額
       * --------------------------------------------- */

      const balance =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "差引"
            ]
          ]
        ) || 0;


      const cash =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "現金入金"
            ]
          ]
        ) || 0;


      const credit =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "クレ入金"
            ]
          ]
        ) || 0;


      const otherPaymentIndex =
        indexes[
          "他入金"
        ];


      const otherPayment =
        (
          otherPaymentIndex !==
            undefined &&
          otherPaymentIndex !== -1
        )
          ? (
              cfPayrollAuditNumber_(
                row[
                  otherPaymentIndex
                ]
              ) || 0
            )
          : 0;


      const difference =
        balance -
        (
          cash +
          credit +
          otherPayment
        );


      if (
        Math.abs(
          difference
        ) >
        0.000001
      ) {

        paymentMismatchRows++;


        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          "AUDIT_PAYMENT_DIFFERENCE",
          "入金差額",
          difference,
          "円",
          "差引-(現金+クレ+他入金)"
        );

      }

    }


    /* =================================================
     * 売上伝票
     * ================================================= */

    else if (
      slipType ===
      "売上伝票"
    ) {

      recordType =
        "SALES_SLIP";


      salesSlipRows++;


      cfPayrollPushAuditFact_(
        facts,
        itemTotals,
        "AUDIT_SALES_SLIP_COUNT",
        "売上伝票行数",
        1,
        "件",
        "伝票種類"
      );

    }


    /* =================================================
     * 入金明細
     * ================================================= */

    else if (
      detailType ===
      "入金"
    ) {

      recordType =
        "PAYMENT_DETAIL";


      paymentDetailRows++;


      cfPayrollPushAuditFact_(
        facts,
        itemTotals,
        "AUDIT_PAYMENT_DETAIL_COUNT",
        "入金明細行数",
        1,
        "件",
        "明細種別"
      );


      if (
        paymentDetailAmount !== null &&
        paymentDetailAmount !== 0
      ) {

        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          "AUDIT_PAYMENT_DETAIL_AMOUNT",
          "入金明細金額",
          paymentDetailAmount,
          "円",
          "入金明細"
        );

      }

    }


    /* =================================================
     * 注文明細
     * ================================================= */

    else if (
      detailType ===
      "注文"
    ) {

      recordType =
        "ORDER_DETAIL";


      orderDetailRows++;


      const quantity =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "数"
            ]
          ]
        );


      const amount =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "金額"
            ]
          ]
        );


      if (
        quantity !== null &&
        quantity !== 0
      ) {

        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          "AUDIT_ORDER_QUANTITY",
          "注文明細数量",
          quantity,
          "個",
          "数"
        );

      }


      if (
        amount !== null &&
        amount !== 0
      ) {

        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          "AUDIT_ORDER_AMOUNT",
          "注文明細金額",
          amount,
          "円",
          "金額"
        );

      }

    }


    /* =================================================
     * 場内
     * ================================================= */

    else if (
      detailType ===
      "場内"
    ) {

      recordType =
        "INHOUSE_DETAIL";


      inhouseDetailRows++;


      const amount =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "金額"
            ]
          ]
        );


      if (
        amount !== null &&
        amount !== 0
      ) {

        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          "AUDIT_INHOUSE_DETAIL_AMOUNT",
          "場内明細金額",
          amount,
          "円",
          "金額"
        );

      }

    }


    /* =================================================
     * 値引調整
     * ================================================= */

    else if (
      detailType ===
      "値引調整"
    ) {

      recordType =
        "ADJUSTMENT_DETAIL";


      adjustmentDetailRows++;


      const amount =
        cfPayrollAuditNumber_(
          row[
            indexes[
              "金額"
            ]
          ]
        );


      if (
        amount !== null &&
        amount !== 0
      ) {

        cfPayrollPushAuditFact_(
          facts,
          itemTotals,
          "AUDIT_ADJUSTMENT_DETAIL_AMOUNT",
          "値引調整明細金額",
          amount,
          "円",
          "金額"
        );

      }

    }


    /* =================================================
     * その他明細
     * ================================================= */

    else if (
      detailType
    ) {

      recordType =
        "OTHER_DETAIL";


      otherDetailRows++;

    }


    else {

      ignoredRows++;

      return;

    }


    /* =================================================
     * プレビュー
     * ================================================= */

    previewRows.push({

      sourceRow:
        sourceRowNumber,

      date:
        "",

      sourceCastId:
        "",

      sourceCastName:
        castName,

      castId:
        castId,

      recordType:
        recordType,

      settlementNumber:
        (
          indexes[
            "精算"
          ] !== undefined &&
          indexes[
            "精算"
          ] !== -1
        )
          ? String(
              row[
                indexes[
                  "精算"
                ]
              ] || ""
            ).trim()
          : "",

      depositNumber:
        (
          indexes[
            "入金"
          ] !== undefined &&
          indexes[
            "入金"
          ] !== -1
        )
          ? String(
              row[
                indexes[
                  "入金"
                ]
              ] || ""
            ).trim()
          : "",

      slipNumber:
        String(
          row[
            indexes[
              "伝票"
            ]
          ] || ""
        ).trim(),

      slipLine:
        String(
          row[
            indexes[
              "行"
            ]
          ] || ""
        ).trim(),

      detailType:
        detailType,

      paymentMethod:
        paymentMethod,

      detailName:
        String(
          row[
            indexes[
              "名前"
            ]
          ] || ""
        ).trim(),

      productCode:
        String(
          row[
            indexes[
              "Ｃ"
            ]
          ] || ""
        ).trim(),

      quantity:
        cfPayrollAuditNumber_(
          row[
            indexes[
              "数"
            ]
          ]
        ),

      unitPrice:
        cfPayrollAuditNumber_(
          row[
            indexes[
              "単価"
            ]
          ]
        ),

      detailAmount:
        cfPayrollAuditNumber_(
          row[
            indexes[
              "金額"
            ]
          ]
        ),

      paymentDetailAmount:
        paymentDetailAmount,

      facts:
        facts

    });

  });


  /* =================================================
   * 結果
   * ================================================= */

  const result = {

    success:
      true,

    mode:
      "PREVIEW_ONLY",


    file: {

      id:
        prepared.originalFileId,

      name:
        prepared.originalFileName,

      type:
        prepared.originalFileType,

      temporaryConversion:
        prepared.temporary,

      sheetName:
        sourceSheet.getName()

    },


    importSource: {

      sourceId:
        source.sourceId,

      displayName:
        source.displayName,

      dataType:
        source.dataType,

      parserMode:
        "SETTLEMENT_DETAIL"

    },


    summary: {

      sourceRows:
        values.length,

      recognizedRows:
        previewRows.length,

      skippedEmptyRows:
        skippedEmptyRows,

      ignoredRows:
        ignoredRows,

      blankDateRows:
        0,

      mappedColumns:
        requiredHeaders.length,

      unmappedColumns:
        0,

      recognizedCasts:
        Object.keys(
          resolvedCastIds
        ).length,

      /*
       * 監査専用なので
       * CAST_ID未登録で本取込停止しない。
       */
      unknownCastCount:
        0,

      unlinkedCastNameCount:
        Object.keys(
          unlinkedCastNames
        ).length,

      settlementRows:
        settlementRows,

      salesSlipRows:
        salesSlipRows,

      paymentDetailRows:
        paymentDetailRows,

      orderDetailRows:
        orderDetailRows,

      inhouseDetailRows:
        inhouseDetailRows,

      adjustmentDetailRows:
        adjustmentDetailRows,

      otherDetailRows:
        otherDetailRows,

      paymentMismatchRows:
        paymentMismatchRows

    },


    itemTotals:
      itemTotals,


    unknownCasts:
      [],


    unlinkedCastNames:
      Object.keys(
        unlinkedCastNames
      )
        .map(function(
          key
        ) {

          return unlinkedCastNames[
            key
          ];

        }),


    rows:
      previewRows

  };


  Logger.log(
    "精算入金監査解析V2 OK: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 28-B. 精算文字列正規化
 * ===================================================== */

function cfPayrollNormalizeSettlementText_(
  value
) {

  let text =
    String(
      value || ""
    ).trim();


  if (
    !text
  ) {

    return "";

  }


  try {

    text =
      text.normalize(
        "NFKC"
      );

  } catch (
    error
  ) {

    /*
     * normalize非対応でも継続
     */

  }


  return text.replace(
    /\s+/g,
    ""
  );

}


/* =====================================================
 * 28-C. 精算数値変換
 * ===================================================== */

function cfPayrollAuditNumber_(
  value
) {

  if (
    value === "" ||
    value === null ||
    value === undefined
  ) {

    return null;

  }


  if (
    typeof value ===
    "number"
  ) {

    return Number.isFinite(
      value
    )
      ? value
      : null;

  }


  const text =
    String(
      value
    )
      .trim()
      .replace(
        /,/g,
        ""
      );


  if (
    !text
  ) {

    return null;

  }


  const number =
    Number(
      text
    );


  return Number.isFinite(
    number
  )
    ? number
    : null;

}


/* =====================================================
 * 28-D. 監査fact追加
 *
 * AUDIT_で始まる項目は
 * 給与計算対象外。
 * ===================================================== */

function cfPayrollPushAuditFact_(
  facts,
  itemTotals,
  itemKey,
  displayName,
  value,
  unit,
  sourceHeader
) {

  facts.push({

    itemKey:
      itemKey,

    displayName:
      displayName,

    rawValue:
      value,

    value:
      value,

    unit:
      unit,

    multiplier:
      1,

    sourceHeader:
      sourceHeader

  });


  if (
    typeof value !==
    "number"
  ) {

    return;

  }


  if (
    !Object.prototype
      .hasOwnProperty.call(
        itemTotals,
        itemKey
      )
  ) {

    itemTotals[
      itemKey
    ] = {

      displayName:
        displayName,

      unit:
        unit,

      total:
        0

    };

  }


  itemTotals[
    itemKey
  ].total +=
    value;

}


/* =====================================================
 * 28-E. 精算入金 未分類明細 診断
 *
 * OTHER_DETAILになった明細種別だけを
 * 分かりやすくログへ出す。
 *
 * データ書込・移動なし。
 * ===================================================== */

function debugPayrollSettlementOtherDetails() {

  const inbox =
    cfPayrollGetInboxFolder_();


  const sources =
    cfPayrollGetImportSources_();


  const files =
    inbox.getFiles();


  const results =
    [];


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    const fileInfo =
      cfPayrollBuildImportFileInfo_(
        file
      );


    const detection =
      cfPayrollDetectImportSourceForFile_(
        fileInfo,
        sources
      );


    if (
      detection.status !==
      "MATCHED" ||
      detection.sourceId !==
      "SRC_SETTLEMENT"
    ) {

      continue;

    }


    const source =
      cfPayrollGetImportSourceById_(
        "SRC_SETTLEMENT"
      );


    if (!source) {

      continue;

    }


    cfPayrollWithReadableSpreadsheet_(
      file.getId(),
      function(
        sourceBook,
        prepared
      ) {

        const preview =
          cfPayrollAnalyzeSettlementDetail_(
            sourceBook,
            prepared,
            source
          );


        const others =
          (
            preview.rows ||
            []
          )
            .filter(function(
              row
            ) {

              return (
                row.recordType ===
                "OTHER_DETAIL"
              );

            })
            .map(function(
              row
            ) {

              return {

                sourceRow:
                  row.sourceRow,

                detailType:
                  row.detailType,

                detailName:
                  row.detailName,

                productCode:
                  row.productCode,

                detailAmount:
                  row.detailAmount

              };

            });


        if (
          others.length > 0
        ) {

          results.push({

            fileName:
              prepared.originalFileName,

            count:
              others.length,

            rows:
              others

          });

        }

      }
    );

  }


  const result = {

    success:
      true,

    filesWithOtherDetails:
      results.length,

    totalOtherDetails:
      results.reduce(function(
        total,
        item
      ) {

        return (
          total +
          item.count
        );

      }, 0),

    details:
      results

  };


  Logger.log(
    "精算入金 未分類明細診断: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 29. 本取込前 総合安全テスト V3
 *
 * 79ファイル一括ではApps Scriptの
 * 実行時間上限に達するため、
 *
 * ORDER
 * ATTENDANCE
 * PERFORMANCE
 * SETTLEMENT
 * SLIP
 *
 * を個別実行する。
 *
 * 各結果はScriptPropertiesへ
 * コンパクトに保存。
 *
 * PAY_FACTS：
 *   書き込まない
 *
 * PAY_IMPORT_LOG：
 *   書き込まない
 *
 * ファイル：
 *   移動しない
 *
 * 最後に
 * showPayrollPreflightOverallResult()
 * で総合判定する。
 * ===================================================== */


/* =====================================================
 * 29-A. 種別別 実行入口
 * ===================================================== */

function runPayrollPreflightOrder() {

  return cfPayrollRunPreflightForSource_(
    "SRC_ORDER_SALES"
  );

}


function runPayrollPreflightAttendance() {

  return cfPayrollRunPreflightForSource_(
    "SRC_CAST_ATTENDANCE"
  );

}


function runPayrollPreflightPerformance() {

  return cfPayrollRunPreflightForSource_(
    "SRC_CAST_PERFORMANCE"
  );

}


function runPayrollPreflightSettlement() {

  return cfPayrollRunPreflightForSource_(
    "SRC_SETTLEMENT"
  );

}


function runPayrollPreflightSlip() {

  return cfPayrollRunPreflightForSource_(
    "SRC_SLIP_LIST"
  );

}


/* =====================================================
 * 29-B. 1種類分の総合安全テスト
 * ===================================================== */

function cfPayrollRunPreflightForSource_(
  targetSourceId
) {

  const sourceId =
    String(
      targetSourceId || ""
    ).trim();


  if (!sourceId) {

    throw new Error(
      "取込設定IDが指定されていません。"
    );

  }


  const source =
    cfPayrollGetImportSourceById_(
      sourceId
    );


  if (!source) {

    throw new Error(
      "有効な取込設定が見つかりません：" +
      sourceId
    );

  }


  const inbox =
    cfPayrollGetInboxFolder_();


  const sources =
    cfPayrollGetImportSources_();


  const completedFileIds =
    cfPayrollPreflightCompletedFileIds_();


  /* =================================================
   * 実行前状態
   * ================================================= */

  const inboxCountBefore =
    cfPayrollPreflightCountFiles_(
      inbox
    );


  const factsCountBefore =
    cfPayrollPreflightCountSheetDataRows_(
      "PAY_FACTS"
    );


  const importLogCountBefore =
    cfPayrollPreflightCountSheetDataRows_(
      "PAY_IMPORT_LOG"
    );


  /* =================================================
   * 集計
   * ================================================= */

  const stat = {

    sourceId:
      sourceId,

    displayName:
      String(
        source.displayName || ""
      ),

    parserMode:
      String(
        source.parserMode || ""
      )
        .trim()
        .toUpperCase(),

    matchedFiles:
      0,

    previewSuccess:
      0,

    previewErrors:
      0,

    sourceRows:
      0,

    recognizedRows:
      0,

    zeroRecognizedFiles:
      0,

    blankDateRows:
      0,

    unknownCastCount:
      0,

    unknownProductCount:
      0,

    vipRows:
      0,

    settlementUnsafeFacts:
      0,

    paymentMismatchRows:
      0,

    otherDetailRows:
      0,

    ignoredRows:
      0,

    alreadyImportedFiles:
      0

  };


  const errorSamples =
    [];


  /* =================================================
   * ファイル走査
   * ================================================= */

  const files =
    inbox.getFiles();


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    const fileInfo =
      cfPayrollBuildImportFileInfo_(
        file
      );


    const detection =
      cfPayrollDetectImportSourceForFile_(
        fileInfo,
        sources
      );


    /*
     * 今回対象の種類以外は触らない
     */
    if (
      detection.status !==
        "MATCHED" ||
      String(
        detection.sourceId || ""
      ).trim() !== sourceId
    ) {

      continue;

    }


    stat.matchedFiles++;


    /* =================================================
     * 既取込確認
     * ================================================= */

    if (
      completedFileIds[
        fileInfo.id
      ]
    ) {

      stat.alreadyImportedFiles++;

    }


    /* =================================================
     * プレビュー
     * ================================================= */

    try {

      /*
       * ★ 正しい引数
       *
       * 第1引数：
       * DriveファイルID
       *
       * 第2引数：
       * sourceId
       */
      const preview =
        cfPayrollPreviewImport_(
          fileInfo.id,
          sourceId
        );


      if (
        !preview ||
        preview.success !== true
      ) {

        throw new Error(
          "プレビュー結果がsuccess=trueではありません。"
        );

      }


      /*
       * ★ 前回抜けていた箇所
       *
       * 必ずここでsummaryを宣言する。
       */
      const summary =
        preview.summary ||
        {};


      stat.previewSuccess++;


      /* =================================================
       * 基本件数
       * ================================================= */

      const sourceRows =
        cfPayrollPreflightNumber_(
          summary.sourceRows
        );


      const recognizedRows =
        cfPayrollPreflightNumber_(
          summary.recognizedRows
        );


      stat.sourceRows +=
        sourceRows;


      stat.recognizedRows +=
        recognizedRows;


      stat.ignoredRows +=
        cfPayrollPreflightNumber_(
          summary.ignoredRows
        );


      if (
        recognizedRows <= 0
      ) {

        stat.zeroRecognizedFiles++;

      }


      /* =================================================
       * 営業日
       *
       * Parser内部ログではdate:""でも、
       * cfPayrollPreviewImport_後に
       * Section 19で営業日が設定される。
       *
       * ここでは最終previewを確認する。
       * ================================================= */

      const dateMethod =
        String(
          source.dateMethod ||
          "NONE"
        )
          .trim()
          .toUpperCase();


      let actualBlankDateRows =
        0;


      if (
        dateMethod !==
          "NONE"
      ) {

        (
          preview.rows ||
          []
        )
          .forEach(function(
            row
          ) {

            if (
              row.date === "" ||
              row.date === null ||
              row.date === undefined
            ) {

              actualBlankDateRows++;

            }

          });

      }


      const blankDateRows =
        Math.max(

          actualBlankDateRows,

          cfPayrollPreflightNumber_(
            summary.blankDateRows
          )

        );


      stat.blankDateRows +=
        blankDateRows;


      /* =================================================
       * 未登録キャスト
       *
       * SETTLEMENT_DETAILは監査専用なので
       * unknownCastCountを本取込ブロックに使わない。
       * ================================================= */

      const unknownCasts =
        Array.isArray(
          preview.unknownCasts
        )
          ? preview.unknownCasts
          : [];


      const unknownCastCount =
        Math.max(

          cfPayrollPreflightNumber_(
            summary.unknownCastCount
          ),

          unknownCasts.length

        );


      stat.unknownCastCount +=
        unknownCastCount;


      /* =================================================
       * 未知商品
       * ================================================= */

      const unknownProducts =
        Array.isArray(
          preview.unknownProducts
        )
          ? preview.unknownProducts
          : [];


      const unknownProductCount =
        Math.max(

          cfPayrollPreflightNumber_(
            summary.unknownProductCount
          ),

          cfPayrollPreflightNumber_(
            summary.unknownProductRows
          ),

          cfPayrollPreflightNumber_(
            summary.unmappedProductRows
          ),

          unknownProducts.length

        );


      stat.unknownProductCount +=
        unknownProductCount;


      /* =================================================
       * VIP
       * ================================================= */

      stat.vipRows +=
        cfPayrollPreflightNumber_(
          summary.vipRows
        );


      /* =================================================
       * 精算入金安全性
       * ================================================= */

      if (
        stat.parserMode ===
        "SETTLEMENT_DETAIL"
      ) {

        let unsafeFacts =
          0;


        (
          preview.rows ||
          []
        )
          .forEach(function(
            row
          ) {

            (
              row.facts ||
              []
            )
              .forEach(function(
                fact
              ) {

                const itemKey =
                  String(
                    fact.itemKey || ""
                  ).trim();


                if (
                  itemKey &&
                  itemKey.indexOf(
                    "AUDIT_"
                  ) !== 0
                ) {

                  unsafeFacts++;

                }

              });

          });


        stat.settlementUnsafeFacts +=
          unsafeFacts;


        stat.paymentMismatchRows +=
          cfPayrollPreflightNumber_(
            summary.paymentMismatchRows
          );


        stat.otherDetailRows +=
          cfPayrollPreflightNumber_(
            summary.otherDetailRows
          );

      }


      Logger.log(
        "PRECHECK OK: " +
        file.getName() +
        " / " +
        sourceId
      );

    } catch (
      error
    ) {

      stat.previewErrors++;


      if (
        errorSamples.length < 10
      ) {

        errorSamples.push({

          fileName:
            file.getName(),

          error:
            String(
              error &&
              error.message
                ? error.message
                : error
            )

        });

      }


      Logger.log(
        "PRECHECK ERROR: " +
        file.getName() +
        " / " +
        String(
          error &&
          error.message
            ? error.message
            : error
        )
      );

    }

  }


  /* =================================================
   * 実行後状態
   * ================================================= */

  const inboxCountAfter =
    cfPayrollPreflightCountFiles_(
      inbox
    );


  const factsCountAfter =
    cfPayrollPreflightCountSheetDataRows_(
      "PAY_FACTS"
    );


  const importLogCountAfter =
    cfPayrollPreflightCountSheetDataRows_(
      "PAY_IMPORT_LOG"
    );


  /* =================================================
   * ブロッカー判定
   * ================================================= */

  const blockers =
    [];


  const warnings =
    [];


  if (
    stat.matchedFiles <= 0
  ) {

    blockers.push(
      "対象ファイルが0件です。"
    );

  }


  if (
    stat.previewErrors > 0
  ) {

    blockers.push(
      "プレビューエラー：" +
      stat.previewErrors +
      "件"
    );

  }


  if (
    stat.zeroRecognizedFiles > 0
  ) {

    blockers.push(
      "取込可能データ0件のファイル：" +
      stat.zeroRecognizedFiles +
      "件"
    );

  }


  if (
    stat.blankDateRows > 0
  ) {

    blockers.push(
      "営業日未確定行：" +
      stat.blankDateRows +
      "件"
    );

  }


  if (
    stat.parserMode !==
      "SETTLEMENT_DETAIL" &&
    stat.unknownCastCount > 0
  ) {

    blockers.push(
      "未登録キャスト：" +
      stat.unknownCastCount +
      "件"
    );

  }


  if (
    stat.unknownProductCount > 0
  ) {

    blockers.push(
      "未知商品：" +
      stat.unknownProductCount +
      "件"
    );

  }


  if (
    stat.alreadyImportedFiles > 0
  ) {

    blockers.push(
      "既取込ファイル：" +
      stat.alreadyImportedFiles +
      "件"
    );

  }


  if (
    stat.settlementUnsafeFacts > 0
  ) {

    blockers.push(
      "精算入金にAUDIT_以外のFact：" +
      stat.settlementUnsafeFacts +
      "件"
    );

  }


  if (
    stat.paymentMismatchRows > 0
  ) {

    blockers.push(
      "精算入金差額：" +
      stat.paymentMismatchRows +
      "件"
    );

  }


  if (
    stat.otherDetailRows > 0
  ) {

    warnings.push(
      "精算入金OTHER_DETAIL：" +
      stat.otherDetailRows +
      "件"
    );

  }


  /* =================================================
   * PREVIEW安全性
   * ================================================= */

  if (
    inboxCountBefore !==
      inboxCountAfter
  ) {

    blockers.push(
      "PREVIEW中に取込待ちフォルダのファイル数が変化しました。"
    );

  }


  if (
    factsCountBefore !==
      factsCountAfter
  ) {

    blockers.push(
      "PREVIEW中にPAY_FACTSの行数が変化しました。"
    );

  }


  if (
    importLogCountBefore !==
      importLogCountAfter
  ) {

    blockers.push(
      "PREVIEW中にPAY_IMPORT_LOGの行数が変化しました。"
    );

  }


  const status =
    blockers.length === 0
      ? "PASS"
      : "BLOCKED";


  /* =================================================
   * 保存する結果
   *
   * 詳細79件は保存せず
   * 集計だけ保存する。
   * ================================================= */

  const result = {

    success:
      true,

    status:
      status,

    sourceId:
      sourceId,

    runAt:
      new Date().toISOString(),

    stat:
      stat,

    integrity: {

      inboxBefore:
        inboxCountBefore,

      inboxAfter:
        inboxCountAfter,

      factsBefore:
        factsCountBefore,

      factsAfter:
        factsCountAfter,

      importLogBefore:
        importLogCountBefore,

      importLogAfter:
        importLogCountAfter

    },

    blockers:
      blockers,

    warnings:
      warnings,

    errorSamples:
      errorSamples

  };


  cfPayrollSavePreflightResult_(
    sourceId,
    result
  );


  Logger.log(
    "========================================"
  );


  Logger.log(
    "給与安全テスト " +
    sourceId +
    " : " +
    status
  );


  Logger.log(
    JSON.stringify(
      result
    )
  );


  Logger.log(
    "========================================"
  );


  return result;

}


/* =====================================================
 * 29-C. 種別結果保存
 * ===================================================== */

function cfPayrollSavePreflightResult_(
  sourceId,
  result
) {

  PropertiesService
    .getScriptProperties()
    .setProperty(
      "CF_PAYROLL_PREFLIGHT_" +
        sourceId,
      JSON.stringify(
        result
      )
    );

}


/* =====================================================
 * 29-D. 種別結果読込
 * ===================================================== */

function cfPayrollLoadPreflightResult_(
  sourceId
) {

  const value =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "CF_PAYROLL_PREFLIGHT_" +
          sourceId
      );


  if (!value) {

    return null;

  }


  try {

    return JSON.parse(
      value
    );

  } catch (
    error
  ) {

    return null;

  }

}


/* =====================================================
 * 29-E. 総合判定
 * ===================================================== */

function showPayrollPreflightOverallResult() {

  const sourceIds = [

    "SRC_ORDER_SALES",
    "SRC_CAST_ATTENDANCE",
    "SRC_CAST_PERFORMANCE",
    "SRC_SETTLEMENT",
    "SRC_SLIP_LIST"

  ];


  const results =
    {};


  const missing =
    [];


  const blockers =
    [];


  const warnings =
    [];


  sourceIds.forEach(function(
    sourceId
  ) {

    const result =
      cfPayrollLoadPreflightResult_(
        sourceId
      );


    if (!result) {

      missing.push(
        sourceId
      );

      return;

    }


    results[
      sourceId
    ] =
      result;


    if (
      result.status !==
      "PASS"
    ) {

      blockers.push(
        sourceId +
        " がBLOCKED"
      );

    }


    (
      result.warnings ||
      []
    )
      .forEach(function(
        warning
      ) {

        warnings.push(
          sourceId +
          " / " +
          warning
        );

      });

  });


  if (
    missing.length > 0
  ) {

    blockers.push(
      "未実行の安全テスト：" +
      missing.join(
        ", "
      )
    );

  }


  /* =================================================
   * VIP確認
   * ================================================= */

  const slipResult =
    results[
      "SRC_SLIP_LIST"
    ];


  if (
    slipResult &&
    slipResult.stat &&
    Number(
      slipResult.stat.vipRows || 0
    ) <= 0
  ) {

    warnings.push(
      "伝票一覧でVIP実データが0件です。"
    );

  }


  /* =================================================
   * ファイル数差
   * ================================================= */

  const fileCounts =
    {};


  sourceIds.forEach(function(
    sourceId
  ) {

    const result =
      results[
        sourceId
      ];


    if (
      result &&
      result.stat
    ) {

      fileCounts[
        sourceId
      ] =
        Number(
          result.stat.matchedFiles || 0
        );

    }

  });


  const finalStatus =
    blockers.length === 0
      ? "READY_FOR_IMPORT"
      : (
          missing.length > 0
            ? "INCOMPLETE"
            : "BLOCKED"
        );


  const overall = {

    success:
      true,

    finalStatus:
      finalStatus,

    fileCounts:
      fileCounts,

    blockers:
      blockers,

    warnings:
      warnings,

    results:
      results

  };


  Logger.log(
    "========================================"
  );


  Logger.log(
    "給与総合安全テスト結果: " +
    finalStatus
  );


  Logger.log(
    JSON.stringify(
      overall
    )
  );


  Logger.log(
    "========================================"
  );


  return overall;

}


/* =====================================================
 * 29-F. 保存済みテスト結果リセット
 * ===================================================== */

function resetPayrollPreflightResults() {

  const sourceIds = [

    "SRC_ORDER_SALES",
    "SRC_CAST_ATTENDANCE",
    "SRC_CAST_PERFORMANCE",
    "SRC_SETTLEMENT",
    "SRC_SLIP_LIST"

  ];


  const properties =
    PropertiesService
      .getScriptProperties();


  sourceIds.forEach(function(
    sourceId
  ) {

    properties.deleteProperty(
      "CF_PAYROLL_PREFLIGHT_" +
        sourceId
    );

  });


  const result = {

    success:
      true,

    cleared:
      sourceIds.length

  };


  Logger.log(
    "給与安全テスト結果リセットOK: " +
    JSON.stringify(
      result
    )
  );


  return result;

}


/* =====================================================
 * 29-G. 数値安全変換
 * ===================================================== */

function cfPayrollPreflightNumber_(
  value
) {

  const number =
    Number(
      value || 0
    );


  return Number.isFinite(
    number
  )
    ? number
    : 0;

}


/* =====================================================
 * 29-H. フォルダ内ファイル数
 * ===================================================== */

function cfPayrollPreflightCountFiles_(
  folder
) {

  let count =
    0;


  const files =
    folder.getFiles();


  while (
    files.hasNext()
  ) {

    files.next();

    count++;

  }


  return count;

}


/* =====================================================
 * 29-I. シートデータ行数
 * ===================================================== */

function cfPayrollPreflightCountSheetDataRows_(
  sheetName
) {

  const ss =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    return 0;

  }


  return Math.max(
    0,
    sheet.getLastRow() - 1
  );

}


/* =====================================================
 * 29-J. 既取込完了ファイルID
 * ===================================================== */

function cfPayrollPreflightCompletedFileIds_() {

  const result =
    {};


  const ss =
    cfPayrollOpenSpreadsheet_();


  const sheet =
    ss.getSheetByName(
      "PAY_IMPORT_LOG"
    );


  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    return result;

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  const headers =
    values[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  const fileIdIndex =
    headers.indexOf(
      "ファイルID"
    );


  const statusIndex =
    headers.indexOf(
      "処理状態"
    );


  if (
    fileIdIndex === -1 ||
    statusIndex === -1
  ) {

    return result;

  }


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const fileId =
      String(
        values[i][
          fileIdIndex
        ] || ""
      ).trim();


    const status =
      String(
        values[i][
          statusIndex
        ] || ""
      ).trim();


    if (
      fileId &&
      status ===
        "取込完了"
    ) {

      result[
        fileId
      ] =
        true;

    }

  }


  return result;

}


/* =====================================================
 * 30. 営業日別 5種類ファイル照合
 *
 * PAY_IMPORT_QUEUEで確定済みの営業日と
 * 01_取込待ちの実ファイルを
 * ファイルIDで照合する。
 *
 * Excel本体は開かない。
 *
 * PAY_FACTS：
 *   書き込まない
 *
 * PAY_IMPORT_QUEUE：
 *   変更しない
 *
 * ファイル：
 *   移動しない
 *
 * 出力：
 *   PAY_DATE_AUDIT
 * ===================================================== */


/* =====================================================
 * 30-A. 営業日別照合 実行
 * ===================================================== */

function buildPayrollBusinessDateAudit() {

  const expectedSources = [

    "SRC_ORDER_SALES",
    "SRC_CAST_ATTENDANCE",
    "SRC_CAST_PERFORMANCE",
    "SRC_SETTLEMENT",
    "SRC_SLIP_LIST"

  ];


  const sourceLabels = {

    SRC_ORDER_SALES:
      "ORDER",

    SRC_CAST_ATTENDANCE:
      "ATTENDANCE",

    SRC_CAST_PERFORMANCE:
      "PERFORMANCE",

    SRC_SETTLEMENT:
      "SETTLEMENT",

    SRC_SLIP_LIST:
      "SLIP"

  };


  const payrollSs =
    cfPayrollOpenSpreadsheet_();


  const timezone =
    payrollSs.getSpreadsheetTimeZone() ||
    "Asia/Tokyo";


  /* =================================================
   * PAY_IMPORT_QUEUE 読込
   * ================================================= */

  const queueSheet =
    payrollSs.getSheetByName(
      "PAY_IMPORT_QUEUE"
    );


  if (!queueSheet) {

    throw new Error(
      "PAY_IMPORT_QUEUE が見つかりません。"
    );

  }


  if (
    queueSheet.getLastRow() < 2
  ) {

    throw new Error(
      "PAY_IMPORT_QUEUE にデータがありません。"
    );

  }


  const queueRange =
    queueSheet.getDataRange();


  const queueValues =
    queueRange.getValues();


  const queueDisplayValues =
    queueRange.getDisplayValues();


  const queueHeaders =
    queueDisplayValues[0]
      .map(function(
        value
      ) {

        return String(
          value || ""
        ).trim();

      });


  const queueFileIdIndex =
    queueHeaders.indexOf(
      "ファイルID"
    );


  const queueSourceIdIndex =
    queueHeaders.indexOf(
      "取込設定ID"
    );


  const queueBusinessDateIndex =
    queueHeaders.indexOf(
      "営業日"
    );


  if (
    queueFileIdIndex === -1
  ) {

    throw new Error(
      "PAY_IMPORT_QUEUE に「ファイルID」列がありません。"
    );

  }


  if (
    queueSourceIdIndex === -1
  ) {

    throw new Error(
      "PAY_IMPORT_QUEUE に「取込設定ID」列がありません。"
    );

  }


  if (
    queueBusinessDateIndex === -1
  ) {

    throw new Error(
      "PAY_IMPORT_QUEUE に「営業日」列がありません。"
    );

  }


  /* =================================================
   * ファイルID → Queue行
   *
   * 同じファイルIDが複数存在する場合も
   * 検出できるよう配列で保持。
   * ================================================= */

  const queueByFileId =
    {};


  for (
    let rowIndex = 1;
    rowIndex < queueValues.length;
    rowIndex++
  ) {

    const fileId =
      String(
        queueDisplayValues[
          rowIndex
        ][
          queueFileIdIndex
        ] || ""
      ).trim();


    if (!fileId) {

      continue;

    }


    const sourceId =
      String(
        queueDisplayValues[
          rowIndex
        ][
          queueSourceIdIndex
        ] || ""
      ).trim();


    const businessDate =
      cfPayrollNormalizeBusinessDateForAudit_(

        queueValues[
          rowIndex
        ][
          queueBusinessDateIndex
        ],

        queueDisplayValues[
          rowIndex
        ][
          queueBusinessDateIndex
        ],

        timezone

      );


    if (
      !queueByFileId[
        fileId
      ]
    ) {

      queueByFileId[
        fileId
      ] =
        [];

    }


    queueByFileId[
      fileId
    ].push({

      queueRow:
        rowIndex + 1,

      sourceId:
        sourceId,

      businessDate:
        businessDate

    });

  }


  /* =================================================
   * 取込待ちフォルダ走査
   * ================================================= */

  const inbox =
    cfPayrollGetInboxFolder_();


  const sources =
    cfPayrollGetImportSources_();


  const files =
    inbox.getFiles();


  const dateMatrix =
    {};


  const seenInboxFileIds =
    {};


  const fileIssues =
    [];


  let totalFiles =
    0;


  let matchedFiles =
    0;


  let unmatchedFiles =
    0;


  let ambiguousFiles =
    0;


  let unsupportedFiles =
    0;


  let queueMissingFiles =
    0;


  let queueDuplicateFiles =
    0;


  let unresolvedDateFiles =
    0;


  let sourceMismatchFiles =
    0;


  while (
    files.hasNext()
  ) {

    const file =
      files.next();


    totalFiles++;


    const fileInfo =
      cfPayrollBuildImportFileInfo_(
        file
      );


    seenInboxFileIds[
      fileInfo.id
    ] =
      true;


    const detection =
      cfPayrollDetectImportSourceForFile_(
        fileInfo,
        sources
      );


    /* =================================================
     * MATCHED以外
     * ================================================= */

    if (
      detection.status !==
      "MATCHED"
    ) {

      if (
        detection.status ===
        "UNMATCHED"
      ) {

        unmatchedFiles++;

      } else if (
        detection.status ===
        "AMBIGUOUS"
      ) {

        ambiguousFiles++;

      } else {

        unsupportedFiles++;

      }


      fileIssues.push({

        fileName:
          file.getName(),

        issue:
          String(
            detection.status ||
            "UNKNOWN"
          )

      });


      continue;

    }


    matchedFiles++;


    const detectedSourceId =
      String(
        detection.sourceId || ""
      ).trim();


    /* =================================================
     * Queue確認
     * ================================================= */

    const queueRows =
      queueByFileId[
        fileInfo.id
      ] ||
      [];


    if (
      queueRows.length === 0
    ) {

      queueMissingFiles++;


      fileIssues.push({

        fileName:
          file.getName(),

        sourceId:
          detectedSourceId,

        issue:
          "QUEUE_MISSING"

      });


      continue;

    }


    if (
      queueRows.length > 1
    ) {

      queueDuplicateFiles++;


      fileIssues.push({

        fileName:
          file.getName(),

        sourceId:
          detectedSourceId,

        issue:
          "QUEUE_DUPLICATE",

        queueRows:
          queueRows.map(function(
            row
          ) {

            return row.queueRow;

          })

      });


      continue;

    }


    const queueInfo =
      queueRows[0];


    /* =================================================
     * Queueの取込設定IDと
     * 実ファイル判定が一致するか
     * ================================================= */

    if (
      queueInfo.sourceId &&
      queueInfo.sourceId !==
        detectedSourceId
    ) {

      sourceMismatchFiles++;


      fileIssues.push({

        fileName:
          file.getName(),

        detectedSourceId:
          detectedSourceId,

        queueSourceId:
          queueInfo.sourceId,

        issue:
          "SOURCE_MISMATCH"

      });


      continue;

    }


    /* =================================================
     * 営業日確定確認
     * ================================================= */

    if (
      !queueInfo.businessDate
    ) {

      unresolvedDateFiles++;


      fileIssues.push({

        fileName:
          file.getName(),

        sourceId:
          detectedSourceId,

        queueRow:
          queueInfo.queueRow,

        issue:
          "BUSINESS_DATE_UNRESOLVED"

      });


      continue;

    }


    const businessDate =
      queueInfo.businessDate;


    /* =================================================
     * 営業日マトリクス
     * ================================================= */

    if (
      !dateMatrix[
        businessDate
      ]
    ) {

      dateMatrix[
        businessDate
      ] =
        {};

    }


    if (
      !dateMatrix[
        businessDate
      ][
        detectedSourceId
      ]
    ) {

      dateMatrix[
        businessDate
      ][
        detectedSourceId
      ] =
        [];

    }


    dateMatrix[
      businessDate
    ][
      detectedSourceId
    ].push({

      fileId:
        fileInfo.id,

      fileName:
        file.getName(),

      queueRow:
        queueInfo.queueRow

    });

  }


  /* =================================================
   * QueueにあるがInboxにない行
   *
   * 今回は警告扱い。
   * ================================================= */

  const orphanQueueRows =
    [];


  Object.keys(
    queueByFileId
  )
    .forEach(function(
      fileId
    ) {

      if (
        seenInboxFileIds[
          fileId
        ]
      ) {

        return;

      }


      (
        queueByFileId[
          fileId
        ] ||
        []
      )
        .forEach(function(
          row
        ) {

          if (
            expectedSources.indexOf(
              row.sourceId
            ) === -1
          ) {

            return;

          }


          orphanQueueRows.push({

            fileId:
              fileId,

            sourceId:
              row.sourceId,

            businessDate:
              row.businessDate,

            queueRow:
              row.queueRow

          });

        });

    });


  /* =================================================
   * 営業日別チェック
   * ================================================= */

  const businessDates =
    Object.keys(
      dateMatrix
    )
      .sort();


  const auditRows =
    [];


  const missingByDate =
    [];


  const duplicateByDate =
    [];


  businessDates.forEach(function(
    businessDate
  ) {

    const matrix =
      dateMatrix[
        businessDate
      ];


    const outputRow = {

      businessDate:
        businessDate,

      statusBySource:
        {},

      overall:
        "OK",

      notes:
        []

    };


    expectedSources.forEach(function(
      sourceId
    ) {

      const records =
        matrix[
          sourceId
        ] ||
        [];


      let status;


      if (
        records.length === 0
      ) {

        status =
          "MISSING";


        outputRow.overall =
          "BLOCKED";


        outputRow.notes.push(
          sourceLabels[
            sourceId
          ] +
          "なし"
        );


        missingByDate.push({

          businessDate:
            businessDate,

          sourceId:
            sourceId,

          label:
            sourceLabels[
              sourceId
            ]

        });

      } else if (
        records.length === 1
      ) {

        status =
          "OK";

      } else {

        status =
          "DUPLICATE(" +
          records.length +
          ")";


        outputRow.overall =
          "BLOCKED";


        outputRow.notes.push(
          sourceLabels[
            sourceId
          ] +
          "重複" +
          records.length +
          "件"
        );


        duplicateByDate.push({

          businessDate:
            businessDate,

          sourceId:
            sourceId,

          label:
            sourceLabels[
              sourceId
            ],

          count:
            records.length,

          files:
            records.map(function(
              record
            ) {

              return record.fileName;

            })

        });

      }


      outputRow.statusBySource[
        sourceId
      ] =
        status;

    });


    auditRows.push(
      outputRow
    );

  });


  /* =================================================
   * PAY_DATE_AUDIT 出力
   * ================================================= */

  let auditSheet =
    payrollSs.getSheetByName(
      "PAY_DATE_AUDIT"
    );


  if (!auditSheet) {

    auditSheet =
      payrollSs.insertSheet(
        "PAY_DATE_AUDIT"
      );

  }


  auditSheet.clearContents();


  const outputValues = [

    [
      "営業日",
      "ORDER",
      "ATTENDANCE",
      "PERFORMANCE",
      "SETTLEMENT",
      "SLIP",
      "総合判定",
      "備考"
    ]

  ];


  auditRows.forEach(function(
    row
  ) {

    outputValues.push([

      row.businessDate,

      row.statusBySource[
        "SRC_ORDER_SALES"
      ],

      row.statusBySource[
        "SRC_CAST_ATTENDANCE"
      ],

      row.statusBySource[
        "SRC_CAST_PERFORMANCE"
      ],

      row.statusBySource[
        "SRC_SETTLEMENT"
      ],

      row.statusBySource[
        "SRC_SLIP_LIST"
      ],

      row.overall,

      row.notes.join(
        " / "
      )

    ]);

  });


  auditSheet
    .getRange(
      1,
      1,
      outputValues.length,
      outputValues[0].length
    )
    .setValues(
      outputValues
    );


  auditSheet.setFrozenRows(
    1
  );


  auditSheet.autoResizeColumns(
    1,
    outputValues[0].length
  );


  /* =================================================
   * 最終判定
   * ================================================= */

  const blockers =
    [];


  const warnings =
    [];


  if (
    unmatchedFiles > 0
  ) {

    blockers.push(
      "UNMATCHED：" +
      unmatchedFiles +
      "件"
    );

  }


  if (
    ambiguousFiles > 0
  ) {

    blockers.push(
      "AMBIGUOUS：" +
      ambiguousFiles +
      "件"
    );

  }


  if (
    unsupportedFiles > 0
  ) {

    blockers.push(
      "UNSUPPORTED：" +
      unsupportedFiles +
      "件"
    );

  }


  if (
    queueMissingFiles > 0
  ) {

    blockers.push(
      "PAY_IMPORT_QUEUE未登録：" +
      queueMissingFiles +
      "件"
    );

  }


  if (
    queueDuplicateFiles > 0
  ) {

    blockers.push(
      "PAY_IMPORT_QUEUE重複：" +
      queueDuplicateFiles +
      "件"
    );

  }


  if (
    unresolvedDateFiles > 0
  ) {

    blockers.push(
      "営業日未確定：" +
      unresolvedDateFiles +
      "件"
    );

  }


  if (
    sourceMismatchFiles > 0
  ) {

    blockers.push(
      "取込設定ID不一致：" +
      sourceMismatchFiles +
      "件"
    );

  }


  if (
    missingByDate.length > 0
  ) {

    blockers.push(
      "営業日別MISSING：" +
      missingByDate.length +
      "件"
    );

  }


  if (
    duplicateByDate.length > 0
  ) {

    blockers.push(
      "営業日別DUPLICATE：" +
      duplicateByDate.length +
      "件"
    );

  }


  if (
    orphanQueueRows.length > 0
  ) {

    warnings.push(
      "Inboxに存在しないQueue行：" +
      orphanQueueRows.length +
      "件"
    );

  }


  const finalStatus =
    blockers.length === 0
      ? "PASS"
      : "BLOCKED";


  const result = {

    success:
      true,

    finalStatus:
      finalStatus,

    summary: {

      totalFiles:
        totalFiles,

      matchedFiles:
        matchedFiles,

      businessDateCount:
        businessDates.length,

      unmatchedFiles:
        unmatchedFiles,

      ambiguousFiles:
        ambiguousFiles,

      unsupportedFiles:
        unsupportedFiles,

      queueMissingFiles:
        queueMissingFiles,

      queueDuplicateFiles:
        queueDuplicateFiles,

      unresolvedDateFiles:
        unresolvedDateFiles,

      sourceMismatchFiles:
        sourceMismatchFiles,

      missingByDateCount:
        missingByDate.length,

      duplicateByDateCount:
        duplicateByDate.length,

      orphanQueueRows:
        orphanQueueRows.length

    },

    missingByDate:
      missingByDate,

    duplicateByDate:
      duplicateByDate,

    fileIssues:
      fileIssues,

    blockers:
      blockers,

    warnings:
      warnings

  };


  Logger.log(
    "========================================"
  );


  Logger.log(
    "営業日別5種類照合結果: " +
    finalStatus
  );


  Logger.log(
    JSON.stringify(
      result
    )
  );


  Logger.log(
    "========================================"
  );


  return result;

}


/* =====================================================
 * 30-B. 営業日正規化
 *
 * 安全のため
 *
 * yyyy/MM/dd
 * yyyy-MM-dd
 * yyyy.MM.dd
 *
 * またはDate型だけ受け付ける。
 *
 * MM/ddだけ等は推測しない。
 * ===================================================== */

function cfPayrollNormalizeBusinessDateForAudit_(
  rawValue,
  displayValue,
  timezone
) {

  if (
    rawValue instanceof Date &&
    !isNaN(
      rawValue.getTime()
    )
  ) {

    return Utilities.formatDate(
      rawValue,
      timezone || "Asia/Tokyo",
      "yyyy-MM-dd"
    );

  }


  const text =
    String(
      displayValue ||
      rawValue ||
      ""
    ).trim();


  if (!text) {

    return "";

  }


  const match =
    text.match(
      /^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/
    );


  if (!match) {

    return "";

  }


  const year =
    Number(
      match[1]
    );


  const month =
    Number(
      match[2]
    );


  const day =
    Number(
      match[3]
    );


  if (
    year < 2000 ||
    year > 2100 ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {

    return "";

  }


  /*
   * 実在日チェック
   */
  const testDate =
    new Date(
      Date.UTC(
        year,
        month - 1,
        day
      )
    );


  if (
    testDate.getUTCFullYear() !==
      year ||
    testDate.getUTCMonth() !==
      month - 1 ||
    testDate.getUTCDate() !==
      day
  ) {

    return "";

  }


  return (
    String(
      year
    ) +
    "-" +
    String(
      month
    ).padStart(
      2,
      "0"
    ) +
    "-" +
    String(
      day
    ).padStart(
      2,
      "0"
    )
  );

}