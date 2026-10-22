/* =====================================================
 * CAST FLOW
 * SystemTest.gs
 * 管理者向けテスト集
 *
 * Apps Script上部の関数一覧から、日本語名のテストを選択して実行。
 * テスト02 / 03 はLINE送信あり。
 * ===================================================== */

const CF_TEST_DATE = "2099-12-31";
const CF_TEST_ID_PREFIX = "TEST_WEBHOOK_";
const CF_TEST_NOTE = "CAST FLOW 自動テスト専用行";


/* =====================================================
 * テスト01：全体診断（送信なし・更新なし）
 * ===================================================== */
function テスト01_全体診断() {
  const p = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const result = {
    実行日時: Utilities.formatDate(new Date(), cfTestTimezone_(), "yyyy/MM/dd HH:mm:ss"),

    秘密情報: {
      LINE送信用アクセストークン: p.getProperty("LINE_CHANNEL_ACCESS_TOKEN") ? "設定済み" : "未設定",
      LINEチャネルシークレット_GAS側: p.getProperty("LINE_CHANNEL_SECRET") ? "設定済み" : "未設定（構成によりGAS側では未使用の場合あり）",
      Cloudflare_GAS共有鍵: p.getProperty("CLOUDFLARE_WEBHOOK_SECRET") ? "設定済み" : "未設定【出勤OKの保存に必要】"
    },

    シート構成: { 正常: [], 不足: [] },
    主要機能: { 正常: [], 不足: [] },
    トリガー: {},
    SETTINGS: {},
    管理者権限基盤: {},
    本日の出勤確認: {}
  };

  if (typeof CONFIG === "undefined" || !CONFIG || !CONFIG.SHEETS) {
    result.シート構成.不足.push("CONFIG.SHEETS が見つかりません");
  } else {
    Object.keys(CONFIG.SHEETS).forEach(key => {
      const name = String(CONFIG.SHEETS[key] || "").trim();
      if (name && ss.getSheetByName(name)) {
        result.シート構成.正常.push(key + " : " + name);
      } else {
        result.シート構成.不足.push(key + " : " + (name || "未設定"));
      }
    });
  }

  const functions = [
    ["Web API受付", typeof doPost === "function"],
    ["月間シフト取得", typeof getMonthlyShiftData === "function"],
    ["シフト一括保存", typeof saveMultiShiftsByLine === "function"],
    ["月間シフト提出", typeof submitMonthlyShiftByLine === "function"],
    ["Webhook出勤OK処理", typeof confirmAttendanceOkByWebhook === "function"],
    ["出勤確認LINE生成", typeof pushAttendanceConfirmationMessage_ === "function"],
    ["管理者LINE送信", typeof sendAdminLineMessage_ === "function"],
    ["月間提出状況", typeof getMonthlySubmissionStatus_ === "function"],
    ["本日の出勤確認", typeof getAdminTodayAttendance_ === "function"],
    ["同伴・来店予定取得", typeof getCustomerVisitsForCastDate_ === "function"],
    ["エラーログ", typeof writeErrorLog_ === "function"]
  ];

  functions.forEach(([name, ok]) => {
    (ok ? result.主要機能.正常 : result.主要機能.不足).push(name);
  });

  result.トリガー = cfTestTriggerStatus_();
  result.SETTINGS = cfTestSettingsStatus_();
  result.管理者権限基盤 = cfTestAdminRoleStatus_();
  result.本日の出勤確認 = cfTestTodayAttendanceStatus_();

  cfTestLog_("CAST FLOW 全体診断", result);
  return result;
}


/* =====================================================
 * テスト02：管理者LINE送信
 * GAS → LINE の送信確認
 * ===================================================== */
function テスト02_管理者LINE送信() {
  if (typeof sendAdminLineMessage_ !== "function") {
    throw new Error("sendAdminLineMessage_ が見つかりません");
  }

  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("LINE_CHANNEL_ACCESS_TOKEN")) {
    throw new Error("LINE_CHANNEL_ACCESS_TOKEN が未設定です");
  }

  const message = [
    "CAST FLOW",
    "",
    "【管理者LINE送信テスト】",
    "",
    "このメッセージが届けば、",
    "GAS → LINE の送信経路は正常です。",
    "",
    Utilities.formatDate(new Date(), cfTestTimezone_(), "yyyy/MM/dd HH:mm:ss")
  ].join("\n");

  sendAdminLineMessage_(message);

  const result = { success: true, 結果: "管理者LINEへテストメッセージを送信しました" };
  cfTestLog_("管理者LINE送信", result);
  return result;
}


/* =====================================================
 * テスト03：出勤確認Webhookテスト送信
 * LINE → Cloudflare → GAS → Sheets を確認
 * ===================================================== */
function テスト03_出勤確認を送る() {
  if (typeof pushAttendanceConfirmationMessage_ !== "function") {
    throw new Error("pushAttendanceConfirmationMessage_ が見つかりません");
  }
  if (typeof confirmAttendanceOkByWebhook !== "function") {
    throw new Error("confirmAttendanceOkByWebhook が見つかりません");
  }

  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("LINE_CHANNEL_ACCESS_TOKEN")) {
    throw new Error("LINE_CHANNEL_ACCESS_TOKEN が未設定です");
  }
  if (!p.getProperty("CLOUDFLARE_WEBHOOK_SECRET")) {
    throw new Error("CLOUDFLARE_WEBHOOK_SECRET が本番GASに未設定です");
  }

  const admin = cfTestGetAdminCast_();
  if (!admin) {
    throw new Error("管理者一覧の有効管理者と紐づくCASTが見つかりません");
  }

  const sheet = cfTestSheet_("ATTENDANCE_CONFIRMATIONS");
  cfTestDeleteRows_(sheet);

  sheet.appendRow([
    CF_TEST_ID_PREFIX + Utilities.getUuid(),
    cfTestParseDate_(CF_TEST_DATE),
    admin.castId,
    admin.castName,
    admin.lineUserId,
    "確認待ち",
    "",
    new Date(),
    CF_TEST_NOTE
  ]);

  const message = [
    "CAST FLOW",
    "",
    "【出勤確認 Webhookテスト】",
    "",
    "管理者専用テストです。",
    "本日の出勤データには影響しません。",
    "",
    "下の『出勤OK』だけ押してください。",
    "",
    "押した後にApps Scriptで",
    "『テスト04_出勤確認結果』を実行してください。"
  ].join("\n");

  pushAttendanceConfirmationMessage_(admin.lineUserId, message, CF_TEST_DATE);

  const result = {
    success: true,
    送信先: admin.castName,
    テスト日: CF_TEST_DATE,
    次の操作: "LINEで出勤OKを押す → テスト04_出勤確認結果"
  };

  cfTestLog_("出勤確認Webhookテスト送信", result);
  return result;
}


/* =====================================================
 * テスト04：出勤確認Webhook結果
 * ===================================================== */
function テスト04_出勤確認結果() {
  const sheet = cfTestSheet_("ATTENDANCE_CONFIRMATIONS");
  const rows = cfTestRows_(sheet);

  if (rows.length === 0) {
    const result = {
      success: false,
      状態: "テスト行がありません",
      案内: "先にテスト03_出勤確認を送る を実行してください"
    };
    cfTestLog_("出勤確認Webhook結果", result);
    return result;
  }

  const row = rows[rows.length - 1].values;
  const status = String(row[5] || "").trim();

  const result = {
    success: status === "出勤OK",
    状態: status || "空欄",
    確認日時: cfTestFormatDateTime_(row[6]),
    更新日時: cfTestFormatDateTime_(row[7]),
    判定: status === "出勤OK"
      ? "正常：LINE → Cloudflare → GAS → Sheets が通っています"
      : "未完了：共有鍵・Worker・GASデプロイ・LINE設定を確認"
  };

  cfTestLog_("出勤確認Webhook結果", result);
  return result;
}


/* =====================================================
 * テスト05：出勤確認テストデータ削除
 * TEST_WEBHOOK_ の行だけ削除
 * ===================================================== */
function テスト05_出勤確認テストデータ削除() {
  const sheet = cfTestSheet_("ATTENDANCE_CONFIRMATIONS");
  const deleted = cfTestDeleteRows_(sheet);
  const result = { success: true, 削除件数: deleted };
  cfTestLog_("出勤確認テストデータ削除", result);
  return result;
}


/* =====================================================
 * テスト06：本日の出勤確認状態（送信なし）
 * ===================================================== */
function テスト06_本日の出勤確認状態() {
  const result = cfTestTodayAttendanceStatus_();
  cfTestLog_("本日の出勤確認状態", result);
  return result;
}


/* =====================================================
 * テスト07：管理者権限分離確認（送信なし・更新なし）
 * 管理者一覧が権限の正本として使える状態か確認
 * ===================================================== */
function テスト07_管理者権限分離確認() {
  const result = cfTestAdminRoleStatus_();
  cfTestLog_("管理者権限分離確認", result);
  return result;
}


/* =====================================================
 * 共通処理
 * ===================================================== */
function cfTestTriggerStatus_() {
  const handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const expected = [
    "sendTodayAttendanceConfirmations",
    "resendTodayPendingAttendanceConfirmations",
    "sendTodayAttendanceAdminSummary"
  ];
  const status = {};
  expected.forEach(name => status[name] = handlers.includes(name) ? "設定あり" : "見つかりません");
  return { 登録数: handlers.length, 登録関数: handlers, 出勤確認関連: status };
}

function cfTestSettingsStatus_() {
  if (typeof getShiftSettings !== "function") {
    return { 正常: false, エラー: "getShiftSettings が見つかりません" };
  }
  try {
    const s = getShiftSettings();
    return {
      正常: true,
      出勤時間候補数: Array.isArray(s.startHourList) ? s.startHourList.length : 0,
      退勤時間候補数: Array.isArray(s.endHourList) ? s.endHourList.length : 0,
      分刻み: s.minuteStep,
      LAST: s.enableLast ? "有効" : "無効"
    };
  } catch (e) {
    return { 正常: false, エラー: e.message || String(e) };
  }
}

function cfTestTodayAttendanceStatus_() {
  try {
    const shiftSheet = cfTestSheet_("SHIFT_CURRENT");
    const confirmSheet = cfTestSheet_("ATTENDANCE_CONFIRMATIONS");
    const today = Utilities.formatDate(new Date(), cfTestTimezone_(), "yyyy-MM-dd");

    const working = new Set();
    if (shiftSheet.getLastRow() >= 2) {
      shiftSheet.getRange(2, 1, shiftSheet.getLastRow() - 1, 13).getValues().forEach(row => {
        if (
          cfTestFormatDate_(row[2]) === today &&
          String(row[4] || "").trim() &&
          String(row[6] || "").trim() === "出勤"
        ) {
          working.add(String(row[4]).trim());
        }
      });
    }

    const latest = {};
    if (confirmSheet.getLastRow() >= 2) {
      confirmSheet.getRange(2, 1, confirmSheet.getLastRow() - 1, 9).getValues().forEach(row => {
        const castId = String(row[2] || "").trim();
        if (cfTestFormatDate_(row[1]) === today && castId) {
          latest[castId] = String(row[5] || "").trim();
        }
      });
    }

    const count = { 確認待ち: 0, 出勤OK: 0, 変更あり: 0, 確認行未作成: 0 };
    working.forEach(castId => {
      const s = latest[castId];
      if (s === "出勤OK") count.出勤OK++;
      else if (s === "変更あり") count.変更あり++;
      else if (s === "確認待ち") count.確認待ち++;
      else count.確認行未作成++;
    });

    return {
      success: true,
      対象日: today,
      出勤予定: working.size,
      確認待ち: count.確認待ち,
      出勤OK: count.出勤OK,
      変更あり: count.変更あり,
      確認行未作成: count.確認行未作成
    };
  } catch (e) {
    return { success: false, エラー: e.message || String(e) };
  }
}

function cfTestAdminRoleStatus_() {
  try {
    const checks = {
      CONFIG_ADMIN_MASTER: !!(
        typeof CONFIG !== "undefined" &&
        CONFIG &&
        CONFIG.SHEETS &&
        CONFIG.SHEETS.ADMIN_MASTER
      ),
      管理者検索関数: typeof getAdminByLineUserId_ === "function",
      管理者認証関数: typeof requireAdmin_ === "function"
    };

    if (!checks.CONFIG_ADMIN_MASTER) {
      return {
        success: false,
        判定: "NG",
        エラー: "CONFIG.SHEETS.ADMIN_MASTER が見つかりません",
        確認項目: checks
      };
    }

    const adminSheet = cfTestSheet_("ADMIN_MASTER");
    const lastRow = adminSheet.getLastRow();
    const lastColumn = adminSheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return {
        success: false,
        判定: "NG",
        エラー: "管理者一覧に管理者データがありません",
        確認項目: checks
      };
    }

    const headers = adminSheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(v => String(v || "").trim());

    const idx = {
      id: headers.indexOf("管理者ID"),
      name: headers.indexOf("管理者名"),
      line: headers.indexOf("LINEユーザーID"),
      role: headers.indexOf("権限"),
      active: headers.indexOf("有効状態")
    };

    const missing = [];
    [
      ["管理者ID", idx.id],
      ["管理者名", idx.name],
      ["LINEユーザーID", idx.line],
      ["権限", idx.role],
      ["有効状態", idx.active]
    ].forEach(([name, index]) => {
      if (index < 0) missing.push(name);
    });

    if (missing.length) {
      return {
        success: false,
        判定: "NG",
        エラー: "管理者一覧の必要な見出しが不足しています：" + missing.join(" / "),
        確認項目: checks
      };
    }

    const values = adminSheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .getValues();

    let target = null;

    for (const row of values) {
      const lineUserId = String(row[idx.line] || "").trim();
      const activeStatus = String(row[idx.active] || "").trim();
      const isActive = typeof isAdminActiveStatus_ === "function"
        ? isAdminActiveStatus_(row[idx.active])
        : ["有効", "TRUE", "1", "YES", "ACTIVE", "ENABLED"]
            .includes(activeStatus.toUpperCase());

      if (lineUserId && isActive) {
        target = {
          adminId: String(row[idx.id] || "").trim(),
          adminName: String(row[idx.name] || "").trim(),
          lineUserId,
          role: String(row[idx.role] || "").trim(),
          activeStatus
        };
        break;
      }
    }

    if (!target) {
      return {
        success: false,
        判定: "NG",
        エラー: "有効な管理者が管理者一覧に見つかりません",
        確認項目: checks
      };
    }

    const admin = getAdminByLineUserId_(target.lineUserId);
    const cast = typeof getCastByLineUserId_ === "function"
      ? getCastByLineUserId_(target.lineUserId)
      : null;

    const adminLookupOk = !!(
      admin &&
      admin.adminId === target.adminId &&
      admin.isActive === true
    );

    const castCompatibilityOk = !cast || cast.isAdmin === true;

    return {
      success: adminLookupOk && castCompatibilityOk,
      判定: adminLookupOk && castCompatibilityOk ? "OK" : "NG",
      権限の正本: CONFIG.SHEETS.ADMIN_MASTER,
      有効管理者: {
        管理者ID: target.adminId,
        管理者名: target.adminName,
        権限: target.role,
        有効状態: target.activeStatus
      },
      管理者一覧から検索: adminLookupOk ? "正常" : "異常",
      キャスト互換isAdmin: cast
        ? (castCompatibilityOk ? "正常" : "異常")
        : "同一LINEのキャスト登録なし",
      注意: "requireAdmin_ の実IDトークン認証はDEV LIFF実機確認で最終確認します",
      確認項目: checks
    };

  } catch (e) {
    return {
      success: false,
      判定: "NG",
      エラー: e.message || String(e)
    };
  }
}


function cfTestGetAdminCast_() {
  const adminSheet = cfTestSheet_("ADMIN_MASTER");
  const lastRow = adminSheet.getLastRow();
  const lastColumn = adminSheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return null;
  }

  const headers = adminSheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());

  const idx = {
    id: headers.indexOf("管理者ID"),
    name: headers.indexOf("管理者名"),
    line: headers.indexOf("LINEユーザーID"),
    active: headers.indexOf("有効状態")
  };

  if (idx.id < 0 || idx.name < 0 || idx.line < 0 || idx.active < 0) {
    throw new Error("管理者一覧の必要な見出しが不足しています");
  }

  const values = adminSheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  for (const row of values) {
    const lineUserId = String(row[idx.line] || "").trim();
    const active = typeof isAdminActiveStatus_ === "function"
      ? isAdminActiveStatus_(row[idx.active])
      : String(row[idx.active] || "").trim() === "有効";

    if (!lineUserId || !active) {
      continue;
    }

    const cast = typeof getCastByLineUserId_ === "function"
      ? getCastByLineUserId_(lineUserId)
      : null;

    if (!cast || !cast.castId || !cast.castName) {
      continue;
    }

    return {
      castId: cast.castId,
      castName: cast.castName,
      lineUserId: lineUserId,
      adminId: String(row[idx.id] || "").trim(),
      adminName: String(row[idx.name] || "").trim()
    };
  }

  return null;
}


function cfTestTruthy_(value) {
  if (value === true || value === 1) return true;
  return ["TRUE", "1", "YES", "Y", "管理者", "あり", "○", "〇"]
    .includes(String(value || "").trim().toUpperCase());
}

function cfTestSheet_(key) {
  if (typeof CONFIG === "undefined" || !CONFIG || !CONFIG.SHEETS || !CONFIG.SHEETS[key]) {
    throw new Error("CONFIG.SHEETS." + key + " が見つかりません");
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS[key]);
  if (!sheet) throw new Error(CONFIG.SHEETS[key] + " シートが見つかりません");
  return sheet;
}

function cfTestRows_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const rows = [];
  values.forEach((row, i) => {
    if (String(row[0] || "").trim().startsWith(CF_TEST_ID_PREFIX)) {
      rows.push({ rowNumber: i + 2, values: row });
    }
  });
  return rows;
}

function cfTestDeleteRows_(sheet) {
  const rows = cfTestRows_(sheet).map(x => x.rowNumber).sort((a, b) => b - a);
  rows.forEach(r => sheet.deleteRow(r));
  return rows.length;
}

function cfTestParseDate_(text) {
  const m = String(text || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error("テスト日付が不正です");
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function cfTestFormatDate_(value) {
  if (value === null || value === undefined || value === "") return "";
  const d = Object.prototype.toString.call(value) === "[object Date]" ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value || "").trim();
  return Utilities.formatDate(d, cfTestTimezone_(), "yyyy-MM-dd");
}

function cfTestFormatDateTime_(value) {
  if (value === null || value === undefined || value === "") return "";
  const d = Object.prototype.toString.call(value) === "[object Date]" ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value || "");
  return Utilities.formatDate(d, cfTestTimezone_(), "yyyy/MM/dd HH:mm:ss");
}

function cfTestTimezone_() {
  return (typeof CONFIG !== "undefined" && CONFIG && CONFIG.TIMEZONE)
    ? CONFIG.TIMEZONE
    : "Asia/Tokyo";
}

function cfTestLog_(title, value) {
  Logger.log("========================================");
  Logger.log("【" + title + "】");
  Logger.log(JSON.stringify(value, null, 2));
  Logger.log("========================================");
}
