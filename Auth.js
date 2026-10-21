/* =====================================================
 * ROYAL SHIFT SYSTEM
 * Auth.gs
 * ===================================================== */
/* =====================================================
 * 01. LINE Login 設定
 * ===================================================== */
/**
 * LINE Login Channel ID
 *
 * ROYAL SHIFT
 */
const LINE_LOGIN_CHANNEL_ID =
  "2011488723";
/* =====================================================
 * 02. LINE IDトークン検証
 * ===================================================== */
/**
 * LIFFから受け取ったIDトークンを
 * LINE公式APIで検証する。
 *
 * @param {string} idToken
 * @return {Object}
 */
function verifyLineIdToken_(
  idToken
) {
  const token =
    String(
      idToken || ""
    ).trim();
  if (!token) {
    throw new Error(
      "LINE認証情報がありません"
    );
  }
  const response =
    UrlFetchApp.fetch(
      "https://api.line.me/oauth2/v2.1/verify",
      {
        method:
          "post",
        contentType:
          "application/x-www-form-urlencoded",
        payload: {
          id_token:
            token,
          client_id:
            LINE_LOGIN_CHANNEL_ID
        },
        muteHttpExceptions:
          true
      }
    );
  const statusCode =
    response.getResponseCode();
  const responseText =
    response.getContentText();
  let data = {};
  try {
    data =
      JSON.parse(
        responseText
      );
  } catch (error) {
    throw new Error(
      "LINE認証レスポンスの解析に失敗しました"
    );
  }
  if (
    statusCode !== 200
  ) {
    const message =
      data.error_description ||
      data.error ||
      "LINE認証に失敗しました";
    throw new Error(
      message
    );
  }
  /*
   * aud は
   * このIDトークンが発行された
   * LINE Login Channel ID
   */
  if (
    String(
      data.aud || ""
    ) !==
    LINE_LOGIN_CHANNEL_ID
  ) {
    throw new Error(
      "LINE Login Channel IDが一致しません"
    );
  }
  /*
   * sub が
   * LINEユーザーID
   */
  const lineUserId =
    String(
      data.sub || ""
    ).trim();
  if (!lineUserId) {
    throw new Error(
      "LINEユーザーIDを取得できませんでした"
    );
  }
  return {
    success:
      true,
    lineUserId:
      lineUserId,
    displayName:
      String(
        data.name || ""
      ).trim(),
    pictureUrl:
      String(
        data.picture || ""
      ).trim()
  };
}
/* =====================================================
 * 03. LINEユーザーIDからCAST / 管理者検索
 * ===================================================== */

/**
 * LINEユーザーIDから
 * キャスト一覧を検索する。
 *
 * 管理者判定はキャスト一覧の「管理者権限」列ではなく、
 * 管理者一覧を参照する。
 *
 * 既存画面との互換性維持のため、返却値の isAdmin は残す。
 *
 * @param {string} lineUserId
 * @return {Object|null}
 */
function getCastByLineUserId_(
  lineUserId
) {

  const normalizedLineUserId =
    String(
      lineUserId || ""
    ).trim();

  if (!normalizedLineUserId) {
    return null;
  }

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

  const lineDisplayNameIndex =
    findCastHeaderIndex_(
      headers,
      [
        "LINE表示名"
      ]
    );

  const lineUserIdIndex =
    findCastHeaderIndex_(
      headers,
      [
        "LINEユーザーID"
      ]
    );

  const activeStatusIndex =
    findCastHeaderIndex_(
      headers,
      [
        "在籍状況"
      ]
    );

  const registeredAtIndex =
    findCastHeaderIndex_(
      headers,
      [
        "登録日"
      ]
    );

  if (
    lineUserIdIndex < 0
  ) {
    throw new Error(
      CONFIG.SHEETS.CAST +
      "シートにLINEユーザーID列がありません"
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

  for (
    let index = 0;
    index < values.length;
    index++
  ) {

    const row =
      values[index];

    const rowLineUserId =
      String(
        row[lineUserIdIndex] || ""
      ).trim();

    if (
      rowLineUserId !==
      normalizedLineUserId
    ) {
      continue;
    }

    const admin =
      getAdminByLineUserId_(
        normalizedLineUserId
      );

    return {

      rowNumber:
        index + 2,

      castId:
        castIdIndex >= 0
          ? String(
              row[castIdIndex] || ""
            ).trim()
          : "",

      castName:
        castNameIndex >= 0
          ? String(
              row[castNameIndex] || ""
            ).trim()
          : "",

      lineDisplayName:
        lineDisplayNameIndex >= 0
          ? String(
              row[lineDisplayNameIndex] || ""
            ).trim()
          : "",

      lineUserId:
        rowLineUserId,

      activeStatus:
        activeStatusIndex >= 0
          ? String(
              row[activeStatusIndex] || ""
            ).trim()
          : "",

      registeredAt:
        registeredAtIndex >= 0
          ? row[registeredAtIndex]
          : "",

      /*
       * 互換用。
       * 権限の正本は管理者一覧。
       */
      isAdmin:
        Boolean(
          admin &&
          admin.isActive
        )

    };
  }

  return null;

}


/**
 * LINEユーザーIDから
 * 管理者一覧を検索する。
 *
 * @param {string} lineUserId
 * @return {Object|null}
 */
function getAdminByLineUserId_(
  lineUserId
) {

  const normalizedLineUserId =
    String(
      lineUserId || ""
    ).trim();

  if (!normalizedLineUserId) {
    return null;
  }

  if (
    !CONFIG.SHEETS ||
    !CONFIG.SHEETS.ADMIN_MASTER
  ) {
    throw new Error(
      "CONFIG.SHEETS.ADMIN_MASTERが設定されていません"
    );
  }

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.ADMIN_MASTER
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

  const adminIdIndex =
    headers.indexOf(
      "管理者ID"
    );

  const adminNameIndex =
    headers.indexOf(
      "管理者名"
    );

  const lineDisplayNameIndex =
    headers.indexOf(
      "LINE表示名"
    );

  const lineUserIdIndex =
    headers.indexOf(
      "LINEユーザーID"
    );

  const roleIndex =
    headers.indexOf(
      "権限"
    );

  const notificationIndex =
    headers.indexOf(
      "通知受信"
    );

  const activeStatusIndex =
    headers.indexOf(
      "有効状態"
    );

  const registeredAtIndex =
    headers.indexOf(
      "登録日時"
    );

  const updatedAtIndex =
    headers.indexOf(
      "更新日時"
    );

  const noteIndex =
    headers.indexOf(
      "備考"
    );

  const requiredHeaders = [
    ["管理者ID", adminIdIndex],
    ["管理者名", adminNameIndex],
    ["LINEユーザーID", lineUserIdIndex],
    ["権限", roleIndex],
    ["有効状態", activeStatusIndex]
  ];

  const missingHeaders =
    requiredHeaders
      .filter(
        item =>
          item[1] < 0
      )
      .map(
        item =>
          item[0]
      );

  if (
    missingHeaders.length
  ) {
    throw new Error(
      CONFIG.SHEETS.ADMIN_MASTER +
      "シートに必要な列がありません：" +
      missingHeaders.join(" / ")
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

  for (
    let index = 0;
    index < values.length;
    index++
  ) {

    const row =
      values[index];

    const rowLineUserId =
      String(
        row[lineUserIdIndex] || ""
      ).trim();

    if (
      rowLineUserId !==
      normalizedLineUserId
    ) {
      continue;
    }

    const activeStatus =
      activeStatusIndex >= 0
        ? row[activeStatusIndex]
        : "";

    return {

      rowNumber:
        index + 2,

      adminId:
        String(
          row[adminIdIndex] || ""
        ).trim(),

      adminName:
        String(
          row[adminNameIndex] || ""
        ).trim(),

      lineDisplayName:
        lineDisplayNameIndex >= 0
          ? String(
              row[lineDisplayNameIndex] || ""
            ).trim()
          : "",

      lineUserId:
        rowLineUserId,

      role:
        String(
          row[roleIndex] || ""
        ).trim(),

      notificationsEnabled:
        notificationIndex >= 0
          ? convertBooleanValue_(
              row[notificationIndex]
            )
          : false,

      activeStatus:
        String(
          activeStatus || ""
        ).trim(),

      isActive:
        isAdminActiveStatus_(
          activeStatus
        ),

      registeredAt:
        registeredAtIndex >= 0
          ? row[registeredAtIndex]
          : "",

      updatedAt:
        updatedAtIndex >= 0
          ? row[updatedAtIndex]
          : "",

      note:
        noteIndex >= 0
          ? String(
              row[noteIndex] || ""
            ).trim()
          : ""

    };
  }

  return null;

}


/**
 * 管理者一覧の有効状態を判定する。
 *
 * @param {*} value
 * @return {boolean}
 */
function isAdminActiveStatus_(
  value
) {

  if (
    value === true
  ) {
    return true;
  }

  const text =
    String(
      value || ""
    )
      .trim()
      .toUpperCase();

  return (
    text === "有効" ||
    text === "TRUE" ||
    text === "1" ||
    text === "YES" ||
    text === "ACTIVE" ||
    text === "ENABLED"
  );

}


/* =====================================================
 * 04. 名前からCAST検索
 * ===================================================== */
/**
 * 入力された名前と
 * CASTシートのキャスト名を照合する。
 *
 * 完全一致のみ。
 *
 * @param {string} castName
 * @return {Object|null}
 */
function getCastByName_(
  castName
) {
  const normalizedName =
    normalizeCastName_(
      castName
    );
  if (!normalizedName) {
    return null;
  }
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
  const lineUserIdIndex =
    findCastHeaderIndex_(
      headers,
      [
        "LINEユーザーID"
      ]
    );
  if (
    castNameIndex < 0
  ) {
    throw new Error(
      "CASTシートにキャスト名列がありません"
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
  const matches = [];
  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    const row =
      values[index];
    const rowCastName =
      String(
        row[castNameIndex] || ""
      ).trim();
    if (
      normalizeCastName_(
        rowCastName
      ) !==
      normalizedName
    ) {
      continue;
    }
    matches.push({
      rowNumber:
        index + 2,
      castId:
        castIdIndex >= 0
          ? String(
              row[castIdIndex] || ""
            ).trim()
          : "",
      castName:
        rowCastName,
      lineUserId:
        lineUserIdIndex >= 0
          ? String(
              row[lineUserIdIndex] || ""
            ).trim()
          : ""
    });
  }
  /*
   * 同姓同名が複数いる場合は
   * 自動で決めない。
   */
  if (
    matches.length > 1
  ) {
    throw new Error(
      "同じ名前のキャストが複数登録されています"
    );
  }
  return (
    matches.length === 1
      ? matches[0]
      : null
  );
}
/* =====================================================
 * 05. 初期管理者登録
 * ===================================================== */

/**
 * 最初の管理者だけ登録する。
 *
 * 初期キャストは S001 固定。
 * 管理者権限は「管理者一覧」を正本として登録する。
 *
 * @param {string} idToken
 * @param {string} castName
 * @return {Object}
 */
function registerInitialAdmin_(
  idToken,
  castName
) {

  const verified =
    verifyLineIdToken_(
      idToken
    );

  let cast =
    getCastByLineUserId_(
      verified.lineUserId
    );

  if (!cast) {

    cast =
      getCastByName_(
        castName
      );

    if (!cast) {
      throw new Error(
        "入力された名前がキャスト一覧に登録されていません"
      );
    }

    if (
      cast.castId !==
      "S001"
    ) {
      throw new Error(
        "初期管理者として登録できません"
      );
    }

    if (
      cast.lineUserId &&
      cast.lineUserId !==
      verified.lineUserId
    ) {
      throw new Error(
        "S001はすでに別のLINEアカウントへ登録されています"
      );
    }

    bindInitialAdminCast_(
      cast,
      verified
    );

    cast =
      getCastByLineUserId_(
        verified.lineUserId
      );
  }

  if (
    !cast ||
    cast.castId !==
    "S001"
  ) {
    throw new Error(
      "初期管理者として登録できません"
    );
  }

  const admin =
    ensureInitialAdminMaster_(
      verified,
      cast
    );

  return {

    success:
      true,

    registered:
      true,

    cast: {

      castId:
        cast.castId,

      castName:
        cast.castName,

      lineDisplayName:
        cast.lineDisplayName ||
        verified.displayName,

      lineUserId:
        verified.lineUserId,

      isAdmin:
        true

    },

    admin:
      admin

  };

}


/**
 * 初期管理者のLINE情報をキャスト一覧へ紐付ける。
 * 「管理者権限」列は更新しない。
 */
function bindInitialAdminCast_(
  cast,
  verified
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.CAST
    );

  const lastColumn =
    sheet.getLastColumn();

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

  const lineDisplayNameIndex =
    headers.indexOf(
      "LINE表示名"
    );

  const lineUserIdIndex =
    headers.indexOf(
      "LINEユーザーID"
    );

  const activeStatusIndex =
    headers.indexOf(
      "在籍状況"
    );

  const registeredAtIndex =
    headers.indexOf(
      "登録日"
    );

  if (
    lineDisplayNameIndex < 0 ||
    lineUserIdIndex < 0 ||
    registeredAtIndex < 0
  ) {
    throw new Error(
      CONFIG.SHEETS.CAST +
      "シートの登録用列が不足しています"
    );
  }

  sheet
    .getRange(
      cast.rowNumber,
      lineDisplayNameIndex + 1
    )
    .setValue(
      verified.displayName
    );

  sheet
    .getRange(
      cast.rowNumber,
      lineUserIdIndex + 1
    )
    .setValue(
      verified.lineUserId
    );

  if (
    activeStatusIndex >= 0
  ) {
    sheet
      .getRange(
        cast.rowNumber,
        activeStatusIndex + 1
      )
      .setValue(
        "在籍"
      );
  }

  sheet
    .getRange(
      cast.rowNumber,
      registeredAtIndex + 1
    )
    .setValue(
      new Date()
    );

}


/**
 * 初期管理者を管理者一覧へ登録・更新する。
 */
function ensureInitialAdminMaster_(
  verified,
  cast
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.ADMIN_MASTER
    );

  const lastColumn =
    sheet.getLastColumn();

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

  const names = [
    "管理者ID",
    "管理者名",
    "LINE表示名",
    "LINEユーザーID",
    "権限",
    "通知受信",
    "有効状態",
    "登録日時",
    "更新日時",
    "備考"
  ];

  const indexes = {};

  names.forEach(
    name => {
      indexes[name] =
        headers.indexOf(
          name
        );
    }
  );

  const missingHeaders =
    names.filter(
      name =>
        indexes[name] < 0
    );

  if (
    missingHeaders.length
  ) {
    throw new Error(
      CONFIG.SHEETS.ADMIN_MASTER +
      "シートに必要な列がありません：" +
      missingHeaders.join(" / ")
    );
  }

  const lastRow =
    sheet.getLastRow();

  const values =
    lastRow >= 2
      ? sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            lastColumn
          )
          .getValues()
      : [];

  let rowNumber =
    -1;

  let currentRow =
    null;

  for (
    let index = 0;
    index < values.length;
    index++
  ) {

    const row =
      values[index];

    const rowAdminId =
      String(
        row[indexes["管理者ID"]] || ""
      ).trim();

    const rowAdminName =
      String(
        row[indexes["管理者名"]] || ""
      ).trim();

    const rowLineUserId =
      String(
        row[indexes["LINEユーザーID"]] || ""
      ).trim();

    if (
      rowLineUserId ===
        verified.lineUserId ||
      rowAdminId ===
        "A001" ||
      (
        normalizeCastName_(
          rowAdminName
        ) ===
        normalizeCastName_(
          cast.castName
        ) &&
        !rowLineUserId
      )
    ) {
      rowNumber =
        index + 2;
      currentRow =
        row;
      break;
    }
  }

  if (
    currentRow
  ) {

    const existingLineUserId =
      String(
        currentRow[indexes["LINEユーザーID"]] || ""
      ).trim();

    if (
      existingLineUserId &&
      existingLineUserId !==
      verified.lineUserId
    ) {
      throw new Error(
        "初期管理者にはすでに別のLINEアカウントが登録されています"
      );
    }
  }

  const now =
    new Date();

  const registeredAt =
    currentRow &&
    currentRow[indexes["登録日時"]]
      ? currentRow[indexes["登録日時"]]
      : now;

  const notificationValue =
    currentRow &&
    currentRow[indexes["通知受信"]] !== ""
      ? currentRow[indexes["通知受信"]]
      : true;

  const role =
    currentRow &&
    String(
      currentRow[indexes["権限"]] || ""
    ).trim()
      ? String(
          currentRow[indexes["権限"]]
        ).trim()
      : "システム管理者";

  const note =
    currentRow
      ? String(
          currentRow[indexes["備考"]] || ""
        ).trim()
      : "初期管理者";

  const rowValues =
    new Array(
      lastColumn
    ).fill("");

  rowValues[indexes["管理者ID"]] =
    currentRow
      ? String(
          currentRow[indexes["管理者ID"]] ||
          "A001"
        ).trim()
      : "A001";

  rowValues[indexes["管理者名"]] =
    cast.castName;

  rowValues[indexes["LINE表示名"]] =
    verified.displayName;

  rowValues[indexes["LINEユーザーID"]] =
    verified.lineUserId;

  rowValues[indexes["権限"]] =
    role;

  rowValues[indexes["通知受信"]] =
    notificationValue;

  rowValues[indexes["有効状態"]] =
    "有効";

  rowValues[indexes["登録日時"]] =
    registeredAt;

  rowValues[indexes["更新日時"]] =
    now;

  rowValues[indexes["備考"]] =
    note;

  if (
    rowNumber >= 2
  ) {
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastColumn
      )
      .setValues([
        rowValues
      ]);
  } else {
    sheet.appendRow(
      rowValues
    );
  }

  return getAdminByLineUserId_(
    verified.lineUserId
  );

}


/* =====================================================
 * 06. 認証状態確認
 * ===================================================== */

/**
 * LIFFから受け取ったIDトークンで
 * 現在の登録状態と利用可能ロールを確認する。
 *
 * 管理者権限は「管理者一覧」を正本とする。
 *
 * @param {string} idToken
 * @return {Object}
 */
function getLineAuthStatus_(
  idToken
) {

  const verified =
    verifyLineIdToken_(
      idToken
    );

  const cast =
    getCastByLineUserId_(
      verified.lineUserId
    );

  const admin =
    getAdminByLineUserId_(
      verified.lineUserId
    );

  const hasAdminRole =
    Boolean(
      admin &&
      admin.isActive
    );

  const availableModes = [];

  if (cast) {
    availableModes.push(
      "cast"
    );
  }

  if (hasAdminRole) {
    availableModes.push(
      "admin"
    );
  }

  /*
   * キャスト未登録
   *
   * 現在のフロント画面との互換性のため、
   * registered はキャスト登録状態を表す。
   * 管理者のみのアカウント対応は
   * モード切替UI追加時に有効化する。
   */
  if (!cast) {

    const pendingRequest =
      getPendingRegistrationRequest_(
        verified.lineUserId
      );

    return {

      success:
        true,

      registered:
        false,

      pending:
        Boolean(
          pendingRequest
        ),

      pendingRequest:
        pendingRequest,

      lineDisplayName:
        verified.displayName,

      hasCastRole:
        false,

      hasAdminRole:
        hasAdminRole,

      availableModes:
        availableModes,

      admin:
        hasAdminRole
          ? {
              adminId:
                admin.adminId,
              adminName:
                admin.adminName,
              role:
                admin.role,
              notificationsEnabled:
                admin.notificationsEnabled
            }
          : null

    };
  }

  return {

    success:
      true,

    registered:
      true,

    hasCastRole:
      true,

    hasAdminRole:
      hasAdminRole,

    availableModes:
      availableModes,

    cast: {

      castId:
        cast.castId,

      castName:
        cast.castName,

      lineDisplayName:
        cast.lineDisplayName ||
        verified.displayName,

      /*
       * 既存フロント互換用。
       * 値は管理者一覧から判定する。
       */
      isAdmin:
        hasAdminRole,

      admin:
        hasAdminRole

    },

    admin:
      hasAdminRole
        ? {
            adminId:
              admin.adminId,
            adminName:
              admin.adminName,
            role:
              admin.role,
            notificationsEnabled:
              admin.notificationsEnabled
          }
        : null

  };

}


/* =====================================================
 * 07. 共通処理
 * ===================================================== */
/**
 * 名前照合用
 */
function normalizeCastName_(
  value
) {
  return String(
    value || ""
  )
    .trim()
    .replace(
      /\s+/g,
      ""
    )
    .toLowerCase();
}
/**
 * TRUE/FALSE変換
 */
function convertBooleanValue_(
  value
) {
  if (
    value === true
  ) {
    return true;
  }
  const text =
    String(
      value || ""
    )
      .trim()
      .toUpperCase();
  return (
    text === "TRUE" ||
    text === "1" ||
    text === "YES" ||
    text === "あり"
  );
}
/* =====================================================
 * 08. 登録申請
 * ===================================================== */
/**
 * 未登録ユーザーの登録申請
 *
 * 新人CASTも申請可能。
 *
 * 本人が
 * ・キャスト番号
 * ・キャスト名
 * を入力。
 *
 * 申請作成後、
 * 管理者へボタン付きLINE通知を送信する。
 *
 * CAST / REGISTRATION_REQUESTSともに
 * 列番号ではなくヘッダー名で処理する。
 *
 * @param {string} idToken
 * @param {string} castId
 * @param {string} castName
 * @return {Object}
 */
function submitRegistrationRequest_(
  idToken,
  castId,
  castName
) {
  const verified =
    verifyLineIdToken_(
      idToken
    );
  const normalizedCastId =
    String(
      castId || ""
    ).trim();
  const normalizedName =
    String(
      castName || ""
    ).trim();
  if (!normalizedCastId) {
    throw new Error(
      "キャスト番号を入力してください"
    );
  }
  if (!normalizedName) {
    throw new Error(
      "キャスト名を入力してください"
    );
  }
  /* =================================================
   * このLINEがすでにCAST登録済みか
   * ================================================= */
  const existingLineCast =
    getCastByLineUserId_(
      verified.lineUserId
    );
  if (existingLineCast) {
    return {
      success:
        true,
      alreadyRegistered:
        true,
      cast:
        existingLineCast
    };
  }
  /* =================================================
   * CAST取得
   * ================================================= */
  const castSheet =
    getRequiredSheet_(
      CONFIG.SHEETS.CAST
    );
  const castLastColumn =
    castSheet.getLastColumn();
  if (
    castLastColumn < 1
  ) {
    throw new Error(
      "CASTシートの列がありません"
    );
  }
  const castHeaders =
    castSheet
      .getRange(
        1,
        1,
        1,
        castLastColumn
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
      castHeaders,
      [
        "キャストID",
        "CAST_ID"
      ]
    );
  const castNameIndex =
    findCastHeaderIndex_(
      castHeaders,
      [
        "キャスト名",
        "CAST_NAME"
      ]
    );
  const castLineUserIdIndex =
    findCastHeaderIndex_(
      castHeaders,
      [
        "LINEユーザーID"
      ]
    );
  if (
    castIdIndex < 0
  ) {
    throw new Error(
      "CASTシートにキャストID列がありません"
    );
  }
  if (
    castNameIndex < 0
  ) {
    throw new Error(
      "CASTシートにキャスト名列がありません"
    );
  }
  let existingCast =
    null;
  const castLastRow =
    castSheet.getLastRow();
  if (
    castLastRow >= 2
  ) {
    const castValues =
      castSheet
        .getRange(
          2,
          1,
          castLastRow - 1,
          castLastColumn
        )
        .getValues();
    for (
      let index = 0;
      index < castValues.length;
      index++
    ) {
      const row =
        castValues[index];
      const rowCastId =
        String(
          row[
            castIdIndex
          ] || ""
        ).trim();
      if (
        rowCastId !==
        normalizedCastId
      ) {
        continue;
      }
      existingCast = {
        rowNumber:
          index + 2,
        castId:
          rowCastId,
        castName:
          String(
            row[
              castNameIndex
            ] || ""
          ).trim(),
        lineUserId:
          castLineUserIdIndex >= 0
            ? String(
                row[
                  castLineUserIdIndex
                ] || ""
              ).trim()
            : ""
      };
      break;
    }
  }
  /* =================================================
   * 既存CASTなら名前照合
   * ================================================= */
  if (
    existingCast &&
    normalizeCastName_(
      existingCast.castName
    ) !==
    normalizeCastName_(
      normalizedName
    )
  ) {
    throw new Error(
      "キャスト番号とキャスト名が一致しません"
    );
  }
  /*
   * 既存CASTに別LINEが
   * すでに登録されている場合
   */
  if (
    existingCast &&
    existingCast.lineUserId &&
    existingCast.lineUserId !==
      verified.lineUserId
  ) {
    throw new Error(
      "このキャスト番号はすでにLINE登録されています"
    );
  }
  /* =================================================
   * REGISTRATION_REQUESTS
   * ================================================= */
  const requestSheet =
    getRequiredSheet_(
      CONFIG.SHEETS.REGISTRATION_REQUESTS
    );
  const requestLastColumn =
    requestSheet.getLastColumn();
  if (
    requestLastColumn < 1
  ) {
    throw new Error(
      "REGISTRATION_REQUESTSの列がありません"
    );
  }
  const requestHeaders =
    requestSheet
      .getRange(
        1,
        1,
        1,
        requestLastColumn
      )
      .getDisplayValues()[0]
      .map(
        value =>
          String(
            value || ""
          ).trim()
      );
  const requestIdIndex =
    requestHeaders.indexOf(
      "申請ID"
    );
  const requestedAtIndex =
    requestHeaders.indexOf(
      "申請日時"
    );
  const requestCastNameIndex =
    requestHeaders.indexOf(
      "入力名"
    );
  const requestLineDisplayNameIndex =
    requestHeaders.indexOf(
      "LINE表示名"
    );
  const requestLineUserIdIndex =
    requestHeaders.indexOf(
      "LINEユーザーID"
    );
  const requestStatusIndex =
    requestHeaders.indexOf(
      "状態"
    );
  const processedAtIndex =
    requestHeaders.indexOf(
      "処理日時"
    );
  const processedByIndex =
    requestHeaders.indexOf(
      "処理者"
    );
  const requestCastIdIndex =
    requestHeaders.indexOf(
      "CAST_ID"
    );
  const requestNoteIndex =
    requestHeaders.indexOf(
      "備考"
    );
  const missingHeaders =
    [];
  [
    [
      "申請ID",
      requestIdIndex
    ],
    [
      "申請日時",
      requestedAtIndex
    ],
    [
      "入力名",
      requestCastNameIndex
    ],
    [
      "LINE表示名",
      requestLineDisplayNameIndex
    ],
    [
      "LINEユーザーID",
      requestLineUserIdIndex
    ],
    [
      "状態",
      requestStatusIndex
    ],
    [
      "処理日時",
      processedAtIndex
    ],
    [
      "処理者",
      processedByIndex
    ],
    [
      "CAST_ID",
      requestCastIdIndex
    ],
    [
      "備考",
      requestNoteIndex
    ]
  ].forEach(
    ([name, index]) => {
      if (
        index < 0
      ) {
        missingHeaders.push(
          name
        );
      }
    }
  );
  if (
    missingHeaders.length
  ) {
    throw new Error(
      "REGISTRATION_REQUESTSに必要な列がありません：" +
      missingHeaders.join(
        " / "
      )
    );
  }
  /* =================================================
   * 承認待ち重複確認
   * ================================================= */
  const requestLastRow =
    requestSheet.getLastRow();
  if (
    requestLastRow >= 2
  ) {
    const requestValues =
      requestSheet
        .getRange(
          2,
          1,
          requestLastRow - 1,
          requestLastColumn
        )
        .getValues();
    for (
      let index = 0;
      index < requestValues.length;
      index++
    ) {
      const row =
        requestValues[index];
      const rowLineUserId =
        String(
          row[
            requestLineUserIdIndex
          ] || ""
        ).trim();
      const rowStatus =
        String(
          row[
            requestStatusIndex
          ] || ""
        ).trim();
      const rowCastId =
        String(
          row[
            requestCastIdIndex
          ] || ""
        ).trim();
      /*
       * 同じLINEの承認待ち
       *
       * 新しい通知は送らない。
       */
      if (
        rowLineUserId ===
          verified.lineUserId &&
        rowStatus ===
          "承認待ち"
      ) {
        return {
          success:
            true,
          pending:
            true,
          requestId:
            String(
              row[
                requestIdIndex
              ] || ""
            ).trim(),
          castId:
            rowCastId,
          castName:
            String(
              row[
                requestCastNameIndex
              ] || ""
            ).trim(),
          lineDisplayName:
            String(
              row[
                requestLineDisplayNameIndex
              ] || ""
            ).trim()
        };
      }
      /*
       * 同じCAST番号を
       * 別LINEが申請中
       */
      if (
        rowCastId ===
          normalizedCastId &&
        rowStatus ===
          "承認待ち"
      ) {
        throw new Error(
          "このキャスト番号は現在登録申請中です"
        );
      }
    }
  }
  /* =================================================
   * 新規申請作成
   * ================================================= */
  const requestId =
    createId(
      "REG"
    );
  const registrationType =
    existingCast
      ? "existing"
      : "new";
  const newRow =
    new Array(
      requestLastColumn
    ).fill(
      ""
    );
  newRow[
    requestIdIndex
  ] =
    requestId;
  newRow[
    requestedAtIndex
  ] =
    new Date();
  newRow[
    requestCastNameIndex
  ] =
    normalizedName;
  newRow[
    requestLineDisplayNameIndex
  ] =
    verified.displayName;
  newRow[
    requestLineUserIdIndex
  ] =
    verified.lineUserId;
  newRow[
    requestStatusIndex
  ] =
    "承認待ち";
  newRow[
    processedAtIndex
  ] =
    "";
  newRow[
    processedByIndex
  ] =
    "";
  newRow[
    requestCastIdIndex
  ] =
    normalizedCastId;
  newRow[
    requestNoteIndex
  ] =
    existingCast
      ? "既存CAST"
      : "新規CAST";
  requestSheet.appendRow(
    newRow
  );
  /* =================================================
   * 管理者へボタン付きLINE通知
   * ================================================= */
  try {
    sendAdminRegistrationRequestNotification_({
      requestId:
        requestId,
      castId:
        normalizedCastId,
      castName:
        normalizedName,
      lineDisplayName:
        verified.displayName,
      registrationType:
        registrationType
    });
  } catch (
    notificationError
  ) {
    /*
     * LINE通知だけ失敗しても
     * 登録申請自体は成功扱い
     */
    Logger.log(
      "登録申請管理者通知に失敗：" +
      (
        notificationError?.message ||
        String(
          notificationError
        )
      )
    );
  }
  /* =================================================
   * 完了
   * ================================================= */
  return {
    success:
      true,
    pending:
      true,
    requestId:
      requestId,
    castId:
      normalizedCastId,
    castName:
      normalizedName,
    lineDisplayName:
      verified.displayName,
    registrationType:
      registrationType
  };
}
/* =====================================================
 * 09. 登録申請状態確認
 * ===================================================== */
/**
 * LINEユーザーの承認待ち登録申請を取得する。
 *
 * REGISTRATION_REQUESTSは
 * 列番号ではなくヘッダー名を基準に処理する。
 *
 * @param {string} lineUserId
 * @return {Object|null}
 */
function getPendingRegistrationRequest_(
  lineUserId
) {
  const normalizedLineUserId =
    String(
      lineUserId || ""
    ).trim();
  if (!normalizedLineUserId) {
    return null;
  }
  /* =================================================
   * REGISTRATION_REQUESTS取得
   * ================================================= */
  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.REGISTRATION_REQUESTS
    );
  const lastColumn =
    sheet.getLastColumn();
  if (
    lastColumn < 1
  ) {
    throw new Error(
      "REGISTRATION_REQUESTSの列がありません"
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
        (value) =>
          String(
            value || ""
          ).trim()
      );
  /* =================================================
   * 列位置
   * ================================================= */
  const requestIdIndex =
    headers.indexOf(
      "申請ID"
    );
  const castNameIndex =
    headers.indexOf(
      "入力名"
    );
  const lineDisplayNameIndex =
    headers.indexOf(
      "LINE表示名"
    );
  const lineUserIdIndex =
    headers.indexOf(
      "LINEユーザーID"
    );
  const statusIndex =
    headers.indexOf(
      "状態"
    );
  const castIdIndex =
    headers.indexOf(
      "CAST_ID"
    );
  /* =================================================
   * 必須列チェック
   * ================================================= */
  const missingHeaders =
    [];
  if (
    requestIdIndex < 0
  ) {
    missingHeaders.push(
      "申請ID"
    );
  }
  if (
    castNameIndex < 0
  ) {
    missingHeaders.push(
      "入力名"
    );
  }
  if (
    lineDisplayNameIndex < 0
  ) {
    missingHeaders.push(
      "LINE表示名"
    );
  }
  if (
    lineUserIdIndex < 0
  ) {
    missingHeaders.push(
      "LINEユーザーID"
    );
  }
  if (
    statusIndex < 0
  ) {
    missingHeaders.push(
      "状態"
    );
  }
  if (
    castIdIndex < 0
  ) {
    missingHeaders.push(
      "CAST_ID"
    );
  }
  if (
    missingHeaders.length
  ) {
    throw new Error(
      "REGISTRATION_REQUESTSに必要な列がありません：" +
      missingHeaders.join(
        " / "
      )
    );
  }
  /* =================================================
   * データ取得
   * ================================================= */
  const lastRow =
    sheet.getLastRow();
  if (
    lastRow < 2
  ) {
    return null;
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
  /* =================================================
   * LINEユーザーの承認待ち申請を検索
   * ================================================= */
  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    const row =
      values[index];
    const rowLineUserId =
      String(
        row[
          lineUserIdIndex
        ] || ""
      ).trim();
    if (
      rowLineUserId !==
      normalizedLineUserId
    ) {
      continue;
    }
    const status =
      String(
        row[
          statusIndex
        ] || ""
      ).trim();
    if (
      status !==
      "承認待ち"
    ) {
      continue;
    }
    return {
      requestId:
        String(
          row[
            requestIdIndex
          ] || ""
        ).trim(),
      castId:
        String(
          row[
            castIdIndex
          ] || ""
        ).trim(),
      castName:
        String(
          row[
            castNameIndex
          ] || ""
        ).trim(),
      lineDisplayName:
        String(
          row[
            lineDisplayNameIndex
          ] || ""
        ).trim(),
      status:
        status
    };
  }
  return null;
}
/* =====================================================
 * 10. 管理者認証
 * ===================================================== */

/**
 * IDトークンから管理者本人を確認する。
 *
 * 管理者権限はキャスト一覧ではなく
 * 管理者一覧のみを参照する。
 *
 * @param {string} idToken
 * @return {Object}
 */
function requireAdmin_(
  idToken
) {

  const verified =
    verifyLineIdToken_(
      idToken
    );

  const admin =
    getAdminByLineUserId_(
      verified.lineUserId
    );

  if (!admin) {
    throw new Error(
      "管理者として登録されていません"
    );
  }

  if (!admin.isActive) {
    throw new Error(
      "管理者権限が無効になっています"
    );
  }

  const cast =
    getCastByLineUserId_(
      verified.lineUserId
    );

  return {

    verified:
      verified,

    admin:
      admin,

    cast:
      cast

  };

}


/* =====================================================
 * 11. 登録申請一覧取得
 * ===================================================== */
/**
 * 承認待ちの登録申請を取得する。
 *
 * REGISTRATION_REQUESTS / CAST ともに
 * 列番号ではなくヘッダー名を基準に処理する。
 *
 * 管理者専用。
 *
 * @param {string} idToken
 * @return {Object}
 */
function getRegistrationRequests_(
  idToken
) {
  requireAdmin_(
    idToken
  );
  /* =================================================
   * REGISTRATION_REQUESTS取得
   * ================================================= */
  const requestSheet =
    getRequiredSheet_(
      CONFIG.SHEETS.REGISTRATION_REQUESTS
    );
  const requestLastColumn =
    requestSheet.getLastColumn();
  if (
    requestLastColumn < 1
  ) {
    throw new Error(
      "REGISTRATION_REQUESTSの列がありません"
    );
  }
  const requestHeaders =
    requestSheet
      .getRange(
        1,
        1,
        1,
        requestLastColumn
      )
      .getDisplayValues()[0]
      .map(
        (value) =>
          String(
            value || ""
          ).trim()
      );
  /* =================================================
   * REGISTRATION_REQUESTS 列位置
   * ================================================= */
  const requestIdIndex =
    requestHeaders.indexOf(
      "申請ID"
    );
  const requestedAtIndex =
    requestHeaders.indexOf(
      "申請日時"
    );
  const requestedNameIndex =
    requestHeaders.indexOf(
      "入力名"
    );
  const lineDisplayNameIndex =
    requestHeaders.indexOf(
      "LINE表示名"
    );
  const statusIndex =
    requestHeaders.indexOf(
      "状態"
    );
  const requestCastIdIndex =
    requestHeaders.indexOf(
      "CAST_ID"
    );
  /* =================================================
   * 必須列チェック
   * ================================================= */
  const missingHeaders =
    [];
  if (
    requestIdIndex < 0
  ) {
    missingHeaders.push(
      "申請ID"
    );
  }
  if (
    requestedAtIndex < 0
  ) {
    missingHeaders.push(
      "申請日時"
    );
  }
  if (
    requestedNameIndex < 0
  ) {
    missingHeaders.push(
      "入力名"
    );
  }
  if (
    lineDisplayNameIndex < 0
  ) {
    missingHeaders.push(
      "LINE表示名"
    );
  }
  if (
    statusIndex < 0
  ) {
    missingHeaders.push(
      "状態"
    );
  }
  if (
    requestCastIdIndex < 0
  ) {
    missingHeaders.push(
      "CAST_ID"
    );
  }
  if (
    missingHeaders.length
  ) {
    throw new Error(
      "REGISTRATION_REQUESTSに必要な列がありません：" +
      missingHeaders.join(
        " / "
      )
    );
  }
  /* =================================================
   * 申請がない場合
   * ================================================= */
  const requestLastRow =
    requestSheet.getLastRow();
  if (
    requestLastRow < 2
  ) {
    return {
      success:
        true,
      requests:
        []
    };
  }
  /* =================================================
   * CAST取得
   * ================================================= */
  const castSheet =
    getRequiredSheet_(
      CONFIG.SHEETS.CAST
    );
  const castLastColumn =
    castSheet.getLastColumn();
  if (
    castLastColumn < 1
  ) {
    throw new Error(
      "CASTシートの列がありません"
    );
  }
  const castHeaders =
    castSheet
      .getRange(
        1,
        1,
        1,
        castLastColumn
      )
      .getDisplayValues()[0]
      .map(
        (value) =>
          String(
            value || ""
          ).trim()
      );
  const castIdIndex =
    castHeaders.indexOf(
      "キャストID"
    );
  const castNameIndex =
    castHeaders.indexOf(
      "キャスト名"
    );
  if (
    castIdIndex < 0 ||
    castNameIndex < 0
  ) {
    throw new Error(
      "CASTシートに必要な列がありません：キャストID / キャスト名"
    );
  }
  /* =================================================
   * CAST一覧を取得
   * ================================================= */
  const castMap =
    {};
  const castLastRow =
    castSheet.getLastRow();
  if (
    castLastRow >= 2
  ) {
    const castValues =
      castSheet
        .getRange(
          2,
          1,
          castLastRow - 1,
          castLastColumn
        )
        .getValues();
    for (
      let index = 0;
      index < castValues.length;
      index++
    ) {
      const row =
        castValues[index];
      const castId =
        String(
          row[
            castIdIndex
          ] || ""
        ).trim();
      if (!castId) {
        continue;
      }
      castMap[
        castId
      ] = {
        castId:
          castId,
        castName:
          String(
            row[
              castNameIndex
            ] || ""
          ).trim()
      };
    }
  }
  /* =================================================
   * 承認待ち申請を取得
   * ================================================= */
  const requestValues =
    requestSheet
      .getRange(
        2,
        1,
        requestLastRow - 1,
        requestLastColumn
      )
      .getValues();
  const requests =
    [];
  for (
    let index = 0;
    index < requestValues.length;
    index++
  ) {
    const row =
      requestValues[index];
    const status =
      String(
        row[
          statusIndex
        ] || ""
      ).trim();
    if (
      status !==
      "承認待ち"
    ) {
      continue;
    }
    const castId =
      String(
        row[
          requestCastIdIndex
        ] || ""
      ).trim();
    const requestedCastName =
      String(
        row[
          requestedNameIndex
        ] || ""
      ).trim();
    const existingCast =
      castMap[
        castId
      ] || null;
    const castExists =
      Boolean(
        existingCast
      );
    let nameMatches =
      null;
    if (
      existingCast
    ) {
      nameMatches =
        normalizeCastName_(
          existingCast.castName
        ) ===
        normalizeCastName_(
          requestedCastName
        );
    }
    requests.push({
      requestId:
        String(
          row[
            requestIdIndex
          ] || ""
        ).trim(),
      requestedAt:
        formatRegistrationDate_(
          row[
            requestedAtIndex
          ]
        ),
      castId:
        castId,
      castName:
        requestedCastName,
      lineDisplayName:
        String(
          row[
            lineDisplayNameIndex
          ] || ""
        ).trim(),
      status:
        status,
      registrationType:
        castExists
          ? "existing"
          : "new",
      castExists:
        castExists,
      existingCastName:
        existingCast
          ? existingCast.castName
          : "",
      nameMatches:
        nameMatches
    });
  }
  /* =================================================
   * 完了
   * ================================================= */
  return {
    success:
      true,
    requests:
      requests
  };
}
/* =====================================================
 * 12. 登録申請承認
 * ===================================================== */
/**
 * 登録申請を承認する。
 *
 * 既存CAST：
 * キャスト番号で既存CASTを特定し、
 * LINE情報を紐付ける。
 *
 * 新規CAST：
 * CASTに存在しない番号なら
 * 新しいCASTとして登録する。
 *
 * 承認完了後、
 * 申請者本人へLINE通知する。
 *
 * @param {string} idToken
 * @param {string} requestId
 * @return {Object}
 */
function approveRegistrationRequest_(
  idToken,
  requestId
) {
  const admin =
    requireAdmin_(
      idToken
    );
  const normalizedRequestId =
    String(
      requestId || ""
    ).trim();
  if (!normalizedRequestId) {
    throw new Error(
      "申請IDがありません"
    );
  }
  const lock =
    LockService.getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    /* =================================================
     * REGISTRATION_REQUESTS
     * ================================================= */
    const requestSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.REGISTRATION_REQUESTS
      );
    const requestLastColumn =
      requestSheet.getLastColumn();
    const requestHeaders =
      requestSheet
        .getRange(
          1,
          1,
          1,
          requestLastColumn
        )
        .getDisplayValues()[0]
        .map(
          value =>
            String(
              value || ""
            ).trim()
        );
    const requestIdIndex =
      requestHeaders.indexOf(
        "申請ID"
      );
    const requestedNameIndex =
      requestHeaders.indexOf(
        "入力名"
      );
    const lineDisplayNameIndex =
      requestHeaders.indexOf(
        "LINE表示名"
      );
    const lineUserIdIndex =
      requestHeaders.indexOf(
        "LINEユーザーID"
      );
    const statusIndex =
      requestHeaders.indexOf(
        "状態"
      );
    const processedAtIndex =
      requestHeaders.indexOf(
        "処理日時"
      );
    const processedByIndex =
      requestHeaders.indexOf(
        "処理者"
      );
    const requestCastIdIndex =
      requestHeaders.indexOf(
        "CAST_ID"
      );
    const requestNoteIndex =
      requestHeaders.indexOf(
        "備考"
      );
    const missingRequestHeaders =
      [];
    [
      ["申請ID", requestIdIndex],
      ["入力名", requestedNameIndex],
      ["LINE表示名", lineDisplayNameIndex],
      ["LINEユーザーID", lineUserIdIndex],
      ["状態", statusIndex],
      ["処理日時", processedAtIndex],
      ["処理者", processedByIndex],
      ["CAST_ID", requestCastIdIndex],
      ["備考", requestNoteIndex]
    ].forEach(
      ([name, index]) => {
        if (index < 0) {
          missingRequestHeaders.push(
            name
          );
        }
      }
    );
    if (
      missingRequestHeaders.length
    ) {
      throw new Error(
        "REGISTRATION_REQUESTSに必要な列がありません：" +
        missingRequestHeaders.join(
          " / "
        )
      );
    }
    const requestLastRow =
      requestSheet.getLastRow();
    if (
      requestLastRow < 2
    ) {
      throw new Error(
        "登録申請が見つかりません"
      );
    }
    const requestValues =
      requestSheet
        .getRange(
          2,
          1,
          requestLastRow - 1,
          requestLastColumn
        )
        .getValues();
    let requestRowNumber =
      -1;
    let requestRow =
      null;
    for (
      let index = 0;
      index < requestValues.length;
      index++
    ) {
      const row =
        requestValues[index];
      if (
        String(
          row[
            requestIdIndex
          ] || ""
        ).trim() ===
        normalizedRequestId
      ) {
        requestRowNumber =
          index + 2;
        requestRow =
          row;
        break;
      }
    }
    if (
      !requestRow ||
      requestRowNumber < 2
    ) {
      throw new Error(
        "登録申請が見つかりません"
      );
    }
    const requestStatus =
      String(
        requestRow[
          statusIndex
        ] || ""
      ).trim();
    if (
      requestStatus !==
      "承認待ち"
    ) {
      throw new Error(
        "この申請はすでに処理されています"
      );
    }
    const castId =
      String(
        requestRow[
          requestCastIdIndex
        ] || ""
      ).trim();
    const castName =
      String(
        requestRow[
          requestedNameIndex
        ] || ""
      ).trim();
    const lineDisplayName =
      String(
        requestRow[
          lineDisplayNameIndex
        ] || ""
      ).trim();
    const lineUserId =
      String(
        requestRow[
          lineUserIdIndex
        ] || ""
      ).trim();
    if (!castId) {
      throw new Error(
        "申請のキャスト番号がありません"
      );
    }
    if (!castName) {
      throw new Error(
        "申請のキャスト名がありません"
      );
    }
    if (!lineUserId) {
      throw new Error(
        "申請のLINEユーザーIDがありません"
      );
    }
    /*
     * 同じLINEが
     * 別CASTに登録されていないか
     */
    const existingLineCast =
      getCastByLineUserId_(
        lineUserId
      );
    if (
      existingLineCast &&
      String(
        existingLineCast.castId || ""
      ).trim() !==
      castId
    ) {
      throw new Error(
        "このLINEアカウントはすでに別のキャストに登録されています"
      );
    }
    /* =================================================
     * CAST
     * ================================================= */
    const castSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.CAST
      );
    const castLastColumn =
      castSheet.getLastColumn();
    const castHeaders =
      castSheet
        .getRange(
          1,
          1,
          1,
          castLastColumn
        )
        .getDisplayValues()[0]
        .map(
          value =>
            String(
              value || ""
            ).trim()
        );
    const castIdIndex =
      castHeaders.indexOf(
        "キャストID"
      );
    const castNameIndex =
      castHeaders.indexOf(
        "キャスト名"
      );
    const defaultStartIndex =
      castHeaders.indexOf(
        "DEFAULT_START_TIME"
      );
    const defaultEndIndex =
      castHeaders.indexOf(
        "DEFAULT_END_TIME"
      );
    const castLineDisplayNameIndex =
      castHeaders.indexOf(
        "LINE表示名"
      );
    const castLineUserIdIndex =
      castHeaders.indexOf(
        "LINEユーザーID"
      );
    const activeStatusIndex =
      castHeaders.indexOf(
        "在籍状況"
      );
    const registeredAtIndex =
      castHeaders.indexOf(
        "登録日"
      );
    const adminIndex =
      castHeaders.indexOf(
        "管理者権限"
      );
    const castNoteIndex =
      castHeaders.indexOf(
        "備考"
      );
    const missingCastHeaders =
      [];
    [
      ["キャストID", castIdIndex],
      ["キャスト名", castNameIndex],
      ["LINE表示名", castLineDisplayNameIndex],
      ["LINEユーザーID", castLineUserIdIndex],
      ["在籍状況", activeStatusIndex],
      ["登録日", registeredAtIndex]
    ].forEach(
      ([name, index]) => {
        if (index < 0) {
          missingCastHeaders.push(
            name
          );
        }
      }
    );
    if (
      missingCastHeaders.length
    ) {
      throw new Error(
        "CASTシートに必要な列がありません：" +
        missingCastHeaders.join(
          " / "
        )
      );
    }
    const castLastRow =
      castSheet.getLastRow();
    let castRowNumber =
      -1;
    let existingCastName =
      "";
    let existingCastLineUserId =
      "";
    if (
      castLastRow >= 2
    ) {
      const castValues =
        castSheet
          .getRange(
            2,
            1,
            castLastRow - 1,
            castLastColumn
          )
          .getValues();
      for (
        let index = 0;
        index < castValues.length;
        index++
      ) {
        const row =
          castValues[index];
        const rowCastId =
          String(
            row[
              castIdIndex
            ] || ""
          ).trim();
        if (
          rowCastId !==
          castId
        ) {
          continue;
        }
        castRowNumber =
          index + 2;
        existingCastName =
          String(
            row[
              castNameIndex
            ] || ""
          ).trim();
        existingCastLineUserId =
          String(
            row[
              castLineUserIdIndex
            ] || ""
          ).trim();
        break;
      }
    }
    let registrationType =
      "";
    /* =================================================
     * 既存CAST
     * ================================================= */
    if (
      castRowNumber >= 2
    ) {
      registrationType =
        "existing";
      if (
        normalizeCastName_(
          existingCastName
        ) !==
        normalizeCastName_(
          castName
        )
      ) {
        throw new Error(
          "キャスト番号とキャスト名が一致しません"
        );
      }
      if (
        existingCastLineUserId &&
        existingCastLineUserId !==
        lineUserId
      ) {
        throw new Error(
          "このキャスト番号にはすでに別のLINEアカウントが登録されています"
        );
      }
      castSheet
        .getRange(
          castRowNumber,
          castLineDisplayNameIndex + 1
        )
        .setValue(
          lineDisplayName
        );
      castSheet
        .getRange(
          castRowNumber,
          castLineUserIdIndex + 1
        )
        .setValue(
          lineUserId
        );
      castSheet
        .getRange(
          castRowNumber,
          activeStatusIndex + 1
        )
        .setValue(
          "在籍"
        );
      castSheet
        .getRange(
          castRowNumber,
          registeredAtIndex + 1
        )
        .setValue(
          new Date()
        );
    }
    /* =================================================
     * 新規CAST
     * ================================================= */
    else {
      registrationType =
        "new";
      const newCastRow =
        new Array(
          castLastColumn
        ).fill(
          ""
        );
      newCastRow[
        castIdIndex
      ] =
        castId;
      newCastRow[
        castNameIndex
      ] =
        castName;
      if (
        defaultStartIndex >= 0
      ) {
        newCastRow[
          defaultStartIndex
        ] = "";
      }
      if (
        defaultEndIndex >= 0
      ) {
        newCastRow[
          defaultEndIndex
        ] = "";
      }
      newCastRow[
        castLineDisplayNameIndex
      ] =
        lineDisplayName;
      newCastRow[
        castLineUserIdIndex
      ] =
        lineUserId;
      newCastRow[
        activeStatusIndex
      ] =
        "在籍";
      newCastRow[
        registeredAtIndex
      ] =
        new Date();
      if (

        adminIndex >= 0

      ) {

        newCastRow[

          adminIndex

        ] =

          false;

      }
      if (
        castNoteIndex >= 0
      ) {
        newCastRow[
          castNoteIndex
        ] = "";
      }
      castSheet.appendRow(
        newCastRow
      );
    }
    /* =================================================
     * 申請を承認済みに更新
     * ================================================= */
    requestSheet
      .getRange(
        requestRowNumber,
        statusIndex + 1
      )
      .setValue(
        "承認済み"
      );
    requestSheet
      .getRange(
        requestRowNumber,
        processedAtIndex + 1
      )
      .setValue(
        new Date()
      );
    requestSheet
      .getRange(
        requestRowNumber,
        processedByIndex + 1
      )
      .setValue(
        admin.admin.adminName
      );
    requestSheet
      .getRange(
        requestRowNumber,
        requestCastIdIndex + 1
      )
      .setValue(
        castId
      );
    requestSheet
      .getRange(
        requestRowNumber,
        requestNoteIndex + 1
      )
      .setValue(
        registrationType === "new"
          ? "新規CASTとして登録"
          : "既存CASTへ紐付け"
      );
    /* =================================================
     * 本人へLINE通知
     * ================================================= */
    const approvedCastName =
      registrationType === "existing"
        ? existingCastName
        : castName;
    const approvalMessage = [
      "CAST FLOW｜登録完了",
      "",
      approvedCastName + " さん",
      "",
      "登録申請が承認されました。",
      "",
      "キャスト番号：" +
        castId,
      "",
      "CAST FLOWをご利用いただけます。"
    ].join(
      "\n"
    );
    /*
     * LINE通知失敗で
     * 承認処理自体は取り消さない
     */
    try {
      pushLineMessage_(
        lineUserId,
        approvalMessage
      );
    } catch (
      notificationError
    ) {
      Logger.log(
        "登録承認LINE通知に失敗：" +
        (
          notificationError?.message ||
          String(
            notificationError
          )
        )
      );
    }
    return {
      success:
        true,
      approved:
        true,
      requestId:
        normalizedRequestId,
      registrationType:
        registrationType,
      createdNewCast:
        registrationType === "new",
      cast: {
        castId:
          castId,
        castName:
          approvedCastName,
        lineDisplayName:
          lineDisplayName
      }
    };
  } finally {
    lock.releaseLock();
  }
}
/* =====================================================
 * 13. 登録申請却下
 * ===================================================== */
/**
 * 登録申請を却下する。
 *
 * 管理者専用。
 * 却下後は申請者本人へLINE通知する。
 *
 * @param {string} idToken
 * @param {string} requestId
 * @param {string} note
 * @return {Object}
 */
function rejectRegistrationRequest_(
  idToken,
  requestId,
  note
) {
  const admin =
    requireAdmin_(
      idToken
    );
  const normalizedRequestId =
    String(
      requestId || ""
    ).trim();
  const normalizedNote =
    String(
      note || ""
    ).trim();
  if (!normalizedRequestId) {
    throw new Error(
      "申請IDがありません"
    );
  }
  const lock =
    LockService.getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    const sheet =
      getRequiredSheet_(
        CONFIG.SHEETS.REGISTRATION_REQUESTS
      );
    const lastColumn =
      sheet.getLastColumn();
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
    const requestIdIndex =
      headers.indexOf(
        "申請ID"
      );
    const castNameIndex =
      headers.indexOf(
        "入力名"
      );
    const lineUserIdIndex =
      headers.indexOf(
        "LINEユーザーID"
      );
    const statusIndex =
      headers.indexOf(
        "状態"
      );
    const processedAtIndex =
      headers.indexOf(
        "処理日時"
      );
    const processedByIndex =
      headers.indexOf(
        "処理者"
      );
    const castIdIndex =
      headers.indexOf(
        "CAST_ID"
      );
    const noteIndex =
      headers.indexOf(
        "備考"
      );
    const missingHeaders =
      [];
    [
      ["申請ID", requestIdIndex],
      ["入力名", castNameIndex],
      ["LINEユーザーID", lineUserIdIndex],
      ["状態", statusIndex],
      ["処理日時", processedAtIndex],
      ["処理者", processedByIndex],
      ["CAST_ID", castIdIndex],
      ["備考", noteIndex]
    ].forEach(
      ([name, index]) => {
        if (index < 0) {
          missingHeaders.push(
            name
          );
        }
      }
    );
    if (
      missingHeaders.length
    ) {
      throw new Error(
        "REGISTRATION_REQUESTSに必要な列がありません：" +
        missingHeaders.join(
          " / "
        )
      );
    }
    const lastRow =
      sheet.getLastRow();
    if (
      lastRow < 2
    ) {
      throw new Error(
        "登録申請が見つかりません"
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
    let rowNumber =
      -1;
    let requestRow =
      null;
    for (
      let index = 0;
      index < values.length;
      index++
    ) {
      const row =
        values[index];
      if (
        String(
          row[
            requestIdIndex
          ] || ""
        ).trim() ===
        normalizedRequestId
      ) {
        rowNumber =
          index + 2;
        requestRow =
          row;
        break;
      }
    }
    if (
      !requestRow ||
      rowNumber < 2
    ) {
      throw new Error(
        "登録申請が見つかりません"
      );
    }
    const status =
      String(
        requestRow[
          statusIndex
        ] || ""
      ).trim();
    if (
      status !==
      "承認待ち"
    ) {
      throw new Error(
        "この申請はすでに処理されています"
      );
    }
    const castName =
      String(
        requestRow[
          castNameIndex
        ] || ""
      ).trim();
    const castId =
      String(
        requestRow[
          castIdIndex
        ] || ""
      ).trim();
    const lineUserId =
      String(
        requestRow[
          lineUserIdIndex
        ] || ""
      ).trim();
    /*
     * 状態
     */
    sheet
      .getRange(
        rowNumber,
        statusIndex + 1
      )
      .setValue(
        "却下"
      );
    /*
     * 処理日時
     */
    sheet
      .getRange(
        rowNumber,
        processedAtIndex + 1
      )
      .setValue(
        new Date()
      );
    /*
     * 処理者
     */
    sheet
      .getRange(
        rowNumber,
        processedByIndex + 1
      )
      .setValue(
        admin.admin.adminName
      );
    /*
     * 却下理由
     */
    sheet
      .getRange(
        rowNumber,
        noteIndex + 1
      )
      .setValue(
        normalizedNote
      );
    /* =================================================
     * 本人へLINE通知
     * ================================================= */
    if (
      lineUserId
    ) {
      const rejectionMessage = [
        "CAST FLOW｜登録申請",
        "",
        castName
          ? castName + " さん"
          : "",
        "",
        "登録申請が却下されました。",
        castId
          ? "キャスト番号：" + castId
          : "",
        normalizedNote
          ? ""
          : null,
        normalizedNote
          ? "理由：" + normalizedNote
          : null,
        "",
        "内容をご確認のうえ、必要に応じて再度申請してください。"
      ]
        .filter(
          value =>
            value !== null
        )
        .join(
          "\n"
        );
      try {
        pushLineMessage_(
          lineUserId,
          rejectionMessage
        );
      } catch (
        notificationError
      ) {
        Logger.log(
          "登録却下LINE通知に失敗：" +
          (
            notificationError?.message ||
            String(
              notificationError
            )
          )
        );
      }
    }
    return {
      success:
        true,
      rejected:
        true,
      requestId:
        normalizedRequestId,
      castId:
        castId,
      castName:
        castName,
      note:
        normalizedNote
    };
  } finally {
    lock.releaseLock();
  }
}
/* =====================================================
 * 14. 登録申請日時表示
 * ===================================================== */
function formatRegistrationDate_(
  value
) {
  if (!value) {
    return "";
  }
  try {
    return Utilities.formatDate(
      new Date(value),
      CONFIG.TIMEZONE ||
      "Asia/Tokyo",
      "yyyy/MM/dd HH:mm"
    );
  } catch (error) {
    return String(
      value
    );
  }
}
/* =====================================================
 * 99. 開発用：ダミー登録申請作成
 * ===================================================== */
/**
 * 開発テスト専用。
 *
 * CASTに S999 / テストキャスト が存在する前提で、
 * REGISTRATION_REQUESTS に承認待ち申請を1件作成する。
 */
function createTestRegistrationRequest() {
  const cast =
    getCastByName_(
      "テストキャスト"
    );
  if (!cast) {
    throw new Error(
      "CASTに「テストキャスト」が登録されていません"
    );
  }
  if (
    cast.lineUserId
  ) {
    throw new Error(
      "テストキャストにはすでにLINEユーザーIDが登録されています"
    );
  }
  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.REGISTRATION_REQUESTS
    );
  const testLineUserId =
    "TEST_LINE_USER_001";
  /*
   * 同じテスト申請がすでに承認待ちなら
   * 二重作成しない
   */
  const lastRow =
    sheet.getLastRow();
  if (
    lastRow >= 2
  ) {
    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          10
        )
        .getValues();
    for (
      let index = 0;
      index < values.length;
      index++
    ) {
      const row =
        values[index];
      const lineUserId =
        String(
          row[4] || ""
        ).trim();
      const status =
        String(
          row[5] || ""
        ).trim();
      if (
        lineUserId === testLineUserId &&
        status === "承認待ち"
      ) {
        Logger.log(
          "すでに承認待ちのテスト申請があります"
        );
        return;
      }
    }
  }
  const requestId =
    "TEST_" +
    Utilities.getUuid();
  sheet.appendRow([
    requestId,
    new Date(),
    cast.castName,
    "ROYAL TEST",
    testLineUserId,
    "承認待ち",
    "",
    "",
    cast.castId,
    "開発テスト用"
  ]);
  Logger.log(
    "テスト登録申請を作成しました：" +
    requestId
  );
}
