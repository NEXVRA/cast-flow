/* =====================================================

 * ROYAL SHIFT SYSTEM

 * Line.gs

 *

 * LINE Messaging API

 * ===================================================== */





/* =====================================================

 * 01. LINE設定取得

 * ===================================================== */



function getLineChannelAccessToken_() {



  const token =

    PropertiesService

      .getScriptProperties()

      .getProperty(

        "LINE_CHANNEL_ACCESS_TOKEN"

      );



  if (!token) {

    throw new Error(

      "LINE_CHANNEL_ACCESS_TOKEN が設定されていません。"

    );

  }



  return token;

}





function getLineChannelSecret_() {



  const secret =

    PropertiesService

      .getScriptProperties()

      .getProperty(

        "LINE_CHANNEL_SECRET"

      );



  if (!secret) {

    throw new Error(

      "LINE_CHANNEL_SECRET が設定されていません。"

    );

  }



  return secret;

}





/* =====================================================

 * 01-B. 実行環境・LINE送信安全制御

 * ===================================================== */



/**
 * 現在のGAS環境を返す。
 *
 * CONFIG.ENVIRONMENT が未設定の場合は
 * 既存本番との互換性を優先して production 扱いにする。
 */
function getGasEnvironment_() {

  const environment =
    String(
      CONFIG &&
      CONFIG.ENVIRONMENT ||
      "production"
    )
      .trim()
      .toLowerCase();


  return environment ||
    "production";

}



/**
 * LINE Messaging APIへの実送信が許可されているか。
 *
 * production:
 *   従来どおり送信する。
 *
 * development:
 *   CONFIG.SAFETY.LINE_SEND_ENABLED === true
 *   の場合だけ送信する。
 */
function isLineSendEnabled_() {

  if (
    getGasEnvironment_() !==
    "development"
  ) {

    return true;

  }


  return Boolean(
    CONFIG &&
    CONFIG.SAFETY &&
    CONFIG.SAFETY.LINE_SEND_ENABLED === true
  );

}



/**
 * LINE API送信の共通入口。
 *
 * DEVでLINE送信が無効な場合は
 * 実際のUrlFetchApp.fetchを実行せず、
 * 成功相当のモックレスポンスを返す。
 *
 * これにより、DEVデータ上の保存・画面確認などは
 * 本番に近い流れのままテストできる。
 */
function fetchLineApi_(
  url,
  options
) {

  if (
    !isLineSendEnabled_()
  ) {

    Logger.log(
      "[CAST FLOW DEV] LINE送信をスキップしました: " +
      String(
        url || ""
      )
    );


    return {

      getResponseCode:
        function() {
          return 200;
        },

      getContentText:
        function() {
          return JSON.stringify({
            success: true,
            skipped: true,
            environment:
              getGasEnvironment_()
          });
        }

    };

  }


  return UrlFetchApp.fetch(
    url,
    options
  );

}



/**
 * 環境に応じたCAST FLOWのLIFF IDを返す。
 */
function getCastFlowLiffId_() {

  if (
    getGasEnvironment_() ===
    "development"
  ) {

    return "2011488723-CUmNFMxv";

  }


  return "2011488723-gI54Hlt3";

}



/**
 * 環境に応じたCAST FLOWのLIFF URLを返す。
 */
function getCastFlowLiffUrl_() {

  return (
    "https://liff.line.me/" +
    getCastFlowLiffId_()
  );

}



/* =====================================================

 * 02. 管理者LINE USER ID

 * ===================================================== */



/**

 * 管理者LINE USER IDを取得

 */

function getAdminLineUserId_() {



  return (

    PropertiesService

      .getScriptProperties()

      .getProperty(

        "ADMIN_LINE_USER_ID"

      ) || ""

  );



}





/**

 * 管理者LINE USER IDを保存

 */

function saveAdminLineUserId_(

  userId

) {



  if (!userId) {

    throw new Error(

      "LINE USER ID がありません。"

    );

  }



  PropertiesService

    .getScriptProperties()

    .setProperty(

      "ADMIN_LINE_USER_ID",

      userId

    );



}





/* =====================================================

 * 03. Webhookイベント処理

 * ===================================================== */



/**

 * LINEから受け取ったイベントを処理

 */

function handleLineWebhook_(

  body

) {



  if (

    !body ||

    !Array.isArray(body.events)

  ) {

    return;

  }





  body.events.forEach(

    event => {



      handleLineEvent_(

        event

      );



    }

  );



}





/**

 * LINEイベント1件を処理

 */

function handleLineEvent_(

  event

) {



  if (!event) {

    return;

  }





  const source =

    event.source || {};





  const userId =

    source.userId || "";





  /*

   * テキストメッセージのみ処理

   */

  if (

    event.type === "message" &&

    event.message &&

    event.message.type === "text"

  ) {



    const text =

      String(

        event.message.text || ""

      ).trim();





    handleLineTextMessage_(

      userId,

      event.replyToken,

      text

    );



  }



}





/* =====================================================

 * 04. テキストメッセージ処理

 * ===================================================== */



function handleLineTextMessage_(

  userId,

  replyToken,

  text

) {



  /*

   * 管理者登録

   */

  if (

    text === "管理者登録"

  ) {



    saveAdminLineUserId_(

      userId

    );





    replyLineMessage_(

      replyToken,

      "ROYAL SHIFTの管理者LINEとして登録しました。"

    );





    return;

  }





  /*

   * 接続確認

   */

  if (

    text === "接続確認"

  ) {



    replyLineMessage_(

      replyToken,

      "ROYAL SHIFTとLINEの接続は正常です。"

    );





    return;

  }



}





/* =====================================================

 * 05. Reply Message

 * ===================================================== */



/**

 * LINEのメッセージに返信

 */

function replyLineMessage_(

  replyToken,

  message

) {



  if (

    !replyToken ||

    !message

  ) {

    return;

  }





  const url =

    "https://api.line.me/v2/bot/message/reply";





  const payload = {



    replyToken:

      replyToken,



    messages: [

      {

        type:

          "text",



        text:

          String(

            message

          )

      }

    ]



  };





  fetchLineApi_(

    url,

    {



      method:

        "post",



      contentType:

        "application/json",



      headers: {



        Authorization:

          "Bearer " +

          getLineChannelAccessToken_()



      },



      payload:

        JSON.stringify(

          payload

        ),



      muteHttpExceptions:

        true



    }

  );



}





/* =====================================================

 * 06. Push Message

 * ===================================================== */



/**

 * 指定ユーザーへLINEを送信

 */

function pushLineMessage_(

  userId,

  message

) {



  if (!userId) {



    throw new Error(

      "送信先LINE USER IDがありません。"

    );



  }





  const url =

    "https://api.line.me/v2/bot/message/push";





  const payload = {



    to:

      userId,





    /*

     * LINE側のプッシュ通知を

     * 無効化しないことを明示

     */

    notificationDisabled:

      false,





    messages: [



      {



        type:

          "text",



        text:

          String(

            message

          )



      }



    ]



  };





  const response =

    fetchLineApi_(

      url,

      {



        method:

          "post",



        contentType:

          "application/json",



        headers: {



          Authorization:

            "Bearer " +

            getLineChannelAccessToken_()



        },



        payload:

          JSON.stringify(

            payload

          ),



        muteHttpExceptions:

          true



      }

    );





  const statusCode =

    response.getResponseCode();





  if (

    statusCode < 200 ||

    statusCode >= 300

  ) {



    throw new Error(

      "LINE送信エラー: " +

      response.getContentText()

    );



  }





  return true;



}



/* =====================================================

 * 07. 管理者通知

 * ===================================================== */



/**

 * CASTシートから

 * LINE通知対象の管理者USER IDを取得する

 *

 * 条件：

 * ・管理者権限 = TRUE

 * ・LINEユーザーIDあり

 * ・在籍状況が空欄 または 在籍

 *

 * @return {string[]}

 */

function getAdminLineUserIds_() {



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





  const lineUserIdIndex =

    headers.indexOf(

      "LINEユーザーID"

    );





  const adminIndex =

    headers.indexOf(

      "管理者権限"

    );





  const activeStatusIndex =

    headers.indexOf(

      "在籍状況"

    );





  if (

    lineUserIdIndex < 0

  ) {



    throw new Error(

      "CASTシートにLINEユーザーID列がありません。"

    );



  }





  if (

    adminIndex < 0

  ) {



    throw new Error(

      "CASTシートに管理者権限列がありません。"

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





  const userIds =

    [];





  values.forEach(

    row => {



      const lineUserId =

        String(

          row[

            lineUserIdIndex

          ] || ""

        ).trim();





      if (!lineUserId) {



        return;



      }





      const adminValue =

        row[

          adminIndex

        ];





      const isAdmin =

        adminValue === true ||

        String(

          adminValue || ""

        )

          .trim()

          .toUpperCase() === "TRUE" ||

        String(

          adminValue || ""

        ).trim() === "1";





      if (!isAdmin) {



        return;



      }





      /*

       * 在籍状況が存在する場合

       * 在籍のみ通知対象。

       *

       * 空欄は移行期間を考慮して許可。

       */

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





      userIds.push(

        lineUserId

      );



    }

  );





  return [

    ...new Set(

      userIds

    )

  ];



}





/**

 * 管理者通知対象USER IDを取得

 *

 * CASTシートに管理者がいない場合のみ

 * 旧ADMIN_LINE_USER_IDをフォールバック使用。

 *

 * @return {string[]}

 */

function getAdminNotificationUserIds_() {



  let adminUserIds =

    getAdminLineUserIds_();





  if (

    adminUserIds.length === 0

  ) {



    const legacyAdminUserId =

      getAdminLineUserId_();





    if (

      legacyAdminUserId

    ) {



      adminUserIds = [

        legacyAdminUserId

      ];



    }



  }





  return [

    ...new Set(

      adminUserIds

    )

  ];



}





/**

 * 管理者全員へ通常テキストLINE通知

 */

function sendAdminLineMessage_(

  message

) {



  const adminUserIds =

    getAdminNotificationUserIds_();





  if (

    adminUserIds.length === 0

  ) {



    throw new Error(

      "LINE通知できる管理者が登録されていません。"

    );



  }





  adminUserIds.forEach(

    userId => {



      pushLineMessage_(

        userId,

        message

      );



    }

  );





  return true;



}





/**

 * 登録申請専用

 * 管理者LINE通知

 *

 * @param {Object} request

 * @return {boolean}

 */

function sendAdminRegistrationRequestNotification_(

  request

) {



  const adminUserIds =

    getAdminNotificationUserIds_();





  if (

    adminUserIds.length === 0

  ) {



    throw new Error(

      "LINE通知できる管理者が登録されていません。"

    );



  }





  adminUserIds.forEach(

    userId => {



      pushAdminRegistrationRequestMessage_(

        userId,

        request

      );



    }

  );





  return true;



}





/**

 * 登録申請専用メッセージを

 * 管理者1名へ送信

 */

function pushAdminRegistrationRequestMessage_(

  userId,

  request

) {



  if (!userId) {



    throw new Error(

      "送信先LINE USER IDがありません。"

    );



  }





  const requestId =

    String(

      request &&

      request.requestId ||

      ""

    ).trim();





  const castId =

    String(

      request &&

      request.castId ||

      ""

    ).trim();





  const castName =

    String(

      request &&

      request.castName ||

      ""

    ).trim();





  const lineDisplayName =

    String(

      request &&

      request.lineDisplayName ||

      ""

    ).trim();





  const registrationType =

    String(

      request &&

      request.registrationType ||

      ""

    ).trim();





  const typeLabel =

    (

      registrationType === "existing" ||

      registrationType === "既存CAST"

    )

      ? "既存CAST"

      : "新規CAST";





  /*

   * 将来的にこのURLを受け取ったら

   * 管理者登録申請画面へ直接遷移させる。

   */

  let reviewUrl =

    getCastFlowLiffUrl_() +

    "?action=admin-registration";





  if (

    requestId

  ) {



    reviewUrl +=

      "&requestId=" +

      encodeURIComponent(

        requestId

      );



  }





  /*

   * LINEテンプレートの本文が

   * 長くなりすぎないよう制限

   */

  const templateText =

    [



      "キャスト番号：" +

        (

          castId ||

          "-"

        ),



      "キャスト名：" +

        (

          castName ||

          "-"

        ),



      "LINE表示名：" +

        (

          lineDisplayName ||

          "-"

        ),



      "登録区分：" +

        typeLabel



    ]

      .join(

        "\n"

      )

      .slice(

        0,

        150

      );





  const payload = {



    to:

      userId,



    notificationDisabled:

      false,



    messages: [



      {



        type:

          "template",



        altText:

          "CAST FLOW｜新しい登録申請があります",



        template: {



          type:

            "buttons",



          title:

            "CAST FLOW｜登録申請",



          text:

            templateText,



          actions: [



            {



              type:

                "uri",



              label:

                "登録申請を確認",



              uri:

                reviewUrl



            }



          ]



        }



      }



    ]



  };





  const response =

    fetchLineApi_(



      "https://api.line.me/v2/bot/message/push",



      {



        method:

          "post",



        contentType:

          "application/json",



        headers: {



          Authorization:

            "Bearer " +

            getLineChannelAccessToken_()



        },



        payload:

          JSON.stringify(

            payload

          ),



        muteHttpExceptions:

          true



      }



    );





  const statusCode =

    response.getResponseCode();





  if (

    statusCode < 200 ||

    statusCode >= 300

  ) {



    throw new Error(

      "登録申請LINE通知エラー: " +

      response.getContentText()

    );



  }





  return true;



}





/**

 * 登録申請通知の手動テスト

 *

 * GASエディタから

 * この関数を手動実行する。

 */

function testAdminRegistrationRequestNotification() {



  return sendAdminRegistrationRequestNotification_({



    requestId:

      "TEST-" +

      Utilities.getUuid(),



    castId:

      "S002",



    castName:

      "タブレット",



    lineDisplayName:

      "CAST FLOW TEST",



    registrationType:

      "new"



  });



}





/**

 * 月間シフト提出を管理者へ通知

 */

function sendMonthlyShiftSubmittedAdminNotification_(

  user,

  targetMonth,

  entries,

  submittedAt,

  isResubmission

) {



  const counts = {



    working:

      0,



    holiday:

      0,



    undecided:

      0



  };





  entries.forEach(

    item => {



      const row =

        item.values || [];





      const status =

        convertStatusToClient_(

          row[6]

        );





      if (

        Object.prototype.hasOwnProperty.call(

          counts,

          status

        )

      ) {



        counts[

          status

        ]++;



      }



    }

  );





  const monthParts =

    String(

      targetMonth || ""

    )

      .split(

        "-"

      );





  const displayMonth =

    monthParts.length === 2



      ? (

          monthParts[0] +

          "年" +

          Number(

            monthParts[1]

          ) +

          "月"

        )



      : targetMonth;





  const submittedTime =

    Utilities.formatDate(



      submittedAt,



      CONFIG.TIMEZONE,



      "MM/dd HH:mm"



    );





  const title =

    isResubmission



      ? "シフトが再提出されました"



      : "シフトが提出されました";





  const message = [



    "ROYAL SHIFT",



    "",



    title,



    "",



    user.castName +

      "（" +

      user.castId +

      "）",



    "対象月：" +

      displayMonth,



    "",



    "出勤：" +

      counts.working +

      "日",



    "休み：" +

      counts.holiday +

      "日",



    "未定：" +

      counts.undecided +

      "日",



    "",



    "入力済み：" +

      entries.length +

      "日",



    "提出日時：" +

      submittedTime



  ].join(

    "\n"

  );





  return sendAdminLineMessage_(

    message

  );



}



/* =====================================================

 * 08. 接続テスト

 * ===================================================== */



/**

 * 管理者LINEへの接続テスト

 */

function testAdminLinePush() {



  sendAdminLineMessage_(



    [

      "CAST FLOW",

      "",

      "LINE接続テスト成功",

      "",

      "管理者通知を送信できる状態になりました。"

    ].join(

      "\n"

    )



  );



}





/**

 * 出勤確認postbackの実機テスト

 *

 * 管理者本人へ

 * 「出勤OK / 変更あり」

 * の確認ボタンを送信する。

 *

 * ATTENDANCE_CONFIRMATIONSに

 * 今日の確認行が存在することを前提とする。

 */

function testAttendanceConfirmationPush() {



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

      "M月d日"

    );





  /*

   * 管理者LINE USER IDを取得

   */

  const adminUserIds =

    getAdminLineUserIds_();





  if (

    !adminUserIds ||

    adminUserIds.length === 0

  ) {



    throw new Error(

      "管理者のLINE USER IDが見つかりません。"

    );



  }





  /*

   * 今回は最初の管理者だけに送る

   */

  const userId =

    adminUserIds[0];





  const message = [



    "CAST FLOW",



    "",



    "【出勤確認テスト】",



    "",



    displayDate,



    "",



    "Webhook / postbackの動作確認です。",



    "",



    "下の「出勤OK」を押してください。"



  ].join(

    "\n"

  );





  /*

   * セクション10の

   * 本番用ボタン送信関数をそのまま使う

   */

  pushAttendanceConfirmationMessage_(

    userId,

    message,

    dateString

  );





  return {



    success:

      true,



    date:

      dateString,



    message:

      "出勤確認テストLINEを送信しました。"



  };



}





/* =====================================================

 * 09. 当日出勤確認LINE

 * ===================================================== */



/**

 * 指定日の出勤予定キャストへ

 * 出勤確認LINEを送信する。

 *

 * LINE送信成功後に

 * ATTENDANCE_CONFIRMATIONSへ確認行を作成する。

 *

 * @param {Date} targetDate

 * @return {Object}

 */

function sendDailyAttendanceConfirmations_(

  targetDate

) {



  const currentSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.SHIFT_CURRENT

    );





  const castSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.CAST

    );





  const confirmationSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.ATTENDANCE_CONFIRMATIONS

    );





  const dateString =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "yyyy-MM-dd"

    );





  const displayDate =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "M月d日"

    );





  const weekday =

    getJapaneseWeekday_(

      targetDate

    );





  /*

   * =========================================

   * SHIFT_CURRENTから当日の出勤予定者を取得

   * =========================================

   */



  const lastRow =

    currentSheet.getLastRow();





  if (

    lastRow < 2

  ) {



    return {

      success: true,

      sentCount: 0,

      skippedCount: 0,

      failedCount: 0

    };



  }





  const shiftValues =

    currentSheet

      .getRange(

        2,

        1,

        lastRow - 1,

        13

      )

      .getValues();





  const workingRows =

    shiftValues.filter(

      row => {



        const rowDate =

          formatDateValue_(

            row[2]

          );





        const status =

          String(

            row[6] || ""

          ).trim();





        return (

          rowDate === dateString &&

          status === "出勤"

        );



      }

    );





  if (

    workingRows.length === 0

  ) {



    return {

      success: true,

      sentCount: 0,

      skippedCount: 0,

      failedCount: 0

    };



  }





  /*

   * =========================================

   * CASTシート読み込み

   * =========================================

   */



  const castLastRow =

    castSheet.getLastRow();





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





  const lineUserIdIndex =

    castHeaders.indexOf(

      "LINEユーザーID"

    );





  const activeStatusIndex =

    castHeaders.indexOf(

      "在籍状況"

    );





  if (

    castIdIndex < 0 ||

    lineUserIdIndex < 0

  ) {



    throw new Error(

      "CASTシートの必要列が見つかりません"

    );



  }





  const castValues =

    castLastRow >= 2



      ? castSheet

          .getRange(

            2,

            1,

            castLastRow - 1,

            castLastColumn

          )

          .getValues()



      : [];





  const castMap =

    {};





  castValues.forEach(

    row => {



      const castId =

        String(

          row[castIdIndex] || ""

        ).trim();





      if (!castId) {

        return;

      }





      const lineUserId =

        String(

          row[lineUserIdIndex] || ""

        ).trim();





      const activeStatus =

        activeStatusIndex >= 0



          ? String(

              row[activeStatusIndex] || ""

            ).trim()



          : "";





      castMap[

        castId

      ] = {



        lineUserId:

          lineUserId,



        activeStatus:

          activeStatus



      };



    }

  );





  /*

   * =========================================

   * 各キャストへ送信

   * =========================================

   */



  let sentCount =

    0;





  let skippedCount =

    0;





  let failedCount =

    0;





  workingRows.forEach(

    row => {



      const castId =

        String(

          row[4] || ""

        ).trim();





      const castName =

        String(

          row[5] || ""

        ).trim();





      const startTime =

        formatTimeValue_(

          row[7]

        );





      const endTime =

        formatTimeValue_(

          row[8]

        );





      const cast =

        castMap[

          castId

        ];





      /*

       * LINE未登録・無効キャストはスキップ

       */

      if (

        !cast ||

        !cast.lineUserId ||

        (

          cast.activeStatus &&

          cast.activeStatus !== "在籍"

        )

      ) {



        skippedCount++;



        return;



      }





      /*

       * すでに同日同キャストの

       * 確認データが存在する場合は

       * 初回LINEを二重送信しない。

       */

      const existing =

        findAttendanceConfirmation_(

          confirmationSheet,

          dateString,

          castId

        );





      if (

        existing

      ) {



        skippedCount++;



        return;



      }





      /*

       * =====================================

       * LINE本文

       * =====================================

       */



      const message = [



        "CAST FLOW",



        "",



        "今日の出勤確認です",



        "",



        displayDate +

          "（" +

          weekday +

          "）",



        castName,



        "",



        "出勤 " +

          startTime +

          "〜" +

          endTime,



        "",



        "下のボタンから確認してください"



      ].join(

        "\n"

      );





      /*

       * =====================================

       * LINE送信

       * =====================================

       *

       * LINE送信成功後にだけ

       * 確認行を作成する。

       */



      try {



        pushAttendanceConfirmationMessage_(

          cast.lineUserId,

          message,

          dateString

        );





        const confirmationId =

          createId(

            "ATTENDANCE"

          );





        const now =

          new Date();





        confirmationSheet.appendRow([

          confirmationId,

          targetDate,

          castId,

          castName,

          cast.lineUserId,

          "確認待ち",

          "",

          now,

          ""

        ]);





        sentCount++;





      } catch (

        error

      ) {



        failedCount++;





        writeErrorLog_(

          "sendDailyAttendanceConfirmations_",

          error,

          {



            date:

              dateString,



            castId:

              castId,



            castName:

              castName



          }

        );



      }



    }

  );





  /*

   * =========================================

   * 実行結果

   * =========================================

   */



  return {



    success:

      true,



    date:

      dateString,



    targetCount:

      workingRows.length,



    sentCount:

      sentCount,



    skippedCount:

      skippedCount,



    failedCount:

      failedCount



  };



}





/**

 * 今日の出勤予定キャストへ

 * 出勤確認LINEを送信する。

 *

 * 13時台の時間主導トリガー用。

 */

function sendTodayAttendanceConfirmations() {



  return sendDailyAttendanceConfirmations_(

    new Date()

  );



}





/**

 * ATTENDANCE_CONFIRMATIONSから

 * 指定日・指定キャストの確認データを探す。

 *

 * 09 / 出勤OK / 変更あり などから共通利用。

 *

 * @param {Sheet} sheet

 * @param {string} dateString YYYY-MM-DD

 * @param {string} castId

 * @return {Object|null}

 */

function findAttendanceConfirmation_(

  sheet,

  dateString,

  castId

) {



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

        9

      )

      .getValues();





  const targetCastId =

    String(

      castId || ""

    ).trim();





  for (

    let i = 0;

    i < values.length;

    i++

  ) {



    const row =

      values[i];





    const rowDate =

      formatDateValue_(

        row[1]

      );





    const rowCastId =

      String(

        row[2] || ""

      ).trim();





    if (

      rowDate === dateString &&

      rowCastId === targetCastId

    ) {



      return {



        rowNumber:

          i + 2,



        confirmationId:

          String(

            row[0] || ""

          ).trim(),



        status:

          String(

            row[5] || ""

          ).trim(),



        confirmedAt:

          row[6] || "",



        updatedAt:

          row[7] || "",



        values:

          row



      };



    }



  }





  return null;



}





/* =====================================================

 * 10. 当日確認ボタン付きLINE

 * ===================================================== */



/**

 * 出勤確認用のボタン付きLINEを送信する

 *

 * 出勤OK：

 * LINE postback → Cloudflare Worker → GAS

 *

 * 変更あり：

 * LIFFを開いてシフト内容を編集

 */

function pushAttendanceConfirmationMessage_(

  userId,

  message,

  dateString

) {



  if (!userId) {

    throw new Error(

      "送信先LINE USER IDがありません。"

    );

  }





  /*

   * 日付形式を事前チェック

   */

  parseDateString_(

    dateString

  );





  const url =

    "https://api.line.me/v2/bot/message/push";





  /*

   * =========================================

   * 出勤OK

   *

   * LIFFは開かず、

   * LINE Webhookへpostbackを送信する。

   * =========================================

   */

  const attendanceOkPostbackData =

    "action=attendance-ok" +

    "&date=" +

    encodeURIComponent(

      dateString

    );





  /*

   * =========================================

   * 変更あり

   *

   * こちらはシフト編集が必要なので

   * 従来通りLIFFを開く。

   * =========================================

   */

  const attendanceChangeUrl =

    getCastFlowLiffUrl_() +

    "?action=attendance-change" +

    "&date=" +

    encodeURIComponent(

      dateString

    );





  const payload = {



    to:

      userId,



    messages: [



      {

        type:

          "text",



        text:

          String(

            message

          )

      },



      {

        type:

          "template",



        altText:

          "今日の出勤確認",



        template: {



          type:

            "buttons",



          text:

            "本日の出勤に変更はありませんか？",



          actions: [



            {

              type:

                "postback",



              label:

                "出勤OK",



              data:

                attendanceOkPostbackData,



              displayText:

                "出勤OK"

            },



            {

              type:

                "uri",



              label:

                "変更あり",



              uri:

                attendanceChangeUrl

            }



          ]



        }



      }



    ]



  };





  const response =

    fetchLineApi_(

      url,

      {



        method:

          "post",



        contentType:

          "application/json",



        headers: {



          Authorization:

            "Bearer " +

            getLineChannelAccessToken_()



        },



        payload:

          JSON.stringify(

            payload

          ),



        muteHttpExceptions:

          true



      }

    );





  const statusCode =

    response.getResponseCode();





  if (

    statusCode < 200 ||

    statusCode >= 300

  ) {



    throw new Error(

      "LINE送信エラー: " +

      response.getContentText()

    );



  }





  return true;



}





/* =====================================================

 * 11. 管理者向け当日確認サマリー

 * ===================================================== */



/**

 * 指定日の出勤確認状況を集計して

 * 管理者LINEへ送信する。

 *

 * 表示ルール：

 * ・全員確認済み → 短く表示

 * ・変更あり → 変更者だけ詳細表示

 * ・未確認あり → 未確認者を最優先で詳細表示

 *

 * @param {Date} targetDate

 * @return {Object}

 */

function sendDailyAttendanceAdminSummary_(

  targetDate

) {



  const currentSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.SHIFT_CURRENT

    );





  const confirmationSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.ATTENDANCE_CONFIRMATIONS

    );





  const dateString =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "yyyy-MM-dd"

    );





  const displayDate =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "M月d日"

    );





  const weekday =

    getJapaneseWeekday_(

      targetDate

    );





  /*

   * =========================================

   * 当日の出勤予定を取得

   * =========================================

   */



  const currentLastRow =

    currentSheet.getLastRow();





  if (

    currentLastRow < 2

  ) {



    const message = [



      "CAST FLOW｜本日の出勤確認",



      "",



      displayDate +

        "（" +

        weekday +

        "）",



      "",



      "本日の出勤予定者はいません。"



    ].join(

      "\n"

    );





    sendAdminLineMessage_(

      message

    );





    return {



      success:

        true,



      scheduledCount:

        0,



      confirmedCount:

        0,



      okCount:

        0,



      changeCount:

        0,



      unconfirmedCount:

        0



    };



  }





  const shiftValues =

    currentSheet

      .getRange(

        2,

        1,

        currentLastRow - 1,

        13

      )

      .getValues();





  const workingRows =

    shiftValues.filter(

      row => {



        const rowDate =

          formatDateValue_(

            row[2]

          );





        const status =

          String(

            row[6] || ""

          ).trim();





        return (

          rowDate === dateString &&

          status === "出勤"

        );



      }

    );





  /*

   * =========================================

   * 今日の出勤予定が0人

   * =========================================

   */



  if (

    workingRows.length === 0

  ) {



    const message = [



      "CAST FLOW｜本日の出勤確認",



      "",



      displayDate +

        "（" +

        weekday +

        "）",



      "",



      "本日の出勤予定者はいません。"



    ].join(

      "\n"

    );





    sendAdminLineMessage_(

      message

    );





    return {



      success:

        true,



      scheduledCount:

        0,



      confirmedCount:

        0,



      okCount:

        0,



      changeCount:

        0,



      unconfirmedCount:

        0



    };



  }





  /*

   * =========================================

   * ATTENDANCE_CONFIRMATIONSを取得

   * =========================================

   */



  const confirmationMap =

    {};





  const confirmationLastRow =

    confirmationSheet.getLastRow();





  if (

    confirmationLastRow >= 2

  ) {



    const confirmationValues =

      confirmationSheet

        .getRange(

          2,

          1,

          confirmationLastRow - 1,

          9

        )

        .getValues();





    confirmationValues.forEach(

      row => {



        const rowDate =

          formatDateValue_(

            row[1]

          );





        if (

          rowDate !== dateString

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





        confirmationMap[

          castId

        ] = {



          status:

            String(

              row[5] || ""

            ).trim(),



          confirmedAt:

            row[6] || "",



          updatedAt:

            row[7] || ""



        };



      }

    );



  }





  /*

   * =========================================

   * 状態別に分類

   * =========================================

   */



  const okList =

    [];





  const changeList =

    [];





  const pendingList =

    [];





  workingRows.forEach(

    row => {



      const castId =

        String(

          row[4] || ""

        ).trim();





      const castName =

        String(

          row[5] || ""

        ).trim();





      const startTime =

        formatTimeValue_(

          row[7]

        );





      const endTime =

        formatTimeValue_(

          row[8]

        );





      const item = {



        castId:

          castId,



        castName:

          castName,



        timeText:

          startTime +

          "〜" +

          endTime



      };





      const confirmation =

        confirmationMap[

          castId

        ];





      const confirmationStatus =

        confirmation



          ? confirmation.status



          : "確認待ち";





      if (

        confirmationStatus ===

        "出勤OK"

      ) {



        okList.push(

          item

        );



      } else if (

        confirmationStatus ===

        "変更あり"

      ) {



        changeList.push(

          item

        );



      } else {



        pendingList.push(

          item

        );



      }



    }

  );





  /*

   * =========================================

   * 件数

   * =========================================

   */



  const scheduledCount =

    workingRows.length;





  const confirmedCount =

    okList.length +

    changeList.length;





  const unconfirmedCount =

    pendingList.length;





  /*

   * =========================================

   * LINE本文

   * =========================================

   */



  const lines = [



    "CAST FLOW｜本日の出勤確認",



    "",



    displayDate +

      "（" +

      weekday +

      "）"



  ];





  /*

   * =========================================

   * 未確認あり

   * =========================================

   *

   * 一番重要なので最優先表示。

   */



  if (

    unconfirmedCount > 0

  ) {



    lines.push(



      "",



      "⚠️ 未確認：" +

        unconfirmedCount +

        "名",



      "",



      "【要確認】"



    );





    pendingList.forEach(

      item => {



        lines.push(



          "・" +

          item.castName +

          "　" +

          item.timeText



        );



      }

    );





    lines.push(



      "",



      "確認済み：" +

        confirmedCount +

        " / " +

        scheduledCount +

        "名"



    );





    /*

     * 未確認がいても

     * 変更ありは管理者に知らせる

     */



    if (

      changeList.length > 0

    ) {



      lines.push(



        "",



        "【変更あり】"



      );





      changeList.forEach(

        item => {



          lines.push(



            "・" +

            item.castName +

            "　" +

            item.timeText



          );



        }

      );



    }



  }





  /*

   * =========================================

   * 全員確認済み

   * =========================================

   */



  else {



    lines.push(



      "",



      "出勤予定：" +

        scheduledCount +

        "名",



      "全員確認済み"



    );





    /*

     * 変更あり

     */



    if (

      changeList.length > 0

    ) {



      lines.push(



        "",



        "【変更あり】"



      );





      changeList.forEach(

        item => {



          lines.push(



            "・" +

            item.castName +

            "　" +

            item.timeText



          );



        }

      );



    }





    /*

     * 全員OK・変更なし

     */



    else {



      lines.push(



        "",



        "変更あり：なし"



      );



    }



  }





  /*

   * =========================================

   * 管理者LINEへ送信

   * =========================================

   */



  const message =

    lines.join(

      "\n"

    );





  sendAdminLineMessage_(

    message

  );





  /*

   * =========================================

   * 実行結果

   * =========================================

   */



  return {



    success:

      true,



    date:

      dateString,



    scheduledCount:

      scheduledCount,



    confirmedCount:

      confirmedCount,



    okCount:

      okList.length,



    changeCount:

      changeList.length,



    unconfirmedCount:

      unconfirmedCount



  };



}





/**

 * 今日の出勤確認状況を

 * 管理者へ送信する。

 *

 * 15時台の時間主導トリガー用。

 */

function sendTodayAttendanceAdminSummary() {



  return sendDailyAttendanceAdminSummary_(

    new Date()

  );



}





/* =====================================================

 * 12. 未確認キャスト再通知

 * ===================================================== */



/**

 * 指定日の出勤確認が「確認待ち」のキャストへ

 * 再通知LINEを送信する。

 *

 * @param {Date} targetDate

 * @return {Object}

 */

function resendPendingAttendanceConfirmations_(

  targetDate

) {



  const currentSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.SHIFT_CURRENT

    );





  const confirmationSheet =

    getRequiredSheet_(

      CONFIG.SHEETS.ATTENDANCE_CONFIRMATIONS

    );





  const dateString =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "yyyy-MM-dd"

    );





  const displayDate =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "M月d日"

    );





  const weekday =

    getJapaneseWeekday_(

      targetDate

    );





  /*

   * =========================================

   * 最新の当日シフトを取得

   * =========================================

   */



  const currentLastRow =

    currentSheet.getLastRow();





  if (

    currentLastRow < 2

  ) {



    return {

      success: true,

      targetCount: 0,

      sentCount: 0,

      failedCount: 0

    };



  }





  const shiftValues =

    currentSheet

      .getRange(

        2,

        1,

        currentLastRow - 1,

        13

      )

      .getValues();





  const shiftMap = {};





  shiftValues.forEach(

    row => {



      const rowDate =

        formatDateValue_(

          row[2]

        );





      const status =

        String(

          row[6] || ""

        ).trim();





      if (

        rowDate !== dateString ||

        status !== "出勤"

      ) {

        return;

      }





      const castId =

        String(

          row[4] || ""

        ).trim();





      if (!castId) {

        return;

      }





      shiftMap[

        castId

      ] = {



        castName:

          String(

            row[5] || ""

          ).trim(),



        startTime:

          formatTimeValue_(

            row[7]

          ),



        endTime:

          formatTimeValue_(

            row[8]

          )



      };



    }

  );





  /*

   * =========================================

   * ATTENDANCE_CONFIRMATIONS取得

   * =========================================

   */



  const lastRow =

    confirmationSheet.getLastRow();





  if (

    lastRow < 2

  ) {



    return {

      success: true,

      targetCount: 0,

      sentCount: 0,

      failedCount: 0

    };



  }





  const values =

    confirmationSheet

      .getRange(

        2,

        1,

        lastRow - 1,

        9

      )

      .getValues();





  /*

   * =========================================

   * 確認待ちだけ抽出

   * =========================================

   */



  const pendingRows =

    values.filter(

      row => {



        const rowDate =

          formatDateValue_(

            row[1]

          );





        const status =

          String(

            row[5] || ""

          ).trim();





        return (

          rowDate === dateString &&

          status === "確認待ち"

        );



      }

    );





  let sentCount =

    0;





  let failedCount =

    0;





  /*

   * =========================================

   * 未確認者へ再通知

   * =========================================

   */



  pendingRows.forEach(

    row => {



      const castId =

        String(

          row[2] || ""

        ).trim();





      const savedCastName =

        String(

          row[3] || ""

        ).trim();





      const lineUserId =

        String(

          row[4] || ""

        ).trim();





      /*

       * 現在のSHIFT_CURRENTに

       * 出勤予定として存在するか確認

       *

       * 13時以降に休みへ変更された場合は

       * 再通知しない。

       */

      const shift =

        shiftMap[

          castId

        ];





      if (

        !shift ||

        !lineUserId

      ) {

        return;

      }





      const castName =

        shift.castName ||

        savedCastName;





      const message = [



        "CAST FLOW",



        "",



        "出勤確認がまだ完了していません。",



        "",



        displayDate +

          "（" +

          weekday +

          "）",



        castName,



        "",



        "出勤 " +

          shift.startTime +

          "〜" +

          shift.endTime,



        "",



        "お手数ですが、下のボタンから確認してください。"



      ].join(

        "\n"

      );





      try {



        /*

         * 10番の既存関数をそのまま利用。

         *

         * 出勤OK / 変更あり

         * の2ボタン付きで再送する。

         */

        pushAttendanceConfirmationMessage_(

          lineUserId,

          message,

          dateString

        );





        sentCount++;





      } catch (

        error

      ) {



        failedCount++;





        writeErrorLog_(

          "resendPendingAttendanceConfirmations_",

          error,

          {



            date:

              dateString,



            castId:

              castId,



            castName:

              castName



          }

        );



      }



    }

  );





  return {



    success:

      true,



    date:

      dateString,



    targetCount:

      pendingRows.length,



    sentCount:

      sentCount,



    failedCount:

      failedCount



  };



}





/**

 * 今日の未確認キャストだけに

 * 出勤確認LINEを再送する。

 *

 * 14時台の時間主導トリガー用。

 */

function resendTodayPendingAttendanceConfirmations() {



  return resendPendingAttendanceConfirmations_(

    new Date()

  );



}





/* =====================================================

 * 13. 当日変更受付通知

 * ===================================================== */



/**

 * 「変更あり」から内容を保存したあと、

 *

 * ・キャスト本人

 * ・管理者

 *

 * へ最新内容をLINE通知する。

 *

 * @param {Object} user

 * @param {string} dateString YYYY-MM-DD

 * @param {Object|null} shift

 * @param {Array} customerVisits

 */

function sendAttendanceChangeAcceptedNotifications_(

  user,

  dateString,

  shift,

  customerVisits

) {



  const targetDate =

    parseDateString_(

      dateString

    );





  const displayDate =

    Utilities.formatDate(

      targetDate,

      CONFIG.TIMEZONE,

      "M月d日"

    );





  const weekday =

    getJapaneseWeekday_(

      targetDate

    );





  /*

   * =========================================

   * 変更後内容

   * =========================================

   */



  const detailLines =

    buildAttendanceChangeDetailLines_(

      shift,

      customerVisits

    );





  /*

   * =========================================

   * キャスト本人

   * =========================================

   */



  const castMessage = [



    "CAST FLOW",



    "",



    "変更を受け付けました。",



    "",



    displayDate +

      "（" +

      weekday +

      "）",



    "",



    "【変更後】",



    ...detailLines,



    "",



    "内容をご確認ください。"



  ].join(

    "\n"

  );





  /*

   * 本人への送信

   */

  pushLineMessage_(

    user.lineUserId,

    castMessage

  );





  /*

   * =========================================

   * 管理者

   * =========================================

   */



  const adminMessage = [



    "CAST FLOW｜出勤予定変更",



    "",



    user.castName +

      "さんから変更を受け付けました。",



    "",



    displayDate +

      "（" +

      weekday +

      "）",



    "",



    "【変更後】",



    ...detailLines,



    "",



    "キャストID：" +

      user.castId



  ].join(

    "\n"

  );





  /*

   * 管理者全員へ送信

   */

  sendAdminLineMessage_(

    adminMessage

  );





  return true;



}





/**

 * 変更後内容の表示行を作る

 *

 * @param {Object|null} shift

 * @param {Array} customerVisits

 * @return {string[]}

 */

function buildAttendanceChangeDetailLines_(

  shift,

  customerVisits

) {



  const lines =

    [];





  /*

   * =========================================

   * シフト

   * =========================================

   */



  if (!shift) {



    lines.push(

      "シフト：登録なし"

    );





  } else if (

    shift.status === "working"

  ) {



    const startTime =

      shift.startTime || "-";





    const endTime =

      shift.endTime || "-";





    lines.push(

      "出勤：" +

      startTime +

      "〜" +

      endTime

    );





  } else if (

    shift.status === "holiday"

  ) {



    lines.push(

      "勤務区分：休み"

    );





  } else {



    lines.push(

      "勤務区分：未定"

    );



  }





  /*

   * =========================================

   * 同伴・来店

   * =========================================

   */



  const visits =

    Array.isArray(

      customerVisits

    )



      ? customerVisits



      : [];





  const companions =

    visits.filter(

      item =>

        item.type === "同伴"

    );





  const arrivals =

    visits.filter(

      item =>

        item.type === "来店"

    );





  /*

   * 同伴

   */

  if (

    companions.length === 0

  ) {



    lines.push(

      "同伴：なし"

    );





  } else {



    companions.forEach(

      (

        item,

        index

      ) => {



        const prefix =

          index === 0



            ? "同伴："



            : "　　　";





        lines.push(

          prefix +

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





  /*

   * 来店

   */

  if (

    arrivals.length === 0

  ) {



    lines.push(

      "来店：なし"

    );





  } else {



    arrivals.forEach(

      (

        item,

        index

      ) => {



        const prefix =

          index === 0



            ? "来店："



            : "　　　";





        lines.push(

          prefix +

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





  /*

   * =========================================

   * 備考

   * =========================================

   */



  if (

    shift &&

    shift.note

  ) {



    lines.push(

      "備考：" +

      shift.note

    );



  }





  return lines;



}