/* =====================================================

 * ROYAL SHIFT SYSTEM

 * WebApp.gs

 * ===================================================== */





/* =====================================================

 * 01. Webアプリ表示 / API

 * ===================================================== */



function doGet(e) {



  const action =

    e &&

    e.parameter &&

    e.parameter.action



      ? String(e.parameter.action).trim()



      : "";





  if (action) {



    return handleApiGet_(

      e

    );



  }





  const page =

    e &&

    e.parameter &&

    e.parameter.page



      ? e.parameter.page



      : "home";





  const template =

    HtmlService

      .createTemplateFromFile(

        "Html"

      );





  template.initialPage =

    page;





  return template

    .evaluate()

    .setTitle(

      "ROYAL STAFF"

    )

    .addMetaTag(

      "viewport",

      "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"

    )

    .setXFrameOptionsMode(

      HtmlService.XFrameOptionsMode.ALLOWALL

    );



}





/* =====================================================

 * 02. POST API

 * ===================================================== */



function doPost(e) {



  try {



    const body =

      parseRequestBody_(

        e

      );





    const action =

      String(

        body.action || ""

      ).trim();





    if (!action) {



      return createJsonResponse_({



        success:

          false,



        error:

          "actionが指定されていません"



      });



    }





    /* =====================================================

     * Cloudflare Webhook：出勤OK

     * ===================================================== */



    if (

      action ===

      "webhookAttendanceOk"

    ) {



      /*

       * Cloudflare → GAS専用共有鍵

       */

      const expectedSecret =

        String(

          PropertiesService

            .getScriptProperties()

            .getProperty(

              "CLOUDFLARE_WEBHOOK_SECRET"

            ) || ""

        ).trim();





      const receivedSecret =

        String(

          body.webhookSecret || ""

        ).trim();





      if (!expectedSecret) {



        throw new Error(

          "Webhook共有鍵がGASに設定されていません"

        );



      }





      if (

        !receivedSecret ||

        receivedSecret !== expectedSecret

      ) {



        throw new Error(

          "Webhook認証に失敗しました"

        );



      }





      const lineUserId =

        String(

          body.lineUserId || ""

        ).trim();





      const dateString =

        String(

          body.date || ""

        ).trim();





      if (!lineUserId) {



        throw new Error(

          "LINEユーザーIDがありません"

        );



      }





      if (!dateString) {



        throw new Error(

          "対象日がありません"

        );



      }





      return createJsonResponse_(



        confirmAttendanceOkByWebhook(

          lineUserId,

          dateString

        )



      );



    }





    /* =====================================================

     * 認証状態確認

     * ===================================================== */



    if (

      action ===

      "authStatus"

    ) {



      return createJsonResponse_(



        getLineAuthStatus_(

          body.idToken

        )



      );



    }





    /* =====================================================

     * シフト本人認証テスト

     * ===================================================== */



    if (

      action ===

      "shiftAuthTest"

    ) {



      return createJsonResponse_(



        getShiftUserAuthTest_(

          body.idToken

        )



      );



    }





    /* =====================================================

     * LINE認証：月間シフト取得

     * ===================================================== */



    if (

      action ===

      "getMonthlyShiftData"

    ) {



      return createJsonResponse_(



        getMonthlyShiftDataByLine(

          body.idToken,

          Number(body.year),

          Number(body.month)

        )



      );



    }





    /* =====================================================

     * LINE認証：シフト保存

     * ===================================================== */



    if (

      action ===

      "saveShiftEntry"

    ) {



      return createJsonResponse_(



        saveShiftEntryByLine(

          body.idToken,

          {



            date:

              body.date,



            status:

              body.status,



            startTime:

              body.startTime,



            endTime:

              body.endTime,



            transport:

              String(

                body.transport

              ) === "true",



            note:

              body.note



          }

        )



      );



    }





    /* =====================================================

     * LINE認証：シフト削除

     * ===================================================== */



    if (

      action ===

      "deleteShiftEntry"

    ) {



      return createJsonResponse_(



        deleteShiftEntryByLine(

          body.idToken,

          body.date

        )



      );



    }





    /* =====================================================

     * LINE認証：月間シフト提出

     * ===================================================== */



    if (

      action ===

      "submitMonthlyShift"

    ) {



      return createJsonResponse_(



        submitMonthlyShiftByLine(

          body.idToken,

          Number(

            body.year

          ),

          Number(

            body.month

          )

        )



      );



    }





    /* =====================================================

     * LINE認証：曜日一括入力

     * ===================================================== */



    if (

      action ===

      "saveBulkShifts"

    ) {



      let weekdays = [];





      try {



        weekdays =

          JSON.parse(

            body.weekdays || "[]"

          );



      } catch (error) {



        weekdays = [];



      }





      return createJsonResponse_(



        saveBulkShiftsByLine(

          body.idToken,

          {



            targetMonth:

              body.targetMonth,



            weekdays:

              weekdays,



            overwrite:

              String(

                body.overwrite

              ) === "true",



            status:

              body.status,



            startTime:

              body.startTime,



            endTime:

              body.endTime,



            note:

              body.note



          }

        )



      );



    }





    /* =====================================================

     * LINE認証：選択日一括入力

     * ===================================================== */



    if (

      action ===

      "saveMultiShifts"

    ) {



      let dates = [];





      try {



        dates =

          JSON.parse(

            body.dates || "[]"

          );



      } catch (error) {



        dates = [];



      }





      return createJsonResponse_(



        saveMultiShiftsByLine(

          body.idToken,

          {



            dates:

              dates,



            status:

              body.status,



            startTime:

              body.startTime,



            endTime:

              body.endTime,



            note:

              body.note,



            overwrite:

              String(

                body.overwrite

              ) === "true"



          }

        )



      );



    }





    /* =====================================================

     * LINE認証：契約時間取得

     * ===================================================== */



    if (

      action ===

      "getCurrentCastDefaultTimes"

    ) {



      return createJsonResponse_(



        getCurrentCastDefaultTimesByLine(

          body.idToken

        )



      );



    }





    /* =====================================================

     * LINE認証：当日出勤OK

     * ===================================================== */



    if (

      action ===

      "confirmAttendanceOk"

    ) {



      return createJsonResponse_(



        confirmAttendanceOkByLine(

          body.idToken,

          body.date

        )



      );



    }





    /* =====================================================

     * LINE認証：当日出勤変更あり

     * ===================================================== */



    if (

      action ===

      "confirmAttendanceChange"

    ) {



      return createJsonResponse_(



        confirmAttendanceChangeByLine(

          body.idToken,

          body.date

        )



      );



    }





    /* =====================================================

     * LINE認証：同伴・来店予定取得

     * ===================================================== */



    if (

      action ===

      "getMyCustomerVisits"

    ) {



      return createJsonResponse_(



        getMyCustomerVisitsByLine(

          body.idToken,

          body.date

        )



      );



    }





    /* =====================================================

     * LINE認証：同伴・来店予定保存

     * ===================================================== */



    if (

      action ===

      "saveMyCustomerVisits"

    ) {



      let visits = [];





      try {



        /*

         * GitHub側からJSON文字列で来ても

         * 配列で来ても両方受け付ける

         */

        if (

          Array.isArray(

            body.visits

          )

        ) {



          visits =

            body.visits;



        } else {



          visits =

            JSON.parse(

              body.visits || "[]"

            );



        }



      } catch (error) {



        throw new Error(

          "同伴・来店予定のデータ形式が不正です。"

        );



      }





      return createJsonResponse_(



        saveMyCustomerVisitsByLine(

          body.idToken,

          body.date,

          visits

        )



      );



    }





    /* =====================================================

     * 初期管理者登録

     * ===================================================== */



    if (

      action ===

      "registerInitialAdmin"

    ) {



      return createJsonResponse_(



        registerInitialAdmin_(

          body.idToken,

          body.castName

        )



      );



    }





    /* =====================================================

     * 一般キャスト登録申請

     * ===================================================== */



    if (

      action ===

      "submitRegistrationRequest"

    ) {



      return createJsonResponse_(



        submitRegistrationRequest_(

         body.idToken,

         body.castId,

         body.castName

         )



        );



    }





            /* =====================================================
     * 管理者：本日の出勤確認
     * ===================================================== */

    if (
      action ===
      "getAdminTodayAttendance"
    ) {

      return createJsonResponse_(

        getAdminTodayAttendance_(
          body.idToken
        )

      );

    }


/* =====================================================
     * 管理者：月間提出状況取得
     * ===================================================== */

    if (
      action ===
      "getMonthlySubmissionStatus"
    ) {

      return createJsonResponse_(

        getMonthlySubmissionStatus_(
          body.idToken,
          Number(
            body.year
          ),
          Number(
            body.month
          )
        )

      );

    }


/* =====================================================

     * 管理者：登録申請一覧取得

     * ===================================================== */



    if (

      action ===

      "getRegistrationRequests"

    ) {



      return createJsonResponse_(



        getRegistrationRequests_(

          body.idToken

        )



      );



    }





    /* =====================================================

     * 管理者：登録申請承認

     * ===================================================== */



    if (

      action ===

      "approveRegistrationRequest"

    ) {



      return createJsonResponse_(



        approveRegistrationRequest_(

          body.idToken,

          body.requestId

        )



      );



    }





    /* =====================================================

     * 管理者：登録申請却下

     * ===================================================== */



    if (

      action ===

      "rejectRegistrationRequest"

    ) {



      return createJsonResponse_(



        rejectRegistrationRequest_(

          body.idToken,

          body.requestId,

          body.note

        )



      );



    }





    return createJsonResponse_({



      success:

        false,



      error:

        "不明なactionです"



    });





  } catch (error) {



    writeErrorLog_(

      "doPost",

      error,

      {}

    );





    return createJsonResponse_({



      success:

        false,



      error:

        error.message ||

        String(error)



    });



  }



}





/* =====================================================

 * 03. GET API

 * ===================================================== */



function handleApiGet_(

  e

) {



  try {



    const action =

      String(

        e.parameter.action || ""

      ).trim();





    if (

      action ===

      "ping"

    ) {



      return createJsonResponse_({



        success:

          true,



        system:

          "ROYAL SHIFT",



        message:

          "API OK",



        timestamp:

          new Date().toISOString()



      });



    }





    return createJsonResponse_({



      success:

        false,



      error:

        "不明なactionです"



    });





  } catch (error) {



    return createJsonResponse_({



      success:

        false,



      error:

        error.message ||

        String(error)



    });



  }



}





/* =====================================================

 * 04. リクエスト解析

 * ===================================================== */



function parseRequestBody_(

  e

) {



  if (

    !e ||

    !e.postData

  ) {



    return {};



  }





  const content =

    String(

      e.postData.contents || ""

    );





  /*

   * JSON形式

   */

  if (content) {



    try {



      return JSON.parse(

        content

      );



    } catch (error) {



      // JSONでなければform形式へ

    }



  }





  return e.parameter || {};



}





/* =====================================================

 * 05. JSONレスポンス

 * ===================================================== */



function createJsonResponse_(

  data

) {



  return ContentService

    .createTextOutput(

      JSON.stringify(

        data

      )

    )

    .setMimeType(

      ContentService.MimeType.JSON

    );



}





/* =====================================================

 * 06. HTML include

 * ===================================================== */



function include(

  filename

) {



  return HtmlService

    .createHtmlOutputFromFile(

      filename

    )

    .getContent();



}