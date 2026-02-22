// 定数設定
const PROPERTIES = PropertiesService.getScriptProperties();
// 以下のプロパティを事前にGASの設定から登録しておく必要があります
const SLACK_BOT_TOKEN = PROPERTIES.getProperty('SLACK_BOT_TOKEN');
const SPREADSHEET_ID = PROPERTIES.getProperty('SPREADSHEET_ID');
const TARGET_CHANNEL_ID = PROPERTIES.getProperty('TARGET_CHANNEL_ID'); // 対象のチャンネルID（例: C01234567）

/**
 * SlackからのEvents API / Webhooks 受信エンドポイント
 */
function doPost(e) {
  // --- [DEBUG] 全リクエストのダンプを最優先で記録 ---
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let dumpSheet = ss.getSheetByName("System_RequestDump");
    if (!dumpSheet) {
      dumpSheet = ss.insertSheet("System_RequestDump");
      dumpSheet.appendRow(["日時", "生データ全文(JSON)"]);
    }
    const rawData = e ? JSON.stringify(e) : "No event object";
    dumpSheet.appendRow([new Date(), rawData]);
  } catch (logErr) {
    console.error("Dump failed", logErr);
  }
  // --------------------------------------------------

  // 1. SlackのURL検証(url_verification)対応
  const postData = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : null;
  if (postData && postData.type === 'url_verification') {
    return ContentService.createTextOutput(postData.challenge);
  }

  // 2. イベントデータの取得
  if (!postData || !postData.event) {
    return ContentService.createTextOutput("OK");
  }

  const event = postData.event;

  // メッセージイベントかつファイルが存在する場合に処理
  if (event.type === 'message' && event.files && event.files.length > 0) {
    // 自身のbotメッセージは無視
    if (event.bot_id) return ContentService.createTextOutput("OK");

    // 指定された対象チャンネルからの投稿のみ処理する
    if (TARGET_CHANNEL_ID && event.channel !== TARGET_CHANNEL_ID) {
      return ContentService.createTextOutput("OK");
    }

    try {
      // 最初のファイルを処理（複数ある場合は拡張可能）
      const file = event.files[0];

      // 画像ファイルかどうかの確認
      if (file.mimetype.indexOf('image/') !== 0 && file.mimetype !== 'application/pdf') {
        return ContentService.createTextOutput("OK");
      }

      // 3. Driveへの保存処理
      const driveFile = saveFileToDrive(file, SLACK_BOT_TOKEN);

      // 4. OCR処理による文字抽出
      const ocrText = extractTextWithOCR(driveFile.getId());

      // 5. 名目推論や金額などの抽出（簡易実装）
      const extractedData = parseOCRText(ocrText);

      // 6. スプレッドシートへの記録
      recordToSpreadsheet(extractedData, driveFile.getUrl(), event);

      // 7. （オプション）Slackへのフィードバック通知
      // notifyToSlack(event.channel, "レシートを処理しました。", extractedData);

    } catch (error) {
      console.error("Error processing message: " + error.toString() + "\nStack: " + error.stack);

      // 原因究明のため、スプレッドシート側に強制的にエラーログを書き出す
      try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let logSheet = ss.getSheetByName("System_ErrorLog");
        if (!logSheet) {
          logSheet = ss.insertSheet("System_ErrorLog");
          logSheet.appendRow(["発生日時", "エラー内容", "スタックトレース", "イベント内容(JSON)"]);
        }
        logSheet.appendRow([
          new Date(),
          error.toString(),
          error.stack,
          JSON.stringify(event)
        ]);
      } catch (logError) {
        // ログ書き込みすら失敗した場合（SPREADSHEET_IDのミス等）
        console.error("Failed to write error log to spreadsheet", logError);
      }

      return ContentService.createTextOutput("OK"); // 返却値自体はOKを返し、Slackからのリトライ地獄を防ぐ
    }
  }

  return ContentService.createTextOutput("OK");
}

/**
 * OCRテキストから簡易的に金額や店舗名を抽出する
 */
function parseOCRText(text) {
  // ※実装計画に基づいて、まずは簡易な抽出ロジックを用意
  const data = {
    date: Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"),
    storeName: "OCR抽出結果要確認",
    totalAmount: 0,
    category: "未分類",
    rawText: text
  };

  // 金額の抽出（「合計 ¥1,000」などを想定したシンプルな正規表現）
  const amountMatch = text.match(/[\¥\\]\s*([0-9,]+)/);
  if (amountMatch && amountMatch[1]) {
    data.totalAmount = parseInt(amountMatch[1].replace(/,/g, ''), 10);
  }

  return data;
}
