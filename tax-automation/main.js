// 定数設定
const PROPERTIES = PropertiesService.getScriptProperties();
// 以下のプロパティを事前にGASの設定から登録しておく必要があります
const SLACK_BOT_TOKEN = PROPERTIES.getProperty('SLACK_BOT_TOKEN');
const SPREADSHEET_ID = PROPERTIES.getProperty('SPREADSHEET_ID');
const TARGET_CHANNEL_ID = PROPERTIES.getProperty('TARGET_CHANNEL_ID'); // 対象のチャンネルID（例: C01234567）
const GEMINI_API_KEY = PROPERTIES.getProperty('GEMINI_API_KEY'); // 追加: Gemini APIキー

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
 * OCRテキストからGemini APIを利用して金額や店舗名、カテゴリを抽出する
 */
function parseOCRText(text) {
  // 環境変数のチェック
  if (!GEMINI_API_KEY) {
    console.error("GEMINI_API_KEY is not set.");
    return {
      date: Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"),
      storeName: "GEMINI_API_KEY未設定",
      totalAmount: 0,
      category: "未分類",
      rawText: text
    };
  }

  // Gemini APIエンドポイント (gemini-1.5-flashを推奨)
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;

  // プロンプトの構築
  const prompt = `
以下のテキストはレシート（または領収書）をOCRで読み取った結果です。
このテキストから以下の4つの情報を抽出し、JSONフォーマットのみで出力してください。Markdownのコードブロック( \`\`\`json など )は含めないでください。

【抽出項目】
1. "date": 支払日付 (yyyy/MM/dd形式。年が不明な場合は推測するか現在の年を使用)
2. "storeName": 店舗名や支払先
3. "totalAmount": 合計金額 (数値のみ。カンマや「円」などは取り除く)
4. "category": 経費のカテゴリ (例: 消耗品費、交通費、交際費、会議費など。推測で構いません。不明な場合は "未分類")

【制約事項】
- 返答は必ず純粋なJSON文字列のみにしてください。
- フォーマット外の会話や説明は一切不要です。
- 合計金額が取得できなかった場合は 0 を指定してください。

【テキスト】
${text}
`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      // Geminiのレスポンス構造からテキストを抽出
      let resultText = data.candidates[0].content.parts[0].text;
      
      // もしMarkdownのコードブロックが含まれていた場合は除去する
      resultText = resultText.replace(/```json/g, "").replace(/```/g, "").trim();

      const parsedJSON = JSON.parse(resultText);

      return {
        date: parsedJSON.date || Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"),
        storeName: parsedJSON.storeName || "不明",
        totalAmount: parseInt(parsedJSON.totalAmount) || 0,
        category: parsedJSON.category || "未分類",
        rawText: text
      };
    } else {
      console.error(`Gemini API Error: ${responseCode} - ${responseBody}`);
    }
  } catch (e) {
    console.error(`parseOCRText Error: ${e.message}\n${e.stack}`);
  }

  // API呼び出しに失敗した場合のフォールバック
  return {
    date: Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd"),
    storeName: "OCR解析エラー",
    totalAmount: 0,
    category: "未分類",
    rawText: text
  };
}
