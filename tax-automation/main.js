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

  // 1. SlackのURL検証(url_verification)対応、およびボタン押下(Interactive)対応
  let postData = null;
  let interactivePayload = null;

  // SlackのコマンドやInteractiveメッセージは application/x-www-form-urlencoded で payload というキーに入ってくる
  if (e && e.parameter && e.parameter.payload) {
    interactivePayload = JSON.parse(e.parameter.payload);
  } else if (e && e.postData && e.postData.contents) {
    postData = JSON.parse(e.postData.contents);
  }

  // --- インタラクティブ (ボタン押下) の処理 ---
  if (interactivePayload) {
    return handleSlackInteractivePayload(interactivePayload);
  }

  // --- URL検証の処理 ---
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

    // --- Slackの自動リトライ（3秒ルール）対策 ---
    // SlackからのイベントIDを取得 (postData.event_id に存在)
    const eventId = postData.event_id;
    if (eventId) {
      const cache = CacheService.getScriptCache();
      const isProcessed = cache.get(eventId);
      if (isProcessed) {
        // すでに処理済み（あるいは処理中）のイベントなので無視
        console.log(`Duplicate event ignored: ${eventId}`);
        return ContentService.createTextOutput("OK");
      }
      // キャッシュに保存（6分間保持 = 360秒）
      // ※Slackのリトライは通常最大5回（最初から数えて約5分以内）行われるため
      cache.put(eventId, 'processed', 360);
    }
    // ------------------------------------------

    try {
      // 添付された全てのファイルをループして処理
      for (let i = 0; i < event.files.length; i++) {
        const file = event.files[i];

        // 画像ファイルまたはPDFかどうかの確認 (MIMEタイプが空にされる場合へのフォールバック対応)
        const mime = file.mimetype || "";
        const isImage = mime.indexOf('image/') === 0 || file.name.match(/\.(jpg|jpeg|png|gif)$/i);
        const isPdf = mime === 'application/pdf' || file.name.match(/\.pdf$/i);

        if (!isImage && !isPdf) {
          continue; // 対象外のファイルはスキップして次へ
        }

        try {
          // 3. Driveへの保存処理
          const driveFile = saveFileToDrive(file, SLACK_BOT_TOKEN);

          // 4. OCR処理による文字抽出
          const ocrText = extractTextWithOCR(driveFile.getId());

          // 設定シートの準備と読み込み (sheet_handler.js側で定義した関数を利用)
          const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
          const settingsMap = ensureSettingsSheet(ss);

          // 5. 名目推論や金額などの抽出
          const extractedData = parseOCRText(ocrText, settingsMap);

          // 6. スプレッドシートへの記録
          const recordResult = recordToSpreadsheet(extractedData, driveFile.getUrl(), event);

          // 7. Slackへの対話型ボタン付き通知（Approve / Reject）
          sendSlackInteractiveMessage(event.channel, extractedData, recordResult, settingsMap);
        } catch (fileError) {
          // 個別のファイル処理でのエラーは、ここだけでキャッチし、次のループへ進む
          console.error(`Error processing file ${file.name}:`, fileError);
          try {
            const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
            let logSheet = ss.getSheetByName("System_ErrorLog");
            if (!logSheet) {
              logSheet = ss.insertSheet("System_ErrorLog");
              logSheet.appendRow(["発生日時", "エラー内容", "スタックトレース", "イベント内容(JSON)"]);
            }
            logSheet.appendRow([
              new Date(),
              `[File: ${file.name}] ` + fileError.toString(),
              fileError.stack,
              JSON.stringify(event)
            ]);
          } catch (logError) {
            console.error("Failed to write individual file error log", logError);
          }
        }
      }
    } catch (globalError) {
      // ループ外などの致命的なエラー
      try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let logSheet = ss.getSheetByName("System_ErrorLog");
        if (!logSheet) {
          logSheet = ss.insertSheet("System_ErrorLog");
          logSheet.appendRow(["発生日時", "エラー内容", "スタックトレース", "イベント内容(JSON)"]);
        }
        logSheet.appendRow([
          new Date(),
          globalError.toString(),
          globalError.stack,
          JSON.stringify(event)
        ]);
      } catch (logError) {
        console.error("Failed to write global error log to spreadsheet", logError);
      }
    }
  }

  return ContentService.createTextOutput("OK");
}

/**
 * OCRテキストからGemini APIを利用して金額や店舗名、カテゴリを抽出する
 */
function parseOCRText(text, settingsMap) {
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

  // 利用可能なGeminiモデルを動的に取得して使用する（バージョンアップによる404エラー対策）
  let modelName = "models/gemini-2.0-flash"; // デフォルト
  try {
    const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${GEMINI_API_KEY}`;
    const listResponse = UrlFetchApp.fetch(listUrl, { muteHttpExceptions: true });
    if (listResponse.getResponseCode() === 200) {
      const data = JSON.parse(listResponse.getContentText());
      if (data.models && data.models.length > 0) {
        // flashモデルかつgenerateContent対応のものを優先
        const flashModel = data.models.find(m => m.name.includes("flash") && m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"));
        if (flashModel) {
          modelName = flashModel.name;
        } else {
          // なければgenerateContent対応の最初のモデル
          const anyModel = data.models.find(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"));
          if (anyModel) modelName = anyModel.name;
        }
      }
    }
  } catch (e) {
    console.error("Failed to list models: " + e);
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${GEMINI_API_KEY}`;

  // 利用可能なカテゴリーのリストを生成
  const availableCategories = Object.keys(settingsMap || {}).join("、");
  const categoryHint = availableCategories ? availableCategories : "消耗品費、交通費、交際費、会議費など";

  // プロンプトの構築
  const prompt = `
以下のテキストはレシート（または領収書）をOCRで読み取った結果です。
このテキストから以下の4つの情報を抽出し、JSONフォーマットのみで出力してください。Markdownのコードブロック( \`\`\`json など )は含めないでください。

【抽出項目】
1. "date": 支払日付 (yyyy/MM/dd形式。年が不明な場合は推測するか現在の年を使用)
2. "storeName": 店舗名や支払先
3. "totalAmount": 合計金額 (数値のみ。カンマや「円」などは取り除く)
4. "category": 経費のカテゴリ（必ず以下のいずれかから推測して完全一致で選択してください: ${categoryHint}。推測できない場合は "未分類"）

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
      logToErrorSheet(`Gemini API HTTP Error: ${responseCode}`, responseBody, text);
    }
  } catch (e) {
    console.error(`parseOCRText Error: ${e.message}\n${e.stack}`);
    logToErrorSheet(`parseOCRText Catch Error: ${e.message}`, e.stack, text);
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

/**
 * プロセス中に発生した非致死的なエラーをスプレッドシートに記録する
 */
function logToErrorSheet(errorMsg, detail, extraInfo) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName("System_ErrorLog");
    if (!logSheet) {
      logSheet = ss.insertSheet("System_ErrorLog");
      logSheet.appendRow(["発生日時", "エラー内容", "詳細（スタック等）", "付加情報"]);
    }
    logSheet.appendRow([
      new Date(),
      errorMsg,
      detail,
      extraInfo
    ]);
  } catch (e) {
    console.error("Failed to write to ErrorSheet: " + e.toString());
  }
}

