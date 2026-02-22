/**
 * スプレッドシートへのデータ書き込み処理
 */
function recordToSpreadsheet(extractedData, fileUrl, event) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 対象の年を取得（"2026/02/22" -> "2026"）
    const yearStr = extractedData.date.split('/')[0];
    let sheet = ss.getSheetByName(yearStr);

    if (!sheet) {
        // シートが存在しない場合は新規作成し、ヘッダーを設定
        sheet = ss.insertSheet(yearStr);
        const headers = [
            "入力元", "日付", "支払先", "名目", "総額", "按分率",
            "経費計上額", "経費フラグ", "証憑URL", "備考/修正メモ", "重複警告"
        ];
        sheet.appendRow(headers);
        // 最初の行を固定
        sheet.setFrozenRows(1);
    }

    // 入力元の判定（SlackからのWebhookの場合は"Slack"）
    const source = "Slack";

    // デフォルトの按分率（仕様書のプリセットに応じて今回は簡易的に1.0とする。高度な実装は後続フェーズ）
    const presetRatio = 1.0;

    // データ配列の作成
    const rowData = [
        source,                   // A列: 入力元
        extractedData.date,       // B列: 日付
        extractedData.storeName,  // C列: 支払先
        extractedData.category,   // D列: 名目
        extractedData.totalAmount,// E列: 総額
        presetRatio,              // F列: 按分率
        `=Erow*Frow`,             // G列: 経費計上額（実際には書き込み行番号に置換される）
        "ON",                     // H列: 経費フラグ
        fileUrl,                  // I列: 証憑URL
        `OCR生データ: ${extractedData.rawText.substring(0, 50)}...`, // J列: 備考/修正メモ
        ""                        // K列: 重複警告
    ];

    // シートの最終行の次の行に追記
    sheet.appendRow(rowData);

    // G列の数式（=E行番号*F行番号）を実際の行番号に置換する補正処理
    const lastRow = sheet.getLastRow();
    const formulaCell = sheet.getRange(lastRow, 7); // G列
    formulaCell.setFormula(`=E${lastRow}*F${lastRow}`);
}
