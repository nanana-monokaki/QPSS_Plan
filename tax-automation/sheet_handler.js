/**
 * スプレッドシートへのデータ書き込み処理
 */
function recordToSpreadsheet(extractedData, fileUrl, event) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 設定シートの準備と読み込み
    const settings = ensureSettingsSheet(ss);

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

        // D列（名目）とH列（経費フラグ）にプルダウン（入力規則）を設定
        applyDataValidationToCategoryColumn(sheet, ss);
        applyDataValidationToExpenseFlagColumn(sheet);
    }

    // 重複チェック
    const isDuplicate = checkDuplicate(sheet, extractedData);

    // 入力元の判定（SlackからのWebhookの場合は"Slack"）
    const source = "Slack";

    // カテゴリに応じたデフォルト按分率の適用
    let presetRatio = 1.0;
    if (settings[extractedData.category] !== undefined) {
        presetRatio = settings[extractedData.category];
    }

    // データ配列の作成
    const rowData = [
        source,                   // A列: 入力元
        extractedData.date,       // B列: 日付
        extractedData.storeName,  // C列: 支払先
        extractedData.category,   // D列: 名目
        extractedData.totalAmount,// E列: 総額
        presetRatio,              // F列: 按分率
        `=Erow*Frow`,             // G列: 経費計上額（実際には書き込み行番号に置換される）
        "経費",                   // H列: 経費フラグ (ON -> 経費)
        fileUrl,                  // I列: 証憑URL
        `OCR生データ: ${extractedData.rawText.substring(0, 50)}...`, // J列: 備考/修正メモ
        isDuplicate ? "重複の可能性あり" : "" // K列: 重複警告
    ];

    // シートの最終行の次の行に追記
    sheet.appendRow(rowData);

    // G列の数式（=E行番号*F行番号）を実際の行番号に置換する補正処理
    const lastRow = sheet.getLastRow();
    const formulaCell = sheet.getRange(lastRow, 7); // G列
    formulaCell.setFormula(`=E${lastRow}*F${lastRow}`);

    return {
        sheetName: yearStr,
        rowNumber: lastRow,
        isDuplicate: isDuplicate
    };
}

/**
 * 「設定」シートを確認し、存在しなければ作成して初期データを入れる。
 * 戻り値として { "名目名": 按分率 } のようなマップオブジェクトを返す。
 */
function ensureSettingsSheet(ss) {
    let sheet = ss.getSheetByName("設定");
    if (!sheet) {
        sheet = ss.insertSheet("設定");
        sheet.appendRow(["名目", "デフォルト按分率"]);
        sheet.setFrozenRows(1);

        // 初期プリセット
        const presets = [
            ["旅費交通費", 1.0],
            ["消耗品費", 1.0],
            ["機材費", 1.0],
            ["地代家賃", 0.5],
            ["通信費", 0.4],
            ["接待交際費", 1.0],
            ["新聞図書費", 1.0],
            ["水道光熱費", 0.3],
            ["未分類", 1.0]
        ];
        sheet.getRange(2, 1, presets.length, 2).setValues(presets);
    }

    // 設定シートのデータを読み込んでMapを返す
    let data = [];
    if (sheet.getLastRow() > 1) {
        data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    }

    const settingsMap = {};
    for (const row of data) {
        if (row[0]) {
            settingsMap[row[0]] = parseFloat(row[1]) || 1.0;
        }
    }
    return settingsMap;
}

/**
 * シートのD列（名目）にプルダウンリストを設定する
 * プルダウンのリスト元は「設定」シートのA列
 */
function applyDataValidationToCategoryColumn(sheet, ss) {
    const settingsSheet = ss.getSheetByName("設定");
    if (!settingsSheet) return;

    // 設定シートのA列（2行目以降）を範囲とする入力規則を作成
    const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(settingsSheet.getRange("A2:A"), true)
        .setAllowInvalid(true) // AI推論などで完全一致しなくてもとりあえず許容する
        .build();

    // D列（2行目以降）に設定
    sheet.getRange("D2:D").setDataValidation(rule);
}

/**
 * シートのH列（経費フラグ）に「経費」「-」のプルダウンリストを設定する
 */
function applyDataValidationToExpenseFlagColumn(sheet) {
    const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["経費", "-"], true)
        .setAllowInvalid(true)
        .build();

    // H列（2行目以降）に設定
    sheet.getRange("H2:H").setDataValidation(rule);
}

/**
 * 過去のデータを読み取って同一・類似の領収書が既に登録されていないか確認する
 * 条件: 日付が±1日以内 ＆ 総額が一致 ＆ 支払先が部分一致
 */
function checkDuplicate(sheet, extractedData) {
    if (sheet.getLastRow() <= 1) return false;

    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 4).getValues();
    // [0]=日付(B), [1]=支払先(C), [2]=名目(D), [3]=総額(E)

    const targetDate = new Date(extractedData.date).getTime();
    const targetAmount = parseInt(extractedData.totalAmount);
    const targetStore = extractedData.storeName || "";

    // 日時の差分許容値 (1日 = 86400000 ミリ秒)
    const oneDayMs = 24 * 60 * 60 * 1000;

    for (const row of data) {
        const rowDateMs = new Date(row[0]).getTime();
        const rowStore = row[1] ? row[1].toString() : "";
        const rowAmount = parseInt(row[3]) || 0;

        // 金額が一致するか
        if (targetAmount > 0 && rowAmount === targetAmount) {
            // 日付が±1日以内か
            if (Math.abs(rowDateMs - targetDate) <= oneDayMs) {
                // 支払先の名前が部分一致（または一方がもう一方を含む）するか
                if (targetStore.includes(rowStore) || rowStore.includes(targetStore)) {
                    return true;
                }
            }
        }
    }
    return false;
}

