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
        "-",                      // H列: 経費フラグ (デフォルトは「-」)
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

    // B列（日付）を昇順にソートする（月別に見やすくするため）
    if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 2, ascending: true });
    }

    // フィルタが設定されていなければ設定する
    if (!sheet.getFilter()) {
        sheet.getDataRange().createFilter();
    }

    // 月間・年間・カテゴリ別サマリーシートの作成/更新
    ensureSummarySheet(ss, yearStr, settings);

    return {
        sheetName: yearStr,
        fileUrl: fileUrl,
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
            ["地代家賃", 0.3],     // 面積の3割
            ["水道光熱費", 0.3],   // 平日8時間＋αの稼働時間で推計
            ["通信費", 0.8],       // 仕事柄常に使用している（スマホ/ネット同回線）
            ["旅費交通費", 1.0],
            ["消耗品費", 1.0],
            ["機材費", 1.0],
            ["接待交際費", 1.0],
            ["新聞図書費", 1.0],
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
 * 「サマリー」シートを確認し、存在しなければ作成する。
 * 対象の年（yearStr）に関する「全体合計」および「カテゴリ別合計」の行を追加・更新する。
 */
function ensureSummarySheet(ss, yearStr, settingsMap) {
    let sheet = ss.getSheetByName("サマリー");
    if (!sheet) {
        sheet = ss.insertSheet("サマリー", 0); // 先頭に作成
        const headers = ["年", "種別", "年間合計", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
        sheet.setFrozenColumns(2);

        // ヘッダー行の書式設定
        sheet.getRange(1, 1, 1, 15).setFontWeight("bold").setBackground("#e0e0e0");
    }

    // すでに今年度のサマリー行が存在するかチェック
    const data = sheet.getDataRange().getValues();
    let yearExists = false;
    let yearStartIdx = -1;
    for (let i = 1; i < data.length; i++) {
        // A列が対象の年（文字列表記や数値表記）の行があれば、もう作成済みとみなす
        if (data[i][0].toString() === yearStr) {
            yearExists = true;
            yearStartIdx = i;
            break;
        }
    }

    // 指定した年用のサマリー行がなければ追加
    if (!yearExists) {
        const rowsToAppend = [];

        // 1. 全体経費
        const expenseRow = [yearStr, "【全体】経費"];
        expenseRow.push(`=SUMIF('${yearStr}'!H:H, "経費", '${yearStr}'!G:G)`); // 年間合計
        for (let m = 1; m <= 12; m++) {
            const startD = `${yearStr}/${('0' + m).slice(-2)}/01`;
            const endD = `${yearStr}/${('0' + m).slice(-2)}/31`;
            expenseRow.push(`=SUMIFS('${yearStr}'!G:G, '${yearStr}'!H:H, "経費", '${yearStr}'!B:B, ">=${startD}", '${yearStr}'!B:B, "<=${endD}")`);
        }
        rowsToAppend.push(expenseRow);

        // 2. 全体経費外
        const nonExpenseRow = [yearStr, "【全体】経費外（対象外等）"];
        nonExpenseRow.push(`=SUM('${yearStr}'!E:E) - SUMIF('${yearStr}'!H:H, "経費", '${yearStr}'!G:G)`); // 総額(E) - 経費計上額(G)
        for (let m = 1; m <= 12; m++) {
            const startD = `${yearStr}/${('0' + m).slice(-2)}/01`;
            const endD = `${yearStr}/${('0' + m).slice(-2)}/31`;
            nonExpenseRow.push(`=SUMIFS('${yearStr}'!E:E, '${yearStr}'!B:B, ">=${startD}", '${yearStr}'!B:B, "<=${endD}") - SUMIFS('${yearStr}'!G:G, '${yearStr}'!H:H, "経費", '${yearStr}'!B:B, ">=${startD}", '${yearStr}'!B:B, "<=${endD}")`);
        }
        rowsToAppend.push(nonExpenseRow);

        // 3. カテゴリごとの経費
        const categories = Object.keys(settingsMap || {});
        for (const cat of categories) {
            const catRow = [yearStr, `[名目] ${cat}`];
            catRow.push(`=SUMIFS('${yearStr}'!G:G, '${yearStr}'!H:H, "経費", '${yearStr}'!D:D, "${cat}")`);
            for (let m = 1; m <= 12; m++) {
                const startD = `${yearStr}/${('0' + m).slice(-2)}/01`;
                const endD = `${yearStr}/${('0' + m).slice(-2)}/31`;
                catRow.push(`=SUMIFS('${yearStr}'!G:G, '${yearStr}'!H:H, "経費", '${yearStr}'!D:D, "${cat}", '${yearStr}'!B:B, ">=${startD}", '${yearStr}'!B:B, "<=${endD}")`);
            }
            rowsToAppend.push(catRow);
        }

        // 4. 空白行（次年度との区切り用）
        rowsToAppend.push(new Array(15).fill(""));

        // スプレッドシートへ書き込み
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rowsToAppend.length, 15).setValues(rowsToAppend);

        // 金額列（C〜O）の表示形式を数字（カンマ区切り）に設定
        sheet.getRange(startRow, 3, rowsToAppend.length, 13).setNumberFormat("#,##0");

        // 行ごとの色分けなどの装飾（任意）
        sheet.getRange(startRow, 1, 1, 15).setBackground("#fff2cc"); // 全体経費
        sheet.getRange(startRow + 1, 1, 1, 15).setBackground("#f4cccc"); // 全体経費外
    } else {
        // すでに年ブロックが存在する場合、不足している「名目（カテゴリ）」があれば追記する
        const existingCategories = new Set();
        let endIdx = yearStartIdx;

        for (let i = yearStartIdx; i < data.length; i++) {
            const rowYearStr = data[i][0].toString();
            // 対象年、または区切りの空白行である間は探索
            if (rowYearStr === yearStr || rowYearStr === "") {
                endIdx = i;
                const typeStr = data[i][1] ? data[i][1].toString() : "";
                if (typeStr.startsWith("[名目] ")) {
                    existingCategories.add(typeStr.replace("[名目] ", ""));
                }
            } else {
                break; // 次の年のブロックに入った
            }
        }

        const categories = Object.keys(settingsMap || {});
        // 設定シートにあるが、サマリーシートにないカテゴリを抽出
        const missingCategories = categories.filter(cat => !existingCategories.has(cat));

        if (missingCategories.length > 0) {
            const rowsToInsert = [];
            for (const cat of missingCategories) {
                const catRow = [yearStr, `[名目] ${cat}`];
                catRow.push(`=SUMIFS('${yearStr}'!G:G, '${yearStr}'!H:H, "経費", '${yearStr}'!D:D, "${cat}")`);
                for (let m = 1; m <= 12; m++) {
                    const startD = `${yearStr}/${('0' + m).slice(-2)}/01`;
                    const endD = `${yearStr}/${('0' + m).slice(-2)}/31`;
                    catRow.push(`=SUMIFS('${yearStr}'!G:G, '${yearStr}'!H:H, "経費", '${yearStr}'!D:D, "${cat}", '${yearStr}'!B:B, ">=${startD}", '${yearStr}'!B:B, "<=${endD}")`);
                }
                rowsToInsert.push(catRow);
            }

            // endIdx は現在のブロックの最終行 (0-indexed)
            let insertAtRow = endIdx + 1;
            if (data[endIdx] && data[endIdx][0].toString() !== "") {
                insertAtRow = endIdx + 2;
            }

            // 行の挿入とデータ書き込み
            if (insertAtRow > sheet.getLastRow()) {
                const startR = sheet.getLastRow() + 1;
                sheet.getRange(startR, 1, rowsToInsert.length, 15).setValues(rowsToInsert);
                sheet.getRange(startR, 3, rowsToInsert.length, 13).setNumberFormat("#,##0");
            } else {
                sheet.insertRowsBefore(insertAtRow, rowsToInsert.length);
                sheet.getRange(insertAtRow, 1, rowsToInsert.length, 15).setValues(rowsToInsert);
                sheet.getRange(insertAtRow, 3, rowsToInsert.length, 13).setNumberFormat("#,##0");
            }
        }
    }
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

/**
 * サマリーシートを作成・更新するための関数（GASエディタからの手動実行用、またはイベント発火時）
 */
function createSummarySheetManually() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // main.jsで定義されているものを利用
    const settings = ensureSettingsSheet(ss);

    // 現在の年を取得して、そのシートが存在するか確認する
    const currentYearStr = new Date().getFullYear().toString();

    // サマリーシートを作成または更新
    ensureSummarySheet(ss, currentYearStr, settings);
    console.log(`${currentYearStr}年のサマリーシート作成プロセスを実行しました。`);
}

