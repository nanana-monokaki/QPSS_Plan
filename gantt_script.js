/**
 * 聞き耳アワー ガントチャート・決定事項管理ツール
 */

function onOpen(e) {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('ガントツール')
        .addItem('1. ガントチャート初期構築(リセット)', 'setupGanttChart')
        .addItem('2. 決定事項一覧シート初期構築', 'setupListSheet')
        .addItem('3. 企画書ドキュメント出力', 'exportToDocument')
        .addToUi();
}

/**
 * 1. ガントチャートのフォーマット構築
 */
function setupGanttChart() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ganttSheetName = "ガントチャート";
    let ganttSheet = ss.getSheetByName(ganttSheetName);

    if (!ganttSheet) {
        ganttSheet = ss.insertSheet(ganttSheetName);
    } else {
        const ui = SpreadsheetApp.getUi();
        const res = ui.alert('確認', '【重要】ガントチャートを【完全にリセット】して再構築しますか？\n（※既存のデータは全て消去されます）', ui.ButtonSet.YES_NO);
        if (res !== ui.Button.YES) return;
        ganttSheet.clear();
    }

    // 行列数をある程度確保
    if (ganttSheet.getMaxRows() < 100) {
        ganttSheet.insertRowsAfter(ganttSheet.getMaxRows(), 100 - ganttSheet.getMaxRows());
    }

    // === 列構成の変更（決定/未定をB列へ）===
    // 1:ステータス(A), 2:決定/未定(B), 3:セクション(C), 4:項目(D), 5:内容(E), 
    // 6:担当者(F), 7:開始日(G), 8:終了日(H), 9:所要日数(I), 10:備考(J), 
    // 11:TaskID(K・非表示), 12:セパレータ(L), 13〜:カレンダー(M〜)
    const headers = ["ステータス", "決定/未定", "セクション", "項目", "内容", "担当者", "開始日", "終了日", "所要日数", "備考", "TaskID", " "];
    ganttSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    ganttSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
    ganttSheet.setFrozenRows(1);
    ganttSheet.setFrozenColumns(5); // E列(内容)まで固定

    // 日付ヘッダーの生成 (M列 = 13列目から)
    const numDays = 150; // 5ヶ月分
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const dateValues = [];
    const startCol = 13; // カレンダー開始位置（M列）
    for (let i = 0; i < numDays; i++) {
        const d = new Date(today.getTime());
        d.setDate(today.getDate() + i);
        dateValues.push(d);
    }

    const targetMaxCol = startCol + numDays - 1;
    if (ganttSheet.getMaxColumns() < targetMaxCol) {
        ganttSheet.insertColumnsAfter(ganttSheet.getMaxColumns(), targetMaxCol - ganttSheet.getMaxColumns());
    }

    ganttSheet.getRange(1, startCol, 1, numDays).setValues([dateValues]);
    ganttSheet.getRange(1, startCol, 1, numDays).setNumberFormat("m/d");
    ganttSheet.getRange(1, startCol, 1, numDays).setFontWeight("bold").setBackground("#cfe2f3");

    // 列幅の調整
    ganttSheet.setColumnWidth(1, 100); // A:ステータス
    ganttSheet.setColumnWidth(2, 100); // B:決定/未定 (変更)
    ganttSheet.setColumnWidth(3, 100); // C:セクション (変更)
    ganttSheet.setColumnWidth(4, 200); // D:項目 (変更)
    ganttSheet.setColumnWidth(5, 300); // E:内容 (変更)
    ganttSheet.setColumnWidth(6, 100); // F:担当者 (変更)
    ganttSheet.setColumnWidth(7, 100); // G:開始日
    ganttSheet.setColumnWidth(8, 100); // H:終了日
    ganttSheet.setColumnWidth(9, 80);  // I:所要日数
    ganttSheet.setColumnWidth(10, 200); // J:備考

    ganttSheet.hideColumns(11); // K:TaskID (非表示)
    ganttSheet.setColumnWidth(12, 20); // L:セパレータ

    for (let i = 0; i < numDays; i++) {
        ganttSheet.setColumnWidth(startCol + i, 30); // M列〜カレンダー幅
    }

    // ==========================================
    // 入力規則・数式の整理と正確な範囲指定
    // ==========================================
    const maxRow = ganttSheet.getMaxRows();
    const numRowsToApply = maxRow - 1;

    ganttSheet.getRange(2, 1, numRowsToApply, ganttSheet.getMaxColumns()).clearDataValidations();

    // 入力規則（ステータス: A列）
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(["未着手", "進行中", "完了", "保留"]).setAllowInvalid(false).build();
    ganttSheet.getRange(2, 1, numRowsToApply, 1).setDataValidation(statusRule);

    // 入力規則（決定/未定: B列に変更）
    const flagRule = SpreadsheetApp.newDataValidation().requireValueInList(["未定", "決定"]).setAllowInvalid(false).build();
    ganttSheet.getRange(2, 2, numRowsToApply, 1).setDataValidation(flagRule);

    const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
    ganttSheet.getRange(2, 7, numRowsToApply, 2).setDataValidation(dateRule); // G, H列

    ganttSheet.getRange(2, 9, numRowsToApply, 1).setFormula('=IF(AND(G2<>"", H2<>""), H2-G2+1, "")');

    ganttSheet.clearConditionalFormatRules();
    const rules = [];

    // 1. 【行グレーアウト】B列(決定/未定)が「決定」のとき A〜J列を網羅
    const ruleItemGrayout = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B2="決定"') // B列に変更
        .setBackground("#efefef")
        .setFontColor("#999999")
        .setRanges([ganttSheet.getRange(2, 1, numRowsToApply, 10)]) // A〜J列
        .build();
    rules.push(ruleItemGrayout);

    const calRange = ganttSheet.getRange(2, startCol, numRowsToApply, numDays);

    // 2. 決定チャート (B列参照)
    const ruleDecided = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(M$1>=$G2, M$1<=$H2, $B2="決定")')
        .setBackground("#3c78d8")
        .setRanges([calRange])
        .build();
    rules.push(ruleDecided);

    // 3. 未定チャート (B列参照)
    const ruleTemp = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(M$1>=$G2, M$1<=$H2, $B2<>"決定")')
        .setBackground("#a4c2f4")
        .setRanges([calRange])
        .build();
    rules.push(ruleTemp);

    // 4. 今日ハイライト
    const ruleToday = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=M$1=TODAY()')
        .setBackground("#fff2cc")
        .setRanges([calRange])
        .build();
    rules.push(ruleToday);

    ganttSheet.setConditionalFormatRules(rules);
    SpreadsheetApp.getUi().alert("ガントチャートの設定が完了しました。");
}

/**
 * 2. 決定事項一覧シートの構築（企画書風）
 */
function setupListSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listSheetName = "決定事項一覧";
    let listSheet = ss.getSheetByName(listSheetName);

    if (!listSheet) {
        listSheet = ss.insertSheet(listSheetName);
    } else {
        const ui = SpreadsheetApp.getUi();
        const res = ui.alert('確認', '【重要】「決定事項一覧」シートをリセットして初期化しますか？\n（既存データは消えます）', ui.ButtonSet.YES_NO);
        if (res === ui.Button.YES) {
            listSheet.clear();
        } else {
            return;
        }
    }

    listSheet.setHiddenGridlines(true);

    // === 1行目：自由入力できるタイトルエリア ===
    listSheet.getRange("B1:C1").merge().setValue("〇〇企画概要書");
    listSheet.getRange("B1").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("left");

    // 少し隙間を空ける
    listSheet.setRowHeight(1, 40);
    listSheet.setRowHeight(2, 10);

    // === 3行目：ヘッダー ===
    // A:TaskID(非表示), B:セクション, C:項目(太字), D:内容
    const listHeaders = ["TaskID", "セクション", "項目", "内容"];
    listSheet.getRange(3, 1, 1, listHeaders.length).setValues([listHeaders]);

    const headerRange = listSheet.getRange(3, 2, 1, 3);
    headerRange.setFontWeight("bold")
        .setBackground("#4c1130") // 暗めのシックな赤紫系
        .setFontColor("#ffffff")
        .setFontSize(12)
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setBorder(true, true, true, true, true, true, "#333333", SpreadsheetApp.BorderStyle.SOLID);

    listSheet.setRowHeight(3, 30);

    listSheet.setColumnWidth(1, 0); // TaskID列は非表示
    listSheet.hideColumns(1);       // GASの機能でA列を隠す

    // ★念のためA列の文字色を完全な白にして背景に同化させる
    listSheet.getRange("A:A").setFontColor("#ffffff");

    listSheet.setColumnWidth(2, 150); // セクション(B列)
    listSheet.setColumnWidth(3, 250); // 項目(C列)
    listSheet.setColumnWidth(4, 500); // 内容(D列)

    // 項目列（C列）全体を太字に設定
    listSheet.getRange("C:C").setFontWeight("bold");

    // 文字を上揃え＆折り返し
    listSheet.getRange("B4:D").setWrap(true).setVerticalAlignment("top");

    // 同期は4行目以降（ヘッダー下）に行うため、ヘッダーで固定
    listSheet.setFrozenRows(3);
    SpreadsheetApp.getUi().alert("「決定事項一覧」シートを初期化しました。\nタイトル（B1セル）は自由に変更可能です。");
}

/**
 * 決定事項の順序をガントチャート順に並び替える（または全再構築する）共通処理
 */
function syncGanttToList() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ganttSheet = ss.getSheetByName("ガントチャート");
    const listSheet = ss.getSheetByName("決定事項一覧");

    if (!ganttSheet || !listSheet) return;

    // 1. ガントチャートから最新の「決定」済みの全データを取得
    const ganttData = ganttSheet.getDataRange().getValues();
    const decidedTasks = []; // ガント順に格納される

    // ヘッダー(1行目)をスキップ
    for (let i = 1; i < ganttData.length; i++) {
        const rowData = ganttData[i];
        const flag = rowData[1];    // B:決定/未定
        const section = rowData[2]; // C:セクション
        const item = rowData[3];    // D:項目
        const content = rowData[4]; // E:内容
        const taskId = rowData[10]; // K:TaskID

        let idStr = taskId;
        if (flag === "決定" && !taskId) {
            idStr = Utilities.getUuid();
            ganttSheet.getRange(i + 1, 11).setValue(idStr); // K列にセット
        }

        if (flag === "決定" && idStr) {
            decidedTasks.push({
                taskId: idStr,
                section: section,
                item: item,
                content: content
            });
        }
    }

    // 2. 決定事項一覧側のデータを取得 (4行目から)
    const listMaxRow = listSheet.getMaxRows();
    let listData = [];
    const manualRows = [];

    const lastDataRow = listSheet.getLastRow();
    if (lastDataRow >= 4) {
        listData = listSheet.getRange(4, 1, lastDataRow - 3, 4).getValues();
        // 手動追加分をキープ
        for (let j = 0; j < listData.length; j++) {
            if (listData[j][1] || listData[j][2] || listData[j][3]) {
                if (!listData[j][0]) {
                    manualRows.push(listData[j]);
                }
            }
        }
        // 古いデータを一旦消去 (4行目からシートの最後まで)
        listSheet.getRange(4, 1, listMaxRow - 3, 4).clearContent();
        listSheet.getRange(4, 1, listMaxRow - 3, 4).setBorder(false, false, false, false, false, false);
    }

    // 3. データ作り直し
    const finalRowsToWrite = [];
    for (let k = 0; k < decidedTasks.length; k++) {
        const t = decidedTasks[k];
        finalRowsToWrite.push([t.taskId, t.section, t.item, t.content]);
    }
    for (let m = 0; m < manualRows.length; m++) {
        finalRowsToWrite.push(manualRows[m]);
    }

    // 4. 書き込み、罫線を引く
    if (finalRowsToWrite.length > 0) {
        listSheet.getRange(4, 1, finalRowsToWrite.length, 4).setValues(finalRowsToWrite);

        for (let r = 0; r < finalRowsToWrite.length; r++) {
            listSheet.getRange(4 + r, 2, 1, 3).setBorder(false, false, true, false, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
        }
    }

    // ★同期のたびに、強制的にA列(TaskID)を非表示＆白文字にする
    listSheet.hideColumns(1);
    listSheet.getRange("A:A").setFontColor("#ffffff");
}

/**
 * onEditトリガー
 */
function onEdit(e) {
    if (!e || !e.range) return;
    const sheet = e.source.getActiveSheet();
    // ガントチャートが編集された時のみ、即座に同期（並び替え再構築）を実行
    if (sheet.getName() === "ガントチャート") {
        syncGanttToList();
    }
}

/**
 * 3. 企画書ドキュメント出力
 */
function exportToDocument() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listSheet = ss.getSheetByName("決定事項一覧");
    if (!listSheet) {
        SpreadsheetApp.getUi().alert("「決定事項一覧」シートが見つかりません。");
        return;
    }

    // タイトルB1の取得
    let docTitleBase = listSheet.getRange("B1").getValue() || "〇〇企画概要書";

    const listMaxRow = listSheet.getLastRow();
    if (listMaxRow < 4) {
        SpreadsheetApp.getUi().alert("出力する内容がありません。");
        return;
    }

    const data = listSheet.getRange(4, 1, listMaxRow - 3, 4).getDisplayValues();

    const docTitle = docTitleBase + "_" + Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
    const doc = DocumentApp.create(docTitle);
    const body = doc.getBody();

    body.insertParagraph(0, docTitleBase).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("出力日時: " + Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm"));
    body.appendParagraph("------------------------------------------------------------------");

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const section = row[1]; // B列相当
        const item = row[2];    // C列相当
        const content = row[3]; // D列相当

        // 内容が何もなければ弾く
        if (!item && !content) continue;

        // 見出し
        body.appendParagraph("【" + (section || "一般") + "】 " + (item || "")).setHeading(DocumentApp.ParagraphHeading.HEADING3);

        // 内容
        if (content) {
            body.appendParagraph(content);
        }

        // 隙間を開ける
        body.appendParagraph("");
    }

    doc.saveAndClose();

    // 指定のフォルダに格納 (190G8TOeMmgt8cEbVdVji0IiE0RYzArOD)
    const folderId = "190G8TOeMmgt8cEbVdVji0IiE0RYzArOD";
    try {
        const file = DriveApp.getFileById(doc.getId());
        const folder = DriveApp.getFolderById(folderId);
        file.moveTo(folder);
        SpreadsheetApp.getUi().alert("ドキュメントを作成しました！\n指定のご共有フォルダに保存されました。\n\nファイル名: " + docTitle);
    } catch (e) {
        SpreadsheetApp.getUi().alert("ドキュメントは作成されましたが、指定フォルダへの移動に失敗しました。\n\n作成先: マイドライブ直下");
    }
}
