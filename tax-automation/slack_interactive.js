/**
 * 解析完了後にSlackにBlock Kitを使ってインタラクティブな通知（承認ボタン等）を送信する
 */
function sendSlackInteractiveMessage(channelId, extractedData, recordResult, settingsMap) {
    const url = "https://slack.com/api/chat.postMessage";

    // 計算結果
    const ratio = settingsMap[extractedData.category] !== undefined ? settingsMap[extractedData.category] : 1.0;
    const finalAmount = Math.floor(extractedData.totalAmount * ratio);

    // コールバックに渡すための識別情報（スプレッドシートのどこを更新すべきか）
    // 数式やソート等で行が変動するため、証憑URLを使って行を特定する
    const callbackContext = {
        sheetName: recordResult.sheetName,
        fileUrl: recordResult.fileUrl || extractedData.fileUrl || ""
    };

    let duplicateWarningText = "";
    if (recordResult.isDuplicate) {
        duplicateWarningText = ":warning: *【注意】類似するデータが既に登録されている可能性があります。*\n";
    }

    const blocks = [
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": `${duplicateWarningText}レシートの解析が完了し、スプレッドシートの仮登録を行いました。\n内容を確認し、経費計上するか判断してください。`
            }
        },
        {
            "type": "section",
            "fields": [
                {
                    "type": "mrkdwn",
                    "text": `*日付:*\n${extractedData.date}`
                },
                {
                    "type": "mrkdwn",
                    "text": `*支払先:*\n${extractedData.storeName}`
                },
                {
                    "type": "mrkdwn",
                    "text": `*名目（推論）:*\n${extractedData.category}`
                },
                {
                    "type": "mrkdwn",
                    "text": `*総額:*\n¥${extractedData.totalAmount.toLocaleString()}`
                },
                {
                    "type": "mrkdwn",
                    "text": `*デフォルト按分率:*\n${ratio * 100}%`
                },
                {
                    "type": "mrkdwn",
                    "text": `*経費計上額（予定）:*\n¥${finalAmount.toLocaleString()}`
                }
            ]
        },
        {
            "type": "actions",
            "block_id": "receipt_approval_actions",
            "elements": [
                {
                    "type": "button",
                    "text": {
                        "type": "plain_text",
                        "emoji": true,
                        "text": "経費として計上 (Approve)"
                    },
                    "style": "primary",
                    "value": JSON.stringify({ action: "approve", ...callbackContext }),
                    "action_id": "btn_approve"
                },
                {
                    "type": "button",
                    "text": {
                        "type": "plain_text",
                        "emoji": true,
                        "text": "対象外・却下 (Reject)"
                    },
                    "style": "danger",
                    "value": JSON.stringify({ action: "reject", ...callbackContext }),
                    "action_id": "btn_reject"
                }
            ]
        }
    ];

    const payload = {
        channel: channelId,
        text: "レシート解析完了（確認をお願いします）",
        blocks: blocks
    };

    const options = {
        method: "post",
        contentType: "application/json",
        headers: { "Authorization": `Bearer ${SLACK_BOT_TOKEN}` },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const res = UrlFetchApp.fetch(url, options);
        const result = JSON.parse(res.getContentText());
        if (!result.ok) {
            console.error("Slack postMessage error: " + res.getContentText());
            logToErrorSheet("Slack postMessage Error", res.getContentText(), JSON.stringify(payload));
        }
    } catch (e) {
        console.error("sendSlackInteractiveMessage Error: " + e);
    }
}

/**
 * Slack Interactive (ボタンのクリックなど) のペイロードを処理する
 */
function handleSlackInteractivePayload(payload) {
    // 返信先のURL (Slack側メッセージを更新するため)
    const responseUrl = payload.response_url;
    const user = payload.user.name;

    if (payload.actions && payload.actions.length > 0) {
        const action = payload.actions[0];
        const valueData = JSON.parse(action.value);

        // スプレッドシート側のフラグを更新
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(valueData.sheetName);

        if (sheet) {
            // URLを使用して対象行を検索する (ソート等で行番号が変わる可能性があるため)
            const lastRow = sheet.getLastRow();
            let targetRow = -1;

            if (lastRow >= 2) {
                const urlData = sheet.getRange(2, 9, lastRow - 1, 1).getValues(); // I列(9列目)が証憑URL
                for (let i = 0; i < urlData.length; i++) {
                    if (urlData[i][0] === valueData.fileUrl) {
                        targetRow = i + 2; // 1行目がヘッダー、iが0始まりのため+2
                        break;
                    }
                }
            }

            if (targetRow === -1) {
                UrlFetchApp.fetch(responseUrl, {
                    method: "post",
                    contentType: "application/json",
                    payload: JSON.stringify({
                        replace_original: false,
                        text: ":warning: 対象となるスプレッドシートの行が見つからなかったため、更新に失敗しました（URL不一致または削除済み）。"
                    })
                });
                return ContentService.createTextOutput("OK");
            }

            const flagRange = sheet.getRange(targetRow, 8); // H列（経費フラグ）

            let newMessageText = "";

            if (valueData.action === "approve") {
                // Approveの場合、すでにデフォルトで「経費」なので特に変えなくてもよいが、念のためチェック
                flagRange.setValue("経費");
                newMessageText = `:white_check_mark: *@[${user}]* によって「経費として計上 (Approve)」されました。`;
            } else if (valueData.action === "reject") {
                // Rejectの場合、フラグを「-」にする
                flagRange.setValue("-");
                newMessageText = `:x: *@[${user}]* によって「対象外 (Reject)」とされました。`;
            }

            // 元のメッセージブロックを「処理済み」に更新する
            const updatedBlocks = [
                ...payload.message.blocks.slice(0, 2), // 元のテキストと詳細フィールドは残す
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": newMessageText
                    }
                }
            ];

            const updatePayload = {
                replace_original: true,
                text: newMessageText,
                blocks: updatedBlocks
            };

            UrlFetchApp.fetch(responseUrl, {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify(updatePayload)
            });

        } else {
            // 対象シートが見つからなかった場合のエラー処理
            UrlFetchApp.fetch(responseUrl, {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify({
                    replace_original: false,
                    text: ":warning: 対象となるスプレッドシートの行が見つからなかったため、更新に失敗しました。"
                })
            });
        }
    }

    return ContentService.createTextOutput("OK");
}
