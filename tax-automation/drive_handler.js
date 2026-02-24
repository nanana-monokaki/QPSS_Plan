/**
 * 定数定義 - Slackから画像をダウンロードするための準備や、ルートフォルダなどを設定
 */
const DRIVE_ROOT_FOLDER_ID = '1S-kfpYK6oWyvh2qzLmO2oJlIfqbW1c61'; // 「確定申告レシート」フォルダ

/**
 * 送信された画像ファイルをDriveに一時保存する
 * 
 * 1. Slackから画像を取得
 * 2. 所定の親フォルダ（「確定申告レシート」）に一時的な名前で保存
 */
function saveFileToDrive(fileEventAuth, slackToken) {
    // 1. ファイルのダウンロード
    const fileUrl = fileEventAuth.url_private_download;
    const downloadOptions = {
        headers: { 'Authorization': 'Bearer ' + slackToken }
    };

    const response = UrlFetchApp.fetch(fileUrl, downloadOptions);
    const blob = response.getBlob();

    // 一時的な名前の生成
    const extension = getFileExtension(fileEventAuth.mimetype, fileEventAuth.name);
    const tempFileName = `temp_${new Date().getTime()}.${extension}`;

    blob.setName(tempFileName);

    // 2. 親フォルダ「11_確定申告」の取得
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);

    // 3. ファイルの一時保存
    const driveFile = rootFolder.createFile(blob);

    return driveFile;
}

/**
 * 抽出された日付情報をもとに、ファイルを正しい「年-月」フォルダに移動しリネームする
 */
function moveFileToYearFolder(driveFile, receiptDate, fileEventAuth) {
    // receiptDate is like "2026/02/10"
    const today = new Date();
    let year = today.getFullYear().toString();
    // '02'のように0埋めされているものを数値として扱い再文字列化することで「2」にする
    let month = (today.getMonth() + 1).toString();

    if (receiptDate) {
        const parts = receiptDate.split('/');
        if (parts.length >= 2 && parts[0].length === 4) {
            year = parts[0];
            // '02'のような文字列を数値化して再文字列化（例: 2）
            month = parseInt(parts[1], 10).toString();
        }
    }

    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);

    // 「2026-2」のような年-月フォルダの取得または作成
    const targetFolderName = `${year}-${month}`;
    let targetFolder;
    const folderIter = rootFolder.getFoldersByName(targetFolderName);
    if (folderIter.hasNext()) {
        targetFolder = folderIter.next();
    } else {
        targetFolder = rootFolder.createFolder(targetFolderName);
    }

    // ファイルの名前を正式なものに変更
    const extension = getFileExtension(fileEventAuth.mimetype, fileEventAuth.name);

    // date部分からYYYYMMDDを作成 (例 "2026/02/10" -> "20260210")
    let formattedDate = year;
    if (receiptDate && receiptDate.includes('/')) {
        formattedDate = receiptDate.replace(/\//g, '');
    }

    const randomStr = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    const newFileName = `receipt_${formattedDate}_${randomStr}.${extension}`;

    driveFile.setName(newFileName);

    // ファイルを指定の年-月フォルダに移動
    driveFile.moveTo(targetFolder);
}

/**
 * MimeTypeやファイル名から拡張子を簡易推定
 */
function getFileExtension(mimeType, fileName) {
    if (mimeType === 'image/jpeg') return 'jpg';
    if (mimeType === 'image/png') return 'png';
    if (mimeType === 'application/pdf') return 'pdf';

    // Fallback to filename
    if (fileName) {
        const extMatch = fileName.match(/\.([a-zA-Z0-9]+)$/);
        if (extMatch) return extMatch[1].toLowerCase();
    }

    return 'bin';
}

/**
 * 画像ファイルをGoogle Docsに一時変換してGoogleの無料OCRを利用する
 */
function extractTextWithOCR(fileId) {
    // 変換元のファイル（画像）を取得
    const sourceImage = DriveApp.getFileById(fileId);

    // Google Docsとして新しくファイルを作成する設定（OCRが自動でかかる）
    const resource = {
        title: sourceImage.getName() + '_ocr_temp',
        mimeType: 'application/vnd.google-apps.document' // Google Docs形式
    };

    // OCRを実行（Drive APIを使用）
    const file = Drive.Files.copy(resource, fileId, { ocr: true, ocrLanguage: 'ja' });

    // 抽出されたテキストをDocumentAppで取得
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();

    // 一時ファイルの削除
    DriveApp.getFileById(file.id).setTrashed(true);

    return text;
}
