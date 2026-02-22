/**
 * 定数定義 - Slackから画像をダウンロードするための準備や、ルートフォルダなどを設定
 */
const DRIVE_ROOT_FOLDER_ID = '1pD7JPAz-Tm-N0s6KibyaCwnRQ4TXaG-J';
const DRIVE_YEARLY_FOLDER_PREFIX = '年度_確定申告レシート';

/**
 * 送信された画像ファイルをDriveに保存する
 * 
 * 1. Slackから画像を取得
 * 2. 所定のフォルダ構成（確定申告/2026年度_確定申告レシート/2026-02等）を作成/取得
 * 3. ファイルをリネーム（receipt_YYYYMMDD_HHmm.jpg等）して保存
 */
function saveFileToDrive(fileEventAuth, slackToken) {
    // 1. ファイルのダウンロード
    const fileUrl = fileEventAuth.url_private_download;
    const downloadOptions = {
        headers: { 'Authorization': 'Bearer ' + slackToken }
    };

    const response = UrlFetchApp.fetch(fileUrl, downloadOptions);
    const blob = response.getBlob();

    // 名前の生成
    const today = new Date();
    const year = today.getFullYear();
    const month = ('0' + (today.getMonth() + 1)).slice(-2);
    const date = ('0' + today.getDate()).slice(-2);
    const hours = ('0' + today.getHours()).slice(-2);
    const minutes = ('0' + today.getMinutes()).slice(-2);

    const formattedDate = `${year}${month}${date}_${hours}${minutes}`;
    const extension = getFileExtension(fileEventAuth.mimetype);
    const newFileName = `receipt_${formattedDate}.${extension}`;

    blob.setName(newFileName);

    // 2. フォルダの取得・作成
    const folder = getTargetFolder(year, month);

    // 3. ファイルの保存
    const driveFile = folder.createFile(blob);

    return driveFile;
}

/**
 * フォルダ構成の取得／作成
 */
function getTargetFolder(year, month) {
    // 1. 親フォルダ「11_確定申告」の取得
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);

    // 2. 年度別サブフォルダ「YYYY年度_確定申告レシート」の取得
    const yearlyFolderName = `${year}${DRIVE_YEARLY_FOLDER_PREFIX}`;
    let yearlyFolder;
    const yearlyIter = rootFolder.getFoldersByName(yearlyFolderName);
    if (yearlyIter.hasNext()) {
        yearlyFolder = yearlyIter.next();
    } else {
        yearlyFolder = rootFolder.createFolder(yearlyFolderName);
    }

    // 3. 月別フォルダ「YYYY-MM」の取得
    const monthFolderName = `${year}-${month}`;
    let monthFolder;
    const monthIter = yearlyFolder.getFoldersByName(monthFolderName);
    if (monthIter.hasNext()) {
        monthFolder = monthIter.next();
    } else {
        monthFolder = yearlyFolder.createFolder(monthFolderName);
    }

    return monthFolder;
}

/**
 * MimeTypeから拡張子を簡易推定
 */
function getFileExtension(mimeType) {
    if (mimeType === 'image/jpeg') return 'jpg';
    if (mimeType === 'image/png') return 'png';
    if (mimeType === 'application/pdf') return 'pdf';
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
