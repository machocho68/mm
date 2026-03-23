/**
 * PDF リネームスクリプト (Google Apps Script)
 *
 * 機能:
 *   1. 指定した Google Drive フォルダ内の PDF を取得
 *   2. Google Drive の OCR 機能で PDF のテキストを抽出
 *   3. テキスト内容からファイル名を自動生成
 *   4. Google Drive 上でリネーム
 *   5. リネーム記録を Google スプレッドシートに保存
 *
 * セットアップ:
 *   1. https://script.google.com で新規プロジェクトを作成
 *   2. このコードを貼り付け
 *   3. CONFIG セクションの FOLDER_ID と SPREADSHEET_ID を設定
 *   4. main() を実行（初回はアクセス許可を求められます）
 */

// ============================================================
// CONFIG - ここを環境に合わせて変更してください
// ============================================================
const CONFIG = {
  // PDF が格納されている Google Drive フォルダの ID
  // URL: https://drive.google.com/drive/folders/XXXXX の XXXXX 部分
  FOLDER_ID: "1nX01Z1GH6TzqzlG3pEvVUi5tUimQy4QC",

  // ログを書き込むスプレッドシートの ID（空の場合は自動作成）
  // URL: https://docs.google.com/spreadsheets/d/XXXXX の XXXXX 部分
  SPREADSHEET_ID: "",

  // ログシート名
  SHEET_NAME: "リネーム記録",

  // ファイル名の最大文字数
  MAX_FILENAME_LENGTH: 80,

  // ドライラン（true にするとリネームせずにログだけ出力）
  DRY_RUN: false,

  // OCR で抽出するテキストの最大文字数（処理対象）
  MAX_OCR_TEXT_LENGTH: 3000,
};

// ============================================================
// メイン処理
// ============================================================

/**
 * エントリーポイント: PDF のリネームを実行する
 */
function main() {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const sheet = getOrCreateLogSheet_();
  const pdfFiles = getPdfFiles_(folder);

  Logger.log(`対象フォルダ: ${folder.getName()}`);
  Logger.log(`PDF ファイル数: ${pdfFiles.length}`);

  if (pdfFiles.length === 0) {
    Logger.log("リネーム対象の PDF ファイルが見つかりません。");
    return;
  }

  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  for (const file of pdfFiles) {
    try {
      const result = processFile_(file);

      if (result.skipped) {
        skippedCount++;
        Logger.log(`スキップ: ${result.oldName} - ${result.reason}`);
        continue;
      }

      // スプレッドシートに記録
      appendLog_(sheet, {
        timestamp: new Date(),
        oldName: result.oldName,
        newName: result.newName,
        extractedText: result.extractedText,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
      });

      processedCount++;
      Logger.log(`リネーム完了: ${result.oldName} → ${result.newName}`);
    } catch (e) {
      errorCount++;
      Logger.log(`エラー: ${file.getName()} - ${e.message}`);
      appendLog_(sheet, {
        timestamp: new Date(),
        oldName: file.getName(),
        newName: "ERROR",
        extractedText: e.message,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
      });
    }
  }

  Logger.log(`\n=== 処理結果 ===`);
  Logger.log(`処理成功: ${processedCount} 件`);
  Logger.log(`スキップ: ${skippedCount} 件`);
  Logger.log(`エラー: ${errorCount} 件`);

  if (CONFIG.SPREADSHEET_ID) {
    Logger.log(
      `記録シート: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`
    );
  }
}

// ============================================================
// PDF 処理
// ============================================================

/**
 * 1つの PDF ファイルを処理する
 */
function processFile_(file) {
  const oldName = file.getName();

  // テキスト抽出（Google Drive OCR を利用）
  const extractedText = extractTextFromPdf_(file);

  if (!extractedText || extractedText.trim().length === 0) {
    return { skipped: true, oldName: oldName, reason: "テキストを抽出できませんでした" };
  }

  // テキストからファイル名を生成
  const newBaseName = generateFileName_(extractedText, oldName);

  if (!newBaseName) {
    return { skipped: true, oldName: oldName, reason: "ファイル名を生成できませんでした" };
  }

  const newName = newBaseName + ".pdf";

  // 同名チェック
  if (oldName === newName) {
    return { skipped: true, oldName: oldName, reason: "ファイル名が同一です" };
  }

  // リネーム実行
  if (!CONFIG.DRY_RUN) {
    file.setName(newName);
  } else {
    Logger.log(`[DRY RUN] ${oldName} → ${newName}`);
  }

  return {
    skipped: false,
    oldName: oldName,
    newName: newName,
    extractedText: extractedText.substring(0, 200),
  };
}

/**
 * Google Drive の OCR 機能を使って PDF からテキストを抽出する
 *
 * PDFをGoogle Docsに一時変換してテキストを取得し、変換したDocsは削除する
 * Drive REST API v3 を直接呼び出す方式（サービス追加のバージョンに依存しない）
 */
function extractTextFromPdf_(file) {
  let tempDocId = null;
  try {
    const blob = file.getBlob();
    const fileName = file.getName().replace(/\.pdf$/i, "") + "_ocr_temp";

    // Drive API v3 REST エンドポイントで PDF → Google Docs に変換（OCR適用）
    const metadata = {
      name: fileName,
      mimeType: "application/vnd.google-apps.document",
    };

    const boundary = "===BOUNDARY===";
    const requestBody =
      "--" + boundary + "\r\n" +
      "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
      JSON.stringify(metadata) + "\r\n" +
      "--" + boundary + "\r\n" +
      "Content-Type: application/pdf\r\n" +
      "Content-Transfer-Encoding: base64\r\n\r\n" +
      Utilities.base64Encode(blob.getBytes()) + "\r\n" +
      "--" + boundary + "--";

    const response = UrlFetchApp.fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=ja",
      {
        method: "post",
        contentType: "multipart/related; boundary=" + boundary,
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
        payload: requestBody,
        muteHttpExceptions: true,
      }
    );

    if (response.getResponseCode() !== 200) {
      throw new Error("Drive API エラー: " + response.getContentText());
    }

    const result = JSON.parse(response.getContentText());
    tempDocId = result.id;

    // Google Docs からテキストを取得
    const doc = DocumentApp.openById(tempDocId);
    let text = doc.getBody().getText();

    // テキストの長さを制限
    if (text.length > CONFIG.MAX_OCR_TEXT_LENGTH) {
      text = text.substring(0, CONFIG.MAX_OCR_TEXT_LENGTH);
    }

    return text;
  } catch (e) {
    Logger.log(`OCR エラー (${file.getName()}): ${e.message}`);
    return null;
  } finally {
    // 一時ファイルを削除
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (e) {
        Logger.log(`一時ファイル削除エラー: ${e.message}`);
      }
    }
  }
}

/**
 * 抽出されたテキストからファイル名を生成する
 *
 * ルール:
 *   1. 日付が含まれていれば先頭に付与
 *   2. タイトル・件名・文書種類を検出
 *   3. 不要な文字を除去してファイル名に適した形に整形
 */
function generateFileName_(text, originalName) {
  const lines = text.split("\n").filter((line) => line.trim().length > 0);

  if (lines.length === 0) return null;

  // 日付の検出
  const dateStr = extractDate_(text);

  // 文書種類の検出
  const docType = detectDocumentType_(text);

  // タイトル候補の抽出（最初の有意な行を使用）
  let title = extractTitle_(lines);

  // ファイル名の組み立て
  let parts = [];
  if (dateStr) parts.push(dateStr);
  if (docType) parts.push(docType);
  if (title) parts.push(title);

  if (parts.length === 0) return null;

  let fileName = parts.join("_");

  // ファイル名に使えない文字を除去
  fileName = sanitizeFileName_(fileName);

  // 長さ制限
  if (fileName.length > CONFIG.MAX_FILENAME_LENGTH) {
    fileName = fileName.substring(0, CONFIG.MAX_FILENAME_LENGTH);
  }

  return fileName;
}

/**
 * テキストから日付を抽出する
 */
function extractDate_(text) {
  // 様々な日付フォーマットに対応
  const patterns = [
    // 2024年1月15日 / 2024年01月15日
    /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/,
    // 令和6年1月15日
    /令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/,
    // 平成31年1月15日
    /平成\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/,
    // R6.1.15 / R06.01.15
    /R\s*(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/,
    // 2024/01/15 / 2024-01-15
    /(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      if (pattern.source.includes("令和")) {
        const year = 2018 + parseInt(match[1]);
        const month = String(match[2]).padStart(2, "0");
        const day = String(match[3]).padStart(2, "0");
        return `${year}${month}${day}`;
      } else if (pattern.source.includes("平成")) {
        const year = 1988 + parseInt(match[1]);
        const month = String(match[2]).padStart(2, "0");
        const day = String(match[3]).padStart(2, "0");
        return `${year}${month}${day}`;
      } else if (pattern.source.startsWith("R")) {
        const year = 2018 + parseInt(match[1]);
        const month = String(match[2]).padStart(2, "0");
        const day = String(match[3]).padStart(2, "0");
        return `${year}${month}${day}`;
      } else {
        const year = match[1];
        const month = String(match[2]).padStart(2, "0");
        const day = String(match[3]).padStart(2, "0");
        return `${year}${month}${day}`;
      }
    }
  }

  return null;
}

/**
 * テキストから文書種類を検出する
 */
function detectDocumentType_(text) {
  const docTypes = [
    { keywords: ["請求書"], type: "請求書" },
    { keywords: ["見積書", "見積り", "お見積"], type: "見積書" },
    { keywords: ["納品書"], type: "納品書" },
    { keywords: ["領収書", "領収証"], type: "領収書" },
    { keywords: ["契約書"], type: "契約書" },
    { keywords: ["注文書", "発注書"], type: "注文書" },
    { keywords: ["報告書"], type: "報告書" },
    { keywords: ["議事録"], type: "議事録" },
    { keywords: ["通知書", "お知らせ"], type: "通知書" },
    { keywords: ["申請書"], type: "申請書" },
    { keywords: ["届出", "届け出"], type: "届出" },
    { keywords: ["明細書", "明細"], type: "明細書" },
    { keywords: ["証明書"], type: "証明書" },
    { keywords: ["仕様書"], type: "仕様書" },
    { keywords: ["稟議書"], type: "稟議書" },
  ];

  for (const dt of docTypes) {
    for (const kw of dt.keywords) {
      if (text.includes(kw)) {
        return dt.type;
      }
    }
  }

  return null;
}

/**
 * テキストの先頭行からタイトルを抽出する
 */
function extractTitle_(lines) {
  // 短すぎる行や数字だけの行をスキップして、最初の有意な行を取得
  for (const line of lines) {
    const trimmed = line.trim();

    // 空行、短すぎる行、数字・記号だけの行をスキップ
    if (trimmed.length < 2) continue;
    if (/^[\d\s\-\/\.,:;]+$/.test(trimmed)) continue;
    if (/^(page|ページ|\d+\/\d+)\s*$/i.test(trimmed)) continue;

    // タイトルとして使える行を返す
    let title = trimmed;

    // 長すぎる場合は切り詰め
    if (title.length > 40) {
      title = title.substring(0, 40);
    }

    return title;
  }

  return null;
}

/**
 * ファイル名に使えない文字を除去・置換する
 */
function sanitizeFileName_(name) {
  // ファイル名に使用不可の文字を置換
  let sanitized = name
    .replace(/[\/\\:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");

  return sanitized;
}

// ============================================================
// スプレッドシート操作
// ============================================================

/**
 * ログ用のスプレッドシートとシートを取得または作成する
 */
function getOrCreateLogSheet_() {
  let ss;

  if (CONFIG.SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } else {
    // スプレッドシートを新規作成し、PDF フォルダと同じ場所に配置
    ss = SpreadsheetApp.create("PDF リネーム記録");
    const file = DriveApp.getFileById(ss.getId());
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    // CONFIG を更新（ログ出力用）
    CONFIG.SPREADSHEET_ID = ss.getId();
    Logger.log(
      `スプレッドシートを作成しました: https://docs.google.com/spreadsheets/d/${ss.getId()}`
    );
  }

  // シートの取得または作成
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // ヘッダー行を追加
    sheet
      .getRange(1, 1, 1, 6)
      .setValues([
        ["処理日時", "旧ファイル名", "新ファイル名", "抽出テキスト（先頭200文字）", "ファイルID", "ファイルURL"],
      ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);

    // 列幅を調整
    sheet.setColumnWidth(1, 160); // 処理日時
    sheet.setColumnWidth(2, 250); // 旧ファイル名
    sheet.setColumnWidth(3, 250); // 新ファイル名
    sheet.setColumnWidth(4, 400); // 抽出テキスト
    sheet.setColumnWidth(5, 200); // ファイルID
    sheet.setColumnWidth(6, 300); // ファイルURL
  }

  return sheet;
}

/**
 * スプレッドシートにリネーム記録を追加する
 */
function appendLog_(sheet, record) {
  sheet.appendRow([
    record.timestamp,
    record.oldName,
    record.newName,
    record.extractedText || "",
    record.fileId,
    record.fileUrl,
  ]);
}

// ============================================================
// フォルダ操作
// ============================================================

/**
 * フォルダ内の PDF ファイル一覧を取得する
 */
function getPdfFiles_(folder) {
  const files = [];
  const iterator = folder.getFilesByType("application/pdf");

  while (iterator.hasNext()) {
    files.push(iterator.next());
  }

  return files;
}

// ============================================================
// ユーティリティ
// ============================================================

/**
 * ドライラン: リネームせずに結果だけ確認する
 */
function dryRun() {
  CONFIG.DRY_RUN = true;
  main();
}

/**
 * スプレッドシートの URL を表示する
 */
function showSpreadsheetUrl() {
  if (CONFIG.SPREADSHEET_ID) {
    Logger.log(
      `スプレッドシート: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`
    );
  } else {
    Logger.log("スプレッドシートはまだ作成されていません。main() を実行してください。");
  }
}
