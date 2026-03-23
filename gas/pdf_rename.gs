/**
 * PDF リネームスクリプト (Google Apps Script + Gemini API)
 *
 * 機能:
 *   1. 指定した Google Drive フォルダ内の PDF を取得
 *   2. Gemini API でスキャン PDF の内容を読み取り
 *   3. Gemini がファイル名を自動生成
 *   4. Google Drive 上でリネーム
 *   5. リネーム記録を Google スプレッドシートに保存
 *
 * セットアップ:
 *   1. https://script.google.com で新規プロジェクトを作成
 *   2. このコードを貼り付け
 *   3. CONFIG の GEMINI_API_KEY を設定
 *      （https://aistudio.google.com/apikey で取得）
 *   4. FOLDER_ID を設定
 *   5. main() を実行（初回はアクセス許可を求められます）
 */

// ============================================================
// CONFIG - ここを環境に合わせて変更してください
// ============================================================
const CONFIG = {
  // Gemini API キー（Google AI Studio で取得）
  // https://aistudio.google.com/apikey
  GEMINI_API_KEY: "AIzaSyCqjGjFCsbU1I-DFJSoKwcgKwRRiTxaqtM",

  // Gemini モデル
  GEMINI_MODEL: "gemini-2.5-flash",

  // PDF が格納されている Google Drive フォルダの ID
  FOLDER_ID: "1nX01Z1GH6TzqzlG3pEvVUi5tUimQy4QC",

  // ログを書き込むスプレッドシートの ID（空の場合は自動作成）
  SPREADSHEET_ID: "",

  // ログシート名
  SHEET_NAME: "リネーム記録",

  // ファイル名の最大文字数
  MAX_FILENAME_LENGTH: 80,

  // ドライラン（true にするとリネームせずにログだけ出力）
  DRY_RUN: false,
};

// Gemini API エンドポイント
const GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/";

// ============================================================
// メイン処理
// ============================================================

/**
 * エントリーポイント: PDF のリネームを実行する
 */
function main() {
  if (!CONFIG.GEMINI_API_KEY) {
    Logger.log("エラー: GEMINI_API_KEY が設定されていません。");
    Logger.log("https://aistudio.google.com/apikey で API キーを取得し、CONFIG に設定してください。");
    return;
  }

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
        geminiSummary: result.geminiSummary,
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
        geminiSummary: e.message,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
      });
    }

    // Gemini API のレートリミット対策
    Utilities.sleep(2000);
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
// PDF 処理（Gemini API）
// ============================================================

/**
 * 1つの PDF ファイルを処理する
 */
function processFile_(file) {
  const oldName = file.getName();

  // Gemini API で PDF を読み取り、ファイル名を生成
  const geminiResult = analyzeWithGemini_(file);

  if (!geminiResult || !geminiResult.fileName) {
    return { skipped: true, oldName: oldName, reason: "Gemini がファイル名を生成できませんでした" };
  }

  const newBaseName = sanitizeFileName_(geminiResult.fileName);

  if (!newBaseName) {
    return { skipped: true, oldName: oldName, reason: "有効なファイル名を生成できませんでした" };
  }

  // 長さ制限
  let finalBaseName = newBaseName;
  if (finalBaseName.length > CONFIG.MAX_FILENAME_LENGTH) {
    finalBaseName = finalBaseName.substring(0, CONFIG.MAX_FILENAME_LENGTH);
  }

  const newName = finalBaseName + ".pdf";

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
    geminiSummary: geminiResult.summary || "",
  };
}

/**
 * Gemini API で PDF の内容を解析し、ファイル名を生成する
 */
function analyzeWithGemini_(file) {
  const blob = file.getBlob();
  const base64Data = Utilities.base64Encode(blob.getBytes());

  // 元ファイル名から日付を抽出（フォールバック用）
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
  const origName = file.getName();
  const dateMatch = origName.match(/(\d{8})/);
  const fallbackDate = dateMatch ? dateMatch[1] : today;

  const prompt = `あなたはスキャンされたPDF文書を解析するアシスタントです。
このPDFの内容を読み取り、以下のJSON形式で回答してください。
他のテキストは一切出力せず、JSONのみ返してください。

{
  "fileName": "YYYYMMDD_文書種類_概要",
  "summary": "文書の要約（1〜2文）"
}

ファイル名のルール:
- 必ず先頭に日付を YYYYMMDD 形式で付けること
- 文書内に日付があればそれを使う（令和・平成の和暦は西暦に変換）
- 文書内に日付が見つからない場合は「${fallbackDate}」を使う
- 文書種類（請求書、見積書、契約書、領収書、報告書、議事録、通知書、申請書など）を付ける
- 概要は簡潔に（会社名や件名など主要な情報）
- ファイル名に使えない文字（/ \\ : * ? " < > |）は使わない
- スペースはアンダースコア(_)にする
- 最大80文字以内`;

  const url = GEMINI_URL + CONFIG.GEMINI_MODEL + ":generateContent?key=" + CONFIG.GEMINI_API_KEY;

  const payload = {
    contents: [
      {
        parts: [
          {
            inlineData: {
              mimeType: "application/pdf",
              data: base64Data,
            },
          },
          {
            text: prompt,
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 256,
    },
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  if (statusCode !== 200) {
    throw new Error(`Gemini API エラー (${statusCode}): ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());

  // レスポンスからテキストを取得
  const text = result.candidates[0].content.parts[0].text;
  Logger.log(`Gemini 応答: ${text}`);

  // JSON を抽出してパース
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error("Gemini の応答から JSON を抽出できませんでした: " + text);
  }

  return JSON.parse(jsonMatch[0]);
}

// ============================================================
// ユーティリティ
// ============================================================

/**
 * ファイル名に使えない文字を除去・置換する
 */
function sanitizeFileName_(name) {
  if (!name) return null;

  let sanitized = name
    .replace(/\.pdf$/i, "")
    .replace(/[\/\\:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");

  return sanitized || null;
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

    CONFIG.SPREADSHEET_ID = ss.getId();
    Logger.log(
      `スプレッドシートを作成しました: https://docs.google.com/spreadsheets/d/${ss.getId()}`
    );
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet
      .getRange(1, 1, 1, 6)
      .setValues([
        ["処理日時", "旧ファイル名", "新ファイル名", "Gemini要約", "ファイルID", "ファイルURL"],
      ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 400);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 300);
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
    record.geminiSummary || "",
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
// ユーティリティ関数
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
