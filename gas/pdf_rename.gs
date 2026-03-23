const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "",
  GEMINI_MODEL: "gemini-2.5-flash",
  FOLDER_ID: "1nX01Z1GH6TzqzlG3pEvVUi5tUimQy4QC",
  SPREADSHEET_ID: "",
  SHEET_NAME: "リネーム記録",
  MAX_FILENAME_LENGTH: 80,
};

const GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/";

/** エントリーポイント */
function main() {
  if (!CONFIG.GEMINI_API_KEY) {
    Logger.log("エラー: GEMINI_API_KEY が未設定です");
    return;
  }

  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const sheet = getOrCreateLogSheet_();
  const pdfFiles = getPdfFiles_(folder);

  Logger.log(`フォルダ: ${folder.getName()} / PDF: ${pdfFiles.length} 件`);
  if (pdfFiles.length === 0) return;

  let processed = 0, skipped = 0, errors = 0;

  for (const file of pdfFiles) {
    try {
      const result = processFile_(file);
      if (result.skipped) {
        skipped++;
        Logger.log(`スキップ: ${result.oldName} - ${result.reason}`);
        continue;
      }
      appendLog_(sheet, {
        timestamp: new Date(),
        oldName: result.oldName,
        newName: result.newName,
        summary: result.summary,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
      });
      processed++;
      Logger.log(`完了: ${result.oldName} → ${result.newName}`);
    } catch (e) {
      errors++;
      Logger.log(`エラー: ${file.getName()} - ${e.message}`);
      appendLog_(sheet, {
        timestamp: new Date(),
        oldName: file.getName(),
        newName: "ERROR",
        summary: e.message,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
      });
    }
    Utilities.sleep(2000);
  }

  Logger.log(`結果: 成功=${processed} スキップ=${skipped} エラー=${errors}`);
  if (CONFIG.SPREADSHEET_ID) {
    Logger.log(`シート: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`);
  }
}

/** PDF を解析してリネームする */
function processFile_(file) {
  const oldName = file.getName();
  const gemini = analyzeWithGemini_(file);

  if (!gemini || !gemini.fileName) {
    return { skipped: true, oldName, reason: "ファイル名を生成できませんでした" };
  }

  let baseName = sanitizeFileName_(gemini.fileName);
  if (!baseName) {
    return { skipped: true, oldName, reason: "有効なファイル名を生成できませんでした" };
  }
  if (baseName.length > CONFIG.MAX_FILENAME_LENGTH) {
    baseName = baseName.substring(0, CONFIG.MAX_FILENAME_LENGTH);
  }

  const newName = baseName + ".pdf";
  if (oldName === newName) {
    return { skipped: true, oldName, reason: "ファイル名が同一です" };
  }

  file.setName(newName);
  return { skipped: false, oldName, newName, summary: gemini.summary || "" };
}

/** Gemini API で PDF の内容を解析しファイル名を生成する */
function analyzeWithGemini_(file) {
  const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
  const dateMatch = file.getName().match(/(\d{8})/);
  const fallbackDate = dateMatch ? dateMatch[1] : today;

  const prompt = `あなたはスキャンされたPDF文書を解析するアシスタントです。
このPDFの内容を読み取り、以下のJSON形式で回答してください。
他のテキストは一切出力せず、JSONのみ返してください。

{
  "fileName": "YYYYMMDD_文書種類_概要",
  "summary": "文書の要約（1〜2文）"
}

ファイル名のルール:
- 先頭に日付をYYYYMMDD形式で付ける
- 文書内に日付があればそれを使う（和暦は西暦に変換）
- 日付が見つからない場合は「${fallbackDate}」を使う
- 文書種類（請求書、見積書、契約書、領収書、報告書、議事録、通知書、申請書など）を付ける
- 概要は簡潔に（会社名や件名など）
- 使えない文字（/ \\ : * ? " < > |）は使わない
- スペースはアンダースコア(_)にする
- 最大80文字以内`;

  const url = GEMINI_URL + CONFIG.GEMINI_MODEL + ":generateContent?key=" + CONFIG.GEMINI_API_KEY;
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [
        { inlineData: { mimeType: "application/pdf", data: base64Data } },
        { text: prompt },
      ]}],
      generationConfig: { temperature: 0.1, maxOutputTokens: 256 },
    }),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Gemini API エラー (${response.getResponseCode()}): ${response.getContentText()}`);
  }

  let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
  Logger.log(`Gemini応答: ${text}`);

  // マークダウンのコードブロック(```json ... ```)を除去
  text = text.replace(/```json\s*/g, "").replace(/```\s*/g, "");

  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error("Gemini の応答からJSONを抽出できませんでした: " + text);
  }
  return JSON.parse(jsonMatch[0]);
}

/** ファイル名のサニタイズ */
function sanitizeFileName_(name) {
  if (!name) return null;
  const s = name
    .replace(/\.pdf$/i, "")
    .replace(/[\/\\:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
  return s || null;
}

/** ログ用スプレッドシートを取得または作成 */
function getOrCreateLogSheet_() {
  let ss;
  if (CONFIG.SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } else {
    ss = SpreadsheetApp.create("PDF リネーム記録");
    const file = DriveApp.getFileById(ss.getId());
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    CONFIG.SPREADSHEET_ID = ss.getId();
    Logger.log(`シート作成: https://docs.google.com/spreadsheets/d/${ss.getId()}`);
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([
      ["処理日時", "旧ファイル名", "新ファイル名", "Gemini要約", "ファイルID", "ファイルURL"],
    ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);
    [160, 250, 250, 400, 200, 300].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

/** スプレッドシートに記録を追加 */
function appendLog_(sheet, record) {
  sheet.appendRow([
    record.timestamp, record.oldName, record.newName,
    record.summary || "", record.fileId, record.fileUrl,
  ]);
}

/** フォルダ内のPDF一覧を取得 */
function getPdfFiles_(folder) {
  const files = [];
  const it = folder.getFilesByType("application/pdf");
  while (it.hasNext()) files.push(it.next());
  return files;
}
