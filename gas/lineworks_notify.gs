/**
 * LINE WORKS 通知スクリプト (Google Apps Script)
 *
 * 機能:
 *   Google Drive フォルダにスキャンされた PDF が追加されたら
 *   LINE WORKS チャットで通知を送信する
 *
 * セットアップ:
 *   1. https://script.google.com で新規プロジェクトを作成（pdf_rename.gs と同じプロジェクト推奨）
 *   2. このコードを貼り付け
 *   3. スクリプトプロパティに LINEWORKS_PRIVATE_KEY を設定
 *      （ファイル > プロジェクトのプロパティ > スクリプトのプロパティ）
 *   4. setupTrigger() を1回実行してトリガーを設定
 *
 * Drive API の有効化:
 *   エディタ左側 「サービス」> 「+」 > 「Drive API」を追加
 */

// ============================================================
// LINE WORKS 設定
// ============================================================
const LW_CONFIG = {
  CLIENT_ID: "XOIKJqIn7YBZmIn4M8DW",
  CLIENT_SECRET: "7Tpzd_mewT",
  SERVICE_ACCOUNT: "3ort2.serviceaccount@visitcareaoi",
  DOMAIN_ID: "400199767",
  BOT_ID: "10723461",

  // 通知先のチャンネルID（トークルームID）
  // 設定方法: getChannelList() を実行して確認
  CHANNEL_ID: "",

  // または、個人ユーザーに通知（ユーザーIDを指定）
  // 空の場合は CHANNEL_ID を使用
  NOTIFY_USER_ID: "",

  // 監視する Google Drive フォルダ ID
  WATCH_FOLDER_ID: "1nX01Z1GH6TzqzlG3pEvVUi5tUimQy4QC",

  // チェック間隔（分）- トリガーで使用
  CHECK_INTERVAL_MINUTES: 5,
};

// LINE WORKS API エンドポイント
const LW_API = {
  AUTH: "https://auth.worksmobile.com/oauth2/v2.0/token",
  BOT_MESSAGE: `https://www.worksapis.com/v1.0/bots/${LW_CONFIG.BOT_ID}/channels/{channelId}/messages`,
  BOT_USER_MESSAGE: `https://www.worksapis.com/v1.0/bots/${LW_CONFIG.BOT_ID}/users/{userId}/messages`,
  CHANNELS: `https://www.worksapis.com/v1.0/bots/${LW_CONFIG.BOT_ID}/channels`,
};

// ============================================================
// トリガー管理
// ============================================================

/**
 * フォルダ監視トリガーを設定する（初回のみ実行）
 */
function setupTrigger() {
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "checkForNewFiles") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 定期実行トリガーを作成
  ScriptApp.newTrigger("checkForNewFiles")
    .timeDriven()
    .everyMinutes(LW_CONFIG.CHECK_INTERVAL_MINUTES)
    .create();

  Logger.log(
    `トリガーを設定しました: ${LW_CONFIG.CHECK_INTERVAL_MINUTES} 分間隔で checkForNewFiles を実行`
  );

  // 初回のチェックポイントを保存
  const now = new Date();
  PropertiesService.getScriptProperties().setProperty(
    "LAST_CHECK_TIME",
    now.toISOString()
  );
  Logger.log(`初回チェックポイント: ${now.toISOString()}`);
}

/**
 * トリガーを停止する
 */
function stopTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "checkForNewFiles") {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  Logger.log(`${removed} 件のトリガーを削除しました。`);
}

// ============================================================
// ファイル監視
// ============================================================

/**
 * 新しいファイルがフォルダに追加されたかチェックする（トリガーから呼ばれる）
 */
function checkForNewFiles() {
  const props = PropertiesService.getScriptProperties();
  const lastCheckStr = props.getProperty("LAST_CHECK_TIME");
  const lastCheck = lastCheckStr ? new Date(lastCheckStr) : new Date();

  const now = new Date();
  props.setProperty("LAST_CHECK_TIME", now.toISOString());

  const folder = DriveApp.getFolderById(LW_CONFIG.WATCH_FOLDER_ID);
  const files = folder.getFilesByType("application/pdf");

  const newFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    const createdDate = file.getDateCreated();

    if (createdDate > lastCheck) {
      newFiles.push(file);
    }
  }

  if (newFiles.length === 0) {
    return;
  }

  Logger.log(`新規 PDF ファイル: ${newFiles.length} 件`);

  // 通知メッセージを作成
  const message = buildNotificationMessage_(newFiles);

  // LINE WORKS に通知
  try {
    sendLineWorksMessage_(message);
    Logger.log("LINE WORKS に通知を送信しました。");
  } catch (e) {
    Logger.log(`通知送信エラー: ${e.message}`);
  }

  // PDF リネーム処理を実行（pdf_rename.gs の main() が同プロジェクトにある場合）
  try {
    if (typeof main === "function") {
      main();
      Logger.log("PDF リネーム処理を実行しました。");
    }
  } catch (e) {
    Logger.log(`リネーム処理エラー: ${e.message}`);
  }
}

// ============================================================
// LINE WORKS 認証
// ============================================================

/**
 * LINE WORKS のアクセストークンを取得する（JWT 認証）
 */
function getAccessToken_() {
  // キャッシュからトークンを取得
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get("LW_ACCESS_TOKEN");
  if (cachedToken) return cachedToken;

  // JWT を作成
  const jwt = createJwt_();

  // アクセストークンを取得
  const response = UrlFetchApp.fetch(LW_API.AUTH, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      assertion: jwt,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      client_id: LW_CONFIG.CLIENT_ID,
      client_secret: LW_CONFIG.CLIENT_SECRET,
      scope: "bot,bot.message",
    },
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());

  if (result.access_token) {
    // トークンをキャッシュ（有効期限の少し前まで）
    const ttl = (result.expires_in || 3600) - 60;
    cache.put("LW_ACCESS_TOKEN", result.access_token, ttl);
    return result.access_token;
  }

  throw new Error(`認証エラー: ${JSON.stringify(result)}`);
}

/**
 * JWT (JSON Web Token) を作成する
 */
function createJwt_() {
  const header = {
    alg: "RS256",
    typ: "JWT",
  };

  const now = Math.floor(Date.now() / 1000);
  const claimSet = {
    iss: LW_CONFIG.CLIENT_ID,
    sub: LW_CONFIG.SERVICE_ACCOUNT,
    iat: now,
    exp: now + 3600,
  };

  const encodedHeader = base64UrlEncode_(JSON.stringify(header));
  const encodedClaimSet = base64UrlEncode_(JSON.stringify(claimSet));
  const signatureInput = encodedHeader + "." + encodedClaimSet;

  // 秘密鍵で署名
  const privateKey = getPrivateKey_();
  const signature = Utilities.computeRsaSha256Signature(
    signatureInput,
    privateKey
  );
  const encodedSignature = base64UrlEncode_(signature);

  return signatureInput + "." + encodedSignature;
}

/**
 * スクリプトプロパティから秘密鍵を取得する
 */
function getPrivateKey_() {
  const key = PropertiesService.getScriptProperties().getProperty(
    "LINEWORKS_PRIVATE_KEY"
  );
  if (!key) {
    throw new Error(
      "秘密鍵が設定されていません。スクリプトプロパティに LINEWORKS_PRIVATE_KEY を設定してください。"
    );
  }
  return key;
}

/**
 * Base64 URL エンコード
 */
function base64UrlEncode_(input) {
  let encoded;
  if (typeof input === "string") {
    encoded = Utilities.base64Encode(input, Utilities.Charset.UTF_8);
  } else {
    encoded = Utilities.base64Encode(input);
  }
  return encoded.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

// ============================================================
// メッセージ送信
// ============================================================

/**
 * LINE WORKS にメッセージを送信する
 */
function sendLineWorksMessage_(message) {
  const token = getAccessToken_();

  let url;
  if (LW_CONFIG.NOTIFY_USER_ID) {
    url = LW_API.BOT_USER_MESSAGE.replace(
      "{userId}",
      LW_CONFIG.NOTIFY_USER_ID
    );
  } else if (LW_CONFIG.CHANNEL_ID) {
    url = LW_API.BOT_MESSAGE.replace("{channelId}", LW_CONFIG.CHANNEL_ID);
  } else {
    throw new Error(
      "通知先が設定されていません。CHANNEL_ID または NOTIFY_USER_ID を設定してください。"
    );
  }

  const payload = {
    content: {
      type: "text",
      text: message,
    },
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  if (statusCode !== 200 && statusCode !== 201) {
    throw new Error(
      `メッセージ送信失敗 (${statusCode}): ${response.getContentText()}`
    );
  }

  return JSON.parse(response.getContentText());
}

/**
 * 新規ファイルの通知メッセージを組み立てる
 */
function buildNotificationMessage_(files) {
  let message = `📄 新しいスキャン PDF が ${files.length} 件追加されました\n\n`;

  for (const file of files) {
    const name = file.getName();
    const size = formatFileSize_(file.getSize());
    const date = Utilities.formatDate(
      file.getDateCreated(),
      "Asia/Tokyo",
      "yyyy/MM/dd HH:mm"
    );
    message += `• ${name}\n  サイズ: ${size} | 追加日時: ${date}\n`;
    message += `  ${file.getUrl()}\n\n`;
  }

  message += `フォルダ: https://drive.google.com/drive/folders/${LW_CONFIG.WATCH_FOLDER_ID}`;

  return message;
}

/**
 * ファイルサイズを読みやすい形式にフォーマットする
 */
function formatFileSize_(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

// ============================================================
// ユーティリティ（デバッグ用）
// ============================================================

/**
 * Bot が参加しているチャンネル一覧を取得する（CHANNEL_ID 確認用）
 */
function getChannelList() {
  const token = getAccessToken_();

  const response = UrlFetchApp.fetch(LW_API.CHANNELS, {
    method: "get",
    headers: {
      Authorization: "Bearer " + token,
    },
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * テスト通知を送信する
 */
function sendTestNotification() {
  const message =
    "🔔 テスト通知\n\nPDF スキャン監視が正常に動作しています。\n送信日時: " +
    Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  sendLineWorksMessage_(message);
  Logger.log("テスト通知を送信しました。");
}

/**
 * 現在の設定状況を表示する
 */
function showStatus() {
  Logger.log("=== 設定状況 ===");
  Logger.log(`Client ID: ${LW_CONFIG.CLIENT_ID}`);
  Logger.log(`Service Account: ${LW_CONFIG.SERVICE_ACCOUNT}`);
  Logger.log(`Domain ID: ${LW_CONFIG.DOMAIN_ID}`);
  Logger.log(`Bot ID: ${LW_CONFIG.BOT_ID}`);
  Logger.log(`Channel ID: ${LW_CONFIG.CHANNEL_ID || "(未設定)"}`);
  Logger.log(`Notify User ID: ${LW_CONFIG.NOTIFY_USER_ID || "(未設定)"}`);
  Logger.log(`Watch Folder ID: ${LW_CONFIG.WATCH_FOLDER_ID}`);

  const lastCheck = PropertiesService.getScriptProperties().getProperty(
    "LAST_CHECK_TIME"
  );
  Logger.log(`最終チェック: ${lastCheck || "(未実行)"}`);

  const hasKey = PropertiesService.getScriptProperties().getProperty(
    "LINEWORKS_PRIVATE_KEY"
  );
  Logger.log(`秘密鍵: ${hasKey ? "設定済み" : "未設定"}`);

  const triggers = ScriptApp.getProjectTriggers();
  const activeTriggers = triggers.filter(
    (t) => t.getHandlerFunction() === "checkForNewFiles"
  );
  Logger.log(`監視トリガー: ${activeTriggers.length > 0 ? "稼働中" : "停止中"}`);
}
