const CONFIG = {
  FOLDER_ID: "ここにDriveのフォルダIDを貼り付け",
  DATA_SHEET_NAME: "ここにスプレッドシート名を記載", // 使わないなら空でOK
  LOG_SHEET_NAME: "Log",
  PROCESSED_TAG: "[processed]",
  FAILED_TAG: "[failed]",
  OCR_DOC_TRASH: true,         // OCR Docをゴミ箱へ
  OCR_DOC_PREFIX: "OCR_",      // OCR Docの先頭文字
  GEMINI_MODEL_PREFERENCE: "flash" // 自動選択でflash優先
};

function processReceipts() {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const files = folder.getFiles();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getActiveSheet();
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME) || ss.insertSheet(CONFIG.LOG_SHEET_NAME);

  ensureLogHeader_(logSheet);

  while (files.hasNext()) {
    const file = files.next();

    // 画像/PDFのみ
    const mime = file.getMimeType();
    if (![MimeType.JPEG, MimeType.PNG, MimeType.PDF].includes(mime)) {
      log_(logSheet, "SKIP_MIME", file, { mime });
      continue;
    }

    // 処理済みスキップ
    if (file.getName().includes(CONFIG.PROCESSED_TAG)) {
      log_(logSheet, "SKIP_PROCESSED", file, {});
      continue;
    }

    // 失敗済みはスキップ（無限リトライ防止）
if (file.getName().includes(CONFIG.FAILED_TAG)) {
  log_(logSheet, "SKIP_FAILED", file, {});
  continue;
}

    const startedAt = new Date();
    let ocrDocId = "";
    let ocrText = "";
    let parsed = null;

    try {
      // OCR doc生成
      ocrDocId = ocrToGoogleDoc_(file.getId());
      log_(logSheet, "OCR_CREATED", file, { ocrDocId });

      // OCRテキスト取得
      const doc = DocumentApp.openById(ocrDocId);
      ocrText = doc.getBody().getText();

      // Geminiで抽出
      parsed = extractMedicalInfo(ocrText);

      // 形式を正規化＆最低限バリデーション
      const result = normalizeMedicalResult_(parsed);

      if (!result.date || !result.amount) {
        log_(logSheet, "SKIP_INCOMPLETE", file, { result, parsed });
        continue;
      }

      // データ書き込み
      dataSheet.appendRow([
        result.date,
        result.hospital,
        result.amount,
        result.description || "医療費",
        result.payer || "",
        file.getName()
      ]);

      // 元ファイルを処理済みにする（重複防止）
      file.setName(file.getName() + " " + CONFIG.PROCESSED_TAG);

      log_(logSheet, "SUCCESS", file, {
        elapsedMs: new Date() - startedAt,
        ocrDocId,
        result
      });

    } catch (e) {
      log_(logSheet, "ERROR", file, {
        elapsedMs: new Date() - startedAt,
        ocrDocId,
        error: String(e),
        stack: e && e.stack ? e.stack : ""
      });
      // ★追加：failedタグ付け
       try {
    file.setName(file.getName() + " " + CONFIG.FAILED_TAG);
    log_(logSheet, "MARK_FAILED", file, {});
  } catch (nameErr) {
    log_(logSheet, "MARK_FAILED_ERROR", file, { error: String(nameErr) });
  }

    } finally {
      // OCR Doc を確実にゴミ箱へ（フォルダを汚さない）
      if (CONFIG.OCR_DOC_TRASH && ocrDocId) {
        try {
          DriveApp.getFileById(ocrDocId).setTrashed(true);
          log_(logSheet, "OCR_TRASHED", file, { ocrDocId });
        } catch (trashErr) {
          log_(logSheet, "OCR_TRASH_FAILED", file, { ocrDocId, error: String(trashErr) });
        }
      }
    }
  }
}

/**
 * 画像/PDFをOCRしてGoogleドキュメントを作成し、そのDoc IDを返す
 * ※Advanced Google Services の「Drive API」を有効化している前提
 */
function ocrToGoogleDoc_(fileId) {
  const resource = {
    title: "OCR_" + fileId,
    mimeType: MimeType.GOOGLE_DOCS
  };
  const options = {
    ocr: true,
    ocrLanguage: "ja"
  };

  const copied = Drive.Files.copy(resource, fileId, options);

  Logger.log("Drive.Files.copy => %s", JSON.stringify(copied));

  if (!copied || !copied.id) {
    throw new Error("OCR doc creation failed: " + JSON.stringify(copied));
  }
  return copied.id;
}

function extractMedicalInfo(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  // 1) モデル一覧取得
  const listRes = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey,
    { method: "get", muteHttpExceptions: true }
  );
  if (listRes.getResponseCode() !== 200) {
    throw new Error("Failed to list models: " + listRes.getContentText());
  }
  const listJson = JSON.parse(listRes.getContentText());

  // 2) generateContent対応モデルを1つ選ぶ（flash優先）
  const models = (listJson.models || []);
  const candidate = models.find(m =>
    (m.supportedGenerationMethods || []).includes("generateContent") &&
    (m.name || "").includes("flash")
  ) || models.find(m =>
    (m.supportedGenerationMethods || []).includes("generateContent")
  );

  if (!candidate || !candidate.name) {
    throw new Error("No generateContent model available for this API key.");
  }

  // candidate.name は "models/xxxx" 形式
  const modelName = candidate.name.replace("models/", "");
  Logger.log("Using model: %s", modelName);

  const prompt = `
あなたは日本の確定申告用の医療費領収書解析AIです。
以下のOCRテキストから医療費情報を抽出してください。
必ずJSONのみを返してください。

{
 "date": "YYYY-MM-DD",
 "hospital": "医療機関名",
 "amount": 数値のみ,
 "description": "医療費",
 "payer": ""
}

OCRテキスト:
${text}
`.trim();

  // 3) generateContent
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }]
    }),
    muteHttpExceptions: true
  });

  const body = res.getContentText();
  if (res.getResponseCode() !== 200) {
    throw new Error("Gemini error " + res.getResponseCode() + ": " + body);
  }

  const json = JSON.parse(body);
  const raw = json.candidates?.[0]?.content?.parts?.[0]?.text || "";
  const cleaned = raw.replace(/```json/g, "").replace(/```/g, "").trim();

  return JSON.parse(cleaned);
}

function listGeminiModels_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  const url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;

  const res = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
  Logger.log("code=%s", res.getResponseCode());
  Logger.log(res.getContentText());
}

function ensureLogHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["timestamp", "status", "fileName", "fileId", "detailsJson"]);
  }
}

function log_(sheet, status, file, details) {
  sheet.appendRow([
    new Date(),
    status,
    file ? file.getName() : "",
    file ? file.getId() : "",
    JSON.stringify(details || {})
  ]);
}

/**
 * Geminiの返り値を安全に整形
 */
function normalizeMedicalResult_(obj) {
  const r = obj || {};

  // date
  let date = (r.date || "").trim();
  // 2025/05/02 -> 2025-05-02
  date = date.replace(/\//g, "-");

  // hospital
  let hospital = (r.hospital || "").trim();
  hospital = hospital.replace(/\s+/g, ""); // 改行・空白の混入を潰す（医院名分断対策）

  // amount
  let amount = r.amount;
  if (typeof amount === "string") {
    amount = amount.replace(/,/g, "").replace(/[^\d]/g, "");
    amount = amount ? Number(amount) : "";
  }
  if (typeof amount === "number" && !Number.isFinite(amount)) amount = "";

  return {
    date,
    hospital,
    amount,
    description: (r.description || "医療費").trim(),
    payer: (r.payer || "").trim()
  };
}

function cleanupOcrDocsInFolder_() {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const files = folder.getFiles();
  let count = 0;

  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === MimeType.GOOGLE_DOCS && f.getName().startsWith(CONFIG.OCR_DOC_PREFIX)) {
      f.setTrashed(true);
      count++;
    }
  }
  Logger.log("Trashed OCR docs: %s", count);
}
