/**
 * 宮工房 オフィス要件ヒアリングシート — バックエンド (Google Apps Script)
 *
 * このスクリプトを、ヒアリング結果を集約したいGoogleスプレッドシートに
 * コンテナバインド（スプレッドシート付属のApps Script）として設置し、
 * ウェブアプリとしてデプロイして使います。手順は README.md 参照。
 *
 * シート構成（初回実行時に自動作成されます）:
 *   - "Submissions" … 送信された回答の1件1行ログ（設問ごとに列を分けた、日本語見出しの一覧）
 *   - "Companies"   … 企業別リンクのトークン → 会社名 の対応表（任意・記録用）
 *
 * Submissionsシートの見出しは日本語（人が読むため）だが、admin.html等プログラムからの
 * 読み書きは SUBMISSIONS_FIELDS の内部キー（英語・位置固定）で行うため、見出し文言を
 * 変えても読み取りは壊れない。
 */

var SUBMISSIONS_SHEET = "Submissions";
var COMPANIES_SHEET = "Companies";
var DRIVE_FOLDER_NAME = "宮工房ヒアリング添付";

// 内部キー（順序固定・admin.html等が参照） / シートに印字する日本語見出し
// 先頭3列（送信日時・送信元区分・リンクトークン）は固定表示され、誰からの回答かひと目でわかるようにしている。
var SUBMISSIONS_FIELDS = [
  { key: "timestamp", header: "送信日時" },
  { key: "senderType", header: "送信元区分" },
  { key: "token", header: "リンクトークン" },
  { key: "linkCompany", header: "登録会社名（トークン）" },
  { key: "caseName", header: "案件名" },
  { key: "company", header: "会社名（本人記入）" },
  { key: "contact", header: "ご担当者名・役職" },
  { key: "address", header: "現在のオフィス所在地" },
  { key: "moveDate", header: "入居希望日・移転期限" },
  { key: "budget", header: "想定予算" },
  { key: "sizeNote", header: "想定面積" },
  { key: "projectType", header: "プロジェクト種別" },
  { key: "background", header: "検討の背景" },
  { key: "backgroundOther", header: "検討の背景（その他）" },
  { key: "headNow", header: "現在の人数" },
  { key: "headMove", header: "入居時の想定人数" },
  { key: "headFuture", header: "将来の想定人数（3年後目安）" },
  { key: "seatType", header: "座席タイプ" },
  { key: "seatTypeOther", header: "座席タイプ（その他）" },
  { key: "deskSize", header: "希望の机サイズ" },
  { key: "imageKeywords", header: "求める空間イメージ" },
  { key: "imageOther", header: "空間イメージ（その他）" },
  { key: "areas", header: "必要な機能・エリア" },
  { key: "areaOther", header: "その他のエリア" },
  { key: "equipment", header: "設備・技術要件" },
  { key: "equipmentOther", header: "設備・技術要件（その他）" },
  { key: "priorities", header: "優先要件" },
  { key: "priorityOther", header: "優先要件（その他）" },
  { key: "drawings", header: "図面データの有無" },
  { key: "notes", header: "面談メモ・特記事項" },
  { key: "testFitDate", header: "テストフィット希望日" },
  { key: "photosJson", header: "添付写真（リンク・内部用）" },
  { key: "docsJson", header: "添付資料（リンク・内部用）" },
  { key: "summaryText", header: "要件サマリー（全文）" },
  { key: "rawJson", header: "RAW JSON（内部用・編集しないでください）" },
  { key: "id", header: "ID（内部用）" }
];
var SENDER_ID_FROZEN_COLS = 3; // 送信日時・送信元区分・リンクトークン を固定表示
var COMPANIES_HEADERS = ["トークン", "会社名", "発行日時", "備考"];

// 初期の管理者パスワード（admin.html用）。スプレッドシートのメニュー
// 「宮工房ヒアリングツール」→「管理者パスワードを設定…」からいつでも変更できます。
var DEFAULT_ADMIN_PASSWORD = "miyakobo-7k2x9q";

/* ───────────────────────────── 設定 ───────────────────────────── */

function getNotifyEmail_() {
  var v = PropertiesService.getScriptProperties().getProperty("NOTIFY_EMAIL");
  return v || "ishikawa@miyakobo.com";
}
function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || DEFAULT_ADMIN_PASSWORD;
}

/* ───────────────────────────── シート ヘルパー ───────────────────────────── */

function getOrCreatePlainSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatSubmissionsSheet_(sheet, headers) {
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(SENDER_ID_FROZEN_COLS);
  sheet.setColumnWidths(1, headers.length, 160);
  sheet.setColumnWidth(headers.length, 420); // 要件サマリー等は広めに

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setBackground("#e9e9ea");

  var fullRange = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 200), headers.length);
  try { fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); } catch (e) {}

  var senderTypeCol = SUBMISSIONS_FIELDS.map(function (f) { return f.key; }).indexOf("senderType") + 1;
  if (senderTypeCol > 0) {
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith("⚠")
      .setFontColor("#b3382c").setBold(true)
      .setRanges([sheet.getRange(2, senderTypeCol, Math.max(sheet.getMaxRows() - 1, 1), 1)])
      .build();
    sheet.setConditionalFormatRules([rule]);
  }
}

function getOrCreateSubmissionsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  var headers = SUBMISSIONS_FIELDS.map(function (f) { return f.header; });
  if (!sheet) {
    sheet = ss.insertSheet(SUBMISSIONS_SHEET);
    sheet.appendRow(headers);
    formatSubmissionsSheet_(sheet, headers);
  } else if (sheet.getLastRow() === 0) {
    // シートはあるが空（ヘッダー未設定）の場合のみ、日本語見出し・書式を設定する
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatSubmissionsSheet_(sheet, headers);
  }
  return sheet;
}

function getOrCreateDriveFolder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

/* ───────────────────────────── サマリー用の小さな整形ヘルパー ───────────────────────────── */

function withQty_(name, cap, capUnit, count, countUnit) {
  var parts = [];
  if (cap) parts.push(cap + capUnit);
  if (count) parts.push(count + countUnit);
  return parts.length ? name + "（" + parts.join("・") + "）" : name;
}
function roomsSummary_(name, rows, capUnit, countUnit) {
  var list = (rows || []).filter(function (r) { return r && (r.cap || r.count); });
  if (!list.length) return name;
  var parts = list.map(function (r) {
    var p = [];
    if (r.cap) p.push(r.cap + capUnit);
    if (r.count) p.push(r.count + countUnit);
    return p.join("・");
  });
  return name + "（" + parts.join("、") + "）";
}
function computeDeskSize_(checks, radio, text) {
  if ((checks.deskNone || []).length) return "特になし";
  var w = radio.deskW || text.deskWOther, d = radio.deskD || text.deskDOther;
  var parts = [];
  if (w) parts.push("W" + w);
  if (d) parts.push("D" + d);
  return parts.length ? parts.join(" × ") + "mm" : "";
}
function computeSenderType_(token, linkCompany) {
  if (!token) return "社内";
  if (linkCompany) return linkCompany;
  return "⚠未登録トークン：" + token; // 発行し忘れ・URL改ざん等の可能性。Companiesシートを確認してください
}
function computeAreas_(checks, rooms, text) {
  var list = checks.areas || [];
  return list.map(function (a) {
    if (a === "会議室") return roomsSummary_(a, rooms.meeting, "名用", "室");
    if (a === "フォンブース") return withQty_(a, text.capPhoneBooth, "名用", text.qtyPhoneBooth, "個");
    if (a === "倉庫・書庫・ロッカー") {
      var sizes = checks.storageSize || [];
      return sizes.length ? a + "（" + sizes.join("・") + "）" : a;
    }
    return a;
  }).join("、");
}

/* ───────────────────────────── doPost: 回答の受信 ───────────────────────────── */

function saveAttachments_(files, folder) {
  var refs = [], blobs = [];
  (files || []).forEach(function (p) {
    if (!p || !p.content) return;
    try {
      var bytes = Utilities.base64Decode(p.content);
      var blob = Utilities.newBlob(bytes, p.mimeType || "application/octet-stream", p.filename || "file");
      blobs.push(blob);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      refs.push({ url: file.getUrl(), name: p.filename || file.getName() });
    } catch (err) {
      // 1件失敗しても他の処理は続行する
    }
  });
  return { refs: refs, blobs: blobs };
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var text = body.text || {}, checks = body.checks || {}, radio = body.radio || {}, rooms = body.rooms || {};
    var token = String(body.token || "").trim();

    var linkCompany = "";
    if (token) {
      var companies = getOrCreatePlainSheet_(COMPANIES_SHEET, COMPANIES_HEADERS);
      var crows = companies.getDataRange().getValues();
      for (var i = 1; i < crows.length; i++) {
        if (String(crows[i][0]) === token) { linkCompany = String(crows[i][1] || ""); break; }
      }
    }

    var folder = getOrCreateDriveFolder_();
    var photoResult = saveAttachments_(body.photos, folder);
    var docResult = saveAttachments_(body.docs, folder);
    var attachments = photoResult.blobs.concat(docResult.blobs);

    var id = Utilities.getUuid();
    var timestamp = body.submittedAt || new Date().toISOString();

    var data = {
      id: id, timestamp: timestamp, token: token, linkCompany: linkCompany,
      senderType: computeSenderType_(token, linkCompany),
      caseName: text.caseName || "", company: text.company || "", contact: text.contact || "",
      address: text.address || "", moveDate: text.moveDate || "", budget: text.budget || "",
      sizeNote: text.sizeNote || "", projectType: radio.projectType || "",
      background: (checks.background || []).join("、"), backgroundOther: text.backgroundOther || "",
      headNow: text.headNow || "", headMove: text.headMove || "", headFuture: text.headFuture || "",
      seatType: (checks.seatType || []).join("、"), seatTypeOther: text.seatTypeOther || "",
      deskSize: computeDeskSize_(checks, radio, text),
      imageKeywords: (checks.imageKeywords || []).join("、"), imageOther: text.imageOther || "",
      areas: computeAreas_(checks, rooms, text), areaOther: text.areaOther || "",
      equipment: (checks.equipment || []).join("、"), equipmentOther: text.equipmentOther || "",
      priorities: (checks.priorities || []).join("、"), priorityOther: text.priorityOther || "",
      drawings: (checks.drawings || []).join("、"),
      notes: text.notes || "", testFitDate: text.testFitDate || "",
      photosJson: JSON.stringify(photoResult.refs), docsJson: JSON.stringify(docResult.refs),
      summaryText: body.summaryText || "", rawJson: JSON.stringify(body)
    };

    var sheet = getOrCreateSubmissionsSheet_();
    var row = SUBMISSIONS_FIELDS.map(function (f) { return data[f.key] != null ? data[f.key] : ""; });
    sheet.appendRow(row);

    // 通知メール（送信元・記入者に関わらず、常にこのスクリプトの実行アカウントから送信される）
    var subject = "【オフィス要件ヒアリング】" + (text.caseName || text.company || "新規案件")
      + "（" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "）";
    var bodyText = (body.summaryText || "（内容なし）")
      + "\n\n─────────────\n"
      + "リンクトークン: " + (token || "（社内利用・トークンなし）")
      + (linkCompany ? "\n登録会社名: " + linkCompany : "");
    var mailOptions = {};
    if (attachments.length) mailOptions.attachments = attachments;
    GmailApp.sendEmail(getNotifyEmail_(), subject, bodyText, mailOptions);

    return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ───────────────────────────── doGet: 管理者一覧 ───────────────────────────── */

function doGet(e) {
  var mode = e.parameter.mode || "";

  if (mode === "list") {
    var pass = e.parameter.pass || "";
    var adminPass = getAdminPassword_();
    if (!adminPass || pass !== adminPass) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "unauthorized" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getOrCreateSubmissionsSheet_();
    var values = sheet.getDataRange().getValues();
    var submissions = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (!row[0]) continue; // 空行はスキップ
      var obj = {};
      SUBMISSIONS_FIELDS.forEach(function (f, idx) { obj[f.key] = row[idx]; });
      try { obj.photos = JSON.parse(obj.photosJson || "[]"); } catch (e2) { obj.photos = []; }
      try { obj.docs = JSON.parse(obj.docsJson || "[]"); } catch (e3) { obj.docs = []; }
      delete obj.photosJson; delete obj.docsJson; delete obj.rawJson;
      submissions.push(obj);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, submissions: submissions }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: "miyakobo-office-hearing" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────────── 企業別リンクの発行（スプレッドシートのメニューから） ───────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("宮工房ヒアリングツール")
    .addItem("新しい企業リンクを発行…", "promptNewToken")
    .addItem("管理者パスワードを設定…", "promptSetAdminPassword")
    .addToUi();
}

function promptNewToken() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt("新しい企業リンクの発行", "会社名を入力してください（例: A社）", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var companyName = res.getResponseText().trim();
  if (!companyName) return;

  var token = generateToken_();
  var companies = getOrCreatePlainSheet_(COMPANIES_SHEET, COMPANIES_HEADERS);
  companies.appendRow([token, companyName, new Date().toISOString(), ""]);

  ui.alert(
    "発行しました",
    companyName + " 様宛リンクのトークン:\n\n" + token +
    "\n\nフォームのURLに ?c=" + token + " を付けて送付してください。\n" +
    "例: https://<あなたのサイト>/index.html?c=" + token,
    ui.ButtonSet.OK
  );
}

function promptSetAdminPassword() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt("管理者パスワードの設定", "admin.html でのログインに使うパスワードを入力してください", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var pass = res.getResponseText();
  if (!pass) return;
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", pass);
  ui.alert("設定しました。");
}

function generateToken_() {
  var chars = "abcdefghijkmnpqrstuvwxyz23456789"; // 紛らわしい文字(0,1,l,o)は除外
  var out = "";
  for (var i = 0; i < 10; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}
