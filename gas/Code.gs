/**
 * 宮工房 オフィス要件ヒアリングシート — バックエンド (Google Apps Script)
 *
 * このスクリプトを、ヒアリング結果を集約したいGoogleスプレッドシートに
 * コンテナバインド（スプレッドシート付属のApps Script）として設置し、
 * ウェブアプリとしてデプロイして使います。手順は README.md 参照。
 *
 * シート構成（初回実行時に自動作成されます）:
 *   - "Submissions" … 送信された回答の1件1行ログ（管理者一覧・メール送信の元データ）
 *   - "Companies"   … 企業別リンクのトークン → 会社名 の対応表（任意・記録用）
 */

var SUBMISSIONS_SHEET = "Submissions";
var COMPANIES_SHEET = "Companies";
var DRIVE_FOLDER_NAME = "宮工房ヒアリング写真";

var SUBMISSIONS_HEADERS = [
  "id", "timestamp", "token", "linkCompany", "company", "caseName", "contact",
  "address", "moveDate", "budget", "sizeNote", "summaryText", "photosJson", "rawJson"
];
var COMPANIES_HEADERS = ["token", "companyName", "issuedAt", "note"];

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

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateDriveFolder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

/* ───────────────────────────── doPost: 回答の受信 ───────────────────────────── */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var text = body.text || {};
    var token = String(body.token || "").trim();

    var linkCompany = "";
    if (token) {
      var companies = getOrCreateSheet_(COMPANIES_SHEET, COMPANIES_HEADERS);
      var rows = companies.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === token) { linkCompany = String(rows[i][1] || ""); break; }
      }
    }

    // 写真をDriveへ保存し、メール添付用のBlobも同時に作る
    var photoRefs = [];
    var attachments = [];
    (body.photos || []).forEach(function (p) {
      if (!p || !p.content) return;
      try {
        var bytes = Utilities.base64Decode(p.content);
        var blob = Utilities.newBlob(bytes, p.mimeType || "application/octet-stream", p.filename || "photo.jpg");
        attachments.push(blob);
        var folder = getOrCreateDriveFolder_();
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoRefs.push({ url: file.getUrl(), name: p.filename || file.getName() });
      } catch (err) {
        // 1枚失敗しても他の処理は続行する
      }
    });

    var id = Utilities.getUuid();
    var timestamp = body.submittedAt || new Date().toISOString();

    var sheet = getOrCreateSheet_(SUBMISSIONS_SHEET, SUBMISSIONS_HEADERS);
    sheet.appendRow([
      id, timestamp, token, linkCompany, text.company || "", text.caseName || "",
      text.contact || "", text.address || "", text.moveDate || "", text.budget || "",
      text.sizeNote || "", body.summaryText || "",
      JSON.stringify(photoRefs), JSON.stringify(body)
    ]);

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

    var sheet = getOrCreateSheet_(SUBMISSIONS_SHEET, SUBMISSIONS_HEADERS);
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var submissions = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var obj = {};
      headers.forEach(function (h, idx) { obj[h] = row[idx]; });
      try { obj.photos = JSON.parse(obj.photosJson || "[]"); } catch (e2) { obj.photos = []; }
      delete obj.photosJson;
      delete obj.rawJson;
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
  var companies = getOrCreateSheet_(COMPANIES_SHEET, COMPANIES_HEADERS);
  companies.appendRow([token, companyName, new Date().toISOString(), ""]);

  var url = ScriptApp.getService().getUrl(); // Web AppのURL（doGet/doPost用。フォームのURLは別途index.htmlをホストしている場所）
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
