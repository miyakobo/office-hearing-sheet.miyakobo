/**
 * 宮工房 オフィス要件ヒアリングシート — バックエンド (Google Apps Script)
 *
 * このスクリプトを、ヒアリング結果を集約したいGoogleスプレッドシートに
 * コンテナバインド（スプレッドシート付属のApps Script）として設置し、
 * ウェブアプリとしてデプロイして使います。手順は README.md 参照。
 *
 * シート構成（初回実行時に自動作成されます）:
 *   - "Submissions" … 送信された回答の1件1行ログ（01〜08のセクションごとに列をまとめた一覧）
 *   - "Companies"   … 企業別リンクのトークン → 会社名 の対応表（任意・記録用）
 *
 * Submissionsシートの見出しは日本語（人が読むため）だが、admin.html等プログラムからの
 * 読み書きは SUBMISSIONS_FIELDS の内部キー（英語・位置固定）で行うため、見出し文言を
 * 変えても読み取りは壊れない。
 *
 * 「担当者」「ステータス」列はフォームからは送られてこない、社内で直接シートに書き込む
 * 手入力の欄（案件ごとの対応管理用）。
 */

var SUBMISSIONS_SHEET = "Submissions";
var COMPANIES_SHEET = "Companies";

// 内部キー（順序固定・admin.html等が参照） / シートに印字する日本語見出し / どのセクション（01〜08）に属するか。
// section が同じ列は、シート上で見出し行がまとめて結合表示される。
var SUBMISSIONS_FIELDS = [
  { key: "timestamp", header: "送信日時", section: "" },
  { key: "senderType", header: "送信元区分", section: "" },
  { key: "assignee", header: "社内担当者", section: "" },
  { key: "status", header: "対応状況", section: "" },
  { key: "oneLineSummary", header: "ひとことまとめ", section: "" },
  { key: "token", header: "リンクトークン", section: "" },
  { key: "linkCompany", header: "登録会社名（トークン）", section: "" },

  { key: "requesterType", header: "ご依頼主区分", section: "01 プロジェクト概要" },
  { key: "caseName", header: "案件名", section: "01 プロジェクト概要" },
  { key: "company", header: "会社名（本人記入）", section: "01 プロジェクト概要" },
  { key: "contact", header: "ご担当者名・役職", section: "01 プロジェクト概要" },
  { key: "contactEmail", header: "ご担当者メールアドレス", section: "01 プロジェクト概要" },
  { key: "contactPhone", header: "ご担当者電話番号", section: "01 プロジェクト概要" },
  { key: "address", header: "現在のオフィス所在地", section: "01 プロジェクト概要" },
  { key: "moveDate", header: "入居希望日・移転期限", section: "01 プロジェクト概要" },
  { key: "budget", header: "想定予算", section: "01 プロジェクト概要" },
  { key: "sizeNote", header: "想定面積", section: "01 プロジェクト概要" },
  { key: "projectType", header: "プロジェクト種別", section: "01 プロジェクト概要" },

  { key: "background", header: "検討の背景", section: "02 検討の背景" },
  { key: "backgroundOther", header: "検討の背景（その他）", section: "02 検討の背景" },

  { key: "headNow", header: "現在の人数", section: "03 人員・座席" },
  { key: "headMove", header: "入居時の想定人数", section: "03 人員・座席" },
  { key: "headFuture", header: "将来の想定人数（3年後目安）", section: "03 人員・座席" },
  { key: "seatType", header: "座席タイプ", section: "03 人員・座席" },
  { key: "seatTypeOther", header: "座席タイプ（その他）", section: "03 人員・座席" },
  { key: "deskSize", header: "希望の机サイズ", section: "03 人員・座席" },

  { key: "imageKeywords", header: "求める空間イメージ", section: "04 求める空間イメージ" },
  { key: "imageOther", header: "空間イメージ（その他）", section: "04 求める空間イメージ" },

  { key: "areas", header: "必要な機能・エリア", section: "05 必要な機能・エリア" },
  { key: "areaOther", header: "その他のエリア", section: "05 必要な機能・エリア" },

  { key: "equipment", header: "設備・技術要件", section: "06 設備・技術要件" },
  { key: "equipmentOther", header: "設備・技術要件（その他）", section: "06 設備・技術要件" },

  { key: "priorities", header: "優先要件", section: "07 優先要件" },
  { key: "priorityOther", header: "優先要件（その他）", section: "07 優先要件" },

  { key: "drawings", header: "図面データの有無", section: "08 特記事項・物件資料" },
  { key: "notes", header: "面談メモ・特記事項", section: "08 特記事項・物件資料" },
  { key: "testFitDate", header: "テストフィット希望日", section: "08 特記事項・物件資料" },
  { key: "photos", header: "添付写真（ファイル名）", section: "08 特記事項・物件資料" },
  { key: "docs", header: "添付資料（ファイル名）", section: "08 特記事項・物件資料" },

  { key: "summaryText", header: "要件サマリー（全文）", section: "" },
  { key: "rawJson", header: "RAW JSON（内部用・編集しないでください）", section: "" },
  { key: "id", header: "ID（内部用）", section: "" }
];
var FROZEN_COLS = 5; // 送信日時・送信元区分・社内担当者・対応状況・ひとことまとめ を固定表示
var STATUS_OPTIONS = ["未対応", "対応中", "完了"];
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

function formatSubmissionsSheet_(sheet) {
  var n = SUBMISSIONS_FIELDS.length;
  var groupRow = [], labelRow = [];
  SUBMISSIONS_FIELDS.forEach(function (f) { groupRow.push(f.section || ""); labelRow.push(f.header); });

  sheet.getRange(1, 1, 1, n).setValues([groupRow]);
  sheet.getRange(2, 1, 1, n).setValues([labelRow]);

  // 同じセクション名が連続する範囲を結合して、01〜08ごとの見出しにする
  var col = 1;
  while (col <= n) {
    var section = SUBMISSIONS_FIELDS[col - 1].section;
    var start = col;
    while (col <= n && SUBMISSIONS_FIELDS[col - 1].section === section) col++;
    if (section && col - start > 1) {
      sheet.getRange(1, start, 1, col - start).merge();
    }
  }

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(FROZEN_COLS);
  sheet.setColumnWidths(1, n, 160);
  sheet.setColumnWidth(SUBMISSIONS_FIELDS.map(function (f) { return f.key; }).indexOf("oneLineSummary") + 1, 320);
  sheet.setColumnWidth(n, 420); // 最後列（RAW JSON）は広め

  var groupRange = sheet.getRange(1, 1, 1, n);
  groupRange.setFontWeight("bold").setBackground("#dcdcdc").setHorizontalAlignment("center");
  var labelRange = sheet.getRange(2, 1, 1, n);
  labelRange.setFontWeight("bold").setBackground("#e9e9ea");

  var fullRange = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 200), n);
  try { fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); } catch (e) {}

  var keys = SUBMISSIONS_FIELDS.map(function (f) { return f.key; });
  var senderTypeCol = keys.indexOf("senderType") + 1;
  if (senderTypeCol > 0) {
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith("⚠")
      .setFontColor("#b3382c").setBold(true)
      .setRanges([sheet.getRange(3, senderTypeCol, Math.max(sheet.getMaxRows() - 2, 1), 1)])
      .build();
    sheet.setConditionalFormatRules([rule]);
  }

  var statusCol = keys.indexOf("status") + 1;
  if (statusCol > 0) {
    var validation = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(true).build();
    sheet.getRange(3, statusCol, Math.max(sheet.getMaxRows() - 2, 1), 1).setDataValidation(validation);
  }
}

function getOrCreateSubmissionsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SUBMISSIONS_SHEET);
    formatSubmissionsSheet_(sheet);
  } else if (sheet.getLastRow() < 2) {
    // シートはあるが見出し未設定（新規、または旧バージョンの1行見出し）の場合のみ書式を作り直す
    formatSubmissionsSheet_(sheet);
  }
  return sheet;
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
  return parts.length ? parts.join(" × ") + "mm" : "W1200×D600mm（標準サイズで製作）";
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

// 中学生が読んでも内容がわかるような、平易な一言サマリーを作る（管理者一覧でひと目で状況を把握するため）
function buildPlainSummary_(text, radio, checks) {
  var who = text.company || text.caseName || "（会社名未記入）";
  var name = text.contact ? "（" + text.contact + "様）" : "";
  var kindOfPerson = radio.requesterType ? "【" + radio.requesterType + "】" : "";
  var parts = [];
  parts.push(kindOfPerson + who + name + "からの回答です。");

  var kind = radio.projectType || "オフィスの見直し";
  parts.push("内容は「" + kind + "」の検討。");

  var meta = [];
  if (text.moveDate) meta.push("希望時期は" + text.moveDate);
  if (text.budget) meta.push("予算は" + text.budget);
  if (text.sizeNote) meta.push("広さは" + text.sizeNote + "くらい");
  if (meta.length) parts.push(meta.join("、") + "。");

  if (text.headNow || text.headMove) {
    parts.push("人数は現在" + (text.headNow || "？") + "名くらいで、入居時は" + (text.headMove || "？") + "名くらいを想定。");
  }

  var pr = (checks.priorities || []).slice(0, 2);
  if (pr.length) parts.push("特に大事にしたいのは「" + pr.join("」「") + "」とのこと。");

  return parts.join("");
}

/* ───────────────────────────── doPost: 回答の受信 ───────────────────────────── */

function buildAttachments_(files) {
  // Driveには保存せず、メールへの添付だけを行う（Driveの権限が一切不要になる、最もシンプルな構成）。
  var names = [], blobs = [];
  (files || []).forEach(function (p) {
    if (!p || !p.content) return;
    try {
      var bytes = Utilities.base64Decode(p.content);
      var blob = Utilities.newBlob(bytes, p.mimeType || "application/octet-stream", p.filename || "file");
      blobs.push(blob);
      names.push(p.filename || "file");
    } catch (err) {
      // 1件失敗しても他の処理は続行する
    }
  });
  return { names: names, blobs: blobs };
}

function isValidEmail_(s) {
  return !!s && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s).trim());
}

function sendConfirmationToSubmitter_(email, contact, summaryText) {
  if (!isValidEmail_(email)) return;
  try {
    var subject = "【宮工房】オフィス要件ヒアリングを受け付けました";
    var body = (contact ? contact + "様\n\n" : "")
      + "このたびはオフィス要件ヒアリングにご回答いただき、誠にありがとうございます。\n"
      + "担当者に内容が届きました。確認のうえ、担当より折り返しご連絡いたしますので、今しばらくお待ちください。\n\n"
      + "――――――――――――――――\n"
      + "ご回答内容の控え\n"
      + "――――――――――――――――\n"
      + (summaryText || "（内容なし）") + "\n\n"
      + "※本メールは送信専用です。ご不明な点がございましたら、宮工房までお問い合わせください。";
    GmailApp.sendEmail(email, subject, body);
  } catch (err) {
    // 確認メールが送れなくても、本体の送信処理自体は失敗させない
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var text = body.text || {}, checks = body.checks || {}, radio = body.radio || {}, rooms = body.rooms || {};
    var token = String(body.token || "").trim();

    var missing = [];
    if (!radio.requesterType) missing.push("ご依頼主区分");
    if (!String(text.company || "").trim()) missing.push("会社名");
    if (!isValidEmail_(text.contactEmail)) missing.push("メールアドレス");
    if (!String(text.contactPhone || "").trim()) missing.push("お電話番号");
    if (missing.length) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false, error: "必須項目が未入力です：" + missing.join("、")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var linkCompany = "";
    if (token) {
      var companies = getOrCreatePlainSheet_(COMPANIES_SHEET, COMPANIES_HEADERS);
      var crows = companies.getDataRange().getValues();
      for (var i = 1; i < crows.length; i++) {
        if (String(crows[i][0]) === token) { linkCompany = String(crows[i][1] || ""); break; }
      }
    }

    var photoResult = buildAttachments_(body.photos);
    var docResult = buildAttachments_(body.docs);
    var attachments = photoResult.blobs.concat(docResult.blobs);

    var id = Utilities.getUuid();
    var timestamp = body.submittedAt || new Date().toISOString();

    var data = {
      id: id, timestamp: timestamp, token: token, linkCompany: linkCompany,
      senderType: computeSenderType_(token, linkCompany),
      oneLineSummary: buildPlainSummary_(text, radio, checks),
      requesterType: radio.requesterType || "",
      caseName: text.caseName || "", company: text.company || "", contact: text.contact || "",
      contactEmail: text.contactEmail || "", contactPhone: text.contactPhone || "",
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
      photos: photoResult.names.join("、"), docs: docResult.names.join("、"),
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
      + "ご依頼主区分: " + (radio.requesterType || "（未選択）")
      + "\nリンクトークン: " + (token || "（社内利用・トークンなし）")
      + (linkCompany ? "\n登録会社名: " + linkCompany : "")
      + (text.contactEmail ? "\nご担当者メール: " + text.contactEmail : "")
      + (text.contactPhone ? "\nご担当者電話: " + text.contactPhone : "");
    var mailOptions = {};
    if (attachments.length) mailOptions.attachments = attachments;
    GmailApp.sendEmail(getNotifyEmail_(), subject, bodyText, mailOptions);

    // 記入者本人への受付確認メール（メールアドレスの記入があった場合のみ）
    sendConfirmationToSubmitter_(text.contactEmail, text.contact, body.summaryText);

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
    for (var i = 2; i < values.length; i++) { // 0,1行目は見出し（セクション行・項目行）
      var row = values[i];
      if (!row[0]) continue; // 空行はスキップ
      var obj = {};
      SUBMISSIONS_FIELDS.forEach(function (f, idx) { obj[f.key] = row[idx]; });
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
