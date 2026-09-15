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
  { key: "company", header: "会社名（案件）", section: "" },
  { key: "timestamp", header: "送信日時", section: "" },
  { key: "senderType", header: "送信元区分", section: "" },
  { key: "assignee", header: "社内担当者", section: "" },
  { key: "status", header: "対応状況", section: "" },
  { key: "oneLineSummary", header: "ひとことまとめ", section: "" },
  { key: "token", header: "リンクトークン", section: "" },
  { key: "linkCompany", header: "登録会社名（トークン）", section: "" },

  { key: "authorCompany", header: "会社名", section: "00 入力者情報" },
  { key: "authorName", header: "お名前", section: "00 入力者情報" },
  { key: "authorEmail", header: "メールアドレス", section: "00 入力者情報" },
  { key: "authorTel", header: "電話番号", section: "00 入力者情報" },

  { key: "contact", header: "ご担当者名・役職", section: "01 プロジェクト概要" },
  { key: "address", header: "移転先のオフィス所在地", section: "01 プロジェクト概要" },
  { key: "moveDate", header: "入居希望日・移転期限", section: "01 プロジェクト概要" },
  { key: "budget", header: "想定予算", section: "01 プロジェクト概要" },
  { key: "sizeNote", header: "想定面積", section: "01 プロジェクト概要" },
  { key: "projectType", header: "プロジェクト種別", section: "01 プロジェクト概要" },
  { key: "background", header: "検討の背景", section: "01 プロジェクト概要" },
  { key: "backgroundOther", header: "検討の背景（その他）", section: "01 プロジェクト概要" },
  { key: "priorities", header: "優先要件", section: "01 プロジェクト概要" },
  { key: "priorityOther", header: "優先要件（その他）", section: "01 プロジェクト概要" },

  { key: "headNow", header: "現在の人数", section: "02 人員・座席" },
  { key: "headMove", header: "入居時の想定人数", section: "02 人員・座席" },
  { key: "headFuture", header: "将来の想定人数（3年後目安）", section: "02 人員・座席" },
  { key: "seatType", header: "座席タイプ", section: "02 人員・座席" },
  { key: "seatTypeOther", header: "座席タイプ（その他）", section: "02 人員・座席" },
  { key: "deskSize", header: "希望の机サイズ", section: "02 人員・座席" },

  { key: "imageKeywords", header: "求める空間イメージ", section: "03 求める空間イメージ" },
  { key: "imageOther", header: "空間イメージ（その他）", section: "03 求める空間イメージ" },

  { key: "areas", header: "必要な機能・エリア", section: "04 必要な機能・設備" },
  { key: "areaOther", header: "その他のエリア", section: "04 必要な機能・設備" },
  { key: "equipment", header: "設備・技術要件", section: "04 必要な機能・設備" },
  { key: "equipmentOther", header: "設備・技術要件（その他）", section: "04 必要な機能・設備" },

  { key: "drawings", header: "図面データの有無", section: "05 特記事項・物件資料" },
  { key: "notes", header: "面談メモ・特記事項", section: "05 特記事項・物件資料" },
  { key: "testFitDate", header: "テストフィット希望日", section: "05 特記事項・物件資料" },
  { key: "photos", header: "添付写真（ファイル名）", section: "05 特記事項・物件資料" },
  { key: "docs", header: "添付資料（ファイル名）", section: "05 特記事項・物件資料" },

  { key: "summaryText", header: "要件サマリー（全文）", section: "" },
  { key: "rawJson", header: "RAW JSON（内部用・編集しないでください）", section: "" },
  { key: "id", header: "ID（内部用）", section: "" }
];
var FROZEN_COLS = 6; // 会社名（案件）・送信日時・送信元区分・社内担当者・対応状況・ひとことまとめ を固定表示
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

  // セクション（01〜08、および案件名などの識別列）ごとに背景色を交互に変え、
  // 境目に太めの縦線を入れて「どこからどこまでが同じ設問グループか」を一目でわかるようにする
  var SECTION_COLORS = ["#e4e4e7", "#ececee"]; // 01/03/05/07 と 02/04/06/08 で交互
  var IDENT_COLOR = "#d3d3d6"; // 案件名など、セクションを持たない識別列
  var maxRows = Math.max(sheet.getMaxRows(), 200);
  var col = 1, sectionIndex = -1;
  while (col <= n) {
    var section = SUBMISSIONS_FIELDS[col - 1].section;
    var start = col;
    while (col <= n && SUBMISSIONS_FIELDS[col - 1].section === section) col++;
    var width = col - start;
    var color;
    if (section) {
      sectionIndex++;
      if (width > 1) sheet.getRange(1, start, 1, width).merge();
      color = SECTION_COLORS[sectionIndex % SECTION_COLORS.length];
    } else {
      color = IDENT_COLOR;
    }
    sheet.getRange(1, start, 1, width).setBackground(color);
    sheet.getRange(2, start, 1, width).setBackground(color);
    if (start > 1) {
      sheet.getRange(1, start, maxRows, 1)
        .setBorder(null, true, null, null, null, null, "#8a8a8a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(FROZEN_COLS);
  sheet.setColumnWidths(1, n, 160);
  sheet.setColumnWidth(SUBMISSIONS_FIELDS.map(function (f) { return f.key; }).indexOf("company") + 1, 220);
  sheet.setColumnWidth(SUBMISSIONS_FIELDS.map(function (f) { return f.key; }).indexOf("oneLineSummary") + 1, 320);
  sheet.setColumnWidth(n, 420); // 最後列（RAW JSON）は広め

  sheet.getRange(1, 1, 1, n).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(2, 1, 1, n).setFontWeight("bold");

  // 縞模様（1行おきの色分け）はデータ行だけに適用し、見出し2行の色分けは崩さない
  var dataRange = sheet.getRange(3, 1, Math.max(maxRows - 2, 1), n);
  try { dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch (e) {}

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
function storageSummary_(rows) {
  var list = (rows || []).filter(function (r) { return r && (r.w || r.d || r.count); });
  if (!list.length) return "倉庫・書庫・ロッカー";
  var parts = list.map(function (r) {
    var size = [r.w ? "W" + r.w : "", r.d ? "D" + r.d : ""].filter(function (x) { return x; }).join(" × ");
    return [size ? size + "mm" : "", r.count ? r.count + "個" : ""].filter(function (x) { return x; }).join(" ");
  });
  return "倉庫・書庫・ロッカー（" + parts.join("、") + "）";
}
function computeAreas_(checks, rooms) {
  var list = checks.areas || [];
  return list.map(function (a) {
    if (a === "会議室") return roomsSummary_(a, rooms.meeting, "名用", "室");
    if (a === "フォンブース") return roomsSummary_(a, rooms.phoneBooth, "名用", "個");
    if (a === "倉庫・書庫・ロッカー") return storageSummary_(rooms.storage);
    return a;
  }).join("、");
}

// 中学生が読んでも内容がわかるような、平易な一言サマリーを作る（管理者一覧でひと目で状況を把握するため）
function buildPlainSummary_(text, radio, checks) {
  var who = text.company || text.authorCompany || "（会社名未記入）";
  var name = text.contact ? "（" + text.contact + "様）" : "";
  var parts = [];
  parts.push(who + name + "からの回答です。");

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
    if (!String(text.authorCompany || "").trim()) missing.push("会社名");
    if (!String(text.authorName || "").trim()) missing.push("お名前");
    if (!isValidEmail_(text.authorEmail)) missing.push("メールアドレス");
    if (!String(text.authorTel || "").trim()) missing.push("電話番号");
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
      authorCompany: text.authorCompany || "", authorName: text.authorName || "",
      authorEmail: text.authorEmail || "", authorTel: text.authorTel || "",
      company: text.company || "", contact: text.contact || "",
      address: text.address || "", moveDate: text.moveDate || "", budget: text.budget || "",
      sizeNote: text.sizeNote || "", projectType: radio.projectType || "",
      background: (checks.background || []).join("、"), backgroundOther: text.backgroundOther || "",
      priorities: (checks.priorities || []).join("、"), priorityOther: text.priorityOther || "",
      headNow: text.headNow || "", headMove: text.headMove || "", headFuture: text.headFuture || "",
      seatType: (checks.seatType || []).join("、"), seatTypeOther: text.seatTypeOther || "",
      deskSize: computeDeskSize_(checks, radio, text),
      imageKeywords: (checks.imageKeywords || []).join("、"), imageOther: text.imageOther || "",
      areas: computeAreas_(checks, rooms), areaOther: text.areaOther || "",
      equipment: (checks.equipment || []).join("、"), equipmentOther: text.equipmentOther || "",
      drawings: (checks.drawings || []).join("、"),
      notes: text.notes || "", testFitDate: text.testFitDate || "",
      photos: photoResult.names.join("、"), docs: docResult.names.join("、"),
      summaryText: body.summaryText || "", rawJson: JSON.stringify(body)
    };

    var sheet = getOrCreateSubmissionsSheet_();
    var row = SUBMISSIONS_FIELDS.map(function (f) { return data[f.key] != null ? data[f.key] : ""; });
    sheet.appendRow(row);

    // 通知メール（送信元・記入者に関わらず、常にこのスクリプトの実行アカウントから送信される）
    var subject = "【オフィス要件ヒアリング】" + (text.company || text.authorCompany || "新規案件")
      + "（" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "）";
    var bodyText = (body.summaryText || "（内容なし）")
      + "\n\n─────────────\n"
      + "リンクトークン: " + (token || "（社内利用・トークンなし）")
      + (linkCompany ? "\n登録会社名: " + linkCompany : "")
      + (text.authorName ? "\n入力者名: " + text.authorName : "")
      + (text.authorEmail ? "\n入力者メール: " + text.authorEmail : "")
      + (text.authorTel ? "\n入力者電話: " + text.authorTel : "");
    var mailOptions = {};
    if (attachments.length) mailOptions.attachments = attachments;
    GmailApp.sendEmail(getNotifyEmail_(), subject, bodyText, mailOptions);

    // 記入者本人への受付確認メール（メールアドレスは必須項目のため、通常は必ず送られる）
    sendConfirmationToSubmitter_(text.authorEmail, text.authorName, body.summaryText);

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
