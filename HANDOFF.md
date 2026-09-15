# 引き継ぎ資料：宮工房 オフィス要件ヒアリングサイト

このファイルは、別のAI（別ツール／別セッション）にこのプロジェクトを引き継ぐための資料です。
下の「■ 新しいAIへの最初のプロンプト」をそのままコピーして、新しいAIの最初のメッセージとして送ってください。
このファイル自体もあわせて読み込ませる（貼り付ける、またはファイルとして渡す）と、より正確に引き継げます。

---

## プロジェクト概要

宮工房（オフィス移転・改装を手がける会社）向けの「オフィス要件ヒアリングシート」Webツール。
取引先企業・社内スタッフがWebフォームでヒアリング項目に回答すると、
- 内容がGoogleスプレッドシートに自動記録され
- 担当者（ishikawa@miyakobo.com）に通知メールが届き（写真・資料も添付）
- 記入者本人にも「受け付けました」の確認メールが届く

という仕組み。管理者（社内）はパスワード保護された管理画面、またはスプレッドシートで一覧確認する。

**最重要の設計要件**：複数の取引先企業（A社・B社…）に配布しても、企業ごとの提出内容は
お互いに一切見えない・存在にも気づかれないこと。ログインID発行のような運用負担は避けること。

## リポジトリ・環境情報

- GitHubリポジトリ: `miyakobo/office-hearing-sheet.miyakobo`
  （個人アカウント`raraji827`から、URL先頭の`raraji827`表記を消すためOrganization `miyakobo`に移動済み。
  git remoteが古い`raraji827/...`のままでも、GitHubの自動リダイレクトでfetch/pushは動作する）
- 作業ブランチ: `claude/website-build-t9qcds`（これまでの全作業はここにコミット・プッシュ済み）
- 公開サイト（GitHub Pages）:
  - フォーム: `https://miyakobo.github.io/office-hearing-sheet.miyakobo/index.html`
  - 管理画面: `https://miyakobo.github.io/office-hearing-sheet.miyakobo/admin.html`
- バックエンド（Google Apps Script Web App URL、`index.html`/`admin.html`双方にハードコード済み）:
  `https://script.google.com/macros/s/AKfycby3A8SC1FOxnVkloGa3YPX9U0LDtwEQCVEtyVxhBFzR9V9DqXTPLNncykenxI2HTmrZvw/exec`
- 管理画面パスワード: デフォルト `miyakobo-7k2x9q`（スプレッドシートのメニュー「宮工房ヒアリングツール」→
  「管理者パスワードを設定…」でいつでも変更可能。Script Property `ADMIN_PASSWORD` に保存される）
- 通知先メールアドレス: デフォルト `ishikawa@miyakobo.com`（Script Property `NOTIFY_EMAIL` で上書き可能）

## ファイル構成

```
index.html      … お客様（社外・社内）が記入する8ステップのヒアリングフォーム（自己完結HTML/CSS/JS）
admin.html      … 社内管理者専用の回答一覧・詳細（パスワード保護、閲覧専用）
gas/Code.gs     … バックエンド本体（Google Apps Script）。doPost=回答受信、doGet=一覧取得
gas/appsscript.json … GASのマニフェスト（OAuthスコープ・Webアプリ公開設定）
README.md       … セットアップ手順・運用マニュアル（かなり詳しく書いてあるので必読）
photos/mood-1.webp 〜 mood-10.webp … フォーム内の「空間イメージ」参考写真
```

## アーキテクチャ

```
企業担当者・社内スタッフ
   │  https://.../index.html?c=<会社ごとのトークン>（トークンなし=社内利用モード）
   ▼
index.html（GitHub Pagesで静的配信）
   │  fetch POST（JSON、写真・資料はbase64） → Content-Type: text/plain;charset=utf-8
   │  （CORSプリフライトを避けるため。GASはOPTIONSを処理できない）
   ▼
Google Apps Script Web App: doPost(e)
   │─ Companiesシートでトークン→会社名を照合（"送信元区分"列に反映）
   │─ Submissionsシートに1行追記
   │─ ishikawa@miyakobo.com へ通知メール（写真・資料を添付、GmailApp.sendEmail）
   └─ 記入者のメールアドレスがあれば、本人にも受付確認メールを送信

社内管理者
   │  https://.../admin.html → パスワード入力
   ▼
GAS Web App: doGet(?mode=list&pass=...)  ← パスワード一致時のみ全件返却
```

**分離の担保**: フォーム側（index.html）には他社データを読み返す手段が一切ない。
一覧取得エンドポイント（doGet）は管理者パスワード必須。これにより「ログインID発行」なしで
企業ごとの完全な分離を実現している。

## これまでの主な意思決定（背景・理由）

1. **なぜGoogle Apps Script + スプレッドシートか**：当初はClaude Artifactの`claude.use("db"/"mcp")`系
   APIで組もうとしたが、外部の取引先担当者はClaude上にGmail連携等を持たないため不安定と判断。
   固定のGoogleアカウント（miyakobo.com）からメールを送る構成にすることで、
   「誰が入力しても必ずishikawa@miyakobo.comに届く」を確実に担保できる。
2. **なぜトークン付きリンクで、ログインID方式にしなかったか**：ログインID＋パスワードにすると
   企業側にID管理の負担が生じる一方、実現できることはトークンリンクと同じ
   （「そのリンクを持つ人だけがその会社として送信できる」）。追加メリットがないため見送った。
3. **なぜGoogle Driveを使っていないか**：当初は写真・資料をDriveに保存する設計だったが、
   `drive.file`スコープでは`DriveApp.getFoldersByName()`も`DriveApp.createFolder()`も権限不足で
   繰り返しエラーになった。ユーザーから「最低限のことだけやりたい（権限周りのトラブルはもう避けたい）」
   という明確な要望があったため、**Drive連携を完全に廃止**し、写真・資料はメールへの直接添付のみに
   変更した（Gmail 1通25MB上限に注意。フォーム側で写真1枚8MB×5枚、資料1件10MB×5件に制限）。
   結果、必要なOAuthスコープは `spreadsheets.currentonly` と `gmail.send` のみ。
4. **URL先頭の`raraji827`表記について**：当初GitHub Pagesの個人アカウントのプロジェクトページは
   URL先頭に`raraji827.github.io`が出る仕様で変更できず、独自ドメイン（例：hearing.miyakobo.com）か
   別ホスティングへの切り替えを提案していたが、DNS設定が必要でユーザーの「最低限のことだけやりたい」
   という要望に合わないため保留にしていた。**その後、GitHub上で完結する解決策として、リポジトリを
   個人アカウント`raraji827`からOrganization `miyakobo`へ移動**し、URLを
   `https://miyakobo.github.io/office-hearing-sheet.miyakobo/` に変更済み。独自ドメイン化は
   引き続き任意・保留（`README.md`にその旨と手順を明記済み。`CNAME`ファイルは未使用）。
5. **フォームのデザイン・設問構成の正解データ**：ユーザーが本当に使いたかったのは、当初私が独自に
   作ったデザイン・設問ではなく、元ZIPに入っていた`#Uオフィス要件ヒアリング.dc.html`
   （Claude Design Canvasの下書き）の内容そのものだった。見た目はモノトーン（ink/paper調、
   角丸カード、Barlow/Barlow Condensedフォント）で、steel-blueのブループリント調デザインは不採用。
   設問は8ステップ（01 プロジェクト概要 〜 08 特記事項・物件資料）。
6. **管理者スプレッドシートの再設計の経緯**：最初は列見出しが英語＋内容が1列にまとまっていて
   「見にくい」と繰り返し指摘された。→ 日本語見出し化 → 設問ごとに列分割 → 送信元区分列の追加 →
   さらに直近で、01〜08のセクションごとに2段見出しでグループ化し、
   「社内担当者」「対応状況（未対応/対応中/完了のプルダウン）」の手入力列と、
   自動生成される「ひとことまとめ」列（中学生が読んでもわかる1〜2文要約）を追加した。

## 直近で実装済みの内容（このセッションの最終コミット）

コミット `4c3f244`（ブランチ`claude/website-build-t9qcds`にプッシュ済み）：

- `index.html`：01の概要に「ご担当者のメールアドレス」「ご担当者のお電話番号」欄を追加
- `gas/Code.gs`：
  - 送信データに`contactEmail`/`contactPhone`を記録
  - 記入者本人への受付確認メール送信（`sendConfirmationToSubmitter_`関数、メールアドレスが
    妥当な形式のときだけ送信。失敗しても本処理は止めない設計）
  - `SUBMISSIONS_FIELDS`に`section`プロパティを追加し、01〜08のセクションでグルーピング
  - `formatSubmissionsSheet_()`を2段見出し（1段目=セクション名を結合セル表示、2段目=設問名）に刷新
  - `assignee`（社内担当者）・`status`（対応状況、プルダウン検証付き）・`oneLineSummary`
    （`buildPlainSummary_()`で自動生成する平易な要約）列を追加
  - `doGet()`の一覧取得ロジックをヘッダー2行分オフセットするよう修正
- `admin.html`：一覧・詳細に対応状況・社内担当者・ひとことまとめ・担当者メール・電話番号を表示
  （編集はスプレッドシート側で行う運用とし、admin.htmlは完全に閲覧専用）
- `README.md`：上記の説明、および「この変更はOAuthスコープを増やしていないので再認可不要」
  「既存のSubmissionsシートは`Submissions_old`にリネームしてから1件テスト送信すると
  新しい2段見出しシートが自動生成される」という移行手順を追記

ユーザー確認：`node --check`によるgas/Code.gsの構文チェック、Playwrightでのローカル画面確認
（新しいフォーム項目の表示、admin.htmlでのモックデータ表示）は完了。
**GAS本体への実デプロイ・実送信テスト（メール到達確認含む）はユーザー側で実施予定**（このセッションの
サンドボックス環境からは`script.google.com`等への外向き通信がネットワークポリシーでブロックされている
ため、実機テストができない）。

## 既知の制約・ハマりどころ

- **サンドボックスのネットワーク制限**：このAI実行環境からは`script.google.com`・
  `*.github.io`への疎通が組織ポリシーでブロックされている（`curl`すると`connect_rejected`）。
  実際のデプロイ後の動作確認は必ずユーザー側のブラウザで行う必要がある。
- **リポジトリの所有者が個人アカウント`raraji827`からOrganization `miyakobo`に変わった**：
  git remoteのURLは`raraji827/...`のままでも、GitHubの自動リダイレクトによりfetch/pushは問題なく
  動作する。ただしGitHub MCPツール（`add_repo`等）はセッション開始時のオーナーでスコープが固定される
  実装のため、新しいセッションを`miyakobo/office-hearing-sheet.miyakobo`で開始しないと、
  GitHub API経由の操作（Issue/PR作成など）はできない可能性がある。
- **GASのOAuthスコープ変更は要注意**：スコープを変更すると、次回デプロイ時にユーザーへの
  再認可（新しい同意画面）が発生する。ユーザーはこれに一度混乱した経緯があるため、
  スコープを変える変更をする場合は**必ず事前にその旨を明示**すること。
  現在のスコープは `spreadsheets.currentonly` + `gmail.send` のみ（Driveは不使用）。
- **Playwrightでのクリックの不安定さ**：`page.click()`が反応しないことがあり、
  `page.evaluate(() => el.click())`で直接クリックさせる方が確実だった。
- **GASへのPOSTはCORSプリフライトを避けるため`Content-Type: text/plain;charset=utf-8`固定**。
  `application/json`にするとGASがOPTIONSを処理できず失敗する。
- **ユーザーの運用スタンス**：「本当に何もわからないので、最低限のことだけやりたい」と明言している。
  複雑な追加提案（新しい外部サービス連携、DNS設定、ログインID管理など）は基本的に避け、
  今ある構成（GAS+スプレッドシート+GitHub Pages）の中でシンプルに解決する方針を貫くこと。

## 未着手・保留中の項目

- カスタムドメイン（`hearing.miyakobo.com`等）への切り替え（DNS設定がユーザー側で必要なため保留）
- 既存の本番Submissionsシートの見出しは旧バージョン（1段）の可能性があり、新しい2段見出しへの
  移行（`Submissions_old`へのリネーム→再送信によるシート再生成）はユーザーがまだ実施していない
- 実機での送信テスト（社内モード／登録済みトークン／未登録トークンの3パターン）と、
  受付確認メール・通知メールの実際の着信確認はユーザー側で未実施

## 追記（引き継ぎ後もこのセッションで継続することになった）

ユーザーは結局このセッションで作業を継続することにしたため、上の「新しいAIへの最初のプロンプト」は
未使用。ただし内容は引き続き正確なので、今後また移行したくなった場合に備えて残してある。

このあと追加で実装した内容：

- 01の先頭にある会社名・メールアドレス・お電話番号を**必須項目**にした（画面側でNext/送信を
  ブロックし、Apps Script側でも同じ3項目を検証してから記録するダブルチェック）
- 一時的に「ご依頼主区分」（お客様／協力業者様／社内のラジオ選択）も追加したが、
  ユーザーから「いらないので削除して」との指示があり、index.html・gas/Code.gs・admin.html・
  README.mdから完全に削除済み（コミット履歴には残っているが、現在のコードには存在しない）。
  もし今後また「誰の立場で送っているか」を明示的に選ばせたい要望が出た場合は、この経緯を
  踏まえて必要性を確認してから再実装すること。
- その後、ユーザーがClaude Design Canvasで作った新しいHTML（"Bundled Page"形式でアップロード）を
  もとに、index.htmlを全面刷新した。旧8ステップ（01〜08）構成をやめ、**6ステップ（00〜05）＋確認画面**
  に再構成：
  - **00 入力者情報**（新設）: 会社名・お名前・メールアドレス・電話番号の4つを**必須項目**として
    独立させた（これが「ご依頼主区分」削除後の必須化の後継。フォームに実際に記入している人の情報）
  - 01 プロジェクト概要（旧01+02検討の背景+07優先要件を統合。背景・優先要件は「＋詳しく入力する
    （任意）」で開閉するdetailブロックに）
  - 02 人員・座席（旧03。机サイズはdetailブロックへ）
  - 03 求める空間イメージ（旧04。「その他」記述はdetailブロックへ）
  - 04 必要な機能・設備（旧05+06。設備・技術要件はdetailブロックへ）
  - 05 特記事項・物件資料（旧08のまま）
  - **「案件名」フィールドは廃止**（新しいHTMLに存在しないため）。フォンブース・倉庫のサイズ入力も
    「大/中/小」チップ方式から、会議室と同じ「複数行・数値入力（W/D/個数など）」方式に統一
  - `gas/Code.gs`のSUBMISSIONS_FIELDSも合わせて全面更新（`authorCompany`/`authorName`/`authorEmail`/
    `authorTel`を追加、`caseName`/`contactEmail`/`contactPhone`を削除、シート一番左の識別列は
    `company`（案件の会社名）に変更）。受付確認メールの宛先も`authorEmail`に変更。
  - `index.html`のレンダリングエンジンに、セクション内の`detail:{...}`（任意の追加設問ブロック、
    「＋ 詳しく入力する」で開閉）という新しい仕組みを追加した。今後セクションを追加・変更する際は
    この`detail`パターンを使うと、必須の主要項目と任意の詳細項目を分けて見せられる。

---

## ■ 新しいAIへの最初のプロンプト（そのままコピーして送ってください）

```
あなたには、宮工房という会社向けの「オフィス要件ヒアリングシート」Webツールの開発を
引き継いでもらいます。GitHubリポジトリ miyakobo/office-hearing-sheet.miyakobo の
ブランチ claude/website-build-t9qcds に、これまでの実装が全てコミット・プッシュ済みです。

まずリポジトリ内の HANDOFF.md を読んでください。これまでの経緯・設計判断・現在の状態・
既知の制約が詳しくまとめてあります。次に README.md（セットアップ・運用マニュアル）、
index.html（フォーム本体）、admin.html（管理画面）、gas/Code.gs（バックエンド）を
一通り確認し、現状を理解してから作業を始めてください。

重要な運用方針：ユーザーは技術に詳しくなく、「最低限のことだけやりたい」と明言しています。
複雑な追加提案（新しい外部サービス、DNS設定、ログイン管理システムなど）は避け、
今の構成（Google Apps Script + スプレッドシート + GitHub Pages、Driveは不使用）の中で
シンプルに解決する方針を守ってください。また、Google Apps ScriptのOAuthスコープを
変更する提案をする場合は、再認可（同意画面）が必要になることを必ず事前に伝えてください。
```
