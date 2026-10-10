# アーム ねじ穴タップチェックアプリ（野田組）— 引き継ぎメモ

## 現行：GAS版トレーサビリティ「aming」（`gas/`）— 2026-10-07〜
アプリ名は **aming**（画面タイトル・ブラウザタブ・Apps Scriptプロジェクト名・通知メール件名【aming】）。
ユーザー確定：**GAS＋スプレッドシート**で作る。目的は4つ全部（刻印→全履歴／建機号機から逆引き／野田組のタップ記録／出荷前の未チェック検出）。撮影して号機が特定できたら**記録画面を開く**。生産管理Excelは毎日/毎週もらう。

| ファイル | 内容 |
|---|---|
| `gas/01_設定.js` | CONFIG（初期値）と `SETTINGS`。**運用値は台帳の「設定」シートが優先**（項目名で照合、`cfg_`/`cfgNum_`/`cfgList_` で読む）。SS_IDだけコード、APIキーはスクリプトプロパティ |
| `gas/02_メイン.js` | `doGet`、`setup()`、画面API `api*`（検索・詳細・記録・取消・出荷日一覧・未チェック・CSV） |
| `gas/03_取込.js` | Excel(.xlsm)→Googleシート変換→`parseShipValues_`→台帳へ上書きマージ。列は**見出し名で探す** |
| `gas/03_照合.js` | `fitZuban_`/`fitGoki_`/`wdist_`/`rankArms_`（純粋関数・nodeでテスト可） |
| `gas/03_AI読取.js` | `apiRead`：Gemini（既定）か Claude で2段階読取（JSONスキーマ指定）＋刻印写真をDrive保存 |
| `gas/03_キュー.js` | **撮るだけモード**（10/10〜メイン）：`apiQueueAdd`（全体写真を長辺2000で保存→「読取キュー」シート）→`apiQueueProcess`（写真全体をGeminiへ：`P_FULL` で刻印の枠 box[ymin,xmin,ymax,xmax 0-1000]＋読み→照合）→画面が `straightCrop`（枠の周りを広めに取り、打刻の点の行投影で傾き±25°を測って回転→文字範囲2〜98%で切り詰め。失敗時は `cropByBox`）で切り出して `apiQueueCrop`（確定ならひも付け、候補なら切り出し＋強調で読み直し）。`apiQueuePick`/`apiQueueDone`/`apiQueueList`。目的が塗装完了・出荷なら一覧から直接「○○を記録」／「まとめて記録」（`apiQueueRecord`＝工程記録へ追記、キューの「記録」列に記録ID、`apiQueueUnrecord` で取消）。塗装後チェックは詳細を開いてタップ記録。目的は色分け（塗装完了=ブルーグリーン/塗装後チェック=オレンジ/出荷=緑）し撮るボタンに【目的】表示、既定は塗装完了。画面を閉じても `processQueue` トリガー（5分おき、setupTriggersで作成）が待ちを読む |
| `gas/03_フォント見本.js` | 刻印機 Telesis TMC470 の11x16ドットフォント見本（PNG base64）。AI読取の1枚目に添える。元画像 `data/telesis_font_11x16.png` |
| `gas/03_通知.js` | `notifyUnchecked`：次の出荷日（台帳上の翌出荷日）分で「良」がないアームをメール。`TRACK_FROM`（運用開始日 2026/10/08）より前の出荷は対象外 |
| `gas/04_テスト.js` | GASエディタで実行する確認用関数 |
| `gas/index.html` | 画面。見た目はダーク×メカ×パワーショベル、色はコベルコ建機のブランドカラー（白×ブルーグリーン `--accent:#00A7B5`。ハザード帯だけ安全表示の黄黒）（ハザード帯・面取りパネル・ボルト・LED・キャタピラ帯・ヘッダーに設計図風の線画ショベル（20tクラス側面図・寸法線・支点の中心マーク）＝実機比率（ブーム5.7m/アーム2.9m）で「伸ばす→掘る→すくう→持ち上げ→ダンプ」の掘削サイクルアニメ（`docs/excavator_cycle.png`）、アームだけ白＋発光。AI読取中は全画面ショベル `#busy`）。ヘッダーに本日塗装予定（スマホは右上、幅520px以上はAMINGとショベルの間に縦並び＝ショベルをflexに戻す。Excel塗装完了日＝今日の本数と「残り N」（未塗装数、0なら「完了」で緑）、`apiList` の paint、タップで一覧）。ロゴ等は使わずオリジナル描画。スクショ `docs/screen_*.png`。（読取／一覧＝出荷日→図番ごとに号機タイル・進捗バー（塗装後修正済）・絞込み（すべて/未塗装/塗装済/未塗装後修正/塗装後修正済。`armView_` の painted/fixed）・タイルに次工程のExcel予定日（塗＝塗装完了日／修＝塗装後修正完了日、過ぎたら赤。Excelの日付は予定＝ユーザー確認済み）・出荷前未チェックは出荷日ごとに折りたたみ。`apiList` で今日〜21日分をまとめて取得し localStorage `trace.list` に保存→開くと即表示・裏で更新（起動時に先読み、日付切替は通信なし）／記録／設定＝CSVと画面の版のみ。確認者（自分の名前）は読取画面の目的の下で選ぶ（localStorage `trace.me`、未選択で記録しようとすると `needMe()` で点滅）。確認者の追加は設定シート、取込はトリガー任せ）。カメラ枠・切り抜き・凹凸強調は旧版から移植 |
| `test/gas.test.js` | 照合・Excel解析・画面(jsdom＋google.script.runモック)のテスト |

- AI読取の確定：`isSure_`（1位が距離0かつ2位が2以上）**または** `isExact_`（AIの図番・号機が?なしで1位と完全一致）なら候補一覧を出さず記録画面を自動で開く（10/10 ユーザー確定）
- **台帳キャッシュ**：`loadArms_` は CacheService（JSON→gzip→base64を90KBずつ分割、`arms_0..n`＋`arms_n`、6時間）から読む。7,918本で約240KB＝3分割。無ければシートから読んで保存。取込（`mergeArms_`）で `clearArmsCache_()`
- 台帳キー＝`図番_号機`。元Excel「出荷明細」シート（ヘッダー5行目、A列「注文番号（写し）」）。3年分約8,400行→アーム7,918本（ブームブラケット等は `EXCLUDE_KISHU` で除外、図番+号機の重複63件は出荷日が新しい方を採用し「要確認」シートへ）
- **工程ログ**：①現場記録「工程記録」シート（塗装完了／塗装後修正完了／出荷を、設定タブの自分の名前＋日時＋メモで記録。追記のみ・取消フラグ。`STEPS`）②「変更履歴」シート（取込で `TRACK_CHANGE_COLS`＝日付6列＋出荷先・建機号機 が変わったら 前/後/元ファイル を追記。`diffArm_`）。詳細画面の工程ログに Excel日付・日程変更・現場記録を工程ごとに表示。新シートは初回アクセスで自動作成（`makeSheet_`）
- **刻印写真**：撮影→枠で切り出した画像を「aming 刻印写真」フォルダの `年-月/未特定`（撮るだけの全体写真は `年-月/元写真`）に保存（`savePhoto_`/`photoFolder_`）。アームが特定されて詳細を開いた時点（自動一致 or 候補タップ）で `apiLinkPhoto` が `図番_号機_工程_日時.jpg`（工程＝塗装完了/塗装後修正完了(=塗装後チェック)/出荷）に改名して `年-月` へ移動。直下に残った古い写真は画面を開いた時 `apiAutoOrganize`→`organizePhotos` が1回だけ整理（スクリプトプロパティ `PHOTOS_ORGANIZED`）し「刻印写真」シートにひも付け（同じ写真の付け直しは上書き）。詳細画面にサムネ（`apiPhoto` で data URL）。「削除」で `apiDeletePhoto`＝ドライブのゴミ箱へ＋シート行削除
- **撮影の目的と段階**：読取画面で「塗装完了／塗装後チェック／出荷」を選んで撮る（localStorage `trace.purpose`）。特定後の詳細の先頭に目的バナー（塗装完了・出荷はその記録ボタン、塗装後チェック＝ねじ穴タップ欄を強調）。写真名・刻印写真シートの「目的」列にも残す。段階 `stage_`＝未→塗装完了→塗装後チェック済（タップ良）→出荷済、否は要処置。一覧・候補・詳細のバッジは段階表示
- **自動学習**（10/10）：「刻印写真」シートの AI読取 と ひも付けた正解（図番・号機）が教師データ。`learned_()`（1時間キャッシュ、`apiLinkPhoto` で更新）→`learnStats_` が読み間違い（例 `7>0`）を集計 → ①照合 `wdist_` で2回以上の間違いはコスト0.15・1回は0.5（自動確定はしない＝cost0のみ）②プロンプト末尾に `learnHint_`（よくある間違い＋最近の例5件）。間違ったひも付けは刻印写真シートを直せば学習も直る。号機の前に英字（例 `2610-K2022`）が付くことがある
- タップ記録は**追記のみ**。訂正は取消フラグ→再記録。最新の有効記録がそのアームの状態
- ねじ穴数は「穴数マスタ」に図番ごとに記憶し、次から自動入力。「全数OK」＝良＋処置数=ねじ穴数
- 実データでのテスト：`test/fixtures/ship_values.json`（.gitignore済・コミットしない）があれば `npm test` で解析も確認
- AI読取は `AI_PROVIDER`（既定 gemini / `gemini-3.6-flash`、予備 `GEMINI_FALLBACK`＝3.5-flash-lite→3.8-flash。**読取ログ実測(10/9)：3.8-flashは503高負荷と無料枠429が多発、1回の503に最大30秒**。なので503/429/404は即次モデル＋CacheServiceで冷却（503=2分・429=10分・404=1時間）、500系だけ1回再試行。エラー文は画面に出す）。凹凸強調は小ぼかしでノイズ除去→大ぼかし背景との差→2/98%で伸長＋smoothstep。**iPhone Safariは canvas filter 非対応**なので、ぼかしは縮小→拡大で作る（`blurred()`）。画像は幅1100・JPEG0.85。サーバー読取は50秒締切（毎回「読取ログ」シートに 秒・結果・Gemini応答コードを記録＝Sheetsコネクタで原因調査できる）、Geminiは `thinkingLevel:low`（400なら外して再送）。キーはスクリプトプロパティ `GEMINI_API_KEY`（claude なら `ANTHROPIC_API_KEY`）。ID類は空なら `setup()` がスクリプトプロパティに保存

### 作成済みのGoogle側（2026-10-07、Driveコネクタで作成。IDは `01_設定.js` に埋め込み済み → `setup()` 不要）
- フォルダ「aming」 https://drive.google.com/drive/folders/1qEAxXRBG7y0edZhMCNvVLe4QjBjlIZw8（Apps Scriptプロジェクト「aming」もここ）
  - 「aming 台帳」スプレッドシート `13Fj89c-17Ec0YsUUuKdIykSkZV2442SrA35UNoMHUd8`（設定・アーム台帳・タップ記録・穴数マスタ・要確認・取込ログ＋工程記録・変更履歴は自動作成。確認者は設定シート最下部の「確認者」行から下へB列1行1人。A列が空の行は直前項目の続き＝`parseSettingRows_`）
  - 「aming 作業用（取込の一時ファイル）」`14XGF3BnMr2-fdt4wiqFMejk7UD5RxSuW` ／「aming 刻印写真」`1g82k8Y4gCAlBVXAWbeSTkvlHUvcNJ1_0`
- 取込元：既存フォルダ「'001_アーム出荷明細」`1NS4WoClO0xlGWSxvFimFcqFGUT0jOKQL`（`出荷予定　日程表変更A(26年10月7日).xlsm` が随時たまる）。ファイル名の日付で最新1本だけ取込、取込済みは `LAST_IMPORTED`（スクリプトプロパティ）で判定。元ファイルは動かさない
- Apps Scriptプロジェクト：scriptId `1Wm4YGpwWU-pTtsohOmXprg1ssC4NRJqUHo4RjqVP0qK6zm0X9zf_NttW`（`gas/.clasp.json`）。2026-10-07 にクラウド環境から clasp push 済み
- ウェブアプリ @1：https://script.google.com/a/macros/nodagumi40.com/s/AKfycbxASmlgFgrQyYGJ1e5OtipzMPre5pfU6H_Lv_IfasA69LtgXnsZ4LPky-XXTH0S5buI3g/exec
  - 更新は `clasp push` → `clasp deploy -i AKfycbxASmlgFgrQyYGJ1e5OtipzMPre5pfU6H_Lv_IfasA69LtgXnsZ4LPky-XXTH0S5buI3g`（同じURLのまま新版）
- **clasp ログイン（クラウド環境）**：Workspaceの再認証で `invalid_rapt` が出たら `python3 scripts/clasp_login.py start` → URLをユーザーに渡す → 戻ってきた `http://localhost:8888/?state=...&code=...` を `python3 scripts/clasp_login.py finish '<URL>'`。裏で待つ `clasp login` はターン間でプロセスが消えるので使わない（要：既存 ~/.clasprc.json の client_id/secret）
- 注意：ローカルでファイルを消しただけだと `clasp push` が「already up to date」で**リモートのファイルが残る**ことがある。Apps Script API の `projects/{id}/content` を GET→対象を除いて PUT で消す
- 注意：`clasp create` は appsscript.json を初期値で上書きするので git checkout で戻すこと
- 注意：nodagumi40.com のWorkspaceは clasp に spreadsheets/drive スコープ追加を**ブロック**。クラウド環境からは script.google.com にも出られないので、`?fn=`（gas-run.sh）での関数実行はできない。関数実行・権限承認はユーザーがブラウザで

### デプロイ手順
1. clasp：`git pull` → `cd gas && clasp create --type standalone --title "アームトレーサビリティ"` → `clasp push`
   貼り付け：`gas-paste/` の `コード.gs`・`index.html`・`appsscript.json`（エディタ設定で「マニフェストを表示」）を新規プロジェクトに貼る。`npm run gas-paste` で再生成
2. スクリプトプロパティに `GEMINI_API_KEY`（エディタで `testAiKey` 実行で確認）
3. `testImport` 実行（元フォルダの最新Excelを強制取込）
4. デプロイ→ウェブアプリ（実行：自分／アクセス：組織内）
5. トリガー：エディタで `setupTriggers` を1回実行（`importLatest` 6時台・`notifyUnchecked` 7時台、何度実行しても2本だけ）

---
以下は claude.ai アーティファクト版（旧）の引き継ぎメモ。


claude.ai チャットから Claude Code への引き継ぎ。2026-10-07 時点。

## 目的
建機アーム（コベルコ建機向け）の平面に打刻された刻印をスマホで撮影 → AIで読み取り → 出荷明細と照合して「どのアームか」を特定 → **ねじ穴タップのチェック結果を記録**する。

**重要な方針転換（ユーザー確定）：野田組がやるのは「ねじ穴タップのチェックだけ」。**
- 製造後検査（52項目）→ 削除済み
- 図面読取・寸法チェック・3D化 → 不要（図面は CONFIDENTIAL なので扱わない）
- ゴミ噛み／スパッタ／毛羽立ち／錆／膜厚／色／測定者／各種承認欄 → **これから削る**（現行コードには残っている）

## ユーザー
落合雄平（有限会社野田組・DX担当）。返答は短くカジュアルな日本語、すぐ使える完成品を好む。GAS・Python・React Native(Expo)を使う。

## ファイル
| パス | 内容 |
|---|---|
| `src/arm-kensa.html` | **本体（本番版）**。単一HTML、CSS/JS全部入り。現状は「アーム塗装後確認」 |
| `src/arm-kokuin-readonly.html` | 刻印読取だけの旧版（参考） |
| `data/shipping_oct_44.txt` | デモ用出荷明細44件。`図番\|号機\|出荷日\|出荷先\|仕様` |
| `data/demo_kokuin.jpg` | 実際の刻印写真（YN12B20269F1 / 1152号機） |
| `scripts/build_demo.py` | 本番版からデモ版 `dist/arm-kensa-demo.html` を生成（DEMO定数3行を差し替え） |
| `test/smoke.js`, `smoke2.js` | jsdomでデモ版を開いてクリック操作する簡易テスト |

```
npm i && npm run build && npm test   # "errors []" が出ればOK
```

## 公開中URL（claude.ai アーティファクト）
- 本番：https://claude.ai/artifact/ECj3Pk2iafwYZxUpenRsV6
- デモ：https://claude.ai/artifact/SJ4qCw5YbUWzkS7nUhX9qL
- 刻印読取のみ 本番/デモ：https://claude.ai/artifact/UfXWoCLCha7PyTTooYcR2C ／ https://claude.ai/artifact/KR6aAdeTte4bsgQzkjkGnL

⚠ Claude Code からはこのURLに再公開できない。改修後のHTMLを claude.ai チャットに渡して「このURLに再publish」と頼むか、下の「移植」をやる。

## 刻印の仕様
- 1行目 図番：`[A-Z]{2}\d{2}B\d{5}F\d(G\d)?` 例 `LS12B10010F1`, `YY12B00902F1G2`
- 2行目：`◇ + 製作年月YYMM + "-" + 号機 + 検査記号`。例 `◇2609-39 UM`（2026年9月製作・39号機・UT/MT検査済）
- 刻印機は **Telesis TMC470（ドットピン）、11x16フォント**。0は斜線入り、2は角ばった斜線、9は右がまっすぐ（ユーザー提供の取説 5章で確認）
- 白塗装面の浅い打刻でコントラストが非常に低い。**ライトを真横から当てるのが精度の最大要因**

## 実装済みの仕組み（src/arm-kensa.html 内）
1. 2段階読取：1回目自由読取 → 明細から上位5候補 → 2回目に写真と候補を見比べて選択。1回目で完全一致かつ紛らわしい候補なしなら2回目スキップ
2. 枠つき撮影：`getUserMedia` ライブ映像＋黄色ガイド枠、枠内だけ切り出し。使えない時はファイル選択→ドラッグ切り抜き枠（`camNote()`で理由表示）
3. 桁型補正 `fitZuban` / `fitGoki`：数字位置の Z→2, O→0, S→5, B→8 等。例 `YN12BZ0Z69F1`→`YN12B20269F1`, `11S2`→`1152`
4. あいまい照合 `wdist`：打刻で紛らわしいペア(0/O, 9/7, 6/5…)のコスト0.35の重み付き編集距離
5. 号機の連番優先：直前と同じ図番で+1〜+5ならコスト-0.3、「前回の続き番号」表示
6. 凹凸強調画像（ぼかし差分＋パーセンタイル正規化）を元画像と2枚送る
7. 明細PDF取り込み：pdf.js(cdnjs 2.16.105) → `parseMaster()`。テキスト貼り付けにも対応
8. 保存：`claude.use("db")` の `arms/{図番}_{号機}`, `days/{日付}`, `settings/people`, `settings/film`, `master/current`。`patch()` は deepMerge＋直列キュー、update失敗時set
9. CSV出力：`claude.use("downloads")`、BOM付き
10. 確認者サインは「名前タップ＋日時」（設定タブで名前一覧管理）

### claude.ai ランタイム依存（移植時に置き換えが必要）
- `claude.use("sample")` … AI読取（Claude呼び出し）
- `claude.use("db")` … Firestore風の共有DB（`doc/collection/onSnapshot/set/update`）
- `claude.use("user")`, `claude.use("downloads")`
- `window.claude` が無い環境では読取・保存が動かない（デモ版は DEMO_ROWS と擬似読取で動く）

## 次にやること（ユーザー承認待ちだった案）
ねじ穴タップ専用に作り直す：
1. 刻印を撮る（または一覧から号機を選ぶ）→ 該当アームを特定
2. 入力は **タップ 良/否、ねじ穴数（全数）、処置数、確認者（名前タップ）だけ**
3. ねじ穴数は図番ごとに記憶し、2本目以降は自動入力（例：LC12B10556F1=69, LS12B10010F1=33）
4. 「全数OK」ボタンで 良＋処置数=全数 を一発入力
5. 一覧：出荷日ごとに未チェック／済を表示。CSVも上記項目だけに

最後のユーザー発言（音声入力・途中切れ）：「写真パシャッと撮ったら（刻印を）読み取って、うん、で、その容器、この容器は…」
→ 「容器」はおそらく **「号機」の誤変換**。「撮ったら読み取って、この号機は○○」= 撮影→特定→その1本を記録、を一発でやりたい意図と推測。**続きを本人に確認すること**（済/未の表示？記録画面を開く？良を即記録？）。

## 移植の選択肢（claude.ai 外で動かすなら）
- GAS Webアプリ：DB→スプレッドシート、AI読取→GASから Anthropic API（キーはスクリプトプロパティ）。ユーザーの既存スタックに一番近い
- Expo アプリ：カメラ枠が確実に動く。バックエンドは GAS か Firebase
- Claudeアプリ内では枠つきカメラが起動しない場合あり（Safari/Chromeなら動く）

## 未解決
- 現場が入力するには、アーティファクト共有時に編集権限が必要
- 精度の定量評価用に正解付き刻印写真20〜30枚が未収集
- 出荷明細の元PDF：Googleドライブ「アーム出荷明細」fileId `1n4bX21_kUk6XBTC6_qKinZE81Wt0SS9M`
