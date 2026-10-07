# アーム ねじ穴タップチェックアプリ（野田組）— 引き継ぎメモ

## 現行：GAS版トレーサビリティ「aming」（`gas/`）— 2026-10-07〜
アプリ名は **aming**（画面タイトル・ブラウザタブ・Apps Scriptプロジェクト名・通知メール件名【aming】）。
ユーザー確定：**GAS＋スプレッドシート**で作る。目的は4つ全部（刻印→全履歴／建機号機から逆引き／野田組のタップ記録／出荷前の未チェック検出）。撮影して号機が特定できたら**記録画面を開く**。生産管理Excelは毎日/毎週もらう。

| ファイル | 内容 |
|---|---|
| `gas/01_設定.js` | CONFIG（ID類・Excel列名→台帳列名・シート名・モデル） |
| `gas/02_メイン.js` | `doGet`、`setup()`、画面API `api*`（検索・詳細・記録・取消・出荷日一覧・未チェック・CSV） |
| `gas/03_取込.js` | Excel(.xlsm)→Googleシート変換→`parseShipValues_`→台帳へ上書きマージ。列は**見出し名で探す** |
| `gas/03_照合.js` | `fitZuban_`/`fitGoki_`/`wdist_`/`rankArms_`（純粋関数・nodeでテスト可） |
| `gas/03_AI読取.js` | `apiRead`：Gemini（既定）か Claude で2段階読取（JSONスキーマ指定）＋刻印写真をDrive保存 |
| `gas/03_フォント見本.js` | 刻印機 Telesis TMC470 の11x16ドットフォント見本（PNG base64）。AI読取の1枚目に添える。元画像 `data/telesis_font_11x16.png` |
| `gas/03_通知.js` | `notifyUnchecked`：次の出荷日（台帳上の翌出荷日）分で「良」がないアームをメール。`TRACK_FROM`（運用開始日 2026/10/08）より前の出荷は対象外 |
| `gas/04_テスト.js` | GASエディタで実行する確認用関数 |
| `gas/index.html` | 画面（読取／一覧／記録／設定）。カメラ枠・切り抜き・凹凸強調は旧版から移植 |
| `test/gas.test.js` | 照合・Excel解析・画面(jsdom＋google.script.runモック)のテスト |

- 台帳キー＝`図番_号機`。元Excel「出荷明細」シート（ヘッダー5行目、A列「注文番号（写し）」）。3年分約8,400行→アーム7,918本（ブームブラケット等は `EXCLUDE_KISHU` で除外、図番+号機の重複63件は出荷日が新しい方を採用し「要確認」シートへ）
- **工程ログ**：①現場記録「工程記録」シート（塗装完了／塗装後修正完了／出荷を、設定タブの自分の名前＋日時＋メモで記録。追記のみ・取消フラグ。`STEPS`）②「変更履歴」シート（取込で `TRACK_CHANGE_COLS`＝日付6列＋出荷先・建機号機 が変わったら 前/後/元ファイル を追記。`diffArm_`）。詳細画面の工程ログに Excel日付・日程変更・現場記録を工程ごとに表示。新シートは初回アクセスで自動作成（`makeSheet_`）
- タップ記録は**追記のみ**。訂正は取消フラグ→再記録。最新の有効記録がそのアームの状態
- ねじ穴数は「穴数マスタ」に図番ごとに記憶し、次から自動入力。「全数OK」＝良＋処置数=ねじ穴数
- 実データでのテスト：`test/fixtures/ship_values.json`（.gitignore済・コミットしない）があれば `npm test` で解析も確認
- AI読取は `AI_PROVIDER`（既定 gemini / `gemini-3.8-flash`）。キーはスクリプトプロパティ `GEMINI_API_KEY`（claude なら `ANTHROPIC_API_KEY`）。ID類は空なら `setup()` がスクリプトプロパティに保存

### 作成済みのGoogle側（2026-10-07、Driveコネクタで作成。IDは `01_設定.js` に埋め込み済み → `setup()` 不要）
- フォルダ「アームトレーサビリティ」 https://drive.google.com/drive/folders/1qEAxXRBG7y0edZhMCNvVLe4QjBjlIZw8
  - 台帳スプレッドシート `13Fj89c-17Ec0YsUUuKdIykSkZV2442SrA35UNoMHUd8`（6シート・見出し・書式TEXT済、穴数マスタに2件）
  - _取込済 `14XGF3BnMr2-fdt4wiqFMejk7UD5RxSuW`（変換一時置き場）／ アーム刻印写真 `1g82k8Y4gCAlBVXAWbeSTkvlHUvcNJ1_0` ／ アーム出荷明細_取込 `17Wno8F_...`（未使用）
- 取込元：既存フォルダ「'001_アーム出荷明細」`1NS4WoClO0xlGWSxvFimFcqFGUT0jOKQL`（`出荷予定　日程表変更A(26年10月7日).xlsm` が随時たまる）。ファイル名の日付で最新1本だけ取込、取込済みは `LAST_IMPORTED`（スクリプトプロパティ）で判定。元ファイルは動かさない
- Apps Scriptプロジェクト：scriptId `1Wm4YGpwWU-pTtsohOmXprg1ssC4NRJqUHo4RjqVP0qK6zm0X9zf_NttW`（`gas/.clasp.json`）。2026-10-07 にクラウド環境から clasp push 済み
- ウェブアプリ @1：https://script.google.com/a/macros/nodagumi40.com/s/AKfycbxASmlgFgrQyYGJ1e5OtipzMPre5pfU6H_Lv_IfasA69LtgXnsZ4LPky-XXTH0S5buI3g/exec
  - 更新は `clasp push` → `clasp deploy -i AKfycbxASmlgFgrQyYGJ1e5OtipzMPre5pfU6H_Lv_IfasA69LtgXnsZ4LPky-XXTH0S5buI3g`（同じURLのまま新版）
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
