# Claude Code 開発ガイド

## バージョン管理

- 現在のバージョン: **v0.18.3**
- バージョン表記箇所: `templates/index.html` の `<h1>` タグ内 `<span>` 要素
- セマンティックバージョニング (`vMAJOR.MINOR.PATCH`) を使用
  - MAJOR: 未完成のため `0` を維持（正式リリースで `1` に）
  - MINOR: 機能追加・変更時にインクリメント
  - PATCH: バグ修正・軽微な変更時にインクリメント
- **変更をコミットする際は、内容に応じてバージョン番号を必ず更新すること**

## マージ手順

変更を main にマージする際は、以下の手順で行うこと：

1. `origin/main` を取得して最新に同期
2. フィーチャーブランチにコミット＆プッシュ
3. GitHub API で PR を作成（`gh` CLI が使えない環境では `curl` + proxy で代替）
4. GitHub API で PR をマージ（GitHub Actions の自動マージが動かない場合）

### PR 作成時の注意

- `gh pr create` が認証エラーになる場合は、以下のように `curl` + `GLOBAL_AGENT_HTTP_PROXY` で GitHub API を直接呼ぶ：

```bash
TOKEN="<GitHub PAT>"
PROXY_URL=$(printenv GLOBAL_AGENT_HTTP_PROXY)

# PR 作成
curl -s --proxy "$PROXY_URL" \
  -X POST \
  -H "Authorization: token $TOKEN" \
  -H "Accept: application/vnd.github+json" \
  "https://api.github.com/repos/Nishi-Taiga/schedule_automation/pulls" \
  -d '{"title":"...", "head":"<branch>", "base":"main", "body":"..."}'

# PR マージ
curl -s --proxy "$PROXY_URL" \
  -X PUT \
  -H "Authorization: token $TOKEN" \
  -H "Accept: application/vnd.github+json" \
  "https://api.github.com/repos/Nishi-Taiga/schedule_automation/pulls/<PR番号>/merge" \
  -d '{"merge_method":"merge"}'
```

- 「A pull request already exists」エラーが出た場合は、既存 PR を確認してそちらをマージする

## プロジェクト構成

```
schedule_automation/
├── app.py                  # Flask バックエンド（メインロジック）
├── templates/
│   └── index.html          # シングルページ UI（HTML/CSS/JS 一体型）
├── static/                 # 静的ファイル
├── requirements.txt        # Python 依存パッケージ
├── README.md               # ユーザー向け説明
├── CLAUDE.md               # 開発ガイド（このファイル）
├── Procfile                # デプロイ用
└── render.yaml             # Render.com 設定
```

## アーキテクチャのポイント

### バックエンド (app.py)

- **Flask** ベースの Web アプリ
- セッション管理: ディスクベース（`/tmp` 配下にセッションディレクトリ）
- 主要なデータフロー:
  1. `/api/upload` — Excel ファイル（元シート・ブース表）のアップロード
  2. `/api/teachers` — ブース表から講師一覧・ブース希望を取得
  3. `/api/generate` — スケジュール自動生成
  4. `/api/download` — 結果 Excel ダウンロード
  5. `/api/save` — 手動編集の保存
  6. `/api/cloud_save` — スケジュール状態をSupabaseに永続保存 (自動/手動)
  7. `/api/cloud_list` — 保存済みスナップショット一覧取得
  8. `/api/cloud_load` — スナップショットからセッション復元
  9. `/api/cloud_delete` — スナップショット削除
  10. `/api/upload_booth_template` — 結果画面からブース表テンプレートを（再）アップロード

### フロントエンド (templates/index.html)

- シングルページアプリケーション（SPA 風）
- 3 ステップ構成: アップロード → 設定 → 結果
- 結果画面: 週タブ切替、未配置一覧、生徒別カレンダー
- ドラッグ&ドロップでの手動編集対応

### ブース表 Excel の構造

**元シート (src):**
- 各シートが 1 週間分の出勤講師データ
- シート数 = 週数（4〜6 週に動的対応済み）

**ブース表 (booth):**
- 週データシート（第1週、第2週...）
- メタデータシート:「必要コマ数」「一覧表（指導可能科目）」「講師ブース希望」
- `write_excel` ではメタシートを `['必要コマ', '一覧', 'ブース希望']` キーワードで識別して除外

## 開発時の注意事項

### 週数の動的対応

- 週数は **元シート (src) のシート数** から動的に決定される
- `range(4)` のようなハードコードは禁止。必ず `len(weekly_teachers)` や `R.schedule.length` を使用すること
- ブース表の `write_excel` ではメタシートを除外して週シートを特定する

### フロントエンド・バックエンド間のデータ整合性

- `/api/teachers` はブース表アップロード後に呼ばれる（`onF` 関数内）
- **「次へ」ボタンは `/api/teachers` のレスポンス完了後に有効化すること**（競合状態防止）
- `BP`（ブース希望）は API レスポンスで更新され、設定画面遷移時に `renderBP()` で描画される

### Excel 入出力

- openpyxl を使用
- 結合セルへの書き込みは `try/except` で保護（結合セルの非左上セルへの書き込みは例外を出す）
- レイアウト定数 (`LAYOUT`, `DAY_COLS`, `SRC_TIME_SLOTS` 等) は app.py 上部で定義

### コード品質

- app.py の関数は責務ごとに分離:
  - `load_*` 系: Excel からのデータ読み込み
  - `build_schedule`: スケジュール生成（Phase1: 固定授業 → Phase2: 通常配置）
  - `write_excel`: 結果の Excel 出力
- フロントエンドは index.html に HTML/CSS/JS をインライン記述（分離不要）
- 変数名の慣例: `wi`=週インデックス, `ts`=時間帯短縮名, `bi`=ブースインデックス

### 学習システム (v0.12.0+)

- **概要**: 手動編集パターンを学習し、`find_slot()`のスコアリング重みを自動調整
- **データフロー**: 生成→スナップショット保存→手動編集→Excelダウンロード時に差分比較→重み調整
- **永続化**: Supabase (`qchdlmmfpkhkkqunziqb`) の `schedule_learning_data` / `schedule_edit_history` テーブル
- **環境変数**: `SUPABASE_URL`, `SUPABASE_SERVICE_KEY` (Render.comで設定)
- **重み**: `DEFAULT_WEIGHTS` (app.py) に定義。`WEIGHT_BOUNDS` で上下限制約
- **過学習防止**: EMA(alpha=0.3)、最低3セッション、1回最大10%変動、BOUNDS制約
- **API**:
  - `POST /api/submit_feedback` — 差分計算＋重み調整
  - `GET /api/learning_stats` — 学習状況取得
  - `POST /api/reset_learning` — 学習データリセット

### クラウド保存 & 途中から再開 (v0.16.0+)

- **概要**: スケジュール状態をSupabaseに自動保存し、クラウドから再開可能にする
- **テーブル**: `schedule_snapshots` (EdBrioテーブルとは独立)
- **キー**: `year + month + label` で一意。自動保存は `label='latest'`
- **自動保存タイミング**: スケジュール生成後、手動編集後(3秒デバウンス)
- **ヘルパー**: `_build_state_json(sd)` — セッションデータからJSON状態を構築（`download_json` と共用）
- **ブース表テンプレート保存 (v0.17.0+)**:
  - `schedule_snapshots.booth_template` (TEXT) にブース表Excelをbase64エンコードして保存
  - 初回保存（生成後）＋手動クラウド保存時のみ送信（`include_template=true`）
  - 自動保存（3秒デバウンス）では `booth_template` を省略 → PostgRESTの `merge-duplicates` で既存値保持
  - `cloud_load` 時にbase64デコード → セッションディレクトリに復元 → `write_excel` がフォーマット付き出力可能
  - 結果画面の「ファイル更新」パネルおよび復元後セクションからブース表テンプレートを再アップロード可能
  - `/api/upload_booth_template` — アップロード + 同年月の全スナップショットのテンプレートをPATCH更新
- **UI構成 (v0.16.0)**:
  - **メイン**: 「クラウドから再開」— 目立つカードとして上部に配置
  - **詳細オプション**: Excel・JSONからの再開は `<details>` 折りたたみに格納
  - `#postRestoreSection` は `<details>` の外に配置（閉じた状態でも表示可能）
- **用語ルール**: UIでは「復元」ではなく「途中から再開」「再開」を使用すること。バックエンドAPIの内部名称（`cloud_load`, `restore_json` 等）は変更不要

### よくあるトラブルと対処

| 症状 | 原因 | 対処 |
|---|---|---|
| 設定画面にブース希望が表示されない | `/api/teachers` 完了前に画面遷移 | ボタン有効化を API 完了後に |
| 5 週目以降が無視される | `range(4)` のハードコード | 動的に週数を取得 |
| Excel 書き込みエラー | 結合セルへの書き込み | `try/except` で保護 |
| `gh pr create` が認証エラー | ローカルプロキシ環境 | `curl` + proxy で代替 |
| PR 作成時「already exists」 | 同一ブランチに既存 PR あり | 既存 PR を確認・マージ |

## テストルール

### デプロイ前（ローカル）

- **変更をコミット・プッシュする前に、必ずアプリを起動してエラーがないことを確認すること**
- テスト手順:
  1. `python app.py` でローカル起動
  2. コンソールにエラーが出ないことを確認
  3. `http://localhost:5000` にアクセスできることを確認
  4. 確認後 `Ctrl+C` で停止してからコミット・プッシュ

### デプロイ後（本番環境） ★必須 — Playwright E2E

- **コミット・プッシュ後、デプロイ環境で Playwright E2E テストが成功するまで作業を終了しない**
- `.claude/skills/deploy-test/SKILL.md` に従ってテストを実施すること
- テスト手順:
  1. バージョン番号でデプロイ反映を確認（Render free tier は反映に5〜15分かかる場合あり）
  2. Playwright でデプロイ環境 (`https://schedule-automation-386m.onrender.com/`) に対してブラウザテストを実行
  3. パスワード `2321` でログイン → 実データでのファイルアップロード → スケジュール生成 → 結果確認
  4. テスト失敗時は原因特定→修正→再push→再テストを成功するまで繰り返す
- テスト用ファイル: `G:\マイドライブ\塾\糀谷\共有用フォルダ_西\files\`
- 日本語ファイル名はASCII名の一時ファイルにコピーして回避

## セットアップ

セッション開始時に gh CLI が未インストールの場合はインストールすること：

```bash
if ! command -v gh &> /dev/null; then
  GH_VERSION="2.63.2"
  curl -fsSL "https://github.com/cli/cli/releases/download/v${GH_VERSION}/gh_${GH_VERSION}_linux_amd64.tar.gz" -o /tmp/gh.tar.gz
  tar -xzf /tmp/gh.tar.gz -C /tmp
  cp /tmp/gh_${GH_VERSION}_linux_amd64/bin/gh /usr/local/bin/gh
fi
```

gh の認証は `GH_TOKEN` 環境変数が設定されていれば自動で行われる。

## マルチエージェント協調ルール

複数エージェントが同時に修正作業を行う場合、以下のルールに従うこと。

### ゾーンベース・オーナーシップ

`app.py` と `index.html` はモノリスのため、ゾーン単位で排他制御する。

**app.py ゾーン:**

| Zone ID | 行範囲(目安) | 内容 |
|---------|-------------|------|
| `BE-SESSION` | 1-242 | セッション管理・認証 |
| `BE-CONSTANTS` | 243-337 | 定数・名前マッピング |
| `BE-SUPABASE` | 338-465 | Supabaseヘルパー |
| `BE-LEARNING` | 466-665 | 学習シグナル・重み調整 |
| `BE-PARSERS` | 667-953 | Excel/データパーサー |
| `BE-SURVEY` | 954-1489 | アンケート・講師選定 |
| `BE-SCHEDULER` | 1490-1810 | コアスケジューラー |
| `BE-EXCEL-OUT` | 1811-1993 | Excel出力 |
| `BE-API-CORE` | 1994-2594 | メインAPI |
| `BE-API-CLOUD` | 2595-2893 | クラウド保存/復元API |
| `BE-API-RESTORE` | 2894-3310 | Excel/JSON復元・メタ更新 |
| `BE-API-CHECK` | 3311-3593 | スケジュール検証 |
| `BE-API-JSON` | 3595-3819 | JSON復元API |
| `BE-API-STATE` | 3820-4071 | State永続化・学習FB |

**index.html ゾーン:**

| Zone ID | 行範囲(目安) | 内容 |
|---------|-------------|------|
| `FE-CSS` | 9-1361 | CSSスタイル |
| `FE-HTML` | 1362-1569 | HTML構造 |
| `FE-JS-INFRA` | 1570-1604 | ユーティリティJS |
| `FE-JS-CLOUD` | 1605-1677 | クラウドUI |
| `FE-JS-RESTORE` | 1678-1827 | 復元機能JS |
| `FE-JS-UPLOAD` | 1828-1991 | アップロードJS |
| `FE-JS-SETTINGS` | 1992-2068 | 設定画面JS |
| `FE-JS-GENERATE` | 2069-2167 | 生成・DL・学習JS |
| `FE-JS-CHECK` | 2168-2310 | チェックモーダルJS |
| `FE-JS-DND` | 2311-3356 | D&D・結果描画JS |

**排他ルール:**
- エージェントは作業開始前に担当ゾーンを宣言する
- **同一ゾーンを2エージェントが同時編集してはならない**
- 読み取り専用アクセスは常に許可
- 1つの機能が複数ゾーンにまたがる場合、1エージェントが全関連ゾーンを担当する

### 並列実行の判定

| 条件 | 判定 |
|------|------|
| 異なるファイルを編集 | ✅ 並列可能（worktree使用） |
| コード編集 + レビュー（read-only） | ✅ 並列可能 |
| 同じファイルの異なるゾーン | ⚠️ worktree使用で並列可能だが要注意 |
| 同じファイルの同じゾーン | ❌ 逐次実行必須 |
| app.py + index.html 両方を変更（フルスタック） | ❌ 1エージェントが担当 |

### マージ順序

```
1. バックエンド単独変更 (app.py) を先にマージ
2. フロントエンド単独変更 (index.html) を次にマージ
3. フルスタック変更 (両ファイル) を最後にマージ
4. e2e_test.py 更新は全コード変更後
```

- 後続エージェントは必ず `rebase` してからマージ（force-merge禁止）
- マージ後は `python app.py` で起動確認必須

### バージョン管理

- 個別のフィーチャーブランチではバージョンを変更しない
- **全変更がマージ完了後、最終コミットでバージョンを一括更新する**
- 更新対象: `templates/index.html`、`e2e_test.py`、`CLAUDE.md`

### テスト協調

- 各エージェントはマージ前に `python app.py` でローカルスモークテスト必須
- デプロイ先は1つなので、テストはキュー管理する（変更A マージ→デプロイ→E2E→変更B…）
- デプロイ待ち時間（5-15分）は他エージェントがレビュー・計画を行う

### 作業開始テンプレート

```
STEP 1: 変更を分類（各変更の対象ゾーンを列挙）
STEP 2: 並列性を判定（上記の判定表に照らす）
STEP 3: 実行順序を決定
  Phase 1（並列可能な変更を同時実行）
  Phase 2（依存関係のある変更を逐次実行）
  Phase 3: バージョン更新 → E2Eテスト → デプロイ確認
```
