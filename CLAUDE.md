# Claude Code 開発ガイド

## バージョン管理

- 現在のバージョン: **v0.5.9**
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

### よくあるトラブルと対処

| 症状 | 原因 | 対処 |
|---|---|---|
| 設定画面にブース希望が表示されない | `/api/teachers` 完了前に画面遷移 | ボタン有効化を API 完了後に |
| 5 週目以降が無視される | `range(4)` のハードコード | 動的に週数を取得 |
| Excel 書き込みエラー | 結合セルへの書き込み | `try/except` で保護 |
| `gh pr create` が認証エラー | ローカルプロキシ環境 | `curl` + proxy で代替 |
| PR 作成時「already exists」 | 同一ブランチに既存 PR あり | 既存 PR を確認・マージ |

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
