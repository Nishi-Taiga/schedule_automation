# Claude Code 開発ガイド

## マージ手順

変更を main にマージする際は、以下の手順で行うこと：

1. フィーチャーブランチにコミット＆プッシュ
2. `gh pr create` で PR を作成
3. `gh pr merge --merge` で main にマージ

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
