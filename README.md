# ブース表自動生成 (Cloud版)

## デプロイ手順 (Render)

### 1. GitHubリポジトリを作成
```bash
cd booth_cloud
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR_USER/booth-scheduler.git
git push -u origin main
```

### 2. Renderでデプロイ
1. https://render.com にログイン（GitHubアカウントで登録可）
2. 「New +」→「Web Service」
3. GitHubリポジトリを選択
4. 設定:
   - **Name**: booth-scheduler
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Instance Type**: Free
5. 環境変数を設定:
   - `APP_PASSWORD`: 塾内で共有するパスワード
   - `SECRET_KEY`: 「Generate」ボタンで自動生成
6. 「Create Web Service」

### 3. アクセス
- `https://booth-scheduler.onrender.com` のようなURLが割り当てられます
- パスワードを入力してログイン

## 注意事項
- 無料プラン: 15分間アクセスがないとスリープ（再アクセスで30秒程度で復帰）
- ファイルは一時保存（サーバー再起動で消去）
- アップロード上限: 10MB
