# マツリカちゃん Streamlitアプリケーション デプロイガイド

## Streamlit Community Cloudへのデプロイ手順

### 前提条件
- GitHubアカウント
- Streamlit Community Cloudアカウント（無料）

### デプロイ手順

1. **GitHubリポジトリの作成**
   - このプロジェクトをGitHubリポジトリにアップロード
   - リポジトリ名: `matsurica-streamlit-app`（任意）

2. **Streamlit Community Cloudへの接続**
   - [Streamlit Community Cloud](https://share.streamlit.io/)にアクセス
   - GitHubアカウントでログイン
   - "New app"ボタンをクリック

3. **アプリ設定**
   - Repository: 作成したGitHubリポジトリを選択
   - Branch: `main`（または適切なブランチ）
   - Main file path: `app.py`
   - Advanced settings:
     - Python version: 3.9以上
     - Requirements file: `requirements.txt`

4. **デプロイ実行**
   - "Deploy!"ボタンをクリック
   - デプロイが完了するまで待機（数分）

### ファイル構成
```
matsurica-streamlit-app/
├── app.py                 # メインのStreamlitアプリ
├── requirements.txt       # Python依存関係
├── setup.sh              # セットアップスクリプト（Linux用）
├── matsurica_integrated_tool.py  # 統合処理ツール
├── matsurica_gui.py       # 元のGUIアプリ（参考）
├── SegUIVar.ttf          # フォントファイル
├── マツリカちゃん統合仕様書.txt  # 仕様書
└── DEPLOYMENT_GUIDE.md   # このファイル
```

### ローカルテスト
デプロイ前にローカルでテストする場合：

```bash
# 仮想環境の作成（推奨）
python -m venv venv
source venv/bin/activate  # Linux/Mac
# または
venv\Scripts\activate     # Windows

# 依存関係のインストール
pip install -r requirements.txt

# アプリの実行
streamlit run app.py
```

### 注意事項

1. **ファイルサイズ制限**
   - Streamlit Community Cloudにはファイルサイズ制限があります
   - 大きなExcelファイルの処理には注意が必要です

2. **タイムアウト**
   - 長時間かかる処理はタイムアウトする可能性があります
   - 進捗表示を適切に実装してください

3. **セキュリティ**
   - アップロードされたファイルは一時的に処理されます
   - 機密データの取り扱いに注意してください

### トラブルシューティング

**デプロイエラーが発生した場合：**
1. requirements.txtの依存関係を確認
2. Pythonバージョンの互換性を確認
3. ログを確認してエラー内容を特定

**アプリが動作しない場合：**
1. ローカル環境でテスト
2. 必要なファイルがすべてリポジトリに含まれているか確認
3. ファイルパスを確認

### カスタムドメイン（オプション）
カスタムドメインを設定する場合：
1. Streamlit Community Cloudの設定からカスタムドメインを追加
2. DNS設定を更新

### バージョン管理
- 定期的に依存関係の更新を確認
- Streamlitの新バージョンに対応
- セキュリティアップデートを適用

## サポート
問題が発生した場合は以下に連絡：
- メール: tomoharu.kobayashi@konicaminolta.com
- エラーログを添付して報告
