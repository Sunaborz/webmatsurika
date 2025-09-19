#!/bin/bash
# Streamlit Community Cloud用セットアップスクリプト

echo "マツリカちゃん Streamlitアプリケーションのセットアップを開始します..."

# 必要なディレクトリを作成
mkdir -p .streamlit

# Streamlit設定ファイルを作成
cat > .streamlit/config.toml << EOF
[server]
headless = true
enableCORS = false
enableXsrfProtection = false

[browser]
serverAddress = "0.0.0.0"
serverPort = 8501

[theme]
base = "light"
EOF

echo "セットアップが完了しました！"
echo "以下のコマンドでローカルテストができます:"
echo "streamlit run app.py"
