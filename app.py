# -*- coding: utf-8 -*-
"""
マツリカちゃん Streamlit Webアプリケーション
Streamlit Community Cloud用のWebインターフェース
"""

import streamlit as st
import pandas as pd
import subprocess
import sys
from pathlib import Path
import os
import tempfile
import shutil
from datetime import datetime

# ページ設定
st.set_page_config(
    page_title="アプリ版魔界大帝マツリカ・マツリちゃんV3",
    page_icon="👑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# カスタムCSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #0066CC;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # ヘッダー
    st.markdown('<h1 class="main-header">👑 アプリ版魔界大帝マツリカ・マツリちゃん v3 </h1>', unsafe_allow_html=True)
    
    # サイドバー
    with st.sidebar:
        st.header("設定")
        st.info("週報Excelからマツリカ取込用CSVを生成するのじゃ")
        
        # ファイルアップロード
        uploaded_excel = st.file_uploader(
            "活動Excelファイルを選択のじゃ",
            type=['xlsx', 'xls'],
            help="活動データが含まれるExcelファイルを選択してください"
        )
        
        uploaded_customers = st.file_uploader(
            "顧客リストCSVを選択のじゃ",
            type=['csv'],
            help="顧客リストのCSVファイルを選択してください"
        )
        
        output_filename = st.text_input(
            "出力ファイル名",
            value="customer_action_import_format.csv",
            help="生成するCSVファイルの名前"
        )
    
    # メインコンテンツ
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("ファイル情報")
        
        if uploaded_excel:
            st.success(f"✅ Excelファイル: {uploaded_excel.name}")
        else:
            st.warning("⚠️ Excelファイルが選択されていません")
            
        if uploaded_customers:
            st.success(f"✅ 顧客リスト: {uploaded_customers.name}")
        else:
            st.warning("⚠️ 顧客リストが選択されていません")
            
        st.info(f"📁 出力ファイル: {output_filename}")
    
    with col2:
        st.header("処理実行")
        
        if st.button("✨ 変換を実行するのじゃ", type="primary", disabled=not (uploaded_excel and uploaded_customers)):
            if uploaded_excel and uploaded_customers:
                process_files(uploaded_excel, uploaded_customers, output_filename)
            else:
                st.error("必要なファイルがすべて選択されていませんのじゃ")
    
    # 使い方ガイド
    st.header("使い方ガイド")
    with st.expander("詳細な使用方法を見る"):
        st.markdown("""
        ### 処理フロー
        1. **ExcelファイルをCSVに変換** - 活動データを抽出します
        2. **顧客リストとマッチング** - 企業名を照合します  
        3. **マツリカ取込用CSVを生成** - 取込用フォーマットで出力します
        
        ### 必要なファイル
        - **活動Excelファイル**: シート「明細データ」または先頭シートに活動データがあること
        - **顧客リストCSV**: Shift-JISエンコーディング、B列に「取引先名(必須)」があること
        
        ### 出力ファイル
        - 中間ファイル: `matched_activity.xlsx`（企業マッチング結果）
        - 最終出力: `customer_action_import_format.csv`（マツリカ取込用CSV）
        
        ### 注意事項
        - 大きなファイルの処理には時間がかかる場合があります
        - エラーが発生した場合はファイル形式を確認してください
        - 処理中はページを閉じないでください
        """)

def process_files(uploaded_excel, uploaded_customers, output_filename):
    """アップロードされたファイルを処理する"""
    try:
        with st.spinner("魔界の力で変換中... しばらくお待ちくだされ"):
            # 一時ディレクトリを作成
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir_path = Path(temp_dir)
                
                # アップロードされたファイルを一時ディレクトリに保存
                excel_path = temp_dir_path / uploaded_excel.name
                with open(excel_path, "wb") as f:
                    f.write(uploaded_excel.getbuffer())
                
                customers_path = temp_dir_path / "顧客リスト.csv"
                with open(customers_path, "wb") as f:
                    f.write(uploaded_customers.getbuffer())
                
                # 出力パス
                output_path = temp_dir_path / output_filename
                
                # 統合ツールを一時ディレクトリにコピー
                tool_source_path = Path("matsurica_integrated_tool.py")
                tool_dest_path = temp_dir_path / "matsurica_integrated_tool.py"
                if tool_source_path.exists():
                    shutil.copy2(tool_source_path, tool_dest_path)
                
                # 統合ツールを実行
                cmd = [
                    sys.executable, "matsurica_integrated_tool.py",
                    str(excel_path),
                    "--customers", str(customers_path),
                    "--output", str(output_path)
                ]
                
                # カレントディレクトリを一時ディレクトリに変更して実行
                result = subprocess.run(
                    cmd,
                    cwd=temp_dir_path,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace'
                )
                
                if result.returncode == 0:
                    # 成功時の処理
                    if output_path.exists():
                        # 生成されたCSVファイルを読み込み
                        df = pd.read_csv(output_path, encoding='cp932')
                        
                        # 成功メッセージ
                        st.success("✅ 変換が完了したのじゃ！")
                        
                        # 結果の表示
                        st.subheader("変換結果")
                        st.info(f"生成された行数: {len(df)}行")
                        
                        # データプレビュー
                        st.dataframe(df.head(), use_container_width=True)
                        
                        # ダウンロードボタン
                        csv_data = output_path.read_bytes()
                        st.download_button(
                            label="📥 CSVファイルをダウンロード",
                            data=csv_data,
                            file_name=output_filename,
                            mime="text/csv",
                            help="生成されたCSVファイルをダウンロードします"
                        )
                        
                        # ログ表示
                        with st.expander("処理ログを見る"):
                            st.text(result.stdout)
                    else:
                        st.error("出力ファイルが生成されませんでしたのじゃ")
                        with st.expander("エラー詳細"):
                            st.text(result.stdout)
                            st.text(result.stderr)
                else:
                    # エラー時の処理
                    st.error("❌ 変換中にエラーが発生したのじゃ")
                    with st.expander("エラー詳細"):
                        st.text(f"リターンコード: {result.returncode}")
                        st.text("標準出力:")
                        st.text(result.stdout)
                        st.text("標準エラー:")
                        st.text(result.stderr)
                        
    except Exception as e:
        st.error(f"予期せぬエラーが発生したのじゃ: {str(e)}")
        import traceback
        with st.expander("エラー詳細"):
            st.text(traceback.format_exc())

if __name__ == "__main__":
    main()

