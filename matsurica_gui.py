# -*- coding: utf-8 -*-
# 最終更新: 2025-11-17 14:25 (Codexによる追記)
"""
マツリカGUIツール
グラフィカルユーザーインターフェースでマツリカ変換を簡単に実行
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import subprocess
import sys
from pathlib import Path
import os
import logging

GUI_VERSION = "matsurica_gui.py v2025.10.22-02"
try:
    from matsurica_integrated_tool import TOOL_VERSION as INTEGRATED_TOOL_VERSION
except Exception:
    INTEGRATED_TOOL_VERSION = "不明（モジュール未読込）"

class MatsuricaGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("アプリ版魔界大帝マツリカ・マツリちゃん")
        # ウィンドウサイズと最小サイズを1000x300に設定（横4列レイアウト）
        self.root.geometry("1000x300")
        self.root.minsize(900, 250)
        self.root.resizable(True, True)  # サイズ変更可能
        
        # ウィンドウを画面中央に配置
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # 基準サイズと設定
        self.base_button_size = 268  # 基準ボタンサイズ
        self.base_font_size = 16     # 基準フォントサイズ
        self.margin = 24             # 外周余白
        self.gap = 16                # ボタン間ギャップ
        
        # 色設定
        self.MARINE = "#0066CC"   # マリンブルー
        self.BORDER = "#2E7AB8"   # 枠線色
        
        # clamテーマを使用（背景色変更を有効化）
        style = ttk.Style()
        style.theme_use("clam")
        
        # Windows 11風のフォント設定
        self.setup_fonts()
        
        # リサイズイベントのバインド
        self.root.bind("<Configure>", self.on_resize)
        
    def setup_fonts(self):
        """Windows 11風のフォント設定"""
        # Segoe UI Variableフォントを設定
        default_font = ("Segoe UI Variable", 9)
        title_font = ("Segoe UI Variable", 16, "bold")
        
        # 全ウィジェットに適用
        self.root.option_add("*Font", default_font)
        self.root.option_add("*Label.Font", default_font)
        self.root.option_add("*Button.Font", default_font)
        self.root.option_add("*Entry.Font", default_font)
        self.root.option_add("*Text.Font", default_font)
        
        # タイトル用フォント
        self.title_font = title_font
        
        # 変数の初期化
        self.excel_file = tk.StringVar()
        self.customers_file = tk.StringVar(value="顧客リスト.csv")
        self.output_folder = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))
        
        # EXE実行時は実行ファイルのあるディレクトリを基準にする
        if getattr(sys, 'frozen', False):
            # EXE実行時
            base_dir = os.path.dirname(sys.executable)
            self.customers_file = tk.StringVar(value=os.path.join(base_dir, "顧客リスト.csv"))
            self.output_folder = tk.StringVar(value=base_dir)
        else:
            # スクリプト実行時
            self.customers_file = tk.StringVar(value="顧客リスト.csv")
            self.output_folder = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))
        
        # 状態フラグ
        self.input_ready = False
        self.list_ready = False
        self.output_ready = True  # 出力はデフォルトで準備OK
        
        self.setup_ui()
        
    def setup_ui(self):
        # メインフレーム - 1行4列のグリッドレイアウト
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), 
                       padx=self.margin, pady=self.margin)
        
        # 1行4列のグリッド設定
        main_frame.columnconfigure(0, weight=1, uniform="button_col")
        main_frame.columnconfigure(1, weight=1, uniform="button_col")
        main_frame.columnconfigure(2, weight=1, uniform="button_col")
        main_frame.columnconfigure(3, weight=1, uniform="button_col")
        main_frame.rowconfigure(0, weight=1, uniform="button_row")
        
        # ボタン作成
        self.input_btn = ttk.Button(main_frame, text="入力", command=self.browse_excel)
        self.list_btn = ttk.Button(main_frame, text="リスト", command=self.browse_customers)
        self.output_btn = ttk.Button(main_frame, text="出力", command=self.browse_output_folder)
        self.run_button = ttk.Button(main_frame, text="実行", command=self.run_conversion)
        
        # ボタンをグリッドに配置（横4列）
        self.input_btn.grid(row=0, column=0, padx=self.gap//2, pady=self.gap//2, sticky="nsew")
        self.list_btn.grid(row=0, column=1, padx=self.gap//2, pady=self.gap//2, sticky="nsew")
        self.output_btn.grid(row=0, column=2, padx=self.gap//2, pady=self.gap//2, sticky="nsew")
        self.run_button.grid(row=0, column=3, padx=self.gap//2, pady=self.gap//2, sticky="nsew")
        
        # ツールチップの設定
        self.setup_tooltips(self.input_btn, "活動Excelファイルを選択します")
        self.setup_tooltips(self.list_btn, "顧客リストファイルを選択します")
        self.setup_tooltips(self.output_btn, "出力フォルダを選択します")
        self.setup_tooltips(self.run_button, "変換処理を実行します")
        
        # ファイルパス表示エリア（非表示）
        self.file_info_text = tk.Text(main_frame, height=8, width=70, state='disabled')
        self.file_info_text.grid_remove()  # 非表示にする
        
        # ルートウィンドウのグリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 初期サイズ調整
        self.update_button_sizes()
        
        # フラットデザインスタイルの適用
        self.apply_flat_style()
        
        # ファイル情報の初期表示
        self.update_file_info()
        
    def apply_flat_style(self):
        """フラットデザインスタイルを適用"""
        # フラットスタイルの設定
        style = ttk.Style()
        
        # 標準ボタンスタイル（未準備）
        style.configure("Step.TButton",
            background="white",
            foreground="black",
            bordercolor=self.BORDER,
            relief="flat",
            borderwidth=1,
            padding=16,
            font=("Segoe UI Variable", 12, "bold"),
            anchor="center"  # テキストを中央揃え
        )
        
        # 準備OKボタンスタイル（マリンブルー）
        style.configure("Ready.TButton",
            background=self.MARINE,
            foreground="white",
            bordercolor=self.BORDER,
            relief="flat",
            borderwidth=1,
            padding=16,
            font=("Segoe UI Variable", 12, "bold"),
            anchor="center"  # テキストを中央揃え
        )
        
        # 無効化ボタンスタイル
        style.configure("Disabled.TButton",
            background="#cccccc",
            foreground="#666666",
            bordercolor=self.BORDER,
            relief="flat",
            borderwidth=1,
            padding=16,
            font=("Segoe UI Variable", 12, "bold"),
            anchor="center"  # テキストを中央揃え
        )
        
        # ホバー/押下時のスタイルマップ
        style.map("Step.TButton",
            background=[("active", "#f0f0f0"), ("pressed", "#e0e0e0")],
            relief=[("pressed", "sunken"), ("!pressed", "flat")]
        )
        
        style.map("Ready.TButton",
            background=[("active", "#1E7FE0"), ("pressed", "#005BB8")],
            relief=[("pressed", "sunken"), ("!pressed", "flat")]
        )
        
        # スタイルを適用
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Button):
                        if child["text"] == "実行":
                            # 実行ボタンは初期状態で無効化
                            child.configure(style="Disabled.TButton")
                        else:
                            child.configure(style="Step.TButton")
        
        # 初期状態で出力ボタンをマリンブルーに設定
        self.update_button_states()
        
    def update_button_sizes(self):
        """ボタンサイズを更新"""
        # 利用可能な領域を計算
        available_width = self.root.winfo_width() - 2 * self.margin - 3 * self.gap
        available_height = self.root.winfo_height() - 2 * self.margin
        
        # 横4列なので幅を4分割
        button_width = available_width // 4
        button_height = available_height
        
        # 正方形サイズを計算（短辺に合わせる）
        button_size = min(button_width, button_height)
        
        # フォントサイズを計算 (max(12pt, 辺×0.065))
        font_size = max(12, int(button_size * 0.065))
        
        # スタイルを更新
        style = ttk.Style()
        
        # 標準ボタンスタイル（未準備）
        style.configure("Step.TButton",
            background="white",
            foreground="black",
            bordercolor=self.BORDER,
            relief="flat",
            borderwidth=1,
            padding=16,
            font=("Segoe UI Variable", font_size, "bold"),
            anchor="center"  # テキストを中央揃え
        )
        
        # 準備OKボタンスタイル（マリンブルー）
        style.configure("Ready.TButton",
            background=self.MARINE,
            foreground="white",
            bordercolor=self.BORDER,
            relief="flat",
            borderwidth=1,
            padding=16,
            font=("Segoe UI Variable", font_size, "bold"),
            anchor="center"  # テキストを中央揃え
        )
        
        # 無効化ボタンスタイル
        style.configure("Disabled.TButton",
            background="#cccccc",
            foreground="#666666",
            bordercolor=self.BORDER,
            relief="flat",
            borderwidth=1,
            padding=16,
            font=("Segoe UI Variable", font_size, "bold"),
            anchor="center"  # テキストを中央揃え
        )
        
        # ボタンの状態を更新
        self.update_button_states()
        
    def on_resize(self, event):
        """リサイズイベントハンドラー"""
        if event.widget == self.root:
            self.update_button_sizes()
            
    def update_button_states(self):
        """ボタンの状態を更新"""
        # リストボタンの状態チェック（顧客リストファイルが存在するか）
        customers_file = self.customers_file.get()
        if customers_file and os.path.exists(customers_file):
            self.list_ready = True
        else:
            # デフォルトの顧客リストファイルが存在するかチェック
            default_customers = os.path.join(os.path.dirname(os.path.abspath(__file__)), "顧客リスト.csv")
            self.list_ready = os.path.exists(default_customers)
        
        # 入力ボタンの状態チェック
        self.input_ready = bool(self.excel_file.get() and os.path.exists(self.excel_file.get()))
        
        # 出力ボタンの状態（常に準備OK）
        self.output_ready = True
        
        # 実行準備が完了したかチェック
        all_ready = self.input_ready and self.list_ready and self.output_ready
        
        # ボタンのスタイルを更新
        buttons = [self.input_btn, self.list_btn, self.output_btn, self.run_button]
        texts = ["入力", "リスト", "出力", "実行"]
        ready_states = [self.input_ready, self.list_ready, self.output_ready, all_ready]
        
        for btn, text, ready in zip(buttons, texts, ready_states):
            if all_ready and text != "実行":
                # 実行準備OK時は実行ボタン以外をグレーにする
                btn.configure(text=f"{text}\n準備OK", style="Disabled.TButton")
            elif ready:
                btn.configure(text=f"{text}\n準備OK", style="Ready.TButton")
            else:
                btn.configure(text=text, style="Step.TButton")
            
            # 実行ボタンの状態を特別に設定
            if text == "実行":
                if ready:
                    btn.configure(state="normal", style="Ready.TButton")
                else:
                    btn.configure(state="disabled", style="Disabled.TButton")
        
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="活動Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            self.update_file_info()
            
    def browse_customers(self):
        filename = filedialog.askopenfilename(
            title="顧客リストファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.customers_file.set(filename)
            self.update_file_info()
            
            
    def browse_output_folder(self):
        # EXE実行時は実行ファイルのあるディレクトリを初期ディレクトリとして設定
        if getattr(sys, 'frozen', False):
            initial_dir = os.path.dirname(sys.executable)
        else:
            initial_dir = os.path.dirname(os.path.abspath(__file__))
            
        folder = filedialog.askdirectory(
            title="出力フォルダを選択",
            initialdir=initial_dir
        )
        if folder:
            self.output_folder.set(folder)
            self.update_file_info()
            
    def setup_tooltips(self, widget, text):
        """簡易的なツールチップを設定"""
        def on_enter(event):
            tooltip = tk.Toplevel(self.root)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = tk.Label(tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()
            widget.tooltip = tooltip
            
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                delattr(widget, 'tooltip')
                
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
            
    def update_file_info(self):
        """ファイル情報を更新"""
        info = ""
        if self.excel_file.get():
            info += f"入力ファイル: {self.excel_file.get()}\n"
        if self.customers_file.get():
            info += f"顧客リスト: {self.customers_file.get()}\n"
        if self.output_folder.get():
            info += f"出力フォルダ: {self.output_folder.get()}\n"
            info += f"出力ファイル: {os.path.join(self.output_folder.get(), 'customer_action_import_format.csv')}\n"
            
        self.file_info_text.config(state='normal')
        self.file_info_text.delete(1.0, tk.END)
        self.file_info_text.insert(1.0, info)
        self.file_info_text.config(state='disabled')
        
        # ボタンの状態も更新
        self.update_button_states()
            
    def log_message(self, message):
        """ログにメッセージを追加"""
        # ログ表示エリアがない場合はコンソールに出力
        print(message)
        
        # ログファイルにも出力
        try:
            if getattr(sys, 'frozen', False):
                # EXE実行時は実行ファイルのあるディレクトリにログファイルを作成
                log_dir = os.path.dirname(sys.executable)
                log_file = os.path.join(log_dir, "matsurica_conversion.log")
            else:
                # スクリプト実行時はスクリプトのあるディレクトリにログファイルを作成
                log_dir = os.path.dirname(os.path.abspath(__file__))
                log_file = os.path.join(log_dir, "matsurica_conversion.log")
            
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"{message}\n")
        except Exception as e:
            print(f"ログファイル書き込みエラー: {e}")
        
        self.root.update_idletasks()
        
    def run_conversion(self):
        """変換処理を実行"""
        # 入力チェック
        if not self.excel_file.get():
            messagebox.showerror("エラー", "活動Excelファイルを指定してください")
            return
            
        if not Path(self.excel_file.get()).exists():
            messagebox.showerror("エラー", "指定された活動Excelファイルが存在しません")
            return
            
        # ボタンを無効化
        self.run_button.config(state='disabled')
        
        # 別スレッドで処理を実行
        thread = threading.Thread(target=self.execute_conversion)
        thread.daemon = True
        thread.start()
        
    def show_log_window(self):
        """ログウィンドウを表示"""
        log_window = tk.Toplevel(self.root)
        log_window.title("処理ログ")
        log_window.geometry("800x400")
        
        # ログ表示エリア
        log_text = tk.Text(log_window, wrap=tk.WORD)
        log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(log_window, orient=tk.VERTICAL, command=log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        log_text.configure(yscrollcommand=scrollbar.set)
        
        # ログ内容を表示（簡易的な実装）
        log_text.insert(tk.END, "ログ機能は現在開発中です。\n")
        log_text.insert(tk.END, "処理中のログはコンソールに出力されます。\n")
        log_text.config(state='disabled')
            
    def execute_conversion(self):
        """実際の変換処理を実行"""
        try:
            self.log_message("=== マツリカ変換処理開始 ===")
            self.log_message(f"GUIバージョン: {GUI_VERSION}")
            self.log_message(f"統合ツールバージョン: {INTEGRATED_TOOL_VERSION}")
            self.log_message(f"入力ファイル: {self.excel_file.get()}")
            self.log_message(f"顧客リスト: {self.customers_file.get()}")
            self.log_message(f"出力フォルダ: {self.output_folder.get()}")
            
            # コマンドライン引数の構築（相対パスを絶対パスに変換）
            customers_path = self.customers_file.get()
            output_folder = self.output_folder.get()
            
            # 相対パスの場合は絶対パスに変換
            if not os.path.isabs(customers_path):
                # EXE実行時とスクリプト実行時でパス処理を分ける
                if getattr(sys, 'frozen', False):
                    # EXE実行時は実行ファイルのあるディレクトリを基準
                    base_dir = os.path.dirname(sys.executable)
                    customers_path = os.path.join(base_dir, customers_path)
                else:
                    # スクリプト実行時はスクリプトのあるディレクトリを基準
                    customers_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), customers_path)
            
            # 出力ファイルパスを固定ファイル名で生成
            output_path = os.path.join(output_folder, "customer_action_import_format.csv")
            self.log_message(f"出力ファイル: {output_path}")
            
            # EXE実行時は統合ツールを直接インポートして実行
            if getattr(sys, 'frozen', False):
                # EXE実行時は直接関数を呼び出す
                try:
                    self.log_message("EXE実行モード: 統合ツールを直接インポート")
                    # 統合ツールの関数を直接インポートして実行
                    from matsurica_integrated_tool import main as integrated_main
                    import argparse
                    
                    # 引数を構築して直接実行
                    args = argparse.Namespace()
                    args.input_excel = self.excel_file.get()
                    args.customers = customers_path
                    args.output = output_path
                    args.sample = None  # サンプルオプションは使用しない
                    
                    self.log_message("統合ツールのmain関数を実行中...")
                    # 統合ツールのmain関数を直接実行
                    integrated_main(args)
                    
                    self.log_message("=== 処理が正常に完了しました ===")
                    self.log_message(f"出力ファイル: {output_path}")
                    messagebox.showinfo("成功", "マツリカ変換が完了しました")
                    
                except Exception as e:
                    self.log_message(f"統合ツール実行エラー: {str(e)}")
                    self.log_message(f"エラータイプ: {type(e).__name__}")
                    import traceback
                    error_traceback = traceback.format_exc()
                    self.log_message(f"エラー詳細:\n{error_traceback}")
                    messagebox.showerror("エラー", f"変換処理中にエラーが発生しました: {str(e)}")
                    
            else:
                # スクリプト実行時はサブプロセスで実行
                cmd = [
                    sys.executable, "matsurica_integrated_tool.py",
                    self.excel_file.get(),
                    "--customers", customers_path,
                    "--output", output_path
                ]
                
                self.log_message(f"実行コマンド: {' '.join(cmd)}")
                
                # カレントディレクトリを設定
                current_dir = os.path.dirname(os.path.abspath(__file__))
                self.log_message(f"カレントディレクトリ: {current_dir}")
                
                # サブプロセス実行
                process = subprocess.Popen(
                    cmd,
                    cwd=current_dir,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding='cp932',  # cp932エンコーディングを使用
                    errors='replace'
                )
                
                self.log_message("サブプロセス実行中...")
                # 出力をリアルタイムで表示
                for line in process.stdout:
                    self.log_message(f"サブプロセス出力: {line.strip()}")
                    
                # プロセス終了を待機
                return_code = process.wait()
                self.log_message(f"サブプロセス終了コード: {return_code}")
                
                if return_code == 0:
                    self.log_message("=== 処理が正常に完了しました ===")
                    self.log_message(f"出力ファイル: {output_path}")
                    messagebox.showinfo("成功", "マツリカ変換が完了しました")
                else:
                    self.log_message("=== 処理中にエラーが発生しました ===")
                    messagebox.showerror("エラー", "変換処理中にエラーが発生しました")
                
        except Exception as e:
            self.log_message(f"予期せぬエラー: {str(e)}")
            self.log_message(f"エラータイプ: {type(e).__name__}")
            import traceback
            error_traceback = traceback.format_exc()
            self.log_message(f"エラー詳細:\n{error_traceback}")
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました: {str(e)}")
            
        finally:
            # UIを元に戻す
            self.run_button.config(state='normal')
            self.log_message("=== 処理終了 ===")

def main(event=None):
    """メイン関数 - TkinterのイベントやPyQtのchecked引数に対応"""
    # DPI対応設定（高DPIディスプレイでのぼやけ防止）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)  # プロセスDPI対応を有効化
    except:
        pass  # Windows以外や古いバージョンでは無視
    
    root = tk.Tk()
    app = MatsuricaGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
