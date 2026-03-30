import sys
import os
import io

# 1. 最初にメッセージを表示
# Windowsコンソールでの文字化けを最小限にするため、
# インポート直後にエンコーディングを設定
if sys.platform == "win32":
    try:
        # 標準出力がリダイレクトされていないか確認
        if sys.stdout and hasattr(sys.stdout, 'buffer'):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        if sys.stderr and hasattr(sys.stderr, 'buffer'):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass

print("分析・変換をしています。しばらくお待ちください...")
if sys.stdout:
    sys.stdout.flush()

# -*- coding: utf-8 -*-


# 2. 標準ライブラリのインポート
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox

def main():
    # --- PyInstaller EXE の argv を Unicode に強制統一 ---
    # Windows環境での文字化け対策
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        def get_unicode_argv():
            """WindowsのGetCommandLineWを使用してUnicodeの引数リストを取得する"""
            GetCommandLineW = ctypes.windll.kernel32.GetCommandLineW
            GetCommandLineW.restype = wintypes.LPCWSTR
            CommandLineToArgvW = ctypes.windll.shell32.CommandLineToArgvW
            CommandLineToArgvW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_int)]
            CommandLineToArgvW.restype = ctypes.POINTER(wintypes.LPWSTR)
            
            argc = ctypes.c_int(0)
            argv_unicode = CommandLineToArgvW(GetCommandLineW(), ctypes.byref(argc))
            if not argv_unicode:
                return sys.argv
            
            try:
                return [argv_unicode[i] for i in range(argc.value)]
            finally:
                ctypes.windll.kernel32.LocalFree(argv_unicode)

        sys.argv = get_unicode_argv()

    # 重いモジュールのインポート
    try:
        import pythoncom
        from converters import ExcelConverter, WordConverter
    except Exception as e:
        print(f"モジュールの読み込みに失敗しました: {e}")
        if getattr(sys, 'frozen', False):
            input("\nEnterキーを押して終了してください。")
        return

    # 引数チェック
    if len(sys.argv) < 2:
        print("\nWord、Excelファイルをこの実行ファイルにドラッグアンドドロップしてください。")
        if getattr(sys, 'frozen', False):
            input("\nEnterキーを押して終了してください。")
        return

    # 対応拡張子のチェック
    supported_extensions = [".xlsx", ".xls", ".xlsm", ".docx", ".doc"]
    files_to_process = []
    unsupported_found = False

    for file_path in sys.argv[1:]:
        if not file_path or not file_path.strip():
            continue
        
        clean_path = file_path.strip().strip('"')
        if not os.path.exists(clean_path):
            continue

        ext = os.path.splitext(clean_path)[1].lower()
        if ext in supported_extensions:
            files_to_process.append(os.path.abspath(clean_path))
        else:
            unsupported_found = True

    if unsupported_found and not files_to_process:
        print("\nエラー: Word または Excelファイルのみ対応しています。")
        if getattr(sys, 'frozen', False) or sys.stdin.isatty():
            input("\nEnterキーを押して終了してください。")
        return

    if not files_to_process:
        print("\n処理対象のファイルが見つかりませんでした。")
        if getattr(sys, 'frozen', False):
            input("\nEnterキーを押して終了してください。")
        return

    # 保存先フォルダの決定（ポップアップウィンドウ）
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    
    # IMEの制御やエンコーディングの問題を避けるため、
    # ダイアログを表示する前に一度ダミーのウィンドウでフォーカスを制御するなどの対策
    
    base_dir = os.path.dirname(files_to_process[0])
    date_str = datetime.now().strftime("%Y%m%d")
    default_name = f"{date_str}_"
    
    prompt_msg = "元のファイルと同じ場所に新たなフォルダを作ります。\nフォルダ名称を入力してください。"
    
    # simpledialogの代わりに、より安定した方法を検討
    # ここでは標準のダイアログを使用しつつ、親ウィンドウを明示
    user_input = simpledialog.askstring("フォルダ作成", prompt_msg, initialvalue=default_name, parent=root)
    
    if user_input is None:
        print("\nキャンセルされました。")
        return

    # 文字化け対策: Windows環境でTkinterからの入力が稀に文字化けする場合の考慮
    # 通常TkinterはUnicodeで返すが、環境によって不正なバイトが含まれる可能性を考慮
    try:
        folder_name = user_input.strip()
    except Exception:
        # 万が一のフォールバック
        folder_name = default_name
    
    print(f"DEBUG user_input: {repr(folder_name)}")
    if not folder_name:
        folder_name = default_name

    output_dir = os.path.join(base_dir, folder_name)

    try:
        print(f"DEBUG folder_name: {repr(folder_name)}")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"フォルダを作成しました: {folder_name}")
    except Exception as e:
        print(f"フォルダの作成に失敗しました: {e}")
        output_dir = base_dir

    success_info = []
    error_files = []

    # ファイルごとに処理
    for file_path in files_to_process:
        basename = os.path.basename(file_path)
        pdf_name = os.path.splitext(basename)[0] + ".pdf"
        pdf_path = os.path.join(output_dir, pdf_name)

        # 上書き確認
        if os.path.exists(pdf_path):
            confirm_msg = f"ファイル '{pdf_name}' は既に存在します。\n上書きしますか？"
            if not messagebox.askyesno("上書き確認", confirm_msg, parent=root):
                print(f"スキップしました: {basename}")
                continue

        print(f"PDFに変換中: {basename}")

        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        converter = None

        try:
            if ext in [".xlsx", ".xls", ".xlsm"]:
                converter = ExcelConverter(file_path)
            elif ext in [".docx", ".doc"]:
                converter = WordConverter(file_path)

            if converter:
                pdf_path = converter.convert(output_dir=output_dir)
                success_info.append({
                    'dir': output_dir,
                    'file': os.path.basename(pdf_path)
                })
        except Exception as e:
            error_message = str(e).strip()
            if not error_message:
                error_message = "不明なエラーが発生しました"
            error_files.append(f"{basename} (エラー: {error_message})")

    # 結果表示
    if success_info:
        print("\n--- 成功 ---")
        for info in success_info:
            print(f"保存先: {info['dir']}")
            print(f"ファイル: {info['file']}\n")
        
    if error_files:
        print("\n--- 失敗 ---")
        for err in error_files:
            print(err)

    if getattr(sys, 'frozen', False) or sys.stdin.isatty():
        input("\n処理が完了しました。Enterキーを押して終了してください。")


if __name__ == "__main__":
    main()
