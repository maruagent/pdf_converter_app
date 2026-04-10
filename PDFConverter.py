# -*- coding: utf-8 -*-
import sys
import os
import io
import threading
from datetime import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox

# Windowsコンソールでの文字化けを最小限にするため、
# インポート直後にエンコーディングを設定
if sys.platform == "win32":
    try:
        if sys.stdout and hasattr(sys.stdout, 'buffer'):
            sys.stdout = io.TextIOWrapper(
                sys.stdout.buffer, encoding='utf-8', errors='replace')
        if sys.stderr and hasattr(sys.stderr, 'buffer'):
            sys.stderr = io.TextIOWrapper(
                sys.stderr.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass

print("分析・変換をしています。しばらくお待ちください...")
if sys.stdout:
    sys.stdout.flush()

# 対応拡張子の定義
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
WORD_EXTS = {".docx", ".doc"}
PPT_EXTS = {".pptx", ".ppt", ".pptm"}
SUPPORTED_EXTS = EXCEL_EXTS | WORD_EXTS | PPT_EXTS


def _process_group(converter_cls, files, output_dir, success_info,
                   error_files, lock):
    """
    同種ファイルを1つのOfficeアプリインスタンスで一括変換する。
    各スレッドから呼ばれる想定で、COM初期化も内部で行う。
    converter_cls: ExcelConverter / WordConverter / PowerPointConverter
    """
    if not files:
        return

    import pythoncom
    pythoncom.CoInitialize()
    converter = None
    try:
        converter = converter_cls()
        for file_path in files:
            basename = os.path.basename(file_path)
            print(f"PDFに変換中: {basename}")
            try:
                result_path = converter.convert(
                    file_path, output_dir=output_dir)
                with lock:
                    success_info.append({
                        'dir': output_dir,
                        'file': os.path.basename(result_path)
                    })
            except Exception as e:
                error_message = str(e).strip() or "不明なエラーが発生しました"
                with lock:
                    error_files.append(f"{basename} (エラー: {error_message})")
    finally:
        if converter:
            converter.close()
        pythoncom.CoUninitialize()


def main():
    # --- PyInstaller EXE の argv を Unicode に強制統一 ---
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        def get_unicode_argv():
            """WindowsのGetCommandLineWを使用してUnicodeの引数リストを取得する"""
            GetCommandLineW = ctypes.windll.kernel32.GetCommandLineW
            GetCommandLineW.restype = wintypes.LPCWSTR
            CommandLineToArgvW = ctypes.windll.shell32.CommandLineToArgvW
            CommandLineToArgvW.argtypes = [
                wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_int)]
            CommandLineToArgvW.restype = ctypes.POINTER(wintypes.LPWSTR)

            argc = ctypes.c_int(0)
            argv_unicode = CommandLineToArgvW(
                GetCommandLineW(), ctypes.byref(argc))
            if not argv_unicode:
                return sys.argv

            try:
                return [argv_unicode[i] for i in range(argc.value)]
            finally:
                ctypes.windll.kernel32.LocalFree(argv_unicode)

        sys.argv = get_unicode_argv()

    # 重いモジュールのインポート（一度だけ）
    try:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        from converters import (
            ExcelConverter, WordConverter, PowerPointConverter)
    except Exception as e:
        print(f"モジュールの読み込みに失敗しました: {e}")
        if getattr(sys, 'frozen', False):
            input("\nEnterキーを押して終了してください。")
        return

    # 引数チェック
    if len(sys.argv) < 2:
        print("\nWord、Excel、PowerPointファイルをこの実行ファイルにドラッグアンドドロップしてください。")
        if getattr(sys, 'frozen', False):
            input("\nEnterキーを押して終了してください。")
        return

    # 対応拡張子のチェックとファイル収集
    files_to_process = []
    unsupported_found = False

    for file_path in sys.argv[1:]:
        if not file_path or not file_path.strip():
            continue

        clean_path = file_path.strip().strip('"')
        if not os.path.exists(clean_path):
            continue

        ext = os.path.splitext(clean_path)[1].lower()
        if ext in SUPPORTED_EXTS:
            files_to_process.append(os.path.abspath(clean_path))
        else:
            unsupported_found = True

    if unsupported_found and not files_to_process:
        print("\nエラー: Word、Excel、またはPowerPointファイルのみ対応しています。")
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

    base_dir = os.path.dirname(files_to_process[0])
    date_str = datetime.now().strftime("%Y%m%d")
    default_name = f"{date_str}_PDFフォルダ"

    prompt_msg = "元のファイルと同じ場所に新たなフォルダを作ります。\nフォルダ名称を入力してください。"
    user_input = simpledialog.askstring(
        "フォルダ作成", prompt_msg, initialvalue=default_name, parent=root)

    if user_input is None:
        print("\nキャンセルされました。")
        return

    try:
        folder_name = user_input.strip()
    except Exception:
        folder_name = default_name

    if not folder_name:
        folder_name = default_name

    output_dir = os.path.join(base_dir, folder_name)

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"フォルダを作成しました: {folder_name}")
    except Exception as e:
        print(f"フォルダの作成に失敗しました: {e}")
        output_dir = base_dir

    # 上書き確認（GUI）
    final_files = []
    for file_path in files_to_process:
        basename = os.path.basename(file_path)
        pdf_name = os.path.splitext(basename)[0] + ".pdf"
        pdf_path = os.path.join(output_dir, pdf_name)

        if os.path.exists(pdf_path):
            confirm_msg = f"ファイル '{pdf_name}' は既に存在します。\n上書きしますか？"
            if not messagebox.askyesno("上書き確認", confirm_msg, parent=root):
                print(f"スキップしました: {basename}")
                continue

        final_files.append(file_path)

    if not final_files:
        print("\n処理するファイルがありませんでした。")
        if getattr(sys, 'frozen', False) or sys.stdin.isatty():
            input("\n処理が完了しました。Enterキーを押して終了してください。")
        return

    # ファイルをタイプ別にグループ化
    excel_files = [
        f for f in final_files
        if os.path.splitext(f)[1].lower() in EXCEL_EXTS
    ]
    word_files = [
        f for f in final_files
        if os.path.splitext(f)[1].lower() in WORD_EXTS
    ]
    ppt_files = [
        f for f in final_files
        if os.path.splitext(f)[1].lower() in PPT_EXTS
    ]

    success_info = []
    error_files = []
    lock = threading.Lock()

    # Excel / Word / PowerPoint を並列で変換（各スレッドが独立して COM を初期化）
    groups = [
        (ExcelConverter,      excel_files),
        (WordConverter,       word_files),
        (PowerPointConverter, ppt_files),
    ]
    active_groups = [(cls, files) for cls, files in groups if files]

    with ThreadPoolExecutor(max_workers=len(active_groups) or 1) as executor:
        futures = {
            executor.submit(
                _process_group, cls, files, output_dir, success_info,
                error_files, lock): cls.__name__
            for cls, files in active_groups
        }
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"変換スレッドでエラーが発生しました: {e}")

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
