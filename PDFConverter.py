# -*- coding: utf-8 -*-
import sys
import os
import io
import time
import threading
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

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


def wait_and_exit(root=None):
    """3秒待機してから終了する共通処理"""
    if root:
        try:
            root.destroy()  # Tkinterのリソースを安全に解放
        except Exception:
            pass

    if getattr(sys, 'frozen', False) or sys.stdin.isatty():
        print("\n3秒経過後、自動的に終了します...")
        time.sleep(3)


def _convert_single_file(converter_cls, file_path, output_dir, success_info,
                         error_files, lock):
    """
    単一ファイルをOfficeアプリインスタンスで変換する。
    各スレッドから呼ばれる想定で、COM初期化も内部で行う。
    """
    import pythoncom
    pythoncom.CoInitialize()
    converter = None
    try:
        converter = converter_cls()
        basename = os.path.basename(file_path)
        print(f"PDFに変換中: {basename}")
        try:
            result_path = converter.convert(file_path, output_dir=output_dir)
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


def _process_group(converter_cls, files, output_dir, success_info,
                   error_files, lock, max_workers=2):
    """
    同種ファイルを複数のOfficeアプリインスタンスで並列変換する。
    """
    if not files:
        return

    # ファイル数が少なければ逐次処理、多い場合は並列処理
    if len(files) <= 2:
        for file_path in files:
            _convert_single_file(
                converter_cls, file_path, output_dir, success_info,
                error_files, lock)
    else:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        workers = min(len(files), max_workers)
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    _convert_single_file, converter_cls, file_path,
                    output_dir, success_info, error_files, lock): file_path
                for file_path in files
            }
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    file_path = futures[future]
                    err_msg = f"{os.path.basename(file_path)} (エラー: {str(e)})"
                    error_files.append(err_msg)


def main():
    # --- PyInstaller EXE の argv を Unicode に強制統一 ---
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        def get_unicode_argv():
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

    # --- ドラッグアンドドロップされたファイル一覧の表示 ---
    # Unicodeに統一した後に行うことで、ファイル名の文字化けを防ぎます
    if len(sys.argv) > 1:
        print("\n【ドラッグ＆ドロップされたファイル一覧】")
        for arg in sys.argv[1:]:
            clean_arg = arg.strip().strip('"')
            print(f" - {os.path.basename(clean_arg)}")
        print("-" * 50)

    # 重いモジュールのインポート（一度だけ）
    try:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        from converters import (
            ExcelConverter, WordConverter, PowerPointConverter)
    except Exception as e:
        print(f"モジュールの読み込みに失敗しました: {e}")
        wait_and_exit()
        return

    # 引数チェック
    if len(sys.argv) < 2:
        print("\nWord、Excel、PowerPointファイルをこの実行ファイルにドラッグアンドドロップしてください。")
        wait_and_exit()
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
        wait_and_exit()
        return

    if not files_to_process:
        print("\n処理対象のファイルが見つかりませんでした。")
        wait_and_exit()
        return

    # 上書き確認等のため、Tkinterを非表示で初期化
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # --- 保存先フォルダの決定（固定名）とコマンドプロンプトへの表示 ---
    base_dir = os.path.dirname(files_to_process[0])
    date_str = datetime.now().strftime("%Y%m%d")
    folder_name = f"{date_str}_PDF"
    output_dir = os.path.join(base_dir, folder_name)

    print("\n【保存先フォルダ情報】")
    print(f"作成場所: {base_dir}")
    print(f"フォルダ名: {folder_name}")
    print("-" * 50 + "\n")

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
        wait_and_exit(root)
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

    # Excel / Word / PowerPoint を並列で変換
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

    # --- 作成したフォルダをオープン ---
    print("\n処理が完了しました。作成されたフォルダを開きます。")
    if os.path.exists(output_dir):
        try:
            if hasattr(os, 'startfile'):  # Windows
                os.startfile(output_dir)
            elif sys.platform == 'darwin':  # Mac (念のため)
                import subprocess
                subprocess.Popen(['open', output_dir])
        except Exception as e:
            print(f"フォルダを開けませんでした: {e}")

    # 終了処理 (3秒待機してGUIもクリーンアップ)
    wait_and_exit(root)


if __name__ == "__main__":
    main()