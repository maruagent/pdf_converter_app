# PDFConverter Pro

Office (Word, Excel, PowerPoint) ファイルをPDFに高速変換するWindows用アプリケーションです。
ドラッグ＆ドロップだけで、複数のファイルを一括でPDFに変換できます。

## 主な機能
- **Word変換**: `.docx`, `.doc` ファイルをPDFに変換します。
- **Excel変換**: `.xlsx`, `.xls`, `.xlsm` ファイルをPDFに変換します（全シート対応）。
- **PowerPoint変換**: `.pptx`, `.ppt`, `.pptm` ファイルをPDFに変換します。
- **高速一括処理**: アプリインスタンスを再利用することで、複数ファイル処理時の起動コストを大幅に削減しました。
- **自動フォルダ生成**: 実行時に日付ベースの保存用フォルダ名を自動提案します。
- **詳細レポート**: 変換成功時に、保存先のフォルダパスとファイル名を一覧表示します。

## 使用方法
1. `PDFConverter.exe` (または `main.py`) に変換したいファイルをドラッグ＆ドロップします。
2. 保存先フォルダ名を入力（デフォルトは `YYYYMMDD_`）します。
3. 変換が完了すると、指定したフォルダ内にPDFが順次作成されます。

## 動作環境
- **OS**: Windows 10 / 11
- **必須ソフト**: Microsoft Office (Word, Excel, PowerPoint) がインストールされている必要があります。
- **開発環境**: Python 3.10以上

## 開発者向け情報

### 依存ライブラリのインストール
```bash
pip install -r requirements.txt
```

### 実行ファイルのビルド方法
PyInstallerを使用してビルドします。
```bash
pyinstaller PDFConverter.spec
```

## ファイル構成
```
pdf_converter_app/
├── main.py              # アプリケーションのメインエントリーポイント
├── converters/          # 変換エンジンモジュール
│   ├── __init__.py
│   ├── base_converter.py  # コンバーターの基底クラス・共通インターフェース
│   ├── excel_converter.py # Excel用コンバーター
│   ├── word_converter.py  # Word用コンバーター
│   └── powerpoint_converter.py # PowerPoint用コンバーター
├── requirements.txt     # 依存ライブラリ (pywin32, pyinstaller)
├── PDFConverter.spec    # PyInstallerビルド定義
└── README.md            # 本ファイル
```

## 更新履歴

- **v1.3 (Current)**:
  - **PowerPoint対応**: `.pptx`, `.ppt`, `.pptm` の変換をサポート
  - **パフォーマンス改善**: 複数ファイル変換時にOfficeアプリを再起動せず、インスタンスを再利用するよう最適化
  - **コード整理**: 冗長なテストスクリプトや重複するビルド定義を削除
- **v1.2以前**: Excel全シート化、文字化け対策、COM初期化の安定化など
