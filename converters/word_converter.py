import os
import win32com.client
import pythoncom
from .base_converter import BaseConverter


class WordConverter(BaseConverter):
    """
    WordファイルをPDFに変換するクラス。
    __init__でWordアプリを1回だけ起動し、convert()を複数回呼び出せる。
    全ファイルの変換後に close() でアプリを終了する。
    """

    def __init__(self):
        self.word = win32com.client.DispatchEx("Word.Application")
        self.word.Visible = False
        self.word.DisplayAlerts = 0  # wdAlertsNone
        self.word.ScreenUpdating = False
        self.word.Options.UpdateLinksAtOpen = False

    def convert(self, file_path, output_dir=None):
        doc = None
        try:
            abs_path = os.path.abspath(file_path)
            pdf_path = self._build_pdf_path(abs_path, output_dir)

            # 既存PDFがあれば削除
            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                except Exception:
                    pass

            doc = self.word.Documents.Open(
                abs_path,
                ReadOnly=True,
                ConfirmConversions=False
            )

            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17,         # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,           # wdExportOptimizeForPrint
                Range=0,                 # wdExportAllDocument
                From=1,
                To=1,
                Item=0,                  # wdExportDocumentContent
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=0,       # wdExportCreateNoBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )

            return pdf_path

        except Exception as e:
            raise RuntimeError(f"Word to PDF変換に失敗しました: {str(e)}")
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
                except Exception:
                    pass

    def close(self):
        try:
            self.word.Quit()
        except Exception:
            pass
