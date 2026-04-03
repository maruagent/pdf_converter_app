import os
import win32com.client
import pythoncom
from .base_converter import BaseConverter


class ExcelConverter(BaseConverter):
    """
    ExcelファイルをPDFに変換するクラス。
    __init__でExcelアプリを1回だけ起動し、convert()を複数回呼び出せる。
    全ファイルの変換後に close() でアプリを終了する。
    """

    def __init__(self):
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.Interactive = False
        self.excel.ScreenUpdating = False
        self.excel.EnableEvents = False

    def convert(self, file_path, output_dir=None):
        wb = None
        try:
            abs_path = os.path.abspath(file_path)
            pdf_path = self._build_pdf_path(abs_path, output_dir)

            # 既存PDFがあれば削除
            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                except Exception:
                    pass

            wb = self.excel.Workbooks.Open(
                abs_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            )

            wb.ExportAsFixedFormat(
                Type=0,              # xlTypePDF
                Filename=pdf_path,
                OpenAfterPublish=False,
                Quality=0,           # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False
            )

            return pdf_path

        except Exception as e:
            raise RuntimeError(f"Excel to PDF変換に失敗しました: {str(e)}")
        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def close(self):
        try:
            self.excel.Quit()
        except Exception:
            pass
