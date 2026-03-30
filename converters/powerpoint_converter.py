import os
import win32com.client
import pythoncom
from .base_converter import BaseConverter


class PowerPointConverter(BaseConverter):
    """
    PowerPointファイル(.pptx/.ppt/.pptm)をPDFに変換するクラス。
    __init__でPowerPointアプリを1回だけ起動し、convert()を複数回呼び出せる。
    全ファイルの変換後に close() でアプリを終了する。
    """

    def __init__(self):
        self.ppt = win32com.client.DispatchEx("PowerPoint.Application")
        # PowerPointはVisibleプロパティの設定方法がExcel/Wordと異なる
        # WithWindow=False で開けば非表示で処理できる

    def convert(self, file_path, output_dir=None):
        prs = None
        try:
            abs_path = os.path.abspath(file_path)
            pdf_path = self._build_pdf_path(abs_path, output_dir)

            # 既存PDFがあれば削除
            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                except Exception:
                    pass

            # WithWindow=False で非表示のまま開く
            # msoFalse = 0
            prs = self.ppt.Presentations.Open(
                abs_path,
                ReadOnly=True,
                Untitled=False,
                WithWindow=False
            )

            # ExportAsFixedFormat でPDF出力
            # ppFixedFormatTypePDF = 2
            # ppFixedFormatIntentPrint = 2
            # ppPrintAll = 1 (全スライド)
            prs.ExportAsFixedFormat(
                pdf_path,
                2,                   # ppFixedFormatTypePDF
                Intent=2,            # ppFixedFormatIntentPrint
                FrameSlides=False,
                HandoutOrder=1,
                OutputType=1,        # ppPrintOutputSlides
                PrintHiddenSlides=False,
                PrintRange=None,
                RangeType=1,         # ppPrintAll
                SlideShowName="",
                IncludeDocProperties=True,
                KeepIRMSettings=True,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )

            return pdf_path

        except Exception as e:
            raise RuntimeError(f"PowerPoint to PDF変換に失敗しました: {str(e)}")
        finally:
            if prs:
                try:
                    prs.Close()
                except Exception:
                    pass

    def close(self):
        try:
            self.ppt.Quit()
        except Exception:
            pass
