import os

class BaseConverter:
    """
    コンバーターの基底クラス。
    インスタンスはOfficeアプリの起動を表し、複数ファイルを変換後に close() で終了する。
    """

    def convert(self, file_path, output_dir=None):
        """
        単一ファイルをPDFに変換する。
        :param file_path: 変換元ファイルの絶対パス
        :param output_dir: PDF出力先ディレクトリ（Noneの場合は同じディレクトリ）
        :return: 出力されたPDFの絶対パス
        """
        raise NotImplementedError

    def close(self):
        """Officeアプリを終了する。"""
        pass

    @staticmethod
    def _build_pdf_path(file_path, output_dir):
        """PDFの出力パスを生成するユーティリティ。"""
        abs_path = os.path.abspath(file_path)
        basename = os.path.basename(abs_path)
        pdf_name = os.path.splitext(basename)[0] + ".pdf"
        if output_dir:
            return os.path.join(output_dir, pdf_name)
        return os.path.splitext(abs_path)[0] + ".pdf"
