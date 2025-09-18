import xlwings as xw
import sys
import subprocess
from typing import Any, List, Optional, Dict
from ..core.exceptions import ToolExecutionError
from ..core.excel_manager import ExcelManager

class ExcelActions:
    """
    具体的なExcel操作を実行するメソッドを集約したクラス。
    """

    def __init__(self, manager: ExcelManager):
        if not manager or not manager.get_active_workbook():
            raise ValueError("有効なExcelManagerインスタンスが必要です。")
        self.book = manager.get_active_workbook()

    def _get_sheet(self, sheet_name: Optional[str] = None) -> xw.Sheet:
        try:
            return self.book.sheets[sheet_name] if sheet_name else self.book.sheets.active
        except Exception as e:
            raise ToolExecutionError(f"シート '{sheet_name or 'アクティブ'}' の取得に失敗: {e}")

    def write_to_cell(self, cell: str, value: Any, sheet_name: Optional[str] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            sheet.range(cell).value = value
            return f"セル {cell} に値 '{value}' を正常に書き込みました。"
        except Exception as e:
            raise ToolExecutionError(f"セル '{cell}' への書き込み中にエラーが発生しました: {e}")

    def read_cell_value(self, cell: str, sheet_name: Optional[str] = None) -> Any:
        try:
            sheet = self._get_sheet(sheet_name)
            value = sheet.range(cell).value
            return f"セル '{cell}' の値は '{value}' です。"
        except Exception as e:
            raise ToolExecutionError(f"セル '{cell}' の読み取り中にエラーが発生しました: {e}")

    def get_sheet_names(self) -> List[str]:
        try:
            return [sheet.name for sheet in self.book.sheets]
        except Exception as e:
            raise ToolExecutionError(f"シート名の取得中にエラーが発生しました: {e}")

    def set_formula(self, cell: str, formula: str, sheet_name: Optional[str] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            sheet.range(cell).formula = formula
            return f"セル {cell} に数式 '{formula}' を正常に設定しました。"
        except Exception as e:
            raise ToolExecutionError(f"数式 '{formula}' の設定中にエラーが発生しました: {e}")

    def copy_range(self, source_range: str, destination_range: str, sheet_name: Optional[str] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            sheet.range(source_range).copy(sheet.range(destination_range))
            return f"範囲 '{source_range}' を '{destination_range}' に正常にコピーしました。"
        except Exception as e:
            raise ToolExecutionError(f"範囲のコピー中にエラーが発生しました: {e}")

    def read_range(self, cell_range: str, sheet_name: Optional[str] = None) -> List[List[Any]]:
        try:
            sheet = self._get_sheet(sheet_name)
            return sheet.range(cell_range).value
        except Exception as e:
            raise ToolExecutionError(f"範囲 '{cell_range}' の読み取り中にエラーが発生しました: {e}")

    def write_range(self, cell_range: str, data: List[List[Any]], sheet_name: Optional[str] = None) -> str:
        """指定された範囲にデータを書き込む前に、データの次元が範囲と一致するかを検証する。"""
        try:
            sheet = self._get_sheet(sheet_name)
            target_range = sheet.range(cell_range)

            if not isinstance(data, list) or (data and not isinstance(data[0], list)):
                 raise ToolExecutionError("書き込むデータは2次元リストである必要があります。")
            
            data_rows = len(data)
            data_cols = len(data[0]) if data_rows > 0 else 0

            range_rows = target_range.rows.count
            range_cols = target_range.columns.count

            if data_rows != range_rows or data_cols != range_cols:
                error_msg = (
                    f"データの次元が一致しません。書き込み先範囲 ({cell_range}) は {range_rows}行 x {range_cols}列ですが、"
                    f"提供されたデータは {data_rows}行 x {data_cols}列です。読み取ったデータと同じ次元のデータを渡してください。"
                )
                raise ToolExecutionError(error_msg)

            target_range.value = data
            return f"範囲 '{cell_range}' にデータを正常に書き込みました。"
        except Exception as e:
            if isinstance(e, ToolExecutionError):
                raise e
            raise ToolExecutionError(f"範囲 '{cell_range}' への書き込み中に予期せぬエラーが発生しました: {e}")

    def get_active_workbook_and_sheet(self) -> Dict[str, str]:
        """現在アクティブなExcelブックとシート名を取得し、辞書として返す。"""
        try:
            book_name = self.book.name
            sheet_name = self.book.sheets.active.name
            return {"workbook_name": book_name, "sheet_name": sheet_name}
        except Exception as e:
            raise ToolExecutionError(f"ブックとシートの取得中にエラーが発生しました: {e}")

    def format_range(self,
                     cell_range: str,
                     sheet_name: Optional[str] = None,
                     font_name: Optional[str] = None,
                     font_size: Optional[float] = None,
                     font_color_hex: Optional[str] = None,
                     bold: Optional[bool] = None,
                     italic: Optional[bool] = None,
                     fill_color_hex: Optional[str] = None,
                     column_width: Optional[float] = None,
                     row_height: Optional[float] = None,
                     horizontal_alignment: Optional[str] = None,
                     border_style: Optional[Dict[str, Any]] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            rng = sheet.range(cell_range)

            if font_name: rng.font.name = font_name
            if font_size: rng.font.size = font_size
            if font_color_hex: rng.font.color = font_color_hex
            if bold is not None: rng.font.bold = bold
            if italic is not None: rng.font.italic = italic
            if fill_color_hex: rng.color = fill_color_hex
            if column_width: rng.column_width = column_width
            if row_height: rng.row_height = row_height
            if horizontal_alignment:
                h_align_map = {
                    "general": xw.constants.HAlign.xlHAlignGeneral,
                    "left": xw.constants.HAlign.xlHAlignLeft,
                    "center": xw.constants.HAlign.xlHAlignCenter,
                    "right": xw.constants.HAlign.xlHAlignRight,
                    "justify": xw.constants.HAlign.xlHAlignJustify,
                }
                align_const = h_align_map.get(horizontal_alignment.lower())
                if align_const: rng.api.HorizontalAlignment = align_const

            if border_style:
                edges = border_style.get("edges", [])
                style = border_style.get("style", "continuous")
                weight = border_style.get("weight", "thin")
                color_hex = border_style.get("color_hex", "#000000")

                style_map = {
                    "continuous": xw.constants.LineStyle.xlContinuous,
                    "dot": xw.constants.LineStyle.xlDot,
                    "dash": xw.constants.LineStyle.xlDash,
                    "none": xw.constants.LineStyle.xlLineStyleNone
                }
                weight_map = {
                    "thin": xw.constants.BorderWeight.xlThin,
                    "medium": xw.constants.BorderWeight.xlMedium,
                    "thick": xw.constants.BorderWeight.xlThick
                }
                edge_map = {
                    "top": xw.constants.BordersIndex.xlEdgeTop,
                    "bottom": xw.constants.BordersIndex.xlEdgeBottom,
                    "left": xw.constants.BordersIndex.xlEdgeLeft,
                    "right": xw.constants.BordersIndex.xlEdgeRight,
                    "inside_vertical": xw.constants.BordersIndex.xlInsideVertical,
                    "inside_horizontal": xw.constants.BordersIndex.xlInsideHorizontal
                }

                for edge_name in edges:
                    edge_const = edge_map.get(edge_name.lower())
                    if edge_const:
                        border = rng.api.Borders(edge_const)
                        border.LineStyle = style_map.get(style, xw.constants.LineStyle.xlContinuous)
                        border.Weight = weight_map.get(weight, xw.constants.BorderWeight.xlThin)
                        border.Color = int(color_hex.replace("#", ""), 16)

            return f"範囲 '{cell_range}' の書式を正常に設定しました。"
        except Exception as e:
            raise ToolExecutionError(f"範囲 '{cell_range}' の書式設定中にエラー: {e}")

    def insert_shape_in_range(self,
                              cell_range: str,
                              shape_type: str,
                              sheet_name: Optional[str] = None,
                              fill_color_hex: Optional[str] = None,
                              line_color_hex: Optional[str] = None) -> str:
        """指定されたセル範囲に、指定された書式でオートシェイプを挿入します。"""
        try:
            sheet = self._get_sheet(sheet_name)
            target_range = sheet.range(cell_range)

            if sys.platform == 'darwin':
                # Note: This is a workaround. xlwings does not have a direct way to add shapes on mac.
                # This will add a picture of a rectangle, which is not ideal.
                if shape_type == "四角形":
                    import io
                    import os
                    import tempfile
                    from PIL import Image, ImageDraw
                    
                    img = Image.new('RGB', (int(target_range.width), int(target_range.height)), color = 'white')
                    draw = ImageDraw.Draw(img)
                    if fill_color_hex:
                        draw.rectangle([(0,0), img.size], fill=fill_color_hex)
                    if line_color_hex:
                        draw.rectangle([(0,0), (img.size[0]-1, img.size[1]-1)], outline=line_color_hex)

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_image:
                        img.save(temp_image, format='PNG')
                        temp_image_path = temp_image.name
                    
                    sheet.pictures.add(
                        temp_image_path,
                        left=target_range.left,
                        top=target_range.top,
                        width=target_range.width,
                        height=target_range.height
                    )
                    os.remove(temp_image_path)
                else:
                    raise ToolExecutionError(f"Unsupported shape type on macOS: {shape_type}")


            elif sys.platform == 'win32':
                shape_map = {"四角形": 1, "楕円": 9}
                mso_shape_type = shape_map.get(shape_type, 1)
                new_shape = sheet.api.Shapes.AddShape(
                    mso_shape_type, target_range.left, target_range.top,
                    target_range.width, target_range.height
                )
                if fill_color_hex:
                    rgb = tuple(int(fill_color_hex.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
                    new_shape.Fill.ForeColor.RGB = rgb[0] | (rgb[1] << 8) | (rgb[2] << 16)
                    new_shape.Fill.Visible = True
                    new_shape.Fill.Solid()
                if line_color_hex:
                    rgb = tuple(int(line_color_hex.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
                    new_shape.Line.ForeColor.RGB = rgb[0] | (rgb[1] << 8) | (rgb[2] << 16)
                    new_shape.Line.Visible = True
            else:
                raise ToolExecutionError(f"サポートされていないOSです: {sys.platform}")

            return f"範囲 '{cell_range}' に '{shape_type}' を正常に挿入しました。"
        except Exception as e:
            raise ToolExecutionError(f"範囲 '{cell_range}' への図形挿入中にエラーが発生しました: {e}")

    def format_last_shape(self, fill_color_hex: Optional[str] = None, line_color_hex: Optional[str] = None, sheet_name: Optional[str] = None) -> str:
        """この関数は非推奨になりました。insert_shapeを使用してください。"""
        raise ToolExecutionError("format_last_shapeは非推奨です。insert_shapeに色を指定してください。")


