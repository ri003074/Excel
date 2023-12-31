import dataclasses
import win32com.client
import sys
import pandas as pd
from typing import List
from lib._excel import _set_graph_title
from lib._excel import _set_axis_title
from lib._excel import _set_axis_obj
from lib._excel import _set_line_format


@dataclasses.dataclass()
class GraphParameter:
    input_file: str = None
    graph_type: int = None
    graph_title: str = None
    graph_range: str = None
    axis_y_title: str = None
    axis_y_ticks: List = None
    axis_y_font_size: int = None


@dataclasses.dataclass()
class ExcelVariable:
    xlLine: int = 4
    xlLineMarkers: int = 65
    xlXYScatter: int = -4169
    xlCategory: int = 1
    xlValue: int = 2
    msoFalse: int = 0
    msoTrue: int = -1


class Excel:
    def __init__(self):
        self.xl = None
        self.wb = None
        self.ws = None
        self.shape = None
        self.xlsx_file_path = None

    def csv_to_xlsx(self, file_path):
        file_path_length = len(file_path)
        if file_path_length > 255:
            sys.exit(f"file path length {file_path_length} is too long")

        df = pd.read_csv(file_path)
        self.xlsx_file_path = file_path.replace(".csv", ".xlsx")
        try:
            df.to_excel(self.xlsx_file_path, index=False, header=True)
        except PermissionError:
            print(f"file {self.xlsx_file_path} is already opened")

    def open_xlsx_file(self):
        self.xl = win32com.client.Dispatch("Excel.Application")
        self.xl.Visible = True
        self.wb = self.xl.Workbooks.Open(self.xlsx_file_path)

    def setup_active_excel(self):
        self.xl = win32com.client.GetObject(Class="Excel.Application")
        self.wb = self.xl.Workbooks(1)
        self.ws = self.wb.ActiveSheet

    def add_chart(self, graph_type, graph_range=None):
        if graph_range is None:
            self.ws.Range("A1").CurrentRegion.Select()
        elif ":" in graph_range:
            self.ws.Range(graph_range).Select()
        else:
            self.ws.Range(graph_range).CurrentRegion.Select()

        self.shape = self.ws.Shapes.AddChart2(-1, graph_type)

    def delete_shape(self):
        for i in range(self.ws.Shapes.Count, 0, -1):
            self.ws.Shapes(i).Delete()

    def resize_graph(self, height, width):
        for ws in self.wb.Sheets:
            for i in range(ws.Shapes.Count):
                shape = ws.Shapes(i + 1)
                shape.Select()
                shape.Height = height
                shape.Width = width

    def relocate_graph(self, top_left_cell="E2"):
        for ws in self.wb.Sheets:
            for i in range(ws.Shapes.Count):
                shape = ws.Shapes(i + 1)
                shape.Left = ws.Range(top_left_cell).Left
                if i == 0:
                    shape.Top = ws.Range(top_left_cell).Top
                else:
                    shape.Top = ws.Cells(
                        ws.Shapes(i).BottomRightCell.Row + 1,
                        ws.Range(top_left_cell).Left,
                    ).Top

    def save_png(self):
        for ws in self.wb.Sheets:
            for i in range(ws.Shapes.Count):
                title = ws.Shapes(i + 1).Chart.ChartTitle.Text
                ws.Shapes(i + 1).Select()
                save_file_path = (
                    self.wb.path
                    + "/"
                    + self.wb.name.replace(".xlsx", "_")
                    + title
                    + "_"
                    + str(i)
                    + ".png"
                )
                ws.Shapes(i + 1).Chart.Export(save_file_path)

    def set_axis_title(self, axis, title, font_size=None):
        if axis == "x":
            axis_type = ExcelVariable.xlCategory
        else:
            axis_type = ExcelVariable.xlValue

        if self.shape is None:
            for ws in self.wb.Sheets:
                for i in range(ws.Shapes.Count):
                    shape = ws.Shapes(i + 1)
                    _set_axis_title(shape, axis_type, title, font_size)
        else:
            shape = self.shape
            _set_axis_title(shape, axis_type, title, font_size)

    def set_graph_title(self, title):
        if self.shape is None:
            for ws in self.wb.Sheets:
                for i in range(ws.Shapes.Count):
                    shape = ws.Shapes(i + 1)
                    _set_graph_title(shape, title)
        else:
            shape = self.shape
            _set_graph_title(shape, title)

    def set_tick(self, axis, minimum, maximum, resolution):
        if axis == "x":
            axis_type = ExcelVariable.xlCategory
        else:
            axis_type = ExcelVariable.xlValue

        if self.shape is None:
            for ws in self.wb.Sheets:
                for i in range(ws.Shapes.Count):
                    shape = ws.Shapes(i + 1)
                    _set_axis_obj(
                        shape, axis_type, minimum, maximum, resolution
                    )
        else:
            shape = self.shape
            _set_axis_obj(shape, axis_type, minimum, maximum, resolution)

    def set_line_format(self, fill=ExcelVariable.msoTrue):
        if self.shape is None:
            for ws in self.wb.Sheets:
                for i in range(ws.Shapes.Count):
                    shape = ws.Shapes(i + 1)
                    shape.Select()
                    _set_line_format(self.xl, shape, fill)

        else:
            shape = self.shape
            shape.Select()
            _set_line_format(self.xl, shape, fill)

    def save_workbook(self):
        self.xl.DisplayAlerts = False
        self.wb.Save()

    def quit(self):
        self.xl.DisplayAlerts = False
        self.xl.Quit()
