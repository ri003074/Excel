import dataclasses
import win32com.client
from typing import List
from lib._excel import _set_graph_title
from lib._excel import _set_axis_title
from lib._excel import _set_axis_obj


@dataclasses.dataclass()
class GraphParameter:
    input_file: str = None
    graph_type: int = None
    graph_title: str = None
    graph_range: str = None
    axis_y_title: str = None
    axis_y_ticks: List = None


@dataclasses.dataclass()
class ExcelVariable:
    xlLine: int = 4
    xlLineMarkers: int = 65
    xlXYScatter: int = -4169
    xlCategory: int = 1
    xlValue: int = 2


class Excel:
    def __init__(self):
        self.wb = None
        self.ws = None
        self.shape = None

    def setup_active_excel(self):
        xl = win32com.client.GetObject(Class="Excel.Application")
        self.wb = xl.Workbooks(1)
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
                print(save_file_path)

    def set_axis_title(self, axis, title):
        if axis == "x":
            axis_type = ExcelVariable.xlCategory
        else:
            axis_type = ExcelVariable.xlValue

        if self.shape is None:
            for ws in self.wb.Sheets:
                for i in range(ws.Shapes.Count):
                    shape = ws.Shapes(i + 1)
                    _set_axis_title(shape, axis_type, title)
        else:
            shape = self.shape
            _set_axis_title(shape, axis_type, title)

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
