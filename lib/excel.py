import os
import dataclasses
import win32com.client


@dataclasses.dataclass()
class ExcelVariable:
    xlCategory = 1
    xlValue = 2


class Excel:
    def __init__(self):
        self.wb = None

    def setup_active_excel(self):
        xl = win32com.client.GetObject(Class="Excel.Application")
        self.wb = xl.Workbooks(1)

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
        for ws in self.wb.Sheets:
            for i in range(ws.Shapes.Count):
                shape = ws.Shapes(i + 1)

                if axis == "x":
                    axis_type = ExcelVariable.xlCategory
                else:
                    axis_type = ExcelVariable.xlValue

                axis = shape.Chart.Axes(axis_type)
                axis.HasTitle = True
                axis.AxisTitle.Characters.Text = title

    def set_graph_title(self, title):
        for ws in self.wb.Sheets:
            for i in range(ws.Shapes.Count):
                shape = ws.Shapes(i + 1)
                shape.Chart.HasTitle = True
                shape.Chart.ChartTitle.Text = title

    def set_tick(self, axis, minimum, maximum, resolution):
        if axis == "x":
            axis_type = ExcelVariable.xlCategory
        else:
            axis_type = ExcelVariable.xlValue

        for ws in self.wb.Sheets:
            for i in range(ws.Shapes.Count):
                shape = ws.Shapes(i + 1)
                axis_obj = shape.Chart.Axes(axis_type)
                axis_obj.MinimumScale = minimum
                axis_obj.MaximumScale = maximum
                axis_obj.MajorUnit = resolution
