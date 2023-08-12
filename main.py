from app.app import relocate_graph
from app.app import resize_graph
from app.app import save_png
from app.app import set_graph_title
from app.app import set_axis_title
from app.app import set_ticks
from app.app import set_line_format
from app.app import make_graph
from lib.excel import ExcelVariable
from lib.excel import GraphParameter

if __name__ == "__main__":
    resize_graph(200, 300)
    relocate_graph()
    save_png()
    set_axis_title(axis="y", title="mV")
    set_graph_title(title="test")
    set_ticks(axis="y", minimum=0, maximum=120, resolution=30)
    set_line_format(fill=ExcelVariable.msoFalse)

    gp1 = GraphParameter()
    gp1.graph_type = ExcelVariable.xlLineMarkers
    gp1.graph_title = "title1"
    gp1.graph_range = "A1"
    gp1.axis_y_title = "y1"
    gp1.axis_y_ticks = [0, 60, 20]

    gp2 = GraphParameter()
    gp2.graph_type = ExcelVariable.xlLineMarkers
    gp2.graph_title = "title2"
    gp2.graph_range = "A6"
    gp2.axis_y_title = "y2"
    gp2.axis_y_ticks = [0, 120, 20]

    make_graph(graph_parameter=[gp1, gp2])
