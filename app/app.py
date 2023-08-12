from lib.excel import Excel
from lib.excel import GraphParameter
from typing import List


def resize_graph(height, width):
    xl = Excel()
    xl.setup_active_excel()
    xl.resize_graph(height=height, width=width)


def relocate_graph():
    xl = Excel()
    xl.setup_active_excel()
    xl.relocate_graph()


def save_png():
    xl = Excel()
    xl.setup_active_excel()
    xl.save_png()


def set_axis_title(axis, title):
    xl = Excel()
    xl.setup_active_excel()
    xl.set_axis_title(axis=axis, title=title)


def set_graph_title(title):
    xl = Excel()
    xl.setup_active_excel()
    xl.set_graph_title(title=title)


def set_ticks(axis, minimum, maximum, resolution):
    xl = Excel()
    xl.setup_active_excel()
    xl.set_tick(
        axis=axis, minimum=minimum, maximum=maximum, resolution=resolution
    )


def set_line_format(fill):
    xl = Excel()
    xl.setup_active_excel()
    xl.set_line_format(fill)


def make_graph(graph_parameter: List[GraphParameter]):
    xl = Excel()
    xl.setup_active_excel()
    xl.delete_shape()
    for gp in graph_parameter:
        xl.add_chart(graph_type=gp.graph_type, graph_range=gp.graph_range)
        xl.set_graph_title(title=gp.graph_title)
        xl.set_axis_title(axis="y", title=gp.axis_y_title)
        xl.set_tick(
            axis="y",
            minimum=gp.axis_y_ticks[0],
            maximum=gp.axis_y_ticks[1],
            resolution=gp.axis_y_ticks[2],
        )
        xl.set_line_format(0)

    xl.relocate_graph()
