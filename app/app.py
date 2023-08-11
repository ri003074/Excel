from lib.excel import Excel
from lib.excel import GraphParameter


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


def make_graph(graph_parameter: GraphParameter):
    xl = Excel()
    xl.setup_active_excel()
    xl.delete_shape()
    for gp in graph_parameter:
        xl.add_chart(graph_type=gp.graph_type, graph_range=gp.graph_range)
        xl.set_graph_title(title=gp.graph_title)
    xl.relocate_graph()
