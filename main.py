from app.app import relocate_graph
from app.app import resize_graph
from app.app import save_png
from app.app import set_graph_title
from app.app import set_axis_title
from app.app import set_ticks

if __name__ == "__main__":
    resize_graph(200, 300)
    relocate_graph()
    save_png()
    set_axis_title(axis="y", title="mV")
    set_graph_title(title="test")
    set_ticks(axis="y", minimum=0, maximum=120, resolution=20)
