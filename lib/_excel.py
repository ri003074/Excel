def _set_axis_title(shape, axis_type, title):
    axis = shape.Chart.Axes(axis_type)
    axis.HasTitle = True
    axis.AxisTitle.Characters.Text = title


def _set_graph_title(shape, title):
    shape.Chart.HasTitle = True
    shape.Chart.ChartTitle.Text = title


def _set_axis_obj(shape, axis_type, minimum, maximum, resolution):
    axis_obj = shape.Chart.Axes(axis_type)
    axis_obj.MinimumScale = minimum
    axis_obj.MaximumScale = maximum
    axis_obj.MajorUnit = resolution
