def _set_axis_title(shape, axis_type, title, font_size=None):
    axis = shape.Chart.Axes(axis_type)
    axis.HasTitle = True
    axis.AxisTitle.Characters.Text = title
    if font_size is not None:
        axis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = font_size


def _set_graph_title(shape, title):
    shape.Chart.HasTitle = True
    shape.Chart.ChartTitle.Text = title


def _set_axis_obj(shape, axis_type, minimum, maximum, resolution):
    axis_obj = shape.Chart.Axes(axis_type)
    axis_obj.MinimumScale = minimum
    axis_obj.MaximumScale = maximum
    axis_obj.MajorUnit = resolution


def _set_line_format(xl, shape, fill):
    for i in range(1, xl.ActiveChart.SeriesCollection().Count + 1):
        shape.Chart.SeriesCollection(i).Format.Line.Visible = fill
