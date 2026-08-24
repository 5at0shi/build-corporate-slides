from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.util import Pt

from .theme import PALETTE

_CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}

# 系列・要素の配色はsemantic paletteを巡回させ、デッキ全体の配色と揃える。
_SERIES_COLORS = [PALETTE.blue, PALETTE.accent_secondary, PALETTE.grey_500,
                  PALETTE.navy]


def add_native_chart(slide, x, y, w, h, *, chart_type, categories, series,
                     typography, value_format=None):
    """編集可能なネイティブPowerPointグラフを描く。

    chart_typeは column/bar/line/pie。seriesは
    [{"name": ..., "values": [...]}] の配列。PNG画像と異なり、貼り付け後も
    PowerPoint上で数値・系列名を直接編集できる。
    """
    if chart_type not in _CHART_TYPES:
        raise ValueError(f"未対応のchart_typeです: {chart_type}")
    if not series:
        raise ValueError("seriesが空です")

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for s in series:
        chart_data.add_series(s.get("name", ""), s.get("values", []))

    graphic_frame = slide.shapes.add_chart(
        _CHART_TYPES[chart_type], x, y, w, h, chart_data)
    chart = graphic_frame.chart
    chart.has_title = False
    font = typography.body_font
    small = typography.small

    if chart_type == "pie":
        _style_pie(chart, font, small)
    else:
        _style_category_chart(chart, chart_type, font, small, value_format,
                              multi_series=len(series) > 1)
    return graphic_frame


def _style_pie(chart, font, small):
    plot = chart.plots[0]
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.show_category_name = True
    labels.show_percentage = True
    labels.show_value = False
    labels.number_format = "0%"
    labels.number_format_is_linked = False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    labels.font.size = small
    labels.font.name = font
    labels.font.color.rgb = PALETTE.text_primary
    chart.has_legend = False

    points = plot.series[0].points
    for index, point in enumerate(points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = (
            _SERIES_COLORS[index % len(_SERIES_COLORS)])
        point.format.line.color.rgb = PALETTE.surface_base
        point.format.line.width = Pt(1.5)


def _style_category_chart(chart, chart_type, font, small, value_format,
                          *, multi_series):
    plot = chart.plots[0]
    plot.gap_width = 60
    if chart_type != "line":
        plot.overlap = -8 if multi_series else 0

    chart.has_legend = multi_series
    if multi_series:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = small
        chart.legend.font.name = font

    plot.has_data_labels = not multi_series
    if plot.has_data_labels:
        labels = plot.data_labels
        labels.font.size = small
        labels.font.name = font
        labels.font.color.rgb = PALETTE.text_primary
        if value_format:
            labels.number_format = value_format
            labels.number_format_is_linked = False
        if chart_type != "line":
            labels.position = XL_LABEL_POSITION.OUTSIDE_END

    for index, s in enumerate(plot.series):
        color = _SERIES_COLORS[index % len(_SERIES_COLORS)]
        if chart_type == "line":
            s.format.line.color.rgb = color
            s.format.line.width = Pt(2.25)
            s.smooth = False
        else:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = color
            s.format.line.fill.background()

    category_axis = chart.category_axis
    category_axis.has_major_gridlines = False
    category_axis.tick_labels.font.size = small
    category_axis.tick_labels.font.name = font
    category_axis.format.line.color.rgb = PALETTE.line_neutral

    value_axis = chart.value_axis
    value_axis.format.line.fill.background()
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = PALETTE.line_neutral
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.tick_labels.font.size = small
    value_axis.tick_labels.font.name = font
    if value_format:
        value_axis.tick_labels.number_format = value_format
        value_axis.tick_labels.number_format_is_linked = False
