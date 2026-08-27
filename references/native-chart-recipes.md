# Native Chart Recipes

All recipes create editable PowerPoint charts with embedded Excel workbook data. Import only the chart-data class that matches the chart type.

## Restrained chart defaults

- Do not add a chart title when the slide title already identifies the measure.
- Do not add axis titles unless the units or measure would otherwise be ambiguous.
- Remove major and minor gridlines by default.
- For columns, bars, and ranked charts, use direct value labels and hide the value axis when the labels carry the complete scale.
- Put labels outside the end of ordinary columns and bars; put them inside pie, donut, and stacked segments when they remain legible.
- Keep axes for line, area, combination, scatter, and bubble charts when the audience needs the continuous scale.
- Keep a legend only when multiple series cannot be identified directly. Never show a legend that duplicates direct labels.
- Do not use 3-D effects, decorative backgrounds, shadows, or images as chart substitutes.

## Category column or bar chart

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

chart_data = CategoryChartData(number_format="#,##0")
chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
chart_data.add_series("Revenue", (12.4, 14.2, 13.8, 16.1))
chart_data.add_series("Target", (11.5, 13.0, 14.0, 15.0))

graphic_frame = placeholder.insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, chart_data)
chart = graphic_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
```

Use `XL_CHART_TYPE.BAR_CLUSTERED` for horizontal bars.

## Line chart

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

chart_data = CategoryChartData(number_format="0.0")
chart_data.categories = ["Jan", "Feb", "Mar", "Apr"]
chart_data.add_series("North", (4.2, 5.1, 5.8, 6.0))
chart_data.add_series("South", (3.9, 4.3, 4.8, 5.4))

chart = placeholder.insert_chart(XL_CHART_TYPE.LINE_MARKERS, chart_data).chart
chart.value_axis.has_major_gridlines = True
```

## Pie chart

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

chart_data = CategoryChartData(number_format="0%")
chart_data.categories = ["Enterprise", "SMB", "Consumer"]
chart_data.add_series("Share", (0.52, 0.31, 0.17))

chart = placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data).chart
plot = chart.plots[0]
plot.has_data_labels = True
plot.data_labels.number_format = "0%"
plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
```

## XY scatter chart

```python
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE

chart_data = XyChartData(number_format="0.0")
series = chart_data.add_series("Efficiency")
for x, y in [(1.0, 2.4), (2.0, 3.8), (3.0, 4.2)]:
    series.add_data_point(x, y)

chart = placeholder.insert_chart(XL_CHART_TYPE.XY_SCATTER_LINES, chart_data).chart
```

## Bubble chart

```python
from pptx.chart.data import BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE

chart_data = BubbleChartData(number_format="0.0")
series = chart_data.add_series("Markets")
series.add_data_point(1.0, 2.0, 20)
series.add_data_point(2.0, 3.1, 35)
series.add_data_point(3.0, 2.7, 28)

chart = placeholder.insert_chart(XL_CHART_TYPE.BUBBLE, chart_data).chart
```

## Replace data in an existing chart

```python
from pptx.chart.data import CategoryChartData

chart_data = CategoryChartData(number_format="#,##0")
chart_data.categories = ["A", "B", "C"]
chart_data.add_series("Current", (10, 14, 12))

for slide in prs.slides:
    for shape in slide.shapes:
        if getattr(shape, "has_chart", False):
            shape.chart.replace_data(chart_data)
```

Replacement data can have different category and series dimensions from the original chart.
