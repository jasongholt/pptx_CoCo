# Native Chart Patterns for python-pptx

Complete, copy-paste recipes for generating native editable PowerPoint charts from Python data structures. All charts use OOXML chart objects -- fully editable in PowerPoint and Google Slides.

## Required Imports

```python
import io
import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData, BubbleChartData, XyChartData
from lxml import etree
```

## Presentation Setup

Always use widescreen 16:9 with blank layouts:

```python
SW = 13.333  # slide width in inches (16:9 widescreen)
SH = 7.5     # slide height

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
```

## Generic Helper Functions

### add_rect -- colored rectangle background

```python
def add_rect(slide, l, t, w, h, rgb):
    """Add a filled rectangle. l/t/w/h in inches, rgb is RGBColor."""
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.line.fill.background()
    s.fill.solid()
    s.fill.fore_color.rgb = rgb
    return s
```

### add_text -- text box with formatting

```python
def add_text(slide, text, l, t, w, h, size, bold=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    """Add a text box. size is font size in points."""
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txb
```

### slide_header -- create a new slide with a dark header bar

```python
def slide_header(prs, header, sw):
    """Create a blank slide with a dark header bar and title text."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, sw, 0.55, RGBColor(0x26, 0x27, 0x30))
    add_text(s, header, 0.2, 0.05, sw - 0.4, 0.45, 15, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    return s
```

### to_bytes -- serialize presentation to bytes

```python
def to_bytes(prs):
    """Return presentation as bytes for download or stage upload."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()
```

## Slide Recipes

### Title Slide

```python
def build_title_slide(prs, sw, title, subtitle, date_str=None, footer=None):
    BLUE = RGBColor(0x29, 0xB5, 0xE8)
    DARK = RGBColor(0x26, 0x27, 0x30)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    GRAY = RGBColor(0x88, 0x88, 0x88)

    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, sw, 7.5, BLUE)
    add_rect(s, 0, 5.8, sw, 1.7, DARK)
    add_text(s, title, 0.7, 1.6, sw - 1.4, 1.1, 40, bold=True, color=WHITE)
    add_text(s, subtitle, 0.7, 2.75, sw - 1.4, 0.7, 22, color=WHITE)
    if date_str:
        add_text(s, date_str, 0.7, 3.55, 4, 0.5, 16, color=RGBColor(0xCC, 0xEE, 0xF8))
    if footer:
        add_text(s, footer, 0.7, 6.15, sw - 1.4, 0.5, 13, color=GRAY)
    return s
```

### KPI Summary Slide

Accepts a list of `(label, formatted_value, color)` tuples. Supports 1-8 KPIs arranged in rows of 4.

```python
def build_kpi_slide(prs, sw, title, kpis):
    """
    kpis: list of (label: str, value: str, color: RGBColor)
    Arranges up to 8 KPIs in rows of 4 with colored value text.
    """
    DARK = RGBColor(0x26, 0x27, 0x30)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    LGRAY = RGBColor(0xF0, 0xF2, 0xF6)

    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, sw, 0.55, DARK)
    add_text(s, title, 0.2, 0.05, sw - 0.4, 0.45, 15, bold=True, color=WHITE)

    cols = 4
    box_gap = 0.13
    box_w = (sw - 0.36 - (cols - 1) * box_gap) / cols

    rows = [kpis[i:i+cols] for i in range(0, len(kpis), cols)]
    for row_idx, row_kpis in enumerate(rows):
        box_top = 0.72 + row_idx * 1.13
        for i, (label, value, val_color) in enumerate(row_kpis):
            lx = 0.18 + i * (box_w + box_gap)
            add_rect(s, lx, box_top, box_w, 0.95, LGRAY)
            add_text(s, value, lx+0.1, box_top+0.04, box_w-0.2, 0.5, 20,
                     bold=True, color=val_color, align=PP_ALIGN.CENTER)
            add_text(s, label, lx+0.1, box_top+0.56, box_w-0.2, 0.3, 9,
                     color=DARK, align=PP_ALIGN.CENTER)
    return s
```

### Bar Chart Slide (Clustered Column)

```python
def build_bar_chart_slide(prs, sw, title, categories, values, colors=None,
                          series_name="Value", show_labels=True, number_format="#,##0"):
    """
    categories: list of str
    values: list of float
    colors: optional list of RGBColor (one per category for per-bar coloring)
    """
    from references import set_point_color  # see ooxml-color-helpers.md

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, [float(v) for v in values])

    s = slide_header(prs, title, sw)
    chart_frame = s.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.18), Inches(0.65), Inches(sw - 0.36), Inches(6.6),
        chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.gap_width = 80

    if colors:
        series_el = plot.series[0]._element
        for idx, clr in enumerate(colors):
            set_point_color(series_el, idx, clr)

    if show_labels:
        plot.series[0].data_labels.show_value = True
        plot.series[0].data_labels.font.size = Pt(9)
        plot.series[0].data_labels.font.bold = True
        plot.series[0].data_labels.number_format = number_format

    return s
```

### Multi-Series Line Chart Slide

```python
def build_line_chart_slide(prs, sw, title, categories, series_dict, color_map=None):
    """
    categories: list of str (x-axis labels, e.g. month names)
    series_dict: dict of {series_name: [values]}. Use None for missing points.
    color_map: optional dict of {series_name: RGBColor}
    """
    from references import set_series_color  # see ooxml-color-helpers.md

    line_data = CategoryChartData()
    line_data.categories = categories
    for name, vals in series_dict.items():
        line_data.add_series(name, vals)

    s = slide_header(prs, title, sw)
    chart_frame = s.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(0.18), Inches(0.65), Inches(sw - 0.36), Inches(6.6),
        line_data
    )
    chart = chart_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(8)

    if color_map:
        for i, name in enumerate(series_dict.keys()):
            if name in color_map:
                ser = chart.series[i]
                set_series_color(ser, color_map[name])
                ser.smooth = False
                ser.format.line.width = Pt(2)

    return s
```

### Bubble Chart Slide

```python
def build_bubble_chart_slide(prs, sw, title, bubble_points, color_map=None):
    """
    bubble_points: list of (series_name, x_value, y_value, bubble_size)
    color_map: optional dict of {series_name: RGBColor}
    """
    from references import set_series_color

    bubble_data = BubbleChartData()
    for name, x, y, size in bubble_points:
        ser = bubble_data.add_series(name)
        ser.add_data_point(float(x), float(y), float(size))

    s = slide_header(prs, title, sw)
    chart_frame = s.shapes.add_chart(
        XL_CHART_TYPE.BUBBLE,
        Inches(0.18), Inches(0.65), Inches(sw - 0.36), Inches(6.6),
        bubble_data
    )
    chart = chart_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(8)

    if color_map:
        for i, (name, _, _, _) in enumerate(bubble_points):
            if name in color_map:
                set_series_color(chart.series[i], color_map[name])

    return s
```

### Pie Chart Slide

```python
def build_pie_chart_slide(prs, sw, title, labels, values, colors=None):
    """
    labels: list of str (slice labels)
    values: list of float (slice values)
    colors: optional list of RGBColor (one per slice)
    """
    from references import set_point_color

    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("Values", [float(v) for v in values])

    s = slide_header(prs, title, sw)
    chart_frame = s.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(1.5), Inches(0.65), Inches(sw - 3.0), Inches(6.6),
        chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.font.size = Pt(9)

    plot = chart.plots[0]
    plot.series[0].data_labels.show_percentage = True
    plot.series[0].data_labels.show_category_name = True
    plot.series[0].data_labels.font.size = Pt(9)

    if colors:
        series_el = plot.series[0]._element
        for idx, clr in enumerate(colors):
            set_point_color(series_el, idx, clr)

    return s
```

### Smooth Line Chart Slide (Single Series)

Useful for probability curves, trend lines, or forecasts.

```python
def build_smooth_line_slide(prs, sw, title, x_labels, y_values, line_color):
    """
    x_labels: list of str (category labels)
    y_values: list of float
    line_color: RGBColor
    """
    from references import set_series_color

    chart_data = CategoryChartData()
    chart_data.categories = x_labels
    chart_data.add_series("Value", [round(float(v), 2) for v in y_values])

    s = slide_header(prs, title, sw)
    chart_frame = s.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.18), Inches(0.65), Inches(sw - 0.36), Inches(6.6),
        chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = False
    set_series_color(chart.series[0], line_color)
    chart.series[0].smooth = True
    chart.series[0].format.line.width = Pt(3)

    return s
```

### Text Slide (AI Summary, Notes, etc.)

```python
def build_text_slide(prs, sw, title, body_text):
    DARK = RGBColor(0x26, 0x27, 0x30)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, sw, 0.55, DARK)
    add_text(s, title, 0.2, 0.05, sw - 0.4, 0.45, 15, bold=True, color=WHITE)
    add_text(s, body_text, 0.3, 0.75, sw - 0.6, 6.5, 11, color=DARK, wrap=True)
    return s
```

## Full Example: Data to Deck in One Function

```python
def build_deck(title, subtitle, kpis, charts, ai_summary=None):
    """
    title: str
    subtitle: str
    kpis: list of (label, value, color)
    charts: list of dicts with keys: type, title, data, colors
    ai_summary: optional str
    """
    SW = 13.333
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(7.5)

    build_title_slide(prs, SW, title, subtitle,
                      date_str=datetime.date.today().strftime("%B %Y"),
                      footer="Powered by Snowflake + Cortex AI")

    if kpis:
        build_kpi_slide(prs, SW, "Key Metrics", kpis)

    for chart_spec in charts:
        ctype = chart_spec["type"]
        if ctype == "bar":
            build_bar_chart_slide(prs, SW, chart_spec["title"],
                                  chart_spec["categories"], chart_spec["values"],
                                  colors=chart_spec.get("colors"))
        elif ctype == "line":
            build_line_chart_slide(prs, SW, chart_spec["title"],
                                   chart_spec["categories"], chart_spec["series"],
                                   color_map=chart_spec.get("color_map"))
        elif ctype == "bubble":
            build_bubble_chart_slide(prs, SW, chart_spec["title"],
                                     chart_spec["points"],
                                     color_map=chart_spec.get("color_map"))
        elif ctype == "pie":
            build_pie_chart_slide(prs, SW, chart_spec["title"],
                                  chart_spec["labels"], chart_spec["values"],
                                  colors=chart_spec.get("colors"))
        elif ctype == "smooth_line":
            build_smooth_line_slide(prs, SW, chart_spec["title"],
                                    chart_spec["x_labels"], chart_spec["y_values"],
                                    chart_spec["line_color"])

    if ai_summary:
        build_text_slide(prs, SW, "AI-Generated Executive Summary (Snowflake Cortex)", ai_summary)

    return to_bytes(prs)
```

## Color Palettes

### Snowflake-inspired (dark theme)

```python
BLUE   = RGBColor(0x29, 0xB5, 0xE8)
DARK   = RGBColor(0x26, 0x27, 0x30)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY  = RGBColor(0xF0, 0xF2, 0xF6)
RED    = RGBColor(0xD9, 0x40, 0x32)
GREEN  = RGBColor(0x2E, 0xCC, 0x71)
AMBER  = RGBColor(0xFF, 0xBE, 0x2E)
MGRAY  = RGBColor(0x88, 0x88, 0x88)
```

### Categorical palette (up to 8 categories)

```python
CATEGORICAL = [
    RGBColor(0x29, 0xB5, 0xE8),  # blue
    RGBColor(0xFF, 0xBE, 0x2E),  # amber
    RGBColor(0xD9, 0x40, 0x32),  # red
    RGBColor(0x2E, 0xCC, 0x71),  # green
    RGBColor(0xFF, 0x6B, 0x6B),  # coral
    RGBColor(0x9B, 0x59, 0xB6),  # purple
    RGBColor(0x1A, 0xBC, 0x9C),  # teal
    RGBColor(0xF3, 0x9C, 0x12),  # orange
]
```
