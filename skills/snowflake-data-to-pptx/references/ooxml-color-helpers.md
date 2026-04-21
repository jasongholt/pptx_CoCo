# OOXML Color Helpers for python-pptx

python-pptx does not natively support per-point bar coloring or per-series line/bubble coloring. These helpers manipulate the underlying OOXML XML directly via `lxml.etree`.

## Namespace Map

Every OOXML chart lives inside these two namespaces:

```python
from lxml import etree

NSMAP = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
```

## set_point_color

Colors individual data points (bars, pie slices) within a single series. Use this when you want each bar in a bar chart to be a different color.

```python
def set_point_color(series_element, idx, rgb):
    """
    Color a single data point within a chart series.

    Args:
        series_element: The lxml element of the series (plot.series[0]._element)
        idx: Zero-based index of the data point to color
        rgb: RGBColor instance or string like "29B5E8"
    """
    dPt = etree.SubElement(series_element, etree.QName(NSMAP["c"], "dPt"))
    idx_el = etree.SubElement(dPt, etree.QName(NSMAP["c"], "idx"))
    idx_el.set("val", str(idx))
    spPr = etree.SubElement(dPt, etree.QName(NSMAP["c"], "spPr"))
    solidFill = etree.SubElement(spPr, etree.QName(NSMAP["a"], "solidFill"))
    srgbClr = etree.SubElement(solidFill, etree.QName(NSMAP["a"], "srgbClr"))
    srgbClr.set("val", str(rgb))
```

### Usage: Color each bar differently

```python
from pptx.dml.color import RGBColor

colors = [RGBColor(0x2E, 0xCC, 0x71), RGBColor(0xFF, 0xBE, 0x2E), RGBColor(0xD9, 0x40, 0x32)]
series_el = chart.plots[0].series[0]._element
for idx, clr in enumerate(colors):
    set_point_color(series_el, idx, clr)
```

## set_series_color

Colors an entire series (line, bubble, or bar series in a multi-series chart). Sets both the fill and the line/border color.

```python
def set_series_color(series, rgb):
    """
    Color an entire chart series (fill + line).

    Args:
        series: A python-pptx Series object (chart.series[i])
        rgb: RGBColor instance or string like "29B5E8"
    """
    spPr = series._element.find(etree.QName(NSMAP["c"], "spPr"))
    if spPr is None:
        spPr = etree.SubElement(series._element, etree.QName(NSMAP["c"], "spPr"))

    for old in spPr.findall(etree.QName(NSMAP["a"], "solidFill")):
        spPr.remove(old)
    solidFill = etree.SubElement(spPr, etree.QName(NSMAP["a"], "solidFill"))
    srgbClr = etree.SubElement(solidFill, etree.QName(NSMAP["a"], "srgbClr"))
    srgbClr.set("val", str(rgb))

    ln = spPr.find(etree.QName(NSMAP["a"], "ln"))
    if ln is None:
        ln = etree.SubElement(spPr, etree.QName(NSMAP["a"], "ln"))
    for old in ln.findall(etree.QName(NSMAP["a"], "solidFill")):
        ln.remove(old)
    ln_fill = etree.SubElement(ln, etree.QName(NSMAP["a"], "solidFill"))
    ln_clr = etree.SubElement(ln_fill, etree.QName(NSMAP["a"], "srgbClr"))
    ln_clr.set("val", str(rgb))
```

### Usage: Color each line in a multi-series line chart

```python
color_map = {"PRIME": RGBColor(0x29, 0xB5, 0xE8), "SUBPRIME": RGBColor(0xD9, 0x40, 0x32)}
for i, name in enumerate(series_names):
    ser = chart.series[i]
    set_series_color(ser, color_map.get(name, RGBColor(0x88, 0x88, 0x88)))
    ser.smooth = False
    ser.format.line.width = Pt(2)
```

### Usage: Color each bubble in a bubble chart

```python
for i, row in enumerate(scatter_data.iterrows()):
    clr = color_map.get(row["CATEGORY"], RGBColor(0x88, 0x88, 0x88))
    set_series_color(chart.series[i], clr)
```

## When to Use Which

| Scenario | Function | Access Pattern |
|---|---|---|
| Single-series bar chart, each bar a different color | `set_point_color` | `plot.series[0]._element` |
| Multi-series bar chart, each series one color | `set_series_color` | `chart.series[i]` |
| Line chart with colored lines | `set_series_color` | `chart.series[i]` |
| Bubble chart with colored bubbles | `set_series_color` | `chart.series[i]` |
| Pie chart with colored slices | `set_point_color` | `plot.series[0]._element` |

## Important Notes

- `RGBColor.__str__()` returns the hex string (e.g., `"29B5E8"`) which is what OOXML expects for `srgbClr@val`.
- `set_point_color` operates on the raw `_element` (lxml Element), not the python-pptx Series object.
- `set_series_color` operates on the python-pptx Series object (it accesses `._element` internally).
- Always set colors AFTER adding the chart to the slide; the series elements don't exist until the chart is rendered.
- For data labels styling, use the python-pptx API directly -- no XML manipulation needed:
  ```python
  plot.series[0].data_labels.show_value = True
  plot.series[0].data_labels.font.size = Pt(9)
  plot.series[0].data_labels.font.bold = True
  plot.series[0].data_labels.number_format = "#,##0"
  ```
