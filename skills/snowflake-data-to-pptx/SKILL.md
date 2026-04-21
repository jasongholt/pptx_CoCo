---
name: snowflake-data-to-pptx
description: "Generate native-chart PowerPoint presentations from Snowflake data. Use when: creating PPTX from query results, exporting dashboards to PowerPoint, building slide decks from tables, native editable charts, python-pptx with Snowflake. Triggers: pptx from data, powerpoint from snowflake, export to pptx, slide deck from table, data to presentation, native chart, editable chart, python-pptx snowflake, generate powerpoint, export dashboard to pptx, snowflake to slides."
---

# Snowflake Data to PowerPoint

Generate editable PowerPoint presentations with **native OOXML charts** from Snowflake query results. Charts are fully editable in PowerPoint and Google Slides (not images).

This skill complements the generic `pptx` skill. Use this skill when data lives in Snowflake; use the generic `pptx` skill for template-based editing, comments, or speaker notes.

## Dependencies

```
python-pptx>=1.0
lxml
```

These are standard in most Snowflake environments. For Streamlit in Snowflake, add them to `pyproject.toml`.

## Workflow

### Step 1: Discover the Data

Before writing any PPTX code, understand the data shape:

```sql
SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'MY_SCHEMA' AND TABLE_NAME = 'MY_TABLE'
ORDER BY ORDINAL_POSITION;
```

Classify each column:

| Column Type | Role in Charts | Example |
|---|---|---|
| VARCHAR/TEXT with few distinct values | Category axis, grouping, legend series | RISK_TIER, REGION, STATUS |
| DATE/TIMESTAMP | Time axis for line/area charts | ORIGINATION_MONTH, CREATED_DATE |
| NUMBER/FLOAT | Value axis, KPI metric, bubble size | TOTAL_AMOUNT, COUNT, RATE_PCT |
| VARCHAR with many distinct values | Labels, not suitable for axes | CUSTOMER_NAME, ID |

Ask the user:
1. What story should the deck tell?
2. Which columns are the key metrics (KPIs)?
3. Which columns define groups/segments?
4. Should the deck include an AI-generated summary slide?

### Step 2: Plan the Slide Deck

Propose a structure before building. A typical data deck:

| Slide | Type | When to Include |
|---|---|---|
| Title | Always | Every deck |
| KPI Summary | When there are 4-8 key numeric metrics | Most decks |
| Bar Chart | Categories + numeric values | Comparing groups |
| Line Chart | Time series data | Trends over time |
| Bubble Chart | 2 numeric axes + size dimension | Risk/performance mapping |
| Pie Chart | Parts of a whole (max 8 slices) | Composition analysis |
| Scatter Plot | Two continuous variables | Correlation |
| AI Summary | When Cortex AI is available | Executive decks |

Present the plan to the user and get approval before generating code.

### Step 3: Generate the PPTX

Use `python-pptx` to create native chart objects. See `references/native-chart-patterns.md` for complete recipes.

**Chart type selection:**

| Data Pattern | Chart Type | python-pptx Enum |
|---|---|---|
| Categories + single value | Clustered bar | `XL_CHART_TYPE.COLUMN_CLUSTERED` |
| Categories + multiple series | Grouped bar | `XL_CHART_TYPE.COLUMN_CLUSTERED` (multi-series) |
| Categories + stacked values | Stacked bar | `XL_CHART_TYPE.COLUMN_STACKED` |
| Time series + value | Line with markers | `XL_CHART_TYPE.LINE_MARKERS` |
| Time series + groups | Multi-series line | `XL_CHART_TYPE.LINE_MARKERS` (multi-series) |
| 2 numeric axes + bubble size | Bubble | `XL_CHART_TYPE.BUBBLE` |
| Parts of a whole | Pie | `XL_CHART_TYPE.PIE` |
| Continuous x vs y | Scatter | `XL_CHART_TYPE.XY_SCATTER` |
| Smooth curve | Smooth line | `XL_CHART_TYPE.LINE` with `series.smooth = True` |

**Key principles:**
- Always use `Presentation()` with `slide_layouts[6]` (blank layout) for full control
- Use widescreen 16:9: `prs.slide_width = Inches(13.333)`, `prs.slide_height = Inches(7.5)`
- Charts should fill most of the slide: position at `Inches(0.18), Inches(0.65)`, size `Inches(SW-0.36), Inches(6.6)`
- Use OOXML XML manipulation for per-point and per-series colors (see `references/ooxml-color-helpers.md`)
- Return bytes via `io.BytesIO` for maximum flexibility

### Step 4: Persist and Deliver

See `references/stage-persistence.md` for complete patterns.

**In Streamlit:**
```python
st.download_button(
    "Download Presentation",
    data=pptx_bytes,
    file_name="my_deck.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
)
```

**In Notebook:**
```python
tmp_path = '/tmp/my_deck.pptx'
with open(tmp_path, 'wb') as f:
    f.write(pptx_bytes)

session.file.put(tmp_path, '@MY_STAGE', auto_compress=False, overwrite=True)

url = session.sql(
    "SELECT GET_PRESIGNED_URL(@MY_STAGE, 'my_deck.pptx', 3600) AS URL"
).to_pandas()["URL"].iloc[0]

from IPython.display import display, HTML, FileLink
display(HTML(f'<a href="{url}" target="_blank">Download from Snowflake Stage</a>'))
display(FileLink(tmp_path))
```

### Step 5: Cortex AI Summary Slide (Optional)

If the user wants an AI-generated executive summary:

```python
import json

data_context = json.dumps(df.to_dict(orient='records'), default=str)
prompt = (
    f"You are a senior analyst. Summarize this data in 5 bullet points. "
    f"Highlight key trends, risks, and recommended actions.\n\n{data_context}"
).replace("'", "''")

model = "claude-sonnet-4-6"
result = session.sql(
    f"SELECT SNOWFLAKE.CORTEX.COMPLETE('{model}', '{prompt}') AS SUMMARY"
).to_pandas()
ai_text = result["SUMMARY"].iloc[0]
```

Add to the deck as a text slide:
```python
s = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(s, 0, 0, SW, 0.55, DARK)
add_text(s, "AI-Generated Executive Summary (Snowflake Cortex)", 0.2, 0.05, SW-0.4, 0.45, 15, bold=True, color=WHITE)
add_text(s, ai_text, 0.3, 0.75, SW-0.6, 6.5, 11, color=DARK, wrap=True)
```

## Connection Patterns

See `references/dual-runtime-connection.md` for patterns that work in:
- Local Jupyter / VS Code notebooks
- Snowflake container notebooks (Snowsight)
- Streamlit in Snowflake (SiS)
- Local Streamlit (`streamlit run`)

## Common Pitfalls

| Pitfall | Solution |
|---|---|
| Charts render as images, not editable | Use `python-pptx` chart objects, not Plotly `.to_image()` |
| All bars/lines are the same color | Use `set_point_color` or `set_series_color` from `references/ooxml-color-helpers.md` |
| `slide_layouts` index error | Always use `slide_layouts[6]` (blank). Other indices vary by template. |
| `None` values crash chart data | Filter or replace None before adding to `CategoryChartData` |
| `session.file.put()` fails in Streamlit | Write to `/tmp/` first, then PUT. `/tmp` is writable in all runtimes. |
| PPTX too large | Avoid embedding images. Native charts are tiny (~5-10KB per chart). |
| `GET_PRESIGNED_URL` returns error | Ensure the stage has `DIRECTORY = (ENABLE = TRUE)` and run `ALTER STAGE ... REFRESH` after PUT. |

## File Structure for a PPTX Project

```
my_project/
    config.py              # Database, schema, stage, model name
    streamlit_app.py       # Streamlit version with st.download_button
    notebook.ipynb         # Notebook version with stage upload
    snowflake.yml          # SiS deployment config
    pyproject.toml         # Dependencies: python-pptx, lxml, plotly, etc.
    setup_mock_data.sql    # Optional: create test database with sample data
    .streamlit/
        config.toml        # Theme colors
        secrets.toml       # Local credentials (git-ignored)
    .gitignore             # Exclude secrets, __pycache__, .pptx files
```
