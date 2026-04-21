# Stage Persistence and Download Patterns

How to save generated PPTX files to Snowflake stages and provide download URLs across all runtimes.

## Prerequisites

Create a stage with directory enabled (required for `GET_PRESIGNED_URL`):

```sql
CREATE STAGE IF NOT EXISTS MY_STAGE
    DIRECTORY = (ENABLE = TRUE)
    ENCRYPTION = (TYPE = 'SNOWFLAKE_SSE');
```

## Save to Stage (Snowpark Session)

```python
import io

buf = io.BytesIO()
prs.save(buf)
buf.seek(0)
pptx_bytes = buf.getvalue()

tmp_path = '/tmp/my_deck.pptx'
with open(tmp_path, 'wb') as f:
    f.write(pptx_bytes)

session.file.put(tmp_path, '@MY_DB.PUBLIC.MY_STAGE', auto_compress=False, overwrite=True)
```

### Refresh the directory listing

After PUT, refresh the stage directory so `GET_PRESIGNED_URL` and `DIRECTORY()` queries work:

```python
session.sql("ALTER STAGE MY_DB.PUBLIC.MY_STAGE REFRESH").collect()
```

### Verify the upload

```python
result = session.sql("""
    SELECT RELATIVE_PATH, SIZE, LAST_MODIFIED
    FROM DIRECTORY(@MY_DB.PUBLIC.MY_STAGE)
    WHERE RELATIVE_PATH LIKE '%.pptx'
""").to_pandas()
print(result.to_string(index=False))
```

## Generate Presigned Download URL

Presigned URLs are temporary, time-limited download links that work without Snowflake authentication:

```python
url = session.sql("""
    SELECT GET_PRESIGNED_URL(
        @MY_DB.PUBLIC.MY_STAGE,
        'my_deck.pptx',
        3600  -- expiry in seconds (1 hour)
    ) AS URL
""").to_pandas()["URL"].iloc[0]

print(f"Download: {url}")
```

## Delivery by Runtime

### Streamlit (Local or SiS)

Use `st.download_button` -- works identically in both runtimes:

```python
import streamlit as st

st.download_button(
    label="Download Presentation",
    data=pptx_bytes,
    file_name="my_deck.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
)
```

No stage upload required for Streamlit -- the download button serves bytes directly from memory. Stage upload is optional for archival.

### Jupyter Notebook (Local)

Option A -- Presigned URL (works everywhere, requires stage):
```python
from IPython.display import display, HTML

display(HTML(f'<a href="{url}" target="_blank">Download from Snowflake Stage</a>'))
```

Option B -- Local file link (only works in local Jupyter):
```python
from IPython.display import FileLink

tmp_path = '/tmp/my_deck.pptx'
with open(tmp_path, 'wb') as f:
    f.write(pptx_bytes)

display(FileLink(tmp_path))
```

Option C -- Both (recommended):
```python
from IPython.display import display, HTML, FileLink

tmp_path = '/tmp/my_deck.pptx'
with open(tmp_path, 'wb') as f:
    f.write(pptx_bytes)

session.file.put(tmp_path, '@MY_STAGE', auto_compress=False, overwrite=True)
session.sql("ALTER STAGE MY_STAGE REFRESH").collect()
url = session.sql("SELECT GET_PRESIGNED_URL(@MY_STAGE, 'my_deck.pptx', 3600) AS URL").to_pandas()["URL"].iloc[0]

display(HTML(f'<a href="{url}" target="_blank">Download from Snowflake Stage</a>'))
display(FileLink(tmp_path))
print(f"Also saved locally: {tmp_path}")
```

### Snowflake Notebook (Snowsight)

In Snowflake container notebooks, local file links don't work. Use presigned URL only:

```python
from IPython.display import display, HTML

session.file.put('/tmp/my_deck.pptx', '@MY_STAGE', auto_compress=False, overwrite=True)
session.sql("ALTER STAGE MY_STAGE REFRESH").collect()
url = session.sql("SELECT GET_PRESIGNED_URL(@MY_STAGE, 'my_deck.pptx', 3600) AS URL").to_pandas()["URL"].iloc[0]

display(HTML(f'<a href="{url}" target="_blank" style="font-size:16px;">📥 Download Presentation</a>'))
```

## Complete End-to-End Pattern

```python
import io
from IPython.display import display, HTML, FileLink

buf = io.BytesIO()
prs.save(buf)
buf.seek(0)
pptx_bytes = buf.getvalue()

tmp_path = '/tmp/my_deck.pptx'
with open(tmp_path, 'wb') as f:
    f.write(pptx_bytes)

stage = '@MY_DB.PUBLIC.MY_STAGE'
filename = 'my_deck.pptx'

session.file.put(tmp_path, stage, auto_compress=False, overwrite=True)
session.sql(f"ALTER STAGE {stage.lstrip('@')} REFRESH").collect()

url = session.sql(f"""
    SELECT GET_PRESIGNED_URL({stage}, '{filename}', 3600) AS URL
""").to_pandas()["URL"].iloc[0]

print(f"Presentation: {len(pptx_bytes)/1024:.0f} KB, {len(prs.slides)} slides")
display(HTML(f'<a href="{url}" target="_blank">Download from Snowflake Stage</a>'))

try:
    display(FileLink(tmp_path))
except Exception:
    pass  # FileLink not available in Snowflake container
```

## Common Issues

| Issue | Cause | Fix |
|---|---|---|
| `GET_PRESIGNED_URL` returns error | Stage directory not enabled | `ALTER STAGE ... SET DIRECTORY = (ENABLE = TRUE)` |
| File not found after PUT | Directory not refreshed | Run `ALTER STAGE ... REFRESH` after PUT |
| `auto_compress=True` (default) renames file | Adds `.gz` extension | Always use `auto_compress=False` for PPTX |
| URL expires too quickly | Short expiry value | Use `3600` (1 hour) or `86400` (24 hours) |
| `/tmp/` not writable | Permissions issue (rare) | Use `tempfile.mktemp(suffix='.pptx')` instead |
| Large file upload slow | File > 10MB | Native charts are tiny (~5-10KB each); check if images are embedded |
