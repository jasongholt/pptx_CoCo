# Dual-Runtime Connection Patterns

Patterns for code that works in all Snowflake runtimes: local development, Snowflake container notebooks, and Streamlit in Snowflake (SiS).

## Notebook Pattern (Snowpark Session)

This try/except pattern works in both Snowflake Notebooks (Snowsight) and local Jupyter/VS Code:

```python
import os

try:
    from snowflake.snowpark.context import get_active_session
    session = get_active_session()
    _env = "Snowflake Notebook"
except Exception:
    from snowflake.snowpark import Session
    conn_name = os.getenv("SNOWFLAKE_CONNECTION_NAME", "default")
    session = Session.builder.configs({"connection_name": conn_name}).create()
    session.sql("USE DATABASE MY_DATABASE").collect()
    session.sql("USE SCHEMA MY_SCHEMA").collect()
    _env = f"Local (connection: {conn_name})"

print(f"Connected via {_env}")
```

### How it works

| Runtime | What happens |
|---|---|
| Snowflake Notebook (Snowsight) | `get_active_session()` succeeds, returns the container session. Database/schema are already set from notebook settings. |
| Local Jupyter / VS Code | `get_active_session()` raises `SnowparkSessionException`. Fallback creates a session using `~/.snowflake/connections.toml`. |

### Connection name resolution

The `connection_name` parameter references a named connection in `~/.snowflake/connections.toml`:

```toml
[admin]
account = "myorg-myaccount"
user = "myuser"
authenticator = "externalbrowser"
warehouse = "MY_WH"
database = "MY_DB"
schema = "MY_SCHEMA"
role = "MY_ROLE"
```

Use environment variable override for CI/CD or team sharing:
```bash
SNOWFLAKE_CONNECTION_NAME=staging python my_script.py
```

### With config.py

Centralize the connection name in a config file:

```python
# config.py
CONNECTION_NAME = "admin"
DATABASE = "MY_DB"
SCHEMA = "MY_SCHEMA"
```

```python
# notebook cell 1
import os
import config as cfg

try:
    from snowflake.snowpark.context import get_active_session
    session = get_active_session()
except Exception:
    from snowflake.snowpark import Session
    conn_name = os.getenv("SNOWFLAKE_CONNECTION_NAME", cfg.CONNECTION_NAME)
    session = Session.builder.configs({"connection_name": conn_name}).create()
    session.sql(f"USE DATABASE {cfg.DATABASE}").collect()
    session.sql(f"USE SCHEMA {cfg.SCHEMA}").collect()
```

## Streamlit Pattern (st.connection)

`st.connection("snowflake")` auto-detects the runtime:

```python
import streamlit as st

conn = st.connection("snowflake")
session = conn.session()
```

| Runtime | What happens |
|---|---|
| Streamlit in Snowflake (SiS) | Automatically uses the app's service account. Zero config needed. |
| Local `streamlit run` | Reads from `~/.snowflake/connections.toml` (default connection) or `.streamlit/secrets.toml`. |

### Local Streamlit setup

Option A — Use `connections.toml` (recommended, no secrets in project):
```toml
# ~/.snowflake/connections.toml
[default]
account = "myorg-myaccount"
user = "myuser"
authenticator = "externalbrowser"
```

Option B — Use `.streamlit/secrets.toml` (project-specific):
```toml
# .streamlit/secrets.toml  (add to .gitignore!)
[connections.snowflake]
account = "myorg-myaccount"
user = "myuser"
authenticator = "externalbrowser"
warehouse = "MY_WH"
database = "MY_DB"
schema = "MY_SCHEMA"
```

### Running queries

```python
df = conn.query("SELECT * FROM MY_TABLE LIMIT 100")

session = conn.session()
sf_df = session.sql("SELECT * FROM MY_TABLE")
pandas_df = sf_df.to_pandas()
```

## Python Script Pattern (snowflake.connector)

For standalone scripts (not notebooks, not Streamlit):

```python
import os
import snowflake.connector

conn = snowflake.connector.connect(
    connection_name=os.getenv("SNOWFLAKE_CONNECTION_NAME") or "admin"
)
cursor = conn.cursor()
cursor.execute("SELECT CURRENT_ACCOUNT(), CURRENT_ROLE()")
print(cursor.fetchone())
```

## Common Gotchas

| Issue | Cause | Fix |
|---|---|---|
| `get_active_session()` fails locally | Only works in Snowflake container | Use try/except fallback |
| `Invalid connection_name 'default'` | No `[default]` section in `connections.toml` | Use the actual section name (e.g., `admin`) |
| `importlib` caches old config | Edited `config.py` but kernel has old version | Restart kernel or use `importlib.reload(cfg)` |
| `st.connection` fails locally | Missing `connections.toml` or `secrets.toml` | Create one of the two config files |
| Session context lost between cells | Each notebook cell shares the same session | Session persists; no reconnection needed |
| `USE DATABASE` fails in SiS | App already has a database context | Skip `USE` commands when `get_active_session()` succeeds |
