# CoE Full Analysis Tool

Single Databricks export → 4 analysis outputs in one run.

## Agents

| Agent | Output file | Port (local) |
|---|---|---|
| Framework Analysis | `{account} - Framework Analysis - {ts}.xlsm` | — |
| Account Health Analysis | `{account} - Account Health Analysis - {ts}.xlsm` | — |
| Account Mastery Analysis | `{account} - Account Mastery Analysis - {ts}.xlsm` | — |
| Account Strategy Analysis | `{account} — Strategy Analysis {date_range}.xlsm` | — |

All 4 agents run in parallel via `ThreadPoolExecutor`. Each succeeds or fails independently — partial results are always surfaced.

---

## Repo structure

```
coe-full-analysis/
├── app.py                          # Flask app + agent runners
├── requirements.txt
├── render.yaml
├── Procfile
├── templates/
│   └── index.html                  # Frontend UI
│
├── templates/  (xlsm — git-ignored, must be added manually)
│   ├── CoE_Framework_Analysis_Templates.xlsm
│   ├── CoE_Account_Health_Analysis_Templates.xlsm
│   ├── CoE_Account_Mastery_Analysis_Templates.xlsm
│   └── CoE_Account_Strategy_Analysis_Templates_V2.xlsm
│
├── # Framework agent modules
├── reader_databricks.py
├── rules_engine.py
├── writer_framework.py
├── config.py
│
├── # Account Health agent modules
├── reader_databricks_health.py
├── rules_engine_health.py
├── writer_account_health.py
├── config_health.py
│
├── # Account Mastery agent modules
├── reader_databricks_mastery.py
├── rules_engine_mastery.py
├── writer_account_mastery.py
├── config_mastery.py
│
├── # Account Strategy agent modules
└── writer_strategy.py
```

---

## Setup

### 1. Copy all agent modules into the root

From each of your 4 existing repos, copy the following files into the root of this repo:

**Framework:**
```
reader_databricks.py
rules_engine.py
writer_framework.py
config.py
```

**Account Health:**
```
reader_databricks_health.py
rules_engine_health.py
writer_account_health.py
config_health.py
```

**Account Mastery:**
```
reader_databricks_mastery.py
rules_engine_mastery.py
writer_account_mastery.py
config_mastery.py
```

**Strategy:**
```
writer_strategy.py
```

### 2. Add the 4 .xlsm templates

Place all 4 template files inside a `templates/` subfolder:

```
templates/CoE_Framework_Analysis_Templates.xlsm
templates/CoE_Account_Health_Analysis_Templates.xlsm
templates/CoE_Account_Mastery_Analysis_Templates.xlsm
templates/CoE_Account_Strategy_Analysis_Templates_V2.xlsm
```

> Templates are git-ignored. Add them to Render via the dashboard or environment-mounted storage.

### 3. Run locally

```bash
pip install -r requirements.txt
python app.py
# Open http://127.0.0.1:8500
```

### 4. Deploy to Render

1. Push this repo to GitHub
2. Create a new **Web Service** on Render, connect the repo
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300 --workers 1`
5. Upload the 4 `.xlsm` templates to the Render disk or bundle them in the repo

> **Timeout note:** The combined run can take 60–120s depending on export size. The `--timeout 300` gives 5 minutes of headroom.

---

## /healthcheck

`GET /healthcheck` returns:
```json
{
  "status": "ok",
  "missing_templates": []
}
```
Returns `503` if any template file is missing.
