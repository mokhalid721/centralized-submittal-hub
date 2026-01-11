# Centralized Submittal Hub (Local MVP)

Local-first, multi-user web app you run from VS Code.

## What it does
- Create projects
- Upload DOCX templates with placeholders like `«Field»`
- Auto-detect placeholders and generate a clean form
- Configure field labels/types/order (text, multiline, date, dropdown)
- Generate formatted DOCX outputs (preserves bullets/spacing)
- Remove unused bullet lines and “placeholder-only” lines automatically
- Upload attachments and keep everything linked to the submittal
- Auto-maintain Submittal + Transmittal logs (export CSV)
- Batch generation from CSV

## Quick start
### 1) Create and activate a virtual environment
macOS / Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows (PowerShell)
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2) Install dependencies
```bash
python -m pip install -r requirements.txt
```

### 3) Run the app
```bash
python app.py
```

Open: http://127.0.0.1:5001  
(Port 5001 by default; change with `APP_PORT=5050 python app.py`)

### 4) First-time setup
You’ll be redirected to `/setup` to create the first Admin user.

### 5) Generate sample templates (optional)
```bash
python scripts/make_sample_templates.py
```
Uploads go here: Project → Templates

## Placeholders
Use `«Key»` in your DOCX.

The app:
- Replaces placeholders even if Word split them across runs
- Removes empty bullet/list lines
- Removes paragraphs that were only placeholders if the value is blank

## Storage
Saved under `storage/` (gitignored): templates, generated docs, attachments, logs, zips
