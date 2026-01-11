import csv
from datetime import datetime
from pathlib import Path
import zipfile
from typing import Dict, List

from werkzeug.utils import secure_filename

from db import db, Setting, Project, Submittal


def ensure_dirs():
    Path("storage").mkdir(parents=True, exist_ok=True)
    Path("instance").mkdir(parents=True, exist_ok=True)
    Path("sample_templates").mkdir(parents=True, exist_ok=True)
    Path("scripts").mkdir(parents=True, exist_ok=True)


def get_storage_root() -> str:
    return Setting.get("storage_root") or str(Path.cwd() / "storage")


def set_storage_root(path: str):
    Setting.set("storage_root", path)
    Path(path).mkdir(parents=True, exist_ok=True)


def date_format_options():
    return [
        ("mdy_slash", "MM/DD/YYYY"),
        ("month_d_yyyy", "Month D, YYYY"),
        ("iso", "YYYY-MM-DD"),
    ]


def name_format_options():
    return [
        ("first_last", "First Last"),
        ("last_first", "Last, First"),
        ("mrms_last", "Mr./Ms. Last"),
    ]


def format_date(dt: datetime, fmt_key: str) -> str:
    if fmt_key == "month_d_yyyy":
        return dt.strftime("%B %d, %Y").replace(" 0", " ")
    if fmt_key == "iso":
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%m/%d/%Y")


def format_name(raw: str, fmt_key: str) -> str:
    raw = (raw or "").strip()
    if not raw:
        return ""
    parts = raw.split()
    if fmt_key == "last_first":
        if len(parts) >= 2:
            return f"{parts[-1]}, {' '.join(parts[:-1])}"
        return raw
    if fmt_key == "mrms_last":
        first = parts[0].lower().rstrip(".")
        if first in ("mr", "ms", "mrs", "dr"):
            return f"{parts[0]} {parts[-1]}"
        return f"Mr./Ms. {parts[-1]}"
    return raw


def guess_field_type(key: str) -> str:
    k = key.lower()
    if "bulleted" in k or "notes" in k or "address" in k:
        return "textarea"
    if "date" in k:
        return "date"
    return "text"


def parse_dropdown_options(options_text: str) -> List[str]:
    if not options_text:
        return []
    return [ln.strip() for ln in options_text.splitlines() if ln.strip()]


def project_folder(project: Project) -> Path:
    storage_root = Path(get_storage_root())
    return storage_root / "projects" / f"project_{project.id}_{secure_filename(project.name)}"


def submittal_folder(project: Project, submittal: Submittal) -> Path:
    return project_folder(project) / "Submittals" / secure_filename(submittal.sub_no)


def export_logs_csv(project_id: int) -> Dict[str, str]:
    p = db.session.get(Project, project_id)
    base = project_folder(p)
    logs_dir = base / "Logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    submittals = Submittal.query.filter_by(project_id=p.id).order_by(Submittal.created_at.asc()).all()

    sub_path = logs_dir / "submittal_log.csv"
    with sub_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "Submittal No", "Title", "Spec Section", "Rev", "Status",
            "Sent Date", "Returned Date", "Disposition", "Responsible Person",
            "Notes"
        ])
        for s in submittals:
            w.writerow([
                s.sub_no, s.title or "", s.spec_section or "", s.rev or "", s.status or "",
                "", "", s.disposition or "", s.responsible_person or "",
                (s.notes or "").replace("\n", " ").strip()
            ])

    trans_path = logs_dir / "transmittal_log.csv"
    if not trans_path.exists():
        with trans_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Transmittal No", "Date Sent", "Sent To", "Delivery Method", "Related Submittals", "Notes"])

    return {"submittal": str(sub_path), "transmittal": str(trans_path)}


def make_zip_for_submittal(submittal_id: int) -> str:
    from db import Project, Submittal
    s = db.session.get(Submittal, submittal_id)
    p = db.session.get(Project, s.project_id)

    sub_dir = submittal_folder(p, s)
    zip_dir = Path(get_storage_root()) / "zips"
    zip_dir.mkdir(parents=True, exist_ok=True)
    zip_path = zip_dir / f"{secure_filename(p.name)}_{secure_filename(s.sub_no)}.zip"

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        if sub_dir.exists():
            for file in sub_dir.rglob("*"):
                if file.is_file():
                    zf.write(file, arcname=str(file.relative_to(sub_dir)))

    return str(zip_path)
