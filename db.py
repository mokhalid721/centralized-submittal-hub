from datetime import datetime
from typing import Optional

from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class Setting(db.Model):
    __tablename__ = "settings"
    key = db.Column(db.String(120), primary_key=True)
    value = db.Column(db.Text, nullable=False, default="")

    @staticmethod
    def get(key: str) -> Optional[str]:
        obj = Setting.query.filter_by(key=key).first()
        return obj.value if obj else None

    @staticmethod
    def set(key: str, value: str) -> None:
        obj = Setting.query.filter_by(key=key).first()
        if not obj:
            obj = Setting(key=key, value=value)
            db.session.add(obj)
        else:
            obj.value = value
        db.session.commit()


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(30), nullable=False, default="editor")  # admin/editor/viewer
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"

    def set_password(self, password: str):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


class Project(db.Model):
    __tablename__ = "projects"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(240), nullable=False)
    contract_no = db.Column(db.String(120), nullable=True)
    project_number = db.Column(db.String(120), nullable=True)

    # numbering settings
    next_transmittal_seq = db.Column(db.Integer, default=1)
    transmittal_prefix = db.Column(db.String(30), default="T-")
    transmittal_padding = db.Column(db.Integer, default=3)
    revision_style = db.Column(db.String(10), default="dot")  # dot or R

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @staticmethod
    def default_statuses():
        return ["Draft", "Sent", "Returned", "Closed", "Resubmit Needed"]

    @staticmethod
    def default_dispositions():
        return [
            "Authorized / No Exceptions Taken",
            "Authorized / Make Corrections Noted",
            "Revise & Resubmit",
            "Rejected",
            "For Information Only",
        ]

    def make_next_transmittal_no(self) -> str:
        seq = int(self.next_transmittal_seq or 1)
        trans_no = f"{self.transmittal_prefix}{str(seq).zfill(int(self.transmittal_padding or 3))}"
        self.next_transmittal_seq = seq + 1
        db.session.commit()
        return trans_no


class Template(db.Model):
    __tablename__ = "templates"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)
    name = db.Column(db.String(240), nullable=False)
    template_type = db.Column(db.String(30), nullable=False)  # cover/transmittal/response
    file_path = db.Column(db.Text, nullable=False)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class TemplateField(db.Model):
    __tablename__ = "template_fields"
    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey("templates.id"), nullable=False)

    # placeholder key WITHOUT « »
    key = db.Column(db.String(240), nullable=False)
    label = db.Column(db.String(240), nullable=False)
    field_type = db.Column(db.String(30), nullable=False, default="text")  # text, textarea, date, select
    required = db.Column(db.Boolean, default=False)
    options_text = db.Column(db.Text, default="")  # for selects: one option per line
    formatter = db.Column(db.String(30), default="")  # "", "date", "name"

    order_index = db.Column(db.Integer, default=0)


class Submittal(db.Model):
    __tablename__ = "submittals"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)

    sub_no = db.Column(db.String(120), nullable=False)
    title = db.Column(db.String(240), nullable=True)
    spec_section = db.Column(db.String(120), nullable=True)
    rev = db.Column(db.String(30), nullable=True, default="")
    status = db.Column(db.String(60), nullable=False, default="Draft")
    disposition = db.Column(db.String(120), nullable=True, default="")
    responsible_person = db.Column(db.String(120), nullable=True, default="")
    notes = db.Column(db.Text, nullable=True, default="")

    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Transmittal(db.Model):
    __tablename__ = "transmittals"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)

    trans_no = db.Column(db.String(120), nullable=False)
    date_sent = db.Column(db.DateTime, nullable=True)
    sent_to = db.Column(db.String(240), nullable=True, default="")
    delivery_method = db.Column(db.String(120), nullable=True, default="")
    notes = db.Column(db.Text, nullable=True, default="")

    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class DocumentFile(db.Model):
    __tablename__ = "documents"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)
    submittal_id = db.Column(db.Integer, db.ForeignKey("submittals.id"), nullable=True)
    transmittal_id = db.Column(db.Integer, db.ForeignKey("transmittals.id"), nullable=True)

    doc_type = db.Column(db.String(60), nullable=False)  # CoverLetter / Transmittal / etc
    file_path = db.Column(db.Text, nullable=False)

    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Attachment(db.Model):
    __tablename__ = "attachments"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False)
    submittal_id = db.Column(db.Integer, db.ForeignKey("submittals.id"), nullable=False)

    original_filename = db.Column(db.String(255), nullable=False)
    stored_path = db.Column(db.Text, nullable=False)

    uploaded_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
