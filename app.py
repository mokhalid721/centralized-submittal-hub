import os
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, abort
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename

from db import db, User, Project, Template, TemplateField, Submittal, Transmittal, DocumentFile, Attachment, Setting
from docx_engine import extract_placeholders_from_docx, fill_docx_to_bytes
from utils import (
    ensure_dirs, get_storage_root, set_storage_root,
    date_format_options, name_format_options, format_date, format_name,
    guess_field_type, parse_dropdown_options,
    submittal_folder,
    make_zip_for_submittal, export_logs_csv
)

APP_PORT = int(os.environ.get("APP_PORT", "5001"))
ALLOWED_TEMPLATE_EXT = {".docx"}
ALLOWED_ATTACHMENT_EXT = {".pdf", ".docx", ".xlsx", ".xls", ".png", ".jpg", ".jpeg", ".txt", ".csv"}
BASE_DIR = Path(__file__).resolve().parent
DB_DIR = BASE_DIR / "storage" / "db"
DB_DIR.mkdir(parents=True, exist_ok=True)

def create_app():
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{(DB_DIR / 'hub.db').as_posix()}"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    db.init_app(app)

    login = LoginManager()
    login.login_view = "login"
    login.init_app(app)

    @login.user_loader
    def load_user(user_id):
        return db.session.get(User, int(user_id))

    with app.app_context():
        ensure_dirs()
        db.create_all()
        if not Setting.get("storage_root"):
            Setting.set("storage_root", str(Path.cwd() / "storage"))

    # -----------------------------
    # Setup / Auth
    # -----------------------------
    @app.get("/setup")
    def setup():
        if User.query.count() > 0:
            return redirect(url_for("dashboard"))
        return render_template("setup.html")

    @app.post("/setup")
    def setup_post():
        if User.query.count() > 0:
            return redirect(url_for("dashboard"))

        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        if not username or not password:
            flash("Username and password are required.", "danger")
            return redirect(url_for("setup"))

        u = User(username=username, role="admin")
        u.set_password(password)
        db.session.add(u)
        db.session.commit()
        login_user(u)
        flash("Admin user created.", "success")
        return redirect(url_for("dashboard"))

    @app.get("/login")
    def login():
        if User.query.count() == 0:
            return redirect(url_for("setup"))
        return render_template("login.html")

    @app.post("/login")
    def login_post():
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        u = User.query.filter_by(username=username).first()
        if not u or not u.check_password(password):
            flash("Invalid credentials.", "danger")
            return redirect(url_for("login"))
        login_user(u)
        return redirect(url_for("dashboard"))

    @app.get("/logout")
    @login_required
    def logout():
        logout_user()
        return redirect(url_for("login"))

    # -----------------------------
    # Root / Dashboard / Projects
    # -----------------------------
    @app.get("/")
    def root():
        if User.query.count() == 0:
            return redirect(url_for("setup"))
        if not current_user.is_authenticated:
            return redirect(url_for("login"))
        return redirect(url_for("dashboard"))

    @app.get("/dashboard")
    @login_required
    def dashboard():
        projects = Project.query.order_by(Project.created_at.desc()).all()
        return render_template("dashboard.html", projects=projects)

    @app.get("/projects/new")
    @login_required
    def project_new():
        return render_template("project_new.html")

    @app.post("/projects/new")
    @login_required
    def project_new_post():
        name = (request.form.get("name") or "").strip()
        contract_no = (request.form.get("contract_no") or "").strip()
        project_number = (request.form.get("project_number") or "").strip()
        if not name:
            flash("Project name is required.", "danger")
            return redirect(url_for("project_new"))

        p = Project(
            name=name,
            contract_no=contract_no,
            project_number=project_number,
            next_transmittal_seq=1,
            transmittal_prefix="T-",
            transmittal_padding=3,
            revision_style="dot",
        )
        db.session.add(p)
        db.session.commit()
        flash("Project created.", "success")
        return redirect(url_for("project_view", project_id=p.id))

    @app.get("/projects/<int:project_id>")
    @login_required
    def project_view(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)
        submittals = Submittal.query.filter_by(project_id=p.id).order_by(Submittal.created_at.desc()).all()
        return render_template("project_view.html", project=p, submittals=submittals)

    @app.post("/projects/<int:project_id>/settings")
    @login_required
    def project_settings_save(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)
        p.transmittal_prefix = (request.form.get("transmittal_prefix") or "T-").strip()
        p.transmittal_padding = int(request.form.get("transmittal_padding") or "3")
        p.revision_style = (request.form.get("revision_style") or "dot").strip()
        db.session.commit()
        flash("Project settings saved.", "success")
        return redirect(url_for("project_view", project_id=p.id))

    # -----------------------------
    # Templates
    # -----------------------------
    @app.get("/projects/<int:project_id>/templates")
    @login_required
    def template_list(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)
        templates = Template.query.filter_by(project_id=p.id).order_by(Template.created_at.desc()).all()
        return render_template("template_list.html", project=p, templates=templates)

    @app.post("/projects/<int:project_id>/templates/upload")
    @login_required
    def template_upload(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)
        f = request.files.get("template")
        t_type = (request.form.get("template_type") or "cover").strip()
        name = (request.form.get("name") or "").strip() or f"{t_type.title()} Template"

        if not f or not f.filename:
            flash("Choose a DOCX template to upload.", "danger")
            return redirect(url_for("template_list", project_id=p.id))

        ext = Path(f.filename).suffix.lower()
        if ext not in ALLOWED_TEMPLATE_EXT:
            flash("Templates must be .docx", "danger")
            return redirect(url_for("template_list", project_id=p.id))

        storage_root = Path(get_storage_root())
        tpl_dir = storage_root / "templates" / f"project_{p.id}"
        tpl_dir.mkdir(parents=True, exist_ok=True)

        filename = secure_filename(f.filename)
        save_path = tpl_dir / f"{int(datetime.utcnow().timestamp())}_{filename}"
        f.save(save_path)

        placeholders = extract_placeholders_from_docx(save_path)
        if not placeholders:
            try:
                save_path.unlink(missing_ok=True)
            except Exception:
                pass
            flash("No placeholders like «Field» found in that file.", "danger")
            return redirect(url_for("template_list", project_id=p.id))

        tpl = Template(
            project_id=p.id,
            name=name,
            template_type=t_type,
            file_path=str(save_path),
            created_by_user_id=current_user.id,
        )
        db.session.add(tpl)
        db.session.commit()

        # default field config
        for idx, key in enumerate(placeholders):
            ftype = guess_field_type(key)
            tf = TemplateField(
                template_id=tpl.id,
                key=key,
                label=key.replace("_", " "),
                field_type=ftype,
                required=False,
                options_text="",
                order_index=idx,
                formatter=("date" if "date" in key.lower() else ("name" if "name" in key.lower() else "")),
            )
            db.session.add(tf)

        db.session.commit()
        flash("Template uploaded. Configure fields if needed.", "success")
        return redirect(url_for("template_fields", template_id=tpl.id))

    @app.get("/templates/<int:template_id>/fields")
    @login_required
    def template_fields(template_id: int):
        tpl = db.session.get(Template, template_id) or abort(404)
        p = db.session.get(Project, tpl.project_id) or abort(404)
        fields = TemplateField.query.filter_by(template_id=tpl.id).order_by(TemplateField.order_index.asc()).all()
        return render_template("template_fields.html", project=p, template=tpl, fields=fields)

    @app.post("/templates/<int:template_id>/fields")
    @login_required
    def template_fields_save(template_id: int):
        tpl = db.session.get(Template, template_id) or abort(404)
        fields = TemplateField.query.filter_by(template_id=tpl.id).all()
        for f in fields:
            f.label = (request.form.get(f"label_{f.id}") or f.label).strip()
            f.field_type = (request.form.get(f"type_{f.id}") or f.field_type).strip()
            f.required = (request.form.get(f"required_{f.id}") == "on")
            f.options_text = (request.form.get(f"options_{f.id}") or "").strip()
            f.formatter = (request.form.get(f"formatter_{f.id}") or "").strip()
            f.order_index = int(request.form.get(f"order_{f.id}") or f.order_index)
        db.session.commit()
        flash("Template fields saved.", "success")
        return redirect(url_for("template_fields", template_id=tpl.id))

    @app.post("/templates/<int:template_id>/delete")
    @login_required
    def template_delete(template_id: int):
        tpl = db.session.get(Template, template_id) or abort(404)
        project_id = tpl.project_id
        try:
            Path(tpl.file_path).unlink(missing_ok=True)
        except Exception:
            pass
        TemplateField.query.filter_by(template_id=tpl.id).delete()
        db.session.delete(tpl)
        db.session.commit()
        flash("Template deleted.", "info")
        return redirect(url_for("template_list", project_id=project_id))

    # -----------------------------
    # Submittals: Step 1 -> Step 2 -> Create
    # -----------------------------
    @app.get("/projects/<int:project_id>/submittals/new")
    @login_required
    def submittal_new_step1(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)
        templates_cover = Template.query.filter_by(project_id=p.id, template_type="cover").all()
        templates_trans = Template.query.filter_by(project_id=p.id, template_type="transmittal").all()
        return render_template(
            "submittal_step1.html",
            project=p,
            templates_cover=templates_cover,
            templates_trans=templates_trans,
            statuses=Project.default_statuses(),
            dispositions=Project.default_dispositions(),
            date_formats=date_format_options(),
            name_formats=name_format_options(),
        )

    def union_fields(template_ids):
        fields = []
        seen = set()
        for tid in template_ids:
            if not tid:
                continue
            for f in TemplateField.query.filter_by(template_id=tid).order_by(TemplateField.order_index.asc()).all():
                if f.key not in seen:
                    f._options = parse_dropdown_options(f.options_text)
                    fields.append(f)
                    seen.add(f.key)
        return fields

    @app.post("/projects/<int:project_id>/submittals/new/fields")
    @login_required
    def submittal_new_step2(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)

        cover_id = int(request.form.get("cover_template_id") or "0") or None
        trans_id = int(request.form.get("trans_template_id") or "0") or None
        if not cover_id and not trans_id:
            flash("Select at least one template.", "danger")
            return redirect(url_for("submittal_new_step1", project_id=p.id))

        sub_no = (request.form.get("sub_no") or "").strip()
        if not sub_no:
            flash("Submittal No is required.", "danger")
            return redirect(url_for("submittal_new_step1", project_id=p.id))

        carry = dict(request.form)
        carry["cover_template_id"] = str(cover_id or 0)
        carry["trans_template_id"] = str(trans_id or 0)

        fields = union_fields([cover_id, trans_id])

        defaults = {
            "Project_Name": p.name or "",
            "Contract_No": p.contract_no or "",
            "Project_Number": p.project_number or "",
            "Sub_No": sub_no,
            "Sub_Title": (request.form.get("title") or "").strip(),
            "Spec_Section": (request.form.get("spec_section") or "").strip(),
            "Authorization": (request.form.get("disposition") or "").strip(),
        }

        return render_template("submittal_step2.html", project=p, fields=fields, carry=carry, defaults=defaults)

    @app.post("/projects/<int:project_id>/submittals/create")
    @login_required
    def submittal_create(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)

        cover_id = int(request.form.get("cover_template_id") or "0") or None
        trans_id = int(request.form.get("trans_template_id") or "0") or None

        sub_no = (request.form.get("sub_no") or "").strip()
        title = (request.form.get("title") or "").strip()
        spec_section = (request.form.get("spec_section") or "").strip()
        status = (request.form.get("status") or "Draft").strip()
        disposition = (request.form.get("disposition") or "").strip()
        responsible = (request.form.get("responsible") or "").strip()
        notes = (request.form.get("notes") or "").strip()

        date_format_key = (request.form.get("date_format") or "mdy_slash").strip()
        name_format_key = (request.form.get("name_format") or "first_last").strip()

        create_transmittal = (request.form.get("create_transmittal") == "on")
        sent_to = (request.form.get("sent_to") or "").strip()
        delivery_method = (request.form.get("delivery_method") or "").strip()

        if not sub_no:
            flash("Submittal No is required.", "danger")
            return redirect(url_for("submittal_new_step1", project_id=p.id))

        s = Submittal(
            project_id=p.id,
            sub_no=sub_no,
            title=title,
            spec_section=spec_section,
            rev="",
            status=status,
            disposition=disposition,
            responsible_person=responsible,
            notes=notes,
            created_by_user_id=current_user.id,
        )
        db.session.add(s)
        db.session.commit()

        t = None
        if create_transmittal:
            trans_no = p.make_next_transmittal_no()
            t = Transmittal(
                project_id=p.id,
                trans_no=trans_no,
                date_sent=datetime.utcnow(),
                sent_to=sent_to,
                delivery_method=delivery_method,
                created_by_user_id=current_user.id,
            )
            db.session.add(t)
            db.session.commit()

        fields = union_fields([cover_id, trans_id])
        values = {}
        for f in fields:
            values[f.key] = (request.form.get(f"field_{f.id}") or "").strip()

        auto_map = {
            "Project_Name": p.name,
            "Contract_No": p.contract_no,
            "Project_Number": p.project_number,
            "Sub_No": s.sub_no,
            "Sub_Title": s.title,
            "Spec_Section": s.spec_section,
            "Authorization": s.disposition,
        }
        for k, v in auto_map.items():
            if k in values and not values[k]:
                values[k] = v or ""

        for f in fields:
            if f.formatter == "date" or "date" in f.key.lower():
                if not values.get(f.key, ""):
                    values[f.key] = format_date(datetime.utcnow(), date_format_key)
            if f.formatter == "name" or "name" in f.key.lower():
                values[f.key] = format_name(values.get(f.key, ""), name_format_key)

        sub_dir = submittal_folder(p, s)
        gen_dir = sub_dir / "Generated"
        att_dir = sub_dir / "Attachments"
        gen_dir.mkdir(parents=True, exist_ok=True)
        att_dir.mkdir(parents=True, exist_ok=True)

        def generate_doc(tpl_id, label):
            if not tpl_id:
                return None
            tpl = db.session.get(Template, tpl_id)
            if not tpl:
                return None
            doc_bytes = fill_docx_to_bytes(tpl.file_path, values)
            out_path = gen_dir / f"{secure_filename(s.sub_no)}_{label}.docx"
            out_path.write_bytes(doc_bytes.getvalue())

            df = DocumentFile(
                project_id=p.id,
                submittal_id=s.id,
                transmittal_id=(t.id if t else None),
                doc_type=label,
                file_path=str(out_path),
                created_by_user_id=current_user.id,
            )
            db.session.add(df)
            db.session.commit()
            return str(out_path)

        generate_doc(cover_id, "CoverLetter")
        generate_doc(trans_id, "Transmittal")

        # attachments
        files = request.files.getlist("attachments")
        for af in files:
            if not af or not af.filename:
                continue
            ext = Path(af.filename).suffix.lower()
            if ext and ext not in ALLOWED_ATTACHMENT_EXT:
                continue
            fname = secure_filename(af.filename)
            apath = att_dir / f"{int(datetime.utcnow().timestamp())}_{fname}"
            af.save(apath)
            a = Attachment(
                project_id=p.id,
                submittal_id=s.id,
                original_filename=af.filename,
                stored_path=str(apath),
                uploaded_by_user_id=current_user.id,
            )
            db.session.add(a)
        db.session.commit()

        export_logs_csv(p.id)
        flash("Submittal created and documents generated.", "success")
        return redirect(url_for("submittal_view", submittal_id=s.id))

    @app.get("/submittals/<int:submittal_id>")
    @login_required
    def submittal_view(submittal_id: int):
        s = db.session.get(Submittal, submittal_id) or abort(404)
        p = db.session.get(Project, s.project_id) or abort(404)
        docs = DocumentFile.query.filter_by(submittal_id=s.id).order_by(DocumentFile.created_at.desc()).all()
        atts = Attachment.query.filter_by(submittal_id=s.id).order_by(Attachment.uploaded_at.desc()).all()
        return render_template("submittal_view.html", project=p, submittal=s, docs=docs, attachments=atts)

    @app.get("/files")
    @login_required
    def download_file():
        path = request.args.get("path")
        if not path:
            abort(400)
        p = Path(path)
        if not p.exists():
            abort(404)
        return send_file(p, as_attachment=True, download_name=p.name)

    @app.get("/submittals/<int:submittal_id>/zip")
    @login_required
    def submittal_zip(submittal_id: int):
        zip_path = make_zip_for_submittal(submittal_id)
        return send_file(zip_path, as_attachment=True, download_name=Path(zip_path).name)

    # logs export
    @app.get("/projects/<int:project_id>/export/submittal_log.csv")
    @login_required
    def export_submittal_log(project_id: int):
        csv_path = export_logs_csv(project_id)["submittal"]
        return send_file(csv_path, as_attachment=True, download_name="submittal_log.csv")

    # batch
    @app.get("/projects/<int:project_id>/batch")
    @login_required
    def batch_page(project_id: int):
        p = db.session.get(Project, project_id) or abort(404)
        templates_cover = Template.query.filter_by(project_id=p.id, template_type="cover").all()
        templates_trans = Template.query.filter_by(project_id=p.id, template_type="transmittal").all()
        return render_template(
            "batch.html",
            project=p,
            templates_cover=templates_cover,
            templates_trans=templates_trans,
            date_formats=date_format_options(),
            name_formats=name_format_options(),
        )

    @app.post("/projects/<int:project_id>/batch")
    @login_required
    def batch_run(project_id: int):
        import csv

        p = db.session.get(Project, project_id) or abort(404)

        cover_id = int(request.form.get("cover_template_id") or "0") or None
        trans_id = int(request.form.get("trans_template_id") or "0") or None
        if not cover_id and not trans_id:
            flash("Select at least one template.", "danger")
            return redirect(url_for("batch_page", project_id=p.id))

        f = request.files.get("csv_file")
        if not f or not f.filename:
            flash("Upload a CSV.", "danger")
            return redirect(url_for("batch_page", project_id=p.id))

        date_format_key = (request.form.get("date_format") or "mdy_slash").strip()
        name_format_key = (request.form.get("name_format") or "first_last").strip()

        content = f.stream.read().decode("utf-8-sig", errors="replace").splitlines()
        reader = csv.DictReader(content)
        rows = list(reader)

        if not rows:
            flash("CSV had no rows.", "danger")
            return redirect(url_for("batch_page", project_id=p.id))

        fields = union_fields([cover_id, trans_id])
        field_keys = [ff.key for ff in fields]

        created = 0
        for r in rows:
            sub_no = (r.get("Sub_No") or r.get("sub_no") or "").strip()
            if not sub_no:
                continue

            s = Submittal(
                project_id=p.id,
                sub_no=sub_no,
                title=(r.get("Sub_Title") or r.get("Title") or "").strip(),
                spec_section=(r.get("Spec_Section") or "").strip(),
                status=(r.get("Status") or "Draft").strip(),
                disposition=(r.get("Disposition") or "").strip(),
                responsible_person=(r.get("Responsible") or "").strip(),
                notes=(r.get("Notes") or "").strip(),
                created_by_user_id=current_user.id,
            )
            db.session.add(s)
            db.session.commit()

            values = {k: (r.get(k) or "").strip() for k in field_keys}

            auto_map = {
                "Project_Name": p.name,
                "Contract_No": p.contract_no,
                "Project_Number": p.project_number,
                "Sub_No": s.sub_no,
                "Sub_Title": s.title,
                "Spec_Section": s.spec_section,
                "Authorization": s.disposition,
            }
            for k, v in auto_map.items():
                if k in values and not values[k]:
                    values[k] = v or ""

            for ff in fields:
                if ff.formatter == "date" or "date" in ff.key.lower():
                    if not values.get(ff.key, ""):
                        values[ff.key] = format_date(datetime.utcnow(), date_format_key)
                if ff.formatter == "name" or "name" in ff.key.lower():
                    values[ff.key] = format_name(values.get(ff.key, ""), name_format_key)

            sub_dir = submittal_folder(p, s)
            gen_dir = sub_dir / "Generated"
            gen_dir.mkdir(parents=True, exist_ok=True)

            def gen(tpl_id, label):
                if not tpl_id:
                    return
                tpl = db.session.get(Template, tpl_id)
                doc_bytes = fill_docx_to_bytes(tpl.file_path, values)
                out_path = gen_dir / f"{secure_filename(s.sub_no)}_{label}.docx"
                out_path.write_bytes(doc_bytes.getvalue())
                df = DocumentFile(
                    project_id=p.id,
                    submittal_id=s.id,
                    doc_type=label,
                    file_path=str(out_path),
                    created_by_user_id=current_user.id,
                )
                db.session.add(df)
                db.session.commit()

            gen(cover_id, "CoverLetter")
            gen(trans_id, "Transmittal")
            created += 1

        export_logs_csv(p.id)
        flash(f"Batch complete: created {created} submittals.", "success")
        return redirect(url_for("project_view", project_id=p.id))

    # -----------------------------
    # Admin: Users
    # -----------------------------
    @app.get("/admin/users")
    @login_required
    def admin_users():
        if not current_user.is_admin:
            abort(403)
        users = User.query.order_by(User.created_at.desc()).all()
        return render_template("admin_users.html", users=users)

    @app.post("/admin/users")
    @login_required
    def admin_users_create():
        if not current_user.is_admin:
            abort(403)

        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        role = (request.form.get("role") or "editor").strip()

        if not username or not password:
            flash("Username and password required.", "danger")
            return redirect(url_for("admin_users"))

        if User.query.filter_by(username=username).first():
            flash("That username already exists.", "danger")
            return redirect(url_for("admin_users"))

        u = User(username=username, role=role)
        u.set_password(password)
        db.session.add(u)
        db.session.commit()
        flash("User created.", "success")
        return redirect(url_for("admin_users"))

    # -----------------------------
    # Settings
    # -----------------------------
    @app.get("/settings")
    @login_required
    def settings_page():
        storage_root = get_storage_root()
        return render_template("settings.html", storage_root=storage_root)

    @app.post("/settings")
    @login_required
    def settings_save():
        storage_root = (request.form.get("storage_root") or "").strip()
        if not storage_root:
            flash("Storage root cannot be empty.", "danger")
            return redirect(url_for("settings_page"))
        set_storage_root(storage_root)
        flash("Settings saved.", "success")
        return redirect(url_for("settings_page"))

    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, port=APP_PORT)
