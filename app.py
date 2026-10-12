from pathlib import Path
from datetime import datetime, timezone
from functools import wraps
import os

from flask import Flask, jsonify, request, session, send_from_directory, abort, redirect
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import re

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None

BASE = Path(__file__).resolve().parent
UPLOAD_DIR = BASE / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("JOBBRIDGE_SECRET", "change-this-secret-in-production")
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL", f"sqlite:///{BASE / 'jobbridge.db'}"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = str(UPLOAD_DIR)

# To use MySQL instead of SQLite, set for example:
# DATABASE_URL=mysql+pymysql://root:password@localhost/jobbridge

db = SQLAlchemy(app)


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(160), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="seeker")
    phone = db.Column(db.String(30))
    location = db.Column(db.String(120))
    bio = db.Column(db.Text)
    resume_filename = db.Column(db.String(255))
    resume_text = db.Column(db.Text)
    company_name = db.Column(db.String(160))
    company_description = db.Column(db.Text)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def public(self):
        return {
            "id": self.id, "name": self.name, "email": self.email, "role": self.role,
            "phone": self.phone or "", "location": self.location or "", "bio": self.bio or "",
            "resume_filename": self.resume_filename or "", "resume_uploaded": bool(self.resume_filename),
            "company_name": self.company_name or "",
            "company_description": self.company_description or "", "is_active": self.is_active,
        }


class Job(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(180), nullable=False)
    company = db.Column(db.String(180), nullable=False)
    category = db.Column(db.String(80), nullable=False)
    location = db.Column(db.String(120), nullable=False)
    desc = db.Column(db.Text, default="")
    posted_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    status = db.Column(db.String(30), default="Open", nullable=False)
    is_approved = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    owner = db.relationship("User", backref="jobs")

    def public(self):
        return {
            "id": f"j{self.id}", "title": self.title, "company": self.company,
            "category": self.category, "location": self.location, "desc": self.desc or "",
            "postedBy": "recruiter", "status": self.status,
            "isApproved": self.is_approved,
        }


class Application(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    job_id = db.Column(db.Integer, db.ForeignKey("job.id"), nullable=False)
    applicant_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    status = db.Column(db.String(30), default="Pending", nullable=False)
    applied_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    job = db.relationship("Job", backref="applications")
    applicant = db.relationship("User", backref="applications")

    def public(self):
        return {
            "id": f"a{self.id}", "jobId": f"j{self.job_id}",
            "applicantName": self.applicant.name, "applicantId": self.applicant_id,
            "status": self.status, "appliedAt": self.applied_at.isoformat() if self.applied_at else "",
        }


class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    message = db.Column(db.String(500), nullable=False)
    is_read = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    user = db.relationship("User", backref="notifications")

    def public(self):
        return {"id": self.id, "message": self.message, "isRead": self.is_read,
                "createdAt": self.created_at.isoformat() if self.created_at else ""}



# ---------- Resume extraction / matching ----------
MATCH_SKILLS = {
    "python", "java", "javascript", "typescript", "html", "css", "react", "angular", "vue",
    "node", "nodejs", "flask", "django", "fastapi", "sql", "mysql", "postgresql", "mongodb",
    "firebase", "aws", "azure", "gcp", "google cloud", "docker", "kubernetes", "git", "github",
    "machine learning", "deep learning", "tensorflow", "pytorch", "pandas", "numpy", "scikit-learn",
    "data analysis", "data analytics", "power bi", "tableau", "excel", "figma", "ui", "ux",
    "product management", "marketing", "sales", "customer support", "operations", "communication"
}


def extract_resume_text(path, ext):
    if ext == "pdf":
        if PdfReader is None:
            raise RuntimeError("PDF support is not installed. Run pip install -r requirements.txt")
        reader = PdfReader(str(path))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    if ext == "docx":
        if Document is None:
            raise RuntimeError("DOCX support is not installed. Run pip install -r requirements.txt")
        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs)
    if ext == "txt":
        return Path(path).read_text(encoding="utf-8", errors="ignore")
    if ext == "doc":
        raise ValueError("Legacy .doc files are accepted for storage, but automatic text matching requires PDF or DOCX. Please upload PDF/DOCX for resume matching.")
    raise ValueError("Unsupported resume format")


def normalize_words(text):
    return set(re.findall(r"[a-z][a-z0-9+#.-]{1,}", (text or "").lower()))


def resume_skills(text):
    lower = (text or "").lower()
    return [skill for skill in sorted(MATCH_SKILLS, key=len, reverse=True)
            if re.search(r"(?<![a-z0-9])" + re.escape(skill) + r"(?![a-z0-9])", lower)]


def score_resume_for_job(resume_text, job):
    """Explainable keyword relevance; it is not an automated hiring decision."""
    job_text = " ".join([job.title or "", job.category or "", job.desc or "", job.company or ""]).lower()
    r_words = normalize_words(resume_text)
    j_words = normalize_words(job_text)
    overlap = r_words & j_words
    title_words = normalize_words((job.title or "") + " " + (job.category or ""))
    title_overlap = r_words & title_words
    skills = resume_skills(resume_text)
    matched_skills = [s for s in skills if s in job_text]
    lexical = (len(overlap) / max(1, len(j_words))) if j_words else 0
    score = min(100, round((lexical * 100 * 0.35) + min(0.35, len(title_overlap) * 0.12) * 100 + min(0.45, len(matched_skills) * 0.09) * 100))
    return score, matched_skills[:12]

def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    user = db.session.get(User, uid)
    if not user or not user.is_active:
        session.clear()
        return None
    return user


def login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        user = current_user()
        if not user:
            return jsonify({"error": "Authentication required"}), 401
        return fn(user, *args, **kwargs)
    return wrapped


def role_required(*roles):
    def decorator(fn):
        @wraps(fn)
        def wrapped(user, *args, **kwargs):
            if user.role not in roles:
                return jsonify({"error": "Access denied"}), 403
            return fn(user, *args, **kwargs)
        return wrapped
    return decorator


def notify(user_id, message):
    db.session.add(Notification(user_id=user_id, message=message))


@app.route("/")
def index():
    # The supplied initial HTML is kept untouched. The backend adds only an extension script at runtime.
    html = (BASE / "templates" / "Pasted code.html").read_text(encoding="utf-8")
    injection = '\n<script src="/static/app_extension.js"></script>\n'
    return html.replace("</body>", injection + "</body>")


@app.route("/uploads/<path:filename>")
@login_required
def uploaded_file(user, filename):
    # Only the owner or an admin can access resume files.
    if user.role == "admin":
        return send_from_directory(app.config["UPLOAD_FOLDER"], filename)
    if user.resume_filename == filename:
        return send_from_directory(app.config["UPLOAD_FOLDER"], filename)
    # Recruiters can view a resume only when the candidate applied to one of their jobs.
    if user.role == "recruiter":
        allowed = db.session.query(Application.id).join(Job, Application.job_id == Job.id).join(User, Application.applicant_id == User.id).filter(
            Job.posted_by == user.id, User.resume_filename == filename
        ).first()
        if allowed:
            return send_from_directory(app.config["UPLOAD_FOLDER"], filename)
    abort(403)


@app.post("/api/auth/register")
def register():
    data = request.get_json(silent=True) or {}
    name, email, password = data.get("name", "").strip(), data.get("email", "").strip().lower(), data.get("password", "")
    role = data.get("role", "seeker")
    if not name or not email or not password or role not in {"seeker", "recruiter"}:
        return jsonify({"error": "Name, email, password and a valid role are required"}), 400
    if User.query.filter_by(email=email).first():
        return jsonify({"error": "Email already registered"}), 409
    user = User(name=name, email=email, password_hash=generate_password_hash(password), role=role)
    if role == "recruiter":
        user.company_name = data.get("company_name", "").strip()
    db.session.add(user)
    db.session.commit()
    session["user_id"] = user.id
    return jsonify({"user": user.public()})


@app.post("/api/auth/login")
def login():
    data = request.get_json(silent=True) or {}
    email, password = data.get("email", "").strip().lower(), data.get("password", "")
    user = User.query.filter_by(email=email).first()
    if not user or not user.is_active or not check_password_hash(user.password_hash, password):
        return jsonify({"error": "Invalid email or password"}), 401
    session["user_id"] = user.id
    return jsonify({"user": user.public()})


@app.post("/api/auth/logout")
def logout():
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/auth/me")
def me():
    user = current_user()
    return jsonify({"user": user.public() if user else None})


@app.get("/api/jobs")
def get_jobs():
    user = current_user()
    query = Job.query.filter_by(status="Open", is_approved=True)
    category = request.args.get("category", "All")
    location = request.args.get("location", "All")
    q = request.args.get("q", "").strip().lower()
    sort = request.args.get("sort", "newest")
    if category != "All": query = query.filter_by(category=category)
    if location != "All": query = query.filter_by(location=location)
    jobs = query.all()
    if q:
        jobs = [j for j in jobs if q in j.title.lower() or q in j.company.lower() or q in (j.desc or "").lower()]
    if sort == "title": jobs.sort(key=lambda j: j.title.lower())
    elif sort == "company": jobs.sort(key=lambda j: j.company.lower())
    else: jobs.sort(key=lambda j: j.created_at or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return jsonify({"jobs": [j.public() for j in jobs], "locations": sorted({j.location for j in Job.query.filter_by(is_approved=True).all()})})


@app.post("/api/jobs")
@login_required
@role_required("recruiter")
def create_job(user):
    data = request.get_json(silent=True) or {}
    title, location = data.get("title", "").strip(), data.get("location", "").strip()
    if not title or not location:
        return jsonify({"error": "Title and location are required"}), 400
    job = Job(title=title, category=data.get("category", "Engineering"), location=location,
              desc=data.get("desc", "").strip() or "No description added yet.",
              company=user.company_name or "Your company", posted_by=user.id, is_approved=True)
    db.session.add(job)
    db.session.commit()
    return jsonify({"job": job.public()})


@app.put("/api/jobs/<int:job_id>")
@login_required
@role_required("recruiter", "admin")
def edit_job(user, job_id):
    job = db.session.get(Job, job_id)
    if not job or (user.role != "admin" and job.posted_by != user.id):
        return jsonify({"error": "Job not found or access denied"}), 404
    data = request.get_json(silent=True) or {}
    for field in ("title", "category", "location", "desc", "company", "status"):
        if field in data and data[field] is not None:
            setattr(job, field, str(data[field]).strip())
    if user.role == "admin" and "is_approved" in data:
        job.is_approved = bool(data["is_approved"])
    db.session.commit()
    return jsonify({"job": job.public()})


@app.delete("/api/jobs/<int:job_id>")
@login_required
@role_required("recruiter", "admin")
def delete_job(user, job_id):
    job = db.session.get(Job, job_id)
    if not job or (user.role != "admin" and job.posted_by != user.id):
        return jsonify({"error": "Job not found or access denied"}), 404
    db.session.delete(job)
    db.session.commit()
    return jsonify({"ok": True})


@app.post("/api/jobs/<int:job_id>/apply")
@login_required
@role_required("seeker")
def apply(user, job_id):
    job = db.session.get(Job, job_id)
    if not job or job.status != "Open" or not job.is_approved:
        return jsonify({"error": "Job is not available"}), 404
    if Application.query.filter_by(job_id=job.id, applicant_id=user.id).first():
        return jsonify({"error": "You already applied to this job"}), 409
    appn = Application(job_id=job.id, applicant_id=user.id)
    db.session.add(appn)
    db.session.flush()
    notify(job.posted_by, f"{user.name} applied for {job.title}.")
    db.session.commit()
    return jsonify({"application": appn.public()})


@app.get("/api/recruiter/jobs")
@login_required
@role_required("recruiter")
def recruiter_jobs(user):
    return jsonify({"jobs": [j.public() for j in Job.query.filter_by(posted_by=user.id).order_by(Job.created_at.desc()).all()]})


@app.get("/api/applications")
@login_required
def applications(user):
    if user.role == "seeker":
        return jsonify({"applications": [a.public() for a in Application.query.filter_by(applicant_id=user.id).order_by(Application.applied_at.desc()).all()]})
    jobs = Job.query.filter_by(posted_by=user.id).all()
    job_ids = [j.id for j in jobs]
    apps = Application.query.filter(Application.job_id.in_(job_ids)).order_by(Application.applied_at.desc()).all() if job_ids else []
    payload = []
    for a in apps:
        item = a.public()
        item["resumeUrl"] = f"/uploads/{a.applicant.resume_filename}" if a.applicant.resume_filename else ""
        item["applicantBio"] = a.applicant.bio or ""
        payload.append(item)
    return jsonify({"applications": payload})


@app.patch("/api/applications/<int:app_id>")
@login_required
@role_required("recruiter", "admin")
def update_application(user, app_id):
    appn = db.session.get(Application, app_id)
    if not appn or (user.role != "admin" and appn.job.posted_by != user.id):
        return jsonify({"error": "Application not found or access denied"}), 404
    status = (request.get_json(silent=True) or {}).get("status")
    if status not in {"Pending", "Shortlisted", "Hired", "Rejected"}:
        return jsonify({"error": "Invalid status"}), 400
    appn.status = status
    notify(appn.applicant_id, f"Your application for {appn.job.title} is now {status}.")
    db.session.commit()
    return jsonify({"application": appn.public()})


@app.get("/api/profile")
@login_required
def profile(user):
    return jsonify({"user": user.public()})


@app.put("/api/profile")
@login_required
def update_profile(user):
    data = request.get_json(silent=True) or {}
    for field in ("name", "phone", "location", "bio", "company_name", "company_description"):
        if field in data:
            setattr(user, field, str(data[field]).strip())
    db.session.commit()
    return jsonify({"user": user.public()})


@app.post("/api/profile/resume")
@login_required
@role_required("seeker")
def upload_resume(user):
    file = request.files.get("resume")
    if not file or not file.filename:
        return jsonify({"error": "Choose a resume file"}), 400
    allowed = {"pdf", "doc", "docx", "txt"}
    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
    if ext not in allowed:
        return jsonify({"error": "Only PDF, DOC, DOCX or TXT files are allowed"}), 400
    filename = f"user_{user.id}_{secure_filename(file.filename)}"
    path = UPLOAD_DIR / filename
    file.save(path)
    try:
        text = extract_resume_text(path, ext)
    except (ValueError, RuntimeError) as exc:
        path.unlink(missing_ok=True)
        return jsonify({"error": str(exc)}), 400
    if not text.strip():
        path.unlink(missing_ok=True)
        return jsonify({"error": "No readable text was found. Please upload a text-based PDF or DOCX."}), 400
    if user.resume_filename and user.resume_filename != filename:
        (UPLOAD_DIR / user.resume_filename).unlink(missing_ok=True)
    user.resume_filename = filename
    user.resume_text = text[:200000]
    db.session.commit()
    return jsonify({"user": user.public(), "url": f"/uploads/{filename}", "skills": resume_skills(text)})


@app.get("/api/resume/matches")
@login_required
@role_required("seeker")
def resume_matches(user):
    if not user.resume_filename or not (user.resume_text or "").strip():
        return jsonify({"error": "Please upload your resume first from Profile."}), 400
    jobs = Job.query.filter_by(status="Open", is_approved=True).all()
    results = []
    for job in jobs:
        score, matched = score_resume_for_job(user.resume_text, job)
        item = job.public()
        item["matchScore"] = score
        item["matchedSkills"] = matched
        results.append(item)
    results.sort(key=lambda x: (x["matchScore"], x["title"]), reverse=True)
    return jsonify({"jobs": results, "resumeSkills": resume_skills(user.resume_text),
                    "message": "Matches are based on resume text and job listing keywords."})


@app.get("/api/notifications")
@login_required
def notifications(user):
    notes = Notification.query.filter_by(user_id=user.id).order_by(Notification.created_at.desc()).limit(50).all()
    return jsonify({"notifications": [n.public() for n in notes], "unread": sum(not n.is_read for n in notes)})


@app.post("/api/notifications/read")
@login_required
def notifications_read(user):
    Notification.query.filter_by(user_id=user.id, is_read=False).update({"is_read": True})
    db.session.commit()
    return jsonify({"ok": True})


@app.get("/api/admin/stats")
@login_required
@role_required("admin")
def admin_stats(user):
    return jsonify({
        "users": User.query.count(), "seekers": User.query.filter_by(role="seeker").count(),
        "recruiters": User.query.filter_by(role="recruiter").count(), "jobs": Job.query.count(),
        "pending_jobs": Job.query.filter_by(is_approved=False).count(), "applications": Application.query.count(),
    })


@app.get("/api/admin/users")
@login_required
@role_required("admin")
def admin_users(user):
    return jsonify({"users": [u.public() for u in User.query.order_by(User.created_at.desc()).all()]})


@app.patch("/api/admin/users/<int:user_id>")
@login_required
@role_required("admin")
def admin_user_toggle(user, user_id):
    target = db.session.get(User, user_id)
    if not target:
        return jsonify({"error": "User not found"}), 404
    if target.id == user.id:
        return jsonify({"error": "You cannot disable your own admin account"}), 400
    target.is_active = not target.is_active
    db.session.commit()
    return jsonify({"user": target.public()})


@app.get("/api/admin/jobs")
@login_required
@role_required("admin")
def admin_jobs(user):
    return jsonify({"jobs": [j.public() for j in Job.query.order_by(Job.created_at.desc()).all()]})


@app.patch("/api/admin/jobs/<int:job_id>")
@login_required
@role_required("admin")
def admin_job_moderate(user, job_id):
    job = db.session.get(Job, job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    data = request.get_json(silent=True) or {}
    if "is_approved" in data:
        job.is_approved = bool(data["is_approved"])
    if "status" in data and data["status"] in {"Open", "Closed"}:
        job.status = data["status"]
    db.session.commit()
    return jsonify({"job": job.public()})


@app.cli.command("init-db")
def init_db():
    db.create_all()
    if not User.query.filter_by(email="admin@jobbridge.local").first():
        admin = User(name="JobBridge Admin", email="admin@jobbridge.local",
                     password_hash=generate_password_hash("Admin@123"), role="admin")
        db.session.add(admin)
    if User.query.filter_by(role="recruiter").count() == 0:
        recruiter = User(name="Demo Recruiter", email="recruiter@jobbridge.local",
                         password_hash=generate_password_hash("Recruiter@123"), role="recruiter",
                         company_name="Fernbank Labs", company_description="Demo company")
        seeker = User(name="Demo Seeker", email="seeker@jobbridge.local",
                      password_hash=generate_password_hash("Seeker@123"), role="seeker")
        db.session.add_all([recruiter, seeker])
        db.session.flush()
        seed = [
            ("Backend Engineer", "Engineering", "Chennai", "Build and scale our core Java services."),
            ("Product Designer", "Design", "Remote", "Own end-to-end design for our mobile app."),
            ("Data Analyst", "Data", "Bengaluru", "Turn raw usage data into product decisions."),
            ("Product Manager", "Product", "Hyderabad", "Drive roadmap for our recruiter tools."),
            ("Growth Marketer", "Marketing", "Remote", "Own acquisition campaigns end to end."),
            ("Sales Associate", "Sales", "Mumbai", "Close deals with mid-market clients."),
            ("Support Specialist", "Support", "Chennai", "Be the first line of help for our users."),
            ("Operations Lead", "Operations", "Remote", "Keep hiring and vendor logistics on track."),
        ]
        for title, category, location, desc in seed:
            db.session.add(Job(title=title, company="Fernbank Labs", category=category,
                               location=location, desc=desc, posted_by=recruiter.id))
    db.session.commit()
    print("Database initialized. Demo passwords: Admin@123 / Recruiter@123 / Seeker@123")


def seed_demo_data():
    db.create_all()
    if User.query.count() > 0:
        return
    admin = User(name="JobBridge Admin", email="admin@jobbridge.local",
                 password_hash=generate_password_hash("Admin@123"), role="admin")
    recruiter = User(name="Demo Recruiter", email="recruiter@jobbridge.local",
                     password_hash=generate_password_hash("Recruiter@123"), role="recruiter",
                     company_name="Fernbank Labs", company_description="Demo company")
    seeker = User(name="Demo Seeker", email="seeker@jobbridge.local",
                  password_hash=generate_password_hash("Seeker@123"), role="seeker")
    db.session.add_all([admin, recruiter, seeker])
    db.session.flush()
    seed = [
        ("Backend Engineer", "Engineering", "Chennai", "Build and scale our core Java services."),
        ("Product Designer", "Design", "Remote", "Own end-to-end design for our mobile app."),
        ("Data Analyst", "Data", "Bengaluru", "Turn raw usage data into product decisions."),
        ("Product Manager", "Product", "Hyderabad", "Drive roadmap for our recruiter tools."),
        ("Growth Marketer", "Marketing", "Remote", "Own acquisition campaigns end to end."),
        ("Sales Associate", "Sales", "Mumbai", "Close deals with mid-market clients."),
        ("Support Specialist", "Support", "Chennai", "Be the first line of help for our users."),
        ("Operations Lead", "Operations", "Remote", "Keep hiring and vendor logistics on track."),
    ]
    for title, category, location, desc in seed:
        db.session.add(Job(title=title, company="Fernbank Labs", category=category,
                           location=location, desc=desc, posted_by=recruiter.id))
    db.session.commit()


@app.before_request
def ensure_db():
    # Small project convenience: create tables and first-run demo data automatically.
    db.create_all()
    try:
        from sqlalchemy import inspect, text
        columns = {c["name"] for c in inspect(db.engine).get_columns("user")}
        if "resume_text" not in columns:
            db.session.execute(text("ALTER TABLE user ADD COLUMN resume_text TEXT"))
            db.session.commit()
    except Exception:
        db.session.rollback()
    seed_demo_data()


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5000)
