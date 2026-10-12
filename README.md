# JobBridge — Flask backend extension

This project keeps the supplied `templates/Pasted code.html` unchanged. Flask serves that original page and injects `static/app_extension.js` at runtime.

## Resume-based job matching

Job seekers can open **Profile**, upload a PDF/DOCX/TXT resume, and then click **Find jobs from my resume**. The backend extracts readable resume text and ranks open approved jobs using an explainable keyword/skill relevance score. It also displays detected and matched skills.

The matching feature is a relevance aid, not an automated hiring decision.

PDF/DOCX extraction requires `pypdf` and `python-docx`. Legacy `.doc` files can be stored but are not parsed for matching; upload PDF or DOCX for matching.

## Modules

- Authentication: registration, login, logout, sessions, role access
- Job Seeker: profile, resume upload, resume-based job matching, applications
- Recruiter: company profile, job posting/editing/deletion, applicant review, resume access for applicants to their jobs, status updates
- Admin: user moderation, job moderation, statistics
- Search & Filter: keyword/category/location filtering and sorting
- Notification: stored notifications and unread count with periodic refresh
- Database: SQLAlchemy with SQLite by default; MySQL supported through `DATABASE_URL`

## Run

```bash
python -m venv venv
venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`.

Demo accounts are created automatically:

- Admin: `admin@jobbridge.local` / `Admin@123`
- Recruiter: `recruiter@jobbridge.local` / `Recruiter@123`
- Seeker: `seeker@jobbridge.local` / `Seeker@123`
