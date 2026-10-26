import tomllib
import json
from google import genai
from google.genai import types
import hashlib
import os
import sqlite3
import docx
from pptx import Presentation
from pypdf import PdfReader
from PIL import Image
import re
import time
from flask import Flask, render_template, request, redirect, session, url_for, jsonify
from flask import send_file
from flask_session import Session
import scorecard
from cachelib.file import FileSystemCache
import io

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-change-me")

app.config["SESSION_TYPE"] = "cachelib"
app.config["SESSION_CACHELIB"] = FileSystemCache(
    cache_dir=os.path.join(os.getcwd(), "flask_session"), threshold=500, default_timeout=0
)
app.config["SESSION_PERMANENT"] = False
Session(app)

def get_secret(name):
    """Read a key from secrets.toml (or .streamlit/secrets.toml), then the environment."""
    for config_path in [".streamlit/secrets.toml", "secrets.toml"]:
        try:
            with open(config_path, "rb") as f:
                value = tomllib.load(f).get(name)
                if value:
                    return value
        except FileNotFoundError:
            continue
    return os.environ.get(name)


api_key = get_secret("GEMINI_API_KEY")

if not api_key:
    print("WARNING: no GEMINI_API_KEY found (secrets.toml or environment variable).")

client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=90_000))
FAST_MODEL = "gemini-3.5-flash-lite"
SMART_MODEL = "gemini-3.5-flash-lite"
MODEL_NAME = FAST_MODEL

DB_NAME = "users.db"
UPLOAD_DIR = "user_uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY,
            password_hash TEXT,
            salt TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS profiles (
            email TEXT PRIMARY KEY,
            data TEXT,
            FOREIGN KEY(email) REFERENCES users(email)
        )
    """)
    conn.commit()
    conn.close()


init_db()


def hash_password(password, salt=None):
    if salt is None:
        salt = os.urandom(16).hex()
    pw_hash = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 100000
    ).hex()
    return pw_hash, salt


def create_user(email, password):
    email = email.lower().strip()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    pw_hash, salt = hash_password(password)
    try:
        c.execute(
            "INSERT INTO users (email, password_hash, salt) VALUES (?, ?, ?)",
            (email, pw_hash, salt),
        )
        conn.commit()
        conn.close()
        return True, "Account created successfully!"
    except sqlite3.IntegrityError:
        conn.close()
        return False, "An account with this email already exists."


def verify_user(email, password):
    email = email.lower().strip()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT password_hash, salt FROM users WHERE email = ?", (email,))
    row = c.fetchone()
    conn.close()

    if not row:
        return False
    stored_hash, salt = row
    pw_hash, _ = hash_password(password, salt)
    return pw_hash == stored_hash


def save_user_profile(email, profile_dict):
    email = email.lower().strip()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute(
        "INSERT OR REPLACE INTO profiles (email, data) VALUES (?, ?)",
        (email, json.dumps(profile_dict)),
    )
    conn.commit()
    conn.close()


def load_user_profile(email):
    email = email.lower().strip()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT data FROM profiles WHERE email = ?", (email,))
    row = c.fetchone()
    conn.close()
    if row:
        return json.loads(row[0])
    return None


def sync_session_to_db():
    if session.get("logged_in") and session.get("user_email"):
        profile_data = {
            "student_school": session.get("student_school", ""),
            "grade_level": session.get("grade_level", ""),
            "primary_field": session.get("primary_field", ""),
            "gpa_estimate": session.get("gpa_estimate", "N/A"),
            "gpa_weighted": session.get("gpa_weighted", "N/A"),
            "class_rank": session.get("class_rank", "N/A"),
            "test_scores": session.get("test_scores", "N/A"),
            "top_20_colleges": session.get("top_20_colleges", []),
            "four_year_plan": session.get("four_year_plan", {}),
            "chats": session.get("chats", {}),
            "plan_major": session.get("plan_major", ""),
            "saved_schools": session.get("saved_schools", []),
            "last_critique": session.get("last_critique", ""),
            "last_simulation": session.get("last_simulation", ""),
        }
        save_user_profile(session["user_email"], profile_data)

def extract_file_content(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in [".png", ".jpg", ".jpeg", ".webp"]:
            img = Image.open(file_path).convert("RGB")
            vision_prompt = "Extract and transcribe ALL visible text, tables, course schedules, grades, or written transcripts from this image with absolute precision."
            response = client.models.generate_content(
                model=MODEL_NAME, contents=[img, vision_prompt]
            )
            return f"[Visual Content Extracted via Gemini Vision]:\n{response.text}"
        elif ext == ".pptx":
            prs = Presentation(file_path)
            content = []
            for idx, slide in enumerate(prs.slides, 1):
                slide_text = [f"--- Slide {idx} ---"]
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            if paragraph.text.strip():
                                slide_text.append(paragraph.text)
                if len(slide_text) > 1:
                    content.append("\n".join(slide_text))
            return "\n\n".join(content)
        elif ext == ".docx":
            doc = docx.Document(file_path)
            content = [p.text for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    row_data = [
                        cell.text.strip().replace("\n", " ") for cell in row.cells
                    ]
                    if any(row_data):
                        content.append(" | ".join(row_data))
            return "\n".join(content)
        elif ext == ".pdf":
            reader = PdfReader(file_path)
            content = []
            for idx, page in enumerate(reader.pages, 1):
                text = page.extract_text()
                if text:
                    content.append(f"--- Page {idx} ---\n{text}")
            return "\n\n".join(content)
        else:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
    except Exception as e:
        return f"[Error parsing file: {str(e)}]"


def get_uploaded_documents_text():
    user_email = session.get("user_email", "")
    if not user_email or not os.path.exists(UPLOAD_DIR):
        return "No documents uploaded."
    user_prefix = user_email.replace("@", "_").replace(".", "_")
    docs_content = []
    for file_name in os.listdir(UPLOAD_DIR):
        if file_name.startswith(user_prefix):
            original_name = file_name.replace(f"{user_prefix}_", "")
            file_path = os.path.join(UPLOAD_DIR, file_name)
            text = cached_extract(file_path)[:8000]
            docs_content.append(f"=== FILE: {original_name} ===\n{text}")
    return "\n\n".join(docs_content)[:20000] if docs_content else "No documents uploaded."


def get_global_context():
    return f"""
    [STUDENT PROFILE & ACADEMIC DATA]
    - High School: {session.get('student_school', 'Unspecified')} | Grade: {session.get('grade_level', 'Unspecified')} | Intended Major: {session.get('primary_field', 'Unspecified')}
    - Academic Metrics: GPA {session.get('gpa_estimate', 'N/A')} (UW) / {session.get('gpa_weighted', 'N/A')} (W) | Class Rank: {session.get('class_rank', 'N/A')} | SAT/ACT: {session.get('test_scores', 'N/A')}
    - Uploaded Documents, Transcripts, & Portfolios: {get_uploaded_documents_text()}
    """

GRADES = ["9th Grade", "10th Grade", "11th Grade", "12th Grade"]
MAX_COURSES_PER_GRADE = 8
RIGOR_LEVELS = ["Regular", "Honors", "AP", "IB", "Dual Credit"]
CHAT_TOPICS = ["general", "courses", "essays", "career"]


def call_ai(contents, system=None, model=None, json_mode=False, grounded=False, retries=2):
    """One place for every Gemini call: retries on rate limits, raises a clear error otherwise."""
    config = types.GenerateContentConfig(
        system_instruction=system,
        response_mime_type="application/json" if json_mode else None,
        tools=[types.Tool(google_search=types.GoogleSearch())] if grounded else None,
    )
    last_error = None
    for attempt in range(retries + 1):
        try:
            resp = client.models.generate_content(
                model=model or SMART_MODEL, contents=contents, config=config
            )
            if resp.text:
                return resp.text
            last_error = RuntimeError(
                "The AI returned an empty response (it may have been blocked by a safety filter)."
            )
        except Exception as e:
            last_error = e
            msg = str(e)
            transient = any(
                code in msg for code in ("429", "503", "RESOURCE_EXHAUSTED", "UNAVAILABLE")
            )
            if transient and attempt < retries:
                time.sleep(2 * (attempt + 1))
                continue
            break
    raise last_error


def friendly_error(e):
    msg = str(e)
    low = msg.lower()
    if "429" in msg or "resource_exhausted" in low or "quota" in low:
        return "The AI is rate-limited right now (quota). Wait a minute and try again, or enable billing on your Gemini key."
    if "api key" in low or "api_key" in low or "401" in msg or "403" in msg or "permission" in low:
        return "The Gemini API key is missing or invalid. Check GEMINI_API_KEY."
    if "404" in msg or "not found" in low:
        return "That Gemini model name isn't available to your API key. Check SMART_MODEL / FAST_MODEL at the top of main.py."
    if "timed out" in low or "timeout" in low or "deadline" in low:
        return "The AI took too long to respond. Please try again."
    return f"The AI request failed: {msg[:300]}"

def call_ai_with_search(contents, **kwargs):
    """Try Google Search grounding first; fall back to a normal answer if search isn't available."""
    try:
        return call_ai(contents, grounded=True, retries=0, **kwargs)
    except Exception as e:
        print(f"[Search grounding unavailable, using a normal answer]: {e}")
        return call_ai(contents, **kwargs)


def lookup_school_courses(school, major):
    """Search the web for what this specific high school actually offers."""
    if not school:
        return ""
    try:
        return call_ai(
            f"Search the web for the course catalog or program of studies for the high school '{school}'. "
            f"List the AP, IB, Honors, and Dual Credit courses it offers, especially for a student interested in {major}. "
            "Only list courses you found evidence for. If you can't find this school's catalog, reply exactly: NOT FOUND.",
            model=SMART_MODEL, grounded=True, retries=0,
        )
    except Exception as e:
        print(f"[School course lookup failed]: {e}")
        return ""

def normalize_grade(value):
    m = re.search(r"\b(9|10|11|12)", str(value or ""))
    return f"{m.group(1)}th Grade" if m else None


def reset_plan_if_major_changed():
    if session.get("plan_major") and session.get("plan_major") != session.get("primary_field"):
        session["four_year_plan"] = {}


_extract_cache = {}


def cached_extract(file_path):
    """Parse each uploaded file once (images cost a Gemini call every time otherwise)."""
    try:
        key = (file_path, os.path.getmtime(file_path))
    except OSError:
        return ""
    if key not in _extract_cache:
        _extract_cache[key] = extract_file_content(file_path)
    return _extract_cache[key]


def clean_json_response(raw_text):
    cleaned = raw_text.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()

    try:
        return json.loads(cleaned)
    except Exception:
        match_arr = re.search(r"\[.*\]", cleaned, re.DOTALL)
        if match_arr:
            try:
                return json.loads(match_arr.group(0))
            except Exception:
                pass
        match_obj = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match_obj:
            try:
                return json.loads(match_obj.group(0))
            except Exception:
                pass
    raise ValueError(f"Unable to parse JSON from AI response: {raw_text[:200]}")


def enrich_with_scorecard(colleges):
    """Replace AI guesses with real College Scorecard data. Falls back to AI-only if unavailable."""
    key = get_secret("SCORECARD_API_KEY")
    if not key:
        print("WARNING: no SCORECARD_API_KEY found, so college stats are AI-only. Free key: https://api.data.gov/signup/")
        return colleges
    sat = scorecard.parse_test_score(session.get("test_scores"))
    gpa = scorecard.parse_gpa(session.get("gpa_estimate"))
    try:
        return scorecard.enrich_colleges(colleges, key, DB_NAME, sat, gpa)
    except Exception as e:
        print(f"[Scorecard Error]: {e}")
        return colleges


@app.template_filter("pct")
def _pct(v):
    if not isinstance(v, (int, float)):
        return "\u2014"
    return f"{v * 100:.0f}%" if v >= 0.1 else f"{v * 100:.1f}%"

@app.template_filter("sat")
def _sat(v):
    """Turn whatever the student typed (SAT or ACT) into a SAT-equivalent number, or None."""
    return scorecard.parse_test_score(v)


@app.template_filter("usd")
def _usd(v):
    return f"${v:,.0f}" if isinstance(v, (int, float)) else "\u2014"


@app.template_filter("num")
def _num(v):
    return f"{v:,.0f}" if isinstance(v, (int, float)) else "\u2014"


def GenerateTop20Colleges():
    prompt = f"""
    Analyze high school: '{session.get('student_school', 'High School')}' and student profile.
    Student Profile: {get_global_context()}

    CRITICAL INSTRUCTION ON GRADE-LEVEL TRAJECTORY:
    - Take the student's grade level into account. If the student is a Freshman (9th) or Sophomore (10th), evaluate them on TRAJECTORY and POTENTIAL, not final senior-year metrics.
    - Do NOT penalize underclassmen for missing class ranks, unestablished GPAs, or lack of test scores.

    Synthesize a comprehensive, highly tailored list of exactly 20 diverse college matches categorized into Safety, Target, and Reach.
    Aim for roughly 4 Safety, 8 Target, and 8 Reach. Be realistic: a Safety must be a school the student is very likely to get into.

    CRITICAL REQUIREMENT: Return your response as a raw JSON object ONLY, containing a "colleges" array with exactly 20 objects. No markdown formatting outside JSON.
    Each object must have these exact keys:
    - "name": The school's official name as listed by the US Dept. of Education College Scorecard (e.g. "University of California-Berkeley", "Georgia Institute of Technology-Main Campus"). Only real US colleges.
    - "category": Exactly "Safety", "Target", or "Reach"
    - "strategy": 1-sentence strategic admissions advice appropriate for their grade level
    - "reason": 2-3 sentences on academic and extracurricular fit for THIS student's intended major
    """
    try:
        text = call_ai(prompt, model=SMART_MODEL, json_mode=True)
        parsed = clean_json_response(text)
        raw = parsed.get("colleges", []) if isinstance(parsed, dict) else parsed
        cleaned = []
        for c in raw:
            if not isinstance(c, dict) or not c.get("name"):
                continue
            category = str(c.get("category", "Target")).strip().title()
            if category not in ("Safety", "Target", "Reach"):
                category = "Target"
            cleaned.append(
                {
                    "name": c["name"],
                    "category": category,
                    "strategy": c.get("strategy", ""),
                    "reason": c.get("reason") or c.get("rationale", ""),
                }
            )
        return enrich_with_scorecard(cleaned)
    except Exception as e:
        print(f"[AI Research Engine Error - Colleges]: {e}")
        return []


def generate_4year_plan(intended_major):
    catalog = lookup_school_courses(session.get("student_school"), intended_major)
    prompt = f"""   
    You are an experienced high school counselor building a realistic, competitive 4-year course plan
    for a student aiming at selective colleges.

    Student Context & Academic Metrics:
    {get_global_context()}

    Rules:
    1. Intended major: {intended_major}. Tailor rigor and electives to it, but keep a balanced core.
    2. Return EXACTLY {MAX_COURSES_PER_GRADE} courses for EACH of the 4 grades (24 total). Every year needs English, Math, Science,
       Social Studies, and a World Language where it fits; fill the rest with major-relevant electives (PE/Health is fine in 9th-10th).
    3. Use only courses a typical US public high school actually offers (AP, Honors, Dual Credit, or standard courses).
    4. Be realistic about workload: ramp rigor up over time (e.g. 1-2 AP/Honors in 9th-10th, more in 11th-12th). Do not put 6 APs in one year.
    5. If a transcript or uploaded schedule shows courses already taken in a completed year, keep those courses as they are.
    6. Respect the student's current grade level and GPA; do not set up a plan they could not realistically handle.
    7. Here is what we found about the student's real high school course offerings. Prefer courses from this list, and ignore it if it says NOT FOUND:
    {catalog}

    CRITICAL REQUIREMENT: Return ONLY a valid JSON array of 24 course objects, no markdown.
    Each object MUST have:
    - "Grade": exactly one of "9th Grade", "10th Grade", "11th Grade", "12th Grade"
    - "Subject": e.g. Mathematics, Science, English
    - "Course": the specific course title
    - "Rigor": exactly one of "Regular", "Honors", "AP", "IB", "Dual Credit"
    """
    try:
        text = call_ai(prompt, model=SMART_MODEL, json_mode=True)
        parsed = clean_json_response(text)
        if isinstance(parsed, dict):
            lists = [v for v in parsed.values() if isinstance(v, list)]
            parsed = lists[0] if lists else []
        return parsed
    except Exception as e:
        print(f"[AI Research Engine Error - Courses]: {e}")
        return []

@app.route("/")
def index():
    if not session.get("logged_in"):
        return render_template("login.html")
    if not session.get("setup_complete"):
        return render_template("assessment.html")

    colleges = session.get("top_20_colleges", [])
    email = session.get("user_email", "")
    return render_template("dashboard.html", colleges=colleges, email=email, saved=session.get("saved_schools", []))


@app.route("/login", methods=["POST"])
def login():
    email = request.form.get("email")
    password = request.form.get("password")
    if verify_user(email, password):
        session["logged_in"] = True
        session["user_email"] = email.lower().strip()
        user_data = load_user_profile(email)
        if user_data:
            for key, val in user_data.items():
                session[key] = val
            session["setup_complete"] = True
        return redirect(url_for("index"))
    return render_template("login.html", error="Invalid credentials")


@app.route("/signup", methods=["POST"])
def signup():
    email = request.form.get("email")
    password = request.form.get("password")
    success, msg = create_user(email, password)
    if success:
        return render_template("login.html", success="Account created! Please sign in.")
    return render_template("login.html", error=msg)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.route("/assessment", methods=["GET", "POST"])
def assessment():
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    if request.method == "POST":
        session["student_school"] = request.form.get("student_school")
        session["grade_level"] = request.form.get("grade_level")
        session["primary_field"] = request.form.get("primary_field")
        session["gpa_estimate"] = request.form.get("gpa_estimate", "N/A")
        session["gpa_weighted"] = request.form.get("gpa_weighted", "N/A")
        session["class_rank"] = request.form.get("class_rank", "N/A")
        session["test_scores"] = request.form.get("test_scores", "N/A")
        session["setup_complete"] = True

        colleges = GenerateTop20Colleges()
        if colleges:
            session["top_20_colleges"] = colleges
        reset_plan_if_major_changed()

        sync_session_to_db()
        return redirect(url_for("index"))

    return render_template("assessment.html")


@app.route("/chat", defaults={"topic": "general"}, methods=["GET", "POST"])
@app.route("/chat/<topic>", methods=["GET", "POST"])
def advisor_chat(topic):
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    chats = session.get("chats") or {}
    for t in CHAT_TOPICS:
        chats.setdefault(t, [])
    if topic not in chats:
        topic = "general"

    if request.method == "POST":
        user_message = request.form.get("message", "").strip()
        if user_message:
            chats[topic].append({"role": "user", "text": user_message})

            personas = {
                "general": "You are Pathway's college admissions counselor and research expert.",
                "courses": "You are Pathway's high school course planner and academic rigor advisor.",
                "essays": "You are Pathway's college essay coach and admissions storytelling expert. Coach the student; never write their essay for them.",
                "career": "You are Pathway's career and economics mentor for high school students.",
            }
            system = f"""{personas.get(topic, personas['general'])}

Guidelines:
- Give specific, actionable advice grounded in the student's actual profile below. Be honest and realistic; do not just flatter.
- If you are missing something important, ask ONE short clarifying question.
- Keep answers under about 250 words unless the student asks for more. Use short paragraphs or bullets.
- Admission statistics and costs change; say so when you are estimating.
- Use Google Search for anything current (deadlines, requirements, costs, program details) and say when you relied on it.

Student context:
{get_global_context()}
"""
            history = [m for m in chats[topic][-12:] if not m.get("error")]
            while history and history[0]["role"] != "user":
                history.pop(0)
            contents = [
                types.Content(role=m["role"], parts=[types.Part(text=m["text"])])
                for m in history
            ]

            try:
                if topic == "essays":
                    reply = call_ai(contents, system=system, model=SMART_MODEL)
                else:
                    reply = call_ai_with_search(contents, system=system, model=SMART_MODEL)
                chats[topic].append({"role": "model", "text": reply})
            except Exception as e:
                chats[topic].append(
                    {"role": "model", "text": friendly_error(e), "error": True}
                )

            session["chats"] = chats
            session.modified = True
            sync_session_to_db()

        return redirect(url_for("advisor_chat", topic=topic))

    session["chats"] = chats
    return render_template("chat.html", history=chats[topic], active_topic=topic)


@app.route("/courses")
def courses_page():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    plan_error = None
    plan = session.get("four_year_plan") or {}

    if not any(plan.values()):
        major = session.get("primary_field", "General STEM")
        raw_list = generate_4year_plan(major)
        prefilled = {g: [] for g in GRADES}
        for item in raw_list if isinstance(raw_list, list) else []:
            if not isinstance(item, dict) or not item.get("Course"):
                continue
            g = normalize_grade(item.get("Grade"))
            if g and len(prefilled[g]) < MAX_COURSES_PER_GRADE:
                rigor = item.get("Rigor", "Regular")
                if rigor not in RIGOR_LEVELS:
                    rigor = "Regular"
                prefilled[g].append({"Course": item["Course"], "Rigor": rigor})

        if any(prefilled.values()):
            session["four_year_plan"] = prefilled
            session["plan_major"] = major
            sync_session_to_db()
        else:
            plan_error = "The AI couldn't build a plan this time. Refresh the page to try again (check the terminal for the error)."

    return render_template("courses.html", plan_error=plan_error)


MAX_SAVED = 6

@app.route("/feedback", methods=["POST"])
def submit_feedback():
    if not session.get("logged_in"):
        return jsonify(error="Please sign in again."), 401
    data = request.get_json(silent=True) or {}
    rating = data.get("rating")
    if rating not in ("up", "down"):
        return jsonify(error="Invalid rating."), 400
    comment = (data.get("comment") or "").strip()[:500]
    conn = sqlite3.connect(DB_NAME)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS feedback (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT, rating TEXT, comment TEXT, created TEXT DEFAULT CURRENT_TIMESTAMP)"
    )
    conn.execute(
        "INSERT INTO feedback (email, rating, comment) VALUES (?, ?, ?)",
        (session.get("user_email"), rating, comment),
    )
    conn.commit()
    conn.close()
    return jsonify(ok=True)


@app.route("/feedback/summary")
def feedback_summary():
    """Open this in your browser while logged in to see your numbers."""
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    conn = sqlite3.connect(DB_NAME)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS feedback (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT, rating TEXT, comment TEXT, created TEXT DEFAULT CURRENT_TIMESTAMP)"
    )
    rows = conn.execute("SELECT rating, COUNT(*) FROM feedback GROUP BY rating").fetchall()
    comments = conn.execute(
        "SELECT rating, comment FROM feedback WHERE comment != '' ORDER BY id DESC LIMIT 20"
    ).fetchall()
    conn.close()
    counts = dict(rows)
    up, down = counts.get("up", 0), counts.get("down", 0)
    total = up + down
    pct_up = round(100 * up / total) if total else 0
    return jsonify(total=total, thumbs_up=up, thumbs_down=down, percent_positive=pct_up, comments=comments)

@app.route("/colleges/refresh", methods=["POST"])
def colleges_refresh():
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    new_colleges = GenerateTop20Colleges()
    if new_colleges:
        session["top_20_colleges"] = new_colleges
        names = {c["name"] for c in new_colleges}
        session["saved_schools"] = [n for n in session.get("saved_schools", []) if n in names]
        sync_session_to_db()
    return redirect(url_for("index"))

@app.route("/debug/state")
def debug_state():
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    colleges = session.get("top_20_colleges", [])
    return jsonify(
        saved_schools=session.get("saved_schools", []),
        total_colleges=len(colleges),
        with_scorecard_data=sum(1 for c in colleges if c.get("data_source") == "scorecard"),
        scorecard_key_found=bool(get_secret("SCORECARD_API_KEY")),
        first_college=colleges[0] if colleges else None,
    )

@app.route("/schools/toggle", methods=["POST"])
def toggle_school():
    if not session.get("logged_in"):
        return jsonify(error="Please sign in again."), 401
    name = ((request.get_json(silent=True) or {}).get("name") or "").strip()
    known = {c["name"] for c in session.get("top_20_colleges", [])}
    if name not in known:
        return jsonify(error="That school isn't on your list."), 400
    saved = list(session.get("saved_schools", []))
    if name in saved:
        saved.remove(name)
        is_saved = False
    elif len(saved) >= MAX_SAVED:
        return jsonify(error=f"You can compare up to {MAX_SAVED} schools. Remove one first."), 400
    else:
        saved.append(name)
        is_saved = True
    session["saved_schools"] = saved
    session.modified = True
    sync_session_to_db()
    return jsonify(saved=is_saved, count=len(saved))


@app.route("/compare")
def compare_page():
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    by_name = {c["name"]: c for c in session.get("top_20_colleges", [])}
    schools = [by_name[n] for n in session.get("saved_schools", []) if n in by_name]
    return render_template("compare.html", schools=schools)


@app.route("/courses/update", methods=["POST"])
def courses_update():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    updated_plan = {
        "9th Grade": [],
        "10th Grade": [],
        "11th Grade": [],
        "12th Grade": [],
    }

    for grade in updated_plan.keys():
        for i in range(1, 7):
            course_name = request.form.get(f"course_{grade}_{i}")
            rigor = request.form.get(f"rigor_{grade}_{i}")
            if course_name:
                updated_plan[grade].append(
                    {"Course": course_name, "Rigor": rigor or "Regular"}
                )

    session["four_year_plan"] = updated_plan
    sync_session_to_db()
    return redirect(url_for("courses_page"))


@app.route("/courses/reset", methods=["POST"])
def courses_reset():
    session["four_year_plan"] = {}
    sync_session_to_db()
    return redirect(url_for("courses_page"))


@app.route("/essays")
def essays_page():
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    return render_template("essays.html", critique_result=session.get("last_critique"))


@app.route("/essays/critique", methods=["POST"])
def essays_critique():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    essay_prompt = request.form.get("essay_prompt", "").strip()
    essay_text = request.form.get("essay_text", "").strip()

    if len(essay_text.split()) < 30:
        session["last_critique"] = "Your draft is very short. Paste at least a few paragraphs so there's something to critique."
        return redirect(url_for("essays_page"))

    prompt = f"""
    You are an experienced college admissions reader and writing coach.
    Critique the student's essay below. Only comment on what is actually written in the draft.

    ESSAY PROMPT: {essay_prompt}

    STUDENT BACKGROUND (context only):
    {get_global_context()}

    <essay_draft>
    {essay_text}
    </essay_draft>

    Give structured feedback with these headings:
    1. Overall Impression & Strengths (what is genuinely working)
    2. Does it answer the prompt? (be direct)
    3. Voice & Authenticity
    4. Show, Don't Tell (quote short phrases from the draft and say how to make them more concrete)
    5. Structure, Opening & Ending
    6. Top 3 Priorities for the next revision

    Rules: quote the student's own words when giving line-level notes. Do NOT rewrite the whole essay
    or write new paragraphs for them; coach them to improve it themselves.
    """
    try:
        session["last_critique"] = call_ai(prompt, model=SMART_MODEL)
    except Exception as e:
        session["last_critique"] = friendly_error(e)

    session.modified = True
    return redirect(url_for("essays_page"))


@app.route("/career")
def career_page():
    if not session.get("logged_in"):
        return redirect(url_for("index"))
    return render_template(
        "career.html", simulation_result=session.get("last_simulation")
    )


@app.route("/career/simulate", methods=["POST"])
def career_simulate():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    target_role = request.form.get("target_role", "").strip()
    target_city = request.form.get("target_city", "").strip()

    session["last_target_role"] = target_role
    session["last_target_city"] = target_city

    prompt = f"""
    Build a realistic career and financial simulation for a new graduate.
    Role: {target_role}
    City: {target_city}
    Student's intended field: {session.get('primary_field', 'General')}

    Use current data where you can. Give ranges, not single numbers, and label everything as an estimate.
    Include these sections (use markdown headings and a table for the budget):
    1. Typical entry-level salary range for this role in {target_city}
    2. Monthly budget estimate: rent (1-bedroom), federal/state/payroll taxes, utilities, food, transportation, other, savings
    3. Estimated monthly take-home and net discretionary income
    4. 5-year career outlook: typical promotions, salary growth, skills to build
    5. One honest caveat (e.g. cost-of-living swings, job market competitiveness)
    """
    try:
        try:
            text = call_ai(prompt, model=SMART_MODEL, grounded=True, retries=0)
        except Exception as grounded_error:
            print(f"[Career grounded search failed, falling back]: {grounded_error}")
            text = call_ai(prompt, model=SMART_MODEL)
        session["last_simulation"] = text
    except Exception as e:
        session["last_simulation"] = friendly_error(e)

    session.modified = True
    sync_session_to_db()
    return redirect(url_for("career_page"))


@app.route("/upload", methods=["POST"])
def upload_file():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    uploaded_file = request.files.get("file")
    if uploaded_file and uploaded_file.filename:
        user_email = session.get("user_email", "user")
        user_prefix = user_email.replace("@", "_").replace(".", "_")
        safe_filename = f"{user_prefix}_{uploaded_file.filename}"
        file_path = os.path.join(UPLOAD_DIR, safe_filename)
        uploaded_file.save(file_path)
        session.modified = True
        sync_session_to_db()

    return redirect(url_for("index"))


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    if request.method == "POST":
        session["student_school"] = request.form.get(
            "student_school", session.get("student_school")
        )
        session["grade_level"] = request.form.get(
            "grade_level", session.get("grade_level")
        )
        session["primary_field"] = request.form.get(
            "primary_field", session.get("primary_field")
        )
        session["gpa_estimate"] = request.form.get(
            "gpa_estimate", session.get("gpa_estimate")
        )
        session["gpa_weighted"] = request.form.get(
            "gpa_weighted", session.get("gpa_weighted")
        )
        session["class_rank"] = request.form.get(
            "class_rank", session.get("class_rank")
        )
        session["test_scores"] = request.form.get(
            "test_scores", session.get("test_scores")
        )

        new_colleges = GenerateTop20Colleges()
        if new_colleges:
            session["top_20_colleges"] = new_colleges
        reset_plan_if_major_changed()

        sync_session_to_db()
        return redirect(url_for("index"))

    return render_template("settings.html")


@app.route("/export/report")
def export_report():
    if not session.get("logged_in"):
        return redirect(url_for("index"))

    report_content = f"""
    ==================================================
    PATHWAY AI - ADMISSIONS & STRATEGY REPORT
    Student: {session.get('user_email')}
    High School: {session.get('student_school')}
    Intended Major: {session.get('primary_field')}
    GPA: {session.get('gpa_estimate')} (UW) / {session.get('gpa_weighted')} (W)
    ==================================================

    --- TOP COLLEGE MATCHES ---
    """
    for col in session.get("top_20_colleges", []):
        report_content += f"\n• {col.get('name')} ({col.get('category')} - Match: {col.get('match_score')}%)\n  Strategy: {col.get('strategy')}\n"

    report_content += "\n--- 4-YEAR COURSE PLAN ---\n"
    plan = session.get("four_year_plan", {})
    for grade, courses in plan.items():
        report_content += f"\n[{grade}]\n"
        for c in courses:
            report_content += f"  - {c.get('Course')} ({c.get('Rigor')})\n"

    buffer = io.BytesIO()
    buffer.write(report_content.encode("utf-8"))
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="pathway_admissions_report.txt",
        mimetype="text/plain",
    )


if __name__ == "__main__":
    app.run(debug=True)
