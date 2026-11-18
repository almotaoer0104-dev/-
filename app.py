import os, re, io, json, time, html, sqlite3, hashlib, secrets
from pathlib import Path
from dotenv import load_dotenv
from fastapi import FastAPI, Request, UploadFile, File, Form, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from starlette.middleware.sessions import SessionMiddleware
from docx import Document
from docx.oxml import OxmlElement

load_dotenv()
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "").lower()
MAX_MB = 10
ALLOWED = {".txt", ".md", ".docx"}
DATA = Path("data"); (DATA / "up").mkdir(parents=True, exist_ok=True); (DATA / "out").mkdir(exist_ok=True)
app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=os.getenv("SECRET_KEY", secrets.token_hex(32)),
                   max_age=86400 * 7, same_site="lax", https_only=False)

# ---------- قاعدة البيانات ----------
def db():
    c = sqlite3.connect(DATA / "app.db"); c.row_factory = sqlite3.Row; return c

with db() as c:
    c.executescript("""
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, email TEXT UNIQUE, pw TEXT, salt TEXT, is_admin INT DEFAULT 0, created TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS jobs(id INTEGER PRIMARY KEY, user_id INT, filename TEXT, status TEXT, progress INT DEFAULT 0, src TEXT, out TEXT, error TEXT, chars INT DEFAULT 0, created TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS glossary(id INTEGER PRIMARY KEY, en TEXT, ar TEXT);
    """)

def hash_pw(pw, salt): return hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 200_000).hex()
def me(req):
    uid = req.session.get("uid")
    if not uid: return None
    with db() as c: return c.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
esc = html.escape

# ---------- محرك الترجمة (محلي بالكامل، بدون أي API) ----------
import argostranslate.package, argostranslate.translate as argos

def ensure_model():
    """يثبّت نموذج en→ar مرة واحدة (يحتاج إنترنت أول مرة فقط)، وبعدها يعمل دون اتصال."""
    if any(l.code == "ar" for l in argos.get_installed_languages()
           for t in l.translations_from if t.from_lang.code == "en"):
        return
    argostranslate.package.update_package_index()
    pkg = next(p for p in argostranslate.package.get_available_packages()
               if p.from_code == "en" and p.to_code == "ar")
    argostranslate.package.install_from_path(pkg.download())

ensure_model()

PROTECT = re.compile(r"(https?://\S+|www\.\S+|[\w.+-]+@[\w-]+\.[\w.]+|`[^`]*`|\{[^}]*\}|%[sd]|<[^>]+>)")

def glossary_text():
    with db() as c: return [(r["en"], r["ar"]) for r in c.execute("SELECT en, ar FROM glossary")]

def translate_one(text, gloss):
    keep = []
    def stash(val):
        keep.append(val); return f" XQ{len(keep)-1}X "
    text = PROTECT.sub(lambda m: stash(m.group(0)), text)          # حماية الروابط والأكواد
    for en, ar in gloss:                                            # فرض المصطلحات المعتمدة
        text = re.sub(rf"\b{re.escape(en)}\b", lambda m: stash(ar), text, flags=re.I)
    out = argos.translate(text, "en", "ar")
    out = re.sub(r"XQ\s*(\d+)\s*X", lambda m: keep[int(m.group(1))] if int(m.group(1)) < len(keep) else "", out, flags=re.I)
    return re.sub(r"[ \t]{2,}", " ", out).strip()

def call_model(texts, gloss):
    return [translate_one(t, gloss) for t in texts]

def translate_list(texts, progress):
    """يترجم قائمة نصوص على دفعات، ويتخطى ما لا يحتاج ترجمة."""
    gloss = glossary_text()
    idx = [i for i, t in enumerate(texts) if re.search(r"[A-Za-z]{2,}", t)]
    result = list(texts); batch, size, done = [], 0, 0
    def flush():
        nonlocal batch, size, done
        if not batch: return
        out = call_model([texts[i] for i in batch], gloss)
        for i, t in zip(batch, out): result[i] = t
        done += len(batch); progress(int(done * 100 / max(len(idx), 1))); batch, size = [], 0
    for i in idx:
        if size + len(texts[i]) > 3500 or len(batch) >= 40: flush()
        batch.append(i); size += len(texts[i])
    flush(); return result

def set_rtl(p):
    pPr = p._p.get_or_add_pPr()
    if pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi") is None:
        pPr.append(OxmlElement("w:bidi"))
    for r in p.runs:
        rPr = r._r.get_or_add_rPr()
        rPr.append(OxmlElement("w:rtl"))

def all_paragraphs(doc):
    def walk(container):
        for p in container.paragraphs: yield p
        for t in getattr(container, "tables", []):
            for row in t.rows:
                for cell in row.cells: yield from walk(cell)
    yield from walk(doc)
    for s in doc.sections:
        for part in (s.header, s.footer): yield from walk(part)

def do_docx(src, dst, progress):
    doc = Document(src)
    seen, paras = set(), []
    for p in all_paragraphs(doc):
        if id(p._p) not in seen and p.text.strip(): seen.add(id(p._p)); paras.append(p)
    out = translate_list([p.text for p in paras], progress)
    for p, t in zip(paras, out):
        if p.runs:
            p.runs[0].text = t          # يحافظ على خط وحجم ولون أول مقطع
            for r in p.runs[1:]: r.text = ""
        else: p.add_run(t)
        set_rtl(p)
    doc.save(dst)

def do_text(src, dst, progress):
    lines = Path(src).read_text(encoding="utf-8", errors="replace").splitlines()
    todo, fence = [], False
    for i, ln in enumerate(lines):
        if ln.strip().startswith("```"): fence = not fence; continue
        if not fence and ln.strip(): todo.append(i)
    prefixes = [re.match(r"^(\s*(?:[#>*\-+]+|\d+[.)])?\s*)", lines[i]).group(1) for i in todo]
    bodies = [lines[i][len(p):] for i, p in zip(todo, prefixes)]
    out = translate_list(bodies, progress)
    for i, p, t in zip(todo, prefixes, out): lines[i] = p + t
    Path(dst).write_text("\n".join(lines), encoding="utf-8")

def put_text(p, t):
    """يضع الترجمة في أول مقطع (للحفاظ على الخط واللون) ويفرغ الباقي."""
    if p.runs:
        p.runs[0].text = t
        for r in p.runs[1:]: r.text = ""
    else: p.add_run(t)

def do_pptx(src, dst, progress):
    from pptx import Presentation
    prs = Presentation(src)
    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6: yield from walk(sh.shapes)
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame: yield from sh.text_frame.paragraphs
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells: yield from cell.text_frame.paragraphs
    paras = [p for sl in prs.slides for p in walk(sl.shapes) if "".join(r.text for r in p.runs).strip()]
    out = translate_list(["".join(r.text for r in p.runs) for p in paras], progress)
    for p, t in zip(paras, out):
        put_text(p, t); pPr = p._p.get_or_add_pPr(); pPr.set("rtl", "1"); pPr.set("algn", "r")
    prs.save(dst)

def do_xlsx(src, dst, progress):
    import openpyxl
    wb = openpyxl.load_workbook(src)
    cells = [c for ws in wb for row in ws.iter_rows() for c in row
             if c.data_type == "s" and isinstance(c.value, str) and c.value.strip()]
    out = translate_list([c.value for c in cells], progress)
    for c, t in zip(cells, out): c.value = t
    for ws in wb: ws.sheet_view.rightToLeft = True
    wb.save(dst)

def do_csv(src, dst, progress):
    import csv
    rows = list(csv.reader(io.StringIO(Path(src).read_text(encoding="utf-8-sig", errors="replace"))))
    pos = [(i, j) for i, r in enumerate(rows) for j, v in enumerate(r) if v.strip()]
    out = translate_list([rows[i][j] for i, j in pos], progress)
    for (i, j), t in zip(pos, out): rows[i][j] = t
    with open(dst, "w", encoding="utf-8-sig", newline="") as f: csv.writer(f).writerows(rows)

def do_json(src, dst, progress):
    data = json.loads(Path(src).read_text(encoding="utf-8-sig"))
    refs = []
    def walk(o):
        for k, v in (o.items() if isinstance(o, dict) else enumerate(o) if isinstance(o, list) else []):
            if isinstance(v, str): refs.append((o, k))
            else: walk(v)
    walk(data)
    out = translate_list([o[k] for o, k in refs], progress)
    for (o, k), t in zip(refs, out): o[k] = t
    Path(dst).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def do_subs(src, dst, progress):
    lines = Path(src).read_text(encoding="utf-8-sig", errors="replace").splitlines()
    idx = [i for i, l in enumerate(lines) if l.strip() and not l.strip().isdigit()
           and "-->" not in l and not l.startswith(("WEBVTT", "NOTE"))]
    out = translate_list([lines[i] for i in idx], progress)
    for i, t in zip(idx, out): lines[i] = t
    Path(dst).write_text("\n".join(lines), encoding="utf-8")

def do_html(src, dst, progress):
    from bs4 import BeautifulSoup, NavigableString
    soup = BeautifulSoup(Path(src).read_text(encoding="utf-8", errors="replace"), "html.parser")
    nodes = [n for n in soup.find_all(string=True) if type(n) is NavigableString
             and n.parent.name not in ("script", "style", "code", "pre", "textarea") and n.strip()]
    out = translate_list([str(n) for n in nodes], progress)
    for n, t in zip(nodes, out): n.replace_with(t)
    for tag in soup.find_all("html"): tag["dir"] = "rtl"; tag["lang"] = "ar"
    Path(dst).write_text(str(soup), encoding="utf-8")

def do_pdf(src, dst, progress):
    """PDF نصي ← ملف DOCX عربي (يُستخرج النص بالترتيب). الملفات الممسوحة ضوئياً تحتاج OCR."""
    import fitz
    paras = []
    for page in fitz.open(src):
        for b in page.get_text("blocks", sort=True):
            t = " ".join(b[4].split()) if b[6] == 0 else ""
            if t: paras.append(t)
    if not paras: raise RuntimeError("لا يوجد نص قابل للاستخراج (PDF ممسوح ضوئياً)")
    d = Document()
    for t in translate_list(paras, progress): set_rtl(d.add_paragraph(t))
    d.save(dst)

H = {".docx": do_docx, ".pptx": do_pptx, ".xlsx": do_xlsx, ".pdf": do_pdf, ".html": do_html, ".htm": do_html,
     ".csv": do_csv, ".json": do_json, ".srt": do_subs, ".vtt": do_subs}
for _e in (".txt", ".md", ".markdown", ".rst", ".log", ".ini", ".yaml", ".yml", ".tex", ".tsv"): H[_e] = do_text
ACCEPT = ",".join(sorted(H))

def run_job(job_id):
    with db() as c: j = c.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    def prog(n):
        with db() as c: c.execute("UPDATE jobs SET progress=? WHERE id=?", (n, job_id))
    try:
        ext = Path(j["filename"]).suffix.lower()
        dst = str(DATA / "out" / f"{job_id}_ar{'.docx' if ext == '.pdf' else ext}")
        H[ext](j["src"], dst, prog)
        with db() as c: c.execute("UPDATE jobs SET status='done', progress=100, out=? WHERE id=?", (dst, job_id))
    except Exception as e:
        with db() as c: c.execute("UPDATE jobs SET status='error', error=? WHERE id=?", (str(e)[:300], job_id))

# ---------- الواجهة ----------
CSS = """*{box-sizing:border-box}body{margin:0;font-family:'Segoe UI',Tahoma,sans-serif;background:#f4f6fa;color:#1d2433}
nav{background:#1f3a5f;color:#fff;padding:14px 20px;display:flex;gap:16px;align-items:center}nav a{color:#fff;text-decoration:none}nav b{flex:1}
main{max-width:900px;margin:24px auto;padding:0 16px}.card{background:#fff;border-radius:14px;padding:20px;margin-bottom:16px;box-shadow:0 2px 8px #0001}
input,button{font:inherit;padding:10px 12px;border-radius:8px;border:1px solid #cbd3e1;width:100%;margin:6px 0}
button,.btn{background:#2f6fed;color:#fff;border:0;cursor:pointer;text-decoration:none;display:inline-block;padding:8px 14px;border-radius:8px;width:auto}
table{width:100%;border-collapse:collapse;font-size:14px}td,th{padding:8px;border-bottom:1px solid #e5e9f2;text-align:right}
.bar{background:#e5e9f2;border-radius:8px;height:10px}.bar i{display:block;height:10px;background:#2f6fed;border-radius:8px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px}.stat{text-align:center}.stat h2{margin:0;color:#2f6fed}.err{color:#c0392b}"""

def page(body, user=None, refresh=False):
    nav = "<b>🌐 مترجم الملفات</b>"
    if user:
        nav += "<a href='/'>ملفاتي</a>" + ("<a href='/admin'>لوحة التحكم</a>" if user["is_admin"] else "") + "<a href='/logout'>خروج</a>"
    js = "<script>setTimeout(()=>location.reload(),3000)</script>" if refresh else ""
    return HTMLResponse(f"<!doctype html><html lang='ar' dir='rtl'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>مترجم الملفات</title><style>{CSS}</style></head><body><nav>{nav}</nav><main>{body}</main>{js}</body></html>")

def auth_form(kind, err=""):
    t = "تسجيل الدخول" if kind == "login" else "إنشاء حساب"
    other = "<a href='/register'>إنشاء حساب جديد</a>" if kind == "login" else "<a href='/login'>لديّ حساب</a>"
    return page(f"<div class='card'><h2>{t}</h2><p class='err'>{esc(err)}</p><form method='post'><input name='email' type='email' placeholder='البريد الإلكتروني' required><input name='password' type='password' placeholder='كلمة المرور (8 أحرف فأكثر)' minlength='8' required><button>{t}</button></form>{other}</div>")

@app.get("/login")
def login_get(): return auth_form("login")

@app.post("/login")
def login_post(req: Request, email: str = Form(...), password: str = Form(...)):
    with db() as c: u = c.execute("SELECT * FROM users WHERE email=?", (email.lower().strip(),)).fetchone()
    if not u or not secrets.compare_digest(u["pw"], hash_pw(password, u["salt"])): return auth_form("login", "بيانات غير صحيحة")
    req.session["uid"] = u["id"]; return RedirectResponse("/", 303)

@app.get("/register")
def reg_get(): return auth_form("register")

@app.post("/register")
def reg_post(req: Request, email: str = Form(...), password: str = Form(...)):
    email = email.lower().strip()
    if len(password) < 8: return auth_form("register", "كلمة المرور قصيرة")
    salt = secrets.token_hex(16)
    with db() as c:
        first = c.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0
        try:
            cur = c.execute("INSERT INTO users(email,pw,salt,is_admin) VALUES(?,?,?,?)",
                            (email, hash_pw(password, salt), salt, int(first or email == ADMIN_EMAIL)))
        except sqlite3.IntegrityError: return auth_form("register", "البريد مستخدم")
    req.session["uid"] = cur.lastrowid; return RedirectResponse("/", 303)

@app.get("/logout")
def logout(req: Request): req.session.clear(); return RedirectResponse("/login", 303)

@app.get("/", response_class=HTMLResponse)
def home(req: Request):
    u = me(req)
    if not u: return RedirectResponse("/login", 303)
    with db() as c: jobs = c.execute("SELECT * FROM jobs WHERE user_id=? ORDER BY id DESC LIMIT 30", (u["id"],)).fetchall()
    rows = ""
    for j in jobs:
        if j["status"] == "done": st = f"<a class='btn' href='/download/{j['id']}'>تحميل</a>"
        elif j["status"] == "error": st = f"<span class='err'>{esc(j['error'] or 'خطأ')}</span>"
        else: st = f"<div class='bar'><i style='width:{j['progress']}%'></i></div>{j['progress']}%"
        rows += f"<tr><td>{esc(j['filename'])}</td><td>{st}</td></tr>"
    busy = any(j["status"] == "processing" for j in jobs)
    body = f"""<div class='card'><h2>ترجمة ملف (إنجليزي ← عربي)</h2>
    <form action='/upload' method='post' enctype='multipart/form-data'><input type='file' name='file' accept='{ACCEPT}' required>
    <button>ابدأ الترجمة</button></form><small>الصيغ: DOCX, PPTX, XLSX, PDF, HTML, CSV, JSON, SRT, VTT, TXT, MD وغيرها — حتى {MAX_MB}MB</small></div>
    <div class='card'><h3>ملفاتي</h3><table>{rows or '<tr><td>لا توجد ملفات بعد</td></tr>'}</table></div>"""
    return page(body, u, refresh=busy)

@app.post("/upload")
async def upload(req: Request, bg: BackgroundTasks, file: UploadFile = File(...)):
    u = me(req)
    if not u: return RedirectResponse("/login", 303)
    ext = Path(file.filename or "").suffix.lower()
    if ext not in H: raise HTTPException(400, "صيغة غير مدعومة")
    data = await file.read()
    if len(data) > MAX_MB * 1024 * 1024: raise HTTPException(413, "الملف كبير")
    path = DATA / "up" / f"{secrets.token_hex(8)}{ext}"; path.write_bytes(data)
    with db() as c:
        cur = c.execute("INSERT INTO jobs(user_id,filename,status,src,chars) VALUES(?,?,?,?,?)",
                        (u["id"], Path(file.filename).name, "processing", str(path), len(data)))
    bg.add_task(run_job, cur.lastrowid); return RedirectResponse("/", 303)

@app.get("/download/{jid}")
def download(req: Request, jid: int):
    u = me(req)
    if not u: return RedirectResponse("/login", 303)
    with db() as c: j = c.execute("SELECT * FROM jobs WHERE id=?", (jid,)).fetchone()
    if not j or j["status"] != "done" or (j["user_id"] != u["id"] and not u["is_admin"]): raise HTTPException(404)
    return FileResponse(j["out"], filename="AR_" + Path(j["filename"]).stem + Path(j["out"]).suffix)

# ---------- لوحة المطور ----------
def admin_only(req):
    u = me(req)
    if not u or not u["is_admin"]: raise HTTPException(403, "للمشرف فقط")
    return u

@app.get("/admin", response_class=HTMLResponse)
def admin(req: Request):
    u = admin_only(req)
    with db() as c:
        n = lambda q: c.execute(q).fetchone()[0]
        stats = [("المستخدمون", n("SELECT COUNT(*) FROM users")), ("كل الترجمات", n("SELECT COUNT(*) FROM jobs")),
                 ("ناجحة", n("SELECT COUNT(*) FROM jobs WHERE status='done'")), ("فاشلة", n("SELECT COUNT(*) FROM jobs WHERE status='error'"))]
        users = c.execute("SELECT u.id,u.email,u.is_admin,u.created,(SELECT COUNT(*) FROM jobs WHERE user_id=u.id) n FROM users u ORDER BY id DESC LIMIT 50").fetchall()
        jobs = c.execute("SELECT j.*,u.email FROM jobs j JOIN users u ON u.id=j.user_id ORDER BY j.id DESC LIMIT 30").fetchall()
        gl = c.execute("SELECT * FROM glossary ORDER BY id DESC").fetchall()
    s = "".join(f"<div class='card stat'><h2>{v}</h2>{k}</div>" for k, v in stats)
    ut = "".join(f"<tr><td>{esc(x['email'])}</td><td>{x['n']}</td><td>{x['created'][:10]}</td><td>{'مشرف' if x['is_admin'] else ''}</td></tr>" for x in users)
    jt = "".join(f"<tr><td>{esc(x['email'])}</td><td>{esc(x['filename'])}</td><td>{x['status']}</td><td>{esc(x['error'] or '')}</td></tr>" for x in jobs)
    gt = "".join(f"<tr><td>{esc(x['en'])}</td><td>{esc(x['ar'])}</td><td><form method='post' action='/admin/glossary/del/{x['id']}'><button>حذف</button></form></td></tr>" for x in gl)
    body = f"""<div class='grid'>{s}</div>
    <div class='card'><h3>قاموس المصطلحات (يُفرض في كل ترجمة)</h3><form method='post' action='/admin/glossary'>
    <input name='en' placeholder='English term' dir='ltr' required><input name='ar' placeholder='الترجمة المعتمدة' required><button>إضافة</button></form><table>{gt}</table></div>
    <div class='card'><h3>آخر العمليات</h3><table>{jt}</table></div>
    <div class='card'><h3>المستخدمون</h3><table><tr><th>البريد</th><th>الملفات</th><th>التسجيل</th><th></th></tr>{ut}</table></div>"""
    return page(body, u)

@app.post("/admin/glossary")
def gl_add(req: Request, en: str = Form(...), ar: str = Form(...)):
    admin_only(req)
    with db() as c: c.execute("INSERT INTO glossary(en,ar) VALUES(?,?)", (en.strip(), ar.strip()))
    return RedirectResponse("/admin", 303)

@app.post("/admin/glossary/del/{gid}")
def gl_del(req: Request, gid: int):
    admin_only(req)
    with db() as c: c.execute("DELETE FROM glossary WHERE id=?", (gid,))
    return RedirectResponse("/admin", 303)
