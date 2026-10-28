"""Secrétaire Issa — assistant personnel avec rappels automatiques par Gmail.
Lancement : python -m streamlit run CampusClair_XXL.py
Base SQLite locale séparée : campus_clair.db
Dépendances : streamlit, pandas, openpyxl (facultatif pour Excel), PyMuPDF (facultatif pour PDF).
"""
from pathlib import Path
from datetime import date, datetime, time, timedelta
import sqlite3, re, unicodedata, io, math, os, threading, time as time_module, smtplib, ssl
from email.message import EmailMessage
import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "campus_clair.db"
st.set_page_config(page_title="Secrétaire Issa", page_icon="🗂️", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
:root{--issa-navy:#10213c;--issa-blue:#2457a7;--issa-cyan:#3bb7c9;--issa-gold:#f0b84b;--issa-muted:#64748b}
.block-container{max-width:1550px;padding-top:1rem;padding-bottom:2rem}
[data-testid="stAppViewContainer"]{background:linear-gradient(180deg,#f4f8ff 0%,#f8fafc 38%,#f3f7fb 100%)}
[data-testid="stSidebar"]{background:#fff;border-right:1px solid #e1e8f2}
[data-testid="stSidebar"]{visibility:hidden;width:0;min-width:0}
.issa-brand{background:linear-gradient(115deg,#10213c 0%,#1c4b88 62%,#267d9d 100%);color:white;border-radius:24px;padding:24px 28px;margin:0 0 15px;box-shadow:0 12px 28px rgba(16,33,60,.16);position:relative;overflow:hidden}
.issa-brand:after{content:'✦';position:absolute;right:34px;top:-27px;font-size:130px;color:rgba(255,255,255,.08)}
.issa-brand h1{font-size:clamp(1.7rem,3vw,2.55rem);font-weight:850;letter-spacing:-.04em;margin:0;color:#fff}
.issa-brand p{margin:7px 0 0;color:#e0efff;font-size:.98rem;max-width:900px}
.issa-badge{display:inline-block;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.25);border-radius:999px;padding:5px 11px;font-size:.75rem;font-weight:800;letter-spacing:.09em;text-transform:uppercase;margin-bottom:9px}
.cc-hero{background:linear-gradient(120deg,#fff,#edf5ff);border:1px solid #dce7f7;border-radius:20px;padding:22px 25px;margin:12px 0 18px;box-shadow:0 4px 15px rgba(30,64,120,.04)}
.cc-eyebrow{font-size:.75rem;letter-spacing:.12em;text-transform:uppercase;color:#3868a8;font-weight:850}
.cc-title{font-size:clamp(1.45rem,2.3vw,2rem);font-weight:850;color:#152640;margin:.2rem 0}
.cc-sub{color:#60728c;margin:0}
.cc-card{background:#fff;border:1px solid #e0e8f3;border-radius:18px;padding:18px;height:100%;box-shadow:0 5px 18px rgba(25,45,80,.045)}
.cc-label{font-size:.88rem;color:#64748b}.cc-value{font-size:1.9rem;font-weight:850;color:#18243a;margin-top:7px}.cc-note{font-size:.8rem;color:#64748b;margin-top:4px}
.stButton>button{border-radius:11px;font-weight:700;min-height:2.55rem;border:1px solid #dbe5f2;transition:all .18s ease}
.stButton>button:hover{border-color:#3b73c4;transform:translateY(-1px);box-shadow:0 5px 14px rgba(36,87,167,.12)}
[data-testid="stDataFrame"]{border:1px solid #e4e8f0;border-radius:12px}
div[role="radiogroup"]{gap:.35rem;flex-wrap:wrap}
div[role="radiogroup"] label{background:#fff;border:1px solid #dce6f3;border-radius:12px;padding:8px 12px!important;font-weight:700;box-shadow:0 2px 7px rgba(16,33,60,.035)}
div[role="radiogroup"] label:has(input:checked){background:#eaf2ff;border-color:#7ca8e6;color:#164b93}
.issa-nav-title{font-size:.72rem;font-weight:850;text-transform:uppercase;letter-spacing:.12em;color:#52709a;margin:8px 0 5px}
.issa-footer{background:#10213c;color:#eaf2ff;border-radius:18px;padding:16px 20px;margin-top:24px}
.issa-footer strong{color:#fff}.issa-footer small{color:#bfd2eb}
.issa-roles{background:linear-gradient(110deg,#10213c,#1c4b88);border-radius:16px;padding:15px 17px;margin-top:16px;color:#fff;box-shadow:0 7px 20px rgba(16,33,60,.12)}
.issa-roles-title{font-size:.75rem;font-weight:850;letter-spacing:.1em;text-transform:uppercase;color:#cfe5ff;margin-bottom:10px}
.issa-role-pills{display:flex;gap:8px;flex-wrap:wrap}
.issa-role-pill{display:inline-flex;align-items:center;gap:6px;border:1px solid rgba(255,255,255,.24);background:rgba(255,255,255,.1);border-radius:999px;padding:7px 11px;font-size:.78rem;color:#f4f8ff}
@media(max-width:700px){.issa-brand{padding:20px;border-radius:18px}.cc-hero{padding:18px}.block-container{padding-left:1rem;padding-right:1rem}div[role="radiogroup"] label{padding:6px 8px!important;font-size:.8rem}}
</style>""", unsafe_allow_html=True)

# ---------- Données : tables créées automatiquement, base strictement locale ----------
def connect():
    c=sqlite3.connect(DB_PATH, timeout=20)
    c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init_db():
    with connect() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS subjects(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,code TEXT DEFAULT '',semester TEXT DEFAULT '',ue TEXT DEFAULT '',ue_code TEXT DEFAULT '',credits REAL DEFAULT 0,cm REAL DEFAULT 0,td REAL DEFAULT 0,tp REAL DEFAULT 0,tpe REAL DEFAULT 0,vht REAL DEFAULT 0,teacher TEXT DEFAULT '',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS events(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,event_date TEXT NOT NULL,start_time TEXT DEFAULT '',end_time TEXT DEFAULT '',kind TEXT DEFAULT 'Cours',subject_id INTEGER,location TEXT DEFAULT '',notes TEXT DEFAULT '',source TEXT DEFAULT 'Saisie manuelle',created_at TEXT DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE SET NULL);
        CREATE TABLE IF NOT EXISTS tasks(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,due_date TEXT DEFAULT '',priority TEXT DEFAULT 'Normale',status TEXT DEFAULT 'À faire',category TEXT DEFAULT 'Travail personnel',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,category TEXT DEFAULT 'Autre',url TEXT DEFAULT '',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS study_logs(id INTEGER PRIMARY KEY AUTOINCREMENT,subject_id INTEGER,study_date TEXT NOT NULL,hours REAL NOT NULL,method TEXT DEFAULT 'Révision active',confidence INTEGER DEFAULT 3,notes TEXT DEFAULT '',FOREIGN KEY(subject_id) REFERENCES subjects(id) ON DELETE SET NULL);
        CREATE TABLE IF NOT EXISTS goals(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,category TEXT DEFAULT 'Études',target_date TEXT DEFAULT '',progress INTEGER DEFAULT 0,notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS habits(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,frequency TEXT DEFAULT 'Quotidienne',notes TEXT DEFAULT '',active INTEGER DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS habit_logs(id INTEGER PRIMARY KEY AUTOINCREMENT,habit_id INTEGER NOT NULL,log_date TEXT NOT NULL,done INTEGER DEFAULT 1,UNIQUE(habit_id,log_date),FOREIGN KEY(habit_id) REFERENCES habits(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS budgets(id INTEGER PRIMARY KEY AUTOINCREMENT,entry_date TEXT NOT NULL,kind TEXT NOT NULL,category TEXT DEFAULT 'Autre',amount REAL NOT NULL,description TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS life_notes(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,category TEXT DEFAULT 'Personnel',content TEXT DEFAULT '',reminder_date TEXT DEFAULT '',reminder_time TEXT DEFAULT '08:00',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS contacts(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,role TEXT DEFAULT '',phone TEXT DEFAULT '',email TEXT DEFAULT '',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS career_skills(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,category TEXT DEFAULT 'Technique',level TEXT DEFAULT 'Débutant',target_level TEXT DEFAULT 'Intermédiaire',status TEXT DEFAULT 'À apprendre',resource_url TEXT DEFAULT '',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS career_training(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,provider TEXT DEFAULT '',kind TEXT DEFAULT 'Formation',start_date TEXT DEFAULT '',end_date TEXT DEFAULT '',status TEXT DEFAULT 'À envisager',url TEXT DEFAULT '',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS career_opportunities(id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,organization TEXT DEFAULT '',kind TEXT DEFAULT 'Stage',location TEXT DEFAULT '',deadline TEXT DEFAULT '',url TEXT DEFAULT '',status TEXT DEFAULT 'À examiner',contact TEXT DEFAULT '',notes TEXT DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS reminders(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            details TEXT DEFAULT '',
            remind_at TEXT NOT NULL,
            phone TEXT NOT NULL DEFAULT '',
            repeat_rule TEXT DEFAULT 'Aucun',
            enabled INTEGER DEFAULT 1,
            status TEXT DEFAULT 'En attente',
            last_sent_at TEXT DEFAULT '',
            last_error TEXT DEFAULT '',
            source_type TEXT DEFAULT '',
            source_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS app_settings(key TEXT PRIMARY KEY, value TEXT DEFAULT '');
        """)
init_db()
with connect() as _conn:
    _cols = {row['name'] for row in _conn.execute('PRAGMA table_info(reminders)').fetchall()}
    if 'source_type' not in _cols: _conn.execute("ALTER TABLE reminders ADD COLUMN source_type TEXT DEFAULT ''")
    if 'source_id' not in _cols: _conn.execute('ALTER TABLE reminders ADD COLUMN source_id INTEGER')
    _conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_reminders_source ON reminders(source_type, source_id) WHERE source_type <> '' AND source_id IS NOT NULL")
    _note_cols = {row['name'] for row in _conn.execute('PRAGMA table_info(life_notes)').fetchall()}
    if 'reminder_date' not in _note_cols: _conn.execute("ALTER TABLE life_notes ADD COLUMN reminder_date TEXT DEFAULT ''")
    if 'reminder_time' not in _note_cols: _conn.execute("ALTER TABLE life_notes ADD COLUMN reminder_time TEXT DEFAULT '08:00'")

DEFAULT_REMINDER_PHONE = "+221777473170"  # conservé pour compatibilité avec les anciennes bases

def get_config(name, default=""):
    """Lit un paramètre depuis Streamlit Secrets ou les variables d’environnement."""
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default
    return str(os.getenv(name, value) or default).strip()

def get_saved_setting(key, default=""):
    try:
        with connect() as c:
            row = c.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
            return str(row["value"] or default) if row else default
    except Exception:
        return default

def save_setting(key, value):
    with connect() as c:
        c.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value).strip()))

def email_configured():
    """Vérifie les paramètres Gmail SMTP et l'adresse de réception."""
    return all(get_config(k) for k in ("SMTP_USERNAME", "SMTP_PASSWORD")) and bool(get_recipient_email())

def get_recipient_email():
    return get_config("REMINDER_EMAIL") or get_saved_setting("reminder_email", "")

def send_email(subject, body, recipient=None):
    """Envoie un e-mail par Gmail SMTP, avec STARTTLS (587) puis SSL (465) en secours."""
    host = get_config("SMTP_HOST", "smtp.gmail.com") or "smtp.gmail.com"
    username = get_config("SMTP_USERNAME")
    password = get_config("SMTP_PASSWORD").replace(" ", "")
    sender_name = get_config("SMTP_SENDER_NAME", "Secrétaire Issa") or "Secrétaire Issa"
    recipient = (recipient or get_recipient_email()).strip()
    if not username or not password:
        raise RuntimeError("Configuration Gmail incomplète : ajoute SMTP_USERNAME et SMTP_PASSWORD dans les Secrets Streamlit. Utilise le mot de passe d'application Google, pas ton mot de passe Gmail habituel.")
    if not recipient or not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", recipient):
        raise RuntimeError("Adresse de réception absente ou invalide. Renseigne ton adresse dans l’onglet Rappels automatiques.")
    try:
        port = int(get_config("SMTP_PORT", "587") or "587")
    except ValueError:
        raise RuntimeError("SMTP_PORT doit être 587 ou 465.") from None
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = f"{sender_name} <{username}>"
    message["To"] = recipient
    message.set_content(body)
    context = ssl.create_default_context()
    errors = []
    # Essai selon le port configuré, puis bascule vers l'autre mode Gmail.
    modes = [("ssl", 465), ("starttls", 587)] if port == 465 else [("starttls", 587), ("ssl", 465)]
    for mode, candidate_port in modes:
        server = None
        try:
            if mode == "ssl":
                server = smtplib.SMTP_SSL(host, candidate_port, timeout=20, context=context)
            else:
                server = smtplib.SMTP(host, candidate_port, timeout=20)
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
            server.login(username, password)
            server.send_message(message)
            server.quit()
            return True
        except smtplib.SMTPAuthenticationError:
            if server:
                try: server.quit()
                except Exception: pass
            raise RuntimeError("Gmail a refusé l'authentification. Vérifie SMTP_USERNAME et utilise le mot de passe d'application Google à 16 caractères. Ne mets pas ton mot de passe Gmail habituel.") from None
        except Exception as exc:
            errors.append(f"port {candidate_port} ({mode}) : {type(exc).__name__}: {exc}")
            if server:
                try: server.quit()
                except Exception: pass
    detail = " ; ".join(errors)
    raise RuntimeError("Impossible de se connecter à Gmail SMTP après deux tentatives. Vérifie SMTP_HOST=smtp.gmail.com, le port 587/465, l'adresse Gmail et le mot de passe d'application. Détails : " + detail)

def process_due_reminders():
    """Envoie par e-mail chaque rappel arrivé à échéance, puis programme la récurrence."""
    if not email_configured():
        return
    now = datetime.now().replace(microsecond=0)
    now_text = now.strftime("%Y-%m-%d %H:%M:%S")
    try:
        due = read_df("SELECT * FROM reminders WHERE enabled=1 AND status='En attente' AND remind_at<=? ORDER BY remind_at LIMIT 20", (now_text,))
    except Exception:
        return
    for r in due.itertuples():
        try:
            with connect() as c:
                cur = c.execute("UPDATE reminders SET status='Envoi en cours' WHERE id=? AND status='En attente' AND enabled=1", (int(r.id),))
                if cur.rowcount != 1:
                    continue
            body = f"Bonjour,\n\nVoici ton rappel : {r.title}\n\n" + (f"Détails :\n{r.details}\n\n" if r.details else "") + f"Échéance prévue : {r.remind_at}.\n\n— Secrétaire Issa"
            send_email(f"🔔 Rappel Secrétaire Issa : {r.title}", body)
            sent_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if r.repeat_rule == "Quotidien":
                next_at = (datetime.strptime(r.remind_at, "%Y-%m-%d %H:%M:%S") + timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
                run("UPDATE reminders SET remind_at=?,status='En attente',last_sent_at=?,last_error='' WHERE id=?", (next_at, sent_at, int(r.id)))
            elif r.repeat_rule == "Hebdomadaire":
                next_at = (datetime.strptime(r.remind_at, "%Y-%m-%d %H:%M:%S") + timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
                run("UPDATE reminders SET remind_at=?,status='En attente',last_sent_at=?,last_error='' WHERE id=?", (next_at, sent_at, int(r.id)))
            else:
                run("UPDATE reminders SET status='Envoyé',last_sent_at=?,last_error='' WHERE id=?", (sent_at, int(r.id)))
        except Exception as exc:
            run("UPDATE reminders SET status='Erreur',last_error=? WHERE id=?", (str(exc)[:500], int(r.id)))

def _reminder_worker(stop_event):
    while not stop_event.is_set():
        try:
            process_due_reminders()
        except Exception:
            pass
        stop_event.wait(30)

@st.cache_resource
def start_reminder_worker():
    """Démarre un seul vérificateur par processus Streamlit."""
    stop_event = threading.Event()
    thread = threading.Thread(target=_reminder_worker, args=(stop_event,), daemon=True, name="campusclair-reminders")
    thread.start()
    return stop_event

def read_df(sql,params=()):
    with connect() as c: return pd.read_sql_query(sql,c,params=params)
def run(sql,params=()):
    with connect() as c: c.execute(sql,params)

def upsert_auto_reminder(source_type, source_id, title, details, remind_at, repeat_rule='Aucun'):
    """Crée/actualise un rappel lié à une entrée enregistrée, sans repousser l'échéance à chaque rechargement."""
    if not remind_at:
        return
    try:
        dt = datetime.strptime(str(remind_at), '%Y-%m-%d %H:%M:%S')
    except ValueError:
        try: dt = datetime.strptime(str(remind_at)[:16], '%Y-%m-%d %H:%M')
        except ValueError: return
    dt_text = dt.strftime('%Y-%m-%d %H:%M:%S')
    with connect() as c:
        existing = c.execute('SELECT id,status,title,details,remind_at,repeat_rule FROM reminders WHERE source_type=? AND source_id=?', (source_type, int(source_id))).fetchone()
        if existing:
            same_content = (existing['title'] == title and (existing['details'] or '') == (details or '') and existing['repeat_rule'] == repeat_rule)
            # Si le contenu n'a pas changé, conserver la date actuelle (notamment pour les rappels quotidiens).
            if same_content:
                if existing['status'] in ('Désactivé', 'Erreur'):
                    c.execute("UPDATE reminders SET enabled=1,status='En attente',last_error='' WHERE id=?", (int(existing['id']),))
                return
            # Une entrée modifiée dans sa rubrique met à jour le rappel associé.
            c.execute("UPDATE reminders SET title=?,details=?,remind_at=?,phone=?,repeat_rule=?,enabled=1,status='En attente',last_error='' WHERE id=?",
                      (title, details or '', dt_text, DEFAULT_REMINDER_PHONE, repeat_rule, int(existing['id'])))
        else:
            # Les éléments déjà en retard doivent tout de même être rappelés rapidement.
            if dt < datetime.now(): dt_text = (datetime.now() + timedelta(minutes=1)).replace(microsecond=0).strftime('%Y-%m-%d %H:%M:%S')
            c.execute("INSERT OR IGNORE INTO reminders(title,details,remind_at,phone,repeat_rule,enabled,status,source_type,source_id) VALUES(?,?,?,?,?,1,'En attente',?,?)",
                      (title, details or '', dt_text, DEFAULT_REMINDER_PHONE, repeat_rule, source_type, int(source_id)))

def sync_existing_entries_to_reminders():
    """Synchronise les rappels à partir des rubriques déjà remplies, y compris les anciennes entrées."""
    today = date.today()
    # Annule automatiquement les rappels dont l'entrée a été terminée, désactivée ou supprimée.
    try:
        with connect() as c:
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='task' AND source_id IN (SELECT id FROM tasks WHERE status IN ('Terminée','Terminé','Rendue','Rendu')) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='goal' AND source_id IN (SELECT id FROM goals WHERE progress >= 100) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='opportunity' AND source_id IN (SELECT id FROM career_opportunities WHERE status IN ('Refusée','Expirée','Retenue','Candidature envoyée')) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type IN ('training_début','training_fin') AND source_id IN (SELECT id FROM career_training WHERE status IN ('Terminée','Abandonnée')) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='habit' AND source_id IN (SELECT id FROM habits WHERE active=0) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='note' AND source_id IN (SELECT id FROM life_notes WHERE category<>'Rappel' OR reminder_date='') AND status IN ('En attente','Erreur')")
            # Les entrées supprimées n'ont plus de raison de déclencher un e-mail.
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='task' AND source_id NOT IN (SELECT id FROM tasks) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='goal' AND source_id NOT IN (SELECT id FROM goals) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='event' AND source_id NOT IN (SELECT id FROM events) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='opportunity' AND source_id NOT IN (SELECT id FROM career_opportunities) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type LIKE 'training_%' AND source_id NOT IN (SELECT id FROM career_training) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='habit' AND source_id NOT IN (SELECT id FROM habits) AND status IN ('En attente','Erreur')")
            c.execute("UPDATE reminders SET enabled=0,status='Désactivé' WHERE source_type='note' AND source_id NOT IN (SELECT id FROM life_notes) AND status IN ('En attente','Erreur')")
    except Exception:
        pass
    # Emploi du temps : une heure avant le début ; si l'évènement est imminent/passé aujourd'hui, rappel immédiat.
    for r in read_df("SELECT id,title,event_date,start_time,kind,location,notes FROM events WHERE event_date >= ?", (today.isoformat(),)).itertuples():
        try:
            event_dt = datetime.strptime(f"{r.event_date} {r.start_time or '08:00'}", '%Y-%m-%d %H:%M')
            remind = event_dt - timedelta(hours=1)
            if remind < datetime.now(): remind = datetime.now() + timedelta(minutes=1)
            upsert_auto_reminder('event', r.id, f"{r.kind} : {r.title}", f"Date : {r.event_date} à {r.start_time or '08:00'}. Lieu : {r.location or 'non précisé'}. {r.notes or ''}", remind.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception: pass
    # Toutes les tâches non terminées, y compris celles dont la date limite est dépassée.
    for r in read_df("SELECT id,title,due_date,priority,status,category,notes FROM tasks WHERE status NOT IN ('Terminée','Terminé','Rendue','Rendu')").itertuples():
        try:
            due = datetime.strptime(str(r.due_date), '%Y-%m-%d').date()
            remind = datetime.combine(due - timedelta(days=1), time(8,0))
            if due <= today and remind < datetime.now(): remind = datetime.now() + timedelta(minutes=1)
            upsert_auto_reminder('task', r.id, f"Échéance : {r.title}", f"Date limite : {due.isoformat()}. Priorité : {r.priority}. Domaine : {r.category}. Statut : {r.status}. {r.notes or ''}", remind.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception: pass
    # Objectifs en cours.
    for r in read_df("SELECT id,title,target_date,category,progress,notes FROM goals WHERE progress < 100").itertuples():
        try:
            if not r.target_date: continue
            due = datetime.strptime(str(r.target_date), '%Y-%m-%d').date()
            remind = datetime.combine(due - timedelta(days=1), time(8,0))
            if due <= today and remind < datetime.now(): remind = datetime.now() + timedelta(minutes=1)
            upsert_auto_reminder('goal', r.id, f"Objectif à suivre : {r.title}", f"Date cible : {due.isoformat()}. Domaine : {r.category}. Progression : {r.progress} %. {r.notes or ''}", remind.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception: pass
    # Stages, concours, bourses et autres opportunités avec échéance.
    for r in read_df("SELECT id,title,deadline,organization,kind,status,notes FROM career_opportunities WHERE deadline <> '' AND status NOT IN ('Refusée','Expirée','Retenue')").itertuples():
        try:
            due = datetime.strptime(str(r.deadline), '%Y-%m-%d').date()
            remind = datetime.combine(due - timedelta(days=1), time(8,0))
            if due <= today and remind < datetime.now(): remind = datetime.now() + timedelta(minutes=1)
            upsert_auto_reminder('opportunity', r.id, f"Date limite : {r.title}", f"Échéance : {due.isoformat()}. {r.kind} · {r.organization or 'Organisation non précisée'} · Statut : {r.status}. {r.notes or ''}", remind.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception: pass
    # Formations : rappel la veille du début et de la fin (deux entrées de rappel distinctes).
    for r in read_df("SELECT id,title,start_date,end_date,provider,status,notes FROM career_training WHERE status NOT IN ('Terminée','Abandonnée')").itertuples():
        try:
            for phase, raw_date in [('début', r.start_date), ('fin', r.end_date)]:
                if not raw_date: continue
                due = datetime.strptime(str(raw_date), '%Y-%m-%d').date()
                if due < today: continue
                remind = datetime.combine(due - timedelta(days=1), time(8,0))
                if due <= today and remind < datetime.now(): remind = datetime.now() + timedelta(minutes=1)
                upsert_auto_reminder(f'training_{phase}', r.id, f"Formation — {phase} : {r.title}", f"Date : {due.isoformat()}. Organisme : {r.provider or 'non précisé'}. État : {r.status}. {r.notes or ''}", remind.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception: pass
    # Habitudes actives : rappel quotidien stable, sans repousser la date à chaque rechargement.
    for r in read_df("SELECT id,title,frequency,notes FROM habits WHERE active=1").itertuples():
        remind = datetime.combine(today, time(8,0))
        if remind < datetime.now(): remind = datetime.now() + timedelta(minutes=1)
        upsert_auto_reminder('habit', r.id, f"Habitude du jour : {r.title}", f"Fréquence : {r.frequency}. {r.notes or ''}", remind.strftime('%Y-%m-%d %H:%M:%S'), 'Quotidien')
    # Notes classées « Rappel » : une date et une heure sont demandées dans la rubrique Notes personnelles.
    for r in read_df("SELECT id,title,content,reminder_date,reminder_time FROM life_notes WHERE category='Rappel' AND reminder_date <> ''").itertuples():
        try:
            dt = datetime.strptime(f"{r.reminder_date} {r.reminder_time or '08:00'}", '%Y-%m-%d %H:%M')
            upsert_auto_reminder('note', r.id, f"Note à ne pas oublier : {r.title}", r.content or '', dt.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception: pass

def norm(s):
    return re.sub(r'[^a-z0-9]+',' ',unicodedata.normalize('NFKD',str(s or '')).encode('ascii','ignore').decode().lower()).strip()
def header(kicker,title,desc):
    st.markdown(f'<div class="cc-hero"><div class="cc-eyebrow">{kicker}</div><div class="cc-title">{title}</div><p class="cc-sub">{desc}</p></div>',unsafe_allow_html=True)
def card(label,value,note=''):
    st.markdown(f'<div class="cc-card"><div class="cc-label">{label}</div><div class="cc-value">{value}</div><div class="cc-note">{note}</div></div>',unsafe_allow_html=True)
def export_csv(df,name,key):
    if not df.empty: st.download_button(f"Exporter {name} en CSV",df.to_csv(index=False).encode('utf-8-sig'),file_name=f"campusclair_{key}.csv",mime='text/csv',key=f"csv_{key}")
REMINDER_WORKER = start_reminder_worker()
# Synchronise les entrées puis traite immédiatement les échéances lors de chaque exécution
# du script Streamlit. Cela évite de dépendre uniquement du thread en arrière-plan :
# si Streamlit Cloud suspend le thread, le prochain chargement de l'application rattrape
# les rappels arrivés à échéance. Pour un envoi garanti même quand l'application dort,
# un planificateur externe et une base persistante restent nécessaires.
sync_existing_entries_to_reminders()
process_due_reminders()

def subject_choices():
    df=read_df('SELECT id,title,semester,ue FROM subjects ORDER BY title')
    opts={'— Non associée —':None}
    for r in df.itertuples(): opts[f"{r.title} · {r.semester} · {r.ue}"]=int(r.id)
    return df,opts

PAGES=["Vue d’ensemble","Maquettes & matières","Emploi du temps","Programmes de révision","Suivi & progression","Assistant de vie","Rappels automatiques","Compétences & Carrière","Paramètres & sauvegarde"]
PAGE_SHORT={
    "Vue d’ensemble":"🏠 Accueil XXL", "Maquettes & matières":"🎓 Académie", "Emploi du temps":"🗓️ Planning",
    "Programmes de révision":"🎯 Révisions", "Suivi & progression":"📈 Progression",
    "Assistant de vie":"👤 Vie personnelle", "Rappels automatiques":"📧 Rappels e-mail", "Compétences & Carrière":"🚀 Carrière",
    "Paramètres & sauvegarde":"⚙️ Paramètres"
}
PAGE_ROLE={
    "Vue d’ensemble":"Secrétaire chargé de la coordination générale",
    "Maquettes & matières":"Secrétaire chargé du suivi pédagogique et des maquettes",
    "Emploi du temps":"Secrétaire chargé de la planification et de l’emploi du temps",
    "Programmes de révision":"Secrétaire chargé de l’organisation des révisions",
    "Suivi & progression":"Secrétaire chargé du suivi des résultats et de la progression",
    "Assistant de vie":"Secrétaire chargé de l’organisation personnelle et des démarches",
    "Rappels automatiques":"Secrétaire chargé des rappels et échéances par e-mail",
    "Compétences & Carrière":"Secrétaire chargé du développement des compétences et de l’insertion professionnelle",
    "Paramètres & sauvegarde":"Secrétaire chargé des paramètres, des exports et de la sauvegarde des données"
}
st.markdown('''<div class="issa-brand"><span class="issa-badge">Organisation · Pédagogie · Assistance</span><h1>🗂️ Secrétaire Issa</h1><p>Votre espace de secrétariat intelligent : chaque espace correspond à un secrétaire chargé d’une mission précise.</p></div>''',unsafe_allow_html=True)
st.markdown('<div class="issa-nav-title">Menu principal · Choisissez le secrétaire spécialisé</div>',unsafe_allow_html=True)
if 'top_nav' not in st.session_state: st.session_state['top_nav']=PAGES[0]
st.radio('Navigation principale',PAGES,format_func=lambda x:PAGE_SHORT[x],horizontal=True,label_visibility='collapsed',key='top_nav')
page=st.session_state['top_nav']
st.markdown(f'<div class="cc-card" style="padding:12px 16px;margin:10px 0 16px"><div class="cc-eyebrow">Votre secrétaire spécialisé</div><div style="font-size:1.05rem;font-weight:800;color:#17345c;margin-top:4px">🧭 {PAGE_ROLE[page]}</div></div>',unsafe_allow_html=True)

# RAPPELS TRANSVERSAUX : visibles depuis TOUS les onglets, sans ressaisie.
try:
    _now = datetime.now().replace(microsecond=0)
    _horizon = (_now + timedelta(hours=24)).strftime('%Y-%m-%d %H:%M:%S')
    _now_text = _now.strftime('%Y-%m-%d %H:%M:%S')
    _alerts = read_df("SELECT title,details,remind_at,status FROM reminders WHERE enabled=1 AND status='En attente' AND remind_at<=? ORDER BY remind_at LIMIT 8", (_horizon,))
    if not _alerts.empty:
        _late = _alerts[_alerts['remind_at'].astype(str) <= _now_text]
        if not _late.empty:
            st.error(f"🔔 **RAPPEL À FAIRE : {len(_late)} tâche(s) ou échéance(s) sont arrivées à échéance.** Ouvre l’onglet « Rappels automatiques » pour les consulter.")
        _soon = _alerts[_alerts['remind_at'].astype(str) > _now_text]
        if not _soon.empty:
            st.warning(f"⏰ **À venir dans les prochaines 24 heures : {len(_soon)} rappel(s).** Consulte l’onglet « Rappels automatiques » pour voir les dates et les détails.")
        with st.expander('Voir les rappels à faire ou à venir', expanded=not _late.empty):
            for _item in _alerts.itertuples():
                st.markdown(f"**{_item.remind_at} — {_item.title}**")
                if _item.details:
                    st.caption(str(_item.details))
except Exception:
    # Ne pas bloquer l’application principale si une base ancienne est en cours de migration.
    pass

# ---------- Tableau de bord ----------
if page=="Vue d’ensemble":
    header("ESPACE CENTRAL","Ton tableau de bord 👋","Une vue simple de tes cours, de tes révisions et de tes objectifs. Les indicateurs reposent sur tes saisies réelles.")
    subjects=read_df('SELECT * FROM subjects'); events=read_df('SELECT * FROM events'); tasks=read_df('SELECT * FROM tasks'); logs=read_df('SELECT * FROM study_logs'); goals=read_df('SELECT * FROM goals')
    today=date.today().isoformat()
    upcoming=events[events.event_date>=today] if not events.empty else events
    open_tasks=tasks[~tasks.status.isin(['Terminée','Terminé','Rendue','Rendu'])] if not tasks.empty else tasks
    hours=float(logs.hours.sum()) if not logs.empty else 0
    cols=st.columns(4)
    with cols[0]: card('Matières enregistrées',len(subjects),'Maquette pédagogique')
    with cols[1]: card('Séances à venir',len(upcoming),'Cours, TD, TP, examens…')
    with cols[2]: card('Tâches ouvertes',len(open_tasks),'À faire ou en cours')
    with cols[3]: card('Heures étudiées',round(hours,1),'Journal de révision')
    # Résumé des rappels issus des autres rubriques, visible dès l'ouverture de l'application.
    sync_existing_entries_to_reminders()
    soon=read_df("SELECT title,details,remind_at,status FROM reminders WHERE enabled=1 AND status='En attente' ORDER BY remind_at LIMIT 6")
    st.subheader('🔔 Ce que tu dois faire bientôt')
    if soon.empty:
        st.info('Aucun rappel en attente. Les rappels apparaîtront automatiquement dès que tu enregistres une tâche, un cours, un objectif, une formation, une opportunité ou une note de type « Rappel ».')
    else:
        for rr in soon.itertuples():
            st.markdown(f"- **{rr.title}** — {rr.remind_at}")
            if rr.details: st.caption(rr.details)
    st.subheader('📅 Prochaines séances')
    if upcoming.empty: st.info('Aucune séance future. Saisis ton emploi du temps dans l’onglet correspondant.')
    else: st.dataframe(upcoming.sort_values(['event_date','start_time']).head(8)[['event_date','start_time','end_time','title','kind','location']].rename(columns={'event_date':'Date','start_time':'Début','end_time':'Fin','title':'Séance','kind':'Type','location':'Lieu'}),use_container_width=True,hide_index=True)
    a,b=st.columns(2)
    with a:
        st.subheader('🎯 Tâches prioritaires')
        if open_tasks.empty: st.success('Aucune tâche ouverte enregistrée.')
        else: st.dataframe(open_tasks.head(7)[['title','due_date','priority','status']].rename(columns={'title':'Tâche','due_date':'Échéance','priority':'Priorité','status':'Statut'}),use_container_width=True,hide_index=True)
    with b:
        st.subheader('🚀 Objectifs personnels')
        if goals.empty: st.info('Ajoute tes objectifs dans Assistant de vie → Objectifs.')
        else:
            for r in goals.sort_values('target_date').head(5).itertuples(): st.write(f"**{r.title}** — {r.progress}%" ); st.progress(int(r.progress)/100)
    st.caption('Conseil : commence par saisir ta maquette, puis ton emploi du temps. Le programme de révision sera calculé à partir de ces informations.')

# ---------- Maquettes ----------
elif page=="Maquettes & matières":
    header('PÉDAGOGIE','Maquettes & matières','Saisie manuelle, import en lot par CSV/Excel et contrôle des doublons. Les PDF peuvent être lus, mais leur conversion nécessite une vérification.')
    tab1,tab2,tab3=st.tabs(['➕ Saisie directe','📥 Import en lot','📚 Maquette enregistrée'])
    with tab1:
        with st.form('add_subject'):
            a,b,c=st.columns(3); title=a.text_input('Intitulé de la matière *'); code=b.text_input('Code matière'); semester=c.selectbox('Semestre',['À préciser']+[f'Semestre {i}' for i in range(1,7)])
            d,e,f=st.columns(3); ue=d.text_input('Unité d’enseignement (UE)'); ue_code=e.text_input('Code UE'); teacher=f.text_input('Enseignant')
            x=st.columns(6); cm=x[0].number_input('CM (h)',min_value=0.,step=1.); td=x[1].number_input('TD (h)',min_value=0.,step=1.); tp=x[2].number_input('TP (h)',min_value=0.,step=1.); tpe=x[3].number_input('TPE (h)',min_value=0.,step=1.); vht=x[4].number_input('VHT (h)',min_value=0.,step=1.); credits=x[5].number_input('Crédits',min_value=0.,step=1.)
            notes=st.text_area('Notes / compétences / coefficient')
            ok=st.form_submit_button('Enregistrer la matière',type='primary')
        if ok:
            if not title.strip(): st.error('L’intitulé est obligatoire.')
            else:
                run('INSERT INTO subjects(title,code,semester,ue,ue_code,credits,cm,td,tp,tpe,vht,teacher,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(title.strip(),code.strip(),semester,ue.strip(),ue_code.strip(),credits,cm,td,tp,tpe,vht,teacher.strip(),notes.strip()))
                st.success('Matière enregistrée.'); st.rerun()
    with tab2:
        st.markdown('**Pour une saisie massive**, prépare un CSV ou Excel avec les colonnes ci-dessous. Les noms de colonnes sont reconnus sans tenir compte des majuscules et accents.')
        template=pd.DataFrame([{'title':'Pédologie','code':'AGR101','semester':'Semestre 1','ue':'Sciences du sol','ue_code':'UE01','credits':4,'cm':20,'td':10,'tp':15,'tpe':10,'vht':55,'teacher':'Nom enseignant','notes':''}])
        st.download_button('Télécharger un modèle CSV',template.to_csv(index=False).encode('utf-8-sig'),'modele_maquette.csv','text/csv')
        upl=st.file_uploader('Importer les matières (CSV ou XLSX)',type=['csv','xlsx'],key='subjects_bulk')
        if upl:
            try:
                imp=pd.read_csv(upl) if upl.name.lower().endswith('.csv') else pd.read_excel(upl)
                st.dataframe(imp.head(30),use_container_width=True,hide_index=True)
                if st.button('Importer les lignes affichées',type='primary'):
                    aliases={'intitule':'title','matiere':'title','matiere title':'title','code matiere':'code','semestre':'semester','unite enseignement':'ue','code ue':'ue_code','credits ects':'credits','enseignant':'teacher','heures cm':'cm','heures td':'td','heures tp':'tp'}
                    imp=imp.rename(columns={col:aliases.get(norm(col),norm(col).replace(' ','_')) for col in imp.columns})
                    allowed=['title','code','semester','ue','ue_code','credits','cm','td','tp','tpe','vht','teacher','notes']
                    count=0
                    for _,r in imp.iterrows():
                        t=str(r.get('title','')).strip()
                        if not t or t.lower()=='nan': continue
                        exists=read_df('SELECT id FROM subjects WHERE lower(title)=lower(?) AND lower(semester)=lower(?) AND lower(ue)=lower(?)',(t,str(r.get('semester','À préciser')),str(r.get('ue',''))))
                        if not exists.empty: continue
                        vals=[]
                        for k in allowed:
                            v=r.get(k,'')
                            if pd.isna(v): v=''
                            if k in ['credits','cm','td','tp','tpe','vht']:
                                try: v=float(v or 0)
                                except: v=0.
                            vals.append(v)
                        run('INSERT INTO subjects(title,code,semester,ue,ue_code,credits,cm,td,tp,tpe,vht,teacher,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',tuple(vals))
                        count+=1
                    st.success(f'{count} matière(s) importée(s). Les doublons exacts ont été ignorés.'); st.rerun()
            except Exception as ex: st.error(f'Lecture impossible : {ex}')
        st.caption('Pour un PDF, utilise la saisie directe après vérification du texte extrait : les tableaux PDF ont souvent des colonnes difficiles à reconstruire automatiquement.')
        pdf=st.file_uploader('Lecture de contrôle facultative d’une maquette PDF',type=['pdf'],key='subject_pdf')
        if pdf and st.button('Extraire le texte du PDF',key='read_subject_pdf'):
            try:
                import fitz
                doc=fitz.open(stream=pdf.getvalue(),filetype='pdf'); st.text_area('Texte extrait à contrôler', '\n\n'.join(pg.get_text() for pg in doc),height=250)
            except ImportError: st.warning('Installe PyMuPDF : python -m pip install PyMuPDF')
            except Exception as ex: st.error(f'PDF illisible : {ex}')
    with tab3:
        df=read_df('SELECT * FROM subjects ORDER BY semester,ue,title')
        if df.empty: st.info('Aucune matière enregistrée.')
        else:
            st.dataframe(df.drop(columns=['created_at'],errors='ignore'),use_container_width=True,hide_index=True)
            export_csv(df,'la maquette','maquette')
            opts={f"#{r.id} — {r.title} ({r.semester})":int(r.id) for r in df.itertuples()}; chosen=st.selectbox('Fiche à supprimer (si nécessaire)',list(opts),key='del_subject_pick')
            if st.button('Supprimer la matière choisie',key='delete_subject'):
                run('DELETE FROM subjects WHERE id=?',(opts[chosen],)); st.success('Matière supprimée.'); st.rerun()
            dup=df.assign(n=df.title.map(norm),s=df.semester.map(norm),u=df.ue.map(norm)).groupby(['n','s','u']).filter(lambda g:len(g)>1)
            if not dup.empty: st.warning('Doublons potentiels détectés : vérifie-les avant suppression.'); st.dataframe(dup[['id','title','semester','ue']],hide_index=True)

# ---------- EDT ----------
elif page=="Emploi du temps":
    header('AGENDA','Emploi du temps','Saisis les séances une à une ou importe un tableau CSV/Excel. Toute ligne importée est affichée avant enregistrement.')
    subjects,subopts=subject_choices(); tab1,tab2,tab3=st.tabs(['➕ Ajouter une séance','📥 Importer un EDT','🗓️ Agenda enregistré'])
    with tab1:
        with st.form('add_event'):
            a,b,c=st.columns(3); title=a.text_input('Cours / événement *'); day=b.date_input('Date',date.today()); kind=c.selectbox('Type',['Cours','TD','TP','Examen','Devoir','Révision','Stage','Personnel','Autre'])
            d,e,f=st.columns(3); start=d.time_input('Début',time(8,0)); end=e.time_input('Fin',time(9,0)); location=f.text_input('Salle / lieu')
            sl=st.selectbox('Matière associée',list(subopts)); notes=st.text_input('Notes')
            ok=st.form_submit_button('Ajouter à l’agenda',type='primary')
        if ok:
            if not title.strip(): st.error('Le titre est obligatoire.')
            elif end<=start: st.error('La fin doit être après le début.')
            else:
                run('INSERT INTO events(title,event_date,start_time,end_time,kind,subject_id,location,notes) VALUES(?,?,?,?,?,?,?,?)',(title.strip(),day.isoformat(),start.strftime('%H:%M'),end.strftime('%H:%M'),kind,subopts[sl],location.strip(),notes.strip()))
                new_id = int(read_df('SELECT id FROM events WHERE title=? AND event_date=? ORDER BY id DESC LIMIT 1', (title.strip(),day.isoformat())).iloc[0]['id'])
                reminder_dt = datetime.combine(day, start) - timedelta(hours=1)
                if reminder_dt < datetime.now(): reminder_dt = datetime.combine(day, start)
                upsert_auto_reminder('event', new_id, f'{kind} : {title.strip()}', f'Lieu : {location.strip() or "non précisé"}. {notes.strip()}', reminder_dt.strftime('%Y-%m-%d %H:%M:%S'))
                st.success('Séance ajoutée. Un rappel est automatiquement créé dans Rappels.'); st.rerun()
    with tab2:
        st.markdown('Colonnes conseillées : **title, event_date, start_time, end_time, kind, subject, location, notes**. Date au format AAAA-MM-JJ et heures HH:MM.')
        template=pd.DataFrame([{'title':'Cours de pédologie','event_date':date.today().isoformat(),'start_time':'08:00','end_time':'10:00','kind':'Cours','subject':'Pédologie','location':'Salle 2','notes':''}])
        st.download_button('Télécharger le modèle EDT CSV',template.to_csv(index=False).encode('utf-8-sig'),'modele_edt.csv','text/csv')
        up=st.file_uploader('Importer un EDT CSV/XLSX',type=['csv','xlsx'],key='edt_bulk')
        if up:
            try:
                imp=pd.read_csv(up) if up.name.lower().endswith('.csv') else pd.read_excel(up)
                st.dataframe(imp,use_container_width=True,hide_index=True)
                if st.button('Enregistrer les séances importées',type='primary'):
                    alias={'intitule':'title','cours':'title','date':'event_date','jour':'event_date','debut':'start_time','heure debut':'start_time','fin':'end_time','heure fin':'end_time','type':'kind','matiere':'subject','salle':'location','lieu':'location','remarques':'notes'}
                    imp=imp.rename(columns={col:alias.get(norm(col),norm(col).replace(' ','_')) for col in imp.columns})
                    count=0; errors=[]
                    for ix,r in imp.iterrows():
                        try:
                            t=str(r.get('title','')).strip(); ds=pd.to_datetime(r.get('event_date'),dayfirst=True,errors='raise').date().isoformat()
                            stime=str(r.get('start_time','08:00')).strip()[:5]; etime=str(r.get('end_time','09:00')).strip()[:5]
                            datetime.strptime(stime,'%H:%M'); datetime.strptime(etime,'%H:%M')
                            if not t or etime<=stime: raise ValueError('titre manquant ou horaires invalides')
                            subname=norm(r.get('subject','')); sid=None
                            if subname:
                                found=subjects[subjects.title.map(norm)==subname]
                                if not found.empty: sid=int(found.iloc[0].id)
                            run('INSERT INTO events(title,event_date,start_time,end_time,kind,subject_id,location,notes,source) VALUES(?,?,?,?,?,?,?,?,?)',(t,ds,stime,etime,str(r.get('kind','Cours')),sid,str(r.get('location','')),str(r.get('notes','')),'Import CSV/Excel'))
                            count+=1
                        except Exception as ex: errors.append(f'Ligne {ix+2}: {ex}')
                    st.success(f'{count} séance(s) enregistrée(s).');
                    if errors: st.warning('Certaines lignes ont été ignorées :'); st.write('\n'.join(errors[:20]))
                    st.rerun()
            except Exception as ex: st.error(f'Fichier impossible à lire : {ex}')
        pdf=st.file_uploader('Lecture de contrôle facultative d’un EDT PDF',type=['pdf'],key='edt_pdf')
        if pdf and st.button('Lire le texte du PDF',key='read_edt_pdf'):
            try:
                import fitz
                doc=fitz.open(stream=pdf.getvalue(),filetype='pdf'); st.text_area('Texte extrait — à vérifier', '\n\n'.join(pg.get_text() for pg in doc),height=250)
                st.info('L’extraction PDF ne garantit pas la reconnaissance des cellules. Vérifie les jours, dates et horaires puis saisis/importes un tableau propre.')
            except ImportError: st.warning('Installe PyMuPDF : python -m pip install PyMuPDF')
            except Exception as ex: st.error(str(ex))
    with tab3:
        ev=read_df('SELECT e.*,s.title AS subject_title FROM events e LEFT JOIN subjects s ON s.id=e.subject_id ORDER BY event_date,start_time')
        if ev.empty: st.info('Aucune séance enregistrée.')
        else:
            d1,d2=st.columns(2); filt=d1.selectbox('Période',['Toutes','À venir','Passées']); query=d2.text_input('Rechercher')
            view=ev.copy()
            if filt=='À venir': view=view[view.event_date>=date.today().isoformat()]
            if filt=='Passées': view=view[view.event_date<date.today().isoformat()]
            if query.strip(): view=view[view.title.str.contains(query,case=False,na=False)]
            st.dataframe(view[['id','event_date','start_time','end_time','title','kind','subject_title','location','source']].rename(columns={'id':'ID','event_date':'Date','start_time':'Début','end_time':'Fin','title':'Événement','kind':'Type','subject_title':'Matière','location':'Lieu','source':'Origine'}),use_container_width=True,hide_index=True)
            export_csv(view,'l’emploi du temps','edt')
            eopts={f"#{r.id} — {r.event_date} {r.title}":int(r.id) for r in ev.itertuples()}; pick=st.selectbox('Séance à supprimer',list(eopts))
            if st.button('Supprimer la séance sélectionnée'): run('DELETE FROM events WHERE id=?',(eopts[pick],)); st.rerun()

# ---------- Génération du programme de révision ----------
elif page=="Programmes de révision":
    header('PLANIFICATION INTELLIGENTE','Programme de révision','Un planning indicatif construit à partir de ta maquette, de tes disponibilités et de la date cible. Tu peux ajuster les séances ensuite.')
    subjects=read_df('SELECT * FROM subjects ORDER BY title')
    if subjects.empty: st.warning('Commence par enregistrer tes matières dans « Maquettes & matières ».'); st.stop()
    with st.form('revision_generator'):
        a,b,c=st.columns(3); startday=a.date_input('Début du programme',date.today()); endday=b.date_input('Date cible / examen',date.today()+timedelta(days=21)); weekly=c.slider('Heures disponibles par semaine',2,40,10)
        d,e=st.columns(2); days=d.multiselect('Jours de révision',['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'],default=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi']); session=e.selectbox('Durée d’une séance',['30 minutes','45 minutes','1 heure','1 h 30','2 heures'],index=2)
        generate=st.form_submit_button('Générer mon programme',type='primary')
    if generate:
        if endday<startday: st.error('La date cible doit être après le début.')
        elif not days: st.error('Choisis au moins un jour de révision.')
        else:
            daymap={'Lundi':0,'Mardi':1,'Mercredi':2,'Jeudi':3,'Vendredi':4,'Samedi':5,'Dimanche':6}; allowed=[daymap[d] for d in days]
            duration={'30 minutes':.5,'45 minutes':.75,'1 heure':1,'1 h 30':1.5,'2 heures':2}[session]
            date_list=[]; cur=startday
            while cur<=endday:
                if cur.weekday() in allowed: date_list.append(cur)
                cur+=timedelta(days=1)
            if not date_list: st.error('Aucune date ne correspond à tes choix.')
            else:
                weights=[]
                for r in subjects.itertuples(): weights.append(max(float(r.vht or 0),float(r.credits or 0)*8,4))
                total=sum(weights); total_hours=min(weekly*max(1,(endday-startday).days/7),sum(weights))
                quota=[total_hours*w/total for w in weights]; remaining=list(quota); rows=[]; idx=0
                for day in date_list:
                    daily=min(weekly/len(allowed),max(duration,1))
                    slots=max(1,int(daily/duration))
                    for _ in range(slots):
                        available=[i for i,x in enumerate(remaining) if x>=min(duration,.25)]
                        if not available: break
                        i=max(available,key=lambda j:remaining[j]); hrs=min(duration,remaining[i]); r=subjects.iloc[i]
                        methods=['Relire le cours puis faire une fiche','Rappel actif sans regarder le cours','Exercices / problèmes corrigés','Questions-réponses et mini-test']
                        rows.append({'Date':day.isoformat(),'Jour':day.strftime('%A'),'Horaire conseillé':f'{8+(idx%8)*1:02d}:00','Matière':r.title,'Durée (h)':round(hrs,2),'Séance proposée':methods[idx%len(methods)],'État':'À faire'})
                        remaining[i]-=hrs; idx+=1
                plan=pd.DataFrame(rows)
                st.session_state['generated_plan']=plan
                st.success(f'Programme indicatif généré : {len(plan)} séances. Il ne connaît pas le contenu réel des cours ; ajuste les priorités selon les examens et difficultés.')
    if st.session_state.get('generated_plan') is not None:
        plan=st.session_state['generated_plan']
        edited=st.data_editor(plan,use_container_width=True,num_rows='dynamic',hide_index=True,key='plan_editor')
        st.download_button('Télécharger le programme CSV',edited.to_csv(index=False).encode('utf-8-sig'),'programme_revision.csv','text/csv')
        if st.button('Créer des tâches à partir de ce programme',type='primary'):
            for r in edited.to_dict('records'):
                run('INSERT INTO tasks(title,due_date,priority,status,category,notes) VALUES(?,?,?,?,?,?)',(f"Réviser : {r.get('Matière','')}",str(r.get('Date','')), 'Normale', 'À faire','Révision',f"{r.get('Durée (h)','')} h · {r.get('Séance proposée','')}"))
            st.success('Séances ajoutées à Tâches & révisions.'); st.rerun()
    st.divider(); st.subheader('Tâches de révision déjà planifiées')
    rt=read_df("SELECT * FROM tasks WHERE category='Révision' ORDER BY due_date")
    if rt.empty: st.caption('Aucune tâche de révision enregistrée pour le moment.')
    else: st.dataframe(rt[['title','due_date','priority','status','notes']],use_container_width=True,hide_index=True)

# ---------- Journal et progression ----------
elif page=="Suivi & progression":
    header('BAROMÈTRE PÉDAGOGIQUE','Suivi & progression','Enregistre chaque séance réellement étudiée, ton niveau de maîtrise et tes heures cumulées par matière.')
    subjects=read_df('SELECT id,title,semester,vht FROM subjects ORDER BY title')
    if subjects.empty: st.info('Ajoute d’abord les matières dans la maquette.')
    else:
        opts={f"{r.title} · {r.semester}":int(r.id) for r in subjects.itertuples()}
        with st.form('study_log_form'):
            a,b,c=st.columns(3); subject=a.selectbox('Matière',list(opts)); day=b.date_input('Date de travail',date.today()); hours=c.number_input('Durée étudiée (heures)',min_value=.25,max_value=16.,value=1.,step=.25)
            d,e=st.columns(2); method=d.selectbox('Méthode',['Rappel actif','Exercices','Fiche de synthèse','Lecture active','Révision espacée','Travail pratique','Préparation examen']); confidence=e.slider('Maîtrise après la séance (1 à 5)',1,5,3)
            notes=st.text_area('Ce qui est compris / reste à revoir')
            ok=st.form_submit_button('Enregistrer ma séance',type='primary')
        if ok:
            run('INSERT INTO study_logs(subject_id,study_date,hours,method,confidence,notes) VALUES(?,?,?,?,?,?)',(opts[subject],day.isoformat(),hours,method,confidence,notes.strip())); st.success('Séance enregistrée.'); st.rerun()
        logs=read_df('SELECT l.*,s.title AS subject FROM study_logs l LEFT JOIN subjects s ON s.id=l.subject_id ORDER BY study_date DESC,id DESC')
        summary=read_df('SELECT s.id,s.title,s.semester,s.vht,COALESCE(SUM(l.hours),0) AS studied,AVG(l.confidence) AS confidence FROM subjects s LEFT JOIN study_logs l ON l.subject_id=s.id GROUP BY s.id ORDER BY s.title')
        st.subheader('Progression par matière')
        for r in summary.itertuples():
            target=float(r.vht or 0); studied=float(r.studied or 0); progress=min(1,studied/target) if target>0 else 0
            st.markdown(f'**{r.title}** · {r.semester} — {studied:.1f} h étudiée(s) / {target:.1f} h de VHT')
            st.progress(progress, text=f'{progress*100:.0f}% de la VHT' if target>0 else 'VHT non renseignée')
            st.caption(f"Maîtrise moyenne déclarée : {float(r.confidence):.1f}/5" if pd.notna(r.confidence) else 'Aucune évaluation de maîtrise enregistrée.')
        st.subheader('Historique des séances')
        if logs.empty: st.info('Le journal est vide.')
        else:
            st.dataframe(logs[['study_date','subject','hours','method','confidence','notes']].rename(columns={'study_date':'Date','subject':'Matière','hours':'Heures','method':'Méthode','confidence':'Maîtrise /5','notes':'Notes'}),use_container_width=True,hide_index=True)
            export_csv(logs,'le journal d’étude','journal_etude')
            choices={f"#{r.id} · {r.study_date} · {r.subject}":int(r.id) for r in logs.itertuples()}; sel=st.selectbox('Entrée à supprimer',list(choices))
            if st.button('Supprimer cette entrée du journal'): run('DELETE FROM study_logs WHERE id=?',(choices[sel],)); st.rerun()

# ---------- Assistant de vie avec sous-onglets ----------
elif page=="Assistant de vie":
    header('ESPACE PERSONNEL XXL','Assistant de vie 🧭','Un espace central pour tes objectifs, habitudes, budget, tâches, routines, notes et contacts. Les informations restent dans la base SQLite locale.')
    tabs=st.tabs(['🎯 Objectifs','✅ Habitudes','💰 Budget','📝 Tâches & démarches','📓 Notes personnelles','👥 Contacts','🌿 Bien-être & routines'])
    with tabs[0]:
        with st.form('goal_form'):
            a,b,c=st.columns(3); title=a.text_input('Objectif *'); category=b.selectbox('Domaine',['Études','Carrière','Finances','Famille','Personnel','Projet','Santé / équilibre']); target=c.date_input('Date cible',date.today()+timedelta(days=30))
            progress=st.slider('Progression actuelle (%)',0,100,0); notes=st.text_area('Étapes / détails'); ok=st.form_submit_button('Ajouter l’objectif',type='primary')
        if ok:
            if title.strip():
                run('INSERT INTO goals(title,category,target_date,progress,notes) VALUES(?,?,?,?,?)',(title.strip(),category,target.isoformat(),progress,notes.strip()))
                new_id=int(read_df('SELECT id FROM goals WHERE title=? ORDER BY id DESC LIMIT 1',(title.strip(),)).iloc[0]['id'])
                remind=datetime.combine(max(date.today(),target-timedelta(days=1)),time(8,0))
                if remind < datetime.now(): remind=datetime.combine(target,time(8,0))
                upsert_auto_reminder('goal',new_id,f'Objectif à suivre : {title.strip()}',f'Domaine : {category}. Progression : {progress} %. {notes.strip()}',remind.strftime('%Y-%m-%d %H:%M:%S'))
                st.rerun()
            else: st.error('Donne un titre à ton objectif.')
        goals=read_df('SELECT * FROM goals ORDER BY target_date')
        if not goals.empty:
            for r in goals.itertuples():
                with st.container(border=True):
                    st.write(f"**{r.title}** · {r.category} · cible : {r.target_date or 'non définie'}")
                    new=st.slider(f'Progression — {r.title}',0,100,int(r.progress),key=f'goalprog_{r.id}')
                    if st.button('Mettre à jour',key=f'goalupdate_{r.id}'): run('UPDATE goals SET progress=? WHERE id=?',(new,r.id)); st.rerun()
                    if r.notes: st.caption(r.notes)
                    if st.button('Supprimer objectif',key=f'goaldel_{r.id}'): run('DELETE FROM goals WHERE id=?',(r.id,)); st.rerun()
    with tabs[1]:
        with st.form('habit_form'):
            title=st.text_input('Habitude à suivre *',placeholder='Ex. lire 20 minutes, réviser, ranger le bureau'); freq=st.selectbox('Fréquence',['Quotidienne','En semaine','Hebdomadaire']); notes=st.text_input('Pourquoi / rappel'); ok=st.form_submit_button('Créer une habitude')
        if ok:
            if title.strip():
                run('INSERT INTO habits(title,frequency,notes) VALUES(?,?,?)',(title.strip(),freq,notes.strip()))
                new_id=int(read_df('SELECT id FROM habits WHERE title=? ORDER BY id DESC LIMIT 1',(title.strip(),)).iloc[0]['id'])
                remind=datetime.combine(date.today(),time(8,0))
                if remind < datetime.now(): remind += timedelta(days=1)
                upsert_auto_reminder('habit',new_id,f'Habitude du jour : {title.strip()}',f'Fréquence : {freq}. {notes.strip()}',remind.strftime('%Y-%m-%d %H:%M:%S'),'Quotidien')
                st.rerun()
            else: st.error('Saisis le nom de l’habitude.')
        habits=read_df('SELECT * FROM habits WHERE active=1 ORDER BY title')
        if habits.empty: st.info('Aucune habitude suivie.')
        for r in habits.itertuples():
            a,b,c=st.columns([3,2,1]); a.write(f"**{r.title}** · {r.frequency}"); b.caption(r.notes)
            done=read_df('SELECT id FROM habit_logs WHERE habit_id=? AND log_date=? AND done=1',(r.id,date.today().isoformat()))
            if not done.empty: c.success('Fait aujourd’hui')
            elif c.button('Marquer fait',key=f'habitdone_{r.id}'):
                run('INSERT OR REPLACE INTO habit_logs(habit_id,log_date,done) VALUES(?,?,1)',(r.id,date.today().isoformat())); st.rerun()
            last7=read_df('SELECT COUNT(*) AS n FROM habit_logs WHERE habit_id=? AND done=1 AND log_date>=?',(r.id,(date.today()-timedelta(days=6)).isoformat())).iloc[0]['n']
            st.caption(f'Régularité sur les 7 derniers jours : {int(last7)}/7')
            if st.button('Supprimer habitude',key=f'habitdel_{r.id}'): run('DELETE FROM habits WHERE id=?',(r.id,)); st.rerun()
    with tabs[2]:
        with st.form('budget_form'):
            a,b,c=st.columns(3); day=a.date_input('Date',date.today()); kind=b.selectbox('Mouvement',['Dépense','Revenu']); amount=c.number_input('Montant (FCFA)',min_value=0.,step=500.)
            d,e=st.columns(2); category=d.selectbox('Catégorie',['Alimentation','Transport','Logement','Études','Téléphone / Internet','Santé','Loisirs','Salaire / aide','Épargne','Autre']); desc=e.text_input('Description')
            ok=st.form_submit_button('Enregistrer le mouvement')
        if ok:
            if amount>0: run('INSERT INTO budgets(entry_date,kind,category,amount,description) VALUES(?,?,?,?,?)',(day.isoformat(),kind,category,amount,desc.strip())); st.rerun()
            else: st.error('Le montant doit être supérieur à zéro.')
        budget=read_df('SELECT * FROM budgets ORDER BY entry_date DESC,id DESC')
        if not budget.empty:
            incomes=budget.loc[budget.kind=='Revenu','amount'].sum(); expenses=budget.loc[budget.kind=='Dépense','amount'].sum(); a,b,c=st.columns(3)
            with a: card('Revenus',f'{incomes:,.0f} FCFA')
            with b: card('Dépenses',f'{expenses:,.0f} FCFA')
            with c: card('Solde enregistré',f'{incomes-expenses:,.0f} FCFA')
            st.dataframe(budget,use_container_width=True,hide_index=True); export_csv(budget,'le budget','budget')
            opts={f"#{r.id} · {r.entry_date} · {r.kind} · {r.amount:g} FCFA":int(r.id) for r in budget.itertuples()}; pick=st.selectbox('Mouvement à supprimer',list(opts))
            if st.button('Supprimer le mouvement'): run('DELETE FROM budgets WHERE id=?',(opts[pick],)); st.rerun()
        else: st.info('Aucun mouvement saisi. Le solde est calculé uniquement à partir des mouvements enregistrés ici.')
    with tabs[3]:
        with st.form('life_task_form'):
            a,b,c=st.columns(3); title=a.text_input('Tâche / démarche *'); due=b.date_input('Date limite',date.today()+timedelta(days=1)); priority=c.selectbox('Priorité',['Basse','Normale','Haute','Urgente'])
            category=st.selectbox('Domaine',['Administratif','Études','Famille','Finances','Projet','Personnel','Santé','Travail']); notes=st.text_area('Détails / pièces à prévoir'); ok=st.form_submit_button('Ajouter à ma liste')
        if ok:
            if title.strip():
                run('INSERT INTO tasks(title,due_date,priority,status,category,notes) VALUES(?,?,?,?,?,?)',(title.strip(),due.isoformat(),priority,'À faire',category,notes.strip()))
                new_id=int(read_df('SELECT id FROM tasks WHERE title=? ORDER BY id DESC LIMIT 1',(title.strip(),)).iloc[0]['id'])
                remind=datetime.combine(max(date.today(), due-timedelta(days=1)), time(8,0))
                if remind < datetime.now(): remind=datetime.combine(due,time(8,0))
                upsert_auto_reminder('task',new_id,f'Échéance : {title.strip()}',f'Priorité : {priority}. Domaine : {category}. {notes.strip()}',remind.strftime('%Y-%m-%d %H:%M:%S'))
                st.rerun()
            else: st.error('Saisis le titre de la tâche.')
        tasks=read_df('SELECT * FROM tasks ORDER BY CASE priority WHEN "Urgente" THEN 1 WHEN "Haute" THEN 2 WHEN "Normale" THEN 3 ELSE 4 END,due_date')
        if not tasks.empty:
            st.dataframe(tasks[['id','title','due_date','priority','status','category','notes']],use_container_width=True,hide_index=True)
            opts={f"#{r.id} · {r.title}":int(r.id) for r in tasks.itertuples()}; sel=st.selectbox('Tâche à mettre à jour',list(opts)); tid=opts[sel]
            status=st.selectbox('Nouveau statut',['À faire','En cours','Terminée'],key='life_task_status')
            if st.button('Enregistrer le statut',key='life_task_save'): run('UPDATE tasks SET status=? WHERE id=?',(status,tid)); st.rerun()
            if st.button('Supprimer cette tâche',key='life_task_delete'): run('DELETE FROM tasks WHERE id=?',(tid,)); st.rerun()
        else: st.info('Ta liste est vide.')
    with tabs[4]:
        with st.form('note_form'):
            title=st.text_input('Titre de la note *'); category=st.selectbox('Catégorie',['Idée','Personnel','Études','Projet','Rappel','Journal','Autre']); content=st.text_area('Contenu',height=140)
            if category == 'Rappel':
                a,b=st.columns(2); reminder_date=a.date_input('Date du rappel',date.today()+timedelta(days=1)); reminder_time=b.time_input('Heure du rappel',time(8,0))
            else:
                reminder_date=None; reminder_time=None
            ok=st.form_submit_button('Enregistrer la note')
        if ok:
            if title.strip():
                rd=reminder_date.isoformat() if reminder_date else ''
                rt=reminder_time.strftime('%H:%M') if reminder_time else '08:00'
                run('INSERT INTO life_notes(title,category,content,reminder_date,reminder_time) VALUES(?,?,?,?,?)',(title.strip(),category,content.strip(),rd,rt))
                new_id=int(read_df('SELECT id FROM life_notes WHERE title=? ORDER BY id DESC LIMIT 1',(title.strip(),)).iloc[0]['id'])
                if category == 'Rappel' and rd:
                    dt=datetime.combine(reminder_date,reminder_time)
                    upsert_auto_reminder('note',new_id,f'Note à ne pas oublier : {title.strip()}',content.strip(),dt.strftime('%Y-%m-%d %H:%M:%S'))
                    st.success('Note enregistrée et rappel automatiquement programmé.')
                st.rerun()
            else: st.error('Le titre est obligatoire.')
        notes=read_df('SELECT * FROM life_notes ORDER BY created_at DESC')
        if notes.empty: st.info('Aucune note personnelle.')
        for r in notes.itertuples():
            with st.container(border=True):
                st.markdown(f'**{r.title}** · {r.category}'); st.write(r.content)
                if r.category == 'Rappel' and getattr(r, 'reminder_date', ''): st.caption(f"🔔 Rappel prévu le {r.reminder_date} à {getattr(r, 'reminder_time', '08:00')}")
                if st.button('Supprimer la note',key=f'note_del_{r.id}'): run('DELETE FROM life_notes WHERE id=?',(r.id,)); st.rerun()
    with tabs[5]:
        with st.form('contact_form'):
            a,b=st.columns(2); name=a.text_input('Nom / organisation *'); role=b.text_input('Rôle / relation'); c,d=st.columns(2); phone=c.text_input('Téléphone'); email=d.text_input('E-mail'); notes=st.text_area('Notes de contact'); ok=st.form_submit_button('Enregistrer le contact')
        if ok:
            if name.strip(): run('INSERT INTO contacts(name,role,phone,email,notes) VALUES(?,?,?,?,?)',(name.strip(),role.strip(),phone.strip(),email.strip(),notes.strip())); st.rerun()
            else: st.error('Le nom est obligatoire.')
        contacts=read_df('SELECT * FROM contacts ORDER BY name')
        if contacts.empty: st.info('Aucun contact enregistré.')
        else:
            st.dataframe(contacts.drop(columns=['created_at'],errors='ignore'),use_container_width=True,hide_index=True)
            opts={f"#{r.id} · {r.name}":int(r.id) for r in contacts.itertuples()}; pick=st.selectbox('Contact à supprimer',list(opts))
            if st.button('Supprimer le contact'): run('DELETE FROM contacts WHERE id=?',(opts[pick],)); st.rerun()
    with tabs[6]:
        st.markdown('### Petit bilan personnel')
        st.caption('Cet espace aide à organiser tes habitudes, sans remplacer un professionnel de santé. Choisis des routines réalistes et adapte-les à tes contraintes.')
        prompts={'Début de journée':'Quelle est ma priorité du jour ?','Études':'Quel sujet dois-je comprendre plutôt que simplement mémoriser ?','Équilibre':'Ai-je prévu une pause, du sommeil et un moment sans écran ?','Fin de journée':'Qu’ai-je accompli ? Quelle est la prochaine petite étape ?'}
        for label,q in prompts.items():
            with st.expander(label): st.write(q); st.text_area('Ma réponse',key='routine_'+norm(label))
        st.markdown('**Routine suggérée :** planifier 3 priorités, travailler en séances de 25–50 minutes, faire une courte pause, puis noter ce qui reste à revoir.')

# ---------- Rappels automatiques issus des rubriques enregistrées ----------
elif page=="Rappels automatiques":
    header('SECRÉTARIAT PERSONNEL','Rappels automatiques par e-mail 📧','Les rappels sont générés à partir de ce que tu enregistres déjà dans les onglets : agenda, tâches, objectifs, habitudes, formations et opportunités. Tu n’as pas à ressaisir les mêmes informations ici.')
    sync_existing_entries_to_reminders()
    current_email = get_recipient_email()
    with st.container(border=True):
        st.markdown('#### 📬 Adresse qui recevra les rappels')
        with st.form('recipient_email_form'):
            recipient_input = st.text_input('Ton adresse e-mail', value=current_email, placeholder='exemple@gmail.com', help='Cette adresse est enregistrée dans la base locale de l’application.')
            save_email = st.form_submit_button('Enregistrer mon adresse e-mail', type='primary')
        if save_email:
            if not re.match(r'^[^\s@]+@[^\s@]+\.[^\s@]+$', recipient_input.strip()):
                st.error('Entre une adresse e-mail valide.')
            else:
                save_setting('reminder_email', recipient_input.strip())
                st.success('Adresse de réception enregistrée.')
                st.rerun()
    if email_configured():
        st.success(f"Envoi e-mail configuré. Les rappels seront envoyés à **{get_recipient_email()}** lorsque l’application fonctionne.")
    else:
        st.warning('Les rappels sont préparés automatiquement, mais l’envoi nécessite les identifiants Gmail SMTP dans les Secrets Streamlit.')
    with st.expander('⚙️ Configuration Gmail gratuite (à faire une seule fois)', expanded=not bool(get_config('SMTP_USERNAME') and get_config('SMTP_PASSWORD'))):
        st.markdown('Utilise ton adresse Gmail et un **mot de passe d’application Google de 16 caractères**. Ne mets jamais ton mot de passe Gmail habituel et ne partage pas ce secret.')
        st.code('SMTP_HOST = \"smtp.gmail.com\"\nSMTP_PORT = \"587\"\nSMTP_USERNAME = \"tonadresse@gmail.com\"\nSMTP_PASSWORD = \"mot_de_passe_application_16_caracteres\"\nSMTP_SENDER_NAME = \"Secrétaire Issa\"', language='toml')
        st.caption('Dans Streamlit Cloud : Manage app → Settings → Secrets. L’adresse de réception se règle ci-dessus. Le code essaie STARTTLS sur le port 587 puis SSL sur le port 465 si nécessaire.')
    if st.button('🧪 Envoyer un e-mail de test', disabled=not (get_config('SMTP_USERNAME') and get_config('SMTP_PASSWORD') and get_recipient_email())):
        try:
            send_email('Test — Secrétaire Issa', 'Bonjour,\n\nCeci est un e-mail de test. Si tu le reçois, les rappels e-mail de Secrétaire Issa sont correctement configurés.')
            st.success(f"E-mail de test transmis à {get_recipient_email()}.")
        except Exception as exc:
            st.error(str(exc))
    reminders=read_df("SELECT id,title,details,remind_at,phone,repeat_rule,status,enabled,last_sent_at,last_error,source_type,source_id FROM reminders ORDER BY remind_at")
    if reminders.empty:
        st.info('Aucune échéance détectée. Ajoute une tâche avec une date limite, un cours dans l’agenda, un objectif daté ou une opportunité : le rappel sera créé automatiquement.')
    else:
        pending=reminders[(reminders.enabled==1)&(reminders.status=='En attente')]
        overdue=pending[pending.remind_at.astype(str)<=datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
        c1,c2,c3=st.columns(3)
        with c1: card('Rappels en attente',len(pending),'Échéances à venir ou à traiter')
        with c2: card('À traiter maintenant',len(overdue),'Date/heure du rappel atteinte')
        with c3: card('E-mails envoyés',int((reminders.status=='Envoyé').sum()),'Rappels ponctuels déjà transmis')
        display=reminders.rename(columns={'title':'Rappel automatique','details':'Détails / rubrique','remind_at':'Date et heure','repeat_rule':'Répétition','status':'État e-mail','enabled':'Actif','last_sent_at':'Dernier e-mail','last_error':'Erreur éventuelle','source_type':'Rubrique source','source_id':'N° entrée'})
        st.dataframe(display[['Rappel automatique','Détails / rubrique','Date et heure','Rubrique source','État e-mail','Actif','Dernier e-mail','Erreur éventuelle']],use_container_width=True,hide_index=True)
        choices={f"#{int(r.id)} — {r.title} — {r.remind_at}":int(r.id) for r in reminders.itertuples()}
        selected=st.selectbox('Gérer un rappel automatique',list(choices),key='auto_reminder_manage_pick')
        rid=choices[selected]; a,b=st.columns(2)
        if a.button('Désactiver / réactiver ce rappel',key='auto_reminder_toggle'):
            cur=read_df('SELECT enabled FROM reminders WHERE id=?',(rid,))
            if not cur.empty:
                enabled=0 if int(cur.iloc[0]['enabled']) else 1
                run("UPDATE reminders SET enabled=?,status=? WHERE id=?",(enabled,'En attente' if enabled else 'Désactivé',rid))
                st.rerun()
        if b.button('Envoyer ce rappel maintenant par e-mail',key='auto_reminder_send_now',disabled=not email_configured()):
            row=read_df('SELECT * FROM reminders WHERE id=?',(rid,))
            if not row.empty:
                r=row.iloc[0]
                try:
                    send_email(f"🔔 Rappel Secrétaire Issa : {r['title']}",f"Rappel : {r['title']}\n\n{r['details']}\n\nDate prévue : {r['remind_at']}\n— Secrétaire Issa")
                    run("UPDATE reminders SET last_sent_at=?,last_error='' WHERE id=?",(datetime.now().strftime('%Y-%m-%d %H:%M:%S'),rid))
                    st.success('E-mail transmis au serveur de messagerie.')
                except Exception as exc: st.error(str(exc))
    st.caption('Vérification des rappels au démarrage de chaque exécution de l’application et contrôle en arrière-plan toutes les 30 secondes tant que le processus Streamlit reste actif. Les rappels sont issus de l’agenda (1 h avant), des tâches et objectifs (veille à 08 h ou immédiatement si déjà en retard), des formations, des stages/concours/bourses, des notes « Rappel » et des habitudes actives. Gmail utilise le port 587 STARTTLS avec secours sur 465 SSL. Streamlit Cloud peut suspendre l’application lorsqu’elle est inactive : un envoi garanti même application endormie nécessite un planificateur externe et une base de données persistante.')

# ---------- Compétences & Carrière ----------
elif page=="Compétences & Carrière":
    header('DÉVELOPPEMENT PROFESSIONNEL','Compétences & Carrière 🚀','Prépare ton avenir professionnel : compétences, formations, certifications, stages et opportunités. Les informations sont conservées dans SQLite.')
    skills=read_df('SELECT * FROM career_skills')
    trainings=read_df('SELECT * FROM career_training')
    opportunities=read_df('SELECT * FROM career_opportunities')
    applied=opportunities[opportunities.status.isin(['Candidature envoyée','Entretien','Retenue','Refusée'])] if not opportunities.empty else opportunities
    metrics=st.columns(4)
    with metrics[0]: card('Compétences suivies',len(skills),'À acquérir ou à renforcer')
    with metrics[1]: card('Formations',len(trainings),'Formations et certifications')
    with metrics[2]: card('Opportunités',len(opportunities),'Stages, emplois, bourses…')
    with metrics[3]: card('Candidatures suivies',len(applied),'Démarches déjà engagées')
    tabs=st.tabs(['🧠 Compétences','🎓 Formations & certifications','🌍 Stages & opportunités','📊 Tableau de bord'])
    with tabs[0]:
        st.markdown('#### Ajouter une compétence')
        with st.form('career_skill_form'):
            a,b,c=st.columns(3); name=a.text_input('Compétence *'); category=b.selectbox('Domaine',['Agronomie','Sciences du sol','SIG / QGIS','Informatique','Bureautique','Communication','Gestion de projet','Langues','Recherche scientifique','Autre']); level=c.selectbox('Niveau actuel',['À évaluer','Débutant','Élémentaire','Intermédiaire','Avancé','Maîtrisé'])
            d,e=st.columns(2); target=d.selectbox('Niveau visé',['Élémentaire','Intermédiaire','Avancé','Maîtrisé']); status=e.selectbox('Statut',['À apprendre','En cours','À pratiquer','Acquise'])
            url=st.text_input('Lien utile (cours, tutoriel ou formation)'); notes=st.text_area('Plan d’apprentissage / observations')
            ok=st.form_submit_button('Enregistrer la compétence',type='primary')
        if ok:
            if not name.strip(): st.error('Le nom de la compétence est obligatoire.')
            elif url.strip() and not re.match(r'^https?://',url.strip()): st.error('Le lien doit commencer par http:// ou https://')
            else:
                run('INSERT INTO career_skills(name,category,level,target_level,status,resource_url,notes) VALUES(?,?,?,?,?,?,?)',(name.strip(),category,level,target,status,url.strip(),notes.strip()))
                st.success('Compétence enregistrée.'); st.rerun()
        skills=read_df('SELECT * FROM career_skills ORDER BY category,name')
        if skills.empty: st.info('Aucune compétence enregistrée. Commence par ajouter les logiciels, méthodes ou savoir-faire que tu veux maîtriser.')
        else:
            for r in skills.itertuples():
                with st.container(border=True):
                    a,b=st.columns([3,1]); a.markdown(f'**{r.name}**'); a.caption(f'{r.category} · Niveau : {r.level} · Objectif : {r.target_level}'); b.markdown(f'**{r.status}**')
                    if r.resource_url: st.markdown(f'[Ouvrir la ressource]({r.resource_url})')
                    if r.notes: st.write(r.notes)
                    with st.expander('Modifier ou supprimer cette compétence'):
                        with st.form(f'edit_skill_{r.id}'):
                            x,y,z=st.columns(3); ename=x.text_input('Nom',value=r.name); elevel=y.selectbox('Niveau actuel',['À évaluer','Débutant','Élémentaire','Intermédiaire','Avancé','Maîtrisé'],index=['À évaluer','Débutant','Élémentaire','Intermédiaire','Avancé','Maîtrisé'].index(r.level) if r.level in ['À évaluer','Débutant','Élémentaire','Intermédiaire','Avancé','Maîtrisé'] else 0); estatus=z.selectbox('Statut',['À apprendre','En cours','À pratiquer','Acquise'],index=['À apprendre','En cours','À pratiquer','Acquise'].index(r.status) if r.status in ['À apprendre','En cours','À pratiquer','Acquise'] else 0)
                            enotes=st.text_area('Notes',value=r.notes or ''); save=st.form_submit_button('Enregistrer les modifications')
                        if save:
                            run('UPDATE career_skills SET name=?,level=?,status=?,notes=? WHERE id=?',(ename.strip(),elevel,estatus,enotes.strip(),int(r.id))); st.rerun()
                        if st.button('Supprimer la compétence',key=f'del_skill_{r.id}'):
                            run('DELETE FROM career_skills WHERE id=?',(int(r.id),)); st.rerun()
            export_csv(skills,'les compétences','competences_carriere')
    with tabs[1]:
        with st.form('career_training_form'):
            title=st.text_input('Nom de la formation ou certification *'); provider=st.text_input('Organisme / plateforme'); kind=st.selectbox('Type',['Formation','Certification','Cours en ligne','Atelier','Conférence','Autre'])
            a,b=st.columns(2); start=a.date_input('Date de début (approximative)',value=date.today()); end=b.date_input('Date de fin prévue',value=date.today()+timedelta(days=30))
            status=st.selectbox('État',['À envisager','Inscription à faire','En cours','Terminée','Abandonnée']); url=st.text_input('Lien officiel'); notes=st.text_area('Notes / coût / prérequis')
            ok=st.form_submit_button('Enregistrer la formation')
        if ok:
            if not title.strip(): st.error('Le titre est obligatoire.')
            elif url.strip() and not re.match(r'^https?://',url.strip()): st.error('Le lien doit commencer par http:// ou https://')
            elif end<start: st.error('La date de fin ne peut pas précéder la date de début.')
            else:
                run('INSERT INTO career_training(title,provider,kind,start_date,end_date,status,url,notes) VALUES(?,?,?,?,?,?,?,?)',(title.strip(),provider.strip(),kind,start.isoformat(),end.isoformat(),status,url.strip(),notes.strip())); new_id=int(read_df('SELECT id FROM career_training WHERE title=? ORDER BY id DESC LIMIT 1',(title.strip(),)).iloc[0]['id']); sync_existing_entries_to_reminders(); st.success('Formation enregistrée : ses rappels de début et de fin sont programmés.'); st.rerun()
        trainings=read_df('SELECT * FROM career_training ORDER BY start_date DESC,id DESC')
        if trainings.empty: st.info('Aucune formation ou certification enregistrée.')
        else:
            st.dataframe(trainings[['title','provider','kind','start_date','end_date','status','url','notes']].rename(columns={'title':'Formation','provider':'Organisme','kind':'Type','start_date':'Début','end_date':'Fin','status':'État','url':'Lien','notes':'Notes'}),use_container_width=True,hide_index=True)
            opts={f"#{r.id} · {r.title}":int(r.id) for r in trainings.itertuples()}; pick=st.selectbox('Formation à supprimer',list(opts),key='career_training_delete_pick')
            if st.button('Supprimer la formation sélectionnée',key='career_training_delete'): run('DELETE FROM career_training WHERE id=?',(opts[pick],)); st.rerun()
            export_csv(trainings,'les formations','formations_carriere')
    with tabs[2]:
        with st.form('career_opportunity_form'):
            title=st.text_input('Intitulé de l’opportunité *'); organization=st.text_input('Organisation / entreprise'); kind=st.selectbox('Type',['Stage','Emploi','Bourse','Formation','Concours','Bénévolat','Projet','Autre']); location=st.text_input('Lieu / modalité')
            a,b=st.columns(2); deadline=a.date_input('Date limite (si connue)',value=date.today()+timedelta(days=30)); status=b.selectbox('Statut',['À examiner','À préparer','Candidature envoyée','Entretien','Retenue','Refusée','Expirée'])
            url=st.text_input('Lien de l’annonce'); contact=st.text_input('Contact / référence'); notes=st.text_area('Pièces à fournir / prochaines étapes')
            ok=st.form_submit_button('Enregistrer l’opportunité',type='primary')
        if ok:
            if not title.strip(): st.error('L’intitulé est obligatoire.')
            elif url.strip() and not re.match(r'^https?://',url.strip()): st.error('Le lien doit commencer par http:// ou https://')
            else:
                run('INSERT INTO career_opportunities(title,organization,kind,location,deadline,url,status,contact,notes) VALUES(?,?,?,?,?,?,?,?,?)',(title.strip(),organization.strip(),kind,location.strip(),deadline.isoformat(),url.strip(),status,contact.strip(),notes.strip())); sync_existing_entries_to_reminders(); st.success('Opportunité enregistrée : rappel de date limite programmé.'); st.rerun()
        opportunities=read_df('SELECT * FROM career_opportunities ORDER BY deadline,id DESC')
        if opportunities.empty: st.info('Aucune opportunité enregistrée. Ajoute ici les stages, emplois, bourses ou formations qui t’intéressent.')
        else:
            for r in opportunities.itertuples():
                with st.container(border=True):
                    st.markdown(f'**{r.title}** · {r.kind} · {r.organization or "Organisation non précisée"}')
                    st.caption(f'Lieu : {r.location or "Non précisé"} · Échéance : {r.deadline or "Non précisée"} · Statut : {r.status}')
                    if r.url: st.markdown(f'[Consulter l’annonce]({r.url})')
                    if r.contact: st.write(f'Contact : {r.contact}')
                    if r.notes: st.write(r.notes)
                    a,b=st.columns([3,1]); newstatus=a.selectbox('Mettre à jour le statut',['À examiner','À préparer','Candidature envoyée','Entretien','Retenue','Refusée','Expirée'],index=['À examiner','À préparer','Candidature envoyée','Entretien','Retenue','Refusée','Expirée'].index(r.status) if r.status in ['À examiner','À préparer','Candidature envoyée','Entretien','Retenue','Refusée','Expirée'] else 0,key=f'opp_status_{r.id}')
                    if b.button('Enregistrer',key=f'opp_save_{r.id}'):
                        run('UPDATE career_opportunities SET status=? WHERE id=?',(newstatus,int(r.id))); st.rerun()
                    if st.button('Supprimer',key=f'del_opp_{r.id}'):
                        run('DELETE FROM career_opportunities WHERE id=?',(int(r.id),)); st.rerun()
            export_csv(opportunities,'les opportunités','opportunites_carriere')
    with tabs[3]:
        skills=read_df('SELECT * FROM career_skills'); trainings=read_df('SELECT * FROM career_training'); opportunities=read_df('SELECT * FROM career_opportunities')
        done_skills=int((skills.status=='Acquise').sum()) if not skills.empty else 0
        in_progress=int((trainings.status=='En cours').sum()) if not trainings.empty else 0
        sent=int(opportunities.status.isin(['Candidature envoyée','Entretien','Retenue']).sum()) if not opportunities.empty else 0
        a,b,c=st.columns(3)
        with a: card('Compétences acquises',done_skills,f'sur {len(skills)} suivie(s)')
        with b: card('Formations en cours',in_progress,f'sur {len(trainings)} enregistrée(s)')
        with c: card('Candidatures actives',sent,'Envoyées, entretien ou retenues')
        st.markdown('#### Mes prochaines échéances')
        if not opportunities.empty:
            upcoming=opportunities[opportunities.deadline.astype(str)>=date.today().isoformat()].sort_values('deadline').head(8)
            if upcoming.empty: st.info('Aucune échéance future enregistrée.')
            else: st.dataframe(upcoming[['deadline','title','organization','kind','status']].rename(columns={'deadline':'Date limite','title':'Opportunité','organization':'Organisation','kind':'Type','status':'Statut'}),use_container_width=True,hide_index=True)
        else: st.info('Les échéances apparaîtront ici dès que tu enregistres des opportunités.')
        st.markdown('#### Conseils de départ')
        st.markdown('- Choisis 3 compétences prioritaires à développer ce mois-ci.\n- Ajoute les formations réellement commencées ou envisagées.\n- Enregistre les liens officiels des stages et opportunités pour retrouver facilement les annonces.')

# ---------- Sauvegarde ----------
elif page=="Paramètres & sauvegarde":
    header('PARAMÈTRES & DONNÉES','Paramètres & sauvegarde','Télécharge tes données pour garder une copie. Fais régulièrement une sauvegarde du fichier SQLite.')
    st.warning(f'Base active : {DB_PATH}. Cette application utilise une base distincte ; ne la remplace pas par la base d’une autre application.')
    tables={'Matières':'subjects','Emploi du temps':'events','Tâches':'tasks','Ressources archivées (ancien espace)':'documents','Journal d’étude':'study_logs','Objectifs':'goals','Habitudes':'habits','Historique habitudes':'habit_logs','Budget':'budgets','Notes personnelles':'life_notes','Contacts':'contacts','Compétences':'career_skills','Formations':'career_training','Opportunités professionnelles':'career_opportunities','Rappels e-mail':'reminders'}
    for label,table in tables.items():
        df=read_df(f'SELECT * FROM {table}')
        with st.expander(f'{label} — {len(df)} enregistrement(s)'):
            export_csv(df,label,table)
    if st.button('Préparer une sauvegarde complète SQLite',type='primary'):
        st.session_state['db_backup']=DB_PATH.read_bytes()
    if st.session_state.get('db_backup'):
        st.download_button('Télécharger la base complète (.db)',st.session_state['db_backup'],file_name=f'campusclair_sauvegarde_{date.today().isoformat()}.db',mime='application/octet-stream')
    st.markdown('### Dépendances facultatives')
    st.code('streamlit\npandas\nopenpyxl\nPyMuPDF',language='text')
    st.caption('openpyxl permet les imports Excel. PyMuPDF permet de lire le texte des PDF. L’application reste utilisable sans ces deux dépendances si tu utilises CSV et la saisie manuelle.')

st.markdown('<div class="issa-footer"><strong>🗂️ Secrétaire Issa</strong><br><small>Un espace personnel de secrétariat, de suivi et d’organisation · Données stockées localement dans SQLite.</small></div>',unsafe_allow_html=True)
role_icons={
    "Vue d’ensemble":"🏠", "Maquettes & matières":"🎓", "Emploi du temps":"🗓️",
    "Programmes de révision":"🎯", "Suivi & progression":"📈", "Assistant de vie":"👤", "Rappels automatiques":"🔔",
    "Compétences & Carrière":"🚀", "Paramètres & sauvegarde":"⚙️"
}
roles_html=''.join(f'<span class="issa-role-pill">{role_icons[p]} {PAGE_ROLE[p]}</span>' for p in PAGES)
st.markdown(f'<div class="issa-roles"><div class="issa-roles-title">Les rôles de vos secrétaires spécialisés</div><div class="issa-role-pills">{roles_html}</div></div>',unsafe_allow_html=True)
st.caption('Navigation disponible dans le menu supérieur · Base locale : '+DB_PATH.name)
