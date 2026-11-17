"""
Grade VIII-A Mathematics — Learning Improvement Dashboard (Streamlit)

Phase 1  Performance & Personalized Learning  (built from the Excel result analysis)
Phase 2  Continuous Learning Improvement      (weekly assessments, interventions,
                                               concept mastery, before/after)

KEY OBJECTIVE
  "Every student has an identified learning need, an assigned improvement
   activity, and measurable evidence of progress."

Run:   streamlit run app.py
Data:  the Excel result-analysis workbook sitting in the same folder as this file.
       Phase-2 records are saved to ./data/*.csv (created automatically).
"""
import io
import re
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Grade VIII-A Maths · Learning Improvement Dashboard",
    page_icon="📐",
    layout="wide",
)

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
DATA_DIR.mkdir(exist_ok=True)
FILES = {
    "weekly": DATA_DIR / "weekly_assessments.csv",
    "interventions": DATA_DIR / "interventions.csv",
    "diagnostics": DATA_DIR / "diagnostics.csv",
}
SCHEMA = {
    "weekly": ["roll", "week", "date", "assessment", "concept", "score", "max_score"],
    "interventions": ["roll", "activity", "addon", "support_format", "owner", "start_date",
                      "target_date", "sessions_planned", "sessions_done", "status",
                      "target_pct", "notes"],
    "diagnostics": ["roll", "verification_status", "verified_gap", "verified_by",
                    "verified_date", "notes"],
}

# Edit this list to match your syllabus plan — it drives the concept-mastery view.
CONCEPTS = [
    "Rational Numbers", "Linear Equations", "Quadrilaterals", "Data Handling",
    "Squares & Square Roots", "Cubes & Cube Roots", "Comparing Quantities",
    "Algebraic Expressions & Identities", "Mensuration", "Exponents & Powers",
    "Proportions", "Factorisation",
]
ASSESSMENT_TYPES = ["Weekly Quiz", "Concept Check", "Re-test (Assessment-II)"]
STATUSES = ["Assigned", "In Progress", "Completed", "Paused"]

# Band definitions (lowest → highest): short name, colour, lower %, upper % (excl.)
BANDS = [
    ("Critical", "#991b1b", 0, 40),
    ("Conceptual Gaps", "#dc2626", 40, 50),
    ("Foundational Cracks", "#ea580c", 50, 60),
    ("Developing", "#d97706", 60, 70),
    ("Competent", "#ca8a04", 70, 80),
    ("Proficient", "#65a30d", 80, 90),
    ("Near-Mastery", "#16a34a", 90, 100),
    ("Mastery", "#0f766e", 100, 101),
]
BAND_NAMES = [b[0] for b in BANDS]
BAND_COLOR = {b[0]: b[1] for b in BANDS}
BAND_COLOR["Incomplete"] = "#6b7280"
BAND_ORDER_ALL = BAND_NAMES[::-1] + ["Incomplete"]  # best → worst → incomplete

# Recommended improvement activity per band (follows the workbook's Learning Path Guide)
ACT = {
    "Mastery": dict(
        focus="Challenge with multi-chapter, Olympiad-style problems; peer-explainer role",
        activity="Cross-chapter Olympiad challenge set + author a problem with model solution",
        fmt="Peer-explainer / self-paced", freq="1 challenge set per week", sessions=4, weeks=4, gain=0,
        metric="Challenge-set score ≥ 85% and one authored problem per fortnight"),
    "Near-Mastery": dict(
        focus="Close the last 1–7 marks; stretch with harder application problems",
        activity="Error-classification of lost marks + harder application set",
        fmt="Self-paced + teacher check", freq="1 timed set per week", sessions=4, weeks=4, gain=3,
        metric="Lost marks classified (concept / calculation / presentation / time); timed set within 2 marks of full"),
    "Proficient": dict(
        focus="Fix 1–2 weak sub-topics; timed mixed practice",
        activity="Targeted sub-topic drills (5–6 problems + 1 'trap' question)",
        fmt="Small group", freq="2 sessions per week", sessions=8, weeks=4, gain=5,
        metric="Trap question answered correctly; weekly 15–20 min mixed set ≥ 90%"),
    "Competent": dict(
        focus="Rebuild the reasoning behind the steps; weekly mixed practice with same-week feedback",
        activity="'Why-first' re-teach of 2 weakest topics + reworded-question practice",
        fmt="Small group", freq="1 mixed set (10–12 Qs) per week", sessions=5, weeks=5, gain=8,
        metric="Student explains each step aloud; reworded-question accuracy ≥ 80%"),
    "Developing": dict(
        focus="Re-teach the 2–3 weakest chapters in small steps; short daily practice",
        activity="Small-step re-teach of weakest chapters + 10-min daily practice with tick-chart",
        fmt="Small group + peer pair", freq="Daily 10 min · 1 check-in per week", sessions=12, weeks=6, gain=10,
        metric="Daily-practice tick-chart ≥ 80% complete; weekly check ≥ 70%"),
    "Foundational Cracks": dict(
        focus="Test prerequisite skills; rebuild fluency with short daily practice",
        activity="Prerequisite diagnostic + 15-min daily fluency practice",
        fmt="Small group / one-to-one", freq="Daily 15 min · weekly small group", sessions=12, weeks=6, gain=12,
        metric="Prerequisite checklist items cleared every 2 weeks; weekly quiz ≥ 60%"),
    "Conceptual Gaps": dict(
        focus="Private conversation first, then twice-weekly small-group remediation from the true working level",
        activity="Below-grade restart: twice-weekly small-group remediation",
        fmt="Small group (twice weekly)", freq="2 sessions per week", sessions=12, weeks=6, gain=15,
        metric="Skill checklist reviewed every 2 weeks; weekly short quiz ≥ 55%"),
    "Critical": dict(
        focus="Immediate diagnostic, counsellor and parent involvement, below-grade restart",
        activity="One-to-one diagnostic + concrete/visual rebuild with counsellor & parent involvement",
        fmt="One-to-one", freq="Weekly one-to-one · daily short tasks", sessions=8, weeks=8, gain=20,
        metric="Working level identified in week 1; weekly success-task completion; upward quiz trend"),
    "Incomplete": dict(
        focus="Complete the data — confirm the reason for EX and schedule a make-up cycle test",
        activity="Make-up cycle test + core-skills diagnostic",
        fmt="One-to-one", freq="Weekly check-in", sessions=2, weeks=2, gain=0,
        metric="Make-up cycle test completed and band assigned"),
}

STATUS_COLOR = {
    "Target met": "#0f766e", "Improving": "#3f8f6b",
    "Flat — adjust activity": "#b7791f", "Declining — review": "#b42318",
    "Evidence building (1 week)": "#356b82", "No evidence yet": "#c05621",
    "No activity assigned": "#991b1b",
}

# ──────────────────────────────────────────────────────────────────────────────
# STYLE + SMALL UI HELPERS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
.block-container{padding-top:1.6rem;max-width:1400px}
body{font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
h1,h2,h3,h4,h5,h6{color:#17324d}
.hero{background:#17324d;color:#fff;padding:20px 26px;border-radius:10px;margin-bottom:18px}
.hero h1{margin:0;font-size:1.65rem;color:#fff}
.hero p{margin:5px 0 0 0;opacity:.88;font-size:.95rem}
.kpi{background:#fff;border:1px solid #dce5e7;border-radius:8px;padding:14px 16px;
     box-shadow:0 1px 2px rgba(23,50,77,.05);height:100%}
.kpi .l{font-size:.70rem;color:#607586;text-transform:uppercase;letter-spacing:.05em;font-weight:600}
.kpi .v{font-size:1.65rem;font-weight:700;line-height:1.25;color:#17324d}
.kpi .s{font-size:.76rem;color:#607586}
.chip{display:inline-block;padding:2px 11px;border-radius:999px;color:#fff;font-size:.8rem;font-weight:600}
.card{background:#fff;border:1px solid #dce5e7;border-left:4px solid #0f766e;border-radius:8px;
      padding:14px 16px;margin-bottom:10px;color:#28445b}
.card h4{margin:0 0 4px 0;font-size:.86rem;color:#17324d;text-transform:uppercase;letter-spacing:.04em}
.card p{margin:0;font-size:.93rem;line-height:1.5}
.step{background:#f7f9fa;border:1px solid #dce5e7;border-radius:7px;padding:9px 12px;margin-bottom:6px;
      color:#28445b;font-size:.92rem}
.objective{background:#edf7f5;border:1px solid #9ed5ca;border-radius:8px;padding:14px 18px;color:#155e59;
      font-size:1.02rem;margin-bottom:12px}
section[data-testid="stSidebar"]{background:#f1f5f5;border-right:1px solid #dce5e7}
div[data-testid="stMetric"]{background:#fff;border:1px solid #dce5e7;border-radius:8px;padding:10px}
</style>
""",
    unsafe_allow_html=True,
)


def kpi(label, value, sub="", color="#17324d"):
    st.markdown(
        f'<div class="kpi"><div class="l">{label}</div>'
        f'<div class="v" style="color:{color}">{value}</div><div class="s">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def chip(text, color):
    return f'<span class="chip" style="background:{color}">{text}</span>'


def card(title, text, color="#0f766e"):
    st.markdown(
        f'<div class="card" style="border-left-color:{color}"><h4>{title}</h4><p>{text}</p></div>',
        unsafe_allow_html=True,
    )


def chart_note(text):
    st.caption(f"**What this means:** {text}")


def show_df(df, **kw):
    """st.dataframe stretched to full width, tolerant of Streamlit version differences."""
    try:
        st.dataframe(df, width="stretch", **kw)
    except Exception:
        st.dataframe(df, use_container_width=True, **kw)


def show_fig(fig, height=360, **kw):
    fig.update_layout(
        template="plotly_white", height=height, margin=dict(l=10, r=10, t=45, b=10),
        font=dict(size=12, color="#17324d"),
        colorway=["#0f766e", "#17324d", "#b7791f", "#3f8f6b", "#356b82"],
        paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
        legend=dict(orientation="h", y=-0.2),
    )
    try:
        st.plotly_chart(fig, width="stretch", **kw)
    except Exception:
        st.plotly_chart(fig, use_container_width=True, **kw)


def editor(df, **kw):
    try:
        return st.data_editor(df, width="stretch", **kw)
    except Exception:
        return st.data_editor(df, use_container_width=True, **kw)


def pct(x, d=1):
    return "—" if pd.isna(x) else f"{x:.{d}f}%"


# ──────────────────────────────────────────────────────────────────────────────
# DATA LOADING — Excel (Phase 1)
# ──────────────────────────────────────────────────────────────────────────────
COLS = ["roll", "name", "written80", "written60", "cycle20", "cycle10", "total", "pct", "rank",
        "written_pct", "cycle_pct", "gap", "marks_lost", "band", "priority", "format_pattern",
        "milestone", "snapshot", "diagnostic", "mindset", "format_insight", "path", "format_action"]


def short_band(full: str) -> str:
    f = str(full)
    for key, name in [("True Mastery", "Mastery"), ("Near-Mastery", "Near-Mastery"),
                      ("Proficient", "Proficient"), ("Competent", "Competent"),
                      ("Developing", "Developing"), ("Foundational", "Foundational Cracks"),
                      ("Conceptual", "Conceptual Gaps"), ("Critical", "Critical"),
                      ("Not Taken", "Incomplete")]:
        if key in f:
            return name
    return "Incomplete"


def band_of_pct(p):
    if pd.isna(p):
        return "Incomplete"
    p = round(float(p), 6)
    for name, _, lo, hi in reversed(BANDS):
        if p >= lo:
            return name
    return "Critical"


def addon_for(pattern: str) -> str:
    p = str(pattern)
    if "Cycle-Dominant" in p or "Cycle-Leaning" in p:
        return "Weekly cumulative quiz (1 current + 2 earlier chapters) + timed section practice"
    if "Written-Leaning" in p or "Written-Dominant" in p:
        return "8–10 min timed drills + 'first-two-minutes' routine for short tests"
    return "—"


@st.cache_data(show_spinner=False)
def load_excel(data: bytes):
    raw = pd.read_excel(io.BytesIO(data), sheet_name="Learning Paths", header=2)
    raw = raw.iloc[:, : len(COLS)]
    raw.columns = COLS
    S = raw[pd.to_numeric(raw["roll"], errors="coerce").notna()].copy()
    S["roll"] = S["roll"].astype(int)
    S["name"] = S["name"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    for c in ["written80", "written60", "cycle20", "cycle10", "total", "pct", "rank",
              "written_pct", "cycle_pct", "gap", "marks_lost"]:
        S[c] = pd.to_numeric(S[c], errors="coerce")
    for c in ["pct", "written_pct", "cycle_pct"]:
        S[c] = S[c] * 100
    S["short"] = S["band"].map(short_band)
    S["color"] = S["short"].map(BAND_COLOR)
    S["complete"] = S["total"].notna()
    S["baseline"] = S["pct"].fillna(S["written_pct"])  # Rudra: written-only fallback
    S["label"] = S["roll"].astype(str).str.zfill(2) + " · " + S["name"]
    S["focus"] = S["short"].map(lambda k: ACT[k]["focus"])
    S["pattern_short"] = S["format_pattern"].astype(str).str.replace(r"\s*\(.*\)", "", regex=True)
    S["need"] = S.apply(
        lambda r: r["focus"] + ("" if "Balanced" in r["pattern_short"] or "Not assessable" in r["pattern_short"]
                                else f"  ·  Format signal: {r['pattern_short']}"), axis=1)
    S["activity"] = S["short"].map(lambda k: ACT[k]["activity"])
    S["addon"] = S["format_pattern"].map(addon_for)
    S["target_pct"] = S.apply(lambda r: min(95.0, r["baseline"] + ACT[r["short"]]["gain"]), axis=1)
    S = S.sort_values("roll").reset_index(drop=True)

    insights = []
    try:
        ins = pd.read_excel(io.BytesIO(data), sheet_name="Class Patterns & Insights", header=None)[0].dropna().tolist()
        i = 0
        while i < len(ins):
            if re.match(r"^\d+\.\s", str(ins[i])) and i + 1 < len(ins):
                insights.append((str(ins[i]), str(ins[i + 1])))
                i += 2
            else:
                i += 1
    except Exception:
        pass
    return S, insights


def find_workbook():
    cands = sorted(BASE.glob("*.xlsx"))
    cands = [c for c in cands if not c.name.startswith("~$")]
    pref = [c for c in cands if "result" in c.name.lower() or "analysis" in c.name.lower()]
    return (pref or cands or [None])[0]


# ──────────────────────────────────────────────────────────────────────────────
# DATA STORE — Phase 2 (CSV on disk; in-memory when Demo mode is on)
# ──────────────────────────────────────────────────────────────────────────────
def coerce(name, df):
    df = df.copy()
    for c in SCHEMA[name]:
        if c not in df:
            df[c] = np.nan
    df = df[SCHEMA[name]]
    df["roll"] = pd.to_numeric(df["roll"], errors="coerce")
    df = df.dropna(subset=["roll"])
    df["roll"] = df["roll"].astype(int)
    if name == "weekly":
        df["week"] = pd.to_numeric(df["week"], errors="coerce")
        df = df.dropna(subset=["week"])
        df["week"] = df["week"].astype(int)
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["score"] = pd.to_numeric(df["score"], errors="coerce")
        df["max_score"] = pd.to_numeric(df["max_score"], errors="coerce").fillna(10)
        df = df.dropna(subset=["score"])
        df["assessment"] = df["assessment"].fillna("Weekly Quiz").astype(str)
        df["concept"] = df["concept"].fillna("").astype(str)
    elif name == "interventions":
        for c in ["start_date", "target_date"]:
            df[c] = pd.to_datetime(df[c], errors="coerce")
        for c in ["sessions_planned", "sessions_done"]:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        df["target_pct"] = pd.to_numeric(df["target_pct"], errors="coerce")
        for c in ["activity", "addon", "support_format", "owner", "status", "notes"]:
            df[c] = df[c].fillna("").astype(str)
        df["status"] = df["status"].replace("", "Assigned")
    else:
        df["verified_date"] = pd.to_datetime(df["verified_date"], errors="coerce")
        for c in ["verification_status", "verified_gap", "verified_by", "notes"]:
            df[c] = df[c].fillna("").astype(str)
        df["verification_status"] = df["verification_status"].replace("", "Not reviewed")
    return df.reset_index(drop=True)


def load_table(name):
    if st.session_state.get("demo"):
        key = f"demo_{name}"
        if key not in st.session_state:
            return coerce(name, pd.DataFrame(columns=SCHEMA[name]))
        return coerce(name, st.session_state[key])
    p = FILES[name]
    df = pd.read_csv(p) if p.exists() else pd.DataFrame(columns=SCHEMA[name])
    return coerce(name, df)


def save_table(name, df):
    df = coerce(name, df)
    if st.session_state.get("demo"):
        st.session_state[f"demo_{name}"] = df
    else:
        df.to_csv(FILES[name], index=False)


def make_intervention_row(r, start=None):
    a = ACT[r["short"]]
    start = start or date.today()
    return dict(
        roll=int(r["roll"]), activity=a["activity"], addon=r["addon"], support_format=a["fmt"],
        owner="Class teacher", start_date=pd.Timestamp(start),
        target_date=pd.Timestamp(start + timedelta(weeks=a["weeks"])),
        sessions_planned=a["sessions"], sessions_done=0, status="Assigned",
        target_pct=round(r["target_pct"], 1), notes="",
    )


def generate_demo(S):
    """Synthetic, clearly-labelled demo data so the Phase-2 views can be explored."""
    rng = np.random.default_rng(8)
    start = date(2026, 9, 28)
    rows = []
    skipped = set(rng.choice(S["roll"].to_numpy(), 4, replace=False).tolist())
    for _, r in S.iterrows():
        slope = 0.0035 if r["roll"] in skipped else rng.uniform(0.012, 0.024)
        if r["short"] in ("Mastery", "Near-Mastery"):
            slope *= 0.2
        for w in range(1, 9):
            for k in range(2):
                c = CONCEPTS[(2 * (w - 1) + k) % 8]
                p = np.clip(r["baseline"] / 100 + slope * w + rng.normal(0, 0.06), 0.05, 1.0)
                if rng.random() < 0.04:
                    continue  # occasional absence
                rows.append(dict(roll=r["roll"], week=w, date=pd.Timestamp(start + timedelta(weeks=w - 1)),
                                 assessment="Weekly Quiz", concept=c, score=round(p * 10), max_score=10))
    wk = pd.DataFrame(rows)
    iv = []
    for _, r in S.iterrows():
        if r["roll"] in skipped:
            continue
        row = make_intervention_row(r, start)
        row["status"] = str(rng.choice(["In Progress", "In Progress", "Completed", "Assigned"]))
        row["sessions_done"] = int(row["sessions_planned"] * rng.uniform(0.3, 0.9))
        iv.append(row)
    return coerce("weekly", wk), coerce("interventions", pd.DataFrame(iv))


# ──────────────────────────────────────────────────────────────────────────────
# PHASE-2 ANALYTICS
# ──────────────────────────────────────────────────────────────────────────────
def level_of(p):
    if pd.isna(p):
        return "No data"
    return "Mastered" if p >= 85 else "Secure" if p >= 70 else "Developing" if p >= 50 else "Not yet"


LEVEL_COLOR = {"Mastered": "#0f766e", "Secure": "#65a30d", "Developing": "#d97706",
               "Not yet": "#dc2626", "No data": "#9ca3af"}


def with_pct(W):
    W = W.copy()
    W["pct"] = W["score"] / W["max_score"] * 100
    return W


def build_progress(S, W, I):
    """One row per student: need (always from Excel) + activity + evidence + loop status."""
    P = S[["roll", "name", "label", "short", "color", "need", "baseline", "target_pct"]].copy()
    P = P.rename(columns={"target_pct": "target_default"})
    # evidence
    if len(W):
        Wp = with_pct(W)
        wk = Wp.groupby(["roll", "week"])["pct"].mean().reset_index()
        rows = []
        for roll, g in wk.groupby("roll"):
            g = g.sort_values("week")
            n = len(g)
            k = 2 if n >= 4 else 1
            first, latest = g["pct"].iloc[:k].mean(), g["pct"].iloc[-k:].mean()
            rows.append(dict(roll=roll, n_weeks=n, first=first, latest=latest,
                             delta=(latest - first) if n >= 2 else np.nan))
        P = P.merge(pd.DataFrame(rows), on="roll", how="left")
    else:
        P[["n_weeks", "first", "latest", "delta"]] = np.nan
    P["n_weeks"] = P["n_weeks"].fillna(0).astype(int)
    # activity (latest assignment per student)
    if len(I):
        li = I.sort_values("start_date").groupby("roll").tail(1)
        P = P.merge(li[["roll", "activity", "addon", "support_format", "owner", "status",
                        "sessions_planned", "sessions_done", "target_pct", "start_date", "target_date"]],
                    on="roll", how="left")
    else:
        for c in ["activity", "addon", "support_format", "owner", "status", "sessions_planned",
                  "sessions_done", "target_pct", "start_date", "target_date"]:
            P[c] = np.nan
    P["has_activity"] = P["activity"].notna() & (P["activity"].astype(str) != "")
    P["target"] = P["target_pct"].fillna(P["target_default"])

    def loop(r):
        if not r["has_activity"]:
            return "No activity assigned"
        if r["n_weeks"] == 0:
            return "No evidence yet"
        if r["n_weeks"] == 1:
            return "Evidence building (1 week)"
        if r["latest"] >= r["target"]:
            return "Target met"
        if r["delta"] >= 3:
            return "Improving"
        if r["delta"] <= -3:
            return "Declining — review"
        return "Flat — adjust activity"

    P["loop"] = P.apply(loop, axis=1)
    P["need_ok"] = True
    P["evidence_ok"] = P["n_weeks"] >= 2
    P["objective_met"] = P["need_ok"] & P["has_activity"] & P["evidence_ok"]
    P["change"] = P["latest"] - P["baseline"]
    return P


# ──────────────────────────────────────────────────────────────────────────────
# PHASE 1 — PERFORMANCE & PERSONALIZED LEARNING
# ──────────────────────────────────────────────────────────────────────────────
def phase1(S, insights, pass_mark):
    C = S[S["complete"]]
    I = load_table("interventions")
    n_all, n_c = len(S), len(C)
    passed = int((C["pct"] >= pass_mark - 1e-9).sum())

    D = load_table("diagnostics")
    t1, t2, t3, t4, t5 = st.tabs(["🏠 Home – Class Overview", "🧑‍🎓 Student Performance",
                                  "🎯 Learning Improvement Plan", "🔎 Areas for Improvement",
                                  "📄 Reports"])

    # ---------- Class analysis ----------
    with t1:
        st.markdown("#### Existing marks and results")
        st.caption("This section shows the class's current Assessment-I results. Future progress is tracked separately.")
        st.markdown(
            '<span style="color:#16a34a;font-weight:600">● Doing well</span>&nbsp;&nbsp;'
            '<span style="color:#d97706;font-weight:600">● Needs practice</span>&nbsp;&nbsp;'
            '<span style="color:#dc2626;font-weight:600">● Needs immediate support</span>',
            unsafe_allow_html=True,
        )
        c = st.columns(6)
        with c[0]: kpi("Class strength", n_all, f"{n_c} complete /70 results")
        with c[1]: kpi("Class average", pct(C["pct"].mean()), f"{C['total'].mean():.1f} / 70 · median {C['total'].median():.0f}")
        with c[2]: kpi("Pass %", pct(passed / n_all * 100), f"{passed} of {n_all} at ≥ {pass_mark}%", "#0f766e")
        with c[3]: kpi("Scoring ≥ 80%", int((C["pct"] >= 80).sum()), f"{(C['pct'] >= 80).mean() * 100:.0f}% of complete", "#16a34a")
        with c[4]: kpi("Scoring < 50%", int((C["pct"] < 50).sum()), "need small-group support", "#dc2626")
        with c[5]: kpi("Result pending", int((~S["complete"]).sum()), "cycle test not taken (EX)", "#6b7280")
        st.write("")

        a, b = st.columns(2)
        with a:
            cnt = S.groupby("short").size().reindex(BAND_ORDER_ALL).dropna().reset_index()
            cnt.columns = ["Marks category", "Students"]
            fig = px.bar(cnt, x="Marks category", y="Students", color="Marks category", color_discrete_map=BAND_COLOR,
                         text="Students", title="Students by performance band")
            fig.update_layout(showlegend=False, xaxis_title=None)
            show_fig(fig)
            chart_note("This groups students by their current marks category. Green means students are doing well; amber means they need practice; red means they need immediate support.")
        with b:
            idx = np.floor(C["pct"].round(6) / 10).clip(0, 10).astype(int)
            labels = ["0–9%", "10–19%", "20–29%", "30–39%", "40–49%", "50–59%", "60–69%", "70–79%",
                      "80–89%", "90–99%", "100%"]
            h = idx.value_counts().reindex(range(11), fill_value=0)
            fig = go.Figure(go.Bar(x=labels, y=h.values, text=h.values, marker_color="#2563eb"))
            fig.update_layout(title="Percentage distribution (10% brackets)", xaxis_title=None, yaxis_title="Students")
            fig.add_vrect(x0=-0.5, x1=pass_mark / 10 - 0.5, fillcolor="#fecaca", opacity=0.35, line_width=0,
                          annotation_text="below pass mark", annotation_position="top left")
            show_fig(fig)
            chart_note("This shows how many students fall into each marks range. The shaded area is below the pass mark.")

        a, b = st.columns(2)
        with a:
            comp = pd.DataFrame({
                "Measure": ["Written (/60)", "Cycle Test (/10)", "Total (/70)"],
                "Class average %": [C["written_pct"].mean(), C["cycle_pct"].mean(), C["pct"].mean()],
            })
            fig = px.bar(comp, x="Measure", y="Class average %", text=comp["Class average %"].round(1),
                         color="Measure", color_discrete_sequence=["#2563eb", "#f59e0b", "#0f766e"],
                         title="The format gap: written paper vs cycle test")
            fig.update_layout(showlegend=False, xaxis_title=None, yaxis_range=[0, 105])
            show_fig(fig)
            chart_note("This compares average marks in the written paper, cycle test, and total result.")
        with b:
            fig = px.scatter(C, x="written_pct", y="cycle_pct", color="short", color_discrete_map=BAND_COLOR,
                             hover_name="name", size=C["marks_lost"].clip(lower=3),
                             labels={"written_pct": "Written %", "cycle_pct": "Cycle test %", "short": "Marks category"},
                             title="Written vs cycle test (below the line = written lags)")
            fig.add_shape(type="line", x0=0, y0=0, x1=100, y1=100, line=dict(dash="dash", color="#9ca3af"))
            show_fig(fig)
            chart_note("A point below the diagonal means the student's cycle-test marks are lower than their written-paper marks.")

        a, b = st.columns([3, 2])
        with a:
            st.markdown("#### Teacher action required")
            today = pd.Timestamp(date.today())
            overdue = I[(I["target_date"].notna()) & (I["target_date"] < today) &
                        (~I["status"].isin(["Completed"]))] if len(I) else pd.DataFrame()
            pending = S[~S["complete"]]
            follow_up = S[S["short"].isin(["Critical", "Conceptual Gaps", "Foundational Cracks"])]
            actions = []
            for _, row in follow_up.iterrows():
                actions.append({"Student": row["name"], "Action": "Review learning need and assign support",
                                "Priority": row["priority"]})
            for _, row in pending.iterrows():
                actions.append({"Student": row["name"], "Action": "Schedule pending assessment / make-up",
                                "Priority": "Assessment pending"})
            if len(overdue):
                labels = dict(zip(S["roll"], S["name"]))
                for _, row in overdue.iterrows():
                    actions.append({"Student": labels.get(row["roll"], f"Roll {row['roll']}"),
                                    "Action": f"Overdue intervention: {row['activity']}", "Priority": "Overdue"})
            if actions:
                show_df(pd.DataFrame(actions).drop_duplicates(), hide_index=True, height=260)
            else:
                st.success("No urgent follow-up, pending assessments, or overdue interventions.")
        with b:
            pat = S["pattern_short"].value_counts().reset_index()
            pat.columns = ["Pattern", "Students"]
            fig = px.bar(pat, y="Pattern", x="Students", orientation="h", text="Students",
                         title="Format patterns (retention & stamina signal)", color_discrete_sequence=["#7c3aed"])
            fig.update_layout(yaxis_title=None)
            fig.update_yaxes(autorange="reversed")
            show_fig(fig, height=420)

    # ---------- Student profile ----------
    with t2:
        query = st.text_input("Search by roll number or name", key="p1_search",
                              placeholder="e.g. 12 or Priya")
        options = S["label"].tolist()
        if query.strip():
            q = query.strip().lower()
            options = S[S["label"].str.lower().str.contains(q, regex=False)]["label"].tolist()
        if not options:
            st.warning("No student matches that search.")
            return
        sel = st.selectbox("Open student profile", options, key="p1_student")
        r = S[S["label"] == sel].iloc[0]
        st.markdown(f"### {r['name']}  &nbsp; {chip(r['short'], r['color'])} &nbsp; "
                    f"<span style='color:#6b7280;font-size:1rem'>Roll {r['roll']} · Priority: {r['priority']}</span>",
                    unsafe_allow_html=True)
        c = st.columns(6)
        with c[0]: kpi("Total", "—" if pd.isna(r["total"]) else f"{r['total']:.0f}/70", pct(r["pct"]), r["color"])
        with c[1]: kpi("Class rank", "—" if pd.isna(r["rank"]) else f"{r['rank']:.0f}", f"of {n_c}")
        with c[2]: kpi("Written", pct(r["written_pct"]), f"{r['written60']:.0f}/60")
        with c[3]: kpi("Cycle test", pct(r["cycle_pct"]), "—" if pd.isna(r["cycle10"]) else f"{r['cycle10']:.0f}/10")
        with c[4]: kpi("Marks lost", "—" if pd.isna(r["marks_lost"]) else f"{r['marks_lost']:.0f}", "out of 70")
        with c[5]: kpi("Next milestone", "", str(r["milestone"]))
        st.write("")

        a, b = st.columns([2, 3])
        with a:
            fig = go.Figure()
            cats = ["Written %", "Cycle test %", "Total %"]
            fig.add_bar(name=r["name"].split()[0], x=cats, y=[r["written_pct"], r["cycle_pct"], r["pct"]],
                        marker_color=r["color"], text=[pct(v, 0) for v in [r["written_pct"], r["cycle_pct"], r["pct"]]])
            fig.add_bar(name="Class average", x=cats, y=[C["written_pct"].mean(), C["cycle_pct"].mean(), C["pct"].mean()],
                        marker_color="#cbd5e1")
            fig.update_layout(barmode="group", title="Student vs class average", yaxis_range=[0, 105])
            show_fig(fig, height=340)
        with b:
            card("Score snapshot", r["snapshot"], r["color"])
            card("Learning need", r["diagnostic"], "#2563eb")
            card("Format pattern insight", r["format_insight"], "#0891b2")

        st.markdown("#### Confirm learning need")
        st.caption("The workbook suggestion is a starting point. A teacher can confirm or revise it after observing the student.")
        existing = D[D["roll"] == r["roll"]].tail(1)
        current = existing.iloc[0] if len(existing) else None
        current_status = current.get("verification_status", "Not reviewed") if current is not None else "Not reviewed"
        with st.form(f"diagnostic_form_{r['roll']}"):
            dc1, dc2 = st.columns(2)
            with dc1:
                st.text_area("Suspected gap (from workbook)", str(r["diagnostic"]), disabled=True)
                verification = st.selectbox(
                    "Learning need status",
                    ["Not reviewed", "Teacher verified", "Teacher revised", "Not observed"],
                    index=["Not reviewed", "Teacher verified", "Teacher revised", "Not observed"].index(
                        current_status if current_status in
                        ["Not reviewed", "Teacher verified", "Teacher revised", "Not observed"] else "Not reviewed"),
                )
                verified_gap = st.text_input(
                    "Teacher-confirmed weakness / need",
                    "" if current is None else str(current.get("verified_gap", "")),
                )
            with dc2:
                verified_by = st.text_input("Verified by", "" if current is None else str(current.get("verified_by", "")))
                verified_date = st.date_input(
                    "Verification date",
                    date.today() if current is None or pd.isna(current.get("verified_date", pd.NaT))
                    else current.get("verified_date").date(),
                )
                diagnostic_notes = st.text_area(
                    "Teacher notes",
                    "" if current is None else str(current.get("notes", "")),
                )
            save_diagnostic = st.form_submit_button("Save learning-need review", type="primary")
        if save_diagnostic:
            row = pd.DataFrame([{
                "roll": r["roll"], "verification_status": verification, "verified_gap": verified_gap,
                "verified_by": verified_by, "verified_date": pd.Timestamp(verified_date),
                "notes": diagnostic_notes,
            }])
            save_table("diagnostics", pd.concat([D[D["roll"] != r["roll"]], row], ignore_index=True))
            st.success("Learning-need review saved.")
            st.rerun()

        st.markdown("#### 🧭 Personalized learning path")
        steps = [re.sub(r"^\d\)\s*", "", s).strip() for s in re.split(r"\s*(?=\d\)\s)", str(r["path"])) if s.strip()]
        for i, s in enumerate(steps, 1):
            st.markdown(f'<div class="step"><b>{i}.</b> {s}</div>', unsafe_allow_html=True)
        card("Format-specific action", r["format_action"], "#0891b2")
        a = ACT[r["short"]]
        card("Recommended weekly improvement activity",
             f"<b>{a['activity']}</b><br>Format: {a['fmt']} · {a['freq']} · {a['weeks']} weeks<br>"
             f"<b>Evidence of progress:</b> {a['metric']}", "#0f766e")

    # ---------- Learning activities ----------
    with t3:
        st.markdown("Each marks category maps to one improvement activity and one **measurable evidence** rule — "
                    "these feed the Phase 2 tracker.")
        rows = []
        for band in BAND_ORDER_ALL:
            g = S[S["short"] == band]
            if g.empty:
                continue
            a = ACT[band]
            rows.append({"Marks category": band, "Students": len(g), "Priority": g["priority"].iloc[0],
                         "Improvement activity": a["activity"], "Support format": a["fmt"],
                         "Frequency": a["freq"], "Evidence of progress": a["metric"]})
        show_df(pd.DataFrame(rows), hide_index=True)

        st.markdown("#### Students in each group")
        for band in BAND_ORDER_ALL:
            g = S[S["short"] == band]
            if g.empty:
                continue
            with st.expander(f"{band} — {len(g)} student(s) · {ACT[band]['focus']}"):
                st.markdown(" ".join(chip(f"{x.name_} ({x.roll_})", BAND_COLOR[band])
                                     for x in g.rename(columns={"name": "name_", "roll": "roll_"}).itertuples()),
                            unsafe_allow_html=True)

        st.markdown("#### 🤝 Suggested peer-learning pairs")
        top = S[S["short"].isin(["Mastery", "Near-Mastery"])].sort_values("pct", ascending=False)
        low = S[S["short"].isin(["Developing", "Foundational Cracks"])].sort_values("pct")
        pairs = [{"Peer explainer": t["name"], "Explainer band": t["short"], "Learner": l["name"], "Learner band": l["short"]}
                 for (_, t), (_, l) in zip(top.iterrows(), low.iterrows())]
        if pairs:
            show_df(pd.DataFrame(pairs), hide_index=True)
            st.caption("Suggestion only — match on personality and seating. Explainers gain depth by teaching.")

        st.markdown("#### Interactive class improvement plan")
        pc1, pc2 = st.columns(2)
        plan_bands = pc1.multiselect("Show marks categories", BAND_ORDER_ALL, default=[], key="plan_bands")
        plan_search = pc2.text_input("Find a student in the plan", key="plan_search")
        plan = S[["roll", "name", "short", "priority", "need", "activity", "addon", "target_pct"]].copy()
        if plan_bands:
            plan = plan[plan["short"].isin(plan_bands)]
        if plan_search:
            plan = plan[plan["name"].str.contains(plan_search, case=False, na=False)]
        plan = plan.rename(columns={
            "roll": "Roll", "name": "Student", "short": "Marks category", "priority": "Priority",
            "need": "Learning need", "activity": "Improvement activity", "addon": "Extra practice",
            "target_pct": "Target %"})
        if not plan.empty:
            graph_data = (plan["Marks category"].value_counts()
                          .reindex(BAND_ORDER_ALL, fill_value=0)
                          .rename_axis("Marks category")
                          .reset_index(name="Students"))
            fig = px.bar(
                graph_data[graph_data["Students"] > 0],
                x="Marks category",
                y="Students",
                color="Marks category",
                color_discrete_map=BAND_COLOR,
                text="Students",
                title="Students by marks category",
                hover_data={"Students": True},
            )
            fig.update_layout(showlegend=False, xaxis_title=None, yaxis_title="Students")
            show_fig(fig, height=330)
            chart_note("Use the filters above to explore how many students need practice or immediate support.")
        show_df(plan, hide_index=True, column_config={
            "Learning need": st.column_config.TextColumn(width="large"),
            "Improvement activity": st.column_config.TextColumn(width="large"),
            "Extra practice": st.column_config.TextColumn(width="large"),
            "Target %": st.column_config.NumberColumn(format="%.0f")})

    # ---------- Insights ----------
    with t4:
        if insights:
            for head, body in insights:
                with st.expander(head, expanded=head.startswith(("1.", "2."))):
                    st.write(body)
        else:
            st.info("No additional class improvement notes were found in the workbook.")

    # ---------- Data table ----------
    with t5:
        f1, f2 = st.columns(2)
        bands = f1.multiselect("Filter by marks category", BAND_ORDER_ALL, default=[])
        q = f2.text_input("Search student name")
        D = S.copy()
        if bands:
            D = D[D["short"].isin(bands)]
        if q:
            D = D[D["name"].str.contains(q, case=False)]
        show_df(D[["roll", "name", "total", "pct", "rank", "written_pct", "cycle_pct", "gap", "marks_lost",
                   "short", "priority", "pattern_short", "milestone"]].rename(columns={
            "roll": "Roll", "name": "Student", "total": "Total /70", "pct": "Total %", "rank": "Rank",
            "written_pct": "Written %", "cycle_pct": "Cycle %", "gap": "Gap (pts)", "marks_lost": "Marks lost",
            "short": "Marks category", "priority": "Priority", "pattern_short": "Format pattern", "milestone": "Next milestone"}),
            hide_index=True,
            column_config={"Total %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.1f"),
                           "Written %": st.column_config.NumberColumn(format="%.1f"),
                           "Cycle %": st.column_config.NumberColumn(format="%.1f")})


# ──────────────────────────────────────────────────────────────────────────────
# PHASE 2 — CONTINUOUS LEARNING IMPROVEMENT
# ──────────────────────────────────────────────────────────────────────────────
def phase2(S):
    W, I = load_table("weekly"), load_table("interventions")
    D = load_table("diagnostics")
    P = build_progress(S, W, I)
    n = len(P)

    if st.session_state.get("demo"):
        st.warning("🧪 **Demo mode** — the weekly results and interventions below are *simulated* and are not saved to disk. "
                   "Switch demo mode off in the sidebar to work with your real records.", icon="🧪")
    elif W.empty and I.empty:
        st.info("No follow-up records yet. Start in **Learning Improvement Plan → Auto-assign**, then enter results in "
                "**Weekly Assessments**. Want to explore first? Turn on *Demo mode* in the sidebar.")

    tabs = st.tabs(["✅ Home – Class Overview", "📝 Student Progress", "🛠 Learning Improvement Plan",
                    "🔎 Areas for Improvement",                     "📈 Change in Marks", "🧑‍🎓 Track Student Progress",
                    "📄 Reports"])

    # ---------- Objective tracker ----------
    with tabs[0]:
        st.markdown("#### Future progress")
        st.caption("These results are collected after the existing marks above. Improvement is measured only after follow-up assessments.")
        st.markdown('<div class="objective">🎯 <b>Key objective:</b> every student has an <b>identified learning need</b>, '
                    'an <b>assigned improvement activity</b>, and <b>measurable evidence of progress</b>.</div>',
                    unsafe_allow_html=True)
        met = int(P["objective_met"].sum())
        c = st.columns(4)
        with c[0]: kpi("1 · Learning need identified", f"{int(P['need_ok'].sum())}/{n}", "from marks category + format signal", "#0f766e")
        with c[1]: kpi("2 · Activity assigned", f"{int(P['has_activity'].sum())}/{n}", f"{n - int(P['has_activity'].sum())} still unassigned",
                       "#0f766e" if P["has_activity"].all() else "#d97706")
        with c[2]: kpi("3 · Evidence (≥ 2 weeks)", f"{int(P['evidence_ok'].sum())}/{n}", "students with measurable trend",
                       "#0f766e" if P["evidence_ok"].all() else "#d97706")
        with c[3]: kpi("Objective met", f"{met}/{n}", f"{met / n * 100:.0f}% of the class", "#16a34a" if met == n else "#2563eb")
        st.progress(met / n)

        st.markdown("##### Student progress status")
        ls = P["loop"].value_counts().reset_index()
        ls.columns = ["Status", "Students"]
        a, b = st.columns([2, 3])
        with a:
            fig = px.pie(ls, names="Status", values="Students", hole=0.55, color="Status", color_discrete_map=STATUS_COLOR)
            fig.update_layout(showlegend=True, legend=dict(orientation="v", y=0.5, x=1.0))
            show_fig(fig, height=300)
        with b:
            att = P[P["loop"].isin(["No activity assigned", "No evidence yet", "Declining — review",
                                    "Flat — adjust activity"])].copy()
            att["_o"] = att["short"].map({k: i for i, k in enumerate(BAND_NAMES)})
            att = att.sort_values(["_o"])
            st.markdown(f"**⚠️ Needs attention ({len(att)})** — lowest bands first")
            if att.empty:
                st.success("Every student has a need, an activity and positive evidence. 🎉")
            else:
                show_df(att[["name", "short", "loop"]].rename(columns={"name": "Student", "short": "Marks category", "loop": "Why"}),
                        hide_index=True, height=250)

        st.markdown("##### Student-by-student")
        f = st.multiselect("Filter by status", list(STATUS_COLOR), default=[])
        T = P if not f else P[P["loop"].isin(f)]
        tbl = T[["roll", "name", "short", "need", "activity", "support_format", "status", "sessions_done",
                 "sessions_planned", "n_weeks", "baseline", "latest", "target", "change", "loop"]].rename(columns={
            "roll": "Roll", "name": "Student", "short": "Marks category", "need": "Learning need", "activity": "Assigned activity",
            "support_format": "Support", "status": "Activity status", "sessions_done": "Sessions done",
            "sessions_planned": "Sessions planned", "n_weeks": "Weeks of evidence", "baseline": "Baseline %",
            "latest": "Latest %", "target": "Target %", "change": "Change (pts)", "loop": "Loop status"})
        sty = tbl.style.map(lambda v: f"color:{STATUS_COLOR.get(v, '#111')};font-weight:600", subset=["Loop status"])
        show_df(sty, hide_index=True, column_config={
            "Baseline %": st.column_config.NumberColumn(format="%.1f"),
            "Latest %": st.column_config.NumberColumn(format="%.1f"),
            "Target %": st.column_config.NumberColumn(format="%.0f"),
            "Change (pts)": st.column_config.NumberColumn(format="%+.1f"),
            "Learning need": st.column_config.TextColumn(width="large"),
            "Assigned activity": st.column_config.TextColumn(width="large")})
        st.caption("Use the filters above to explore the tracker by current student progress status.")

    # ---------- Weekly assessments ----------
    with tabs[1]:
        if len(W):
            Wp = with_pct(W)
            tr = Wp.groupby("week").agg(avg=("pct", "mean"), students=("roll", "nunique")).reset_index()
            fig = go.Figure()
            fig.add_scatter(x=tr["week"], y=tr["avg"], mode="lines+markers", name="Class average %",
                            line=dict(color="#2563eb", width=3))
            fig.add_hline(y=S["baseline"].mean(), line_dash="dot", line_color="#9ca3af",
                          annotation_text="Assessment-I baseline")
            fig.update_layout(title="Class average by week", xaxis_title="Week", yaxis_title="Average %",
                              yaxis_range=[0, 100])
            show_fig(fig, height=300)
        st.markdown("#### Enter follow-up assessment results")
        with st.form("weekly_form"):
            nxt = int(W["week"].max()) + 1 if len(W) else 1
            c = st.columns(5)
            wk = c[0].number_input("Week #", 1, 60, nxt)
            dt = c[1].date_input("Date", date.today())
            at = c[2].selectbox("Assessment type", ASSESSMENT_TYPES)
            cp = c[3].selectbox("Concept", CONCEPTS)
            mx = c[4].number_input("Max score", 1, 100, 10)
            grid = editor(pd.DataFrame({"roll": S["roll"], "student": S["name"], "score": np.nan}),
                          hide_index=True, disabled=["roll", "student"], height=320, key="grid",
                          column_config={"score": st.column_config.NumberColumn("Score (leave blank if absent)", min_value=0)})
            ok = st.form_submit_button("💾 Save results", type="primary")
        if ok:
            new = grid.dropna(subset=["score"]).copy()
            if (new["score"] > mx).any():
                st.error("A score is higher than the max score — please correct it.")
            elif new.empty:
                st.warning("No scores entered.")
            else:
                new = new.assign(week=wk, date=pd.Timestamp(dt), assessment=at, concept=cp, max_score=mx)
                keep = W[~((W["week"] == wk) & (W["assessment"] == at) & (W["concept"] == cp))]
                save_table("weekly", pd.concat([keep, new[SCHEMA["weekly"]]], ignore_index=True))
                st.success(f"Saved {len(new)} scores for week {wk} · {cp}.")
                st.rerun()

        with st.expander("📤 Bulk upload / template / edit raw records"):
            tpl = pd.DataFrame({"roll": S["roll"], "week": 1, "date": str(date.today()), "assessment": "Weekly Quiz",
                                "concept": CONCEPTS[0], "score": "", "max_score": 10})
            st.download_button("Download CSV template", tpl.to_csv(index=False).encode("utf-8"),
                               "weekly_template.csv", "text/csv")
            up = st.file_uploader("Upload completed CSV (columns: roll, week, date, assessment, concept, score, max_score)",
                                  type="csv", key="wk_up")
            if up is not None and st.button("Import uploaded CSV"):
                try:
                    inc = coerce("weekly", pd.read_csv(up))
                    inc = inc[inc["roll"].isin(S["roll"])]
                    allw = pd.concat([W, inc]).drop_duplicates(["roll", "week", "assessment", "concept"], keep="last")
                    save_table("weekly", allw)
                    st.success(f"Imported {len(inc)} rows.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not import: {e}")
            st.caption("Edit or delete individual records below, then save.")
            ed = editor(W, num_rows="dynamic", key="wk_raw", hide_index=True, height=260)
            if st.button("Save edited records"):
                save_table("weekly", ed)
                st.success("Saved.")
                st.rerun()

    # ---------- Interventions ----------
    with tabs[2]:
        un = P[~P["has_activity"]]
        c = st.columns(4)
        with c[0]: kpi("Students with activity", f"{int(P['has_activity'].sum())}/{n}")
        with c[1]: kpi("In progress", int((I["status"] == "In Progress").sum()) if len(I) else 0, color="#2563eb")
        with c[2]: kpi("Completed", int((I["status"] == "Completed").sum()) if len(I) else 0, color="#0f766e")
        done = I["sessions_done"].sum() / I["sessions_planned"].sum() * 100 if len(I) and I["sessions_planned"].sum() else 0
        with c[3]: kpi("Sessions delivered", f"{done:.0f}%", "of all planned sessions")
        st.write("")
        a, b = st.columns([3, 2])
        a.markdown(f"**{len(un)} student(s) have no assigned activity.**")
        if len(un) and b.button(f"⚡ Auto-assign recommended activities to {len(un)} student(s)", type="primary"):
            add = pd.DataFrame([make_intervention_row(S[S["roll"] == r].iloc[0]) for r in un["roll"]])
            save_table("interventions", pd.concat([I, add], ignore_index=True))
            st.rerun()
        st.caption("Recommended activity comes from the student's marks category. Edit any cell — owner, format, status, "
                   "sessions done, target — then save. Use the + row to add a second activity if needed.")
        lab = dict(zip(S["roll"], S["label"]))
        rev = {v: k for k, v in lab.items()}
        E = I.copy()
        E.insert(0, "student", E["roll"].map(lab))
        E = E.drop(columns="roll")
        ed = editor(E, num_rows="dynamic", hide_index=True, key="iv_ed", height=420, column_config={
            "student": st.column_config.SelectboxColumn("Student", options=S["label"].tolist(), required=True),
            "status": st.column_config.SelectboxColumn("Status", options=STATUSES, required=True),
            "start_date": st.column_config.DateColumn("Start"), "target_date": st.column_config.DateColumn("Review by"),
            "sessions_planned": st.column_config.NumberColumn("Sessions planned", min_value=0),
            "sessions_done": st.column_config.NumberColumn("Sessions done", min_value=0),
            "target_pct": st.column_config.NumberColumn("Target %", min_value=0, max_value=100),
            "activity": st.column_config.TextColumn("Activity", width="large"),
            "addon": st.column_config.TextColumn("Format add-on", width="large"),
            "support_format": "Support format", "owner": "Owner", "notes": "Notes"})
        if st.button("💾 Save interventions"):
            out = ed.dropna(subset=["student"]).copy()
            out.insert(0, "roll", out["student"].map(rev))
            save_table("interventions", out.drop(columns="student"))
            st.success("Saved.")
            st.rerun()

    # ---------- Concept mastery ----------
    with tabs[3]:
        if W.empty or (W["concept"] == "").all():
            st.info("Concept mastery appears once weekly results are tagged with a concept.")
        else:
            Wp = with_pct(W[W["concept"] != ""])
            # latest score per student × concept
            last = (Wp.sort_values(["week", "date"]).groupby(["roll", "concept"]).tail(1)
                    .groupby(["roll", "concept"])["pct"].mean().reset_index())
            piv = last.pivot(index="roll", columns="concept", values="pct").reindex(S["roll"])
            piv = piv[[c for c in CONCEPTS if c in piv.columns] + [c for c in piv.columns if c not in CONCEPTS]]
            cls = piv.mean().sort_values()
            c = st.columns(4)
            for col, lv in zip(c, ["Mastered", "Secure", "Developing", "Not yet"]):
                share = piv.map(lambda v: level_of(v) == lv).sum().sum() / max(piv.notna().sum().sum(), 1) * 100
                with col: kpi(f"{lv}", f"{share:.0f}%", {"Mastered": "≥ 85%", "Secure": "70–84%", "Developing": "50–69%",
                                                          "Not yet": "< 50%"}[lv], LEVEL_COLOR[lv])
            st.write("")
            a, b = st.columns([2, 3])
            with a:
                cd = cls.reset_index()
                cd.columns = ["Concept", "Class mastery %"]
                cd["Level"] = cd["Class mastery %"].map(level_of)
                fig = px.bar(cd, y="Concept", x="Class mastery %", color="Level", color_discrete_map=LEVEL_COLOR,
                             orientation="h", text=cd["Class mastery %"].round(0), title="Class mastery by concept")
                fig.update_layout(xaxis_range=[0, 100], yaxis_title=None)
                show_fig(fig, height=420)
            with b:
                order = S.sort_values("baseline")["roll"].tolist()
                z = piv.reindex(order)
                names = S.set_index("roll").loc[order, "name"]
                fig = go.Figure(go.Heatmap(
                    z=z.values, x=z.columns, y=names, zmin=0, zmax=100, xgap=2, ygap=2,
                    colorscale=[[0, "#dc2626"], [0.5, "#fbbf24"], [0.7, "#a3e635"], [1, "#0f766e"]],
                    hovertemplate="%{y}<br>%{x}: %{z:.0f}%<extra></extra>", colorbar=dict(title="%")))
                fig.update_layout(title="Student × concept (latest score) — weakest students at top")
                fig.update_yaxes(autorange="reversed", tickfont=dict(size=9))
                fig.update_xaxes(tickangle=-40, tickfont=dict(size=9))
                show_fig(fig, height=520)

            st.markdown("##### 🔁 Areas needing more practice — concepts where students are 'Not yet' (< 50%)")
            rt = []
            for cname in piv.columns:
                low = piv[cname][piv[cname] < 50].index
                if len(low):
                    rt.append({"Concept": cname, "Class mastery %": round(piv[cname].mean(), 1), "Students below 50%": len(low),
                               "Who": ", ".join(S.set_index("roll").loc[low, "name"].str.split().str[0])})
            if rt:
                show_df(pd.DataFrame(rt).sort_values("Students below 50%", ascending=False), hide_index=True)
            else:
                st.success("No student is below 50% on any assessed concept.")

    # ---------- Before & after ----------
    with tabs[4]:
        if W.empty:
            st.info("Enter at least two weeks of results to see before-and-after analysis.")
        else:
            types = st.multiselect("Assessment types to include", sorted(W["assessment"].unique()),
                                   default=sorted(W["assessment"].unique()))
            Wf = W[W["assessment"].isin(types)]
            Pf = build_progress(S, Wf, I)
            B = Pf[Pf["n_weeks"] >= 1].copy()
            B["band_after"] = B["latest"].map(band_of_pct)
            B["moved"] = B["band_after"].map(BAND_NAMES.index) - B["short"].map(lambda k: BAND_NAMES.index(k) if k in BAND_NAMES else 0)
            c = st.columns(5)
            with c[0]: kpi("Average before", pct(B["baseline"].mean()), "Assessment-I (written + cycle)")
            with c[1]: kpi("Average now", pct(B["latest"].mean()), "latest weeks", "#2563eb")
            with c[2]: kpi("Average change", f"{B['change'].mean():+.1f} pts", "", "#0f766e" if B["change"].mean() >= 0 else "#dc2626")
            with c[3]: kpi("Students improved", f"{int((B['change'] > 0).sum())}/{len(B)}", "latest above baseline", "#16a34a")
            with c[4]: kpi("Moved up a band", int((B["moved"] > 0).sum()), f"{int((B['moved'] < 0).sum())} moved down", "#2563eb")
            st.write("")

            a, b = st.columns(2)
            with a:
                D = B.sort_values("baseline")
                fig = go.Figure()
                for _, r in D.iterrows():
                    fig.add_scatter(x=[r["baseline"], r["latest"]], y=[r["name"]] * 2, mode="lines",
                                    line=dict(color="#cbd5e1", width=2), showlegend=False, hoverinfo="skip")
                fig.add_scatter(x=D["baseline"], y=D["name"], mode="markers", name="Before (Assessment-I)",
                                marker=dict(color="#9ca3af", size=9))
                fig.add_scatter(x=D["latest"], y=D["name"], mode="markers", name="Now",
                                marker=dict(color=np.where(D["change"] >= 0, "#16a34a", "#dc2626"), size=10))
                fig.update_layout(title="Each student: before → now", xaxis_title="%", xaxis_range=[0, 105])
                fig.update_yaxes(tickfont=dict(size=9))
                show_fig(fig, height=620)
            with b:
                g = B.groupby("short").agg(Before=("baseline", "mean"), Now=("latest", "mean"), n=("roll", "count")).reindex(
                    [x for x in BAND_NAMES if x in B["short"].values]).reset_index()
                fig = go.Figure()
                fig.add_bar(x=g["short"], y=g["Before"], name="Before", marker_color="#cbd5e1")
                fig.add_bar(x=g["short"], y=g["Now"], name="Now", marker_color="#2563eb")
                fig.update_layout(barmode="group", title="By starting band — average before vs now", yaxis_title="%")
                show_fig(fig, height=300)

                if B["has_activity"].nunique() > 1:
                    gg = B.groupby("has_activity")["change"].mean()
                    fig = go.Figure(go.Bar(x=["No activity assigned", "Activity assigned"],
                                           y=[gg.get(False, 0), gg.get(True, 0)], marker_color=["#f87171", "#16a34a"],
                                           text=[f"{gg.get(False, 0):+.1f}", f"{gg.get(True, 0):+.1f}"]))
                    fig.update_layout(title="Average change (pts): with vs without an assigned activity")
                    show_fig(fig, height=300)
                    st.caption("Indicative only — groups are small and not randomly chosen.")

            st.markdown("##### Change in marks category")
            mv = B[["name", "short", "baseline", "band_after", "latest", "moved"]].copy()
            mv["Movement"] = mv["moved"].map(lambda m: "⬆️ Up" if m > 0 else "⬇️ Down" if m < 0 else "➡️ Same")
            show_df(mv.drop(columns="moved").rename(columns={"name": "Student", "short": "Marks category before", "baseline": "Before %",
                                                            "band_after": "Marks category now", "latest": "Now %"}).sort_values("Before %"),
                    hide_index=True, column_config={"Before %": st.column_config.NumberColumn(format="%.1f"),
                                                    "Now %": st.column_config.NumberColumn(format="%.1f")})
            with st.expander("⚠️ How to read this fairly"):
                st.write("The baseline is the Assessment-I total (written /60 + cycle /10) while 'now' is the average of recent "
                         "weekly assessments, which are shorter and may differ in difficulty. For the cleanest comparison, run a "
                         "**Re-test (Assessment-II)** on the same blueprint as Assessment-I and filter to that type above. "
                         "Treat class-level movement as stronger evidence than any single student's change.")

    # ---------- Student progress ----------
    with tabs[5]:
        query = st.text_input("Search by roll number or name", key="p2_search",
                              placeholder="e.g. 12 or Priya")
        options = S["label"].tolist()
        if query.strip():
            q = query.strip().lower()
            options = S[S["label"].str.lower().str.contains(q, regex=False)]["label"].tolist()
        if not options:
            st.warning("No student matches that search.")
            return
        sel = st.selectbox("Select a student", options, key="p2_student")
        r = S[S["label"] == sel].iloc[0]
        pr = P[P["roll"] == r["roll"]].iloc[0]
        st.markdown(f"### {r['name']} &nbsp; {chip(r['short'], r['color'])} &nbsp; "
                    f"{chip(pr['loop'], STATUS_COLOR.get(pr['loop'], '#6b7280'))}", unsafe_allow_html=True)
        c = st.columns(4)
        with c[0]:         card("1 · Learning need", r["need"], "#0f766e")
        with c[1]:
            card("2 · Assigned activity", (f"<b>{pr['activity']}</b><br>{pr['support_format']} · {pr['status']} · "
                                           f"{int(pr['sessions_done'])}/{int(pr['sessions_planned'])} sessions")
                 if pr["has_activity"] else "⚠️ Not assigned yet — go to Learning Improvement Plan.", "#2563eb" if pr["has_activity"] else "#dc2626")
        with c[2]:
            card("3 · Evidence of progress", f"{pr['n_weeks']} week(s) of data · baseline {pct(pr['baseline'], 0)} → "
                                             f"latest {pct(pr['latest'], 0)} (target {pct(pr['target'], 0)})", "#7c3aed")
        with c[3]:
            card("Success measure", ACT[r["short"]]["metric"], "#d97706")
        Ws = with_pct(W[W["roll"] == r["roll"]]) if len(W) else pd.DataFrame()
        if len(Ws):
            wk = Ws.groupby("week")["pct"].mean().reset_index()
            fig = go.Figure()
            fig.add_scatter(x=wk["week"], y=wk["pct"], mode="lines+markers", name="Weekly %", line=dict(color=r["color"], width=3))
            cl = with_pct(W).groupby("week")["pct"].mean().reset_index()
            fig.add_scatter(x=cl["week"], y=cl["pct"], mode="lines", name="Class average", line=dict(color="#cbd5e1", dash="dot"))
            fig.add_hline(y=r["baseline"], line_dash="dash", line_color="#9ca3af", annotation_text="baseline")
            fig.add_hline(y=pr["target"], line_dash="dash", line_color="#16a34a", annotation_text="target")
            fig.update_layout(title="Weekly progress", xaxis_title="Week", yaxis_title="%", yaxis_range=[0, 105])
            a, b = st.columns([3, 2])
            with a:
                show_fig(fig, height=340)
            with b:
                cm = Ws[Ws["concept"] != ""].sort_values("week").groupby("concept").tail(1)[["concept", "pct"]]
                if len(cm):
                    cm["Level"] = cm["pct"].map(level_of)
                    show_df(cm.rename(columns={"concept": "Concept", "pct": "Latest %"}).sort_values("Latest %"), hide_index=True,
                            column_config={"Latest %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f")})
        else:
            st.info("No weekly results recorded for this student yet.")

    # ---------- Reports ----------
    with tabs[6]:
        st.markdown("Generate a private, student-wise improvement report for sharing with the teacher or parent.")
        query = st.text_input("Search report student by roll number or name", key="report_search",
                              placeholder="e.g. 12 or Priya")
        options = S["label"].tolist()
        if query.strip():
            q = query.strip().lower()
            options = S[S["label"].str.lower().str.contains(q, regex=False)]["label"].tolist()
        if not options:
            st.warning("No student matches that search.")
        else:
            report_sel = st.selectbox("Student", options, key="report_student")
            student = S[S["label"] == report_sel].iloc[0]
            progress = P[P["roll"] == student["roll"]].iloc[0]
            diagnostic = D[D["roll"] == student["roll"]].tail(1)
            diagnostic_record = diagnostic.iloc[0] if len(diagnostic) else None
            student_weekly = with_pct(W[W["roll"] == student["roll"]]) if len(W) else pd.DataFrame()
            report_rows = [
                ("Student", student["name"]),
                ("Roll number", int(student["roll"])),
                ("Baseline", pct(student["baseline"], 1)),
                ("Marks category", student["short"]),
                ("Learning need", student["need"]),
                ("Assigned activity", progress["activity"] if progress["has_activity"] else "Not assigned"),
                ("Responsible teacher", progress["owner"] if progress["has_activity"] else "Not assigned"),
                ("Activity status", progress["status"] if progress["has_activity"] else "Not assigned"),
                ("Target date", str(progress["target_date"].date()) if progress["has_activity"] and pd.notna(progress["target_date"])
                 else "Not set"),
                ("Latest progress", pct(progress["latest"], 1)),
                ("Change from baseline", pct(progress["change"], 1)),
                ("Diagnostic verification", diagnostic_record.get("verification_status", "Not reviewed")
                 if diagnostic_record is not None else "Not reviewed"),
                ("Teacher-verified weakness", diagnostic_record.get("verified_gap", "")
                 if diagnostic_record is not None else ""),
            ]
            report = pd.DataFrame(report_rows, columns=["Measure", "Value"])
            report["Value"] = report["Value"].fillna("").astype(str)
            st.markdown("#### Student summary")
            rc = st.columns(4)
            with rc[0]:
                kpi("Marks category", student["short"], "current result", student["color"])
            with rc[1]:
                kpi("Baseline", pct(student["baseline"], 0), "existing result")
            with rc[2]:
                kpi("Latest progress", pct(progress["latest"], 0), "follow-up assessments", "#2563eb")
            with rc[3]:
                kpi("Change in marks", pct(progress["change"], 0), "from baseline",
                    "#16a34a" if pd.notna(progress["change"]) and progress["change"] >= 0 else "#dc2626")
            show_df(report, hide_index=True)
            if len(student_weekly):
                st.markdown("##### Weekly assessment record")
                weekly_chart = student_weekly.groupby("week", as_index=False)["pct"].mean()
                fig = px.line(weekly_chart, x="week", y="pct", markers=True,
                              title="Student progress over time",
                              labels={"week": "Assessment week", "pct": "Marks %"})
                fig.update_layout(yaxis_range=[0, 100])
                show_fig(fig, height=280)
                show_df(student_weekly[["week", "date", "assessment", "concept", "score", "max_score", "pct"]],
                        hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────
def main():
    st.markdown('<div class="hero"><h1>📐 Grade VIII-A Mathematics · Learning Improvement Dashboard</h1>'
                '<p>Every student has an identified learning need, an assigned improvement activity, '
                'and measurable evidence of progress.</p></div>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("Navigation")
        phase = st.radio("Dashboard", ["Home – Class Overview", "Track Student Progress"],
                         label_visibility="collapsed")
        with st.expander("How to use this dashboard"):
            st.markdown(
                "1. Start with **Home – Class Overview** to understand the class results.\n"
                "2. Open **Student Performance** to review one student's marks and learning need.\n"
                "3. Use **Areas for Improvement** to identify where students need more practice.\n"
                "4. Create a **Learning Improvement Plan** and record follow-up assessments.\n"
                "5. Use **Track Student Progress** only after new assessments are entered."
            )
        st.divider()
        st.subheader("Data")
        wb = find_workbook()
        up = st.file_uploader("Override workbook (optional)", type="xlsx")
        if up is not None:
            data, src = up.getvalue(), up.name
        elif wb is not None:
            data, src = wb.read_bytes(), wb.name
        else:
            st.error("No .xlsx workbook found next to app.py. Place it in the same folder or upload it here.")
            st.stop()
        st.caption(f"📄 {src}")
        pass_mark = st.slider("Pass mark (%)", 20, 60, 40, 1, help="Same input as the Excel 'Pass Percentage' sheet.")
        st.divider()
        demo = st.toggle("🧪 Demo mode (simulated Phase-2 data)", value=st.session_state.get("demo", False),
                         help="Fills Phase 2 with made-up weekly results and interventions so you can explore. Nothing is saved.")
        st.session_state["demo"] = demo

    try:
        S, insights = load_excel(data)
    except Exception as e:
        st.error(f"Could not read the 'Learning Paths' sheet: {e}")
        st.stop()

    if demo and "demo_weekly" not in st.session_state:
        st.session_state["demo_weekly"], st.session_state["demo_interventions"] = generate_demo(S)
    if not demo:
        for k in ("demo_weekly", "demo_interventions"):
            st.session_state.pop(k, None)

    if phase == "Home – Class Overview":
        phase1(S, insights, pass_mark)
    else:
        phase2(S)


main()
