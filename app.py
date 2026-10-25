
import re
import io
import json
import html
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# PAGE CONFIGURATION
st.set_page_config(
    page_title="ClauseGuard AI | Contract Intelligence",
    page_icon="🧑‍⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# DESIGN
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Playfair+Display:wght@500;600;700&display=swap');

html, body, [data-testid="stApp"] {
    font-family: 'DM Sans', sans-serif;
    color: #243B53;
}

.stApp {
    background: #F7F9FC;
}

.block-container {
    max-width: 1080px;
    padding-top: 2rem;
    padding-bottom: 4rem;
}

h1, h2, h3 {
    font-family: 'Playfair Display', serif !important;
    color: #193C58 !important;
}

p, label, button, input, textarea {
    font-family: 'DM Sans', sans-serif !important;
}

.hero {
    background: linear-gradient(120deg, #FFFFFF, #EAF3FA);
    border: 1px solid #DDE8F2;
    border-radius: 22px;
    padding: 38px 42px;
    margin-bottom: 28px;
    box-shadow: 0 8px 28px rgba(30, 65, 95, 0.04);
}

.logo {
    display: inline-flex;
    align-items: center;
    gap: 12px;
    color: #193C58;
    font-size: 21px;
    font-weight: 700;
}

.lawyer {
    background: #DDEDF8;
    border-radius: 16px;
    padding: 10px;
    font-size: 30px;
}

.eyebrow {
    color: #537F9E;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-top: 30px;
}

.hero-title {
    font-family: 'Playfair Display', serif;
    font-size: 42px;
    line-height: 1.2;
    color: #193C58;
    font-weight: 600;
    margin: 12px 0;
}

.hero-subtitle {
    color: #63798B;
    font-size: 16px;
    line-height: 1.7;
    max-width: 690px;
}

.section-label {
    font-size: 12px;
    letter-spacing: 1.5px;
    font-weight: 700;
    color: #6C879E;
    text-transform: uppercase;
    margin-bottom: 6px;
}

.metric-box {
    background: white;
    border: 1px solid #E0E8F0;
    border-radius: 15px;
    padding: 24px;
    margin-bottom: 12px;
}

.metric-value {
    font-size: 34px;
    font-weight: 700;
}

.metric-label {
    color: #708398;
    font-size: 13px;
}

.finding {
    background: white;
    border: 1px solid #E1E9F1;
    border-left: 5px solid #7D9DB8;
    padding: 20px 24px;
    border-radius: 12px;
    margin: 14px 0;
}

.finding-title {
    font-size: 18px;
    font-weight: 700;
    color: #193C58;
}

.risk-tag {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
    padding: 5px 10px;
    border-radius: 20px;
}

.stButton > button[kind="primary"] {
    background: #28658D;
    border: none;
    color: white;
    border-radius: 10px;
    font-weight: 600;
    padding: 10px 26px;
}

.stButton > button[kind="primary"]:hover {
    background: #194D70;
    color: white;
}

div[data-testid="stFileUploader"] {
    background: white;
    border-radius: 12px;
}

.footer {
    text-align: center;
    font-size: 12px;
    color: #8294A4;
    padding: 25px;
    margin-top: 40px;
    border-top: 1px solid #E0E8F0;
}
</style>
""", unsafe_allow_html=True)

# RISK DETECTION RULES
RULES = {
    "Unlimited liability": {
        "pattern": r"unlimited liability|without limitation of liability|liability shall not be limited",
        "risk": "HIGH",
        "explanation": "This provision may create uncapped financial exposure.",
        "recommendation": "Consider a reasonable financial cap and appropriate exclusions."
    },
    "Unilateral termination": {
        "pattern": r"terminate at any time|terminate without notice|terminate for any reason",
        "risk": "HIGH",
        "explanation": "One party may have broad rights to terminate the agreement.",
        "recommendation": "Review notice periods, grounds for termination and reciprocity."
    },
    "Broad indemnification": {
        "pattern": r"indemnify.*all losses|indemnify.*any and all|hold harmless.*all claims",
        "risk": "HIGH",
        "explanation": "An extensive indemnity may transfer substantial liabilities.",
        "recommendation": "Review scope, causation, exclusions and liability caps."
    },
    "Automatic renewal": {
        "pattern": r"automatically renew|automatic renewal|renewed automatically",
        "risk": "MEDIUM",
        "explanation": "The contract may renew without express further agreement.",
        "recommendation": "Check renewal periods and cancellation deadlines."
    },
    "Unilateral price changes": {
        "pattern": r"change prices at any time|increase fees at any time|modify fees without notice",
        "risk": "MEDIUM",
        "explanation": "One party may be entitled to alter the commercial terms.",
        "recommendation": "Consider advance notice and objective price-adjustment criteria."
    },
    "Exclusive jurisdiction": {
        "pattern": r"exclusive jurisdiction|courts of .* shall have exclusive",
        "risk": "REVIEW",
        "explanation": "Disputes may need to be resolved before specified courts.",
        "recommendation": "Verify the suitability of the forum and applicable law."
    },
    "Non-compete restriction": {
        "pattern": r"non.compet|shall not compete|prohibited from competing",
        "risk": "REVIEW",
        "explanation": "This provision may restrict future commercial activities.",
        "recommendation": "Review duration, territory, scope and enforceability."
    }
}

CATEGORIES = {
    "Liability": "liability damages losses indemnification financial exposure",
    "Termination": "termination cancellation expiry ending notice agreement",
    "Payment": "payment price fees invoice charges financial",
    "Confidentiality": "confidential information secrecy nondisclosure",
    "Dispute resolution": "jurisdiction courts arbitration governing law",
    "Intellectual property": "copyright trademarks intellectual property ownership",
    "Other": "general contractual provision miscellaneous"
}

SAMPLE = """
SERVICE AGREEMENT

1. Liability
The supplier shall have unlimited liability for all damages.

2. Termination
The customer may terminate this agreement at any time without notice.

3. Renewal
This agreement shall automatically renew for successive periods of twelve months.

4. Indemnification
The supplier shall indemnify the customer against any and all losses.

5. Payment
The customer shall pay invoices within thirty days.

6. Confidentiality
Both parties shall keep confidential information private.

7. Jurisdiction
The courts of London shall have exclusive jurisdiction.
"""

# DOCUMENT PROCESSING
def extract_text(file):
    data = file.getvalue()
    name = file.name.lower()

    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    if name.endswith(".docx"):
        document = Document(io.BytesIO(data))
        return "\n".join(p.text for p in document.paragraphs)

    return data.decode("utf-8", errors="replace")


def split_clauses(text):
    return [
        clause.strip()
        for clause in re.split(r"\n\s*\n", text.strip())
        if clause.strip()
    ]


def categorise(clauses):
    descriptions = list(CATEGORIES.values())
    vectorizer = TfidfVectorizer(stop_words="english")

    matrix = vectorizer.fit_transform(descriptions + clauses)
    similarities = cosine_similarity(
        matrix[len(descriptions):],
        matrix[:len(descriptions)]
    )

    labels = list(CATEGORIES.keys())

    return [
        labels[row.argmax()] if row.max() > 0 else "Other"
        for row in similarities
    ]


def analyse_contract(text):
    clauses = split_clauses(text)

    if not clauses:
        return []

    categories = categorise(clauses)
    findings = []

    for clause, category in zip(clauses, categories):
        for name, rule in RULES.items():
            if re.search(rule["pattern"], clause, re.I | re.S):
                findings.append({
                    "Issue": name,
                    "Risk": rule["risk"],
                    "Category": category,
                    "Clause": clause,
                    "Explanation": rule["explanation"],
                    "Recommendation": rule["recommendation"]
                })

    return findings


# HEADER
st.markdown("""
<div class="hero">
    <div class="logo">
        <span class="lawyer">🧑‍⚖️</span>
        ClauseGuard AI
    </div>
    <div class="eyebrow">INTELLIGENT CONTRACT REVIEW</div>
    <div class="hero-title">
        Clarity in every clause.
    </div>
    <div class="hero-subtitle">
        Understand your contracts with confidence.
        Identify potential legal risks, discover important provisions,
        and make more informed decisions with AI-assisted analysis.
    </div>
</div>
""", unsafe_allow_html=True)

# MAIN INTERFACE
st.markdown('<div class="section-label">01 / DOCUMENT REVIEW</div>',
            unsafe_allow_html=True)

st.header("Analyse your contract")

st.write(
    "Choose a sample agreement, upload your own document, "
    "or paste contract text below."
)

mode = st.radio(
    "Select input method",
    ["Sample contract", "Upload document", "Paste text"],
    horizontal=True
)

text = ""

if mode == "Sample contract":
    st.info("A fictional service agreement is ready for analysis.")
    with st.expander("Preview sample contract"):
        st.text(SAMPLE)
    text = SAMPLE

elif mode == "Upload document":
    uploaded = st.file_uploader(
        "Choose your contract",
        type=["pdf", "docx", "txt"]
    )

    if uploaded:
        try:
            text = extract_text(uploaded)
            if text.strip():
                st.success("Document loaded successfully.")
            else:
                st.warning(
                    "No readable text found. Scanned PDFs may require OCR."
                )
        except Exception as error:
            st.error(f"Unable to read this document: {error}")

else:
    text = st.text_area(
        "Paste contract text",
        height=240,
        placeholder="Paste your agreement here..."
    )

st.write("")

if st.button("🔍 Analyse Contract", type="primary", use_container_width=True):
    if not text.strip():
        st.warning("Please provide a contract first.")
    else:
        with st.spinner("Reviewing contractual provisions..."):
            st.session_state["findings"] = analyse_contract(text)
            st.session_state["analysed"] = True

# ANALYSIS REPORT
if st.session_state.get("analysed", False):
    findings = st.session_state.get("findings", [])

    st.divider()

    st.markdown(
        '<div class="section-label">02 / ANALYSIS RESULTS</div>',
        unsafe_allow_html=True
    )

    st.header("Your contract risk report")

    st.write(
        "The following findings identify provisions requiring "
        "closer consideration. Risk ratings are indicative."
    )

    high = sum(f["Risk"] == "HIGH" for f in findings)
    medium = sum(f["Risk"] == "MEDIUM" for f in findings)
    review = sum(f["Risk"] == "REVIEW" for f in findings)

    metrics = [
        ("Issues identified", len(findings), "#193C58"),
        ("High priority", high, "#BE615B"),
        ("Medium priority", medium, "#B88B45"),
        ("For review", review, "#447E9C")
    ]

    columns = st.columns(4)

    for column, (label, number, color) in zip(columns, metrics):
        with column:
            st.markdown(f"""
            <div class="metric-box">
                <div class="metric-value" style="color:{color}">
                    {number}
                </div>
                <div class="metric-label">{label}</div>
            </div>
            """, unsafe_allow_html=True)

    if findings:
        st.subheader("Risk overview")

        chart_data = pd.DataFrame({
            "Risk": ["HIGH", "MEDIUM", "REVIEW"],
            "Number of issues": [high, medium, review]
        }).set_index("Risk")

        st.bar_chart(chart_data, color="#739CB9")

        st.subheader("Detailed findings")

        priority_order = {"HIGH": 0, "MEDIUM": 1, "REVIEW": 2}

        sorted_findings = sorted(
            findings,
            key=lambda x: priority_order.get(x["Risk"], 3)
        )

        for number, finding in enumerate(sorted_findings, 1):
            risk = finding["Risk"]

            color = {
                "HIGH": "#BE615B",
                "MEDIUM": "#B88B45",
                "REVIEW": "#447E9C"
            }.get(risk, "#739CB9")

            title = html.escape(finding["Issue"])

            st.markdown(f"""
            <div class="finding" style="border-left-color:{color}">
                <div style="display:flex;justify-content:space-between;
                            align-items:center;gap:12px">
                    <span class="finding-title">{number}. {title}</span>
                    <span class="risk-tag"
                          style="color:{color};background:#F5F7FA">
                        {risk}
                    </span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.write("**Relevant provision**")
            st.info(finding["Clause"])

            st.write("**Why this matters**")
            st.write(finding["Explanation"])

            st.write("**Suggested review**")
            st.write(finding["Recommendation"])

            st.divider()

        st.subheader("Export your report")

        report = pd.DataFrame(sorted_findings)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                "📥 Download CSV",
                report.to_csv(index=False).encode("utf-8"),
                "ClauseGuard_Report.csv",
                "text/csv",
                use_container_width=True
            )

        with col2:
            st.download_button(
                "📥 Download JSON",
                json.dumps(sorted_findings, indent=2),
                "ClauseGuard_Report.json",
                "application/json",
                use_container_width=True
            )

    else:
        st.success("No predefined risk patterns were detected.")
        st.warning(
            "This does not mean the contract is risk-free. "
            "Only a limited set of issues is screened."
        )

# METHODOLOGY
st.divider()

with st.expander("How does ClauseGuard AI work?"):
    st.write("""
    ClauseGuard combines two methods:

    **1. Natural Language Processing (NLP)**

    TF-IDF and cosine similarity are used to categorise
    contractual provisions by their textual content.

    **2. Explainable risk screening**

    Transparent pattern-matching rules identify selected
    contractual expressions and flag potential concerns.

    Every finding includes its underlying contractual text,
    an explanation, and a suggested review action.

    **Limitations**

    This is an educational prototype, not a trained legal
    risk prediction model. It does not assess the full legal
    effect or enforceability of contractual terms. Results
    may include false positives and missed risks.
    """)

st.markdown("""
<div class="footer">
    🧑‍⚖️ ClauseGuard AI · Contract Intelligence Platform
    <br>
    Built for the Advanced AI Algorithms Project
    <br><br>
    Educational demonstration only · Not legal advice
</div>
""", unsafe_allow_html=True)
