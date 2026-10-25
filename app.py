
import re
import io
import json
import html
import hashlib
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ----------------------------------------
# PAGE SETTINGS
# ----------------------------------------

st.set_page_config(
    page_title="ClauseGuard AI",
    page_icon="🧑‍⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ----------------------------------------
# SOFT OLIVE DESIGN
# ----------------------------------------

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Serif+Display&display=swap');

:root {
    --olive: #697B55;
    --olive-dark: #485A39;
    --sage: #E8EDDF;
    --ivory: #F8F9F4;
    --cream: #FFFFFF;
    --border: #E0E7D8;
    --text: #344235;
    --muted: #71806B;
}

html, body, [data-testid="stApp"] {
    font-family: 'DM Sans', sans-serif !important;
}

.stApp {
    background-color: #F8F9F4;
    color: #344235;
}

.block-container {
    max-width: 1100px;
    padding-top: 1.8rem;
    padding-bottom: 3rem;
}

h1, h2, h3 {
    font-family: 'DM Serif Display', serif !important;
    color: #485A39 !important;
    font-weight: 400 !important;
}

p, label, button, input, textarea {
    font-family: 'DM Sans', sans-serif !important;
}

a {
    color: #536D44;
}

[data-testid="stHeader"] {
    background-color: #F8F9F4;
}

[data-testid="stToolbar"] {
    color: #485A39;
}

[data-testid="stRadio"] label {
    color: #485A39 !important;
}

div[role="radiogroup"] {
    background-color: #F0F3EA;
    border-radius: 14px;
    padding: 10px;
}

div[role="radiogroup"] label {
    border-radius: 10px;
    padding: 6px 10px;
}

.stButton > button[kind="primary"] {
    background-color: #697B55;
    color: white;
    border: 1px solid #697B55;
    border-radius: 12px;
    font-weight: 600;
    padding: 0.7rem 1.5rem;
}

.stButton > button[kind="primary"]:hover {
    background-color: #485A39;
    color: white;
    border-color: #485A39;
}

.stButton > button:not([kind="primary"]) {
    background-color: #FFFFFF;
    border: 1px solid #CBD7C0;
    border-radius: 12px;
    color: #485A39;
}

.stButton > button:not([kind="primary"]):hover {
    background-color: #E8EDDF;
    color: #344235;
    border-color: #697B55;
}

[data-testid="stFileUploader"] {
    background-color: white;
    border-radius: 14px;
    border: 1px solid #E0E7D8;
    padding: 12px;
}

[data-testid="stExpander"] {
    background: white;
    border: 1px solid #E0E7D8;
    border-radius: 14px;
}

[data-testid="stMetric"] {
    background: white;
    border: 1px solid #E0E7D8;
    border-radius: 14px;
    padding: 18px;
}

[data-testid="stAlert"] {
    border-radius: 12px;
}

[data-testid="stTextArea"] textarea {
    background-color: white;
    color: #344235;
    border-radius: 12px;
    border-color: #CBD7C0;
}

div[data-testid="stDownloadButton"] button {
    border-radius: 12px;
    color: #485A39;
    border: 1px solid #CBD7C0;
    background-color: white;
}

.brand {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 12px 0 18px;
}

.brand-icon {
    background: #E8EDDF;
    border-radius: 16px;
    padding: 12px;
    font-size: 28px;
}

.brand-name {
    font-family: 'DM Serif Display', serif;
    color: #485A39;
    font-size: 27px;
}

.brand-caption {
    color: #829177;
    font-size: 10px;
    letter-spacing: 2px;
    font-weight: 600;
}

.hero {
    background: linear-gradient(125deg, #FFFFFF, #EDF2E7);
    border: 1px solid #E0E7D8;
    border-radius: 26px;
    padding: 55px 48px;
    margin-top: 24px;
    margin-bottom: 30px;
}

.eyebrow {
    color: #798B6D;
    font-size: 12px;
    letter-spacing: 2px;
    font-weight: 700;
    text-transform: uppercase;
}

.hero-title {
    font-family: 'DM Serif Display', serif;
    font-size: clamp(40px, 5vw, 65px);
    line-height: 1.14;
    color: #485A39;
    margin: 18px 0;
}

.hero-description {
    color: #667560;
    font-size: 17px;
    line-height: 1.8;
    max-width: 670px;
}

.soft-card {
    background: white;
    border: 1px solid #E0E7D8;
    border-radius: 20px;
    padding: 26px;
    margin: 10px 0;
    min-height: 170px;
}

.card-icon {
    font-size: 25px;
    margin-bottom: 14px;
}

.card-title {
    color: #485A39;
    font-weight: 700;
    font-size: 17px;
    margin-bottom: 8px;
}

.card-text {
    color: #71806B;
    font-size: 14px;
    line-height: 1.7;
}

.page-label {
    font-size: 12px;
    color: #839477;
    letter-spacing: 2px;
    font-weight: 700;
    text-transform: uppercase;
    margin-top: 28px;
}

.page-heading {
    font-family: 'DM Serif Display', serif;
    color: #485A39;
    font-size: 43px;
    margin-top: 10px;
    margin-bottom: 12px;
}

.page-description {
    color: #71806B;
    font-size: 16px;
    line-height: 1.8;
    margin-bottom: 30px;
}

.finding {
    background: white;
    border: 1px solid #E0E7D8;
    border-left: 5px solid #697B55;
    border-radius: 14px;
    padding: 20px 24px;
    margin: 16px 0 12px;
}

.finding-title {
    color: #485A39;
    font-size: 18px;
    font-weight: 700;
}

.tag {
    display: inline-block;
    padding: 5px 12px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    margin-top: 8px;
}

.footer {
    color: #8B9982;
    text-align: center;
    font-size: 12px;
    border-top: 1px solid #E0E7D8;
    margin-top: 60px;
    padding: 30px;
    line-height: 2;
}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------
# CONTRACT RISK RULES
# ----------------------------------------

RULES = {
    "Unlimited liability": {
        "pattern": r"unlimited liability|without limitation of liability|liability shall not be limited",
        "risk": "HIGH",
        "explanation": "The provision may create uncapped financial exposure.",
        "recommendation": "Consider negotiating a liability cap and appropriate exclusions."
    },
    "Unilateral termination": {
        "pattern": r"terminate at any time|terminate without notice|terminate for any reason",
        "risk": "HIGH",
        "explanation": "A party may be able to terminate the agreement without adequate notice.",
        "recommendation": "Review notice periods, termination grounds and reciprocity."
    },
    "Broad indemnification": {
        "pattern": r"indemnify.*all losses|indemnify.*any and all|hold harmless.*all claims",
        "risk": "HIGH",
        "explanation": "The indemnity could transfer extensive liabilities.",
        "recommendation": "Consider the indemnity's scope, exclusions and financial limits."
    },
    "Automatic renewal": {
        "pattern": r"automatically renew|automatic renewal|renewed automatically",
        "risk": "MEDIUM",
        "explanation": "The agreement may renew without further express approval.",
        "recommendation": "Check renewal periods and cancellation requirements."
    },
    "Unilateral price changes": {
        "pattern": r"change prices at any time|increase fees at any time|modify fees without notice",
        "risk": "MEDIUM",
        "explanation": "One party may have broad discretion to change prices.",
        "recommendation": "Consider objective adjustment criteria and advance notice."
    },
    "Exclusive jurisdiction": {
        "pattern": r"exclusive jurisdiction|courts of .* shall have exclusive",
        "risk": "REVIEW",
        "explanation": "Disputes may need to be brought in a specified jurisdiction.",
        "recommendation": "Review the selected courts and governing law."
    },
    "Non-compete restriction": {
        "pattern": r"non.compet|shall not compete|prohibited from competing",
        "risk": "REVIEW",
        "explanation": "The clause may restrict future commercial activity.",
        "recommendation": "Review geographical scope, duration and enforceability."
    }
}

CATEGORIES = {
    "Liability": "liability damages indemnification loss",
    "Termination": "termination cancellation expiry notice",
    "Payment": "fees payment price invoice",
    "Confidentiality": "secrecy confidential nondisclosure",
    "Dispute resolution": "jurisdiction courts arbitration disputes",
    "Intellectual property": "copyright trademarks ownership",
    "Other": "general contract provisions"
}

SAMPLE = """
SERVICE AGREEMENT

1. LIABILITY

The supplier shall have unlimited liability for all damages.

2. TERMINATION

The customer may terminate this agreement at any time without notice.

3. RENEWAL

This agreement shall automatically renew for successive periods of twelve months.

4. INDEMNIFICATION

The supplier shall indemnify the customer against any and all losses.

5. PAYMENT

The customer shall pay invoices within thirty days.

6. CONFIDENTIALITY

Both parties shall keep confidential information private.

7. JURISDICTION

The courts of London shall have exclusive jurisdiction.
"""

# ----------------------------------------
# ANALYSIS FUNCTIONS
# ----------------------------------------

def extract_text(uploaded):
    data = uploaded.getvalue()
    name = uploaded.name.lower()

    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(
            page.extract_text() or ""
            for page in reader.pages
        )

    if name.endswith(".docx"):
        document = Document(io.BytesIO(data))
        return "\n".join(
            paragraph.text for paragraph in document.paragraphs
        )

    return data.decode("utf-8", errors="replace")


def split_clauses(text):
    return [
        item.strip()
        for item in re.split(r"\n\s*\n", text.strip())
        if item.strip()
    ]


def categorise(clauses):
    descriptions = list(CATEGORIES.values())

    matrix = TfidfVectorizer(
        stop_words="english"
    ).fit_transform(descriptions + clauses)

    similarities = cosine_similarity(
        matrix[len(descriptions):],
        matrix[:len(descriptions)]
    )

    names = list(CATEGORIES.keys())

    return [
        names[row.argmax()] if row.max() > 0 else "Other"
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

    order = {"HIGH": 0, "MEDIUM": 1, "REVIEW": 2}

    return sorted(
        findings,
        key=lambda item: order.get(item["Risk"], 3)
    )


# ----------------------------------------
# SITE HEADER AND NAVIGATION
# ----------------------------------------

st.markdown("""
<div class="brand">
    <div class="brand-icon">🧑‍⚖️</div>
    <div>
        <div class="brand-name">ClauseGuard AI</div>
        <div class="brand-caption">LEGAL INTELLIGENCE, SIMPLIFIED</div>
    </div>
</div>
""", unsafe_allow_html=True)

if "page" not in st.session_state:
    st.session_state.page = "Home"

page = st.radio(
    "Navigation",
    [
        "Home",
        "Analyse a Contract",
        "Legal News",
        "The Legal 500"
    ],
    horizontal=True,
    key="page",
    label_visibility="collapsed"
)

# ========================================
# PAGE 1 — HOME
# ========================================

if page == "Home":

    st.markdown("""
    <div class="hero">
        <div class="eyebrow">
            A MORE THOUGHTFUL WAY TO REVIEW CONTRACTS
        </div>

        <div class="hero-title">
            Clarity in<br>every clause.
        </div>

        <div class="hero-description">
            Contracts should inspire confidence, not confusion.
            ClauseGuard AI helps you identify potential contractual
            risks, understand important provisions and approach
            every agreement with greater clarity.
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button(
        "Begin your contract review  →",
        type="primary",
        use_container_width=True
    ):
        st.session_state.page = "Analyse a Contract"
        st.rerun()

    st.write("")
    st.write("")

    st.markdown("""
    <div class="page-label">THE PLATFORM</div>
    <div class="page-heading">
        A clearer perspective on contracts.
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    cards = [
        (
            "📄",
            "Understand",
            "Upload a contract and discover the provisions "
            "that may deserve closer attention."
        ),
        (
            "🌿",
            "Identify",
            "Detect selected contractual risks using "
            "explainable AI-assisted analysis."
        ),
        (
            "✦",
            "Decide",
            "Receive clear explanations and suggested "
            "questions for legal review."
        )
    ]

    for column, (icon, title, description) in zip(
        [col1, col2, col3], cards
    ):
        with column:
            st.markdown(f"""
            <div class="soft-card">
                <div class="card-icon">{icon}</div>
                <div class="card-title">{title}</div>
                <div class="card-text">{description}</div>
            </div>
            """, unsafe_allow_html=True)

    st.write("")
    st.write("")

    st.markdown("""
    <div class="page-label">HOW IT WORKS</div>
    <div class="page-heading">
        Simple by design.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    **01 — Share your document**

    Upload a PDF, Word document or text file.
    Alternatively, use our sample contract.

    **02 — Explore the analysis**

    ClauseGuard scans the document for selected risk
    indicators and categorises contractual language.

    **03 — Review the findings**

    Understand why provisions were flagged and
    download a structured report.
    """)

    st.write("")

    st.info(
        "ClauseGuard is an educational legal-technology prototype. "
        "It assists with preliminary document screening and does "
        "not replace professional legal advice."
    )

# ========================================
# PAGE 2 — CONTRACT ANALYSIS
# ========================================

elif page == "Analyse a Contract":

    st.markdown("""
    <div class="page-label">CONTRACT INTELLIGENCE</div>
    <div class="page-heading">Your contract, made clearer.</div>
    <div class="page-description">
        Start by uploading an agreement, exploring our sample
        contract or pasting your own text.
    </div>
    """, unsafe_allow_html=True)

    input_method = st.radio(
        "How would you like to begin?",
        [
            "Upload a contract",
            "Try a sample",
            "Paste text"
        ],
        horizontal=True
    )

    text = ""

    if input_method == "Upload a contract":

        st.write("#### Upload your document")

        uploaded = st.file_uploader(
            "Select a PDF, Word or text document",
            type=["pdf", "docx", "txt"]
        )

        if uploaded:
            try:
                text = extract_text(uploaded)

                if text.strip():
                    st.success("Your document has been loaded.")
                else:
                    st.warning(
                        "No readable text was found. "
                        "Scanned PDFs may require OCR."
                    )

            except Exception as error:
                st.error(f"Unable to read this file: {error}")

    elif input_method == "Try a sample":

        st.write("#### Sample service agreement")

        st.caption(
            "Explore ClauseGuard using a fictional "
            "commercial agreement."
        )

        text = SAMPLE

        with st.expander("Read the sample contract"):
            st.text(SAMPLE)

    else:

        st.write("#### Paste your agreement")

        text = st.text_area(
            "Contract text",
            height=280,
            placeholder="Paste the terms of your agreement here..."
        )

    # Avoid displaying findings for a different document.
    document_key = hashlib.sha256(
        text.encode("utf-8")
    ).hexdigest()

    st.write("")

    if st.button(
        "Analyse my contract  →",
        type="primary",
        use_container_width=True
    ):

        if not text.strip():
            st.warning("Please provide a contract first.")

        else:
            with st.spinner("Reviewing your contract..."):
                findings = analyse_contract(text)

                st.session_state["findings"] = findings
                st.session_state["document_key"] = document_key

    # RESULTS
    if (
        st.session_state.get("document_key") == document_key
        and "findings" in st.session_state
    ):

        findings = st.session_state["findings"]

        st.divider()

        st.markdown("""
        <div class="page-label">YOUR RESULTS</div>
        <div class="page-heading">Contract review report.</div>
        <div class="page-description">
            An overview of provisions flagged for further consideration.
        </div>
        """, unsafe_allow_html=True)

        high = sum(f["Risk"] == "HIGH" for f in findings)
        medium = sum(f["Risk"] == "MEDIUM" for f in findings)
        review = sum(f["Risk"] == "REVIEW" for f in findings)

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total findings", len(findings))

        with col2:
            st.metric("High priority", high)

        with col3:
            st.metric("Medium priority", medium)

        with col4:
            st.metric("For review", review)

        if findings:

            st.write("")
            st.subheader("Your risk overview")

            chart = pd.DataFrame({
                "Priority": ["High", "Medium", "Review"],
                "Findings": [high, medium, review]
            }).set_index("Priority")

            st.bar_chart(chart, color="#81966D")

            st.write("")
            st.subheader("Detailed findings")

            for index, finding in enumerate(findings, 1):

                risk = finding["Risk"]

                if risk == "HIGH":
                    color = "#B76D63"
                    background = "#F9EDEA"
                elif risk == "MEDIUM":
                    color = "#A68B55"
                    background = "#F8F2E5"
                else:
                    color = "#697B55"
                    background = "#E8EDDF"

                title = html.escape(finding["Issue"])

                st.markdown(f"""
                <div class="finding" style="border-left-color:{color}">
                    <div class="finding-title">
                        {index:02d} · {title}
                    </div>
                    <div class="tag"
                         style="color:{color};background:{background}">
                        {risk} PRIORITY
                    </div>
                </div>
                """, unsafe_allow_html=True)

                with st.expander(
                    "Read the provision and explanation",
                    expanded=True
                ):

                    st.write("**Relevant clause**")
                    st.info(finding["Clause"])

                    st.write("**Why it matters**")
                    st.write(finding["Explanation"])

                    st.write("**Suggested review**")
                    st.write(finding["Recommendation"])

            st.write("")
            st.subheader("Keep a copy of your report")

            report = pd.DataFrame(findings)

            col_a, col_b = st.columns(2)

            with col_a:
                st.download_button(
                    "Download CSV report",
                    report.to_csv(index=False).encode("utf-8"),
                    "ClauseGuard_Report.csv",
                    "text/csv",
                    use_container_width=True
                )

            with col_b:
                st.download_button(
                    "Download JSON report",
                    json.dumps(findings, indent=2),
                    "ClauseGuard_Report.json",
                    "application/json",
                    use_container_width=True
                )

        else:
            st.success(
                "No predefined risk indicators were identified."
            )
            st.warning(
                "This does not establish that the contract "
                "is free from legal or commercial risk."
            )

    st.write("")

    with st.expander("About our analysis methodology"):

        st.write("""
        ClauseGuard uses natural language processing
        (TF-IDF and cosine similarity) to classify
        contractual text.

        It also applies transparent pattern-matching
        rules to flag selected risk indicators.

        The prototype does not use a trained legal
        risk prediction model and cannot determine
        the legal enforceability of a provision.

        Findings may contain false positives or omit
        important risks. Every result should therefore
        be checked by a qualified professional.
        """)

# ========================================
# PAGE 3 — LEGAL NEWS
# ========================================

elif page == "Legal News":

    st.markdown("""
    <div class="page-label">LEGAL INSIGHTS</div>
    <div class="page-heading">Stay informed.</div>
    <div class="page-description">
        Explore legal developments, industry perspectives
        and professional commentary from established publications.
    </div>
    """, unsafe_allow_html=True)

    news_sources = [
        {
            "name": "Reuters — Legal",
            "description": (
                "Reporting on litigation, regulation, "
                "legal business and international developments."
            ),
            "url": "https://www.reuters.com/legal/"
        },
        {
            "name": "The Law Society Gazette",
            "description": (
                "Legal news, professional practice, "
                "regulation and developments in England and Wales."
            ),
            "url": "https://www.lawgazette.co.uk/"
        },
        {
            "name": "Legal Cheek",
            "description": (
                "News about law firms, legal careers, "
                "training and developments in the profession."
            ),
            "url": "https://www.legalcheek.com/"
        },
        {
            "name": "EUR-Lex",
            "description": (
                "Access to EU legislation, case law "
                "and official legal materials."
            ),
            "url": "https://eur-lex.europa.eu/"
        }
    ]

    for source in news_sources:

        st.markdown(f"""
        <div class="soft-card" style="min-height:auto">
            <div class="card-title">
                {html.escape(source["name"])}
            </div>
            <div class="card-text">
                {html.escape(source["description"])}
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.link_button(
            "Visit publication  ↗",
            source["url"],
            use_container_width=False
        )

        st.write("")

    st.caption(
        "These are links to independent external publications. "
        "ClauseGuard does not reproduce or independently "
        "verify their current articles."
    )

# ========================================
# PAGE 4 — THE LEGAL 500
# ========================================

elif page == "The Legal 500":

    st.markdown("""
    <div class="page-label">PROFESSIONAL RESOURCES</div>
    <div class="page-heading">The Legal 500.</div>
    <div class="page-description">
        Explore an established international directory
        of law firms and legal professionals.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="hero">
        <div class="eyebrow">
            DISCOVER THE LEGAL PROFESSION
        </div>

        <div class="hero-title">
            Find expertise.<br>
            Explore excellence.
        </div>

        <div class="hero-description">
            The Legal 500 provides independent research,
            rankings and commentary on law firms and
            legal professionals across jurisdictions
            and practice areas.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.link_button(
        "Visit The Legal 500  ↗",
        "https://www.legal500.com/",
        use_container_width=True
    )

    st.write("")
    st.write("")

    st.markdown("""
    <div class="page-label">WHY EXPLORE THE DIRECTORY?</div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="soft-card">
            <div class="card-icon">🌍</div>
            <div class="card-title">
                International perspective
            </div>
            <div class="card-text">
                Explore law firm research across
                numerous jurisdictions and sectors.
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="soft-card">
            <div class="card-icon">⚖️</div>
            <div class="card-title">
                Professional expertise
            </div>
            <div class="card-text">
                Research established legal practices,
                professionals and areas of specialisation.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.caption(
        "ClauseGuard AI is an independent educational project "
        "and is not affiliated with or endorsed by The Legal 500."
    )

# ----------------------------------------
# FOOTER
# ----------------------------------------

st.markdown("""
<div class="footer">
    🧑‍⚖️ ClauseGuard AI
    <br>
    Clarity in every clause.
    <br>
    An educational legal intelligence project
    <br><br>
    Thoughtfully designed · Not legal advice
</div>
""", unsafe_allow_html=True)
