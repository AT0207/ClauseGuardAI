
import io
import re
import json
import hashlib
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

st.set_page_config(
    page_title="ClauseGuard AI",
    page_icon="🧑‍⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# DESIGN

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Serif+Display&display=swap');

:root {
    color-scheme: light;
}

.stApp {
    background: #F9FAF6;
    color: #354332;
}

.block-container {
    max-width: 1080px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}

html, body, p, div, button, input, textarea, label {
    font-family: 'DM Sans', sans-serif;
}

h1, h2, h3 {
    font-family: 'DM Serif Display', serif !important;
    color: #485B3D !important;
    font-weight: 400 !important;
}

[data-testid="stHeader"] {
    background: #F9FAF6;
}

[data-testid="stRadio"] label {
    color: #485B3D;
}

[data-testid="stRadio"] div[role="radiogroup"] {
    background: #EDF1E8;
    border-radius: 14px;
    padding: 9px 12px;
}

.stButton>button[kind="primary"] {
    background: #6B8058;
    color: white;
    border: 1px solid #6B8058;
    border-radius: 12px;
    font-weight: 600;
}

.stButton>button[kind="primary"]:hover {
    background: #4E6440;
    color: white;
    border-color: #4E6440;
}

.stButton>button:not([kind="primary"]),
.stDownloadButton>button,
div[data-testid="stLinkButton"] a {
    background: #FFFFFF;
    color: #4E6440;
    border: 1px solid #D5E0CD;
    border-radius: 12px;
}

.stButton>button:not([kind="primary"]):hover {
    background: #E9EFE3;
    color: #354332;
    border-color: #6B8058;
}

[data-testid="stFileUploader"],
[data-testid="stExpander"],
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E0E9D9;
    border-radius: 14px;
    padding: 10px;
}

[data-testid="stTextArea"] textarea {
    background: #FFFFFF;
    color: #354332;
    border: 1px solid #D5E0CD;
    border-radius: 12px;
}

.brand {
    display: flex;
    align-items: center;
    gap: 14px;
    margin: 8px 0 14px;
}

.brand-icon {
    background: #E8EEDF;
    border-radius: 17px;
    padding: 10px;
    font-size: 30px;
}

.brand-title {
    font-family: 'DM Serif Display', serif;
    font-size: 29px;
    color: #485B3D;
}

.brand-note {
    letter-spacing: 1.8px;
    color: #819276;
    font-size: 10px;
    font-weight: 700;
}

.hero {
    background: linear-gradient(125deg, #FFFFFF, #EDF2E9);
    border: 1px solid #E0E9D9;
    border-radius: 24px;
    padding: 48px 45px;
    margin: 25px 0 24px;
}

.eyebrow {
    color: #75896A;
    letter-spacing: 2px;
    font-weight: 700;
    font-size: 11px;
}

.hero-title {
    font-family: 'DM Serif Display', serif;
    font-size: clamp(40px, 5vw, 64px);
    line-height: 1.12;
    color: #485B3D;
    margin: 18px 0;
}

.hero-desc {
    max-width: 670px;
    color: #65775F;
    font-size: 17px;
    line-height: 1.75;
}

.page-label {
    letter-spacing: 1.7px;
    font-size: 11px;
    font-weight: 700;
    color: #7B9172;
    margin-top: 25px;
}

.page-title {
    font-family: 'DM Serif Display', serif;
    font-size: 43px;
    color: #485B3D;
    margin: 10px 0;
}

.page-desc {
    font-size: 16px;
    color: #697D64;
    line-height: 1.7;
    margin-bottom: 23px;
}

.info-card {
    background: #FFFFFF;
    border: 1px solid #E0E9D9;
    border-radius: 18px;
    padding: 24px;
    min-height: 175px;
    margin: 10px 0;
}

.info-icon {
    font-size: 25px;
    margin-bottom: 12px;
}

.info-title {
    color: #485B3D;
    font-weight: 700;
    font-size: 17px;
    margin-bottom: 8px;
}

.info-desc {
    color: #71816A;
    font-size: 14px;
    line-height: 1.65;
}

.finding {
    background: #FFFFFF;
    border: 1px solid #E0E9D9;
    border-left: 5px solid #6B8058;
    border-radius: 13px;
    padding: 18px 22px;
    margin: 18px 0 8px;
}

.finding-title {
    font-size: 18px;
    font-weight: 700;
    color: #485B3D;
}

.risk-pill {
    display: inline-block;
    border-radius: 16px;
    font-size: 11px;
    font-weight: 700;
    padding: 5px 10px;
    margin-top: 9px;
}

.footer {
    border-top: 1px solid #E0E9D9;
    text-align: center;
    color: #829279;
    font-size: 12px;
    margin-top: 50px;
    padding-top: 25px;
    line-height: 1.9;
}
</style>
""", unsafe_allow_html=True)

# CONTRACT ANALYSIS RULES

RULES = {
    "Unlimited liability": (
        r"unlimited liability|without limitation of liability|liability shall not be limited",
        "HIGH",
        "Liability may be uncapped.",
        "Consider a negotiated cap and appropriate exceptions."
    ),
    "Unilateral termination": (
        r"terminate at any time|terminate without notice|terminate for any reason",
        "HIGH",
        "One party may terminate with limited notice or justification.",
        "Review notice, grounds for termination, and reciprocity."
    ),
    "Broad indemnification": (
        r"indemnify[^.\n]*all losses|indemnify[^.\n]*any and all|hold harmless[^.\n]*all claims",
        "HIGH",
        "This indemnity may transfer extensive liabilities.",
        "Review causation, scope, exclusions, and caps."
    ),
    "Automatic renewal": (
        r"automatically renew|automatic renewal|renewed automatically",
        "MEDIUM",
        "The contract may renew without further express consent.",
        "Check the renewal period and cancellation deadline."
    ),
    "Unilateral price changes": (
        r"change prices at any time|increase fees at any time|modify fees without notice",
        "MEDIUM",
        "A party may change pricing unilaterally.",
        "Consider objective adjustment criteria and advance notice."
    ),
    "Exclusive jurisdiction": (
        r"exclusive jurisdiction|courts of [^.\n]* shall have exclusive",
        "REVIEW",
        "Disputes may be confined to specified courts.",
        "Check whether the chosen forum is appropriate."
    ),
    "Non-compete restriction": (
        r"non.compet|shall not compete|prohibited from competing",
        "REVIEW",
        "Commercial activities may be restricted.",
        "Review territorial scope, duration, and enforceability."
    ),
}

CATEGORIES = {
    "Liability": "liability indemnity compensation damages losses",
    "Termination": "termination cancellation expiry notice",
    "Payment": "payment price fees invoice charges",
    "Confidentiality": "confidentiality secrecy disclosure privacy",
    "Dispute resolution": "courts jurisdiction arbitration governing law",
    "Intellectual property": "intellectual property copyright trademark ownership",
}

SAMPLE = """SERVICE AGREEMENT

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
The courts of London shall have exclusive jurisdiction."""

# DOCUMENT PROCESSING

def extract_text(upload):
    data = upload.getvalue()
    name = upload.name.lower()

    if name.endswith(".pdf"):
        return "\n".join(
            page.extract_text() or ""
            for page in PdfReader(io.BytesIO(data)).pages
        )

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)

    return data.decode("utf-8", errors="replace")


def analyse_contract(text):
    clauses = [
        s.strip()
        for s in re.split(r"\n\s*\n", text.strip())
        if s.strip()
    ]

    if not clauses:
        return []

    descriptions = list(CATEGORIES.values())

    matrix = TfidfVectorizer(
        stop_words="english"
    ).fit_transform(descriptions + clauses)

    similarity = cosine_similarity(
        matrix[len(descriptions):],
        matrix[:len(descriptions)]
    )

    category_names = list(CATEGORIES)
    findings = []

    for clause, similarities in zip(clauses, similarity):
        category = (
            category_names[similarities.argmax()]
            if similarities.max() > 0
            else "Other"
        )

        for issue, (pattern, risk, explanation, recommendation) in RULES.items():
            if re.search(pattern, clause, flags=re.I | re.S):
                findings.append({
                    "Issue": issue,
                    "Risk": risk,
                    "Category": category,
                    "Clause": clause,
                    "Explanation": explanation,
                    "Recommendation": recommendation
                })

    return sorted(
        findings,
        key=lambda x: {
            "HIGH": 0,
            "MEDIUM": 1,
            "REVIEW": 2
        }[x["Risk"]]
    )


# HEADER

st.markdown(
    '<div class="brand"><div class="brand-icon">🧑‍⚖️</div><div><div class="brand-title">ClauseGuard AI</div><div class="brand-note">LEGAL INTELLIGENCE, SIMPLIFIED</div></div></div>',
    unsafe_allow_html=True
)

# NAVIGATION

if "page" not in st.session_state:
    st.session_state.page = "Home"

page = st.radio(
    "Site navigation",
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

# PAGE 1 — HOME

if page == "Home":

    st.markdown(
        '<div class="hero"><div class="eyebrow">THOUGHTFUL CONTRACT INTELLIGENCE</div><div class="hero-title">Clarity in every clause.</div><div class="hero-desc">Contracts should inspire confidence, not confusion. ClauseGuard AI helps you identify potential contractual risks, understand important provisions and approach every agreement with greater clarity.</div></div>',
        unsafe_allow_html=True
    )

    if st.button(
        "Begin your contract review →",
        type="primary",
        use_container_width=True
    ):
        st.session_state.page = "Analyse a Contract"
        st.rerun()

    st.markdown(
        '<div class="page-label">WHAT WE OFFER</div><div class="page-title">A clearer view of your contracts.</div>',
        unsafe_allow_html=True
    )

    cols = st.columns(3)

    content = [
        (
            "📄",
            "Understand",
            "Upload your contract or try our sample to explore its provisions."
        ),
        (
            "🌿",
            "Identify",
            "Spot selected risk patterns with clear, explainable screening."
        ),
        (
            "✦",
            "Review",
            "Read the relevant wording and export a report of findings."
        )
    ]

    for col, (symbol, title, description) in zip(cols, content):
        with col:
            st.markdown(
                f'<div class="info-card"><div class="info-icon">{symbol}</div><div class="info-title">{title}</div><div class="info-desc">{description}</div></div>',
                unsafe_allow_html=True
            )

    st.markdown(
        '<div class="page-label">THE PROCESS</div><div class="page-title">Simple by design.</div>',
        unsafe_allow_html=True
    )

    st.write(
        "**01 · Add an agreement** — Upload a PDF, DOCX, TXT file, "
        "paste text, or select the sample."
    )

    st.write(
        "**02 · Analyse its language** — ClauseGuard categorises "
        "text and looks for selected legal-risk expressions."
    )

    st.write(
        "**03 · Review the findings** — Read explanations "
        "and download the results."
    )

    st.info(
        "Educational prototype. This tool does not provide "
        "legal advice or determine whether a contract is legally safe."
    )

# PAGE 2 — CONTRACT ANALYSIS

elif page == "Analyse a Contract":

    st.markdown(
        '<div class="page-label">CONTRACT INTELLIGENCE</div><div class="page-title">Your contract, made clearer.</div><div class="page-desc">Choose how to provide a contract, then review any flagged provisions.</div>',
        unsafe_allow_html=True
    )

    method = st.radio(
        "How would you like to begin?",
        ["Upload a contract", "Try a sample", "Paste text"],
        horizontal=True
    )

    text = ""

    if method == "Upload a contract":

        uploaded = st.file_uploader(
            "Choose a PDF, DOCX or TXT document",
            type=["pdf", "docx", "txt"]
        )

        if uploaded is not None:
            try:
                text = extract_text(uploaded)

                if text.strip():
                    st.success("Document loaded successfully.")
                else:
                    st.warning(
                        "No readable text was found. "
                        "Scanned PDFs may need OCR."
                    )

            except Exception as error:
                st.error(f"Could not read the document: {error}")

    elif method == "Try a sample":

        text = SAMPLE
        st.caption("Fictional service agreement for demonstration.")

        with st.expander("Read the sample"):
            st.text(SAMPLE)

    else:

        text = st.text_area(
            "Paste your agreement here",
            height=260
        )

    digest = hashlib.sha256(
        text.encode("utf-8")
    ).hexdigest()

    if st.button(
        "Analyse my contract →",
        type="primary",
        use_container_width=True
    ):

        if not text.strip():
            st.warning("Please provide contract text first.")

        else:
            with st.spinner("Analysing the agreement…"):
                st.session_state.findings = analyse_contract(text)
                st.session_state.analysed_digest = digest

    if (
        text.strip()
        and st.session_state.get("analysed_digest") == digest
    ):

        findings = st.session_state.findings

        st.divider()

        st.markdown(
            '<div class="page-label">YOUR RESULTS</div><div class="page-title">Contract review report.</div><div class="page-desc">Potential issues that merit consideration. Ratings are indicative, not legal conclusions.</div>',
            unsafe_allow_html=True
        )

        high = sum(
            item["Risk"] == "HIGH" for item in findings
        )

        medium = sum(
            item["Risk"] == "MEDIUM" for item in findings
        )

        review = sum(
            item["Risk"] == "REVIEW" for item in findings
        )

        for col, label, value in zip(
            st.columns(4),
            [
                "Total findings",
                "High priority",
                "Medium priority",
                "For review"
            ],
            [len(findings), high, medium, review]
        ):
            col.metric(label, value)

        if findings:

            st.subheader("Risk overview")

            chart = pd.DataFrame({
                "Priority": ["High", "Medium", "Review"],
                "Findings": [high, medium, review]
            }).set_index("Priority")

            st.bar_chart(chart, color="#80946D")

            st.subheader("Detailed findings")

            for i, finding in enumerate(findings, 1):

                risk = finding["Risk"]

                accent, bg = {
                    "HIGH": ("#B66E65", "#F9ECEA"),
                    "MEDIUM": ("#A68D56", "#F8F2E4"),
                    "REVIEW": ("#647F58", "#E9F0E1")
                }[risk]

                st.markdown(
                    f'<div class="finding" style="border-left-color:{accent}"><div class="finding-title">{i:02d} · {finding["Issue"]}</div><span class="risk-pill" style="color:{accent};background:{bg}">{risk} PRIORITY</span></div>',
                    unsafe_allow_html=True
                )

                with st.expander(
                    "Read the clause and explanation",
                    expanded=True
                ):

                    st.write("**Relevant wording**")
                    st.info(finding["Clause"])

                    st.write("**Why it matters**")
                    st.write(finding["Explanation"])

                    st.write("**Suggested review**")
                    st.write(finding["Recommendation"])

            st.subheader("Download your report")

            c1, c2 = st.columns(2)

            c1.download_button(
                "Download CSV report",
                pd.DataFrame(findings).to_csv(
                    index=False
                ).encode("utf-8"),
                "ClauseGuard_Report.csv",
                "text/csv",
                use_container_width=True
            )

            c2.download_button(
                "Download JSON report",
                json.dumps(findings, indent=2),
                "ClauseGuard_Report.json",
                "application/json",
                use_container_width=True
            )

        else:
            st.success("No predefined patterns were detected.")
            st.warning(
                "This does not mean the contract is risk-free."
            )

    with st.expander("How does the analysis work?"):

        st.write(
            "ClauseGuard uses TF-IDF and cosine similarity to "
            "categorise paragraphs, then transparent pattern-matching "
            "rules to flag selected phrases. It is not a trained legal "
            "risk model. Results may miss issues or flag harmless text. "
            "It does not determine enforceability."
        )

# PAGE 3 — LEGAL NEWS

elif page == "Legal News":

    st.markdown(
        '<div class="page-label">LEGAL INSIGHTS</div><div class="page-title">Stay informed.</div><div class="page-desc">Trusted places to read current legal reporting and primary EU materials.</div>',
        unsafe_allow_html=True
    )

    sources = [
        (
            "Reuters — Legal",
            "Coverage of litigation, regulation and legal business.",
            "https://www.reuters.com/legal/"
        ),
        (
            "The Law Society Gazette",
            "Law and professional practice in England and Wales.",
            "https://www.lawgazette.co.uk/"
        ),
        (
            "Legal Cheek",
            "Legal careers, firms and developments in the profession.",
            "https://www.legalcheek.com/"
        ),
        (
            "EUR-Lex",
            "Official EU legislation, judgments and legal materials.",
            "https://eur-lex.europa.eu/"
        )
    ]

    for title, description, url in sources:

        st.markdown(
            f'<div class="info-card" style="min-height:auto"><div class="info-title">{title}</div><div class="info-desc">{description}</div></div>',
            unsafe_allow_html=True
        )

        st.link_button("Open publication ↗", url)

    st.caption(
        "These are external links, not live news feeds or endorsements."
    )

# PAGE 4 — LEGAL 500

else:

    st.markdown(
        '<div class="page-label">PROFESSIONAL RESOURCES</div><div class="page-title">The Legal 500.</div><div class="page-desc">Explore a directory of law firms and legal practitioners around the world.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="hero"><div class="eyebrow">LEGAL RESEARCH & DIRECTORIES</div><div class="hero-title">Discover legal expertise.</div><div class="hero-desc">The Legal 500 publishes research, rankings and commentary on law firms and legal professionals across jurisdictions and practice areas.</div></div>',
        unsafe_allow_html=True
    )

    st.link_button(
        "Visit The Legal 500 ↗",
        "https://www.legal500.com/",
        use_container_width=True
    )

    st.caption(
        "ClauseGuard AI is independent and not affiliated "
        "with or endorsed by The Legal 500."
    )

# FOOTER

st.markdown(
    '<div class="footer">🧑‍⚖️ ClauseGuard AI · Clarity in every clause.<br>Educational demonstration · Not legal advice</div>',
    unsafe_allow_html=True
)
