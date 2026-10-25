
import re
import io
import json
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

st.set_page_config(page_title="ClauseGuard AI", page_icon="⚖️", layout="wide")

RULES = {
    "Unlimited liability": {
        "pattern": r"unlimited liability|without limitation of liability|liability shall not be limited",
        "risk": "HIGH",
        "explanation": "Liability may be uncapped, potentially exposing a party to substantial financial loss.",
        "suggestion": "Consider a negotiated financial liability cap."
    },
    "Unilateral termination": {
        "pattern": r"terminate at any time|terminate without notice|terminate for any reason",
        "risk": "HIGH",
        "explanation": "A party may be entitled to terminate without meaningful notice or justification.",
        "suggestion": "Consider reciprocal termination rights and reasonable notice."
    },
    "Automatic renewal": {
        "pattern": r"automatically renew|automatic renewal|renewed automatically",
        "risk": "MEDIUM",
        "explanation": "The contract may renew without further affirmative agreement.",
        "suggestion": "Check renewal periods, notification requirements and cancellation rights."
    },
    "Broad indemnity": {
        "pattern": r"indemnify.*all losses|indemnify.*any and all|hold harmless.*all claims",
        "risk": "HIGH",
        "explanation": "The indemnity may transfer extensive liabilities to one party.",
        "suggestion": "Review scope, exclusions, causation and any financial caps."
    },
    "Unilateral price changes": {
        "pattern": r"change prices at any time|increase fees at any time|modify fees without notice",
        "risk": "MEDIUM",
        "explanation": "One party may be able to change the agreed price unilaterally.",
        "suggestion": "Consider advance notice and objective price-adjustment criteria."
    },
    "Exclusive jurisdiction": {
        "pattern": r"exclusive jurisdiction|courts of .* shall have exclusive",
        "risk": "REVIEW",
        "explanation": "Disputes may have to be brought before a specified court.",
        "suggestion": "Verify that the chosen forum is appropriate and commercially acceptable."
    },
    "Non-compete restriction": {
        "pattern": r"non.compet|shall not compete|prohibited from competing",
        "risk": "REVIEW",
        "explanation": "The clause may restrict future commercial activities.",
        "suggestion": "Review duration, geographical scope and enforceability."
    }
}

CATEGORIES = {
    "Liability": "Limitation of liability damages losses financial exposure",
    "Termination": "Termination cancellation notice expiry ending agreement",
    "Payment": "Payment fees invoices interest price financial obligations",
    "Confidentiality": "Confidentiality secrets private information disclosure",
    "Dispute resolution": "Jurisdiction courts arbitration governing law disputes",
    "Intellectual property": "Copyright ownership trademarks intellectual property",
    "Other": "General miscellaneous contractual provision"
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

def extract_text(uploaded):
    data = uploaded.getvalue()
    name = uploaded.name.lower()

    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)

    return data.decode("utf-8", errors="replace")


def split_clauses(text):
    # Divide text into paragraphs, preserving complete provisions.
    chunks = re.split(r"\n\s*\n", text.strip())
    return [c.strip() for c in chunks if c.strip()]


def categorise(clauses):
    descriptions = list(CATEGORIES.values())

    vectorizer = TfidfVectorizer(stop_words="english")
    matrix = vectorizer.fit_transform(descriptions + clauses)
    scores = cosine_similarity(matrix[len(descriptions):],
                               matrix[:len(descriptions)])

    names = list(CATEGORIES.keys())
    return [
        names[row.argmax()] if row.max() > 0 else "Other"
        for row in scores
    ]


def analyse(text):
    clauses = split_clauses(text)
    if not clauses:
        return []

    categories = categorise(clauses)
    results = []

    for clause, category in zip(clauses, categories):
        for issue, rule in RULES.items():
            if re.search(rule["pattern"], clause, re.I | re.S):
                results.append({
                    "Issue": issue,
                    "Risk": rule["risk"],
                    "Category": category,
                    "Clause": clause,
                    "Explanation": rule["explanation"],
                    "Recommendation": rule["suggestion"]
                })

    return results


st.title("⚖️ ClauseGuard AI")
st.subheader("Intelligent Contract Risk Detection")
st.write(
    "Upload a contract to identify potentially problematic provisions "
    "and obtain explainable risk assessments."
)

st.info(
    "Educational prototype. It highlights issues for legal review "
    "but does not determine enforceability or replace legal advice."
)

source = st.radio(
    "Choose your input:",
    ["Try sample contract", "Upload contract", "Paste contract text"]
)

text = ""

if source == "Try sample contract":
    text = SAMPLE
    st.text_area("Sample contract", text, height=220)

elif source == "Upload contract":
    uploaded = st.file_uploader(
        "Upload a PDF, DOCX or TXT contract",
        type=["pdf", "docx", "txt"]
    )
    if uploaded:
        try:
            text = extract_text(uploaded)
        except Exception as e:
            st.error(f"Unable to read this file: {e}")

else:
    text = st.text_area("Paste your contract", height=250)

if st.button("🔍 Analyse Contract", type="primary"):
    if not text.strip():
        st.warning("Please provide contract text.")
    else:
        findings = analyse(text)
        st.session_state["findings"] = findings
        st.session_state["analysed"] = True

if st.session_state.get("analysed"):
    findings = st.session_state.get("findings", [])

    st.divider()
    st.header("Contract Risk Report")

    high = sum(x["Risk"] == "HIGH" for x in findings)
    medium = sum(x["Risk"] == "MEDIUM" for x in findings)
    review = sum(x["Risk"] == "REVIEW" for x in findings)

    a, b, c, d = st.columns(4)
    a.metric("Total flags", len(findings))
    b.metric("High risk", high)
    c.metric("Medium risk", medium)
    d.metric("For review", review)

    if findings:
        st.bar_chart(pd.DataFrame(findings)["Risk"].value_counts())

        for i, finding in enumerate(findings, 1):
            with st.expander(
                f"{i}. {finding['Issue']} — {finding['Risk']}",
                expanded=True
            ):
                st.write("**Clause:**", finding["Clause"])
                st.write("**Explanation:**", finding["Explanation"])
                st.write("**Recommendation:**", finding["Recommendation"])

        df = pd.DataFrame(findings)

        st.download_button(
            "Download CSV Report",
            df.to_csv(index=False).encode("utf-8"),
            "ClauseGuard_Report.csv",
            "text/csv"
        )

        st.download_button(
            "Download JSON Report",
            json.dumps(findings, indent=2),
            "ClauseGuard_Report.json",
            "application/json"
        )
    else:
        st.success("No predefined risk patterns detected.")
        st.warning("This does not establish that the contract is risk-free.")

st.divider()
st.caption(
    "ClauseGuard AI | NLP-based classification and explainable "
    "rule-based risk screening | Academic demonstration"
)
