import io
import json
import pandas as pd
import streamlit as st
from docx import Document
from pypdf import PdfReader
from engine import analyse, risk_summary

st.set_page_config(page_title='ClauseGuard AI',page_icon='🛡️',layout='wide')
st.title('🛡️ ClauseGuard AI')
st.caption('Explainable contract issue spotting • educational prototype • English-language contracts')
st.warning('Flags are heuristic review prompts, **not** legal conclusions, legal advice, or a substitute for qualified counsel. No flag does not establish that a contract is safe.')

@st.cache_data
def sample_text():
    with open('data/sample_contract.txt',encoding='utf-8') as f:
        return f.read()

def read_file(f):
    raw=f.getvalue()
    if len(raw)>5_000_000:
        raise ValueError('File exceeds 5 MB limit')
    ext=f.name.rsplit('.',1)[-1].lower()
    if ext=='txt':
        return raw.decode('utf-8-sig')
    if ext=='docx':
        doc=Document(io.BytesIO(raw))
        return '\n\n'.join(p.text for p in doc.paragraphs if p.text.strip())
    if ext=='pdf':
        return '\n\n'.join((page.extract_text() or '') for page in PdfReader(io.BytesIO(raw)).pages)
    raise ValueError('Only TXT, DOCX and text-based PDF are supported')

with st.sidebar:
    st.header('Input')
    mode=st.radio('Contract source',['Demo contract','Upload document','Paste text'])
    st.info('For confidentiality, run locally and use synthetic or authorised documents. Uploaded files are processed in the current Streamlit process; hosting policies vary.')
    st.markdown('**Method:** TF-IDF categorisation + auditable pattern-based flags. No remote AI API or account required.')

if mode=='Demo contract':
    text=sample_text()
elif mode=='Upload document':
    up=st.file_uploader('Choose a contract',type=['txt','docx','pdf'])
    if up is None:
        st.stop()
    try:
        text=read_file(up)
    except Exception as exc:
        st.error(f'Could not read file: {exc}')
        st.stop()
else:
    text=st.text_area('Paste contract',height=230,placeholder='Insert English contract text...')

if not text.strip():
    st.warning('No extractable text found. Scanned PDFs require OCR, which is not included.')
    st.stop()

clauses, findings=analyse(text)
if not clauses:
    st.warning('No substantial clauses detected. Separate clauses using blank lines.')
    st.stop()

counts=risk_summary(findings)
a,b,c,d=st.columns(4)
a.metric('Clauses',len(clauses))
b.metric('Review flags',len(findings))
c.metric('High-priority flags',counts['High'])
d.metric('Medium-priority flags',counts['Medium'])
st.caption('Counts indicate pattern matches only; they are not probabilities, validated risk ratings or legal conclusions.')

tab1,tab2,tab3=st.tabs(['🚩 Review flags','📄 Clause map','📊 Export'])
with tab1:
    if not findings:
        st.success('No configured patterns matched. This is NOT a clean legal bill of health.')
    for f in findings:
        with st.expander(f"{f['severity']} | {f['clause']} | {f['issue']}",expanded=True):
            st.write('**Category:**',f['category'])
            st.write('**Why it deserves review:**',f['explanation'])
            st.write('**Matched language:**',f['evidence'])
            st.write('**Question for counsel:**',f['review_question'])
            st.caption(f"Transparent detection rule: {f['rule']}")
with tab2:
    st.dataframe(pd.DataFrame([{'Clause':x['number'],'Predicted topic':x['category'],'TF-IDF similarity':x['similarity'],'Text':x['text']} for x in clauses]),use_container_width=True,hide_index=True)
    st.caption('The similarity score measures lexical overlap with curated category descriptions. It is not a confidence or legal risk score.')
with tab3:
    report={'tool':'ClauseGuard AI v1.0','method':'TF-IDF + heuristic regex','disclaimer':'Educational issue spotter, not legal advice','clauses':clauses,'findings':findings}
    st.download_button('Download JSON report',json.dumps(report,indent=2,ensure_ascii=False),file_name='clauseguard_report.json',mime='application/json')
    st.download_button('Download findings CSV',pd.DataFrame(findings,columns=['clause','category','severity','issue','explanation','review_question','rule','evidence','similarity']).to_csv(index=False),file_name='clauseguard_findings.csv',mime='text/csv')
