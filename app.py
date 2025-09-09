import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from openai import OpenAI

st.set_page_config(page_title="Supply Chain Agent (PoC)", layout="centered")

# ---------- helpers ----------
def get_openai_key():
    try:
        return st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        return None

def read_any_file(upload):
    name = (upload.name or "").lower()
    data = upload.read()
    bio = BytesIO(data)
    if name.endswith(".csv"):
        try:
            df = pd.read_csv(bio)
            return "table", df
        except Exception:
            return "text", data.decode("utf-8", "ignore")
    if name.endswith(".xlsx"):
        try:
            df = pd.read_excel(bio)
            return "table", df
        except Exception:
            return "text", data.decode("utf-8", "ignore")
    if name.endswith(".txt"):
        return "text", data.decode("utf-8", "ignore")
    if name.endswith(".docx"):
        doc = Document(bio)
        txt = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return "text", txt
    return "text", data.decode("utf-8", "ignore")

def to_docx_bytes(text: str) -> bytes:
    doc = Document()
    for line in (text or "").splitlines():
        if line.startswith("## "):
            doc.add_heading(line.replace("## ", ""), level=2)
        elif line.startswith("### "):
            doc.add_heading(line.replace("### ", ""), level=3)
        elif line.startswith("- "):
            p = doc.add_paragraph(line[2:])
            try:
                p.style = "List Bullet"
            except Exception:
                pass
        else:
            doc.add_paragraph(line)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()

def polish_with_openai(notes, file_text, focus):
    key = get_openai_key()
    if not key:
        return None, "No OpenAI API key found"
    client = OpenAI(api_key=key)

    prompt = f"""
You are a senior supply chain consultant.
Write clear project documentation in Markdown.

Sections:
## A) Project Name Ideas
## B) 2–3 Line Description
## C) Detailed Overview
## D) Quick Insights
## E) Additional Analysis Focus

=== INPUT START ===
Notes:
{notes}

Uploaded Files (combined text or table samples):
{file_text or "(none)"}

Focus:
{focus or "(none)"}
=== INPUT END ===
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=900,
            temperature=0.2,
        )
        return resp.choices[0].message.content, None
    except Exception as e:
        return None, str(e)

# ---------- UI ----------
st.sidebar.write("🔑 OpenAI Key:", "Found ✅" if get_openai_key() else "Missing ❌")

st.title("Supply Chain Agent — Minimal Demo (multi-file + Word export)")

# Multiple files
uploads = st.file_uploader(
    "Upload files (CSV/XLSX/TXT/DOCX) — multiple allowed",
    type=["csv", "xlsx", "txt", "docx"],
    accept_multiple_files=True
)

file_text = ""
if uploads:
    parts = []
    for up in uploads:
        kind, content = read_any_file(up)
        if kind == "table":
            st.write(f"Table: {up.name}")
            st.dataframe(content.head())
            parts.append(content.head(20).to_csv(index=False))
        else:
            st.write(f"Text: {up.name}")
            preview = str(content)[:600]
            st.text_area(f"Preview: {up.name}", preview, height=120)
            parts.append(str(content))
    file_text = "\n\n".join(parts)

with st.form("input_form"):
    notes = st.text_area("Stakeholder notes", "Example: Stakeholder wants analytics MVP.", height=120)
    focus = st.text_area("Additional focus (optional)", "", height=80)
    submitted = st.form_submit_button("Generate Project Draft")

if submitted:
    st.info("Generating with OpenAI…")
    polished, err = polish_with_openai(notes, file_text, focus)
    if polished:
        st.subheader("🔎 Polished Draft")
        st.markdown(polished)

        # Word (.docx) download
        docx_bytes = to_docx_bytes(polished)
        st.download_button(
            "⬇️ Download as Word (.docx)",
            data=docx_bytes,
            file_name="project.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Optional Markdown download
        st.download_button(
            "⬇️ Download as Markdown (.md)",
            data=polished.encode("utf-8"),
            file_name="project.md",
            mime="text/markdown"
        )
    else:
        st.error(err)

