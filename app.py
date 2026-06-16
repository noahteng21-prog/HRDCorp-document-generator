import base64
import io
import os
import re
import zipfile

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# ---------------- CONFIG ----------------

REQUIRED_COLUMNS = [
    "Name of Employer",
    "Name of Trainees",
    "HRDCorp Claim"
]

COLUMN_VARIATIONS = {
    "Name of Employer": [
        "Name of Employer",
        "Name of Employers",
        "Name of Employer(s)",
        "Name of Employers(s)"
    ],
    "Name of Trainees": [
        "Name of Trainees",
        "Name of Trainee",
        "Name of Trainee(s)",
        "Name of Trainees(s)"
    ],
    "HRDCorp Claim": [
        "HRDCorp Claim"
    ]
}

FONT_SIZE = 10

DOCX_TEMPLATE = os.path.join("templates", "New dynamic DOCX PSMB_SBL_KHAS_T3_01 template - Copy.docx")

st.set_page_config(
    page_title="HRDC T3 Generator",
    layout="wide"
)

st.title("HRDC T3 Generator")

course_title = st.text_input("Course Title", "")
training_date = st.text_input("Training Date", "")

st.write(
    "Upload an Excel file to generate T3 attendance lists."
)

# ---------------- FILE UPLOAD ----------------

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is None:
    st.info("Please upload an Excel file.")
    st.stop()

# ---------------- READ EXCEL ----------------

try:
    df = pd.read_excel(
        uploaded_file,
        engine="openpyxl"
    )
    df.columns = df.columns.str.strip()
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

# ---------------- PREVIEW ----------------

st.subheader("Excel Preview")

st.dataframe(df.head())

# ---------------- COLUMN NORMALIZATION ----------------

def find_column(actual_columns, variations):
    for variation in variations:
        if variation in actual_columns:
            return variation
    return None

column_map = {}
missing = []
for canonical, variations in COLUMN_VARIATIONS.items():
    found = find_column(df.columns, variations)
    if found:
        column_map[canonical] = found
    else:
        missing.append(canonical)

if missing:
    st.error(
        f"Missing required columns: {', '.join(missing)}"
    )
    st.stop()

# ---------------- KEEP REQUIRED COLUMNS ----------------

proc_df = df[list(column_map.values())].copy()
proc_df.rename(columns={
    column_map["Name of Employer"]: "Name of Employer",
    column_map["Name of Trainees"]: "Name of Trainees",
    column_map["HRDCorp Claim"]: "HRDCorp Claim"
}, inplace=True)

# ---------------- FILTER HRDC ----------------

proc_df = proc_df[
    proc_df["HRDCorp Claim"]
    .astype(str)
    .str.strip()
    .str.lower()
    == "yes"
]

# ---------------- REMOVE BLANK TRAINEES ----------------

proc_df = proc_df[
    proc_df["Name of Trainees"]
    .notna()
]

proc_df = proc_df[
    proc_df["Name of Trainees"]
    .astype(str)
    .str.strip()
    != ""
]

# ---------------- GROUP BY EMPLOYER ----------------

grouped = (
    proc_df
    .groupby("Name of Employer")
    ["Name of Trainees"]
    .apply(list)
    .reset_index(name="Trainees")
)

grouped["Count"] = grouped[
    "Trainees"
].apply(len)

# ---------------- SUMMARY ----------------

st.subheader("Employers to Generate")

st.dataframe(
    grouped[
        [
            "Name of Employer",
            "Count"
        ]
    ]
)

st.success(
    f"Ready to generate "
    f"{len(grouped)} T3 forms."
)

# ---------------- DEBUG ----------------

with st.expander("View grouped trainees"):

    for _, row in grouped.iterrows():

        st.write(
            f"### {row['Name of Employer']}"
        )

        for trainee in row["Trainees"]:
            st.write(
                f"- {trainee}"
            )


def safe_filename(s):
    s = re.sub(r'[\\/:*?"<>|]', '_', str(s))
    s = s.strip().replace(' ', '_')
    return s


def normalize_docx_placeholders(docx_path):
    placeholder_regex = re.compile(r'\{\{.*?\}\}', flags=re.DOTALL)

    def normalize_match(match):
        block = match.group(0)
        inner = block[2:-2]
        inner = re.sub(r'<[^>]+>', '', inner)
        inner = re.sub(r'\s+', ' ', inner).strip().lower()
        normalized = inner.replace(' ', '_')
        normalized = re.sub(r'_+', '_', normalized)
        return f'{{{{{normalized}}}}}'

    with zipfile.ZipFile(docx_path, 'r') as zin:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    text = data.decode('utf-8', errors='replace')
                    text = placeholder_regex.sub(normalize_match, text)
                    data = text.encode('utf-8')
                zout.writestr(item, data)
        buffer.seek(0)
        return buffer


def expand_docx_participant_rows(template_buffer, target_rows):
    if target_rows <= 0:
        return template_buffer

    if hasattr(template_buffer, "read"):
        template_buffer.seek(0)
        template_data = template_buffer.read()
    else:
        template_data = template_buffer

    with zipfile.ZipFile(io.BytesIO(template_data), "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")

    existing_rows = [int(n) for n in re.findall(r"trainee_(\d+)", doc_xml)]
    current_max = max(existing_rows) if existing_rows else 0
    if current_max >= target_rows:
        return io.BytesIO(template_data)

    row_pattern = re.compile(r"(<w:tr[\s\S]*?<\/w:tr>)")
    rows = row_pattern.findall(doc_xml)
    template_row = next((row for row in rows if "{{trainee_1}}" in row), None)
    if template_row is None:
        return io.BytesIO(template_data)

    placeholder_rows = [row for row in rows if "{{trainee_" in row]
    if placeholder_rows:
        insert_pos = doc_xml.index(placeholder_rows[-1]) + len(placeholder_rows[-1])
    else:
        insert_pos = doc_xml.index(template_row) + len(template_row)

    clones = []
    for index in range(current_max + 1, target_rows + 1):
        clone = template_row
        clone = clone.replace("{{no_1}}", f"{{{{no_{index}}}}}")
        clone = clone.replace("{{trainee_1}}", f"{{{{trainee_{index}}}}}")
        clone = clone.replace("{{employer_1}}", f"{{{{employer_{index}}}}}")
        clones.append(clone)

    doc_xml = doc_xml[:insert_pos] + "".join(clones) + doc_xml[insert_pos:]

    output_buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_data), "r") as zin, zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = doc_xml.encode("utf-8")
            zout.writestr(item, data)
    output_buffer.seek(0)
    return output_buffer


buffers = {}

# Render Word documents from the DOCX template (one per employer)
if DOCX_TEMPLATE and os.path.exists(DOCX_TEMPLATE):
    try:
        normalized_template = normalize_docx_placeholders(DOCX_TEMPLATE)
        for _, row in grouped.iterrows():
            employer = row["Name of Employer"]
            raw_trainees = [str(t) for t in row["Trainees"]]
            target_rows = max(len(raw_trainees), 6)

            expanded_template = expand_docx_participant_rows(
                normalized_template,
                target_rows
            )

            participants = []
            context = {
                "course_title": course_title,
                "training_date": training_date,
                "participants": participants,
            }

            for idx in range(1, target_rows + 1):
                context[f"no_{idx}"] = str(idx)
                if idx <= len(raw_trainees):
                    context[f"trainee_{idx}"] = raw_trainees[idx - 1]
                    context[f"employer_{idx}"] = employer
                    participants.append({
                        "no": idx,
                        "trainee": raw_trainees[idx - 1],
                        "employer": employer,
                    })
                else:
                    context[f"trainee_{idx}"] = ""
                    context[f"employer_{idx}"] = ""

            tpl = DocxTemplate(expanded_template)
            tpl.render(context)

            out = io.BytesIO()
            tpl.save(out)
            out.seek(0)

            buffers[f"T3_{safe_filename(employer)}.docx"] = out.read()
    except Exception as e:
        st.warning(f"Word rendering skipped or failed: {e}")

# (DOCX rendering already performed above)

# Per-employer Word download buttons
st.subheader("Download Word T3 per employer")
st.markdown(
    """
    <style>
    section[data-testid="stHorizontalBlock"] {
        border: 1px solid #d3d3d3 !important;
        border-radius: 8px;
        padding: 0.65rem 0.9rem;
        margin-bottom: 0.35rem;
        box-shadow: rgba(0, 0, 0, 0.03) 0px 1px 2px;
    }
    section[data-testid="stHorizontalBlock"] > div {
        gap: 0.75rem;
        align-items: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

header_col1, header_col2 = st.columns([1, 1])
with header_col1:
    st.markdown("**Employer**")
with header_col2:
    st.markdown("**Action**")

for _, row in grouped.iterrows():
    employer = row["Name of Employer"]
    filename = f"T3_{safe_filename(employer)}.docx"
    data = buffers.get(filename)
    if not data:
        continue

    col1, col2 = st.columns([1, 1])
    with col1:
        st.write(employer)
    with col2:
        st.download_button(
            label="Download",
            data=data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl_docx_{filename}",
            use_container_width=False
        )

# Zip all generated files (DOCX) and provide a red download button
zip_buffer = io.BytesIO()

with zipfile.ZipFile(
    zip_buffer,
    "w",
    zipfile.ZIP_DEFLATED
) as zf:
    for filename, data in buffers.items():
        zf.writestr(filename, data)

zip_buffer.seek(0)

st.markdown("---")

st.markdown(
    "<style>div.stDownloadButton>button {background-color: #dc3545; border-color: #dc3545; color: white;}</style>",
    unsafe_allow_html=True,
)

left_col, _ = st.columns([1, 3])
with left_col:
    st.download_button(
        "Download All T3 DOCX",
        data=zip_buffer,
        file_name="T3_Docx.zip",
        mime="application/zip",
        key="dl_all_docx",
        type="primary",
        use_container_width=False
    )
