"""JD14 source-document extraction, review, and PDF generation."""

from __future__ import annotations

import io
import hashlib
import json
import os
import re
import urllib.error
import urllib.request
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Tuple

import fitz
import streamlit as st


TEMPLATE_PATH = Path(__file__).parent / "templates" / "JD14 template 2026 (2).pdf"
EXTRACTION_VERSION = 6

FIELD_LABELS = {
    "employer_name_address": "Registered Name and Address of Employer",
    "employer_code": "Employer Code",
    "approval_no": "Approval No",
    "course_title": "Course Title",
    "training_start_date": "Training commenced (DD/MM/YYYY)",
    "training_end_date": "Training ended (DD/MM/YYYY)",
    "training_venue": "Training Venue",
    "number_of_trainees": "Number of Trainees",
    "total_fee_approved": "Total Fee Approved (RM)",
    "total_fee_claimed": "Total Fee Claimed (RM)",
}

FIELD_ORDER = list(FIELD_LABELS)
DATE_FIELDS = {"training_start_date", "training_end_date"}
MONEY_FIELDS = {"total_fee_approved", "total_fee_claimed"}


def extract_pdf_text(uploaded_pdf) -> Tuple[str, List[str]]:
    """Extract text locally, attempting PyMuPDF's Tesseract OCR for image pages."""
    warnings: List[str] = []
    document = fitz.open(stream=uploaded_pdf.getvalue(), filetype="pdf")
    if document.is_encrypted:
        raise ValueError("This PDF is password-protected. Please upload an unlocked copy.")
    page_text: List[str] = []
    for page_number, page in enumerate(document, start=1):
        text = page.get_text("text").strip()
        if not text:
            try:
                text_page = page.get_textpage_ocr(language="eng", dpi=200)
                text = page.get_text("text", textpage=text_page).strip()
            except Exception:
                warnings.append(
                    "Page {} appears scanned. OCR could not run; install Tesseract with English language data to enable it.".format(page_number)
                )
        if text:
            page_text.append(text)
    document.close()
    return "\n\n".join(page_text), warnings


def _first(pattern: str, text: str, flags=re.IGNORECASE | re.MULTILINE) -> str:
    match = re.search(pattern, text, flags)
    return match.group(1).strip(" \t:;\r") if match else ""


def _date_to_display(value: str) -> str:
    value = value.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return ""


def _compact_employer_address(value: str) -> str:
    """Keep the employer name on line one and group its address into three lines."""
    parts = [part.strip(" ,\t") for part in value.splitlines() if part.strip(" ,\t")]
    if len(parts) <= 1:
        return "\n".join(parts).upper()
    name, address = parts[0], parts[1:]
    address_line_count = min(3, len(address))
    compacted = []
    offset = 0
    for line_index in range(address_line_count):
        lines_left = address_line_count - line_index
        items_left = len(address) - offset
        take = (items_left + lines_left - 1) // lines_left
        compacted.append(", ".join(address[offset : offset + take]))
        offset += take
    return "\n".join([name] + compacted).upper()


def extract_candidates(text: str) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Map common HRD Corp approval letter labels into editable JD14 values."""
    values = {key: "" for key in FIELD_ORDER}
    notes: Dict[str, str] = {}
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    salutation_index = next(
        (i for i, line in enumerate(lines) if re.search(r"\b(?:Sir\s*/\s*Madam|Dear)\b", line, re.I)),
        len(lines),
    )
    # Approval letters identify HRD Corp in the page header before the actual
    # employer's address. Use the last company-name match before the salutation.
    company_candidates = [
        i
        for i, line in enumerate(lines[:salutation_index])
        if re.search(r"\b(?:SDN\.?\s*BHD\.?|BERHAD)\b", line, re.I)
    ]
    company_index = company_candidates[-1] if company_candidates else None
    if company_index is not None:
        address_lines = [lines[company_index]]
        for line in lines[company_index + 1 : salutation_index]:
            if re.search(r"\bTRAINING GRANT APPLICATION\b", line, re.I):
                break
            if re.search(r"\b(?:PROGRAMME NAME|DATE OF PROGRAM|APPLICATION NUMBER)\b", line, re.I):
                break
            address_lines.append(line)
        values["employer_name_address"] = _compact_employer_address("\n".join(address_lines))

    values["approval_no"] = _first(
        r"(?:APPLICATION\s*(?:NUMBER|NO\.?|#)|APPROVAL\s*(?:NUMBER|NO\.?))\s*:?\s*([^\n]+)", text
    )
    if "_" in values["approval_no"]:
        values["employer_code"] = values["approval_no"].split("_", 1)[0].strip()
        notes["employer_code"] = "Derived from the approval number prefix before the first underscore; confirm it."
    else:
        notes["employer_code"] = "Not derived: the approval number has no underscore. Enter the Employer Code directly."
    values["course_title"] = _first(r"(?:PROGRAMME|COURSE)\s*NAME\s*:?\s*([^\n]+)", text)

    date_match = re.search(
        r"(?:DATE\s*OF\s*PROGRAM|TRAINING\s*DATE)[^\n]*?From\s*:?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})[^\n]*?To\s*:?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        text,
        re.I,
    )
    if date_match:
        values["training_start_date"] = _date_to_display(date_match.group(1))
        values["training_end_date"] = _date_to_display(date_match.group(2))

    # A course-fee line is more specific than the grant's total, which may include allowances.
    fee_match = re.search(
        r"Course\s+Fee\s+\d+\s+([\d,]+\.\d{2})\s+[\d.]+\s+([\d,]+\.\d{2})",
        text,
        re.I,
    )
    if fee_match:
        values["number_of_trainees"] = str(int(re.search(r"Course\s+Fee\s+(\d+)", text, re.I).group(1)))
        values["total_fee_approved"] = fee_match.group(2).replace(",", "")
        notes["total_fee_approved"] = "Matched to the Course Fee row in the approval schedule."
    else:
        trainee_match = re.search(r"(?:Number\s+of\s+(?:Trainees|People)|Trainees?)\s*:?\s*(\d+)\b", text, re.I)
        if trainee_match:
            values["number_of_trainees"] = trainee_match.group(1)

    normalized_text = re.sub(r"\s+", " ", text)
    upfront_payment_match = re.search(
        r"\bUPFRONT\s+PAYMENT(?:\s+TO\s+TRAINING\s+PROVIDER)?\b[^\d]{0,80}([\d,]+(?:\.\d{1,2})?)",
        normalized_text,
        re.I,
    )
    claimable_amount_patterns = (
        r"\bBALANCE\s+AMOUNT\s+TO\s+BE\s+CLAIMED(?:\s+BY\s+TRAINING\s+PROVIDER)?\b[^\d]{0,80}([\d,]+(?:\.\d{1,2})?)",
        r"\b(?:TOTAL\s+)?AMOUNT\s+(?:TO\s+BE\s+CLAIMED|CLAIMABLE)(?:\s+BY\s+TRAINING\s+PROVIDER)?\b[^\d]{0,80}([\d,]+(?:\.\d{1,2})?)",
        r"\b(?:TOTAL\s+)?CLAIMABLE\s+AMOUNT\b[^\d]{0,80}([\d,]+(?:\.\d{1,2})?)",
    )
    for pattern in claimable_amount_patterns:
        claimable_match = re.search(pattern, normalized_text, re.I)
        if claimable_match:
            values["total_fee_claimed"] = claimable_match.group(1).replace(",", "")
            notes["total_fee_claimed"] = "Matched to the explicit balance/claimable amount in the source; confirm before generating."
            break

    if "total_fee_claimed" not in notes and upfront_payment_match and values["total_fee_approved"]:
        upfront_payment = Decimal(upfront_payment_match.group(1).replace(",", ""))
        approved_course_fee = Decimal(values["total_fee_approved"])
        if Decimal(0) <= upfront_payment <= approved_course_fee:
            remaining_balance = approved_course_fee - upfront_payment
            values["total_fee_claimed"] = format(remaining_balance, "f")
            notes["total_fee_claimed"] = (
                "Calculated as the approved Course Fee minus the upfront payment; confirm against the source."
            )

    if "training_venue" not in notes:
        notes["training_venue"] = "Not confidently identified in the uploaded PDF; enter or confirm this value."
    if "total_fee_claimed" not in notes:
        notes["total_fee_claimed"] = "No explicit claimable balance found; enter or confirm this value."
    if values["approval_no"]:
        notes["approval_no"] = "Read from the source application/approval number; confirm it matches the JD14 Approval No."
    return values, notes


def extract_with_ai(text: str) -> Dict[str, str]:
    """Use OpenAI structured output to map the source PDF text into JD14 fields."""
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise ValueError("OPENAI_API_KEY is not configured. Set it in the local environment to enable AI extraction.")

    properties = {key: {"type": ["string", "null"]} for key in FIELD_ORDER}
    schema = {
        "type": "object",
        "properties": properties,
        "required": FIELD_ORDER,
        "additionalProperties": False,
    }
    prompt = (
        "Extract only facts supported by the supplied PDF text for the HRD Corp JD14 form. "
        "Return dates as DD/MM/YYYY and monetary values as plain numeric strings. "
        "Employer Code is the approval/application number prefix before its first underscore. "
        "For total_fee_approved, use the Course Fee approved amount only, excluding trainee allowances. "
        "For total_fee_claimed, use an explicit balance/claimable amount when stated. Otherwise, "
        "if an upfront payment is stated, subtract it from the approved Course Fee. If neither is stated, use null. "
        "Do not mistake HRD Corp's government address for the employer address. Use null for unsupported or absent fields."
    )
    body = {
        "model": os.environ.get("OPENAI_MODEL", "gpt-4o-mini"),
        "messages": [
            {"role": "system", "content": prompt},
            {"role": "user", "content": "Source PDF text:\n" + text[:60000]},
        ],
        "response_format": {
            "type": "json_schema",
            "json_schema": {"name": "jd14_extraction", "strict": True, "schema": schema},
        },
        "temperature": 0,
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": "Bearer " + api_key, "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise ValueError("OpenAI extraction failed (HTTP {}): {}".format(exc.code, detail)) from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise ValueError("Could not connect to OpenAI for extraction: {}".format(exc)) from exc

    content = payload["choices"][0]["message"]["content"]
    extracted = json.loads(content)
    return {key: str(extracted[key]).strip() if extracted.get(key) is not None else "" for key in FIELD_ORDER}


def validate_values(values: Dict[str, str]) -> List[str]:
    errors: List[str] = []
    for key, label in FIELD_LABELS.items():
        value = (values.get(key) or "").strip()
        if not value:
            errors.append("{} is required.".format(label))
            continue
        if key in DATE_FIELDS:
            try:
                datetime.strptime(value, "%d/%m/%Y")
            except ValueError:
                errors.append("{} must use DD/MM/YYYY.".format(label))
        elif key == "number_of_trainees":
            if not value.isdigit() or int(value) <= 0:
                errors.append("Number of Trainees must be a positive whole number.")
        elif key in MONEY_FIELDS:
            try:
                amount = Decimal(value.replace(",", "").replace("RM", "").strip())
                if amount < 0 or amount.as_tuple().exponent < -2:
                    raise InvalidOperation()
            except (InvalidOperation, ValueError):
                errors.append("{} must be a non-negative amount with up to two decimal places.".format(label))
    if values.get("training_start_date") and values.get("training_end_date"):
        try:
            start = datetime.strptime(values["training_start_date"], "%d/%m/%Y")
            end = datetime.strptime(values["training_end_date"], "%d/%m/%Y")
            if end < start:
                errors.append("Training ended date must be on or after the commenced date.")
        except ValueError:
            pass
    return errors


def _put(page, rect: Tuple[float, float, float, float], value: str, label: str, fontsize=8.5, align=0):
    if not value:
        return
    x0, baseline_y, x1, bottom_y = rect
    text_rect = fitz.Rect(x0, baseline_y - fontsize * 1.12, x1, bottom_y - fontsize * 0.12)
    current_size = fontsize
    while current_size >= 5:
        result = page.insert_textbox(
            text_rect,
            value,
            fontsize=current_size,
            fontname="helv",
            color=(0, 0, 0),
            align=align,
        )
        if result >= 0:
            return
        current_size -= 0.5
    raise ValueError("{} is too long to fit its field on the JD14 template.".format(label))


def _money_value(value: str) -> Decimal:
    return Decimal(value.upper().replace("RM", "").replace(",", "").strip())


def _fee_display(value: str) -> str:
    amount = _money_value(value)
    if amount == amount.to_integral_value():
        return "{:,.0f}".format(amount)
    return "{:,.2f}".format(amount)


def _company_output_filename(employer_name_address: str) -> str:
    company = next((line.strip() for line in employer_name_address.splitlines() if line.strip()), "")
    safe_company = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "", company).strip().rstrip(" .")
    return (safe_company or "HRDCorp") + "_JD14.pdf"


def generate_pdf(values: Dict[str, str]) -> bytes:
    if not TEMPLATE_PATH.exists():
        raise ValueError("The JD14 template file is missing from the templates folder.")
    document = fitz.open(TEMPLATE_PATH)
    if len(document) != 1:
        document.close()
        raise ValueError("The configured JD14 template must be a one-page PDF.")
    page = document[0]
    # PyMuPDF coordinates are top-left based; these rectangles align to the blank
    # fields in the supplied 595.44 x 841.68 point A4 template. Existing provider
    # identification and declaration entries are deliberately left as printed.
    employer_address = _compact_employer_address(values["employer_name_address"])
    _put(page, (37, 190, 316, 248), employer_address, "Employer name and address", 8)
    _put(page, (416, 174, 555, 185), values["employer_code"], "Employer Code")
    _put(page, (416, 193, 555, 204), values["approval_no"], "Approval No")
    _put(page, (144, 251.6, 555, 263.6), values["course_title"], "Course Title")
    _put(page, (239, 269, 335, 281), values["training_start_date"], "Training commenced date")
    _put(page, (400, 269, 555, 281), values["training_end_date"], "Training ended date")
    # Match the printed Training Venue label's text line while preserving the value as entered.
    _put(page, (144, 290.12, 555, 303.12), values["training_venue"], "Training Venue")
    # Rebuild Part 2's header row with consistent vertical centering and a
    # little more air between the section title and the table.
    header_top, header_bottom = 321.4, 346.5
    col_edges = (31.44, 184.58, 396.07, 562.66)
    page.draw_rect(fitz.Rect(31.8, 318.4, 562.3, 321.5), color=None, fill=(1, 1, 1), overlay=True)
    page.draw_rect(fitz.Rect(31.8, 321.3, 562.3, 346.2), color=None, fill=(1, 1, 1), overlay=True)
    headings = (
        "Number of Trainee(s)*",
        "Total Fee Approved (RM)",
        "Total Fee Claimed (RM)",
    )
    header_font_size = 9.5
    header_font = fitz.Font("hebo")
    header_center_y = (header_top + header_bottom) / 2
    header_baseline = header_center_y + (
        header_font.ascender + header_font.descender
    ) * header_font_size / 2
    for left, right, heading in zip(col_edges[:-1], col_edges[1:], headings):
        text_width = header_font.text_length(heading, fontsize=header_font_size)
        page.insert_text(
            ((left + right - text_width) / 2, header_baseline),
            heading,
            fontsize=header_font_size,
            fontname="hebo",
            color=(0, 0, 0),
            overlay=True,
        )
    page.draw_line((col_edges[0], header_top), (col_edges[-1], header_top), color=(0, 0, 0), width=0.72, overlay=True)
    page.draw_line((col_edges[0], header_bottom), (col_edges[-1], header_bottom), color=(0, 0, 0), width=0.72, overlay=True)
    for x in col_edges:
        page.draw_line((x, header_top), (x, 367.37), color=(0, 0, 0), width=0.72, overlay=True)
    # Center values inside the three actual Part 2 cells (rather than across
    # boxes that extend over the neighboring columns).
    _put(page, (31, 361.5, 185, 380.5), values["number_of_trainees"], "Number of Trainees", 9, align=1)
    _put(page, (185, 361.5, 396, 380.5), _fee_display(values["total_fee_approved"]), "Total Fee Approved", 9, align=1)
    _put(page, (396, 361.5, 563, 380.5), _fee_display(values["total_fee_claimed"]), "Total Fee Claimed", 9, align=1)
    output = io.BytesIO()
    document.save(output, garbage=3, deflate=True)
    document.close()
    return output.getvalue()


@st.dialog("Review extracted fields", width="large")
def _review_jd14_dialog() -> None:
    st.markdown(
        """
        <style>
        div[data-testid="stDialog"] > div { width: 96vw !important; max-width: 96vw !important; }
        div[data-testid="stDialog"] [role="dialog"] { width: 96vw !important; max-width: 96vw !important; max-height: 94vh !important; }
        div[data-testid="stDialog"] [data-testid="stDialogContent"] { max-height: 82vh; overflow-y: auto; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Review and edit any extracted value, then save your changes to regenerate the PDF.")
    columns = st.columns(2)
    for index, key in enumerate(FIELD_ORDER):
        with columns[index % 2]:
            if key == "training_venue":
                st.text_input(FIELD_LABELS[key], key="jd14_edit_training_venue")
            elif key == "employer_name_address":
                st.text_area(FIELD_LABELS[key], key="jd14_edit_" + key, height=120)
            else:
                st.text_input(FIELD_LABELS[key], key="jd14_edit_" + key)
            if key in st.session_state.jd14_warnings:
                st.caption("Review: " + st.session_state.jd14_warnings[key])

    save_col, cancel_col = st.columns(2)
    with save_col:
        if st.button("Save changes", type="primary", key="jd14_save_review"):
            reviewed_values = {
                key: st.session_state.get("jd14_edit_training_venue", "") if key == "training_venue"
                else st.session_state.get("jd14_edit_" + key, "")
                for key in FIELD_ORDER
            }
            st.session_state.jd14_values = reviewed_values
            st.session_state.jd14_pending_venue_sync = reviewed_values["training_venue"]
            st.session_state.jd14_reviewed = True
            st.session_state.jd14_open_review = False
            st.session_state.jd14_pdf = None
            st.rerun()
    with cancel_col:
        if st.button("Cancel", key="jd14_cancel_review"):
            st.session_state.jd14_open_review = False
            st.rerun()


def render_jd14() -> None:
    st.title("HRD Corp JD14 Generator")
    st.markdown(
        '<p style="color: #d32f2f;">Just type the training venue and let AI handle the rest!</p>',
        unsafe_allow_html=True,
    )

    if "jd14_values" not in st.session_state:
        st.session_state.jd14_values = {key: "" for key in FIELD_ORDER}
    if "jd14_warnings" not in st.session_state:
        st.session_state.jd14_warnings = {}
    if "jd14_pdf" not in st.session_state:
        st.session_state.jd14_pdf = None
    if "jd14_extracted" not in st.session_state:
        st.session_state.jd14_extracted = False
    if "jd14_reviewed" not in st.session_state:
        st.session_state.jd14_reviewed = False
    if "jd14_open_review" not in st.session_state:
        st.session_state.jd14_open_review = False

    uploaded = st.file_uploader("Upload employer grant application", type=["pdf"], key="jd14_source_pdf")
    if uploaded is None:
        st.info("Upload a source PDF to begin. The original template is kept in the application and is not replaced.")
        return

    source_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    if (
        source_hash != st.session_state.get("jd14_source_hash")
        or st.session_state.get("jd14_extraction_version") != EXTRACTION_VERSION
    ):
        st.session_state.jd14_source_hash = source_hash
        st.session_state.jd14_extraction_version = EXTRACTION_VERSION
        st.session_state.jd14_values = {key: "" for key in FIELD_ORDER}
        st.session_state.jd14_warnings = {}
        st.session_state.jd14_pdf = None
        st.session_state.jd14_extracted = False
        st.session_state.jd14_reviewed = False
        st.session_state.jd14_open_review = False
        st.session_state.jd14_venue_input = ""
        st.session_state.jd14_edit_training_venue = ""
        for key in FIELD_ORDER:
            if key != "training_venue":
                st.session_state["jd14_edit_" + key] = ""
        st.session_state.jd14_review_check = False

    try:
        source_text, ocr_warnings = extract_pdf_text(uploaded)
    except Exception as exc:
        st.error("Could not read this PDF: {}".format(exc))
        return
    if not source_text:
        st.error("No selectable text was extracted. Check OCR setup or upload a text-based PDF.")
        return

    st.caption("{} characters read. The extracted text is sent to OpenAI for field mapping.".format(len(source_text)))
    for warning in ocr_warnings:
        st.warning(warning)

    if "jd14_pending_venue_sync" in st.session_state:
        st.session_state.jd14_venue_input = st.session_state.pop("jd14_pending_venue_sync")
    if not st.session_state.jd14_open_review:
        st.text_input(
            "Training Venue",
            key="jd14_venue_input",
            placeholder="Enter the training venue",
        )
        if not st.session_state.jd14_extracted:
            st.caption("The completed JD14 download will appear after you enter the venue.")

    venue = st.session_state.get("jd14_venue_input", "").strip()
    if venue and not st.session_state.jd14_extracted:
        values, notes = extract_candidates(source_text)
        try:
            ai_values = extract_with_ai(source_text)
            # AI performs the field mapping. Preserve local matches for fields
            # with exact approval and course-fee rules used by JD14.
            local_values = values
            values = {key: ai_values.get(key) or local_values.get(key, "") for key in FIELD_ORDER}
            for key in ("approval_no", "employer_code", "number_of_trainees", "total_fee_approved", "total_fee_claimed"):
                if local_values.get(key):
                    values[key] = local_values[key]
            notes["_ai"] = "AI extraction completed. Verify all values against the source PDF."
        except Exception as exc:
            notes["_ai"] = "AI extraction unavailable ({}). Local extraction was used; check every field.".format(exc)
        values["training_venue"] = venue
        notes["training_venue"] = "Entered by the user."
        st.session_state.jd14_values = values
        st.session_state.jd14_warnings = notes
        st.session_state.jd14_pdf = None
        st.session_state.jd14_extracted = True
        st.session_state.jd14_open_review = False
        st.session_state.jd14_edit_training_venue = values.get("training_venue", "")
        for key in FIELD_ORDER:
            if key != "training_venue":
                st.session_state["jd14_edit_" + key] = values.get(key, "")

    if st.session_state.jd14_extracted:
        if st.session_state.jd14_open_review:
            _review_jd14_dialog()
        else:
            st.markdown(
                """
                <style>
                div.st-key-jd14_open_review_button button {
                    background-color: #ffc107 !important;
                    color: #202124 !important;
                    border-color: #e0a800 !important;
                }
                div.st-key-jd14_open_review_button button:hover {
                    background-color: #e0a800 !important;
                    border-color: #c69500 !important;
                    color: #202124 !important;
                }
                </style>
                """,
                unsafe_allow_html=True,
            )
            # The same single venue input remains visible after generation.
            # If it is edited, regenerate the PDF without repeating extraction.
            current_venue = st.session_state.get("jd14_venue_input", "").strip()
            if current_venue != st.session_state.jd14_values.get("training_venue", ""):
                st.session_state.jd14_values["training_venue"] = current_venue
                st.session_state.jd14_pdf = None

            errors = validate_values(st.session_state.jd14_values)
            if errors:
                st.session_state.jd14_pdf = None
                st.warning("Please correct these values before downloading:\n\n" + "\n".join("- " + item for item in errors))
                edit_fields = st.button("✏️ Edit fields", key="jd14_open_review_button")
            else:
                try:
                    st.session_state.jd14_pdf = generate_pdf(st.session_state.jd14_values)
                except Exception as exc:
                    st.session_state.jd14_pdf = None
                    st.error("Could not generate the JD14 PDF: {}".format(exc))
                    edit_fields = st.button("✏️ Edit fields", key="jd14_open_review_button")
                else:
                    download_col, edit_col = st.columns([0.34, 0.66], vertical_alignment="center")
                    with download_col:
                        st.download_button(
                            "Download completed JD14 PDF",
                            data=st.session_state.jd14_pdf,
                            file_name=_company_output_filename(st.session_state.jd14_values["employer_name_address"]),
                            mime="application/pdf",
                            type="primary",
                        )
                    with edit_col:
                        edit_fields = st.button("✏️ Edit fields", key="jd14_open_review_button")

            if edit_fields:
                st.session_state.jd14_edit_training_venue = st.session_state.jd14_values.get("training_venue", "")
                for key in FIELD_ORDER:
                    if key != "training_venue":
                        st.session_state["jd14_edit_" + key] = st.session_state.jd14_values.get(key, "")
                st.session_state.jd14_open_review = True
                st.session_state.jd14_pdf = None
                st.rerun()
