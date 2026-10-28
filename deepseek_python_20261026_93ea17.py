import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- Odvisnosti za branje dokumentov ---
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import docx  # python-docx
except ImportError:
    docx = None

# --- OCR ---
try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    pytesseract = None
    Image = None
    OCR_AVAILABLE = False

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    convert_from_bytes = None
    PDF2IMAGE_AVAILABLE = False


# ---------------------------------------------------------------------------
# 0. Konfiguracija OCR
# ---------------------------------------------------------------------------

def configure_tesseract():
    """
    Na lokalnih sistemih je včasih treba nastaviti pot do tesseract.exe.
    V oblaku (Streamlit Cloud, Linux) je tesseract v PATH in to ni potrebno.
    """
    import os
    import platform
    if pytesseract is None:
        return
    if platform.system() == "Windows":
        for path in [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break


configure_tesseract()


# ---------------------------------------------------------------------------
# 1. Pomožne funkcije (nespremenjene)
# ---------------------------------------------------------------------------

STATUS_COLORS = {
    "Popolno": "C6EFCE",
    "Podaljšano (LoC)": "FFEB9C",
    "Nepopolno": "FFC7CE",
    "Potekel": "FF9999",
    "Ni podatkov": "D9D9D9",
}


def parse_date(value):
    if pd.isna(value) or value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat", ""):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s[: len(fmt) + 2] if "%H" in fmt else s, fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(s, dayfirst=True).to_pydatetime()
    except Exception:
        return None


def get_max_transition_date(razred):
    if pd.isna(razred):
        return None
    r = str(razred).strip().upper()
    if r.startswith("III") or r in ("IIB", "II B"):
        return datetime(2027, 12, 31)
    if r.startswith("II") or r.startswith("I"):
        return datetime(2028, 12, 31)
    return None


# ---------------------------------------------------------------------------
# 2. Branje dokumentov z OCR podporo
# ---------------------------------------------------------------------------

def read_pdf_text_native(file_bytes, max_pages=30):
    """Branje PDF prek pdfplumber (besedilni sloj)."""
    if pdfplumber is None:
        return ""
    text = []
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages[:max_pages]:
                t = page.extract_text() or ""
                text.append(t)
    except Exception:
        pass
    return "\n".join(text)


def read_pdf_with_ocr(file_bytes, lang="slv+eng", dpi=300, max_pages=10):
    """OCR nad PDF prek pdf2image + pytesseract."""
    if not (OCR_AVAILABLE and PDF2IMAGE_AVAILABLE):
        return ""
    try:
        images = convert_from_bytes(file_bytes, dpi=dpi, first_page=1, last_page=max_pages)
    except Exception:
        return ""
    parts = []
    for img in images:
        try:
            parts.append(pytesseract.image_to_string(img, lang=lang))
        except Exception:
            continue
    return "\n".join(parts)


def read_pdf_text(file_bytes, use_ocr=True, ocr_lang="slv+eng", ocr_dpi=300):
    """
    Hibridno branje PDF-ja:
    1. Najprej poskusi pdfplumber.
    2. Če vrne premalo besedila (<50 znakov na stran) in je OCR na voljo,
       preklopi na OCR.
    """
    native = read_pdf_text_native(file_bytes)
    if not use_ocr or not OCR_AVAILABLE or not PDF2IMAGE_AVAILABLE:
        return native

    # Če je besedila premalo, uporabi OCR
    if len(native.strip()) < 100:
        ocr = read_pdf_with_ocr(file_bytes, lang=ocr_lang, dpi=ocr_dpi)
        if len(ocr.strip()) > len(native.strip()):
            return ocr
    return native


def read_image_with_ocr(file_bytes, lang="slv+eng"):
    if not OCR_AVAILABLE:
        return ""
    try:
        img = Image.open(BytesIO(file_bytes))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        return pytesseract.image_to_string(img, lang=lang)
    except Exception:
        return ""


def read_docx_text(file_bytes):
    if docx is None:
        return ""
    try:
        d = docx.Document(BytesIO(file_bytes))
        parts = [p.text for p in d.paragraphs]
        for table in d.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(parts)
    except Exception:
        return ""


def read_txt_text(file_bytes):
    try:
        return file_bytes.decode("utf-8", errors="ignore")
    except Exception:
        return ""


IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif")


def extract_text(file_name, file_bytes, use_ocr=True, ocr_lang="slv+eng", ocr_dpi=300):
    name = file_name.lower()
    if name.endswith(".pdf"):
        return read_pdf_text(file_bytes, use_ocr=use_ocr, ocr_lang=ocr_lang, ocr_dpi=ocr_dpi)
    if name.endswith(".docx"):
        return read_docx_text(file_bytes)
    if name.endswith(".txt"):
        return read_txt_text(file_bytes)
    if name.endswith(IMAGE_EXTS):
        return read_image_with_ocr(file_bytes, lang=ocr_lang)
    return ""


def classify_document_type(file_name):
    n = file_name.lower()
    if any(k in n for k in ["loc", "letter of confirmation", "confirmation letter",
                             "nb letter", "nb_confirmation", "notified body",
                             "confirmation_letter", "confirmation-letter",
                             "podaljš", "extension"]):
        return "LoC"
    if any(k in n for k in ["ce ", "ce_", "_ce", "ec certificate", "ec_certificate",
                             "certifikat", "certificate", "mdr cert", "mdr_cert",
                             "ec cert", "ec_cert"]):
        return "Certifikat"
    if any(k in n for k in ["doc", "declaration of conformity", "declaration_of_conformity",
                             "izjava o skladnosti", "izjava_o_skladnosti",
                             "declaration-of-conformity"]):
        return "DoC"
    return "Drugo"


def find_idents_in_text(text, idents):
    if not text:
        return []
    found = set()
    for ident in idents:
        s = str(ident).strip()
        if not s or len(s) < 4:
            continue
        # Dovoli presledke med številkami (npr. "967 917")
        pattern = r"(?<!\d)" + r"\s*".join(map(re.escape, s)) + r"(?!\d)"
        if re.search(pattern, text):
            found.add(s)
    return sorted(found)


# ---------------------------------------------------------------------------
# 3. Verižno preverjanje (nespremenjeno)
# ---------------------------------------------------------------------------

def classify_row(row, dokumentacija_df, uploaded_index, today=None):
    if today is None:
        today = datetime.today()

    result = {
        "Status_izracun": "Ni podatkov",
        "Opomba_izracun": "",
        "Korak_DoC": "ni najden",
        "Korak_Certifikat": "ni preverjen",
        "Korak_LoC": "ni potreben",
        "Vir_DoC": "",
        "Vir_Certifikat": "",
        "Vir_LoC": "",
        "Nalozen_DoC": "",
        "Nalozen_Certifikat": "",
        "Nalozen_LoC": "",
        "Prehodno_obdobje_do": "",
    }

    ident = str(row.get("IDENT", "")).strip()
    razred = row.get("Razred")
    st_cert = row.get("Št. certifikata")
    veljavnost_cert = parse_date(row.get("Veljavnost certifikata"))
    veljavnost_loc = parse_date(row.get("Veljavnost po letter of confirmation"))
    razred_str = str(razred).strip() if not pd.isna(razred) else ""

    def _uploaded_names(kind):
        names = []
        for d in uploaded_index.get(kind, []):
            if ident in d["idents"]:
                names.append(d["file_name"])
        return "; ".join(names)

    nal_doс = _uploaded_names("DoC")
    nal_cert = _uploaded_names("Certifikat")
    nal_loc = _uploaded_names("LoC")
    result["Nalozen_DoC"] = nal_doс
    result["Nalozen_Certifikat"] = nal_cert
    result["Nalozen_LoC"] = nal_loc

    if razred_str.lower() in ("i", "im", "is", "is+m", "ims"):
        opomba = str(row.get("Opomba", "")).lower()
        if "certifikat ni potreben" in opomba or "certifikat ni potrben" in opomba:
            if nal_doс:
                result.update(
                    Status_izracun="Popolno",
                    Korak_DoC="naložen in vsebuje kataloško št.",
                    Opomba_izracun="Razred I – certifikat ni potreben, DoC naložen.",
                )
            else:
                result.update(
                    Status_izracun="Nepopolno",
                    Korak_DoC="manjka naložen DoC",
                    Opomba_izracun="Razred I – DoC ni naložen.",
                )
            return result

    if not nal_doс:
        result.update(
            Status_izracun="Nepopolno",
            Korak_DoC="manjka naložen DoC",
            Opomba_izracun="DoC za to kataloško številko ni naložen.",
        )
        return result

    result["Korak_DoC"] = "naložen in vsebuje kataloško št."

    cert_number = str(st_cert).strip() if not pd.isna(st_cert) else None
    if not cert_number:
        if "certifikat ni potreben" in str(row.get("Opomba", "")).lower():
            result.update(
                Status_izracun="Popolno",
                Korak_Certifikat="ni potreben",
                Opomba_izracun="Certifikat ni potreben, DoC naložen.",
            )
        else:
            result.update(
                Status_izracun="Nepopolno",
                Korak_Certifikat="manjka v Excelu",
                Opomba_izracun="Številka certifikata ni navedena v vrstici.",
            )
        return result

    if veljavnost_cert is None:
        result.update(
            Status_izracun="Nepopolno",
            Korak_Certifikat="veljavnost neznana",
            Opomba_izracun=f"Certifikat {cert_number} – veljavnost ni razvidna.",
        )
        return result

    if veljavnost_cert >= today:
        if nal_cert:
            result.update(
                Status_izracun="Popolno",
                Korak_Certifikat=f"veljaven do {veljavnost_cert.date()}, naložen",
                Opomba_izracun=f"Certifikat {cert_number} veljaven in naložen.",
            )
        else:
            result.update(
                Status_izracun="Nepopolno",
                Korak_Certifikat=f"veljaven do {veljavnost_cert.date()}, NI naložen",
                Opomba_izracun=f"Certifikat {cert_number} veljaven, a ni naložen.",
            )
        return result

    result["Korak_Certifikat"] = f"POTEKEL ({veljavnost_cert.date()})"

    max_transition = get_max_transition_date(razred)
    if max_transition:
        result["Prehodno_obdobje_do"] = max_transition.strftime("%d.%m.%Y")

    if nal_loc:
        if veljavnost_loc and veljavnost_loc >= today:
            result.update(
                Status_izracun="Podaljšano (LoC)",
                Korak_LoC=f"naložen, veljaven do {veljavnost_loc.date()}",
                Opomba_izracun=(
                    f"Certifikat {cert_number} je potekel, LoC naložen in veljaven "
                    f"do {veljavnost_loc.date()}."
                ),
            )
            return result
        if max_transition and today <= max_transition:
            result.update(
                Status_izracun="Podaljšano (LoC)",
                Korak_LoC=f"naložen, prehodno obdobje do {max_transition.date()}",
                Opomba_izracun=(
                    f"Certifikat {cert_number} je potekel, LoC naložen; prehodno "
                    f"obdobje po Uredbi (EU) 2023/607 traja do {max_transition.date()}."
                ),
            )
            return result
        result.update(
            Status_izracun="Nepopolno",
            Korak_LoC="naložen, a prehodno obdobje je poteklo",
            Opomba_izracun=(
                f"Certifikat {cert_number} je potekel; LoC naložen, a prehodno "
                f"obdobje za razred {razred} je poteklo."
            ),
        )
        return result

    result.update(
        Status_izracun="Potekel",
        Korak_LoC="ni naložen",
        Opomba_izracun=(
            f"Certifikat {cert_number} je potekel ({veljavnost_cert.date()}), "
            f"LoC ni naložen."
        ),
    )
    return result


# ---------------------------------------------------------------------------
# 4. Izgradnja Excel datoteke
# ---------------------------------------------------------------------------

def build_excel(df_out, summary_df):
    wb = Workbook()
    wb.remove(wb.active)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _write_sheet(name, df):
        ws = wb.create_sheet(title=name[:31])
        for c_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            if "Status_izracun" in df.columns:
                status_col = list(df.columns).index("Status_izracun") + 1
                status_val = ws.cell(row=r_idx, column=status_col).value
                color = STATUS_COLORS.get(str(status_val), None)
                if color:
                    fill = PatternFill("solid", fgColor=color)
                    for c_idx in range(1, len(df.columns) + 1):
                        ws.cell(row=r_idx, column=c_idx).fill = fill
        for c_idx, col_name in enumerate(df.columns, start=1):
            max_len = max(
                [len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).tolist()[:200]]
            )
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max(10, max_len + 2), 60)
        ws.freeze_panes = "A2"

    _write_sheet("Rezultati", df_out)
    _write_sheet("Povzetek", summary_df)

    for status, sheet_name in [
        ("Nepopolno", "Nepopolno"),
        ("Potekel", "Poteklo"),
        ("Podaljšano (LoC)", "Podaljšano"),
        ("Ni podatkov", "Ni_podatkov"),
    ]:
        sub = df_out[df_out["Status_izracun"] == status]
        if not sub.empty:
            _write_sheet(sheet_name, sub)

    vir_cols = ["JR", "SK", "IDENT", "NAZIV", "PONUDNIK", "PROIZVAJALEC",
                "Status_izracun", "Nalozen_DoC", "Nalozen_Certifikat", "Nalozen_LoC"]
    vir_cols = [c for c in vir_cols if c in df_out.columns]
    if vir_cols:
        _write_sheet("Viri", df_out[vir_cols])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 5. Streamlit aplikacija
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Preverjanje certifikatov", layout="wide")
st.title("🔍 Samodejno preverjanje dokumentacije medicinskih pripomočkov")
st.caption("Verižno preverjanje: DoC → certifikat → veljavnost → Confirmation Letter (LoC)")

# --- Stranska vrstica: OCR nastavitve -------------------------------------
with st.sidebar:
    st.header("⚙️ Nastavitve OCR")
    if OCR_AVAILABLE:
        st.success("pytesseract je na voljo.")
    else:
        st.warning("pytesseract ni nameščen – OCR ne bo deloval.")
    if PDF2IMAGE_AVAILABLE:
        st.success("pdf2image je na voljo.")
    else:
        st.warning("pdf2image ni nameščen – OCR za PDF ne bo deloval.")

    use_ocr = st.checkbox("Uporabi OCR za skenirane PDF-je in slike", value=True,
                          disabled=not OCR_AVAILABLE)
    ocr_lang = st.selectbox("Jezik OCR", ["slv+eng", "eng", "slv"], index=0)
    ocr_dpi = st.slider("DPI za OCR", min_value=150, max_value=400, value=300, step=50)

# --- Nalaganje Excel datoteke ---------------------------------------------
st.header("1️⃣ Excel datoteka")
uploaded_xlsx = st.file_uploader("Naloži Excel datoteko (Certifikati_preverjeno.xlsx)", type=["xlsx"])

# --- Nalaganje dokumentov --------------------------------------------------
st.header("2️⃣ Dokumenti (DoC, certifikati, LoC)")
st.write(
    "Naložite PDF, DOCX, TXT ali slike (PNG/JPG/TIF). Aplikacija bo v vsebini "
    "poiskala kataloške številke in jih povezala z vrsticami v Excelu."
)
uploaded_docs = st.file_uploader(
    "Naloži dokumente (lahko več hkrati)",
    type=["pdf", "docx", "txt", "png", "jpg", "jpeg", "tif", "tiff", "bmp"],
    accept_multiple_files=True,
)

if uploaded_xlsx and uploaded_docs:
    xls = pd.ExcelFile(uploaded_xlsx)
    sheet_r067 = "R067" if "R067" in xls.sheet_names else xls.sheet_names[0]
    df_r067 = pd.read_excel(xls, sheet_name=sheet_r067)
    df_dok = (
        pd.read_excel(xls, sheet_name="Dokumentacija")
        if "Dokumentacija" in xls.sheet_names
        else pd.DataFrame()
    )

    idents = df_r067["IDENT"].dropna().astype(str).str.strip().unique().tolist()

    col1, col2, col3 = st.columns(3)
    col1.metric("Vrstic v Excelu", len(df_r067))
    col2.metric("Kataloških številk", len(idents))
    col3.metric("Naloženih dokumentov", len(uploaded_docs))

    if st.button("🚀 Zaženi verižno preverjanje"):
        # --- 1) Indeksiraj naložene dokumente ---------------------------
        uploaded_index = {"DoC": [], "Certifikat": [], "LoC": [], "Drugo": []}
        progress = st.progress(0, text="Berem naložene dokumente ...")
        for i, doc in enumerate(uploaded_docs):
            file_bytes = doc.read()
            text = extract_text(
                doc.name, file_bytes,
                use_ocr=use_ocr, ocr_lang=ocr_lang, ocr_dpi=ocr_dpi,
            )
            found = find_idents_in_text(text, idents)
            kind = classify_document_type(doc.name)
            uploaded_index[kind].append({
                "file_name": doc.name,
                "idents": found,
                "ocr_used": (use_ocr and len(text) > 0),
            })
            progress.progress((i + 1) / len(uploaded_docs),
                              text=f"Berem {doc.name} ({i+1}/{len(uploaded_docs)})")
        progress.empty()

        # --- 2) Verižno preverjanje --------------------------------------
        with st.spinner("Pregledujem DoC, certifikate in Confirmation Letters ..."):
            results = [
                classify_row(row, df_dok, uploaded_index)
                for _, row in df_r067.iterrows()
            ]
            df_results = pd.DataFrame(results)
            df_out = pd.concat([df_r067.reset_index(drop=True), df_results], axis=1)

        st.success("Preverjanje zaključeno.")

        # --- 3) Povzetek ------------------------------------------------
        st.subheader("📊 Povzetek rezultatov")
        summary = (
            df_results["Status_izracun"]
            .value_counts()
            .rename_axis("Status")
            .reset_index(name="Število")
        )
        summary["Delež (%)"] = (summary["Število"] / len(df_results) * 100).round(1)
        st.dataframe(summary, use_container_width=True)

        # --- 4) Pregled naloženih dokumentov ----------------------------
        st.subheader("📂 Pregled naloženih dokumentov")
        for kind in ["DoC", "Certifikat", "LoC", "Drugo"]:
            items = uploaded_index.get(kind, [])
            if items:
                st.markdown(f"**{kind}** ({len(items)})")
                for d in items:
                    if d["idents"]:
                        st.write(f"✅ {d['file_name']} → {', '.join(d['idents'])}")
                    else:
                        st.write(f"⚠️ {d['file_name']} → ni najdene kataloške številke")

        # --- 5) Podrobnosti ---------------------------------------------
        st.subheader("📋 Podrobni rezultati")
        st.dataframe(df_out, use_container_width=True, height=500)

        # --- 6) Prenos Excel --------------------------------------------
        excel_bytes = build_excel(df_out, summary)
        filename = f"preverjanje_rezultati_{datetime.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            "💾 Prenesi končno Excel datoteko",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif uploaded_xlsx and not uploaded_docs:
    st.info("Naložili ste Excel, a ne še nobenega dokumenta. Dodajte vsaj en DoC, certifikat ali LoC.")
elif uploaded_docs and not uploaded_xlsx:
    st.info("Naložili ste dokumente, a ne še Excel datoteke.")
