import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. Pomožne funkcije
# ---------------------------------------------------------------------------

STATUS_COLORS = {
    "Popolno": "C6EFCE",           # zelena
    "Podaljšano (LoC)": "FFEB9C",  # rumena
    "Nepopolno": "FFC7CE",         # rdeča
    "Potekel": "FF9999",           # temno rdeča
    "Ni podatkov": "D9D9D9",       # siva
}


def parse_date(value):
    """Pretvori različne zapise datumov v datetime. Vrne None, če ne uspe."""
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
    """Najpoznejši datum prehodnega obdobja po Uredbi (EU) 2023/607."""
    if pd.isna(razred):
        return None
    r = str(razred).strip().upper()
    if r.startswith("III") or r in ("IIB", "II B"):
        return datetime(2027, 12, 31)
    if r.startswith("II") or r.startswith("I"):
        return datetime(2028, 12, 31)
    return None


def _normalize(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _supplier_matches(row, df_dok):
    """
    Če ima df_dok stolpca Proizvajalec / Ponudnik, poskusi omejiti ujemanje
    na istega proizvajalca. Vrne filtriran df_dok (ali original, če ni mogoče).
    """
    if df_dok is None or df_dok.empty:
        return df_dok
    if "Proizvajalec" not in df_dok.columns:
        return df_dok

    proizv = _normalize(row.get("PROIZVAJALEC", ""))
    if not proizv:
        return df_dok

    # Ohlapno ujemanje: prva beseda proizvajalca
    key = proizv.split()[0] if proizv.split() else proizv
    if len(key) < 3:
        return df_dok

    mask = df_dok["Proizvajalec"].astype(str).str.lower().str.contains(key, na=False)
    filtered = df_dok[mask]
    return filtered if not filtered.empty else df_dok


def find_doc_files_for_ident(ident, dokumentacija_df, row=None):
    ident = str(ident).strip()
    if dokumentacija_df is None or dokumentacija_df.empty:
        return pd.DataFrame()
    df = dokumentacija_df.copy()
    for col in ("Datoteka", "DoC"):
        if col not in df.columns:
            df[col] = ""
    mask = (
        df["Datoteka"].astype(str).str.contains(ident, na=False)
        & df["DoC"].astype(str).str.strip().str.lower().isin(["da", "yes", "true", "1"])
    )
    sub = df[mask]
    if row is not None and not sub.empty:
        sub2 = _supplier_matches(row, sub)
        if not sub2.empty:
            return sub2
    return sub


def find_cert_files_for_ident(ident, dokumentacija_df, row=None):
    ident = str(ident).strip()
    if dokumentacija_df is None or dokumentacija_df.empty:
        return pd.DataFrame()
    df = dokumentacija_df.copy()
    for col in ("Datoteka", "Certifikat"):
        if col not in df.columns:
            df[col] = ""
    mask = (
        df["Datoteka"].astype(str).str.contains(ident, na=False)
        & df["Certifikat"].astype(str).str.strip().str.lower().isin(["da", "yes", "true", "1"])
    )
    sub = df[mask]
    if row is not None and not sub.empty:
        sub2 = _supplier_matches(row, sub)
        if not sub2.empty:
            return sub2
    return sub


def find_loc_files_for_ident(ident, dokumentacija_df, row=None):
    ident = str(ident).strip()
    if dokumentacija_df is None or dokumentacija_df.empty:
        return pd.DataFrame()
    df = dokumentacija_df.copy()
    for col in ("Datoteka", "LoC"):
        if col not in df.columns:
            df[col] = ""
    mask = (
        df["Datoteka"].astype(str).str.contains(ident, na=False)
        & df["LoC"].astype(str).str.strip().str.lower().isin(["da", "yes", "true", "1"])
    )
    sub = df[mask]
    if row is not None and not sub.empty:
        sub2 = _supplier_matches(row, sub)
        if not sub2.empty:
            return sub2
    return sub


def extract_cert_number_from_doc(df_doc):
    if df_doc is None or df_doc.empty:
        return None
    if "Št. certifikata" in df_doc.columns:
        vals = df_doc["Št. certifikata"].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if not vals.empty:
            return vals.iloc[0]
    return None


def _files_list(df):
    if df is None or df.empty:
        return ""
    return "; ".join(df["Datoteka"].astype(str).tolist())


def classify_row(row, dokumentacija_df, today=None):
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
        "Prehodno_obdobje_do": "",
    }

    ident = row.get("IDENT")
    razred = row.get("Razred")
    st_cert = row.get("Št. certifikata")
    veljavnost_cert = parse_date(row.get("Veljavnost certifikata"))
    veljavnost_loc = parse_date(row.get("Veljavnost po letter of confirmation"))

    razred_str = str(razred).strip() if not pd.isna(razred) else ""

    # ---- Razred I brez certifikata --------------------------------------
    if razred_str.lower() in ("i", "im", "is", "is+m", "ims"):
        opomba = str(row.get("Opomba", "")).lower()
        if "certifikat ni potreben" in opomba or "certifikat ni potrben" in opomba:
            df_doc = find_doc_files_for_ident(ident, dokumentacija_df, row)
            if not df_doc.empty:
                result.update(
                    Status_izracun="Popolno",
                    Korak_DoC="najden",
                    Opomba_izracun="Razred I – certifikat ni potreben, DoC obstaja.",
                    Vir_DoC=_files_list(df_doc),
                )
            else:
                result.update(
                    Status_izracun="Nepopolno",
                    Korak_DoC="ni najden",
                    Opomba_izracun="Razred I – DoC ni najden v dokumentaciji.",
                )
            return result

    # ---- Korak 1: DoC ---------------------------------------------------
    df_doc = find_doc_files_for_ident(ident, dokumentacija_df, row)
    if df_doc.empty:
        result.update(
            Status_izracun="Nepopolno",
            Korak_DoC="ni najden",
            Opomba_izracun="DoC za to kataloško številko ni najden.",
        )
        return result

    result["Korak_DoC"] = "najden"
    result["Vir_DoC"] = _files_list(df_doc)

    # ---- Korak 2: številka certifikata ---------------------------------
    cert_from_doc = extract_cert_number_from_doc(df_doc)
    cert_number = cert_from_doc or (str(st_cert).strip() if not pd.isna(st_cert) else None)

    if not cert_number:
        if "certifikat ni potreben" in str(row.get("Opomba", "")).lower():
            result.update(
                Status_izracun="Popolno",
                Korak_Certifikat="ni potreben",
                Opomba_izracun="Certifikat ni potreben, DoC obstaja.",
            )
        else:
            result.update(
                Status_izracun="Nepopolno",
                Korak_Certifikat="manjka",
                Opomba_izracun="Številka certifikata ni razvidna iz DoC ali vrstice.",
            )
        return result

    # ---- Korak 3: veljavnost certifikata -------------------------------
    if veljavnost_cert is None:
        df_cert = find_cert_files_for_ident(ident, dokumentacija_df, row)
        if not df_cert.empty and "Veljavnost" in df_cert.columns:
            for v in df_cert["Veljavnost"].dropna():
                d = parse_date(v)
                if d:
                    veljavnost_cert = d
                    break
        result["Vir_Certifikat"] = _files_list(df_cert)

    if veljavnost_cert is None:
        result.update(
            Status_izracun="Nepopolno",
            Korak_Certifikat="veljavnost neznana",
            Opomba_izracun=f"Certifikat {cert_number} – veljavnost ni razvidna.",
        )
        return result

    if veljavnost_cert >= today:
        result.update(
            Status_izracun="Popolno",
            Korak_Certifikat=f"veljaven do {veljavnost_cert.date()}",
            Opomba_izracun=f"Certifikat {cert_number} veljaven.",
        )
        return result

    # ---- Korak 4: potekel – preveri LoC --------------------------------
    result["Korak_Certifikat"] = f"POTEKEL ({veljavnost_cert.date()})"

    df_loc = find_loc_files_for_ident(ident, dokumentacija_df, row)
    loc_found = not df_loc.empty
    result["Vir_LoC"] = _files_list(df_loc)

    if veljavnost_loc is None and loc_found and "Veljavnost LoC" in df_loc.columns:
        for v in df_loc["Veljavnost LoC"].dropna():
            d = parse_date(v)
            if d:
                veljavnost_loc = d
                break

    max_transition = get_max_transition_date(razred)
    if max_transition:
        result["Prehodno_obdobje_do"] = max_transition.strftime("%d.%m.%Y")

    if loc_found and veljavnost_loc is not None and veljavnost_loc >= today:
        result.update(
            Status_izracun="Podaljšano (LoC)",
            Korak_LoC=f"najden, veljaven do {veljavnost_loc.date()}",
            Opomba_izracun=(
                f"Certifikat {cert_number} je potekel, vendar je veljavnost "
                f"podaljšana z Confirmation Letter do {veljavnost_loc.date()}."
            ),
        )
        return result

    if loc_found and max_transition is not None and today <= max_transition:
        result.update(
            Status_izracun="Podaljšano (LoC)",
            Korak_LoC=f"najden, prehodno obdobje do {max_transition.date()}",
            Opomba_izracun=(
                f"Certifikat {cert_number} je potekel, vendar je bilo predloženo "
                f"Confirmation Letter; prehodno obdobje po Uredbi (EU) 2023/607 "
                f"traja do {max_transition.date()}."
            ),
        )
        return result

    if loc_found:
        result.update(
            Status_izracun="Nepopolno",
            Korak_LoC="najden, a prehodno obdobje je poteklo",
            Opomba_izracun=(
                f"Certifikat {cert_number} je potekel; Confirmation Letter obstaja, "
                f"vendar je prehodno obdobje za razred {razred} poteklo."
            ),
        )
        return result

    result.update(
        Status_izracun="Potekel",
        Korak_LoC="ni najden",
        Opomba_izracun=(
            f"Certifikat {cert_number} je potekel ({veljavnost_cert.date()}), "
            f"Confirmation Letter ni najden."
        ),
    )
    return result


# ---------------------------------------------------------------------------
# 2. Izgradnja Excel datoteke z oblikovanjem
# ---------------------------------------------------------------------------

def build_excel(df_out, summary_df):
    wb = Workbook()
    wb.remove(wb.active)  # odstrani privzeti list

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _write_sheet(name, df):
        ws = wb.create_sheet(title=name[:31])
        # Glava
        for c_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        # Podatki
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            # Barvanje po statusu
            if "Status_izracun" in df.columns:
                status_col = list(df.columns).index("Status_izracun") + 1
                status_val = ws.cell(row=r_idx, column=status_col).value
                color = STATUS_COLORS.get(str(status_val), None)
                if color:
                    fill = PatternFill("solid", fgColor=color)
                    for c_idx in range(1, len(df.columns) + 1):
                        ws.cell(row=r_idx, column=c_idx).fill = fill
        # Samodejna širina
        for c_idx, col_name in enumerate(df.columns, start=1):
            max_len = max(
                [len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).tolist()[:200]]
            )
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max(10, max_len + 2), 60)
        ws.freeze_panes = "A2"
        return ws

    # 1) Rezultati
    _write_sheet("Rezultati", df_out)

    # 2) Povzetek
    _write_sheet("Povzetek", summary_df)

    # 3) Ločeni listi po statusu
    for status, sheet_name in [
        ("Nepopolno", "Nepopolno"),
        ("Potekel", "Poteklo"),
        ("Podaljšano (LoC)", "Podaljšano"),
        ("Ni podatkov", "Ni_podatkov"),
    ]:
        sub = df_out[df_out["Status_izracun"] == status]
        if not sub.empty:
            _write_sheet(sheet_name, sub)

    # 4) Viri – samo datoteke, ki so bile uporabljene
    vir_cols = ["JR", "SK", "IDENT", "NAZIV", "PONUDNIK", "PROIZVAJALEC",
                "Status_izracun", "Vir_DoC", "Vir_Certifikat", "Vir_LoC"]
    vir_cols = [c for c in vir_cols if c in df_out.columns]
    if vir_cols:
        _write_sheet("Viri", df_out[vir_cols])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 3. Streamlit aplikacija
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Preverjanje certifikatov", layout="wide")
st.title("🔍 Samodejno preverjanje dokumentacije medicinskih pripomočkov")
st.caption("Verižno preverjanje: DoC → certifikat → veljavnost → Confirmation Letter (LoC)")

uploaded = st.file_uploader("Naloži Excel datoteko (Certifikati_preverjeno.xlsx)", type=["xlsx"])

if uploaded:
    xls = pd.ExcelFile(uploaded)

    sheet_r067 = "R067" if "R067" in xls.sheet_names else xls.sheet_names[0]
    df_r067 = pd.read_excel(xls, sheet_name=sheet_r067)
    df_dok = (
        pd.read_excel(xls, sheet_name="Dokumentacija")
        if "Dokumentacija" in xls.sheet_names
        else pd.DataFrame()
    )

    col1, col2 = st.columns(2)
    col1.metric(f"Vrstic v listu {sheet_r067}", len(df_r067))
    col2.metric("Vrstic v listu Dokumentacija", len(df_dok))

    if st.button("🚀 Zaženi verižno preverjanje"):
        with st.spinner("Pregledujem DoC, certifikate in Confirmation Letters ..."):
            results = [classify_row(row, df_dok) for _, row in df_r067.iterrows()]
            df_results = pd.DataFrame(results)
            df_out = pd.concat([df_r067.reset_index(drop=True), df_results], axis=1)

        st.success("Preverjanje zaključeno.")

        # Povzetek
        st.subheader("📊 Povzetek rezultatov")
        summary = (
            df_results["Status_izracun"]
            .value_counts()
            .rename_axis("Status")
            .reset_index(name="Število")
        )
        summary["Delež (%)"] = (summary["Število"] / len(df_results) * 100).round(1)
        st.dataframe(summary, use_container_width=True)

        # Podrobnosti
        st.subheader("📋 Podrobni rezultati")
        st.dataframe(df_out, use_container_width=True, height=500)

        # Izdelaj Excel
        excel_bytes = build_excel(df_out, summary)
        filename = f"preverjanje_rezultati_{datetime.today().strftime('%Y%m%d')}.xlsx"

        st.download_button(
            "💾 Prenesi končno Excel datoteko",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
