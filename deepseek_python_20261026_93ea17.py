import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="MDR pregled dokumentacije", layout="wide")
st.title("MDR pregled dokumentacije – kataloške številke in certifikati")
st.caption("Naložite Excel preglednico in mapo PDF. Aplikacija bo preverila, kateri dokument vsebuje katero kataloško številko, in označila manjkajoče ter potekle certifikate.")

# ---------- 1. Nalaganje vhodnih podatkov ----------
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Naloži Excel preglednico (.xlsx)", type=["xlsx"])
with col2:
    pdf_files = st.file_uploader(
        "Naloži PDF dokumente (lahko več datotek hkrati)",
        type=["pdf"],
        accept_multiple_files=True,
    )

# ---------- 2. Branje Excela ----------
def preberi_excel(file):
    wb = load_workbook(file, data_only=True)
    ws = wb.active  # privzeto prvi list; po potrebi spremenite
    # Poišči vrstico glave (vsebuje "KATALOŠKA")
    glava_vrstica = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        if row and any(isinstance(c, str) and "KATALOŠKA" in c.upper() for c in row):
            glava_vrstica = i
            break
    if glava_vrstica is None:
        st.error("V preglednici ni bilo mogoče najti vrstice glave s stolpcem 'KATALOŠKA'.")
        return None, None
    glava = [c for c in next(ws.iter_rows(min_row=glava_vrstica, max_row=glava_vrstica, values_only=True))]
    podatki = []
    for row in ws.iter_rows(min_row=glava_vrstica + 1, values_only=True):
        if not any(row):
            continue
        podatki.append(dict(zip(glava, row)))
    df = pd.DataFrame(podatki)
    return df, glava_vrstica

# ---------- 3. Izluščanje besedila iz PDF ----------
def izlusci_besedilo(pdf_file):
    besedilo = []
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for stran in pdf.pages:
                t = stran.extract_text() or ""
                besedilo.append(t)
    except Exception as e:
        return "", f"Napaka pri branju: {e}"
    return "\n".join(besedilo), None

# ---------- 4. Iskanje kataloške številke v besedilu ----------
def najdi_katalosko(besedilo, kataloska):
    if not isinstance(kataloska, str) or not kataloska.strip():
        return False
    k = kataloska.strip()
    # 1) neposredno iskanje (brez regularnih izrazov)
    if k.lower() in besedilo.lower():
        return True
    # 2) iskanje brez presledkov (npr. "MDR 766495" → "MDR766495")
    if re.sub(r"\s+", "", k).lower() in re.sub(r"\s+", "", besedilo).lower():
        return True
    return False

# ---------- 5. Preverjanje veljavnosti ----------
def je_poteklo(datum):
    if datum is None or (isinstance(datum, float) and pd.isna(datum)):
        return False
    if isinstance(datum, datetime):
        return datum.date() < datetime.today().date()
    if isinstance(datum, str):
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%Y/%m/%d"):
            try:
                return datetime.strptime(datum.strip(), fmt).date() < datetime.today().date()
            except ValueError:
                continue
    return False

# ---------- 6. Glavna logika ----------
if excel_file and pdf_files:
    df, glava_vrstica = preberi_excel(excel_file)
    if df is None:
        st.stop()

    # Preveri, ali obstaja stolpec KATALOŠKA
    if "KATALOŠKA" not in df.columns:
        st.error("V preglednici ni stolpca 'KATALOŠKA'.")
        st.stop()

    st.success(f"Prebrana preglednica: {len(df)} vrstic, {len(pdf_files)} PDF datotek.")

    # Izlušči besedilo iz vseh PDF
    with st.spinner("Berem PDF dokumente ..."):
        pdf_besedila = {}
        opozorila_pdf = []
        for f in pdf_files:
            besedilo, napaka = izlusci_besedilo(f)
            if napaka:
                opozorila_pdf.append(f"{f.name}: {napaka}")
            elif not besedilo.strip():
                opozorila_pdf.append(f"{f.name}: brez besedilne plasti (verjetno skenirano) – potreben ročni pregled.")
            pdf_besedila[f.name] = besedilo

    if opozorila_pdf:
        with st.expander("Opozorila pri branju PDF"):
            for o in opozorila_pdf:
                st.warning(o)

    # Primerjava
    with st.spinner("Primerjam kataloške številke z dokumenti ..."):
        rezultati = []
        for _, vrstica in df.iterrows():
            kat = vrstica.get("KATALOŠKA")
            najdene = []
            for ime, besedilo in pdf_besedila.items():
                if najdi_katalosko(besedilo, kat):
                    najdene.append(ime)
            # Preveri veljavnost
            veljavnost = vrstica.get("Veljavnost certifikata")
            poteklo = je_poteklo(veljavnost)
            rezultati.append({
                "Kataloška": kat,
                "Naziv": vrstica.get("NAZIV", ""),
                "Ponudnik": vrstica.get("PONUDNIK", ""),
                "Proizvajalec": vrstica.get("PROIZVAJALEC", ""),
                "Razred": vrstica.get("Razred", ""),
                "Najdeno v dokumentih": "; ".join(najdene) if najdene else "",
                "Število zadetkov": len(najdene),
                "Status": "Najdeno" if najdene else "Ni najdeno",
                "Veljavnost certifikata": veljavnost,
                "Poteklo": "DA" if poteklo else "",
            })

    rez_df = pd.DataFrame(rezultati)

    # Povzetek
    st.subheader("Povzetek")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Skupaj artiklov", len(rez_df))
    c2.metric("Najdeno v dokumentih", int((rez_df["Status"] == "Najdeno").sum()))
    c3.metric("Ni najdeno", int((rez_df["Status"] == "Ni najdeno").sum()))
    c4.metric("Potekli certifikati", int((rez_df["Poteklo"] == "DA").sum()))

    # Prikaz tabele z barvami
    st.subheader("Rezultati pregleda")
    def obarvaj(vrstica):
        if vrstica["Poteklo"] == "DA":
            return ["background-color: #f8d7da"] * len(vrstica)
        if vrstica["Status"] == "Ni najdeno":
            return ["background-color: #fff3cd"] * len(vrstica)
        return ["background-color: #d4edda"] * len(vrstica)

    st.dataframe(
        rez_df.style.apply(obarvaj, axis=1),
        use_container_width=True,
        height=600,
    )

    # Izvoz
    st.subheader("Izvoz rezultatov")
    izvoz = io.BytesIO()
    with pd.ExcelWriter(izvoz, engine="openpyxl") as writer:
        rez_df.to_excel(writer, index=False, sheet_name="Rezultati")
    izvoz.seek(0)
    st.download_button(
        label="Prenesi rezultate kot Excel",
        data=izvoz,
        file_name=f"Preverjeno_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Podrobnosti o dokumentih
    with st.expander("Seznam naloženih PDF in njihova dolžina besedila"):
        for ime, besedilo in pdf_besedila.items():
            st.write(f"**{ime}** – {len(besedilo)} znakov")

else:
    st.info("Naložite Excel preglednico in vsaj en PDF dokument.")