import pandas as pd
import pdfplumber
import re
from datetime import datetime

# --- POMOŽNE FUNKCIJE ZA BRANJE ---

def preberi_pdf(pot_do_datoteke):
    """Prebere besedilo iz PDF datoteke."""
    besedilo = ""
    try:
        with pdfplumber.open(pot_do_datoteke) as pdf:
            for stran in pdf.pages:
                besedilo += stran.extract_text() + "\n"
    except Exception as e:
        print(f"Napaka pri branju PDF {pot_do_datoteke}: {e}")
    return besedilo

def preberi_excel(pot_do_datoteke):
    """Prebere podatke iz Excel datoteke."""
    try:
        df = pd.read_excel(pot_do_datoteke)
        return df
    except Exception as e:
        print(f"Napaka pri branju Excela {pot_do_datoteke}: {e}")
        return None

# --- GLAVNA FUNKCIJA ZA PRIMERJAVO ---

def primerjaj_dokumente(narocilnica_poti, prejemnica_poti, racun_poti):
    """
    Primerja tri dokumente. Tukaj bi v praksi izluščili podatke.
    Za potrebe primera predpostavimo, da so podatki že znani ali jih izluščimo s pravili.
    """
    
    print("--- ZAČETEK PRIMERJAVE ---")
    
    # 1. Tukaj bi v resnici izluščili podatke iz datotek
    # Primer (simulacija, ker ne poznam strukture vaših dokumentov):
    
    # podatki_narocilnice = izlusci_podatke(narocilnica_poti)
    # podatki_prejemnice = izlusci_podatke(prejemnica_poti)
    # podatki_racuna = izlusci_podatke(racun_poti)
    
    # Za testiranje uporabimo "mock" podatke (vstavite svoje)
    podatki_narocilnice = {
        "st_dokumenta": "PO-2023-001",
        "dobavitelj": "Podjetje d.o.o.",
        "znesek_brez_ddv": 1000.00,
        "ddv": 220.00,
        "skupni_znesek": 1220.00,
        "datum": "2023-10-01"
    }
    
    podatki_prejemnice = {
        "st_dokumenta": "PR-2023-055",
        "povezava_na_narocilnico": "PO-2023-001",
        "prejeto_blago": "Da", # Ali je blago prevzeto
        "datum_prejema": "2023-10-05"
    }

    podatki_racuna = {
        "st_dokumenta": "RAC-2023-100",
        "sklic_na_narocilnico": "PO-2023-001",
        "znesek_brez_ddv": 1000.00,
        "ddv": 220.00,
        "skupni_znesek": 1220.00,
        "datum_izdaje": "2023-10-06"
    }

    # --- LOGIKA PRIMERJAVE ---
    
    napake = []

    # 1. Preverjanje povezave Naročilnica <-> Račun
    if podatki_racuna.get("sklic_na_narocilnico") != podatki_narocilnice.get("st_dokumenta"):
        napake.append("NAPAKA: Račun se ne sklicuje na pravo naročilnico.")
    
    # 2. Preverjanje povezave Naročilnica <-> Prejemnica
    if podatki_prejemnice.get("povezava_na_narocilnico") != podatki_narocilnice.get("st_dokumenta"):
        napake.append("NAPAKA: Prejemnica se ne ujema z naročilnico.")

    # 3. Preverjanje zneskov (Račun vs Naročilnica)
    if podatki_racuna.get("skupni_znesek") != podatki_narocilnice.get("skupni_znesek"):
        napake.append(f"NAPAKA: Znesek na računu ({podatki_racuna.get('skupni_znesek')}) se ne ujema z naročilnico ({podatki_narocilnice.get('skupni_znesek')}).")

    # 4. Preverjanje DDV
    if podatki_racuna.get("ddv") != podatki_narocilnice.get("ddv"):
         napake.append("NAPAKA: DDV na računu se ne ujema z naročilnico.")

    # 5. Preverjanje datuma (Prejemnica mora biti pred ali isti dan kot Račun - običajno)
    # Tukaj bi morali pretvoriti stringe v datetime objekte za resno primerjavo
    # datum_prejema = datetime.strptime(podatki_prejemnice["datum_prejema"], "%Y-%m-%d")
    # datum_racuna = datetime.strptime(podatki_racuna["datum_izdaje"], "%Y-%m-%d")
    # if datum_racuna < datum_prejema:
    #    napake.append("OPOZORILO: Račun je izdan pred prejemom blaga.")

    # --- IZPIS REZULTATOV ---
    
    if not napake:
        print("✅ VSI DOKUMENTI SE UJEMAJO. Račun je pripravljen za plačilo.")
    else:
        print("❌ PRIŠLO JE DO NAPAKE PRI PRIMERJAVI:")
        for napaka in napake:
            print(f" - {napaka}")

    return napake

# --- ZAGON ---
if __name__ == "__main__":
    # Tukaj vpišete poti do vaših datotek (če jih imate naložene v repozitorij)
    # Primer: 
    # primerjaj_dokumente("narocilnica.pdf", "prejemnica.pdf", "racun.pdf")
    
    # Za testiranje zaženemo s praznimi potmi, ker uporabljamo mock podatke
    primerjaj_dokumente(None, None, None)
