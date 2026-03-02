"""
revista_porsche.py
Generează documentul Word «Revista_Porsche_2026.docx» — o revistă informativă
profesională despre compania Porsche, în stilul oficial al brandului.

Dependințe: python-docx, requests, Pillow
Instalare: pip install -r requirements.txt
Utilizare:  python revista_porsche.py
"""

import io
import os
import sys

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

# ---------------------------------------------------------------------------
# Constante de culoare / stil
# ---------------------------------------------------------------------------
NEGRU = RGBColor(0x00, 0x00, 0x00)
GRI_INCHIS = RGBColor(0x33, 0x33, 0x33)
ROS_PORSCHE = RGBColor(0xC5, 0x00, 0x19)
ALB = RGBColor(0xFF, 0xFF, 0xFF)
GRI_DESCHIS = RGBColor(0xF5, 0xF5, 0xF5)
GRI_BORDER = RGBColor(0xCC, 0xCC, 0xCC)

FONT_PRINCIPAL = "Arial"

# ---------------------------------------------------------------------------
# Imagini publice (Wikimedia Commons / licență liberă)
# ---------------------------------------------------------------------------
IMAGINI = {
    "porsche_logo": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/2/25/"
        "Porsche_Logo.svg/500px-Porsche_Logo.svg.png"
    ),
    "porsche_911": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/9/90/"
        "Porsche_911_Carrera_4S_%28992%29_--_2020_-%28cropped%29.jpg/"
        "640px-Porsche_911_Carrera_4S_%28992%29_--_2020_-%28cropped%29.jpg"
    ),
    "porsche_taycan": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/"
        "Porsche_Taycan_4S_--_2020_SCSV_-_Rear_%28cropped%29.jpg/"
        "640px-Porsche_Taycan_4S_--_2020_SCSV_-_Rear_%28cropped%29.jpg"
    ),
    "porsche_cayenne": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/6/6c/"
        "Porsche_Cayenne_III_%28facelift%2C_2023%29_--_2024_SCS_Hannover.jpg/"
        "640px-Porsche_Cayenne_III_%28facelift%2C_2023%29_--_2024_SCS_Hannover.jpg"
    ),
}

os.makedirs("images", exist_ok=True)


def _descarca_imagine(cheie: str, url: str) -> str | None:
    """Descarcă imaginea și o salvează local. Returnează calea sau None."""
    cale = os.path.join("images", f"{cheie}.png")
    if os.path.exists(cale):
        return cale
    try:
        resp = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        img = Image.open(io.BytesIO(resp.content)).convert("RGB")
        img.save(cale, "PNG")
        print(f"  [OK] Imaginea '{cheie}' descărcată.")
        return cale
    except Exception as exc:  # noqa: BLE001
        print(f"  [AVERTISMENT] Nu s-a putut descărca '{cheie}': {exc}")
        return None


def _descarca_imagini() -> dict[str, str | None]:
    print("Descărcare imagini...")
    return {cheie: _descarca_imagine(cheie, url) for cheie, url in IMAGINI.items()}


# ---------------------------------------------------------------------------
# Utilitare document
# ---------------------------------------------------------------------------

def _set_cell_bg(cell, hex_color: str) -> None:
    """Setează culoarea de fundal a unei celule de tabel."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _add_page_break(doc: Document) -> None:
    doc.add_page_break()


def _style_run(run, bold=False, size=12, color=None, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = FONT_PRINCIPAL
    if color:
        run.font.color.rgb = color


def _heading(doc: Document, text: str, level: int = 1) -> None:
    """Adaugă un titlu de secțiune stilizat."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if level == 1:
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text.upper())
        _style_run(run, bold=True, size=20, color=NEGRU)
        # Linie roșie sub titlu
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(10)
        run2 = p2.add_run("─" * 60)
        _style_run(run2, bold=False, size=9, color=ROS_PORSCHE)
    elif level == 2:
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        _style_run(run, bold=True, size=14, color=ROS_PORSCHE)
    else:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        _style_run(run, bold=True, size=12, color=GRI_INCHIS)


def _body(doc: Document, text: str, size=11, color=None, indent=False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    run = p.add_run(text)
    _style_run(run, size=size, color=color or GRI_INCHIS)


def _bullet(doc: Document, text: str, bold_prefix: str = "") -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    if bold_prefix:
        run_pref = p.add_run(f"{bold_prefix} ")
        _style_run(run_pref, bold=True, size=11, color=ROS_PORSCHE)
    run = p.add_run(text)
    _style_run(run, size=11, color=GRI_INCHIS)


def _add_image(doc: Document, cale: str | None, width_cm=14, caption="") -> None:
    if cale and os.path.exists(cale):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(cale, width=Cm(width_cm))
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"[ {caption or 'imagine indisponibilă'} ]")
        _style_run(run, size=10, color=GRI_BORDER, italic=True)
    if caption:
        pc = doc.add_paragraph()
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rc = pc.add_run(caption)
        _style_run(rc, size=9, color=GRI_INCHIS, italic=True)
        pc.paragraph_format.space_after = Pt(8)


def _add_footer(doc: Document) -> None:
    """Adaugă footer cu câmpuri PAGE / NUMPAGES."""
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.clear()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        def _fld(fld_type: str):
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), "begin")
            r = OxmlElement("w:r")
            r.append(fld)

            ri = OxmlElement("w:r")
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = f" {fld_type} "
            ri.append(instr)

            rend = OxmlElement("w:r")
            fldend = OxmlElement("w:fldChar")
            fldend.set(qn("w:fldCharType"), "end")
            rend.append(fldend)
            return r, ri, rend

        run_pre = p.add_run("© Porsche AG 2026  |  Pagina ")
        _style_run(run_pre, size=8, color=GRI_INCHIS)

        r1, ri1, re1 = _fld("PAGE")
        for el in (r1, ri1, re1):
            p._p.append(el)

        run_mid = p.add_run(" din ")
        _style_run(run_mid, size=8, color=GRI_INCHIS)

        r2, ri2, re2 = _fld("NUMPAGES")
        for el in (r2, ri2, re2):
            p._p.append(el)

        run_site = p.add_run("  |  www.porsche.com")
        _style_run(run_site, size=8, color=GRI_INCHIS)


# ---------------------------------------------------------------------------
# Secțiunile revistei
# ---------------------------------------------------------------------------

def _cover_page(doc: Document, img_logo: str | None) -> None:
    """1. Foaie de titlu."""
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.0)

    # Spațiu sus
    for _ in range(4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)

    # Logo
    _add_image(doc, img_logo, width_cm=8, caption="")

    for _ in range(2):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)

    # Titlu mare
    p_titlu = doc.add_paragraph()
    p_titlu.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titlu.paragraph_format.space_after = Pt(6)
    r = p_titlu.add_run("PORSCHE")
    r.bold = True
    r.font.size = Pt(56)
    r.font.name = FONT_PRINCIPAL
    r.font.color.rgb = NEGRU

    # Linie decorativă
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_after = Pt(10)
    rl = p_line.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    rl.font.color.rgb = ROS_PORSCHE
    rl.font.size = Pt(14)
    rl.font.name = FONT_PRINCIPAL

    # Subtitlu
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(6)
    rs = p_sub.add_run("Revista Informativă Oficială")
    rs.font.size = Pt(22)
    rs.font.name = FONT_PRINCIPAL
    rs.font.color.rgb = GRI_INCHIS

    # Data
    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_data.paragraph_format.space_after = Pt(6)
    rd = p_data.add_run("Martie 2026")
    rd.font.size = Pt(14)
    rd.font.name = FONT_PRINCIPAL
    rd.font.color.rgb = GRI_INCHIS

    # Linie jos
    p_line2 = doc.add_paragraph()
    p_line2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line2.paragraph_format.space_before = Pt(20)
    rl2 = p_line2.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    rl2.font.color.rgb = ROS_PORSCHE
    rl2.font.size = Pt(14)
    rl2.font.name = FONT_PRINCIPAL

    # Motto
    p_motto = doc.add_paragraph()
    p_motto.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_motto.paragraph_format.space_before = Pt(10)
    rm = p_motto.add_run('"There is no substitute."')
    rm.italic = True
    rm.font.size = Pt(13)
    rm.font.name = FONT_PRINCIPAL
    rm.font.color.rgb = GRI_INCHIS


def _table_of_contents(doc: Document) -> None:
    """2. Cuprins."""
    _add_page_break(doc)
    _heading(doc, "Cuprins", level=1)

    sectiuni = [
        ("1.", "Despre Companie", "3"),
        ("2.", "Istoria Porsche", "5"),
        ("3.", "Gama de Modele 2025", "8"),
        ("4.", "Servicii Porsche", "12"),
        ("5.", "Contact și Informații", "14"),
    ]

    table = doc.add_table(rows=len(sectiuni) + 1, cols=3)
    table.style = "Table Grid"

    # Header
    hdr = table.rows[0].cells
    for i, txt in enumerate(["Nr.", "Secțiune", "Pagina"]):
        hdr[i].text = txt
        for para in hdr[i].paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(11)
                run.font.name = FONT_PRINCIPAL
                run.font.color.rgb = ALB
        _set_cell_bg(hdr[i], "C50019")

    table.columns[0].width = Cm(1.5)
    table.columns[1].width = Cm(12)
    table.columns[2].width = Cm(2)

    for idx, (nr, titlu, pag) in enumerate(sectiuni, start=1):
        row = table.rows[idx].cells
        row[0].text = nr
        row[1].text = titlu
        row[2].text = pag
        bg = "F5F5F5" if idx % 2 == 0 else "FFFFFF"
        for cell in row:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(11)
                    run.font.name = FONT_PRINCIPAL
                    run.font.color.rgb = GRI_INCHIS
            _set_cell_bg(cell, bg)


def _despre_companie(doc: Document) -> None:
    """3. Despre Companie."""
    _add_page_break(doc)
    _heading(doc, "Despre Companie", level=1)

    _body(doc,
          "Porsche AG este un producător german de automobile sportive de lux, cu sediul "
          "în Stuttgart-Zuffenhausen, Germania. Fondată pe 25 aprilie 1931 de inginerul "
          "Ferdinand Porsche, compania a devenit unul dintre cele mai recunoscute branduri "
          "auto din lume.")

    _heading(doc, "Date Cheie", level=2)
    date = [
        ("Denumire oficială:", "Dr. Ing. h.c. F. Porsche AG"),
        ("Sediu:", "Stuttgart-Zuffenhausen, Germania"),
        ("Fondată:", "25 aprilie 1931"),
        ("Fondator:", "Ferdinand Porsche"),
        ("Grup:", "Volkswagen Group (acționar majoritar)"),
        ("Președinte executiv (CEO):", "Oliver Blume"),
        ("Angajați:", "~40.000 (2024)"),
        ("Venituri:", "~40,5 miliarde EUR (2024)"),
        ("Profit operațional:", "~5,6 miliarde EUR (2024)"),
        ("Vehicule livrate:", "~320.000 unități (2024)"),
        ("Bursa de valori:", "Frankfurt Stock Exchange (P911)"),
    ]
    for cheie, val in date:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        rk = p.add_run(f"{cheie} ")
        _style_run(rk, bold=True, size=11, color=NEGRU)
        rv = p.add_run(val)
        _style_run(rv, size=11, color=GRI_INCHIS)

    _heading(doc, "Misiune și Valori", level=2)
    _body(doc,
          "Porsche urmărește excelența în design și inginerie, îmbinând performanța sportivă "
          "cu luxul și inovația tehnologică. Valorile centrale ale companiei sunt:")
    for val in [
        "Inovație — soluții tehnice de vârf, de la motoare boxer la propulsie electrică",
        "Calitate — standarde stricte de fabricație la uzina din Stuttgart-Zuffenhausen",
        "Sustenabilitate — tranziția spre mobilitate electrică cu modelul Taycan",
        "Exclusivitate — personalizare avansată prin Porsche Exclusive Manufaktur",
        "Tradiție — peste 90 de ani de pasiune pentru sportul cu motor",
    ]:
        _bullet(doc, val, bold_prefix="▸")


def _istoria(doc: Document) -> None:
    """4. Istoria Porsche."""
    _add_page_break(doc)
    _heading(doc, "Istoria Porsche", level=1)

    _body(doc,
          "De la un birou modest de consultanță în inginerie la unul dintre cele mai "
          "admirate branduri auto din lume — istoria Porsche este o poveste de viziune, "
          "pasiune și inovație neîntreruptă.")

    momente = [
        ("1931",
         "Ferdinand Porsche înființează biroul de consultanță în inginerie "
         "«Dr. Ing. h.c. F. Porsche GmbH» la Stuttgart, axat pe proiectare și consultanță."),
        ("1938",
         "Porsche proiectează Volkswagen Beetle (Käfer), un automobil destinat maselor, "
         "la solicitarea Guvernului german."),
        ("1948",
         "Ferry Porsche construiește primul automobil cu emblema Porsche — Porsche 356 Nr. 1 "
         "Roadster. Caroseria din aluminiu și motorul VW de 35 CP devin baza unui brand iconic."),
        ("1950",
         "Producția Porsche 356 se mută la Stuttgart-Zuffenhausen, actualul sediu al companiei."),
        ("1963",
         "Debut mondial al legendarului Porsche 911 la Salonul Auto de la Frankfurt. "
         "Modelul cu motor boxer răcit cu aer amplasat în spate devine unul dintre cele mai "
         "longevive sportcaruri din istorie."),
        ("1970–1980",
         "Porsche domină motorsportul: victorii la Le Mans, Targa Florio și Campionatul Mondial "
         "de Anduranță. Sunt lansate 914, 924 și legendarul 928."),
        ("1984",
         "Debutul Porsche 959, un supercar revoluționar cu tracțiune integrală și "
         "turbocompresoare, care a stabilit standarde tehnice pentru decenii."),
        ("1993",
         "Lansarea Porsche 993 — ultima variantă cu motor răcit cu aer a lui 911, "
         "considerată de mulți colecționari cea mai pură expresie a modelului."),
        ("1996",
         "Debutul Porsche Boxster (986), primul roadster mid-engine al companiei, "
         "care aduce brandul spre un public mai larg."),
        ("2002",
         "Lansarea Porsche Cayenne — primul SUV al companiei. Decizia controversată inițial "
         "se dovedește un succes comercial masiv, salvând financiar Porsche."),
        ("2009",
         "Lansarea Porsche Panamera — primul sedan cu patru uși al brandului, îmbinând "
         "sportivitatea cu confortul executiv."),
        ("2009–2012",
         "Integrarea progresivă în Volkswagen AG. Porsche devine marcă de prestigiu "
         "în cadrul VW Group, alături de Audi, Lamborghini, Bentley și Bugatti."),
        ("2014",
         "Debutul Porsche Macan — SUV compact de lux, care devine rapid cel mai vândut model."),
        ("2019",
         "Lansarea Porsche Taycan — primul automobil complet electric al companiei. "
         "Taycan stabilește standarde pentru EV-uri premium: 0–100 km/h în 2,8 secunde, "
         "autonomie peste 400 km (WLTP)."),
        ("2021",
         "Porsche SE listată la Bursa din Frankfurt. Compania este evaluată la "
         "circa 75 miliarde EUR."),
        ("2025",
         "Lansarea ediției speciale Porsche 911 Turbo 50 Years — marcând 50 de ani "
         "de la debutul lui 911 Turbo. Preț de pornire: 267.300 € incl. TVA. "
         "Disponibil în culori alb, negru, roșu și culori metalice (supliment de la 2.530 €), "
         "cu opțiunea Paint to Sample de la 10.130 €."),
    ]

    for an, descriere in momente:
        _heading(doc, an, level=3)
        _body(doc, descriere, indent=True)


def _modele(doc: Document, imagini: dict) -> None:
    """5. Gama de Modele 2025."""
    _add_page_break(doc)
    _heading(doc, "Gama de Modele 2025", level=1)

    _body(doc,
          "Porsche oferă în 2025 o gamă diversificată de modele, de la sportcaruri iconice "
          "la SUV-uri de lux și vehicule electrice de înaltă performanță.")

    # --- 911 ---
    _heading(doc, "Porsche 911", level=2)
    _add_image(doc, imagini.get("porsche_911"),
               width_cm=13,
               caption="Porsche 911 Carrera 4S (Seria 992)")
    _body(doc,
          "Porsche 911 este sportcarul iconic cu motor boxer amplasat în spate, produs "
          "neîntrerupt din 1963. Seria actuală (992) oferă performanțe excepționale "
          "și rafinament tehnologic de vârf.")
    variante_911 = [
        ("Carrera / Carrera S", "Motor 3.0L twin-turbo, 385 CP / 450 CP, 0–100 în 4,2 / 3,7 s"),
        ("Carrera 4 / 4S", "Tracțiune integrală, motor 3.0L twin-turbo 385 CP / 450 CP"),
        ("Turbo / Turbo S", "Motor 3.7L twin-turbo, 580 CP / 650 CP, 0–100 în 2,7 / 2,5 s"),
        ("GT3", "Motor 4.0L atmosferic, 510 CP, 9.000 RPM, 0–100 în 3,4 s"),
        ("GT3 RS", "Motor 4.0L atmosferic, 525 CP, aerodinamică activă de F1"),
        ("Targa 4 / 4S", "Acoperiș retractabil Targa, motor 3.0L twin-turbo"),
        ("Cabriolet", "Cabriolet cu capotă textilă electrică"),
        ("911 Turbo 50 Years 2025",
         "Ediție limitată aniversară — 50 de ani de la debutul lui 911 Turbo. "
         "580 CP, preț de pornire 267.300 € incl. TVA. Culori: alb, negru, roșu, "
         "metalice (supliment de la 2.530 €), Paint to Sample de la 10.130 €."),
    ]
    for var, desc in variante_911:
        _bullet(doc, desc, bold_prefix=f"{var}:")

    # --- 718 ---
    _heading(doc, "Porsche 718 (Boxster & Cayman)", level=2)
    _body(doc,
          "718 Boxster și 718 Cayman sunt sportcaruri mid-engine pure, cu motoare "
          "amplasate central pentru o distribuție optimă a masei. Oferă experiența "
          "de condus cea mai pură din gama Porsche.")
    for var, desc in [
        ("718 Boxster", "Roadster decapotabil, motor 2.0L / 2.5L turbo, până la 300 CP"),
        ("718 Boxster GTS 4.0", "Motor 4.0L boxster atmosferic, 400 CP, 0–100 în 4,0 s"),
        ("718 Cayman", "Coupé ușor, motor 2.0L / 2.5L turbo, până la 300 CP"),
        ("718 Cayman GT4 RS", "Motor 4.0L atmosferic din 911 GT3, 500 CP, 0–100 în 3,4 s"),
        ("718 Spyder RS", "Open-top, motor 4.0L atmosferic, 500 CP, aerodinamică de circuit"),
    ]:
        _bullet(doc, desc, bold_prefix=f"{var}:")

    # --- Panamera ---
    _add_page_break(doc)
    _heading(doc, "Porsche Panamera", level=2)
    _body(doc,
          "Panamera este sedanul sportiv de lux cu patru uși al Porsche, "
          "combinând performanța unui sportcar cu confortul și spațiul unui sedan executiv.")
    for var, desc in [
        ("Panamera", "Motor V6 turbo 2.9L, 330 CP, tracțiune spate sau integrală"),
        ("Panamera 4S", "Motor V6 twin-turbo, 440 CP, tracțiune integrală"),
        ("Panamera Turbo E-Hybrid", "Motor V8 turbo + motor electric, 680 CP, autonomie 50 km EV"),
        ("Panamera Turbo S E-Hybrid", "Motor V8 twin-turbo + electric, 700 CP, 0–100 în 3,2 s"),
        ("Panamera Sport Turismo", "Variantă combi/wagon cu același ADN sportiv"),
    ]:
        _bullet(doc, desc, bold_prefix=f"{var}:")

    # --- Cayenne ---
    _heading(doc, "Porsche Cayenne", level=2)
    _add_image(doc, imagini.get("porsche_cayenne"),
               width_cm=13,
               caption="Porsche Cayenne (Seria III, facelift 2023)")
    _body(doc,
          "Cayenne este SUV-ul de lux al Porsche, disponibil în versiunile clasică și Coupé. "
          "Rămâne unul dintre cele mai vândute modele ale brandului.")
    for var, desc in [
        ("Cayenne", "Motor V6 turbo 3.0L, 353 CP, SUV clasic 5 locuri"),
        ("Cayenne S", "Motor V8 turbo 2.9L biturbo, 474 CP"),
        ("Cayenne GTS", "Motor V8 4.0L biturbo, 500 CP, setare sportivă"),
        ("Cayenne Turbo E-Hybrid", "Motor V8 + electric, 739 CP, autonomie ~45 km EV"),
        ("Cayenne Coupé", "Caroserie Coupé mai aerodinamică, aceleași grupuri motopropulsoare"),
    ]:
        _bullet(doc, desc, bold_prefix=f"{var}:")

    # --- Macan ---
    _add_page_break(doc)
    _heading(doc, "Porsche Macan", level=2)
    _body(doc,
          "Macan este SUV-ul compact de lux al Porsche, cel mai vândut model al companiei. "
          "Generația nouă (2024+) este disponibilă exclusiv în variantă electrică.")
    for var, desc in [
        ("Macan Electric", "Motor electric 408 CP (versiunea standard), 0–100 în 5,2 s, "
                           "autonomie ~613 km WLTP"),
        ("Macan Turbo Electric",
         "Motor electric dublu, 639 CP, 0–100 în 3,3 s, autonomie ~591 km WLTP"),
    ]:
        _bullet(doc, desc, bold_prefix=f"{var}:")

    # --- Taycan ---
    _heading(doc, "Porsche Taycan", level=2)
    _add_image(doc, imagini.get("porsche_taycan"),
               width_cm=13,
               caption="Porsche Taycan 4S")
    _body(doc,
          "Taycan este primul automobil complet electric al Porsche, lansat în 2019. "
          "Combină performanța unui sportcar pur cu zero emisii locale.")
    for var, desc in [
        ("Taycan", "Propulsie spate sau integrală, 300–400 CP, autonomie ~484 km WLTP"),
        ("Taycan 4S", "Propulsie integrală, 490 CP, autonomie ~464 km WLTP"),
        ("Taycan Turbo", "Propulsie integrală, 680 CP, 0–100 în 3,2 s"),
        ("Taycan Turbo S", "Propulsie integrală, 761 CP, 0–100 în 2,4 s"),
        ("Taycan Sport Turismo", "Variantă wagon cu portbagaj extins"),
        ("Taycan Cross Turismo", "Variantă crossover cu gardă la sol mărită"),
        ("Taycan GTS", "Setare sportivă, 598 CP, orientat spre circuitul de curse"),
    ]:
        _bullet(doc, desc, bold_prefix=f"{var}:")


def _servicii(doc: Document) -> None:
    """6. Servicii Porsche."""
    _add_page_break(doc)
    _heading(doc, "Servicii Porsche", level=1)

    _body(doc,
          "Porsche pune la dispoziția clienților săi un ecosistem complet de servicii, "
          "de la personalizare avansată la soluții financiare și digitale.")

    servicii = [
        (
            "Porsche Exclusive Manufaktur",
            "Serviciul de personalizare premium al Porsche oferă posibilitatea configurării "
            "unui vehicul unic: culori Paint to Sample, tapițerie pe comandă, inserții "
            "decorative speciale și echipamente exclusive. Fiecare detaliu poate fi adaptat "
            "gustului personal al clientului.",
        ),
        (
            "Porsche Approved",
            "Programul de vehicule rulate certificate de Porsche. Fiecare automobil "
            "Porsche Approved trece printr-o inspecție de minimum 111 puncte, beneficiind "
            "de garanție de până la 2 ani fără limită de kilometri.",
        ),
        (
            "Porsche Financial Services",
            "Soluții de finanțare și leasing adaptate stilului de viață Porsche: "
            "leasing operațional, finanțare clasică, asigurări auto și produse de protecție "
            "a valorii vehiculului.",
        ),
        (
            "Porsche Experience Centers",
            "Centre dedicate experienței de condus Porsche, situate în Atlanta, Los Angeles, "
            "Leipzig, Silverstone și Hockenheim. Clienții pot testa modele pe circuite "
            "special amenajate și parcursuri off-road.",
        ),
        (
            "Porsche Connect",
            "Ecosistemul digital Porsche: navigație online, servicii de informații despre trafic "
            "în timp real, comandă de la distanță a vehiculului, actualizări OTA (Over-the-Air) "
            "și integrare Apple CarPlay / Android Auto.",
        ),
        (
            "Porsche Tequipment",
            "Gama oficială de accesorii și piese originale Porsche: roți, seturi aerodinamice, "
            "sisteme audio premium, echipamente pentru stil de viață și articole de colecție.",
        ),
        (
            "Porsche Driver's Selection",
            "Colecția oficială de îmbrăcăminte, accesorii și obiecte de lifestyle Porsche, "
            "disponibilă online și în centrele Porsche.",
        ),
        (
            "Porsche Classic",
            "Servicii de restaurare și întreținere pentru modele Porsche clasice. "
            "Piese originale disponibile pentru modele cu o vechime de peste 10 ani, "
            "inclusiv reproduceri fidele pentru modele rare.",
        ),
    ]

    for titlu, desc in servicii:
        _heading(doc, titlu, level=2)
        _body(doc, desc)


def _pagina_finala(doc: Document) -> None:
    """7. Pagina finală."""
    _add_page_break(doc)
    _heading(doc, "Contact și Informații", level=1)

    _heading(doc, "Date de Contact Porsche AG", level=2)
    contacte = [
        ("Adresă:", "Porscheplatz 1, 70435 Stuttgart, Germania"),
        ("Telefon:", "+49 711 911-0"),
        ("E-mail:", "info@porsche.de"),
        ("Website:", "www.porsche.com"),
        ("Relații investitori:", "investor.relations@porsche.de"),
        ("Presă:", "presse@porsche.de"),
    ]
    for cheie, val in contacte:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        rk = p.add_run(f"{cheie} ")
        _style_run(rk, bold=True, size=11, color=NEGRU)
        rv = p.add_run(val)
        _style_run(rv, size=11, color=GRI_INCHIS)

    _heading(doc, "Rețele Sociale", level=2)
    for retea, handle in [
        ("Instagram:", "@porsche"),
        ("Facebook:", "Porsche"),
        ("YouTube:", "Porsche"),
        ("LinkedIn:", "Porsche AG"),
        ("X (Twitter):", "@Porsche"),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        rk = p.add_run(f"{retea} ")
        _style_run(rk, bold=True, size=11, color=NEGRU)
        rv = p.add_run(handle)
        _style_run(rv, size=11, color=GRI_INCHIS)

    # Linie separator
    for _ in range(2):
        doc.add_paragraph()

    p_sep = doc.add_paragraph()
    p_sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rs = p_sep.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    rs.font.color.rgb = ROS_PORSCHE
    rs.font.size = Pt(10)
    rs.font.name = FONT_PRINCIPAL

    # Disclaimer
    _heading(doc, "Disclaimer", level=2)
    disclaimer = (
        "Acest document a fost elaborat exclusiv în scop informativ și educațional. "
        "Toate informațiile tehnice, prețurile și specificațiile menționate sunt orientative "
        "și pot varia în funcție de piața locală, configurație și data achiziției. "
        "Prețurile menționate includ TVA și sunt valabile pe piața din Germania (conform "
        "listelor de prețuri Porsche AG, Martie 2026).\n\n"
        "Porsche, emblema ecuson Porsche, 911, Taycan, Cayenne, Macan, Panamera, 718 și alte "
        "denumiri de modele sunt mărci înregistrate ale Dr. Ing. h.c. F. Porsche AG, Stuttgart, "
        "Germania. Toate drepturile rezervate.\n\n"
        "© 2026 Dr. Ing. h.c. F. Porsche AG. Toate drepturile rezervate."
    )
    p_disc = doc.add_paragraph()
    p_disc.paragraph_format.left_indent = Cm(0.8)
    p_disc.paragraph_format.space_before = Pt(4)
    rd = p_disc.add_run(disclaimer)
    rd.font.size = Pt(9)
    rd.font.name = FONT_PRINCIPAL
    rd.font.color.rgb = GRI_INCHIS
    rd.italic = True

    # Tagline final
    for _ in range(3):
        doc.add_paragraph()
    p_tag = doc.add_paragraph()
    p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = p_tag.add_run("PORSCHE  —  There is no substitute.")
    rt.bold = True
    rt.font.size = Pt(16)
    rt.font.name = FONT_PRINCIPAL
    rt.font.color.rgb = NEGRU


# ---------------------------------------------------------------------------
# Punct de intrare
# ---------------------------------------------------------------------------

def genereaza_revista(output_path: str = "Revista_Porsche_2026.docx") -> None:
    print("=" * 60)
    print("  Generator Revistă Porsche 2026")
    print("=" * 60)

    imagini = _descarca_imagini()

    print("\nCreare document Word...")
    doc = Document()

    # Setări globale stil
    style = doc.styles["Normal"]
    style.font.name = FONT_PRINCIPAL
    style.font.size = Pt(11)

    _cover_page(doc, imagini.get("porsche_logo"))
    _table_of_contents(doc)
    _despre_companie(doc)
    _istoria(doc)
    _modele(doc, imagini)
    _servicii(doc)
    _pagina_finala(doc)
    _add_footer(doc)

    doc.save(output_path)
    print(f"\n✓ Document salvat: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else "Revista_Porsche_2026.docx"
    genereaza_revista(output)
