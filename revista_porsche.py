"""
revista_porsche.py — Generează revista informativă Porsche în format Word (.docx).
Structură: Copertă (nenumerotată) + 7 pagini numerotate.
"""

import io
import os

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

# ---------------------------------------------------------------------------
# Culori
# ---------------------------------------------------------------------------
NEGRU = RGBColor(0x00, 0x00, 0x00)
GRI = RGBColor(0x33, 0x33, 0x33)
ROS_PORSCHE = RGBColor(0xC5, 0x00, 0x19)
ALB = RGBColor(0xFF, 0xFF, 0xFF)

# ---------------------------------------------------------------------------
# Imagini (Wikimedia Commons — licență liberă)
# ---------------------------------------------------------------------------
IMAGINI = {
    "logo": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/8/8e/"
        "Porsche_logo.svg/320px-Porsche_logo.svg.png"
    ),
    "911": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/9/90/"
        "Porsche_911_Carrera_4S_%28992%2C_facelift%2C_2024%29%2C_front_8.27.24.jpg/"
        "320px-Porsche_911_Carrera_4S_%28992%2C_facelift%2C_2024%29%2C_front_8.27.24.jpg"
    ),
    "taycan": (
        "https://upload.wikimedia.org/wikipedia/commons/thumb/5/54/"
        "2020_Porsche_Taycan_Turbo_S_%28facelift%2C_grey%29%2C_front_8.15.21.jpg/"
        "320px-2020_Porsche_Taycan_Turbo_S_%28facelift%2C_grey%29%2C_front_8.15.21.jpg"
    ),
}


def descarca_imagine(url: str) -> io.BytesIO | None:
    """Descarcă o imagine și returnează un obiect BytesIO sau None dacă eșuează."""
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            return io.BytesIO(r.content)
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Utilitare document
# ---------------------------------------------------------------------------


def set_margini(document: Document, top: float = 2.0, bottom: float = 2.0,
                left: float = 2.5, right: float = 2.5) -> None:
    """Setează marginile paginii (cm)."""
    sectiune = document.sections[0]
    sectiune.top_margin = Cm(top)
    sectiune.bottom_margin = Cm(bottom)
    sectiune.left_margin = Cm(left)
    sectiune.right_margin = Cm(right)


def adauga_run_formatat(paragraph, text: str, bold: bool = False,
                        size: int = 11, culoare: RGBColor = GRI,
                        font_name: str = "Arial") -> None:
    """Adaugă un run formatat la un paragraf."""
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = culoare
    run.font.name = font_name


def adauga_titlu(doc: Document, text: str, nivel: int = 1,
                 culoare: RGBColor = NEGRU, size: int = 18) -> None:
    """Adaugă un titlu de secțiune."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    run.font.color.rgb = culoare
    run.font.name = "Arial"
    if nivel == 1:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        # Linie decorativă roșie sub titlul principal
        adauga_linie_decorativa(doc)
    else:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)


def adauga_linie_decorativa(doc: Document) -> None:
    """Adaugă o linie roșie subțire sub titlu."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run("─" * 60)
    run.font.color.rgb = ROS_PORSCHE
    run.font.size = Pt(9)
    run.font.name = "Arial"


def adauga_text(doc: Document, text: str, size: int = 11,
                culoare: RGBColor = GRI, bold: bool = False,
                aliniere=WD_ALIGN_PARAGRAPH.JUSTIFY) -> None:
    """Adaugă un paragraf de text normal."""
    p = doc.add_paragraph()
    p.alignment = aliniere
    p.paragraph_format.space_after = Pt(4)
    adauga_run_formatat(p, text, bold=bold, size=size, culoare=culoare)


def adauga_bullet(doc: Document, text: str, size: int = 11) -> None:
    """Adaugă un element de listă cu bullet."""
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(2)
    adauga_run_formatat(p, text, size=size, culoare=GRI)


def adauga_rand_cuprins(doc: Document, nr: str, titlu: str, pagina: str) -> None:
    """Adaugă un rând în cuprins cu numerotare de pagini la dreapta."""
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Inches(5.5))
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(f"{nr}  {titlu}")
    run.font.size = Pt(11)
    run.font.name = "Arial"
    run.font.color.rgb = GRI
    run2 = p.add_run(f"\t{pagina}")
    run2.font.size = Pt(11)
    run2.font.name = "Arial"
    run2.font.color.rgb = ROS_PORSCHE
    run2.bold = True


def adauga_page_break(doc: Document) -> None:
    """Adaugă un salt de pagină."""
    doc.add_page_break()


def adauga_imagine_sau_placeholder(doc: Document, url: str,
                                   latime: float = 6.0,
                                   caption: str = "") -> None:
    """Adaugă o imagine descărcată sau un placeholder text."""
    img_data = descarca_imagine(url)
    if img_data:
        try:
            doc.add_picture(img_data, width=Inches(latime))
            if caption:
                p = doc.paragraphs[-1]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception:
            img_data = None
    if not img_data:
        p = doc.add_paragraph("[Imagine Porsche]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        run.font.size = Pt(10)
        run.font.name = "Arial"
    if caption:
        p = doc.add_paragraph(caption)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        run.font.name = "Arial"
        run.italic = True


# ---------------------------------------------------------------------------
# Numerotare pagini în footer
# ---------------------------------------------------------------------------


def adauga_numerotare_pagini(sectiune) -> None:
    """Adaugă numerotare de pagini în footer (format: Pagina X)."""
    footer = sectiune.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("Pagina ")
    run.font.size = Pt(9)
    run.font.name = "Arial"
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # Câmp număr pagină
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")

    run2 = p.add_run()
    run2.font.size = Pt(9)
    run2.font.name = "Arial"
    run2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    run2._r.append(fldChar1)
    run2._r.append(instrText)
    run2._r.append(fldChar2)


# ---------------------------------------------------------------------------
# Secțiuni
# ---------------------------------------------------------------------------


def adauga_sectiune_noua(doc: Document, include_footer: bool = True,
                         start_numar: int | None = None) -> None:
    """
    Adaugă o secțiune nouă în document cu salt de pagină.
    Dacă include_footer=True, activează numerotarea de pagini.
    """
    doc.add_page_break()
    # Adăugăm o nouă secțiune Word (continuous → nextPage)
    new_section = doc.add_section()
    new_section.start_type = 2  # nextPage (deja tratat de page break)
    new_section.footer.is_linked_to_previous = False
    if include_footer:
        adauga_numerotare_pagini(new_section)
        if start_numar is not None:
            # Setăm numărul de start al paginii
            pgNumType = OxmlElement("w:pgNumType")
            pgNumType.set(qn("w:start"), str(start_numar))
            new_section._sectPr.append(pgNumType)


# ---------------------------------------------------------------------------
# Construcție document
# ---------------------------------------------------------------------------


def construieste_document() -> Document:
    doc = Document()
    set_margini(doc)

    # -----------------------------------------------------------------------
    # Secțiunea 0 — Copertă (fără footer cu numerotare)
    # -----------------------------------------------------------------------
    sectiune_coperta = doc.sections[0]
    sectiune_coperta.footer.is_linked_to_previous = False
    # Footer gol pe copertă
    fp = sectiune_coperta.footer.paragraphs[0] if sectiune_coperta.footer.paragraphs \
        else sectiune_coperta.footer.add_paragraph()
    fp.clear()

    # Spațiu vertical înainte de titlu
    for _ in range(8):
        doc.add_paragraph()

    # Titlu PORSCHE
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("PORSCHE")
    run.bold = True
    run.font.size = Pt(60)
    run.font.color.rgb = NEGRU
    run.font.name = "Arial"

    # Linie roșie
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_line = p_line.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    run_line.font.color.rgb = ROS_PORSCHE
    run_line.font.size = Pt(12)

    doc.add_paragraph()

    # Subtitlu
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Revista Informativă Oficială")
    run_sub.font.size = Pt(22)
    run_sub.font.color.rgb = GRI
    run_sub.font.name = "Arial"

    doc.add_paragraph()

    # Data
    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_data = p_data.add_run("Martie 2026")
    run_data.font.size = Pt(16)
    run_data.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run_data.font.name = "Arial"

    # Logo (dacă disponibil)
    doc.add_paragraph()
    adauga_imagine_sau_placeholder(doc, IMAGINI["logo"], latime=1.5)
    last_p = doc.paragraphs[-1]
    last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # -----------------------------------------------------------------------
    # Secțiunea 1 — Pagina 1: Cuprins
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec1 = doc.add_section()
    sec1.start_type = 2
    sec1.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec1)
    # Pagina 1 pornește de la 1
    pgNumType = OxmlElement("w:pgNumType")
    pgNumType.set(qn("w:start"), "1")
    sec1._sectPr.append(pgNumType)
    set_margini_sectiune(sec1)

    adauga_titlu(doc, "CUPRINS", nivel=1)

    doc.add_paragraph()
    adauga_rand_cuprins(doc, "01", "Despre Compania Porsche", "2")
    adauga_rand_cuprins(doc, "02", "Istoria Porsche", "3")
    adauga_rand_cuprins(doc, "03", "Modelele Porsche: 911 și 718", "4")
    adauga_rand_cuprins(doc, "04", "Modelele Porsche: Panamera și Cayenne", "5")
    adauga_rand_cuprins(doc, "05", "Modelele Porsche: Macan și Taycan", "6")
    adauga_rand_cuprins(doc, "06", "Servicii Porsche și Informații de Contact", "7")

    # -----------------------------------------------------------------------
    # Secțiunea 2 — Pagina 2: Despre companie
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec2 = doc.add_section()
    sec2.start_type = 2
    sec2.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec2)
    set_margini_sectiune(sec2)

    adauga_titlu(doc, "DESPRE COMPANIA PORSCHE", nivel=1)

    adauga_text(doc,
                "Porsche AG — denumire oficială Dr. Ing. h.c. F. Porsche AG — este un "
                "producător german de automobile de lux și sport, cu sediul central în "
                "Stuttgart-Zuffenhausen, Germania. Compania face parte din grupul "
                "Volkswagen AG, unul dintre cei mai mari producători de automobile din lume.")

    adauga_titlu(doc, "Informații generale", nivel=2, size=13, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Fondată pe 25 aprilie 1931 de Ferdinand Porsche")
    adauga_bullet(doc, "Sediul central: Stuttgart-Zuffenhausen, Germania")
    adauga_bullet(doc, "Parte din Volkswagen Group (din 2012)")
    adauga_bullet(doc, "Site oficial: www.porsche.com")
    adauga_bullet(doc, "Peste 40.000 de angajați la nivel mondial")
    adauga_bullet(doc, "Producție anuală: aproximativ 300.000 de vehicule")
    adauga_bullet(doc, "Prezență în peste 120 de piețe globale")

    adauga_titlu(doc, "Site-ul oficial www.porsche.com", nivel=2, size=13, culoare=ROS_PORSCHE)
    adauga_text(doc,
                "Site-ul oficial Porsche oferă o experiență digitală completă pentru "
                "pasionații și clienții mărcii. Printre funcționalitățile principale se numără:")
    adauga_bullet(doc, "Configurator online — personalizare completă a vehiculului dorit")
    adauga_bullet(doc, "Știri și noutăți despre modele noi, competiții și inovații")
    adauga_bullet(doc, "Servicii online: programare service, Porsche Connect, aplicații mobile")
    adauga_bullet(doc, "Localizator de dealeri autorizați în întreaga lume")
    adauga_bullet(doc, "Informații despre Porsche Experience Centers")

    adauga_titlu(doc, "Cifre cheie", nivel=2, size=13, culoare=ROS_PORSCHE)
    adauga_text(doc,
                "Porsche generează o cifră de afaceri de peste 40 de miliarde de euro anual, "
                "cu o marjă de profit operațional remarcabilă în industria auto. În 2023, "
                "compania a livrat aproximativ 320.000 de vehicule la nivel global, "
                "confirmând poziția sa de lider în segmentul premium sportiv.")

    # -----------------------------------------------------------------------
    # Secțiunea 3 — Pagina 3: Istoria Porsche
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec3 = doc.add_section()
    sec3.start_type = 2
    sec3.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec3)
    set_margini_sectiune(sec3)

    adauga_titlu(doc, "ISTORIA PORSCHE", nivel=1)

    adauga_text(doc,
                "Povestea Porsche este una dintre cele mai fascinante din istoria "
                "industriei auto mondiale — de la un birou de inginerie la un imperiu "
                "al performanței, al inovației și al eleganței.")

    evenimente = [
        ("1931", "Fondarea firmei de inginerie de către Ferdinand Porsche în Stuttgart. "
                 "Inițial, compania oferea consultanță și design pentru alți producători auto."),
        ("1938", "Dezvoltarea Volkswagen Beetle — Ferdinand Porsche primește comanda "
                 "guvernamentală pentru crearea unui automobil accesibil pentru popor."),
        ("1948", "Primul automobil cu numele Porsche — modelul 356 — creat de Ferry Porsche "
                 "(fiul lui Ferdinand) în Gmünd, Austria, dintr-un cadru de VW Beetle."),
        ("1951", "Decesul lui Ferdinand Porsche la vârsta de 75 de ani. Compania continuă "
                 "sub conducerea fiului său, Ferry Porsche."),
        ("1963", "Lansarea legendarului Porsche 911 la Salonul Auto de la Frankfurt — "
                 "un automobil care va deveni cel mai iconoc sportcar din istorie."),
        ("1970–1980", "Decadă de aur: victorii la Le Mans, lansarea modelelor 924, 928 și 944, "
                      "extinderea gamei spre un public mai larg."),
        ("1996", "Lansarea Porsche Boxster — un model care salvează financiar compania "
                 "și reînnoiește interesul publicului pentru marca Porsche."),
        ("2002", "Lansarea Porsche Cayenne — primul SUV al companiei, care devine rapid "
                 "cel mai vândut model Porsche și sursa principală de profit."),
        ("2009–2012", "Fuziunea cu Volkswagen AG — Porsche devine oficial parte a grupului VW, "
                      "beneficiind de resurse tehnice și financiare sporite."),
        ("2019", "Lansarea Porsche Taycan — primul automobil complet electric al companiei, "
                 "demonstrând că electrificarea și spiritul sportiv Porsche sunt compatibile."),
        ("2025", "Ediții speciale aniversare (911 Turbo 50 Years), continuarea electrificării "
                 "gamei și extinderea ofertei de vehicule plug-in hybrid."),
    ]

    for an, descriere in evenimente:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        run_an = p.add_run(f"{an}  ")
        run_an.bold = True
        run_an.font.size = Pt(11)
        run_an.font.color.rgb = ROS_PORSCHE
        run_an.font.name = "Arial"
        run_desc = p.add_run(descriere)
        run_desc.font.size = Pt(11)
        run_desc.font.color.rgb = GRI
        run_desc.font.name = "Arial"

    # -----------------------------------------------------------------------
    # Secțiunea 4 — Pagina 4: Modele 911 și 718
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec4 = doc.add_section()
    sec4.start_type = 2
    sec4.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec4)
    set_margini_sectiune(sec4)

    adauga_titlu(doc, "MODELELE PORSCHE: 911 ȘI 718", nivel=1)

    # 911
    adauga_titlu(doc, "Porsche 911", nivel=2, size=14, culoare=NEGRU)
    adauga_imagine_sau_placeholder(doc, IMAGINI["911"], latime=5.5,
                                   caption="Porsche 911 Carrera 4S (2024)")

    adauga_text(doc,
                "Porsche 911 este cel mai iconoc sportcar din lume — un automobil cu o "
                "siluetă inconfundabilă, motor boxer amplasat în spate și o filozofie "
                "inginerească unică. Lansat în 1963, a evoluat continuu fără să-și piardă "
                "esența.")

    adauga_titlu(doc, "Variante disponibile:", nivel=2, size=11, culoare=ROS_PORSCHE)
    for varianta in ["Carrera", "Carrera S", "Carrera 4S", "Targa", "Turbo", "Turbo S",
                     "GT3", "GT3 RS", "Ediție specială: 911 Turbo 50 Years (2025)"]:
        adauga_bullet(doc, varianta)

    adauga_titlu(doc, "Specificații tehnice:", nivel=2, size=11, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Motor: twin-turbo boxer 6 cilindri (spate)")
    adauga_bullet(doc, "Putere: 385 CP (Carrera) — 650 CP (GT3 RS)")
    adauga_bullet(doc, "Transmisie: cutie PDK automată sau manuală 7 trepte")
    adauga_bullet(doc, "Tracțiune: spate (RWD) sau integrală (4S, Carrera 4)")
    adauga_bullet(doc, "Preț: de la ~120.000 € până la 267.300 € (Turbo 50 Years)")

    doc.add_paragraph()

    # 718
    adauga_titlu(doc, "Porsche 718 (Boxster & Cayman)", nivel=2, size=14, culoare=NEGRU)
    adauga_text(doc,
                "Porsche 718 reprezintă moștenirea modelelor mid-engine ale companiei. "
                "Disponibil în versiunea decapotabilă (Boxster) sau coupé (Cayman), "
                "este cel mai accesibil sportcar din gama Porsche.")

    adauga_titlu(doc, "Variante disponibile:", nivel=2, size=11, culoare=ROS_PORSCHE)
    for varianta in ["718 / 718 T", "718 S", "718 GTS 4.0", "718 GT4",
                     "718 GT4 RS", "718 Spyder"]:
        adauga_bullet(doc, varianta)

    adauga_titlu(doc, "Specificații tehnice:", nivel=2, size=11, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Motor: turbo 4 cilindri (versiuni de bază) sau aspirat 4.0L boxer 6 (GTS/GT4)")
    adauga_bullet(doc, "Putere: 300–500 CP")
    adauga_bullet(doc, "Motor amplasat central (mid-engine) — distribuție ideală a greutății")
    adauga_bullet(doc, "Disponibil ca roadster (Boxster) sau coupé (Cayman)")

    # -----------------------------------------------------------------------
    # Secțiunea 5 — Pagina 5: Panamera și Cayenne
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec5 = doc.add_section()
    sec5.start_type = 2
    sec5.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec5)
    set_margini_sectiune(sec5)

    adauga_titlu(doc, "MODELELE PORSCHE: PANAMERA ȘI CAYENNE", nivel=1)

    # Panamera
    adauga_titlu(doc, "Porsche Panamera", nivel=2, size=14, culoare=NEGRU)
    adauga_text(doc,
                "Porsche Panamera este sedanul sportiv de lux cu 4 uși care a redefinit "
                "conceptul de Gran Turismo. Disponibil și în versiunea Sport Turismo "
                "(break), combină confortul unui limuzine cu performanța unui sportcar.")

    adauga_titlu(doc, "Variante disponibile:", nivel=2, size=11, culoare=ROS_PORSCHE)
    for varianta in ["Panamera", "Panamera 4", "Panamera 4S", "Panamera GTS",
                     "Panamera Turbo", "Panamera Turbo S E-Hybrid"]:
        adauga_bullet(doc, varianta)

    adauga_titlu(doc, "Specificații tehnice:", nivel=2, size=11, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Motor: V6 turbo, V8 turbo sau plug-in hybrid")
    adauga_bullet(doc, "Putere: 353–700 CP")
    adauga_bullet(doc, "Tracțiune integrală disponibilă pe toate variantele")
    adauga_bullet(doc, "Combinația perfectă între confort de lux și performanță sportivă")
    adauga_bullet(doc, "Disponibil și ca Sport Turismo (break sportiv)")

    doc.add_paragraph()

    # Cayenne
    adauga_titlu(doc, "Porsche Cayenne", nivel=2, size=14, culoare=NEGRU)
    adauga_text(doc,
                "Porsche Cayenne a revoluționat segmentul SUV-urilor de lux la lansarea "
                "sa în 2002 și a devenit rapid cel mai vândut model Porsche din istorie. "
                "Disponibil ca SUV clasic sau în varianta Coupé cu design mai dinamic.")

    adauga_titlu(doc, "Variante disponibile:", nivel=2, size=11, culoare=ROS_PORSCHE)
    for varianta in ["Cayenne", "Cayenne S", "Cayenne GTS", "Cayenne Turbo",
                     "Cayenne Turbo GT", "Cayenne E-Hybrid", "Cayenne Coupé"]:
        adauga_bullet(doc, varianta)

    adauga_titlu(doc, "Specificații tehnice:", nivel=2, size=11, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Motor: V6 turbo, V8 turbo sau plug-in hybrid")
    adauga_bullet(doc, "Putere: 353–640 CP")
    adauga_bullet(doc, "Capacitate off-road și remorcare până la 3.500 kg")
    adauga_bullet(doc, "Cel mai vândut model Porsche din întreaga istorie")

    # -----------------------------------------------------------------------
    # Secțiunea 6 — Pagina 6: Macan și Taycan
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec6 = doc.add_section()
    sec6.start_type = 2
    sec6.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec6)
    set_margini_sectiune(sec6)

    adauga_titlu(doc, "MODELELE PORSCHE: MACAN ȘI TAYCAN", nivel=1)

    # Macan
    adauga_titlu(doc, "Porsche Macan", nivel=2, size=14, culoare=NEGRU)
    adauga_text(doc,
                "Porsche Macan este SUV-ul compact sportiv al companiei, ideal pentru "
                "utilizarea urbană și extraurbană. Noua generație vine complet electrică, "
                "în timp ce versiunea clasică continuă cu motor cu ardere internă.")

    adauga_titlu(doc, "Variante disponibile:", nivel=2, size=11, culoare=ROS_PORSCHE)
    for varianta in ["Macan", "Macan S", "Macan GTS", "Macan Turbo",
                     "Macan Electric (generația nouă)"]:
        adauga_bullet(doc, varianta)

    adauga_titlu(doc, "Specificații tehnice:", nivel=2, size=11, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Motor: turbo 4 cilindri sau electric (noua generație)")
    adauga_bullet(doc, "Design dinamic, ideal pentru uz urban și sportiv")
    adauga_bullet(doc, "Cel mai accesibil SUV din gama Porsche")
    adauga_bullet(doc, "Tracțiune integrală pe toate variantele")

    doc.add_paragraph()

    # Taycan
    adauga_titlu(doc, "Porsche Taycan", nivel=2, size=14, culoare=NEGRU)
    adauga_imagine_sau_placeholder(doc, IMAGINI["taycan"], latime=5.5,
                                   caption="Porsche Taycan Turbo S")
    adauga_text(doc,
                "Porsche Taycan este primul automobil complet electric al companiei, "
                "lansat în 2019. A dovedit că electrificarea nu compromite ADN-ul Porsche "
                "— performanță, agilitate și plăcerea condusului rămân intacte.")

    adauga_titlu(doc, "Variante disponibile:", nivel=2, size=11, culoare=ROS_PORSCHE)
    for varianta in ["Taycan", "Taycan 4S", "Taycan GTS", "Taycan Turbo", "Taycan Turbo S",
                     "Taycan Sport Turismo", "Taycan Cross Turismo"]:
        adauga_bullet(doc, varianta)

    adauga_titlu(doc, "Specificații tehnice:", nivel=2, size=11, culoare=ROS_PORSCHE)
    adauga_bullet(doc, "Putere: 408–761 CP")
    adauga_bullet(doc, "Autonomie: până la 630 km (WLTP)")
    adauga_bullet(doc, "Încărcare rapidă 800V — 5-80% în ~23 minute")
    adauga_bullet(doc, "Accelerare: 0-100 km/h în 2,8 secunde (Turbo S)")
    adauga_bullet(doc, "Dovada că electrificarea nu compromite ADN-ul Porsche")

    # -----------------------------------------------------------------------
    # Secțiunea 7 — Pagina 7: Servicii și contact
    # -----------------------------------------------------------------------
    doc.add_page_break()
    sec7 = doc.add_section()
    sec7.start_type = 2
    sec7.footer.is_linked_to_previous = False
    adauga_numerotare_pagini(sec7)
    set_margini_sectiune(sec7)

    adauga_titlu(doc, "SERVICII PORSCHE ȘI INFORMAȚII DE CONTACT", nivel=1)

    servicii = [
        ("Porsche Exclusive Manufaktur",
         "Serviciul de personalizare completă a vehiculului, oferind clienților "
         "posibilitatea de a crea un automobil unic. De la culori speciale la "
         "tapițerii exclusiviste și dotări personalizate."),
        ("Porsche Approved",
         "Program de vehicule rulate certificate. Fiecare automobil Porsche Approved "
         "trece prin verificări riguroase și beneficiază de garanție extinsă, "
         "oferind siguranța unui vehicul nou."),
        ("Porsche Financial Services",
         "Soluții complete de finanțare, leasing operațional și financiar, "
         "asigurări auto și servicii de fleet management. Adaptat nevoilor "
         "clienților individuali și corporativi."),
        ("Porsche Experience Centers",
         "Centre dedicate experiențelor de condus, cu circuite special amenajate. "
         "Clienții pot testa limitele vehiculelor Porsche în condiții sigure, "
         "ghidați de instructori profesioniști."),
        ("Porsche Connect",
         "Ecosistemul digital Porsche: aplicație mobilă, navigare avansată, "
         "Remote Services (pornire de la distanță, pre-climatizare), servicii "
         "de conectivitate și actualizări over-the-air."),
        ("Porsche Tequipment",
         "Gama oficială de accesorii și piese originale Porsche. De la jante "
         "și sisteme audio la echipamente sportive și produse de îngrijire — "
         "toate certificate și garantate de Porsche AG."),
    ]

    for titlu_serviciu, descriere_serviciu in servicii:
        adauga_titlu(doc, titlu_serviciu, nivel=2, size=12, culoare=ROS_PORSCHE)
        adauga_text(doc, descriere_serviciu)

    doc.add_paragraph()
    adauga_linie_decorativa(doc)

    adauga_titlu(doc, "Informații de contact", nivel=2, size=13, culoare=NEGRU)
    adauga_bullet(doc, "Site oficial: www.porsche.com")
    adauga_bullet(doc, "Configurator online: www.porsche.com/configurator")
    adauga_bullet(doc, "Sediu central: Porscheplatz 1, 70435 Stuttgart-Zuffenhausen, Germania")
    adauga_bullet(doc, "Rețea globală: peste 900 de dealeri autorizați în 120+ țări")

    doc.add_paragraph()

    # Copyright
    p_copy = doc.add_paragraph()
    p_copy.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_copy = p_copy.add_run("Copyright © 2026 Porsche AG. Toate drepturile rezervate.")
    run_copy.font.size = Pt(9)
    run_copy.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    run_copy.font.name = "Arial"
    run_copy.italic = True

    return doc


def set_margini_sectiune(sectiune, top: float = 2.0, bottom: float = 2.0,
                          left: float = 2.5, right: float = 2.5) -> None:
    """Setează marginile pentru o secțiune specifică."""
    sectiune.top_margin = Cm(top)
    sectiune.bottom_margin = Cm(bottom)
    sectiune.left_margin = Cm(left)
    sectiune.right_margin = Cm(right)


# ---------------------------------------------------------------------------
# Punct de intrare
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    nume_fisier = "Revista_Porsche_2026.docx"
    print("Generare revistă Porsche...")
    doc = construieste_document()
    doc.save(nume_fisier)
    print(f"Document salvat: {nume_fisier}")
    print("Gata! Revista Porsche 2026 a fost generată cu succes.")
