# Porsche — Revistă Informativă

Acest repository conține un script Python care generează o revistă informativă profesională despre compania Porsche în format Word (`.docx`).

## Structura revistei

Documentul are **8 pagini fizice**: copertă (nenumerotată) + 7 pagini numerotate.

| Pagina | Conținut |
|--------|----------|
| Copertă | Titlu, subtitlu, dată |
| 1 | Cuprins |
| 2 | Despre Compania Porsche |
| 3 | Istoria Porsche |
| 4 | Modelele Porsche: 911 și 718 |
| 5 | Modelele Porsche: Panamera și Cayenne |
| 6 | Modelele Porsche: Macan și Taycan |
| 7 | Servicii Porsche și Informații de Contact |

## Fișiere

- `revista_porsche.py` — scriptul Python de generare
- `Revista_Porsche_2026.docx` — documentul Word generat
- `requirements.txt` — dependențele Python necesare

## Instalare și utilizare

```bash
# Instalează dependențele
pip install -r requirements.txt

# Generează documentul
python revista_porsche.py
```

Documentul `Revista_Porsche_2026.docx` va fi creat în directorul curent.

## Dependențe

- `python-docx>=0.8.11` — generare documente Word
- `requests>=2.28.0` — descărcare imagini
- `Pillow>=9.0.0` — procesare imagini
