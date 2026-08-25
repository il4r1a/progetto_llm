"""
genera_report.py

Genera un file `report.html` di sintesi del progetto sui bias di genere
degli LLM, leggendo:

- call_ollama.xlsx
    foglio 1 -> prompts principali
    foglio 4 -> prompts di controllo (mostra solo le colonne relative al
                prompt, a "parsed_choice" e a "gender_result")
- analisi_risultati_prompts.xlsx
    tutti i fogli -> sezione "analisi dei dati"
- grafici_attribuzione_genere.pdf
- confronto_temperature.pdf

Il file HTML prodotto è autonomo (nessun file esterno da tenere accanto):
i PDF vengono convertiti in immagini PNG e incorporati come base64
direttamente nell'HTML, quindi si può aprire report.html ovunque, anche
senza connessione e senza server locale.

Requisiti:
    pip install pandas openpyxl pymupdf

Uso:
    python genera_report.py
"""

import base64
from pathlib import Path

import pandas as pd
import fitz  # PyMuPDF, per convertire le pagine PDF in immagini

# ---------------------------------------------------------------------------
# Configurazione: percorsi dei file di input/output
# ---------------------------------------------------------------------------

CARTELLA = Path(__file__).parent

FILE_EXCEL_PROMPT = CARTELLA / "call_ollama.xlsx"
FILE_EXCEL_ANALISI = CARTELLA / "analisi_risultati_prompts.xlsx"
FILE_PDF_ATTRIBUZIONE = CARTELLA / "grafici_attribuzione_genere.pdf"
FILE_PDF_TEMPERATURE = CARTELLA / "confronto_temperature.pdf"
FILE_OUTPUT = CARTELLA / "report.html"

# Quante righe mostrare al massimo in ogni tabella, per non generare un
# report enorme se il dataset ha migliaia di osservazioni.
MAX_RIGHE_TABELLA = 50

# Risoluzione di rendering dei PDF (zoom): 2.0 = doppia risoluzione,
# utile per leggere bene testo ed etichette negli assi dei grafici.
ZOOM_PDF = 2.0

# Parole chiave (case-insensitive) usate per riconoscere le colonne da
# mantenere nella tabella "prompts di controllo": qualunque colonna il
# cui nome contenga una di queste stringhe.
COLONNE_CONTROLLO_DA_TENERE = ["prompt", "gender_result"]

# Parole chiave (case-insensitive) usate per riconoscere la colonna che
# identifica il "tipo" di prompt di controllo (es. condition/tipo/type),
# usata per tenere una sola riga per tipo invece di tutte le ripetizioni.
COLONNE_TIPO_CONTROLLO = ["condition", "tipo", "type"]


# ---------------------------------------------------------------------------
# Funzioni di supporto
# ---------------------------------------------------------------------------

def leggi_foglio_excel(percorso_excel: Path, foglio) -> pd.DataFrame:
    """
    Legge un foglio di un file Excel e restituisce un DataFrame.
    `foglio` può essere un indice (0-based) o un nome di foglio.
    Se il file o il foglio non esistono, restituisce un DataFrame vuoto
    con una colonna di avviso, così il report si genera comunque.
    """
    try:
        df = pd.read_excel(percorso_excel, sheet_name=foglio, engine="openpyxl")
        return df
    except FileNotFoundError:
        return pd.DataFrame({"avviso": [f"File non trovato: {percorso_excel.name}"]})
    except Exception as errore:
        return pd.DataFrame({"avviso": [f"Impossibile leggere il foglio {foglio}: {errore}"]})


def leggi_tutti_i_fogli_excel(percorso_excel: Path) -> dict[str, pd.DataFrame]:
    """
    Legge TUTTI i fogli di un file Excel e li restituisce come dizionario
    {nome_foglio: DataFrame}. Se il file non esiste o non è leggibile,
    restituisce un dizionario con un solo DataFrame di avviso.
    """
    try:
        return pd.read_excel(percorso_excel, sheet_name=None, engine="openpyxl")
    except FileNotFoundError:
        return {"avviso": pd.DataFrame({"avviso": [f"File non trovato: {percorso_excel.name}"]})}
    except Exception as errore:
        return {"avviso": pd.DataFrame({"avviso": [f"Impossibile leggere il file: {errore}"]})}


def tieni_un_prompt_per_tipo(df: pd.DataFrame, parole_chiave_tipo: list[str]) -> pd.DataFrame:
    """
    Riduce il DataFrame dei prompt di controllo a una sola riga per ogni
    "tipo" di prompt (es. una sola riga per condizione), scartando le
    ripetizioni. Cerca automaticamente la colonna che identifica il tipo
    tra quelle il cui nome contiene una delle parole chiave fornite; se
    non la trova, restituisce il DataFrame invariato (nessuna deduplicazione).
    """
    if df.empty:
        return df

    colonna_tipo = next(
        (
            colonna
            for colonna in df.columns
            if any(parola.lower() in str(colonna).lower() for parola in parole_chiave_tipo)
        ),
        None,
    )

    if colonna_tipo is None:
        return df

    return df.drop_duplicates(subset=[colonna_tipo], keep="first")


def filtra_colonne_controllo(df: pd.DataFrame, parole_chiave: list[str]) -> pd.DataFrame:
    """
    Restituisce un DataFrame contenente solo le colonne il cui nome
    contiene (case-insensitive) una delle parole chiave fornite.
    Se nessuna colonna corrisponde, restituisce il DataFrame originale
    (per non perdere dati silenziosamente in caso di nomi diversi da
    quelli attesi).
    """
    if df.empty:
        return df

    colonne_selezionate = [
        colonna
        for colonna in df.columns
        if any(parola.lower() in str(colonna).lower() for parola in parole_chiave)
    ]

    if not colonne_selezionate:
        return df

    return df[colonne_selezionate]


def dataframe_a_html(df: pd.DataFrame, max_righe: int = MAX_RIGHE_TABELLA) -> str:
    """
    Converte un DataFrame in una tabella HTML leggibile.
    Se il DataFrame supera max_righe, mostra solo le prime max_righe
    e aggiunge una nota che indica quante righe sono state omesse.
    """
    if df.empty:
        return "<p><em>Nessun dato disponibile.</em></p>"

    n_righe_totali = len(df)
    df_mostrato = df.head(max_righe)

    html_tabella = df_mostrato.to_html(
        index=False,
        border=0,
        na_rep="",
        classes="tabella-dati",
        escape=True,
    )

    nota = ""
    if n_righe_totali > max_righe:
        nota = (
            f"<p class='nota-tabella'>Mostrate le prime {max_righe} righe "
            f"su {n_righe_totali} totali.</p>"
        )

    return html_tabella + nota


def pdf_a_immagini_base64(percorso_pdf: Path, zoom: float = ZOOM_PDF) -> list[str]:
    """
    Converte ogni pagina di un PDF in un'immagine PNG codificata in base64,
    pronta per essere incorporata in un tag <img src="data:image/png;base64,...">.
    Restituisce una lista di stringhe base64 (una per pagina).
    Se il file non esiste o non è leggibile, restituisce una lista vuota.
    """
    if not percorso_pdf.exists():
        return []

    immagini_base64 = []
    matrice_zoom = fitz.Matrix(zoom, zoom)

    try:
        with fitz.open(percorso_pdf) as documento:
            for pagina in documento:
                pixmap = pagina.get_pixmap(matrix=matrice_zoom)
                immagine_bytes = pixmap.tobytes("png")
                immagine_b64 = base64.b64encode(immagine_bytes).decode("ascii")
                immagini_base64.append(immagine_b64)
    except Exception:
        return []

    return immagini_base64


def blocco_immagini_html(immagini_base64: list[str], didascalia: str, nome_file: str) -> str:
    """
    Costruisce il blocco HTML per un PDF convertito in immagini:
    un'immagine per pagina, seguite (una sola volta, in fondo al blocco)
    dalla didascalia fornita.
    """
    if not immagini_base64:
        return (
            f"<div class='grafico-mancante'>"
            f"<p><em>File non trovato o non leggibile: {nome_file}</em></p>"
            f"</div>"
        )

    tag_immagini = ""
    for indice, immagine_b64 in enumerate(immagini_base64, start=1):
        etichetta_pagina = f" (pagina {indice})" if len(immagini_base64) > 1 else ""
        tag_immagini += (
            f"<img src='data:image/png;base64,{immagine_b64}' "
            f"alt='{didascalia}{etichetta_pagina}'>"
        )

    return (
        f"<figure class='figura-grafico'>"
        f"{tag_immagini}"
        f"<figcaption>{didascalia}</figcaption>"
        f"</figure>"
    )


def blocco_fogli_analisi_html(fogli: dict[str, pd.DataFrame]) -> str:
    """
    Costruisce il blocco HTML per la sezione "analisi dei dati": un
    sottotitolo (h3) con il nome del foglio seguito dalla relativa tabella,
    per ciascun foglio del file di analisi.
    """
    blocchi = []
    for nome_foglio, df in fogli.items():
        blocchi.append(f"<h3>{nome_foglio}</h3>")
        blocchi.append(dataframe_a_html(df))
    return "\n".join(blocchi)


# ---------------------------------------------------------------------------
# Caricamento dei dati
# ---------------------------------------------------------------------------

df_prompts = leggi_foglio_excel(FILE_EXCEL_PROMPT, foglio=0)  # foglio 1

df_prompts_controllo_completo = leggi_foglio_excel(FILE_EXCEL_PROMPT, foglio=3)  # foglio 4
df_prompts_controllo_un_per_tipo = tieni_un_prompt_per_tipo(
    df_prompts_controllo_completo, COLONNE_TIPO_CONTROLLO
)
df_prompts_controllo = filtra_colonne_controllo(
    df_prompts_controllo_un_per_tipo, COLONNE_CONTROLLO_DA_TENERE
)

fogli_analisi = leggi_tutti_i_fogli_excel(FILE_EXCEL_ANALISI)

immagini_attribuzione = pdf_a_immagini_base64(FILE_PDF_ATTRIBUZIONE)
immagini_temperature = pdf_a_immagini_base64(FILE_PDF_TEMPERATURE)


# ---------------------------------------------------------------------------
# Costruzione dei blocchi HTML dinamici
# ---------------------------------------------------------------------------

html_tabella_prompts = dataframe_a_html(df_prompts)
html_tabella_prompts_controllo = dataframe_a_html(df_prompts_controllo)
html_analisi_dati = blocco_fogli_analisi_html(fogli_analisi)

html_grafico_attribuzione = blocco_immagini_html(
    immagini_attribuzione,
    didascalia=(
        "Distribuzione delle attribuzioni di genere (uomo/donna) per "
        "professione e condizione informativa."
    ),
    nome_file=FILE_PDF_ATTRIBUZIONE.name,
)

html_grafico_temperature = blocco_immagini_html(
    immagini_temperature,
    didascalia=(
        "Confronto tra le distribuzioni ottenute a temperatura 0 "
        "(risposta deterministica) e a temperatura 0,7 (risposta distributiva)."
    ),
    nome_file=FILE_PDF_TEMPERATURE.name,
)


# ---------------------------------------------------------------------------
# Template HTML (f-string, nessuna libreria di templating)
# ---------------------------------------------------------------------------

html_finale = f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<title>Attribuzione stereotipata di genere da parte di un LLM — Report</title>
<style>
    body {{
        font-family: -apple-system, "Segoe UI", Helvetica, Arial, sans-serif;
        max-width: 800px;
        margin: 40px auto;
        padding: 0 20px;
        line-height: 1.5;
        color: #1a1a1a;
        background-color: #ffffff;
    }}

    h1 {{
        font-size: 1.8em;
        border-bottom: 3px solid #333;
        padding-bottom: 10px;
    }}

    h2 {{
        font-size: 1.3em;
        margin-top: 2.5em;
        border-bottom: 1px solid #ccc;
        padding-bottom: 6px;
    }}

    h3 {{
        font-size: 1.05em;
        margin-top: 1.8em;
        color: #333;
    }}

    p {{
        margin: 1em 0;
    }}

    .nota-tabella {{
        font-size: 0.85em;
        color: #666;
        font-style: italic;
    }}

    table.tabella-dati {{
        border-collapse: collapse;
        width: 100%;
        margin: 1em 0;
        font-size: 0.9em;
    }}

    table.tabella-dati th,
    table.tabella-dati td {{
        border: 1px solid #ccc;
        padding: 6px 8px;
        text-align: left;
    }}

    table.tabella-dati th {{
        background-color: #f2f2f2;
    }}

    table.tabella-dati tr:nth-child(even) {{
        background-color: #fafafa;
    }}

    figure.figura-grafico {{
        margin: 1.5em 0;
        text-align: center;
    }}

    figure.figura-grafico img {{
        width: 100%;
        height: auto;
        border: 1px solid #ddd;
        display: block;
        margin: 0 auto 0.5em auto;
    }}

    figcaption {{
        font-size: 0.9em;
        color: #444;
        text-align: left;
    }}

    .grafico-mancante {{
        padding: 12px;
        background-color: #fff3f3;
        border: 1px dashed #cc8888;
        border-radius: 4px;
    }}

    footer {{
        margin-top: 3em;
        padding-top: 1em;
        border-top: 1px solid #ccc;
        font-size: 0.8em;
        color: #888;
    }}
</style>
</head>
<body>

<h1>Attribuzione stereotipata di genere da parte di un LLM</h1>

<p>
Questo report riassume il progetto sperimentale volto a osservare come un
modello linguistico associa un genere a una persona descritta professionalmente,
e come tale associazione cambi al variare della descrizione (componente tecnica,
relazionale, di leadership) e della modalità di risposta (libera vs. scelta
forzata A/B). Di seguito sono riportati i prompt utilizzati, i risultati dei
prompt di controllo, l'analisi dei dati raccolti e i grafici riassuntivi.
</p>

<h2>Prompts</h2>
<p>Elenco dei prompt principali utilizzati nell'esperimento (foglio 1 del file call_ollama.xlsx).</p>
{html_tabella_prompts}

<h2>Prompts di controllo</h2>
<p>Un prompt per ogni tipo di condizione di controllo, con il relativo gender_result — foglio 4 del file call_ollama.xlsx.</p>
{html_tabella_prompts_controllo}

<h2>Analisi dei dati</h2>
<p>Sintesi dell'analisi condotta sui risultati, dati letti da analisi_risultati_prompts.xlsx (tutti i fogli).</p>
{html_analisi_dati}

<h2>Grafici</h2>
<p>Sintesi visiva dei risultati raccolti.</p>

{html_grafico_attribuzione}

{html_grafico_temperature}

<footer>
    Report generato automaticamente a partire da call_ollama.xlsx,
    analisi_risultati_prompts.xlsx, grafici_attribuzione_genere.pdf
    e confronto_temperature.pdf.
</footer>

</body>
</html>
"""

# ---------------------------------------------------------------------------
# Scrittura del file finale
# ---------------------------------------------------------------------------

FILE_OUTPUT.write_text(html_finale, encoding="utf-8")
print(f"Report generato: {FILE_OUTPUT.resolve()}")
