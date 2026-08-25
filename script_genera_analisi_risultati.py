"""
genera_analisi_risultati_prompts.py

Legge i due file di raccolta dati (temperatura 0.8 e temperatura 0) e produce
un UNICO file Excel di analisi: analisi_risultati_prompts.xlsx

Il file di output contiene 3 fogli:

  1. "Dettaglio"
     Una riga per ogni combinazione professione x condizione (C0-C3):
       - numero di risposte 'male' e 'female' (conteggi grezzi, presi cosi'
         come sono dai file sorgente, per t0.8 e per t0)
       - percentuale male/female sul totale (formula)
       - stabilita' tra le ripetizioni = quota della risposta maggioritaria
         (formula, range 50%-100%)
       - differenza di stabilita' tra t0 e t0.8 (formula)
       - variazione in punti percentuali di %male rispetto alla condizione
         base C0, calcolata separatamente per ogni professione (formula)

  2. "Aggregato per condizione"
     Una riga per ciascuna condizione (C0, C1, C2, C3): media tra tutte le
     professioni di %male, variazione vs C0, stabilita', e deviazione
     standard tra le professioni (per capire se la media e' rappresentativa
     o se nasconde professioni che si muovono in direzioni opposte).

  3. "Note metodologiche"
     Definizione di ogni indicatore usato.

Tutti i valori derivati (percentuali, variazione, stabilita', deviazione
standard) sono scritti come FORMULE Excel, non come numeri pre-calcolati:
il file si ricalcola da solo se apri e modifichi un conteggio grezzo.
Aprendolo in Excel o LibreOffice Calc, le formule vengono calcolate
automaticamente.

Il file di input puo' essere:

  (a) UN SOLO file Excel con piu' fogli, uno per la temperatura 0.8 e uno
      per la temperatura 0 (es. fogli chiamati "temperature=0.8" e
      "temperature=0"): lo script individua da solo i due fogli cercando
      "0.8" e "0" nei nomi dei fogli.

        python genera_analisi_risultati_prompts.py \
            --input prompts_piano_di_lancio_t0_8.xlsx \
            --output analisi_risultati_prompts.xlsx

  (b) DUE file separati, uno per temperatura (modalita' precedente):

        python genera_analisi_risultati_prompts.py \
            --file-t08 prompts_piano_di_lancio.xlsx \
            --file-t0  prompts_piano_di_lancio_t0_1.xlsx \
            --output   analisi_risultati_prompts.xlsx

  Se non viene passato nessun argomento, lo script cerca da solo, nella
  cartella corrente, un file .xlsx che contenga fogli riconducibili alle
  due temperature.

Lo script riconosce automaticamente variazioni ragionevoli nei nomi delle
colonne (es. "profession"/"professione") e nel formato dei valori (es.
"valid" scritto come True/False oppure come "yes"/"no", nomi di
professione con maiuscole/minuscole diverse tra i due fogli).

Richiede: pandas, openpyxl
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

CONDITIONS_ORDER = ["C0", "C1", "C2", "C3"]
BASE_CONDITION = "C0"

# Possibili nomi (case-insensitive, spazi/underscore ignorati) per le
# colonne richieste, cosi' lo script si adatta a piccole differenze tra
# file prodotti in momenti diversi.
COLUMN_ALIASES = {
    "profession": {"profession", "professione", "professioni"},
    "condition": {"condition", "condizione"},
    "choice_gender": {"choicegender", "generescelto", "gender", "genere"},
    "valid": {"valid", "valido"},
}

VALID_TRUE_VALUES = {"true", "yes", "y", "si", "sì", "1", "valid", "ok"}

FONT_NAME = "Arial"
HEADER_FILL = PatternFill(start_color="4C72B0", end_color="4C72B0", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
BASE_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, bold=True, size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# --------------------------------------------------------------------------
# Normalizzazione colonne e valori
# --------------------------------------------------------------------------

def _slug(s: str) -> str:
    return str(s).strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def normalize_columns(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Rinomina le colonne trovate secondo gli alias noti, cosi' il resto
    dello script puo' usare sempre gli stessi nomi (profession, condition,
    choice_gender, valid), indipendentemente da come sono chiamate nel
    file originale."""
    rename_map = {}
    for col in df.columns:
        slug = _slug(col)
        for canonical, aliases in COLUMN_ALIASES.items():
            if slug in {_slug(a) for a in aliases}:
                rename_map[col] = canonical
                break
    df = df.rename(columns=rename_map)

    required = {"profession", "condition", "choice_gender", "valid"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"{source_name}: non trovo le colonne {missing} (colonne presenti: "
            f"{list(df.columns)}). Aggiorna COLUMN_ALIASES nello script se il "
            f"nome usato e' diverso da quelli previsti."
        )
    return df


def normalize_valid(series: pd.Series) -> pd.Series:
    """Accetta 'valid' come booleano True/False oppure come testo
    ('yes'/'no', 'true'/'false', 'si'/'no', ecc.)."""
    def is_true(v):
        if isinstance(v, bool):
            return v
        if pd.isna(v):
            return False
        return str(v).strip().lower() in VALID_TRUE_VALUES
    return series.apply(is_true)


def normalize_profession_key(series: pd.Series) -> pd.Series:
    """Chiave di confronto insensibile a maiuscole/minuscole e spazi
    superflui, per unire correttamente professioni scritte in modo
    leggermente diverso tra i due fogli/file (es. 'Nurse' vs 'nurse')."""
    return series.astype(str).str.strip().str.lower()


# --------------------------------------------------------------------------
# Individuazione automatica di file e fogli
# --------------------------------------------------------------------------

def find_matching_sheet(sheet_names, want_08: bool):
    """Cerca, tra i nomi dei fogli di un workbook, quello relativo a
    temperatura 0.8 o a temperatura 0."""
    for s in sheet_names:
        slug = _slug(s)
        has_08 = "0.8" in s or "08" in slug or "t08" in slug
        if want_08 and has_08:
            return s
    for s in sheet_names:
        slug = _slug(s)
        has_08 = "0.8" in s or "08" in slug or "t08" in slug
        if not want_08 and not has_08 and ("temperature" in slug or slug in {"t0", "temp0"} or slug.endswith("=0")):
            return s
    return None


def find_default_input_file() -> Path:
    """Se non viene passato nessun file, cerca nella cartella corrente un
    .xlsx che contenga sia un foglio per temperatura 0.8 sia uno per
    temperatura 0."""
    for candidate in sorted(Path(".").glob("*.xlsx")):
        try:
            xl = pd.ExcelFile(candidate)
        except Exception:
            continue
        s08 = find_matching_sheet(xl.sheet_names, want_08=True)
        s0 = find_matching_sheet(xl.sheet_names, want_08=False)
        if s08 and s0:
            return candidate
    raise FileNotFoundError(
        "Non ho trovato automaticamente un file .xlsx con i fogli per "
        "temperatura 0.8 e temperatura 0 nella cartella corrente. "
        "Specifica il file con --input, oppure due file separati con "
        "--file-t08 e --file-t0."
    )


# --------------------------------------------------------------------------
# Caricamento e unione dei dati grezzi (nessun valore inventato: solo
# conteggi estratti direttamente dai file sorgente)
# --------------------------------------------------------------------------

def load_counts(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Tiene solo le righe valide con scelta binaria, e restituisce i
    conteggi n_male/n_female per profession x condition."""
    df = normalize_columns(df, source_name)
    df = df[normalize_valid(df["valid"])]
    df = df[df["choice_gender"].astype(str).str.strip().str.lower().isin(["male", "female"])]
    df = df.copy()
    df["choice_gender"] = df["choice_gender"].astype(str).str.strip().str.lower()
    df["condition"] = df["condition"].astype(str).str.strip()
    df["profession_key"] = normalize_profession_key(df["profession"])
    # nome visualizzato: prima grafia incontrata per quella chiave
    display_name = df.groupby("profession_key")["profession"].first()

    counts = (
        df.groupby(["profession_key", "condition"])["choice_gender"]
        .value_counts()
        .unstack(fill_value=0)
        .reindex(columns=["male", "female"], fill_value=0)
        .reset_index()
    )
    counts["profession"] = counts["profession_key"].map(display_name)
    return counts[["profession", "profession_key", "condition", "male", "female"]]


def load_input(args) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Restituisce (df_grezzo_t08, df_grezzo_t0) a seconda della modalita'
    scelta (file unico con piu' fogli, oppure due file separati)."""
    if args.file_t08 and args.file_t0:
        df08 = pd.read_excel(args.file_t08)
        df0 = pd.read_excel(args.file_t0)
        return df08, df0, str(args.file_t08), str(args.file_t0)

    input_path = args.input or find_default_input_file()
    xl = pd.ExcelFile(input_path)
    sheet_08 = find_matching_sheet(xl.sheet_names, want_08=True)
    sheet_0 = find_matching_sheet(xl.sheet_names, want_08=False)
    if not sheet_08 or not sheet_0:
        raise ValueError(
            f"{input_path}: non trovo entrambi i fogli richiesti tra "
            f"{xl.sheet_names}. Rinomina i fogli in modo che uno contenga "
            f"'0.8' e l'altro sia relativo a temperatura 0, oppure usa "
            f"--file-t08/--file-t0 per due file separati."
        )
    df08 = xl.parse(sheet_08)
    df0 = xl.parse(sheet_0)
    return df08, df0, f"{input_path}[{sheet_08}]", f"{input_path}[{sheet_0}]"


def build_merged_counts(args) -> pd.DataFrame:
    df08_raw, df0_raw, name08, name0 = load_input(args)

    c08 = load_counts(df08_raw, name08)
    c0 = load_counts(df0_raw, name0)

    merged = c08.merge(
        c0, on=["profession_key", "condition"], suffixes=("_t08", "_t0")
    )
    merged["profession"] = merged["profession_t08"]
    merged["condition"] = pd.Categorical(
        merged["condition"], categories=CONDITIONS_ORDER, ordered=True
    )
    merged = merged.sort_values(["profession", "condition"]).reset_index(drop=True)
    return merged[["profession", "condition", "male_t08", "female_t08", "male_t0", "female_t0"]]


# --------------------------------------------------------------------------
# Foglio "Dettaglio"
# --------------------------------------------------------------------------

def write_dettaglio_sheet(wb: openpyxl.Workbook, merged: pd.DataFrame) -> str:
    ws = wb.active
    ws.title = "Dettaglio"

    headers = [
        "profession", "condition",
        "n_male_t0.8", "n_female_t0.8", "n_totale_t0.8", "%male_t0.8", "%female_t0.8",
        "n_male_t0", "n_female_t0", "n_totale_t0", "%male_t0", "%female_t0",
        "stabilita_t0.8", "stabilita_t0", "differenza_stabilita_t0_meno_t0.8",
        "variazione_pp_vs_C0_t0.8", "variazione_pp_vs_C0_t0",
    ]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[1].height = 30

    n = len(merged)
    for i, row in merged.iterrows():
        r = i + 2
        ws.cell(row=r, column=1, value=row["profession"]).font = BASE_FONT
        ws.cell(row=r, column=2, value=row["condition"]).font = BASE_FONT
        ws.cell(row=r, column=3, value=int(row["male_t08"])).font = BASE_FONT
        ws.cell(row=r, column=4, value=int(row["female_t08"])).font = BASE_FONT
        ws.cell(row=r, column=5, value=f"=C{r}+D{r}").font = BASE_FONT

        ws.cell(row=r, column=6, value=f"=C{r}/E{r}").font = BASE_FONT
        ws.cell(row=r, column=6).number_format = "0.0%"
        ws.cell(row=r, column=7, value=f"=D{r}/E{r}").font = BASE_FONT
        ws.cell(row=r, column=7).number_format = "0.0%"

        ws.cell(row=r, column=8, value=int(row["male_t0"])).font = BASE_FONT
        ws.cell(row=r, column=9, value=int(row["female_t0"])).font = BASE_FONT
        ws.cell(row=r, column=10, value=f"=H{r}+I{r}").font = BASE_FONT

        ws.cell(row=r, column=11, value=f"=H{r}/J{r}").font = BASE_FONT
        ws.cell(row=r, column=11).number_format = "0.0%"
        ws.cell(row=r, column=12, value=f"=I{r}/J{r}").font = BASE_FONT
        ws.cell(row=r, column=12).number_format = "0.0%"

        # stabilita' = quota della risposta maggioritaria = MAX(%male, %female)
        ws.cell(row=r, column=13, value=f"=MAX(F{r},G{r})").font = BASE_FONT
        ws.cell(row=r, column=13).number_format = "0.0%"
        ws.cell(row=r, column=14, value=f"=MAX(K{r},L{r})").font = BASE_FONT
        ws.cell(row=r, column=14).number_format = "0.0%"
        ws.cell(row=r, column=15, value=f"=N{r}-M{r}").font = BASE_FONT
        ws.cell(row=r, column=15).number_format = "0.0%"

        # variazione vs C0: %male attuale meno %male di C0 della stessa
        # professione (stesso file/temperatura), via SUMPRODUCT
        ws.cell(
            row=r, column=16,
            value=f'=F{r}-SUMPRODUCT(($A$2:$A${n+1}=A{r})*($B$2:$B${n+1}="C0")*$F$2:$F${n+1})',
        ).font = BASE_FONT
        ws.cell(row=r, column=16).number_format = "0.0%"
        ws.cell(
            row=r, column=17,
            value=f'=K{r}-SUMPRODUCT(($A$2:$A${n+1}=A{r})*($B$2:$B${n+1}="C0")*$K$2:$K${n+1})',
        ).font = BASE_FONT
        ws.cell(row=r, column=17).number_format = "0.0%"

        for col in range(1, 18):
            ws.cell(row=r, column=col).border = BORDER

    widths = [22, 10, 11, 12, 12, 10, 11, 10, 11, 11, 9, 10, 11, 10, 15, 14, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    return ws.title


def write_aggregato_sheet(wb: openpyxl.Workbook, n_dett: int) -> None:
    ws2 = wb.create_sheet("Aggregato per condizione")

    headers2 = [
        "condition", "%male_medio_t0.8", "%male_medio_t0",
        "variazione_pp_vs_C0_media_t0.8", "variazione_pp_vs_C0_media_t0",
        "stabilita_media_t0.8", "stabilita_media_t0", "differenza_stabilita_media",
        "dev_std_%male_t0.8", "dev_std_%male_t0",
        "dev_std_variazione_t0.8", "dev_std_variazione_t0",
    ]
    for col, h in enumerate(headers2, start=1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws2.row_dimensions[1].height = 30

    for i, cond in enumerate(CONDITIONS_ORDER):
        r = i + 2
        ws2.cell(row=r, column=1, value=cond).font = BASE_FONT

        ws2.cell(row=r, column=2,
                 value=f'=AVERAGEIF(Dettaglio!$B$2:$B${n_dett+1},A{r},Dettaglio!$F$2:$F${n_dett+1})').font = BASE_FONT
        ws2.cell(row=r, column=3,
                 value=f'=AVERAGEIF(Dettaglio!$B$2:$B${n_dett+1},A{r},Dettaglio!$K$2:$K${n_dett+1})').font = BASE_FONT
        ws2.cell(row=r, column=4,
                 value=f'=AVERAGEIF(Dettaglio!$B$2:$B${n_dett+1},A{r},Dettaglio!$P$2:$P${n_dett+1})').font = BASE_FONT
        ws2.cell(row=r, column=5,
                 value=f'=AVERAGEIF(Dettaglio!$B$2:$B${n_dett+1},A{r},Dettaglio!$Q$2:$Q${n_dett+1})').font = BASE_FONT
        ws2.cell(row=r, column=6,
                 value=f'=AVERAGEIF(Dettaglio!$B$2:$B${n_dett+1},A{r},Dettaglio!$M$2:$M${n_dett+1})').font = BASE_FONT
        ws2.cell(row=r, column=7,
                 value=f'=AVERAGEIF(Dettaglio!$B$2:$B${n_dett+1},A{r},Dettaglio!$N$2:$N${n_dett+1})').font = BASE_FONT
        ws2.cell(row=r, column=8, value=f'=G{r}-F{r}').font = BASE_FONT

        count_expr = f'COUNTIF(Dettaglio!$B$2:$B${n_dett+1},A{r})'
        ws2.cell(row=r, column=9,
                 value=f'=SQRT(SUMPRODUCT((Dettaglio!$B$2:$B${n_dett+1}=A{r})*(Dettaglio!$F$2:$F${n_dett+1}-B{r})^2)/({count_expr}-1))').font = BASE_FONT
        ws2.cell(row=r, column=10,
                 value=f'=SQRT(SUMPRODUCT((Dettaglio!$B$2:$B${n_dett+1}=A{r})*(Dettaglio!$K$2:$K${n_dett+1}-C{r})^2)/({count_expr}-1))').font = BASE_FONT
        ws2.cell(row=r, column=11,
                 value=f'=SQRT(SUMPRODUCT((Dettaglio!$B$2:$B${n_dett+1}=A{r})*(Dettaglio!$P$2:$P${n_dett+1}-D{r})^2)/({count_expr}-1))').font = BASE_FONT
        ws2.cell(row=r, column=12,
                 value=f'=SQRT(SUMPRODUCT((Dettaglio!$B$2:$B${n_dett+1}=A{r})*(Dettaglio!$Q$2:$Q${n_dett+1}-E{r})^2)/({count_expr}-1))').font = BASE_FONT

        for col in range(2, 13):
            ws2.cell(row=r, column=col).number_format = "0.0%"
        for col in range(1, 13):
            ws2.cell(row=r, column=col).border = BORDER

    widths2 = [10, 16, 15, 20, 18, 18, 16, 18, 17, 15, 20, 18]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def write_note_sheet(wb: openpyxl.Workbook) -> None:
    ws3 = wb.create_sheet("Note metodologiche")

    notes = [
        ("Fonte dati",
         "Conteggi n_male e n_female estratti direttamente dai due file di raccolta "
         "(temperatura 0.8 e temperatura 0), filtrando solo le righe con valid=True "
         "e choice_gender in {male, female}."),
        ("%male / %female",
         "n_male / (n_male + n_female) per ogni cella profession x condition."),
        ("Stabilita'",
         "Quota della risposta maggioritaria tra le ripetizioni di una stessa cella: "
         "MAX(%male, %female). 50% = massima incertezza (50/50), 100% = accordo totale "
         "tra le ripetizioni."),
        ("Variazione vs C0",
         "%male della condizione (C1/C2/C3) meno %male della condizione C0, calcolata "
         "separatamente per ogni professione, all'interno dello stesso file/temperatura."),
        ("Deviazione standard tra professioni",
         "Nel foglio Aggregato: quanto le 15 professioni si discostano, in media, dal "
         "valore medio di quella condizione. Se e' grande rispetto alla media, la media "
         "nasconde professioni che si muovono in direzioni opposte (va letto il dettaglio "
         "per singola professione, non solo l'aggregato)."),
    ]
    ws3.cell(row=1, column=1, value="Nota").font = HEADER_FONT
    ws3.cell(row=1, column=1).fill = HEADER_FILL
    ws3.cell(row=1, column=2, value="Descrizione").font = HEADER_FONT
    ws3.cell(row=1, column=2).fill = HEADER_FILL
    for i, (k, v) in enumerate(notes, start=2):
        ws3.cell(row=i, column=1, value=k).font = BOLD_FONT
        ws3.cell(row=i, column=2, value=v).font = BASE_FONT
        ws3.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        ws3.row_dimensions[i].height = 45
    ws3.column_dimensions["A"].width = 28
    ws3.column_dimensions["B"].width = 100


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=None,
                         help="Un unico file .xlsx con un foglio per temperatura 0.8 "
                              "e uno per temperatura 0 (individuati automaticamente). "
                              "Se omesso e non vengono passati --file-t08/--file-t0, "
                              "lo script cerca da solo un file adatto nella cartella corrente.")
    parser.add_argument("--file-t08", type=Path, default=None,
                         help="(Modalita' a due file) file dei risultati a temperatura 0.8")
    parser.add_argument("--file-t0", type=Path, default=None,
                         help="(Modalita' a due file) file dei risultati a temperatura 0")
    parser.add_argument("--output", type=Path, default=Path("analisi_risultati_prompts.xlsx"))
    args = parser.parse_args()

    if bool(args.file_t08) != bool(args.file_t0):
        sys.exit("Errore: --file-t08 e --file-t0 vanno passati insieme, non uno solo.")

    merged = build_merged_counts(args)
    n_dett = len(merged)

    wb = openpyxl.Workbook()
    write_dettaglio_sheet(wb, merged)
    write_aggregato_sheet(wb, n_dett)
    write_note_sheet(wb)

    wb.save(args.output)
    print(f"File creato: {args.output.resolve()}")
    print(f"Righe nel foglio Dettaglio: {n_dett}")
    print("Apri il file in Excel o LibreOffice Calc: le formule si calcolano "
          "automaticamente all'apertura.")


if __name__ == "__main__":
    main()
