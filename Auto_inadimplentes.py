from __future__ import annotations
import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd
from openpyxl.utils import get_column_letter
import sys
import os

FINAL_COLUMNS = [
    "Nome segurado",
    "Data contato",
    "Número apólice",
    "Account ID",
    "Seguradora",
    "Dias atraso",
    "Data vencimento",
    "Valor parcela",
    "Total inadimplente",
    "Proprietário da OP",
    "E-mail do proprietário da OP",
]

FLOW_ORDER = [
    "azos",
    "mag",
    "omint",
    "icatu",
    "metlife",
    "prudential",
]

INSURER_ALIASES: Dict[str, Tuple[str, ...]] = {
    "azos": ("azos",),
    "mag": ("mag",),
    "omint": ("omint",),
    "icatu": ("icatu",),
    "metlife": ("metlife", "met life"),
    "prudential": ("prudential", "pru"),
}

@dataclass(frozen=True)
class LayoutConfig:
    seguradora: str
    keep_letters: List[str]
    source_to_target: Dict[str, str]

CONFIGS: Dict[str, LayoutConfig] = {
    "azos": LayoutConfig(
        seguradora="Azos",
        keep_letters=["A", "I", "L", "Q"],
        source_to_target={
            "A": "Nome segurado",
            "I": "Número apólice",
            "L": "Valor parcela",
            "Q": "Data vencimento",
        },
    ),
    "mag": LayoutConfig(
        seguradora="MAG",
        keep_letters=["D", "I", "W", "AB"],
        source_to_target={
            "D": "Nome segurado",
            "I": "Número apólice",
            "W": "Data vencimento",
            "AB": "Valor parcela",
        },
    ),
    "omint": LayoutConfig(
        seguradora="Omint",
        keep_letters=["B", "F", "I", "K"],
        source_to_target={
            "B": "Nome segurado",
            "F": "Número apólice",
            "I": "Data vencimento",
            "K": "Valor parcela",
        },
    ),
    "metlife": LayoutConfig(
        seguradora="MetLife",
        keep_letters=["A", "C", "L", "M"],
        source_to_target={
            "A": "Número apólice",
            "C": "Nome segurado",
            "L": "Valor parcela",
            "M": "Data vencimento",
        },
    ),
    #==================
    "icatu": LayoutConfig(seguradora="Icatu", keep_letters=[], source_to_target={}),
    #==================
    "prudential": LayoutConfig(
        seguradora="Prudential",
        keep_letters=["E", "G", "Q", "S"],
        source_to_target={
            "E": "Número apólice",
            "G": "Nome segurado",
            "Q": "Data vencimento",
            "S": "Valor parcela",
        },
    ),
}

def excel_letter_to_index(letter: str) -> int:
    result = 0
    for ch in letter.upper().strip():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1

def metlife_prefix_policy(value: object) -> object:
    # normaliza primeiro (tira .0, notação científica, etc.)
    v = normalize_policy_number(value)
    if v is None or v is pd.NA:
        return pd.NA

    s = re.sub(r"\D", "", str(v))
    if not s:
        return pd.NA

    # regra que você pediu
    if len(s) == 5:
        return "9100" + s
    if len(s) == 6:
        return "910" + s

    return s

def normalize_policy_number(value: object) -> object:

    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NA

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if pd.isna(value):
            return pd.NA
        return f"{value:.0f}"

    text = str(value).strip()
    if not text:
        return pd.NA

    scientific = text.replace(",", ".")
    if "e" in scientific.lower():
        try:
            return f"{float(scientific):.0f}"
        except ValueError:
            pass

    digits = re.sub(r"\D", "", text)
    return digits if digits else pd.NA


def normalize_brl_number(value: object) -> object:
    """Converte valor monetário textual para número (float) quando possível."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NA
    if isinstance(value, (int, float, Decimal)):
        return float(value)

    text = str(value).strip()
    if not text:
        return pd.NA

    text = text.replace("R$", "").replace(" ", "")
    text = re.sub(r"[^0-9,.-]", "", text)

    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(Decimal(text))
    except (InvalidOperation, ValueError):
        return pd.NA

def save_excel_with_formats(df: pd.DataFrame, path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]
        columns = list(df.columns)

        if "Data vencimento" in columns:
            cidx = columns.index("Data vencimento") + 1
            col_letter = get_column_letter(cidx)
            for row in range(2, ws.max_row + 1):
                ws[f"{col_letter}{row}"].number_format = "DD/MM/YYYY"

        if "Valor parcela" in columns:
            cidx = columns.index("Valor parcela") + 1
            col_letter = get_column_letter(cidx)
            for row in range(2, ws.max_row + 1):
                ws[f"{col_letter}{row}"].number_format = 'R$ #,##0.00'

def read_sheet(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, header=0, dtype=object)

def select_and_rename(df: pd.DataFrame, config: LayoutConfig) -> pd.DataFrame:
    if not config.keep_letters:
        raise ValueError(
            f"Layout da seguradora '{config.seguradora}' não foi configurado (colunas não informadas)."
        )

    selected: Dict[str, pd.Series] = {}
    for letter in config.keep_letters:
        idx = excel_letter_to_index(letter)
        if idx >= len(df.columns):
            raise ValueError(
                f"A coluna {letter} ({idx + 1}) não existe na planilha para {config.seguradora}."
            )
        target = config.source_to_target[letter]
        selected[target] = df.iloc[:, idx]

    cleaned = pd.DataFrame(selected)

    for col in FINAL_COLUMNS:
        if col not in cleaned.columns:
            cleaned[col] = None

    cleaned["Seguradora"] = config.seguradora

    if config.seguradora == "MetLife":
        cleaned["Número apólice"] = cleaned["Número apólice"].apply(metlife_prefix_policy)
    else:
        cleaned["Número apólice"] = cleaned["Número apólice"].apply(normalize_policy_number)
    cleaned["Data vencimento"] = pd.to_datetime(
        cleaned["Data vencimento"], errors="coerce", dayfirst=True
    ).dt.date
    cleaned["Valor parcela"] = cleaned["Valor parcela"].apply(normalize_brl_number)

    today = datetime.now().date()
    cleaned["Dias atraso"] = cleaned["Data vencimento"].apply(
        lambda d: (today - d).days if pd.notnull(d) and d < today else 0
    )

    return cleaned[FINAL_COLUMNS]

def detect_insurer(file_name: str) -> Optional[str]:
    normalized = file_name.lower().replace("_", " ").replace("-", " ")
    compact = normalized.replace(" ", "")
    for key in FLOW_ORDER:
        for alias in INSURER_ALIASES.get(key, (key,)):
            if alias.lower().replace(" ", "") in compact:
                return key
    return None

#Aqui em baixo está o coisa ruim de fazer manutenção (espero que nunca de problema)
def apply_mag_rules(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    d = df.copy()


    d["Número apólice"] = d["Número apólice"].apply(normalize_policy_number)
    d["Data vencimento"] = pd.to_datetime(d["Data vencimento"], errors="coerce", dayfirst=True)
    d["Valor parcela"] = pd.to_numeric(d["Valor parcela"], errors="coerce")

    d = d[d["Número apólice"].notna() & d["Data vencimento"].notna()]


    d["_mes"] = d["Data vencimento"].dt.to_period("M")

    premio_por_mes = (
        d.groupby(["Número apólice", "_mes"], as_index=False)["Valor parcela"]
        .sum()
        .rename(columns={"Valor parcela": "Premio"})
    )

    total_por_apolice = (
        premio_por_mes.groupby("Número apólice", as_index=False)["Premio"]
        .sum()
        .rename(columns={"Premio": "Total_inad"})
    )


    d = d.merge(premio_por_mes, on=["Número apólice", "_mes"], how="left")
    d = d.merge(total_por_apolice, on="Número apólice", how="left")
    d["Valor parcela"] = d["Premio"]
    d["Total inadimplente"] = d["Total_inad"]


    d = d.sort_values("Data vencimento", ascending=False).drop_duplicates(subset=["Número apólice"], keep="first")


    d["Data vencimento"] = d["Data vencimento"].dt.date
    d = d.drop(columns=["_mes", "Premio", "Total_inad"], errors="ignore")

    return d
def apply_metlife_rules(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    d = df.copy()

    def _fix_apolice(x):
        if pd.isna(x):
            return pd.NA
        s = re.sub(r"\D", "", str(x))
        if len(s) == 5:
            return "9100" + s
        if len(s) == 6:
            return "910" + s
        return s if s else pd.NA

    # ✅ APLICA A REGRA AQUI
    d["Número apólice"] = d["Número apólice"].apply(metlife_prefix_policy)
    # resto da sua regra continua igual...
    d["Data vencimento"] = pd.to_datetime(d["Data vencimento"], errors="coerce", dayfirst=True)
    d["Valor parcela"] = pd.to_numeric(d["Valor parcela"], errors="coerce")

    d = d[d["Número apólice"].notna() & d["Data vencimento"].notna()]

    d["_mes"] = d["Data vencimento"].dt.to_period("M")  

    premio_por_mes = (
        d.groupby(["Número apólice", "_mes"], as_index=False)["Valor parcela"]
        .sum()
        .rename(columns={"Valor parcela": "Premio"})
    )

    total_por_apolice = (
        premio_por_mes.groupby("Número apólice", as_index=False)["Premio"]
        .sum()
        .rename(columns={"Premio": "Total_inad"})
    )

    d = d.merge(premio_por_mes, on=["Número apólice", "_mes"], how="left")
    d = d.merge(total_por_apolice, on="Número apólice", how="left")

    d["Valor parcela"] = d["Premio"]
    d["Total inadimplente"] = d["Total_inad"]

    d = d.sort_values("Data vencimento", ascending=False).drop_duplicates(subset=["Número apólice"], keep="first")

    d["Data vencimento"] = d["Data vencimento"].dt.date
    d = d.drop(columns=["_mes", "Premio", "Total_inad"], errors="ignore")

    return d
def apply_azos_rules(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    d = df.copy()
    d["Número apólice"] = d["Número apólice"].apply(normalize_policy_number)
    d["Data vencimento"] = pd.to_datetime(d["Data vencimento"], errors="coerce", dayfirst=True)
    d["Valor parcela"] = pd.to_numeric(d["Valor parcela"], errors="coerce")

    d = d[d["Número apólice"].notna() & d["Data vencimento"].notna()]


    d["_mes"] = d["Data vencimento"].dt.to_period("M")

    d["_mes"] = d["Data vencimento"].dt.to_period("M")
    

    premio_por_mes = (
        d.groupby(["Número apólice", "_mes"], as_index=False)["Valor parcela"]
        .sum()
        .rename(columns={"Valor parcela": "Premio"})
    )

    total_por_apolice = (
        premio_por_mes.groupby("Número apólice", as_index=False)["Premio"]
        .sum()
        .rename(columns={"Premio": "Total_inad"})
    )


    d = d.merge(premio_por_mes, on=["Número apólice", "_mes"], how="left")
    d = d.merge(total_por_apolice, on="Número apólice", how="left")


    d["Valor parcela"] = d["Premio"]
    d["Total inadimplente"] = d["Total_inad"]


    d = d.sort_values("Data vencimento", ascending=False).drop_duplicates(subset=["Número apólice"], keep="first")


    d["Data vencimento"] = d["Data vencimento"].dt.date
    d = d.drop(columns=["_mes", "Premio", "Total_inad"], errors="ignore")

    return d    



def apply_omint_rules(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    d = df.copy()
    d["Número apólice"] = d["Número apólice"].apply(normalize_policy_number)
    d["Data vencimento"] = pd.to_datetime(d["Data vencimento"], errors="coerce", dayfirst=True)
    d["Valor parcela"] = pd.to_numeric(d["Valor parcela"], errors="coerce")

    d = d[d["Número apólice"].notna() & d["Data vencimento"].notna()]


    d["_mes"] = d["Data vencimento"].dt.to_period("M")


    premio_por_mes = (
        d.groupby(["Número apólice", "_mes"], as_index=False)["Valor parcela"]
        .sum()
        .rename(columns={"Valor parcela": "Premio"})
    )


    total_por_apolice = (
        premio_por_mes.groupby("Número apólice", as_index=False)["Premio"]
        .sum()
        .rename(columns={"Premio": "Total_inad"})
    )


    d = d.merge(premio_por_mes, on=["Número apólice", "_mes"], how="left")
    d = d.merge(total_por_apolice, on="Número apólice", how="left")


    d["Valor parcela"] = d["Premio"]
    d["Total inadimplente"] = d["Total_inad"]


    d = d.sort_values("Data vencimento", ascending=False).drop_duplicates(subset=["Número apólice"], keep="first")


    d["Data vencimento"] = d["Data vencimento"].dt.date
    d = d.drop(columns=["_mes", "Premio", "Total_inad"], errors="ignore")

    return d

def apply_prudential_rules(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    d = df.copy()
    d["Número apólice"] = d["Número apólice"].apply(normalize_policy_number)
    d["Data vencimento"] = pd.to_datetime(d["Data vencimento"], errors="coerce", dayfirst=True)
    d["Valor parcela"] = pd.to_numeric(d["Valor parcela"], errors="coerce")

    d = d[d["Número apólice"].notna() & d["Data vencimento"].notna()]


    d["_mes"] = d["Data vencimento"].dt.to_period("M")


    premio_por_mes = (
        d.groupby(["Número apólice", "_mes"], as_index=False)["Valor parcela"]
        .sum()
        .rename(columns={"Valor parcela": "Premio"})
    )


    total_por_apolice = (
        premio_por_mes.groupby("Número apólice", as_index=False)["Premio"]
        .sum()
        .rename(columns={"Premio": "Total_inad"})
    )


    d = d.merge(premio_por_mes, on=["Número apólice", "_mes"], how="left")
    d = d.merge(total_por_apolice, on="Número apólice", how="left")


    d["Valor parcela"] = d["Premio"]
    d["Total inadimplente"] = d["Total_inad"]


    d = d.sort_values("Data vencimento", ascending=False).drop_duplicates(subset=["Número apólice"], keep="first")


    d["Data vencimento"] = d["Data vencimento"].dt.date
    d = d.drop(columns=["_mes", "Premio", "Total_inad"], errors="ignore")

    return d  

def process_folder(folder: Path, output: Path, recursive: bool = True) -> None:
    if not folder.exists():
        raise FileNotFoundError(f"Pasta de entrada não encontrada: {folder}")

    file_iter = folder.rglob("*") if recursive else folder.iterdir()
    files = sorted(
        [p for p in file_iter if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}],
        key=lambda p: p.name.lower(),
    )

    if not files:
        print("Nenhum arquivo .xlsx/.xlsm encontrado na pasta informada.")
        return

    output.mkdir(parents=True, exist_ok=True)
    print(f"Arquivos encontrados: {len(files)}")

    per_insurer: Dict[str, List[pd.DataFrame]] = {k: [] for k in FLOW_ORDER}
    skipped: List[str] = []

    for file in files:
        insurer_key = detect_insurer(file.name)
        if not insurer_key:
            skipped.append(f"{file.name}: seguradora não identificada no nome do arquivo")
            continue

        config = CONFIGS[insurer_key]
        try:
            original = read_sheet(file)
            cleaned = select_and_rename(original, config)
            per_insurer[insurer_key].append(cleaned)

            out_path = output / f"{file.stem}_limpo.xlsx"
            save_excel_with_formats(cleaned, out_path)
            print(f"OK  - {file.name} [{config.seguradora}] -> {out_path.name}")
        except Exception as exc:
            skipped.append(f"{file.name} [{config.seguradora}]: {repr(exc)}")

    consolidated: List[pd.DataFrame] = []
    for insurer_key in FLOW_ORDER:
        consolidated.extend(per_insurer[insurer_key])

    if consolidated:
        adjusted: List[pd.DataFrame] = []

        for part in consolidated:
            seg = str(part["Seguradora"].iloc[0]) if "Seguradora" in part.columns and not part.empty else ""

            if seg == "MAG":
                part = apply_mag_rules(part)
            if seg == "MetLife":
                part = apply_metlife_rules(part)
            if seg == "Azos":
                part = apply_azos_rules(part)
            if seg == "Omint":
                part = apply_omint_rules(part)
            if seg == "Prudential":
                part = apply_prudential_rules(part)

            adjusted.append(part)

        final_df = pd.concat(adjusted, ignore_index=True)
        final_path = output / "relatorio_final.xlsx"
        save_excel_with_formats(final_df, final_path)
        print(f"\nRelatório consolidado criado: {final_path} (linhas: {len(final_df)})")
    else:
        print("\nNenhuma planilha válida processada para consolidação.")

def app_base_dir() -> Path:
    # 1) tenta usar OneDrive corporativo (funciona pra qualquer usuário)
    od = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
    if od:
        fixed = Path(od) / "Financial Planning - Automações - Documentos" / "Seguro de Vida - Documentos" / "Inadimplencia_Automacao"
        if fixed.exists():
            return fixed

    # 2) fallback: ao lado do executável/script
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def parse_args() -> argparse.Namespace:
    base = app_base_dir()

    parser = argparse.ArgumentParser(description="Limpeza e consolidação de planilhas de seguradoras")
    parser.add_argument(
        "--input",
        default=str(base / "ENTRADA"),
        help="Pasta com os arquivos de entrada (padrão: ENTRADA ao lado do app)",
    )
    parser.add_argument(
        "--output",
        default=str(base / "SAIDA"),
        help="Pasta para salvar os arquivos limpos e o relatório final (padrão: SAIDA ao lado do app)",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Desativa busca em subpastas (por padrão a busca é recursiva)",
    )
    return parser.parse_args()

def main() -> None:
    args = parse_args()
    process_folder(Path(args.input), Path(args.output), recursive=not args.no_recursive)

if __name__ == "__main__":
    main()