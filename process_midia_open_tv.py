import os
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


# Month aliases to handle accents/english variants.
MONTH_MAP: Dict[str, int] = {
    "JAN": 1,
    "FEV": 2,
    "FEB": 2,
    "MAR": 3,
    "ABR": 4,
    "APR": 4,
    "MAI": 5,
    "MAIO": 5,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AGO": 8,
    "AUG": 8,
    "SET": 9,
    "SEP": 9,
    "OUT": 10,
    "OCT": 10,
    "NOV": 11,
    "DEZ": 12,
    "DEC": 12,
}


def normalize(text: str) -> str:
    """Uppercase ASCII-only version to simplify matching."""
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ASCII", "ignore")
        .decode("ASCII")
        .upper()
    )


def detect_month(text: object) -> Optional[int]:
    """Try to infer month number from free text or date range like 18/02 A 28/02."""
    if not isinstance(text, str):
        return None
    norm = normalize(text)
    for key, month in MONTH_MAP.items():
        if key in norm:
            return month
    match = re.search(r"(\\d{1,2})/(\\d{1,2})", norm)
    if match:
        return int(match.group(2))
    return None


def is_header(row: pd.Series) -> bool:
    """Check if the row contains the expected column labels."""
    vals = {str(v).strip().upper() for v in row if isinstance(v, str)}
    return {"REGION", "CHANNEL", "TV SHOW", "DAYTIME"}.issubset(vals)


def find_headers(df: pd.DataFrame) -> List[int]:
    """Return indices of header rows."""
    headers = [idx for idx, row in df.iterrows() if is_header(row)]
    headers.append(len(df))
    return headers


def day_columns(header_row: pd.Series) -> List[Tuple[int, int]]:
    """Map column index to day number extracted from header cells (e.g., 'S01', 'S 01')."""
    cols: List[Tuple[int, int]] = []
    pattern = re.compile("(\\d{1,2})\\s*$")
    for col, label in header_row.items():
        # Day columns live before the metrics (Ins, GRP, etc.), which start ~col 42.
        if col >= 42 or not isinstance(label, str):
            continue
        match = pattern.search(label.strip())
        if match:
            cols.append((col, int(match.group(1))))
    return cols


@dataclass
class Record:
    Canal: str
    TV_Show: str
    Data: str
    Horario_inicial: str
    Horario_final: str


def parse_sheet(df: pd.DataFrame, default_year: int) -> List[Record]:
    """Parse a sheet, handling moving headers and month markers."""
    headers = find_headers(df)
    records: List[Record] = []
    last_month: Optional[int] = None

    for i in range(len(headers) - 1):
        h_idx, end_idx = headers[i], headers[i + 1]

        # Look a few rows before the header for month hints.
        for back in range(max(0, h_idx - 5), h_idx):
            for val in df.iloc[back]:
                month = detect_month(val)
                if month:
                    last_month = month

        header_row = df.iloc[h_idx]
        for val in header_row:
            month = detect_month(val)
            if month:
                last_month = month

        block_month = last_month
        days = day_columns(header_row)
        block = df.iloc[h_idx + 1 : end_idx]

        # Try to find month inside block if still unknown.
        if block_month is None:
            for _, row in block.iterrows():
                for val in row:
                    month = detect_month(val)
                    if month:
                        block_month = month
                        last_month = month
                        break
                if block_month:
                    break

        # Skip blocks without day columns or month.
        if not days or block_month is None:
            continue

        for _, row in block.iterrows():
            row_month = block_month

            # If we still don't know the month, try to infer from this row.
            if row_month is None:
                for val in row:
                    month = detect_month(val)
                    if month:
                        last_month = month
                        row_month = month
                        block_month = month
                        break

            # Totals rows: if they carry a month label, update block_month for subsequent rows.
            if isinstance(row[1], str) and row[1].strip().upper().startswith("TOTAL"):
                month = detect_month(row[1])
                # Only set if we don't already know the block month (avoid rolling back).
                if month and block_month is None:
                    last_month = month
                    block_month = month
                continue

            # Use the most recent month hint available.
            row_month = row_month or block_month

            channel, show, daytime = row[2], row[3], row[4]
            if (
                pd.isna(channel)
                or pd.isna(show)
                or pd.isna(daytime)
                or not isinstance(daytime, str)
                or "-" not in daytime
            ):
                continue

            if row_month is None:
                continue

            try:
                start_time, end_time = [t.strip() for t in daytime.split("-")[:2]]
            except Exception:
                continue

            for col, day_num in days:
                val = row[col]
                if pd.notna(val):
                    try:
                        if float(val) > 0:
                            for _ in range(int(val)):
                                records.append(
                                    Record(
                                        Canal=str(channel),
                                        TV_Show=str(show),
                                        Data=f"{default_year:04d}-{row_month:02d}-{day_num:02d}",
                                        Horario_inicial=start_time,
                                        Horario_final=end_time,
                                    )
                                )
                    except Exception:
                        # Non-numeric cell; ignore.
                        pass

    return records


def main() -> None:
    input_dir = Path("INPUT")
    output_dir = Path("OUTPUT")
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    print(f"Coloque seus .xlsx em: {input_dir.resolve()}")
    file_path = input(f"Caminho do arquivo Excel (Enter para listar {input_dir}): ").strip()
    if not file_path:
        files = list(input_dir.glob("*.xlsx"))
        if not files:
            print("Nenhum .xlsx encontrado em INPUT. Informe um caminho completo.", file=sys.stderr)
            sys.exit(1)
        print("Arquivos encontrados em INPUT:")
        for idx, f in enumerate(files, 1):
            print(f"{idx}) {f.name}")
        choice = input("Escolha o número do arquivo: ").strip()
        try:
            file_idx = int(choice) - 1
            file_path = str(files[file_idx])
        except Exception:
            print("Escolha inválida.", file=sys.stderr)
            sys.exit(1)

    if not os.path.exists(file_path):
        print(f"Arquivo não encontrado: {file_path}", file=sys.stderr)
        sys.exit(1)

    sheet_name = input("Nome da aba (ex: 'OPEN TV', 'OPEN TV - GLOBO'): ").strip()
    if not sheet_name:
        print("Nenhuma aba informada. Encerrando.", file=sys.stderr)
        sys.exit(1)

    year_raw = input("Ano (YYYY, Enter para padrão 2025): ").strip()
    year = 2025
    if year_raw:
        try:
            year = int(year_raw)
        except ValueError:
            print("Ano inválido, usando 2025.", file=sys.stderr)

    print("Lendo planilha...")
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    print("Processando inserções...")
    records = parse_sheet(df, default_year=year)
    if not records:
        print("Nenhuma inserção encontrada. Verifique a aba ou formato.", file=sys.stderr)
        sys.exit(1)

    out_default_name = f"insercoes_{Path(file_path).stem}_{sheet_name.replace(' ', '_').lower()}.csv"
    out_path_input = input(f"Salvar CSV em (padrão OUTPUT/{out_default_name}): ").strip()
    if out_path_input:
        out_path = Path(out_path_input)
    else:
        out_path = output_dir / out_default_name

    out_df = pd.DataFrame([r.__dict__ for r in records])
    out_df.sort_values(["Data", "Canal", "TV_Show", "Horario_inicial"], inplace=True)
    out_df.to_csv(out_path, index=False)

    summary = (
        out_df["Data"].str.slice(0, 7).value_counts().sort_index().to_dict()
    )
    print(f"Salvo {len(out_df)} linhas em {out_path}")
    print("Distribuição por mês:", summary)


if __name__ == "__main__":
    main()
