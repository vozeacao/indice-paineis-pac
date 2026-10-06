"""Gera a planilha de equipamentos entregues por ID Governa.

Uso (no terminal do VSCode):
    pip install pandas openpyxl
    python gerar_equipamentos_entregues.py caminho/BD_Equip_Combo_UBS.xlsx
    python gerar_equipamentos_entregues.py                  # usa o .xlsx mais recente da pasta atual

A saída é salva em out/Equipamentos_Entregues_AAAA-MM-DD.xlsx
"""
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment

ABA = "AgSUS_Import"
SAIDA = Path("out")


def localizar_entrada() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    candidatos = [p for p in Path(".").glob("*.xlsx") if not p.name.startswith("~$")]
    if not candidatos:
        sys.exit("Informe o caminho do arquivo .xlsx: python gerar_equipamentos_entregues.py arquivo.xlsx")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def consolidar(df: pd.DataFrame) -> pd.DataFrame:
    base = df.groupby("ID Governa").agg(UF=("UF", "first"), Município=("Município", "first")).reset_index()

    entregues = df[df["Situação logística"] == "Entregue"]
    por_equip = (
        entregues.groupby(["ID Governa", "Nome equipamento"])["Qtd. de Itens"].sum().reset_index()
        .sort_values(["ID Governa", "Nome equipamento"])
    )
    por_equip["texto"] = por_equip["Nome equipamento"] + " (" + por_equip["Qtd. de Itens"].astype(int).astype(str) + ")"
    agg = por_equip.groupby("ID Governa").agg(
        **{
            "Qtd. tipos de equipamentos entregues": ("texto", "size"),
            "Qtd. total de itens entregues": ("Qtd. de Itens", "sum"),
            "Equipamentos entregues": ("texto", "; ".join),
        }
    ).reset_index()

    r = base.merge(agg, how="left", on="ID Governa")
    for c in ["Qtd. tipos de equipamentos entregues", "Qtd. total de itens entregues"]:
        r[c] = r[c].fillna(0).astype(int)
    r["Equipamentos entregues"] = r["Equipamentos entregues"].fillna("Nenhum equipamento entregue até o momento")
    return r


def salvar(r: pd.DataFrame, destino: Path) -> None:
    with pd.ExcelWriter(destino, engine="openpyxl") as w:
        r.to_excel(w, index=False, sheet_name="Entregues por ID")
        ws = w.sheets["Entregues por ID"]
        for col, largura in zip("ABCDEF", [12, 5, 30, 14, 14, 120]):
            ws.column_dimensions[col].width = largura
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for (cel,) in ws.iter_rows(min_row=2, min_col=6, max_col=6):
            cel.alignment = Alignment(wrap_text=True, vertical="top")


def main() -> None:
    entrada = localizar_entrada()
    print(f"Lendo {entrada} (aba {ABA})...")
    df = pd.read_excel(entrada, sheet_name=ABA)

    r = consolidar(df)

    # conferência: itens entregues na saída == itens entregues na base
    esperado = df.loc[df["Situação logística"] == "Entregue", "Qtd. de Itens"].sum()
    assert r["Qtd. total de itens entregues"].sum() == esperado, "Soma de itens não confere com a base"

    SAIDA.mkdir(exist_ok=True)
    destino = SAIDA / f"Equipamentos_Entregues_{date.today():%Y-%m-%d}.xlsx"
    salvar(r, destino)
    print(f"OK: {len(r)} IDs, {int(esperado)} itens entregues -> {destino}")
    sem_status = df["Situação logística"].isna().sum()
    if sem_status:
        print(f"Atenção: {sem_status} linhas sem Situação logística (ignoradas).")


if __name__ == "__main__":
    main()
