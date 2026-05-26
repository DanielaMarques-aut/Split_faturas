import pandas as pd
from pathlib import Path
import sys

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
# Muda estes caminhos para os teus ficheiros reais
FATURAS_FILE = "Módulo de Faturas.xlsx"
PMO_FILE     = "Listagem PMO Abril.xlsx"   # muda o mês conforme necessário
OUTPUT_DIR   = "output_por_pm"
# ─────────────────────────────────────────────────────────────────────────────

def load_faturas(path):
    """Lê todas as sheets do Módulo de Faturas e concatena."""
    xl = pd.ExcelFile(path)
    sheets = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, dtype=str)
        df["_source_sheet"] = sheet
        sheets.append(df)
    return pd.concat(sheets, ignore_index=True)

def load_pmo(path):
    """Lê a listagem de PMO — assume primeira sheet."""
    return pd.read_excel(path, dtype=str)

def match_wbs_to_contract(wbs: str, contract: str) -> bool:
    """
    Verifica se o WBS Element pertence ao contrato.
    Cobre: 5201900075.x.x  5201900075A.x.x  5201900075AA.x.x
    Lógica: o WBS começa com o número do contrato (antes do primeiro ponto).
    """
    if pd.isna(wbs) or pd.isna(contract):
        return False
    wbs = str(wbs).strip()
    contract = str(contract).strip()
    # O prefixo do WBS (antes do primeiro ponto) deve começar com o contrato
    prefix = wbs.split(".")[0]
    return prefix.startswith(contract)

def main():
    # Verificar ficheiros
    for f in [FATURAS_FILE, PMO_FILE]:
        if not Path(f).exists():
            print(f"ERRO: Ficheiro não encontrado: {f}")
            print("Coloca o script na mesma pasta dos ficheiros Excel.")
            sys.exit(1)

    print("A ler Módulo de Faturas...")
    faturas = load_faturas(FATURAS_FILE)

    print("A ler Listagem PMO...")
    pmo = load_pmo(PMO_FILE)

    # Identificar colunas chave — ajusta se os nomes forem ligeiramente diferentes
    wbs_col      = next(c for c in faturas.columns if "WBS" in c.upper())
    contrato_col = next(c for c in pmo.columns if "CONTRATO" in c.upper())
    pm_col       = next(c for c in pmo.columns if "CHEFE" in c.upper() or "PM" in c.upper())

    print(f"  Coluna WBS em faturas: '{wbs_col}'")
    print(f"  Coluna Contrato em PMO: '{contrato_col}'")
    print(f"  Coluna PM em PMO: '{pm_col}'")

    # Criar pasta de output
    out = Path(OUTPUT_DIR)
    out.mkdir(exist_ok=True)

    # Para cada PM, filtrar faturas dos seus contratos
    pms = pmo[[contrato_col, pm_col]].dropna(subset=[pm_col])
    pm_groups = pms.groupby(pm_col)[contrato_col].apply(list)

    stats = {}
    for pm_name, contratos in pm_groups.items():
        mask = faturas[wbs_col].apply(
            lambda wbs: any(match_wbs_to_contract(wbs, c) for c in contratos)
        )
        df_pm = faturas[mask].drop(columns=["_source_sheet"])

        if df_pm.empty:
            stats[pm_name] = 0
            continue

        # Nome do ficheiro: remove caracteres inválidos
        safe_name = "".join(c for c in pm_name if c.isalnum() or c in " _-").strip()
        out_path = out / f"{safe_name}.xlsx"

        df_pm.to_excel(out_path, index=False)
        stats[pm_name] = len(df_pm)
        print(f"  ✓ {pm_name}: {len(df_pm)} faturas → {out_path.name}")

    # Faturas sem PM atribuído
    all_contratos = pmo[contrato_col].dropna().tolist()
    mask_all = faturas[wbs_col].apply(
        lambda wbs: any(match_wbs_to_contract(wbs, c) for c in all_contratos)
    )
    sem_pm = faturas[~mask_all].drop(columns=["_source_sheet"])
    if not sem_pm.empty:
        sem_pm.to_excel(out / "SEM_PM_ATRIBUIDO.xlsx", index=False)
        print(f"\n  ⚠ {len(sem_pm)} faturas sem PM atribuído → SEM_PM_ATRIBUIDO.xlsx")

    print(f"\nConcluído. Ficheiros gerados em: {out.resolve()}")
    print(f"Total de PMs com faturas: {sum(1 for v in stats.values() if v > 0)}")

if __name__ == "__main__":
    main()