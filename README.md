# Split Faturas

Ferramenta simples para dividir o relatório de faturas por PMO, gerando ficheiros Excel separados.

## O que faz

O script `slpit_faturas.py`:
- lê o ficheiro `Módulo de Faturas.xlsx`
- lê o ficheiro `Listagem PMO Abril.xlsx`
- divide as faturas por PM (e, para PMs específicos, por cliente)
- escreve resultados em Excel na pasta `Faturas/`
- aplica cores às linhas com base no estado de pagamento

## Requisitos

- Python 3.12+
- Dependências:
  - `openpyxl`
  - `pandas`

Instalação:

```bash
python -m pip install openpyxl pandas
```

## Uso

1. Coloque os ficheiros de entrada na raiz do projeto:
   - `Módulo de Faturas.xlsx`
   - `Listagem PMO Abril.xlsx`

2. Ajuste a configuração no topo de `slpit_faturas.py` se necessário:
   - `FATURAS_FILE`
   - `PMO_FILE`
   - `OUTPUT_DIR`
   - `SPLIT_BY_CLIENT_PMS`

3. Execute o script:

```bash
python slpit_faturas.py
```

4. Os ficheiros gerados serão criados em `Faturas/`.

## Detalhes de processamento

- As faturas são agrupadas por PM usando a coluna `WBS` do ficheiro de faturas.
- Para PMs listados em `SPLIT_BY_CLIENT_PMS`, o script também divide as faturas por cliente.
- Linhas com pagamento confirmado (`PAGO`) e não vencidas (`NÃO`) ficam com cor verde.
- Linhas não pagas (`NÃO PAGO`) ou vencidas (`SIM`) ficam com cor vermelha.

## Observações

- O arquivo de saída `SEM_PM_ATRIBUIDO.xlsx` contém faturas que não correspondem a nenhum contrato PMO.
- O script espera que as colunas `WBS`, `PAGO` e `VENCIDA` existam no ficheiro de faturas.
- O script espera que as colunas `CONTRATO`, `CHEFE` e `CLIENTE` existam no ficheiro PMO.
