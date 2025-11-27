# LEITOR-DE-PLANOS ArtPlan

Transforme planilhas de mídia instáveis em CSVs prontos para banco com um comando. Este leitor:
- Identifica cabeçalhos que mudam de lugar a cada mês.
- Lê o número do dia no cabeçalho (ex.: `S01`, `D02`) e ignora o nome do dia.
- Captura inserções em blocos, ignora linhas de total e segue lendo o que vem depois.
- Funciona para abas `OPEN TV` e `OPEN TV - GLOBO` (ou qualquer aba no mesmo formato).

## Como usar (rápido)
1) Coloque seu `.xlsx` em `INPUT/`.
2) Rode:
   ```bash
   python process_midia_open_tv.py
   ```
3) Responda aos prompts:
   - Arquivo (Enter lista os .xlsx em `INPUT/`).
   - Aba (ex.: `OPEN TV` ou `OPEN TV - GLOBO`).
   - Ano (Enter usa 2025).
   - Saída (Enter salva em `OUTPUT/insercoes_<arquivo>_<aba>.csv`).
4) Veja no final a contagem de linhas e a distribuição por mês.

## Requisitos
- Python 3.9+  
- Dependências:
  ```bash
  pip install pandas openpyxl
  ```

## Como funciona
- Detecta a linha de cabeçalho pelo conjunto `Region, Channel, TV Show, Daytime`.
- Extrai colunas de dias pelos dígitos no cabeçalho (1–31), ignorando o nome do dia da semana.
- Descobre o mês por rótulos próximos ao cabeçalho ou textos como `18/02 A 28/02`; aplica a todo o bloco até um novo cabeçalho.
- Para cada célula de dia > 0, gera uma linha de inserção com `Canal, TV_Show, Data, Horario_inicial, Horario_final`.
- Linhas que começam com “TOTAL” são puladas, mas não interrompem a leitura do bloco.

## Estrutura do projeto
- `process_midia_open_tv.py` — script CLI.
- `INPUT/` — coloque aqui os Excel.
- `OUTPUT/` — CSVs gerados (já ignorados no Git).

## Notas e dicas
- Use o nome da aba exatamente como está na planilha.
- Para outro ano, informe no prompt ou ajuste `default_year`.
- Se surgirem novos formatos, ajuste `day_columns`, `detect_month` ou `is_header`.
- CSVs em `OUTPUT/` não sobem para o repositório (`.gitignore`).
