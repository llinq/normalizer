# Arquitetura e Fluxo do Código

Este documento descreve como o código do **normalizer** funciona internamente, do ponto de entrada até a geração do resultado JSON.

---

## Visão Geral

```
main.py (CLI)
    └── normalizer.compare()          ← ponto de entrada da biblioteca
            ├── readers.py            ← leitura e normalização das planilhas
            ├── comparator.py         ← lógica de comparação
            └── models.py             ← estruturas de dados do resultado
```

---

## Estrutura de Módulos

| Arquivo | Responsabilidade |
|---|---|
| `main.py` | CLI – parseia argumentos, chama `compare()`, imprime JSON |
| `normalizer/__init__.py` | Re-exporta os símbolos públicos da biblioteca |
| `normalizer/readers.py` | Carrega workbooks, lê cabeçalhos, normaliza fórmulas |
| `normalizer/comparator.py` | Orquestra toda a comparação; contém `compare()` e os helpers internos |
| `normalizer/models.py` | Define `DivergenceType`, `Divergence` e `ComparisonResult` |

---

## Fluxo Detalhado

### 1. Entrada (CLI — `main.py`)

```
$ python main.py template.xlsx user_file.xlsx [opções]
```

1. `build_parser()` define os argumentos aceitos:
   - `template` e `user_file` — caminhos obrigatórios
   - `--header-row N` (padrão `1`)
   - `--no-formulas` — desativa a comparação de fórmulas
   - `--check-types` — ativa comparação de tipos de dados
   - `--max-formula-rows N` — limita quantas linhas de dados são verificadas
   - `--output FILE` — grava o JSON em arquivo em vez de stdout
2. Valida se os arquivos existem e têm extensão `.xlsx`/`.xlsm`.
3. Chama `compare(...)` e serializa o `ComparisonResult` retornado para JSON.
4. Retorna exit-code **0** (sem divergências), **1** (divergências encontradas) ou **2** (arquivo não encontrado).

---

### 2. Ponto de Entrada da Biblioteca — `compare()` (`comparator.py`)

```python
result = compare(template_path, user_path, header_row, check_formulas, check_data_types, max_formula_rows)
```

Sequência de operações:

```
compare()
 ├─ load_workbook_with_formulas(template_path)   # carrega fórmulas como string
 ├─ load_workbook_with_formulas(user_path)
 ├─ get_sheet_names(template_wb)
 ├─ get_sheet_names(user_wb)
 ├─ _compare_sheets()                            # nível 1 – abas
 └─ para cada aba em comum:
      ├─ get_headers(tmpl_sheet, header_row)
      ├─ get_headers(user_sheet, header_row)
      ├─ _compare_columns()                      # nível 2 – colunas
      └─ _compare_cells()                        # nível 3 – células
```

---

### 3. Leitura das Planilhas — `readers.py`

#### `load_workbook_with_formulas(path)`
Usa `openpyxl.load_workbook(data_only=False)` para preservar as strings das fórmulas (ex.: `=SOMA(A1:A10)`) em vez de substituí-las pelos valores calculados.

#### `get_headers(sheet, header_row)`
Lê a linha de cabeçalho e retorna um dicionário `{índice_coluna: nome_cabeçalho}`. Nomes são normalizados com `.strip()`. Colunas sem valor recebem `None`.

#### `is_formula(cell)`
Retorna `True` se `cell.value` é uma string que começa com `=`.

#### `cell_address(row, col)`
Converte coordenadas numéricas `(linha, coluna)` para endereço Excel como `A1`, `C11`.

---

### 4. Comparação de Abas — `_compare_sheets()`

| Condição | Tipo de Divergência |
|---|---|
| Aba no template, ausente no arquivo do usuário | `MISSING_SHEET` |
| Aba no arquivo do usuário, ausente no template | `EXTRA_SHEET` |

Somente as **abas em comum** avançam para as etapas seguintes.

---

### 5. Comparação de Colunas — `_compare_columns()`

Compara os nomes de cabeçalho (não as posições numéricas) entre as duas abas correspondentes:

| Condição | Tipo de Divergência |
|---|---|
| Coluna no template, ausente no usuário | `MISSING_COLUMN` |
| Coluna no usuário, ausente no template | `EXTRA_COLUMN` |
| Mesmas colunas, mas em ordem diferente | `COLUMN_ORDER` |

A ordem é verificada comparando a sequência de nomes comuns, ordenados pela posição em cada planilha.

---

### 6. Comparação de Células — `_compare_cells()`

#### Mapeamento de colunas

Primeiro é construído um dicionário `col_mapping: {tmpl_col_idx → user_col_idx}` baseado nos **nomes** dos cabeçalhos. Isso garante que colunas comparadas são semanticamente equivalentes, mesmo que estejam em posições numéricas diferentes.

#### Iteração sobre linhas

```
para cada linha de dados (header_row + 1 … tmpl_max_row):
    ignorar linhas completamente vazias no template
    para cada par (tmpl_col, user_col) em col_mapping:
        obter tmpl_cell e user_cell
        → verificar fórmulas (se check_formulas=True)
        → verificar tipos de dado (se check_data_types=True)
```

Se a planilha do usuário tiver menos linhas que o template, as linhas extras do template são ignoradas silenciosamente.

#### Verificações de fórmula

| Situação | Tipo de Divergência |
|---|---|
| Template tem fórmula; usuário tem valor não-nulo | `FORMULA_MISSING` |
| Ambos têm fórmula, mas a estrutura difere | `FORMULA_MISMATCH` |
| Template tem valor; usuário tem fórmula | `UNEXPECTED_FORMULA` |

A comparação de fórmulas **não é literal** — ela usa `normalize_formula()` antes de comparar (veja seção 7).

#### Verificação de tipo de dado

Ativada por `--check-types`. Só se aplica a células sem fórmula. Compara os tipos Python dos valores (`int`, `float`, `str`, `bool`, etc.) via `_cell_type_label()`. Divergência só é gerada quando ambas as células têm valor.

---

### 7. Normalização de Fórmulas — `normalize_formula()` (`readers.py`)

Esta função é o coração da detecção de `FORMULA_MISMATCH`. Ela converte a fórmula em uma forma canônica que permite comparação **estrutural** (sem ser enganada por diferenças esperadas como número de linha ou posição de coluna).

```python
normalize_formula(formula, current_col=0, current_row=0) → str
```

Transformações aplicadas **em ordem**:

#### Passo 0 – Caixa alta
Toda a fórmula é convertida para maiúsculas: `=soma(a1)` → `=SOMA(A1)`.

#### Passo 0b – Remoção de aspas em nomes de abas
Excel insere aspas simples em nomes de abas que começam com dígito ou contêm caracteres especiais:
`'3_IPI-PIS-Cofins'!C1` → `3_IPI-PIS-COFINS!C1`

#### Passo 1 – Referências cruzadas de abas (quando `current_col > 0`)
Referências do tipo `NOME_ABA!$B$38` são identificadas primeiro (antes de qualquer relativização). Os números de linha **dentro** delas são substituídos por `#` (valor absoluto fixo), pois apontam para posições absolutas em outra aba e não devem ser relativizados em relação à linha atual.

Exemplo: `3_IPI!$C$5` → `3_IPI!$C#`

#### Passo 2 – Relativização de colunas e linhas da mesma aba (quando `current_col > 0`)
Referências de coluna relativas (não precedidas por `$` ou `!`) são convertidas em **deslocamentos** relativos à coluna atual:

```
Coluna da fórmula: C (índice 3)
Referência A11  → coluna A (1), offset = 1-3 = -2
                   linha 11, offset relativo à linha atual
Resultado: [-2][+0]   (se current_row=11)
```

Isso garante que a mesma fórmula copiada para uma coluna diferente continue comparando como igual, ao mesmo tempo em que detecta mudanças reais de referência de coluna.

#### Passo 3 – Colunas absolutas com linha relativa (quando `current_col > 0`)
Referências do tipo `$D$2` que sobreviveram aos passos anteriores têm apenas o número de linha convertido para offset relativo: `$D$2` (em row 5) → `$D[-3]`.

#### Sem contexto de coluna (`current_col = 0`)
Sem `current_col`, o número de linha ainda é relativizado (se `current_row > 0`) ou substituído por `#` (se também sem `current_row`).

---

#### Exemplos práticos de normalização

| Fórmula original | Célula | Resultado normalizado |
|---|---|---|
| `=SE(A11="";"";"ok")` | linha 11, col 3 | `=SE([-2][+0]="";"";\"OK\")` |
| `=SE(A10="";"";"ok")` | linha 11, col 3 | `=SE([-2][-1]="";"";\"OK\")` ← divergência detectada |
| `=PROCV(E11;3_IPI!C:R;2;FALSO())` | linha 11, col 3 | `=[+2][+0];3_IPI!C#:R#;2;FALSO())` |
| `=SOMA(B2:B100)` | linha 2, col 3 | offset de linha `[+0]` e `[+98]` |
| `'Sheet1'!$B$38` | qualquer linha | `SHEET1!$B#` |

---

### 8. Modelos de Dados — `models.py`

#### `DivergenceType` (Enum)
Todos os tipos de divergência possíveis:
- Nível de aba: `MISSING_SHEET`, `EXTRA_SHEET`
- Nível de coluna: `MISSING_COLUMN`, `EXTRA_COLUMN`, `COLUMN_ORDER`
- Nível de célula: `FORMULA_MISSING`, `FORMULA_MISMATCH`, `UNEXPECTED_FORMULA`, `DATA_TYPE_MISMATCH`

#### `Divergence` (dataclass)
Representa uma única divergência encontrada:
- `type` — qual `DivergenceType`
- `sheet` — nome da aba
- `location` — endereço da célula (ex.: `C11`) ou `None` para divergências estruturais
- `description` — mensagem legível
- `template_value` / `user_value` — valores originais para exibição

#### `ComparisonResult` (dataclass)
Agrega todas as divergências:
- `divergences: list[Divergence]`
- `has_divergences` — `True` se a lista não está vazia
- `by_type(dtype)` — filtra divergências por tipo
- `to_dict()` — serializa para o formato JSON de saída

---

## Diagrama de Fluxo Completo

```
┌─────────────────────────────────────────────────────────────────────┐
│  main.py (CLI)                                                      │
│  argparse → valida arquivos → compare() → JSON → stdout/arquivo     │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  compare()  [comparator.py]                                         │
│                                                                     │
│  1. load_workbook_with_formulas(template)                           │
│  2. load_workbook_with_formulas(user)                               │
│  3. _compare_sheets()  ──► MISSING_SHEET / EXTRA_SHEET              │
│  4. para cada aba em comum:                                         │
│     a. get_headers(template_sheet)                                  │
│     b. get_headers(user_sheet)                                      │
│     c. _compare_columns() ──► MISSING_COLUMN / EXTRA_COLUMN /      │
│                                COLUMN_ORDER                         │
│     d. _compare_cells()                                             │
│        ├─ monta col_mapping por nome de cabeçalho                   │
│        ├─ itera linhas de dados                                     │
│        │   ├─ pula linhas totalmente vazias no template             │
│        │   └─ para cada célula mapeada:                             │
│        │       ├─ is_formula(tmpl) + is_formula(user)               │
│        │       ├─ [check_formulas] normalize_formula() ambas        │
│        │       │   ├─ tmpl_norm == user_norm → OK                   │
│        │       │   └─ tmpl_norm != user_norm → FORMULA_MISMATCH     │
│        │       │   (ou FORMULA_MISSING / UNEXPECTED_FORMULA)        │
│        │       └─ [check_data_types] _cell_type_label()             │
│        │           └─ tipos diferentes → DATA_TYPE_MISMATCH         │
│        └─ retorna ComparisonResult preenchido                        │
└─────────────────────────────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  normalize_formula()  [readers.py]                                  │
│                                                                     │
│  1. upper-case                                                      │
│  2. remover aspas de nomes de abas                                  │
│  3. (se current_col > 0)                                            │
│     a. substituir refs cruzadas de aba → col mantida, linha → #     │
│     b. relativizar col relativa → [±col_offset]                     │
│        + relativizar linha      → [±row_offset]  (se current_row>0) │
│     c. relativizar linha de col absoluta ($D$2 → $D[±row_offset])   │
│  4. (sem current_col) substituir linha por # ou offset              │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Tipos de Divergência e Quando São Geradas

| Tipo | Quando ocorre |
|---|---|
| `MISSING_SHEET` | Aba do template não existe no arquivo do usuário |
| `EXTRA_SHEET` | Aba existe no usuário mas não no template |
| `MISSING_COLUMN` | Cabeçalho do template ausente no usuário |
| `EXTRA_COLUMN` | Cabeçalho do usuário ausente no template |
| `COLUMN_ORDER` | Mesmas colunas, ordem diferente |
| `FORMULA_MISSING` | Template tem fórmula, usuário tem valor não-nulo |
| `FORMULA_MISMATCH` | Ambos têm fórmula, mas estrutura normalizada difere |
| `UNEXPECTED_FORMULA` | Template tem valor, usuário tem fórmula |
| `DATA_TYPE_MISMATCH` | Tipos Python dos valores diferem (requer `--check-types`) |
