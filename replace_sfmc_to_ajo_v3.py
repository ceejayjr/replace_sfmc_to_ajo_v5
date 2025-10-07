import re
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime

# =============================
# CONFIGURAÇÃO
# =============================
html_input  = sys.argv[1] if len(sys.argv) > 1 else "input.html"
html_output = sys.argv[2] if len(sys.argv) > 2 else "output.html"
excel_path  = sys.argv[3] if len(sys.argv) > 3 else "SFMCtoAJOComparision.xlsx"

# =============================
# HELPERS
# =============================
def pick_cols(df):
    """
    Descobre as colunas SFMC/AJO mesmo que mudem o header.
    """
    cols = {str(c).strip().lower(): c for c in df.columns}
    sfmc_col = next((v for k,v in cols.items() if "sfmc" in k), df.columns[0])
    ajo_col  = next((v for k,v in cols.items() if "ajo"  in k), df.columns[1])
    return sfmc_col, ajo_col

def build_flex_regex(sfmc_snippet: str) -> re.Pattern:
    """
    Gera regex tolerante a espaços e que permita '=' com espaços.
    Para aplicar substituições do de-para no HTML.
    """
    s = re.escape(str(sfmc_snippet).strip())
    s = re.sub(r'\\[ \t\r\n\f]+', r'\\s*', s)
    s = s.replace(r'\=', r'\s*=\s*')
    return re.compile(s, flags=re.IGNORECASE | re.DOTALL)

# =============================
# PLANILHA → VARIÁVEIS CONHECIDAS
# =============================
def extract_var_expressions(df):
    """
    Lê a planilha e extrai mapeamentos para @variáveis:
      - Se a célula AJO tiver "{% let Var = ... %}" → usar "Var"
      - Se tiver "{{ expr }}" → usar "expr"
      - Se tiver "profile." ou "context." → usar esse path
    Retorna:
      var_expr: dict var_lower -> expressão Liquid (nome/expr a ser usada)
      covered : set com variáveis que NÃO devem gerar warning
    """
    var_expr = {}
    covered = set()

    let_re   = re.compile(r'{%\s*let\s+([A-Za-z_]\w*)\s*=\s*\(?\s*(.*?)\s*\)?\s*%}', re.IGNORECASE | re.DOTALL)
    print_re = re.compile(r'{{\s*(.*?)\s*}}', re.DOTALL)
    path_re  = re.compile(r'\b(?:profile|context)\.[A-Za-z0-9_\.\[\]]+', re.IGNORECASE)

    for _, row in df.iterrows():
        sfmc = str(row["SFMC"] or "")
        ajo  = str(row["AJO"] or "")

        vars_in_sfmc = re.findall(r'@(\w+)', sfmc)
        if not vars_in_sfmc:
            continue

        expr = None
        m_let = let_re.search(ajo)
        if m_let:
            expr = m_let.group(1).strip()  # usar o nome da variável definida pelo let
        else:
            m_print = print_re.search(ajo)
            if m_print:
                expr = m_print.group(1).strip()
            else:
                m_path = path_re.search(ajo)
                if m_path:
                    expr = m_path.group(0).strip()

        for var in vars_in_sfmc:
            vlow = var.lower()
            if expr:
                var_expr[vlow] = expr
                covered.add(vlow)
            elif any(x in ajo for x in ["profile.", "context.", "fragment", "{{", "{%"]):
                # tem algo em AJO (ainda que não extraímos nome final), não alertar
                covered.add(vlow)

    return var_expr, covered

def resolve_expr(var_name: str, var_expr: dict) -> str:
    """
    Retorna a expressão Liquid a ser usada para @var.
    Se não houver mapeamento, retorna o próprio nome (para sabermos que NÃO devemos converter).
    """
    return var_expr.get(var_name.lower(), var_name)

# =============================
# TRADUÇÃO DE CONDIÇÕES (AMPScript → Liquid AJO)
# =============================
def translate_condition_to_liquid(cond: str, warnings: list, var_expr: dict, covered_vars: set) -> str:
    """
    Traduz:
      NOT EMPTY(@x), EMPTY(@x), CONTAINS(@x,'y'), @x == 'y', @x != 'y', @x (truthy)
    Fecha com {%/if%}.
    """
    cond_original = cond
    cond_l = cond.strip()

    # Coletar variáveis citadas na condição (para avisos)
    for v in re.findall(r'@(\w+)', cond_l, flags=re.IGNORECASE):
        if v.lower() not in covered_vars and v.lower() not in {"email","firstname","lastname","country"}:
            warnings.append(v)

    re_not_empty = re.compile(r'not\s*empty\s*\(\s*@(\w+)\s*\)', re.IGNORECASE)
    re_empty     = re.compile(r'empty\s*\(\s*@(\w+)\s*\)', re.IGNORECASE)
    re_contains  = re.compile(r'contains\s*\(\s*@(\w+)\s*,\s*[\'"]([^\'"]+)[\'"]\s*\)', re.IGNORECASE)
    re_cmp       = re.compile(r'@(\w+)\s*(==|!=)\s*[\'"]([^\'"]+)[\'"]', re.IGNORECASE)
    re_truthy    = re.compile(r'^\s*@(\w+)\s*$', re.IGNORECASE)

    def learn_expr(var_name: str) -> str:
        expr = resolve_expr(var_name, var_expr)
        # se expr mapeada != nome cru, marcar como coberta
        if expr != var_name:
            covered_vars.add(var_name.lower())
        return expr

    m = re_not_empty.search(cond_l)
    if m:
        e = learn_expr(m.group(1))
        return "{% if length(" + e + ") > 0 %}"

    m = re_empty.search(cond_l)
    if m:
        e = learn_expr(m.group(1))
        return "{% if length(" + e + ") == 0 %}"

    m = re_contains.search(cond_l)
    if m:
        e = learn_expr(m.group(1))
        val = m.group(2)
        return "{% if contains(" + e + ", '" + val + "') %}"

    m = re_cmp.search(cond_l)
    if m:
        e = learn_expr(m.group(1))
        op = "==" if m.group(2) == "==" else "!="
        val = m.group(3)
        return "{% if " + e + " " + op + " '" + val + "' %}"

    m = re_truthy.search(cond_l)
    if m:
        e = learn_expr(m.group(1))
        return "{% if " + e + " %}"

    # ELSEIF (fallback simples)
    if cond_l.lower().startswith("elseif "):
        inner = cond_l[7:].strip()
        trans = translate_condition_to_liquid(inner, warnings, var_expr, covered_vars)
        return trans.replace("{% if", "{% elseif")

    return "<!-- Untranslated condition: " + cond_original + " -->"

# =============================
# CONVERSÃO DE BLOCOS IF/ELSE/ENDIF
# =============================
def convert_ampscript_if_blocks(html_text: str, warnings: list, var_expr: dict, covered_vars:set) -> str:
    """
    Converte:
      %%[ IF ... THEN ]%% ... %%[ ELSE ]%% ... %%[ ENDIF ]%%
      %%[ IF ... ]%%      ... %%[ ELSE ]%% ... %%[ ENDIF ]%%
    """
    def repl(match):
        cond = (match.group(1) or "").strip()
        then_block = match.group(2) or ""
        else_block = match.group(4) or ""

        cond_liquid = translate_condition_to_liquid(cond, warnings, var_expr, covered_vars)

        # Substitui prints %%=v(@x)=%% -> {{ expr }} SOMENTE se houver mapeamento
        def _print_sub(mv):
            var = mv.group(1)
            expr = resolve_expr(var, var_expr)
            if expr != var:
                return "{{ " + expr + " }}"
            return mv.group(0)  # sem mapeamento → deixa AMPScript; será comentado e logado depois

        then_block = re.sub(r'%%=v\(@(\w+)\)=%%', _print_sub, then_block, flags=re.IGNORECASE)
        else_block = re.sub(r'%%=v\(@(\w+)\)=%%', _print_sub, else_block, flags=re.IGNORECASE)

        if else_block.strip():
            return cond_liquid + then_block + "{% else %}" + else_block + "{%/if%}"
        else:
            return cond_liquid + then_block + "{%/if%}"

    # Aceita IF com ou sem THEN
    pattern = re.compile(
        r'%%\[\s*IF\s+(.+?)(?:\s+THEN)?\s*\]%%(.*?)((?:%%\[\s*ELSE\s*\]%%(.*?))?)%%\[\s*ENDIF\s*\]%%',
        re.IGNORECASE | re.DOTALL
    )
    return pattern.sub(repl, html_text)

# =============================
# REPLACE GLOBAL DE PRINTS (apenas mapeados)
# =============================
def replace_all_prints(html_text: str, var_expr: dict) -> str:
    def _print_sub(mv):
        var = mv.group(1)
        expr = resolve_expr(var, var_expr)
        if expr != var:
            return "{{ " + expr + " }}"
        return mv.group(0)  # sem mapeamento → deixa AMPScript; será comentado e logado depois
    return re.sub(r'%%=v\(@(\w+)\)=%%', _print_sub, html_text, flags=re.IGNORECASE)

# =============================
# COMENTAR AMPSCRIPT (com HOIST de {% let ... %} existentes)
# =============================
def comment_ampscript_with_hoist(html_text: str):
    """
    - Hoista qualquer `{% let ... %}` que esteja dentro de blocos AMPScript.
    - Comenta o restante do AMPScript (%%=...=%%, %%[ ... ]%%, <script runat=server>...</script>).
    - Retorna (html_modificado, lista_de_blocos_comentados_com_linha)
    """
    block_re = re.compile(
        r'(%%=.+?=%%|%%\[[\s\S]*?\]%%|<script[^>]*\brunat=[\'"]server[\'"][^>]*>[\s\S]*?<\/script>)',
        re.IGNORECASE | re.DOTALL
    )
    let_re = re.compile(r'{%\s*let\s+[^%]+%}', re.IGNORECASE | re.DOTALL)

    commented = []

    # Pré-cálculo de linhas p/ localizar nº de linha
    lines = html_text.splitlines()
    offsets = [0]
    for line in lines:
        offsets.append(offsets[-1] + len(line) + 1)  # + '\n'

    def find_line(pos):
        for i, end in enumerate(offsets):
            if pos < end:
                return max(1, i)
        return len(lines)

    out = []
    last_end = 0

    for m in block_re.finditer(html_text):
        start, end = m.span()
        full = m.group(1)
        line_no = find_line(start)

        # 1) Hoist de LETs existentes
        lets = let_re.findall(full)
        cleaned = let_re.sub("", full)

        # 2) Anexar antes do comentário
        out.append(html_text[last_end:start])
        for l in lets:
            out.append(l)

        # 3) Comentar o restante do bloco AMPScript (se sobrou algo)
        if cleaned.strip():
            commented.append((line_no, cleaned.strip()))
            out.append("<!-- " + cleaned + " -->")

        last_end = end

    out.append(html_text[last_end:])
    return "".join(out), commented

# =============================
# LOG
# =============================
def write_log(commented_blocks, total_found, total_replaced, output_path, warnings, covered_vars):
    """
    Gera um .txt com:
      - Resumo de execução
      - Warnings de variáveis não mapeadas (apenas se realmente não cobertas)
      - Lista de blocos AMPScript comentados (linha + trecho)
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path(f"output_log_{timestamp}.txt")

    lines = []
    lines.append(f"=== EXECUÇÃO EM {datetime.now().isoformat()} ===")
    lines.append(f"Arquivo HTML (entrada): {html_input}")
    lines.append(f"Arquivo HTML (saída)  : {output_path}")
    lines.append(f"Planilha              : {excel_path}")
    lines.append("")
    unresolved = sorted({w.lower() for w in warnings} - set(covered_vars))
    if unresolved:
        lines.append("⚠️ Variáveis não mapeadas detectadas (use profile./context. no AJO ou inclua na planilha):")
        for w in unresolved:
            lines.append(f"   - @{w}")
        lines.append("")
    lines.append(f"Trechos SFMC encontrados (pré-substituição): {total_found}")
    lines.append(f"Substituições aplicadas (tabela de-para): {total_replaced}")
    lines.append(f"Blocos AMPScript comentados: {len(commented_blocks)}")
    lines.append("")
    lines.append("=== BLOCOS COMENTADOS ===")
    for line, snippet in commented_blocks:
        snippet_clean = snippet.replace("\n", "\\n")
        lines.append(f"Linha {line:>5}: {snippet_clean[:400]}{'...' if len(snippet_clean) > 400 else ''}")

    log_path.write_text("\n".join(lines), encoding="utf-8")
    return log_path

# =============================
# PIPELINE
# =============================
# 0) Ler planilha
df = pd.read_excel(excel_path, usecols="A:B")
sfmc_col, ajo_col = pick_cols(df)
df = df[[sfmc_col, ajo_col]].rename(columns={sfmc_col: "SFMC", ajo_col: "AJO"})
df = df.dropna(subset=["SFMC"])
df["SFMC"] = df["SFMC"].astype(str)
df["AJO"]  = df["AJO"].fillna("").astype(str)

# 1) Ler HTML
html = Path(html_input).read_text(encoding="utf-8").replace("\r\n", "\n")

# 2) Variáveis conhecidas via planilha
var_expr, covered_vars = extract_var_expressions(df)
warnings = []

# 3) Traduzir blocos condicionais AMPScript (IF/ELSE/ENDIF)
html_after_if = convert_ampscript_if_blocks(html, warnings, var_expr, covered_vars)

# 4) Substituições simples via tabela (de-para literal)
rows = sorted(df.to_dict("records"), key=lambda r: len(str(r["SFMC"]).strip()), reverse=True)
total_found = 0
total_replaced = 0
work = html_after_if
for r in rows:
    sfmc = str(r["SFMC"]).strip()
    ajo  = str(r["AJO"]).strip()
    if not sfmc:
        continue
    pattern = build_flex_regex(sfmc)
    matches = list(pattern.finditer(work))
    if not matches:
        continue
    total_found += len(matches)
    if ajo:
        work, n = pattern.subn(ajo, work)
        total_replaced += n

# 5) Replace global de prints AMPScript → Liquid (apenas mapeados)
work = replace_all_prints(work, var_expr)

# 6) Comentar AMPScript remanescente (e hoistar LETs existentes)
final_html, commented_blocks = comment_ampscript_with_hoist(work)

# 7) Salvar HTML de saída
Path(html_output).write_text(final_html, encoding="utf-8")

# 8) LOG
log_path = write_log(commented_blocks, total_found, total_replaced, html_output, warnings, {v.lower() for v in covered_vars})

# 9) Console
print("==== SFMC → AJO (relatório) ====")
print(f"Entrada : {html_input}")
print(f"Saída   : {html_output}")
print(f"Planilha: {excel_path}")
print(f"SFMC matches: {total_found}")
print(f"Substituições: {total_replaced}")
print(f"AMPScript comentado: {len(commented_blocks)}")
print(f"Log: {log_path.resolve()}")
unresolved = sorted({w.lower() for w in warnings} - {v.lower() for v in covered_vars})
if unresolved:
    print("\n⚠️  Variáveis não mapeadas detectadas:")
    for w in unresolved:
        print(f"   - @{w}")
