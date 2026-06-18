#!/usr/bin/env python3
"""
Andlz Dashboard — Gerador automático
Lê a planilha Demonstrações_2026_final.xlsx e injeta os dados no template,
produzindo dashboard_andlz.html pronto para publicação no GitHub Pages.

Uso:
    python update_dashboard.py

Requisitos:
    pip install openpyxl
"""

import json
import math
import os
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl --quiet")
    import openpyxl

# ── CONFIGURAÇÃO ──────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
TEMPLATE     = SCRIPT_DIR / "dashboard_template.html"
OUTPUT       = SCRIPT_DIR / "dashboard_andlz.html"

# Aceita o nome da planilha como argumento ou busca automaticamente
if len(sys.argv) > 1:
    PLANILHA = Path(sys.argv[1])
else:
    # Busca qualquer .xlsx na pasta do script
    xlsx_files = sorted(SCRIPT_DIR.glob("*.xlsx"))
    if not xlsx_files:
        print("ERRO: Nenhum arquivo .xlsx encontrado na pasta.")
        print("Coloque a planilha na mesma pasta do script ou passe o caminho como argumento.")
        sys.exit(1)
    PLANILHA = xlsx_files[0]
    print(f"Planilha encontrada: {PLANILHA.name}")

MONTHS = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']

# ── HELPERS ───────────────────────────────────────────────────
def load_wb(path):
    return openpyxl.load_workbook(path, data_only=True)

def cell_val(ws, row, col):
    """Retorna valor numérico de uma célula, 0 se vazia ou não-numérica."""
    v = ws.cell(row=row, column=col).value
    if v is None:
        return 0
    # Trata strings como "0,00" ou "1.234,56"
    if isinstance(v, str):
        v = v.strip().replace('.', '').replace(',', '.')
        try:
            return float(v)
        except ValueError:
            return 0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0

def row_values(ws, row, start_col, n_cols=12):
    """Retorna lista de n_cols valores numéricos a partir de start_col."""
    return [cell_val(ws, row, start_col + i) for i in range(n_cols)]

def find_row(ws, label, col=1, max_rows=200):
    """Busca a linha onde col==label (case-insensitive, strip)."""
    label_clean = label.strip().lower()
    for r in range(1, max_rows + 1):
        v = ws.cell(row=r, column=col).value
        if v and str(v).strip().lower() == label_clean:
            return r
    return None

def find_row_startswith(ws, prefix, col=1, max_rows=200):
    """Busca linha onde a célula começa com prefix."""
    prefix_clean = prefix.strip().lower()
    for r in range(1, max_rows + 1):
        v = ws.cell(row=r, column=col).value
        if v and str(v).strip().lower().startswith(prefix_clean):
            return r
    return None

def arr12(vals, n_realized):
    """Garante array de 12 posições: zeros após nRealized."""
    result = list(vals[:n_realized])
    result += [0] * (12 - len(result))
    return result

def arr12_full(vals):
    """Array de 12 posições sem truncar (para orçamento)."""
    result = list(vals[:12])
    result += [0] * (12 - len(result))
    return result

def safe(v, default=0):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return default
    return v

# ── LEITURA DA PLANILHA ───────────────────────────────────────
print(f"Lendo {PLANILHA.name}...")
wb = load_wb(PLANILHA)
sheets = wb.sheetnames
print(f"Abas encontradas: {', '.join(sheets)}")

# ── ABA: Consolidated ─────────────────────────────────────────
def read_consolidated():
    ws = wb['Consolidated']

    # Layout: rótulos na col D (4), dados Jan–Dez nas colunas E–P (5–16)
    LABEL_COL = 4
    DATA_COL  = 5

    def find_row_in_ws(label):
        label_clean = label.strip().lower()
        for r in range(1, 300):
            v = ws.cell(row=r, column=LABEL_COL).value
            if v and str(v).strip().lower().startswith(label_clean[:15]):
                return r
        return None

    def get_row(label):
        r = find_row_in_ws(label)
        if r is None:
            return [0] * 12
        return row_values(ws, r, DATA_COL, 12)

    rec_bruta   = get_row('01 Receitas Operacion')
    n_realized  = sum(1 for v in rec_bruta if v != 0)
    print(f"Meses realizados detectados: {n_realized} ({', '.join(MONTHS[:n_realized])})")

    # AuM e ROA ficam acima do cabeçalho (linhas 4 e 5)
    r_aum = None
    r_roa = None
    for r in range(1, 10):
        v = ws.cell(row=r, column=LABEL_COL).value
        if v:
            sv = str(v).strip().lower()
            if sv == 'aum':
                r_aum = r
            elif 'roa' in sv:
                r_roa = r

    real_aum_row = row_values(ws, r_aum, DATA_COL, 12) if r_aum else [0]*12
    real_aum = [int(v) for v in real_aum_row[:n_realized] if v != 0]

    real_roa_row = row_values(ws, r_roa, DATA_COL, 12) if r_roa else [0]*12
    real_roa = [v for v in real_roa_row[:n_realized]]

    data = {
        'recBruta':    arr12(rec_bruta, n_realized),
        'consultoria': arr12(get_row('Consultoria'), n_realized),
        'planejamento':arr12(get_row('Planejamento'), n_realized),
        'tesouroMT':   arr12(get_row('Tesouro Mato Grosso'), n_realized),
        'rpps':        arr12(get_row('RPPS'), n_realized),
        'consulta':    arr12(get_row('Consulta'), n_realized),
        'valuation':   arr12(get_row('Valuation'), n_realized),
        'cambio':      arr12(get_row('Câmbio/Bancários'), n_realized),
        'consorcio':   arr12(get_row('Consórcio'), n_realized),
        'previdencia': arr12(get_row('Previdência'), n_realized),
        'segVida':     arr12(get_row('Seguro de vida'), n_realized),
        'outrosRamos': arr12(get_row('Outros Ramos'), n_realized),
        'deducoes':    arr12(get_row('02 Deduções'), n_realized),
        'imp_vendas':  arr12(get_row('02.1 Impostos Sobre Vendas'), n_realized),
        'com_consult': arr12(get_row('02.2 Comissões Consultores'), n_realized),
        'com_mesa':    arr12(get_row('02.3 Comissões Mesa'), n_realized),
        'descontos':   arr12(get_row('02.4 Descontos'), n_realized),
        'devolucoes':  arr12(get_row('02.5 Devoluções'), n_realized),
        'com_parc':    arr12(get_row('02.6 Comissões Parceiros'), n_realized),
        'recLiq':      arr12(get_row('02T Receita Líquida'), n_realized),
        'custos':      arr12(get_row('03 Custos Operacion'), n_realized),
        'sal_enc':     arr12(get_row('03.1 Salários'), n_realized),
        'sistemas_c':  arr12(get_row('03.2 Sistemas'), n_realized),
        'custo_MT':    arr12(get_row('03.3 Custo Tesouro'), n_realized),
        'custo_rpps':  arr12(get_row('03.4 Custo RPPS'), n_realized),
        'regulatorias':arr12(get_row('03.5 Regulatórias'), n_realized),
        'lucroBruto':  arr12(get_row('03T Lucro Bruto'), n_realized),
        'despOp':      arr12(get_row('04 Despesas Operacion'), n_realized),
        'imovel':      arr12(get_row('04.1 Imóvel'), n_realized),
        'servicos':    arr12(get_row('04.2 Serviços'), n_realized),
        'comerciais':  arr12(get_row('04.3 Comerciais'), n_realized),
        'admin':       arr12(get_row('04.4 Admin'), n_realized),
        'colabor':     arr12(get_row('04.5 Colaboradores'), n_realized),
        'diretoria':   arr12(get_row('04.6 Diretoria'), n_realized),
        'lucroOp':     arr12(get_row('04T Lucro / Prejuízo Operac'), n_realized),
        'financeiras': arr12(get_row('05 Receitas e Despesas Financ'), n_realized),
        'outras_nd':   arr12(get_row('06 Outras'), n_realized),
        'lucroLiq':    arr12(get_row('06T Lucro / Prejuízo Líq'), n_realized),
        'inv_emp':     arr12(get_row('07 Despesas com Invest'), n_realized),
        'inv_imob':    arr12(get_row('07.1 Investimentos em Imob'), n_realized),
        'emprestimos': arr12(get_row('07.2 Empréstimos'), n_realized),
        'lucroFinal':  arr12(get_row('07T Lucro / Prejuízo Final'), n_realized),
        'realCaixa':   [v for v in get_row('Caixa') if v != 0][:n_realized],
        'realAuM':     real_aum,
        'realROA':     real_roa,
    }

    # EBITDA = lucroOp + |imp_vendas|
    data['ebitda'] = [
        data['lucroOp'][i] + abs(data['imp_vendas'][i])
        for i in range(12)
    ]

    # Margens (só meses realizados)
    data['margEBITDA'] = [
        round(data['ebitda'][i] / data['recBruta'][i] * 100, 4)
        if data['recBruta'][i] != 0 else 0
        for i in range(n_realized)
    ]
    data['margBruta'] = [
        round(data['lucroBruto'][i] / data['recBruta'][i] * 100, 4)
        if data['recBruta'][i] != 0 else 0
        for i in range(n_realized)
    ]
    data['margLiq'] = [
        round(data['lucroLiq'][i] / data['recBruta'][i] * 100, 4)
        if data['recBruta'][i] != 0 else 0
        for i in range(n_realized)
    ]
    data['margFinal'] = [
        round(data['lucroFinal'][i] / data['recBruta'][i] * 100, 4)
        if data['recBruta'][i] != 0 else 0
        for i in range(n_realized)
    ]

    # Caixa projetado
    r_caixa_proj = find_row_in_ws('Caixa projetado')
    data['projCaixa'] = arr12_full(row_values(ws, r_caixa_proj, DATA_COL, 12) if r_caixa_proj else [0]*12)

    return data, n_realized

# ── ABA: Orçamento 2026 ───────────────────────────────────────
def read_orcamento(n_realized):
    ws = wb['Orçamento 2026']
    LABEL_COL = 2
    DATA_COL  = 3  # Jan na col C

    def find_row_orc(label):
        lc = label.strip().lower()
        for r in range(1, 200):
            v = ws.cell(row=r, column=LABEL_COL).value
            if v and str(v).strip().lower().startswith(lc[:15]):
                return r
        return None

    def get_row(label):
        r = find_row_orc(label)
        return row_values(ws, r, DATA_COL, 12) if r else [0]*12

    orc = {
        'orcRecBruta':  arr12_full(get_row('01 Receitas Ope')),
        'orcDeducoes':  arr12_full(get_row('02 Deduções da')),
        'orcCustos':    arr12_full(get_row('03 Custos Oper')),
        'orcDespOp':    arr12_full(get_row('04 Despesas Oper')),
        'orcLucroOp':   arr12_full(get_row('04T Lucro / Prej')),
        'orcFinal':     arr12_full(get_row('07T Lucro / Prej')),
    }

    marg_row = find_row_orc('Margem EBITDA')
    if marg_row:
        raw_marg = row_values(ws, marg_row, DATA_COL, 12)
        orc['orcMargEBITDA'] = [round(v*100,4) if abs(v) < 2 else round(v,4) for v in raw_marg]
    else:
        orc['orcMargEBITDA'] = [0]*12

    # Meta AuM (célula C5)
    meta_cell = ws.cell(row=5, column=3).value
    orc['metaAuM'] = float(meta_cell) if meta_cell else 280000000

    orc['projFinal'] = ['null']*n_realized + orc['orcFinal'][n_realized:]
    return orc

# ── ABA: AuM ─────────────────────────────────────────────────
def read_aum(n_realized):
    ws = wb['AuM']
    # Col A=ano, Col B=mês, Col C=valor (R$ ou milhões)
    series_by_year = {}
    for r in range(1, 400):
        ano = ws.cell(row=r, column=1).value
        mes = ws.cell(row=r, column=2).value
        val = ws.cell(row=r, column=3).value
        if not ano or not mes or not val:
            continue
        try:
            ano_int = int(ano)
            val_f   = float(val)
        except (TypeError, ValueError):
            continue
        if ano_int < 2018 or ano_int > 2030:
            continue
        mes_map = {
            'janeiro':0,'fevereiro':1,'março':2,'abril':3,'maio':4,'junho':5,
            'julho':6,'agosto':7,'setembro':8,'outubro':9,'novembro':10,'dezembro':11
        }
        mes_idx = mes_map.get(str(mes).lower().strip())
        if mes_idx is not None:
            if ano_int not in series_by_year:
                series_by_year[ano_int] = {}
            series_by_year[ano_int][mes_idx] = val_f

    # Determina se valores estão em R$ ou milhões
    sample_vals = [v for yr in series_by_year.values() for v in yr.values()]
    in_reais = any(v > 1e6 for v in sample_vals[:5]) if sample_vals else False

    aum_raw_series, hist_labels, hist_vals = [], [], []
    for year in sorted(series_by_year.keys()):
        months_data = series_by_year[year]
        n_m = max(months_data.keys()) + 1 if months_data else 0
        m_list = [
            (months_data[i] / 1e6 if in_reais else months_data[i]) if i in months_data else 0
            for i in range(n_m)
        ]
        if m_list:
            aum_raw_series.append({'y': year, 'm': m_list})
        if year < 2026 and 11 in months_data:
            hist_labels.append(str(year))
            hist_vals.append(months_data[11] if not in_reais else int(months_data[11]))

    real_aum_2026 = []
    if 2026 in series_by_year:
        d = series_by_year[2026]
        real_aum_2026 = [
            int(d[i]) if not in_reais else int(d[i])
            for i in range(n_realized) if i in d
        ]

    # projAuM do Orçamento 2026
    proj_aum_full = ['null']*n_realized + ['null']*(12-n_realized)
    try:
        ws_orc = wb['Orçamento 2026']
        for r in range(1, 200):
            v = ws_orc.cell(row=r, column=2).value
            if v and 'aum projetado' in str(v).lower():
                for i in range(n_realized, 12):
                    pv = ws_orc.cell(row=r, column=3+i).value
                    proj_aum_full[i] = int(float(pv)) if pv else 'null'
                break
    except Exception:
        pass

    return {
        'realAuM':      real_aum_2026,
        'projAuM':      proj_aum_full,
        'aumRawSeries': aum_raw_series,
        'aumHistLabels':hist_labels,
        'aumHistVals':  hist_vals,
    }

# ── ABA: Projeção (linha 62 — Caixa contínuo) ─────────────────
def read_projecao():
    ws = wb['Projeção']
    for r in range(55, 75):
        v = ws.cell(row=r, column=1).value
        if not v or 'caixa' not in str(v).lower():
            continue
        vals = row_values(ws, r, 2, 12)
        if sum(1 for x in vals if x != 0) >= 5:
            print(f"  Linha Caixa Projeção encontrada na linha {r}")
            return arr12_full(vals)
    # Try col 4 (D) as label col
    for r in range(55, 75):
        v = ws.cell(row=r, column=4).value
        if not v or 'caixa' not in str(v).lower():
            continue
        vals = row_values(ws, r, 5, 12)
        if sum(1 for x in vals if x != 0) >= 5:
            print(f"  Linha Caixa Projeção encontrada na linha {r} (col D)")
            return arr12_full(vals)
    return [0]*12

# ── ABA: CNS ──────────────────────────────────────────────────
def read_cns(n_realized):
    ws = wb['CNS']
    LABEL_COL = 1
    DATA_COL  = 2
    def get(label):
        r = find_row(ws, label, LABEL_COL) or find_row_startswith(ws, label, LABEL_COL)
        return arr12(row_values(ws, r, DATA_COL, 12) if r else [0]*12, n_realized)
    return {
        'cnsRB':      get('01 Receitas Operacionais'),
        'cnsRL':      get('02T Receita Líquida de Vendas'),
        'cnsFinal':   get('07T Lucro / Prejuízo Final'),
        'cnsConsult': get('Consultoria'),
        'cnsPlan':    get('Planejamento'),
        'cnsMT':      get('Tesouro Mato Grosso'),
        'cnsRPPS':    get('RPPS'),
        'cnsConsulta':get('Consulta'),
    }

# ── ABA: CRR ──────────────────────────────────────────────────
def read_crr(n_realized):
    ws = wb['CRR']
    LABEL_COL = 1
    DATA_COL  = 2
    def get(label):
        r = find_row(ws, label, LABEL_COL) or find_row_startswith(ws, label, LABEL_COL)
        return arr12(row_values(ws, r, DATA_COL, 12) if r else [0]*12, n_realized)
    return {
        'crrRB':       get('01 Receitas Operacionais'),
        'crrRL':       get('02T Receita Líquida de Vendas'),
        'crrFinal':    get('07T Lucro / Prejuízo Final'),
        'crrCambio':   get('Câmbio/Bancários'),
        'crrConsorcio':get('Consórcio'),
        'crrPrevid':   get('Previdência'),
        'crrSegVida':  get('Seguro de Vida'),
    }

# ── ABA: Captação ─────────────────────────────────────────────
# ── ABA: Receita (consultores) ────────────────────────────────
def read_receita(n_realized):
    ws = wb['Receita']
    CONS_COL = 9   # col I = consultor name
    DATA_COL = 12  # col L = Jan
    consultores_dict = {}
    for r in range(15, 500):
        cons = ws.cell(row=r, column=CONS_COL).value
        if not cons or not isinstance(cons, str) or len(cons.strip()) < 2:
            continue
        cons = cons.strip()
        vals = [cell_val(ws, r, DATA_COL + i) for i in range(12)]
        if not any(vals[:n_realized]):
            continue
        if cons not in consultores_dict:
            consultores_dict[cons] = [0.0] * 12
        for i in range(12):
            consultores_dict[cons][i] += vals[i]
    result = [{'n': n, 'v': arr12(v, n_realized)} for n, v in consultores_dict.items()]
    result.sort(key=lambda x: sum(x['v']), reverse=True)
    return result

# ── ABA: Gastos ───────────────────────────────────────────────
def read_gastos(n_realized):
    # Pull aggregated cost lines from Consolidated DRE (already summed)
    ws = wb['Consolidated']
    LABEL_COL = 4
    DATA_COL  = 5
    def get_row_val(label):
        lc = label.strip().lower()
        for r in range(1, 200):
            v = ws.cell(row=r, column=LABEL_COL).value
            if v and str(v).strip().lower().startswith(lc[:15]):
                return [abs(cell_val(ws, r, DATA_COL+i)) for i in range(12)]
        return [0]*12
    cost_map = [
        ('Salários e Encargos', '03.1 Salários',      False),
        ('Mato Grosso',         '03.3 Custo Tesouro',  False),
        ('Sistemas',            '03.2 Sistemas',       False),
        ('Desp. Operacionais',  '04 Despesas Oper',    False),
        ('Financeiras',         '05 Receitas e Desp',  False),
        ('Regulatórias',        '03.5 Regulatórias',   False),
        ('Comissões',           '02.2 Comissões Cons', True),
        ('Impostos s/ Vendas',  '02.1 Impostos Sobre', True),
        ('PLR/Bônus',           None,                  True),
    ]
    gastos = []
    for nome, label, excl in cost_map:
        vals = get_row_val(label) if label else [0]*12
        if sum(vals[:n_realized]) > 0:
            gastos.append({'n': nome, 'v': arr12(vals, n_realized), 'excl': excl})
    return gastos

# ── ABA: Sistemas ─────────────────────────────────────────────
def read_sistemas():
    ws_name = 'Sistemas ' if 'Sistemas ' in sheets else 'Sistemas'
    if ws_name not in sheets:
        return [], 255
    ws = wb[ws_name]
    # Col 5=categoria, Col 6=nome, Col 7=valor
    sistemas = []
    current_cat = 'Geral'
    for r in range(1, 100):
        cat_cell = ws.cell(row=r, column=5).value
        nome     = ws.cell(row=r, column=6).value
        val      = ws.cell(row=r, column=7).value
        if cat_cell and isinstance(cat_cell, str) and not nome:
            current_cat = cat_cell.strip()
            continue
        if nome and isinstance(nome, str) and val is not None:
            nome = nome.strip()
            if nome.lower() in ['total','subtotal'] or nome.isdigit():
                continue
            try:
                v = float(val)
                if v > 0:
                    sistemas.append({'n': nome, 'cat': current_cat, 'v': v})
            except (TypeError, ValueError):
                pass
    return sistemas, 255

# ── ABA: H. Receita ───────────────────────────────────────────
def read_h_receita():
    ws = wb['H. Receita']
    # Col 4=data, Col 5=Receita Op, Col 6=Lucro Op
    from datetime import datetime
    qmap = {1:'1T',2:'1T',3:'1T',4:'2T',5:'2T',6:'2T',7:'3T',8:'3T',9:'3T',10:'4T',11:'4T',12:'4T'}
    qacc, racc = {}, {}
    for r in range(4, 500):
        dt_val = ws.cell(row=r, column=4).value
        rec    = ws.cell(row=r, column=5).value
        lucro  = ws.cell(row=r, column=6).value
        if not dt_val or lucro is None:
            continue
        try:
            dt = dt_val if isinstance(dt_val, datetime) else datetime.strptime(str(dt_val)[:10], '%Y-%m-%d')
            key = qmap[dt.month] + str(dt.year)[-2:]
            qacc[key] = qacc.get(key, 0) + float(lucro or 0)
            racc[key] = racc.get(key, 0) + float(rec or 0)
        except (ValueError, TypeError):
            continue
    hist_labels = sorted(qacc.keys())
    hist_vals   = [round(qacc[k], 2) for k in hist_labels]
    return hist_labels, hist_vals, [], []

def read_captacao(n_realized):
    if 'Captação' not in sheets:
        return None
    ws = wb['Captação']
    LABEL_COL = 6   # col F
    DATA_COL  = 7   # col G = Janeiro

    def row_arr(row_num):
        return arr12(row_values(ws, row_num, DATA_COL, 12), n_realized)

    return {
        'numClientes':   row_arr(5),
        'captLiquida':   row_arr(6),
        'ticketMedio':   row_arr(7),
        'saldoClientes': row_arr(8),
        'captBruta':     row_arr(9),
        'novosContratos':row_arr(10),
        'renovacoes':    row_arr(11),
        'entradaClientes':row_arr(12),
        'churnTotal':    row_arr(13),
        'retiradas':     row_arr(14),
        'cancelamento':  row_arr(15),
        'saidaClientes': row_arr(16),
    }

# ── ABA: DCF Andlz ────────────────────────────────────────────
def read_dcf():
    if 'DCF Andlz' not in sheets:
        return None
    ws = wb['DCF Andlz']
    anos    = ['2027','2028','2029','2030','2031','Perpetuidade']
    fcf_rows = []
    val_rows = []

    # Premissas B4-B8
    dcf = {
        'ebitdaBase':   cell_val(ws, 4, 2),
        'wacc':         cell_val(ws, 5, 2),
        'txCrescAnual': cell_val(ws, 6, 2),
        'txCrescPerp':  cell_val(ws, 7, 2),
        'pctFCF':       cell_val(ws, 8, 2),
        'fcfRows':      [],
        'valRows':      [],
    }

    # Tabela FCF linhas 10-15
    for r in range(10, 16):
        item = ws.cell(row=r, column=1).value
        if not item:
            continue
        row_data = {'Item': str(item).strip()}
        for i, ano in enumerate(anos):
            v = ws.cell(row=r, column=2 + i).value
            row_data[ano] = str(v) if v is not None else '–'
        dcf['fcfRows'].append(row_data)

    # Tabela Valuation linhas 18-24
    for r in range(18, 25):
        item = ws.cell(row=r, column=1).value
        val  = ws.cell(row=r, column=2).value
        if not item:
            continue
        dcf['valRows'].append({
            'Item': str(item).strip(),
            'Valor': str(val) if val is not None else '–'
        })

    return dcf

# ── EXECUÇÃO ──────────────────────────────────────────────────
print("\nExtraindo dados...")

consolidated, n_realized = read_consolidated()
orcamento    = read_orcamento(n_realized)
aum_data     = read_aum(n_realized)
proj_caixa62 = read_projecao()
cns_data     = read_cns(n_realized)
crr_data     = read_crr(n_realized)
consultores  = read_receita(n_realized)
gastos       = read_gastos(n_realized)
sistemas, num_clientes = read_sistemas()
hist_labels, hist_vals, _, _ = read_h_receita()
captacao     = read_captacao(n_realized)
dcf          = read_dcf()

print(f"  Consultores encontrados: {len(consultores)}")
print(f"  Sistemas encontrados: {len(sistemas)}")
print(f"  Captação: {'✓' if captacao else '–'}")
print(f"  DCF: {'✓' if dcf else '–'}")

# ── MONTA BLOCO RAW ───────────────────────────────────────────
def js_arr(lst):
    """Converte lista Python para array JS, preservando null."""
    items = []
    for v in lst:
        if v == 'null' or v is None:
            items.append('null')
        elif isinstance(v, float) and math.isnan(v):
            items.append('0')
        else:
            items.append(str(round(float(v), 6)) if isinstance(v, float) else str(v))
    return '[' + ','.join(items) + ']'

def js_val(v):
    if v is None or v == 'null':
        return 'null'
    if isinstance(v, bool):
        return 'true' if v else 'false'
    if isinstance(v, (int, float)):
        return str(round(v, 6))
    return json.dumps(v, ensure_ascii=False)

# Consultores JS
def cons_js(consultores_list):
    lines = []
    for c in consultores_list:
        lines.append(f"    {{n:{json.dumps(c['n'])},v:{js_arr(c['v'])}}},")
    return '[\n' + '\n'.join(lines) + '\n  ]'

# Gastos JS
def gastos_js(gastos_list):
    lines = []
    for g in gastos_list:
        excl_str = 'true' if g['excl'] else 'false'
        lines.append(f"    {{n:{json.dumps(g['n'])},v:{js_arr(g['v'])},excl:{excl_str}}},")
    return '[\n' + '\n'.join(lines) + '\n  ]'

# Sistemas JS
def sistemas_js(sistemas_list):
    lines = []
    for s in sistemas_list:
        lines.append(f"    {{n:{json.dumps(s['n'])},cat:{json.dumps(s['cat'])},v:{s['v']}}},")
    return '[\n' + '\n'.join(lines) + '\n  ]'

# aumRawSeries JS
def aum_series_js(series):
    lines = []
    for yr in series:
        m_str = ','.join(str(round(v,2)) for v in yr['m'])
        lines.append(f"    {{y:{yr['y']},m:[{m_str}]}},")
    return '[\n' + '\n'.join(lines) + '\n  ]'

# captacao JS
def captacao_js(c):
    if not c:
        # Fallback com arrays zeros
        return '''{
    numClientes:[0,0,0,0,0,0,0,0,0,0,0,0],
    captLiquida:[0,0,0,0,0,0,0,0,0,0,0,0],
    ticketMedio:[0,0,0,0,0,0,0,0,0,0,0,0],
    saldoClientes:[0,0,0,0,0,0,0,0,0,0,0,0],
    captBruta:[0,0,0,0,0,0,0,0,0,0,0,0],
    novosContratos:[0,0,0,0,0,0,0,0,0,0,0,0],
    renovacoes:[0,0,0,0,0,0,0,0,0,0,0,0],
    entradaClientes:[0,0,0,0,0,0,0,0,0,0,0,0],
    churnTotal:[0,0,0,0,0,0,0,0,0,0,0,0],
    retiradas:[0,0,0,0,0,0,0,0,0,0,0,0],
    cancelamento:[0,0,0,0,0,0,0,0,0,0,0,0],
    saidaClientes:[0,0,0,0,0,0,0,0,0,0,0,0],
  }'''
    fields = ['numClientes','captLiquida','ticketMedio','saldoClientes',
              'captBruta','novosContratos','renovacoes','entradaClientes',
              'churnTotal','retiradas','cancelamento','saidaClientes']
    lines = [f"    {f}:{js_arr(c[f])}," for f in fields]
    return '{\n' + '\n'.join(lines) + '\n  }'

# DCF JS
def dcf_js(d):
    if not d:
        return '''{
    ebitdaBase:null,wacc:null,txCrescAnual:null,txCrescPerp:null,pctFCF:null,
    fcfRows:[],valRows:[],
  }'''
    fcf_str  = json.dumps(d['fcfRows'],  ensure_ascii=False)
    val_str  = json.dumps(d['valRows'],  ensure_ascii=False)
    return f'''{{
    ebitdaBase:{js_val(d['ebitdaBase'])},
    wacc:{js_val(d['wacc'])},
    txCrescAnual:{js_val(d['txCrescAnual'])},
    txCrescPerp:{js_val(d['txCrescPerp'])},
    pctFCF:{js_val(d['pctFCF'])},
    fcfRows:{fcf_str},
    valRows:{val_str},
  }}'''

# Histórico trimestral
hist_labels_js = json.dumps(hist_labels, ensure_ascii=False)
hist_vals_js   = js_arr(hist_vals)

# projFinal como array JS (mescla null e valores)
proj_final_js = '[' + ','.join(
    'null' if v == 'null' else str(round(float(v),6))
    for v in orcamento['projFinal']
) + ']'

# projAuM como array JS
proj_aum_js = '[' + ','.join(
    'null' if v == 'null' else str(v)
    for v in aum_data['projAuM']
) + ']'

raw_block = f"""const RAW = {{
  recBruta:{js_arr(consolidated['recBruta'])},
  consultoria:{js_arr(consolidated['consultoria'])},
  planejamento:{js_arr(consolidated['planejamento'])},
  tesouroMT:{js_arr(consolidated['tesouroMT'])},
  rpps:{js_arr(consolidated['rpps'])},
  consulta:{js_arr(consolidated['consulta'])},
  valuation:{js_arr(consolidated['valuation'])},
  cambio:{js_arr(consolidated['cambio'])},
  consorcio:{js_arr(consolidated['consorcio'])},
  previdencia:{js_arr(consolidated['previdencia'])},
  segVida:{js_arr(consolidated['segVida'])},
  outrosRamos:{js_arr(consolidated['outrosRamos'])},
  deducoes:{js_arr(consolidated['deducoes'])},
  imp_vendas:{js_arr(consolidated['imp_vendas'])},
  com_consult:{js_arr(consolidated['com_consult'])},
  com_mesa:{js_arr(consolidated['com_mesa'])},
  descontos:{js_arr(consolidated['descontos'])},
  devolucoes:{js_arr(consolidated['devolucoes'])},
  com_parc:{js_arr(consolidated['com_parc'])},
  recLiq:{js_arr(consolidated['recLiq'])},
  custos:{js_arr(consolidated['custos'])},
  sal_enc:{js_arr(consolidated['sal_enc'])},
  sistemas_c:{js_arr(consolidated['sistemas_c'])},
  custo_MT:{js_arr(consolidated['custo_MT'])},
  custo_rpps:{js_arr(consolidated['custo_rpps'])},
  regulatorias:{js_arr(consolidated['regulatorias'])},
  lucroBruto:{js_arr(consolidated['lucroBruto'])},
  despOp:{js_arr(consolidated['despOp'])},
  imovel:{js_arr(consolidated['imovel'])},
  servicos:{js_arr(consolidated['servicos'])},
  comerciais:{js_arr(consolidated['comerciais'])},
  admin:{js_arr(consolidated['admin'])},
  colabor:{js_arr(consolidated['colabor'])},
  diretoria:{js_arr(consolidated['diretoria'])},
  lucroOp:{js_arr(consolidated['lucroOp'])},
  financeiras:{js_arr(consolidated['financeiras'])},
  outras_nd:{js_arr(consolidated['outras_nd'])},
  lucroLiq:{js_arr(consolidated['lucroLiq'])},
  inv_emp:{js_arr(consolidated['inv_emp'])},
  inv_imob:{js_arr(consolidated['inv_imob'])},
  emprestimos:{js_arr(consolidated['emprestimos'])},
  lucroFinal:{js_arr(consolidated['lucroFinal'])},
  ebitda:{js_arr(consolidated['ebitda'])},
  margEBITDA:{js_arr(consolidated['margEBITDA'])},
  margBruta:{js_arr(consolidated['margBruta'])},
  margLiq:{js_arr(consolidated['margLiq'])},
  margFinal:{js_arr(consolidated['margFinal'])},
  projFinal:{proj_final_js},
  projCaixa:{js_arr(orcamento.get('orcFinal', [0]*12))},
  projCaixaLinha62:{js_arr(proj_caixa62)},
  realAuM:{js_arr(aum_data['realAuM'])},
  projAuM:{proj_aum_js},
  metaAuM:{int(orcamento['metaAuM'])},
  realROA:{js_arr(consolidated['realROA'])},
  realCaixa:{js_arr(consolidated['realCaixa'])},
  orcRecBruta:{js_arr(orcamento['orcRecBruta'])},
  orcDeducoes:{js_arr(orcamento['orcDeducoes'])},
  orcCustos:{js_arr(orcamento['orcCustos'])},
  orcDespOp:{js_arr(orcamento['orcDespOp'])},
  orcLucroOp:{js_arr(orcamento['orcLucroOp'])},
  orcFinal:{js_arr(orcamento['orcFinal'])},
  orcMargEBITDA:{js_arr(orcamento['orcMargEBITDA'])},
  cnsRB:{js_arr(cns_data['cnsRB'])},
  cnsRL:{js_arr(cns_data['cnsRL'])},
  cnsFinal:{js_arr(cns_data['cnsFinal'])},
  cnsConsult:{js_arr(cns_data['cnsConsult'])},
  cnsPlan:{js_arr(cns_data['cnsPlan'])},
  cnsMT:{js_arr(cns_data['cnsMT'])},
  cnsRPPS:{js_arr(cns_data['cnsRPPS'])},
  cnsConsulta:{js_arr(cns_data['cnsConsulta'])},
  crrRB:{js_arr(crr_data['crrRB'])},
  crrRL:{js_arr(crr_data['crrRL'])},
  crrFinal:{js_arr(crr_data['crrFinal'])},
  crrCambio:{js_arr(crr_data['crrCambio'])},
  crrConsorcio:{js_arr(crr_data['crrConsorcio'])},
  crrPrevid:{js_arr(crr_data['crrPrevid'])},
  crrSegVida:{js_arr(crr_data['crrSegVida'])},
  consultores:{cons_js(consultores)},
  gastosAll:{gastos_js(gastos)},
  histLabels:{hist_labels_js},
  histVals:{hist_vals_js},
  histRecLabels:[],
  histRecVals:[],
  aumHistLabels:{json.dumps(aum_data['aumHistLabels'])},
  aumHistVals:{js_arr(aum_data['aumHistVals'])},
  aumRawSeries:{aum_series_js(aum_data['aumRawSeries'])},
  sistemas:{sistemas_js(sistemas)},
  numClientes:{num_clientes},
  captacao:{captacao_js(captacao)},
  dcf:{dcf_js(dcf)},
}};"""

# ── INJETA NO TEMPLATE ────────────────────────────────────────
print("\nLendo template...")
template_content = TEMPLATE.read_text(encoding='utf-8')

# Substitui marcador de dados
injection_marker = "// <<SKILL_DATA_INJECTION>>\nconst RAW = __INJECTED_DATA__;"
if injection_marker not in template_content:
    print("ERRO: Marcador '// <<SKILL_DATA_INJECTION>>' não encontrado no template.")
    print("Verifique se dashboard_template.html é o arquivo correto.")
    sys.exit(1)

output_content = template_content.replace(injection_marker, raw_block)

# Substitui pTo padrão
output_content = re.sub(
    r'let pFrom = 0, pTo = __DEFAULT_PTO__;.*\n',
    f'let pFrom = 0, pTo = {n_realized - 1}; // {MONTHS[n_realized-1]}/26\n',
    output_content
)

# Atualiza seletor "Até" para refletir meses realizados e selecionar o último
month_options = '\n'.join(
    f'          <option value="{i}"{" selected" if i == n_realized-1 else ""}>{MONTHS[i]}</option>'
    for i in range(n_realized)
)
# Substitui bloco de options do sel-to
output_content = re.sub(
    r'(<select class="period-select" id="sel-to">).*?(</select>)',
    f'<select class="period-select" id="sel-to">\n{month_options}\n        </select>',
    output_content,
    flags=re.DOTALL
)
# Substitui bloco de options do sel-from
month_options_from = '\n'.join(
    f'          <option value="{i}">{MONTHS[i]}</option>'
    for i in range(n_realized)
)
output_content = re.sub(
    r'(<select class="period-select" id="sel-from">).*?(</select>)',
    f'<select class="period-select" id="sel-from">\n{month_options_from}\n        </select>',
    output_content,
    flags=re.DOTALL
)

# ── SALVA ─────────────────────────────────────────────────────
OUTPUT.write_text(output_content, encoding='utf-8')
size_kb = OUTPUT.stat().st_size / 1024
print(f"\n✓ Dashboard gerado: {OUTPUT.name} ({size_kb:.0f} KB)")
print(f"  Período: Jan–{MONTHS[n_realized-1]}/2026 ({n_realized} meses realizados)")
print(f"  Meta AuM: R$ {int(orcamento['metaAuM']):,}")
print(f"  Captação: {'lida' if captacao else 'não encontrada (usando zeros)'}")
print(f"  DCF: {'lido' if dcf else 'não encontrado'}")
