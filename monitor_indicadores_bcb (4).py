# -*- coding: utf-8 -*-
"""
Monitor de Indicadores BCB (PTAX, Selic e IPCA) - painel com cards e gráfico.

A cada execução o código consulta a API SGS do Banco Central e atualiza:
  - Dólar PTAX (venda): últimos 30 dias corridos (gráfico de linha)
  - Selic: % a.d., % a.m. e % a.a. (séries oficiais do BCB, sem cálculo aproximado)
  - IPCA: variação mensal e acumulado em 12 meses (séries oficiais)

Saídas: painel HTML, banco SQLite (histórico) e planilha Excel.
"""

import os
import time
import sqlite3
from datetime import datetime, timedelta
from html import escape

import requests
import pandas as pd

# ----------------------------------------------------------------------
# Configuração
# ----------------------------------------------------------------------
BASE_DIR = '/content' if os.path.isdir('/content') else os.getcwd()
CAMINHO_DB = os.path.join(BASE_DIR, 'monitor_financeiro.db')
CAMINHO_HTML = os.path.join(BASE_DIR, 'painel_indicadores.html')
CAMINHO_PNG = os.path.join(BASE_DIR, 'painel_indicadores.png')
CAMINHO_EXCEL = os.path.join(BASE_DIR, 'relatorio_cotacoes.xlsx')

DIAS_GRAFICO = 30

# Séries do SGS/BCB (https://www3.bcb.gov.br/sgspub)
SERIES = {
    'Dolar_PTAX':      {'codigo': 1,     'unidade': 'R$',    'desc': 'Taxa de câmbio - Dólar (venda) - diário'},
    'Selic_diaria':    {'codigo': 11,    'unidade': '% a.d.', 'desc': 'Taxa de juros Selic - diária'},
    'Selic_mensal':    {'codigo': 4390,  'unidade': '% a.m.', 'desc': 'Selic acumulada no mês'},
    'Selic_anual':     {'codigo': 1178,  'unidade': '% a.a.', 'desc': 'Selic anualizada (base 252)'},
    'IPCA_mensal':     {'codigo': 433,   'unidade': '% a.m.', 'desc': 'IPCA - variação mensal'},
    'IPCA_12m':        {'codigo': 13522, 'unidade': '% 12m',  'desc': 'IPCA - acumulado em 12 meses'},
}

HEADERS = {'User-Agent': 'Mozilla/5.0'}
URL_BASE = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados'


# ----------------------------------------------------------------------
# Coleta
# ----------------------------------------------------------------------
def _get_json(url, params=None, tentativas=3):
    ultimo_erro = None
    for i in range(tentativas):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=20)
            r.raise_for_status()
            return r.json()
        except Exception as e:  # noqa
            ultimo_erro = e
            time.sleep(1.5 * (i + 1))
    raise RuntimeError(f'Falha ao consultar {url}: {ultimo_erro}')


def buscar_periodo(nome, dias):
    """Série dos últimos `dias` dias corridos até hoje."""
    hoje = datetime.now()
    params = {
        'formato': 'json',
        'dataInicial': (hoje - timedelta(days=dias)).strftime('%d/%m/%Y'),
        'dataFinal': hoje.strftime('%d/%m/%Y'),
    }
    dados = _get_json(URL_BASE.format(codigo=SERIES[nome]['codigo']), params)
    df = pd.DataFrame(dados)
    if df.empty:
        raise RuntimeError(f'A série {nome} não retornou dados no período.')
    df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y')
    df['valor'] = df['valor'].astype(float)
    return df.sort_values('data').reset_index(drop=True)


def buscar_ultimo(nome):
    """Último valor divulgado da série. Retorna (data, valor) ou (None, None)."""
    try:
        url = URL_BASE.format(codigo=SERIES[nome]['codigo']) + '/ultimos/1'
        item = _get_json(url, {'formato': 'json'})[-1]
        return datetime.strptime(item['data'], '%d/%m/%Y'), float(item['valor'])
    except Exception as e:  # noqa
        print(f'[AVISO] Não foi possível atualizar {nome}: {e}')
        return None, None


# ----------------------------------------------------------------------
# Persistência (SQLite)
# ----------------------------------------------------------------------
def salvar_historico(registros):
    """registros: lista de (indicador, data_referencia(datetime), valor)."""
    conn = sqlite3.connect(CAMINHO_DB)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS historico_indicadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            indicador TEXT NOT NULL,
            serie_sgs INTEGER NOT NULL,
            unidade TEXT NOT NULL,
            data_referencia TEXT NOT NULL,
            valor REAL NOT NULL,
            data_execucao_rpa TEXT NOT NULL,
            UNIQUE (indicador, data_referencia)
        )
    ''')
    agora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    for indicador, data_ref, valor in registros:
        s = SERIES[indicador]
        conn.execute('''
            INSERT INTO historico_indicadores
                (indicador, serie_sgs, unidade, data_referencia, valor, data_execucao_rpa)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT (indicador, data_referencia)
            DO UPDATE SET valor = excluded.valor, data_execucao_rpa = excluded.data_execucao_rpa
        ''', (indicador, s['codigo'], s['unidade'], data_ref.strftime('%Y-%m-%d'), valor, agora))
    conn.commit()
    conn.close()


# ----------------------------------------------------------------------
# Formatação
# ----------------------------------------------------------------------
def br(valor, casas=2):
    if valor is None:
        return 'N/D'
    return f'{valor:,.{casas}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def fmt_sinal(valor, casas=2):
    return ('+' if valor >= 0 else '') + br(valor, casas)


# ----------------------------------------------------------------------
# Gráfico (SVG inline, sem depender de internet)
# ----------------------------------------------------------------------
def gerar_svg_linha(df, largura=700, altura=260):
    pad_x, pad_top, pad_bottom = 16, 16, 34
    vmin, vmax = df['valor'].min(), df['valor'].max()
    margem = (vmax - vmin) * 0.15 or vmax * 0.01
    vmin, vmax = vmin - margem, vmax + margem

    n = len(df)
    area_w = largura - 2 * pad_x
    area_h = altura - pad_top - pad_bottom

    def px(i):
        return pad_x + (area_w * i / (n - 1) if n > 1 else area_w / 2)

    def py(v):
        return pad_top + area_h * (1 - (v - vmin) / (vmax - vmin))

    pontos = [(px(i), py(v)) for i, v in enumerate(df['valor'])]
    linha = ' '.join(f'{x:.1f},{y:.1f}' for x, y in pontos)
    base_y = pad_top + area_h
    area = f'{pontos[0][0]:.1f},{base_y:.1f} {linha} {pontos[-1][0]:.1f},{base_y:.1f}'

    grade = ''
    for k in range(4):
        y = pad_top + area_h * k / 3
        grade += f'<line x1="{pad_x}" y1="{y:.1f}" x2="{largura - pad_x}" y2="{y:.1f}" class="grid"/>'

    bolinhas = ''
    for (x, y), (_, row) in zip(pontos, df.iterrows()):
        txt = f"{row['data'].strftime('%d/%m/%Y')}: R$ {br(row['valor'], 4)}"
        bolinhas += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="5" class="dot"><title>{txt}</title></circle>'

    ini = df['data'].iloc[0].strftime('%d/%m')
    fim = df['data'].iloc[-1].strftime('%d/%m')
    return f'''
<svg viewBox="0 0 {largura} {altura}" width="100%" preserveAspectRatio="xMidYMid meet" role="img"
     aria-label="Evolução do dólar PTAX nos últimos {DIAS_GRAFICO} dias">
  {grade}
  <polygon points="{area}" class="area"/>
  <polyline points="{linha}" class="linha"/>
  {bolinhas}
  <text x="{pad_x}" y="{altura - 10}" class="eixo">Início ({ini})</text>
  <text x="{largura - pad_x}" y="{altura - 10}" class="eixo" text-anchor="end">Hoje ({fim})</text>
  <text x="{pad_x + 4}" y="{pad_top + 12}" class="eixo-pequeno">máx R$ {br(df['valor'].max(), 4)}</text>
  <text x="{pad_x + 4}" y="{base_y - 6}" class="eixo-pequeno">mín R$ {br(df['valor'].min(), 4)}</text>
</svg>'''


# ----------------------------------------------------------------------
# Painel HTML
# ----------------------------------------------------------------------
CSS = '''
.painel{font-family:Segoe UI,Roboto,Arial,sans-serif;background:#1e1e1e;color:#e8eaed;
  padding:14px;border-radius:14px;max-width:760px;margin:auto}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:12px}
.card{background:#2f3541;border-radius:10px;padding:10px 12px}
.card .t{font-size:12px;font-weight:600;color:#c9cdd4}
.card .v{font-size:21px;font-weight:700;margin:3px 0}
.card .s{font-size:11.5px;color:#9aa1ad;line-height:1.5}
.alta{color:#f87171}.queda{color:#4ade80}
.bloco{background:#2a2a2a;border-radius:10px;padding:14px;margin-bottom:12px}
.bloco h3{margin:0 0 10px;font-size:14px}
.grid{stroke:#555;stroke-dasharray:5 5;stroke-width:1}
.area{fill:rgba(96,165,250,.12)}
.linha{fill:none;stroke:#60a5fa;stroke-width:3;stroke-linejoin:round;stroke-linecap:round}
.dot{fill:#60a5fa;opacity:0}.dot:hover{opacity:1}
.eixo{fill:#9aa1ad;font-size:13px}.eixo-pequeno{fill:#7d8490;font-size:11px}
.item{display:flex;align-items:center;gap:8px;margin:7px 0;font-size:13.5px}
.tag{font-size:11px;padding:2px 9px;border-radius:10px;font-weight:600;white-space:nowrap}
.tag.g{background:#14391f;color:#4ade80}.tag.o{background:#4a2a14;color:#fb923c}.tag.b{background:#16304f;color:#60a5fa}
.rodape{font-size:11px;color:#7d8490;text-align:right}
'''


def calcular_resumo(df_ptax):
    """Métricas do dólar usadas no painel HTML e na imagem PNG."""
    ult = df_ptax.iloc[-1]
    ant = df_ptax.iloc[-2] if len(df_ptax) > 1 else ult
    var_dia = (ult['valor'] / ant['valor'] - 1) * 100
    var_periodo = (ult['valor'] / df_ptax.iloc[0]['valor'] - 1) * 100

    if var_periodo > 0.5:
        tendencia = f'Dólar em alta no período ({fmt_sinal(var_periodo)}% em {DIAS_GRAFICO} dias)'
    elif var_periodo < -0.5:
        tendencia = f'Dólar em queda no período ({fmt_sinal(var_periodo)}% em {DIAS_GRAFICO} dias)'
    else:
        tendencia = f'Dólar estável no período ({fmt_sinal(var_periodo)}% em {DIAS_GRAFICO} dias)'
    return {'ult': ult, 'ant': ant, 'var_dia': var_dia, 'var_periodo': var_periodo, 'tendencia': tendencia}


def montar_html(df_ptax, selic, ipca):
    r = calcular_resumo(df_ptax)
    ult, ant, var_dia, tendencia = r['ult'], r['ant'], r['var_dia'], r['tendencia']
    classe_var = 'alta' if var_dia >= 0 else 'queda'

    def ref_mes(d):
        return d.strftime('%m/%Y') if d else 'N/D'

    selic_ad, selic_am, selic_aa = selic['dia'], selic['mes'], selic['ano']
    d_selic = selic['data_dia']
    ipca12, ipca_m = ipca['12m'], ipca['mes']

    cards = f'''
    <div class="card"><div class="t">PTAX</div><div class="v">R$ {br(ult['valor'], 4)}</div>
      <div class="s">Cotação de {ult['data'].strftime('%d/%m/%Y')}</div></div>
    <div class="card"><div class="t">Variação</div><div class="v {classe_var}">{fmt_sinal(var_dia)}%</div>
      <div class="s">Diária (vs. {ant['data'].strftime('%d/%m')})</div></div>
    <div class="card"><div class="t">Selic</div><div class="v">{br(selic_aa)}% a.a.</div>
      <div class="s">{br(selic_ad, 4)}% a.d.<br>{br(selic_am)}% a.m. (ref. {ref_mes(selic['data_mes'])})</div></div>
    <div class="card"><div class="t">IPCA</div><div class="v">{br(ipca12)}%</div>
      <div class="s">Acumulado 12m (ref. {ref_mes(ipca['data_12m'])})<br>Mês: {br(ipca_m)}% ({ref_mes(ipca['data_mes'])})</div></div>'''

    resumo = f'''
    <div class="item"><span class="tag g">Tendência</span><span>{tendencia}</span></div>
    <div class="item"><span class="tag o">Selic</span><span>{br(selic_ad, 4)}% a.d. · {br(selic_am)}% a.m. · {br(selic_aa)}% a.a.
      (taxa diária de {d_selic.strftime('%d/%m/%Y') if d_selic else 'N/D'})</span></div>
    <div class="item"><span class="tag b">Inflação</span><span>IPCA acumulado em 12 meses: {br(ipca12)}%
      (mês de ref. {ref_mes(ipca['data_12m'])})</span></div>'''

    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    return f'''<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Monitor BCB</title>
<style>body{{background:#121212;margin:0;padding:16px}}{CSS}</style></head><body>
<div class="painel">
  <div class="cards">{cards}</div>
  <div class="bloco"><h3>Gráfico PTAX — {DIAS_GRAFICO} dias</h3>{gerar_svg_linha(df_ptax)}</div>
  <div class="bloco"><h3>Resumo automático</h3>{resumo}</div>
  <div class="rodape">Fonte: Banco Central do Brasil (SGS) · Atualizado em {agora}<br>
    Séries: PTAX {SERIES['Dolar_PTAX']['codigo']} · Selic {SERIES['Selic_diaria']['codigo']}/{SERIES['Selic_mensal']['codigo']}/{SERIES['Selic_anual']['codigo']} ·
    IPCA {SERIES['IPCA_mensal']['codigo']}/{SERIES['IPCA_12m']['codigo']}</div>
</div></body></html>'''


# ----------------------------------------------------------------------
# Imagem PNG do painel (para README / GitHub, que não renderiza HTML)
# ----------------------------------------------------------------------
def gerar_png(df_ptax, selic, ipca, caminho):
    from matplotlib.figure import Figure
    from matplotlib.patches import FancyBboxPatch

    r = calcular_resumo(df_ptax)
    ult, ant, var_dia = r['ult'], r['ant'], r['var_dia']

    W, H = 800, 760
    fig = Figure(figsize=(8, 7.6), dpi=150, facecolor='#121212')
    bg = fig.add_axes([0, 0, 1, 1])
    bg.set_xlim(0, W)
    bg.set_ylim(H, 0)
    bg.axis('off')

    def caixa(x, y, w, h, cor, raio=10):
        bg.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f'round,pad=0,rounding_size={raio}',
                                    fc=cor, ec='none'))

    def ref_mes(d):
        return d.strftime('%m/%Y') if d else 'N/D'

    caixa(8, 8, 784, 744, '#1e1e1e', 16)

    # --- Cards ---
    cor_var = '#f87171' if var_dia >= 0 else '#4ade80'
    cards = [
        ('PTAX', f"R$ {br(ult['valor'], 4)}", '#ffffff',
         [f"Cotação de {ult['data'].strftime('%d/%m/%Y')}"]),
        ('Variação', f"{fmt_sinal(var_dia)}%", cor_var,
         [f"Diária (vs. {ant['data'].strftime('%d/%m')})"]),
        ('Selic', f"{br(selic['ano'])}% a.a.", '#ffffff',
         [f"{br(selic['dia'], 4)}% a.d.", f"{br(selic['mes'])}% a.m. (ref. {ref_mes(selic['data_mes'])})"]),
        ('IPCA', f"{br(ipca['12m'])}%", '#ffffff',
         [f"Acumulado 12m (ref. {ref_mes(ipca['data_12m'])})",
          f"Mês: {br(ipca['mes'])}% ({ref_mes(ipca['data_mes'])})"]),
    ]
    larg_card = (752 - 3 * 12) / 4
    for i, (titulo, valor, cor_valor, linhas) in enumerate(cards):
        x = 24 + i * (larg_card + 12)
        caixa(x, 24, larg_card, 116, '#2f3541')
        bg.text(x + 12, 46, titulo, color='#c9cdd4', fontsize=8, fontweight='bold', va='center')
        bg.text(x + 12, 72, valor, color=cor_valor, fontsize=14, fontweight='bold', va='center')
        for k, linha in enumerate(linhas):
            bg.text(x + 12, 98 + k * 16, linha, color='#9aa1ad', fontsize=6.6, va='center')

    # --- Gráfico PTAX ---
    caixa(24, 156, 752, 344, '#2a2a2a')
    bg.text(40, 180, f'Gráfico PTAX — {DIAS_GRAFICO} dias', color='#e8eaed', fontsize=10,
            fontweight='bold', va='center')

    ax = fig.add_axes([48 / W, 1 - 464 / H, 704 / W, 262 / H])
    ax.patch.set_alpha(0)
    xs, ys = df_ptax['data'], df_ptax['valor']
    margem = (ys.max() - ys.min()) * 0.15 or ys.max() * 0.01
    y0, y1 = ys.min() - margem, ys.max() + margem
    ax.set_ylim(y0, y1)
    ax.set_xlim(xs.iloc[0], xs.iloc[-1])
    ax.plot(xs, ys, color='#60a5fa', lw=2.4, solid_capstyle='round', solid_joinstyle='round')
    ax.fill_between(xs, ys, y0, color='#60a5fa', alpha=0.12)
    ax.scatter([xs.iloc[-1]], [ys.iloc[-1]], color='#60a5fa', s=30, zorder=3, clip_on=False)
    ax.set_yticks([y0 + (y1 - y0) * k / 3 for k in range(4)])
    ax.set_yticklabels([])
    ax.grid(axis='y', color='#555555', linestyle=(0, (5, 5)), linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_xticks([xs.iloc[0], xs.iloc[-1]])
    ax.set_xticklabels([f"Início ({xs.iloc[0].strftime('%d/%m')})", f"Hoje ({xs.iloc[-1].strftime('%d/%m')})"],
                       color='#9aa1ad', fontsize=8)
    for lbl, ha in zip(ax.get_xticklabels(), ['left', 'right']):
        lbl.set_ha(ha)
    ax.tick_params(length=0, pad=8)
    ax.text(0.01, 0.97, f"máx R$ {br(ys.max(), 4)}", transform=ax.transAxes, color='#7d8490',
            fontsize=7, va='top')
    ax.text(0.01, 0.03, f"mín R$ {br(ys.min(), 4)}", transform=ax.transAxes, color='#7d8490',
            fontsize=7, va='bottom')

    # --- Resumo automático ---
    caixa(24, 516, 752, 176, '#2a2a2a')
    bg.text(40, 540, 'Resumo automático', color='#e8eaed', fontsize=10, fontweight='bold', va='center')
    d_selic = selic['data_dia'].strftime('%d/%m/%Y') if selic['data_dia'] else 'N/D'
    linhas_resumo = [
        ('Tendência', '#14391f', '#4ade80', r['tendencia']),
        ('Selic', '#4a2a14', '#fb923c',
         f"{br(selic['dia'], 4)}% a.d. · {br(selic['mes'])}% a.m. · {br(selic['ano'])}% a.a. (taxa diária de {d_selic})"),
        ('Inflação', '#16304f', '#60a5fa',
         f"IPCA acumulado em 12 meses: {br(ipca['12m'])}% (mês de ref. {ref_mes(ipca['data_12m'])})"),
    ]
    for k, (tag, fundo, cor, texto) in enumerate(linhas_resumo):
        y = 574 + k * 36
        caixa(40, y - 11, 72, 22, fundo, 11)
        bg.text(76, y, tag, color=cor, fontsize=7.5, fontweight='bold', ha='center', va='center')
        bg.text(124, y, texto, color='#e8eaed', fontsize=8.5, va='center')

    # --- Rodapé ---
    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    bg.text(776, 712, f"Fonte: Banco Central do Brasil (SGS) · Atualizado em {agora}", color='#7d8490',
            fontsize=6.5, ha='right', va='center')
    bg.text(776, 728, (f"Séries: PTAX {SERIES['Dolar_PTAX']['codigo']} · "
                       f"Selic {SERIES['Selic_diaria']['codigo']}/{SERIES['Selic_mensal']['codigo']}/"
                       f"{SERIES['Selic_anual']['codigo']} · "
                       f"IPCA {SERIES['IPCA_mensal']['codigo']}/{SERIES['IPCA_12m']['codigo']}"),
            color='#7d8490', fontsize=6.5, ha='right', va='center')

    fig.savefig(caminho, dpi=150, facecolor=fig.get_facecolor())


# ----------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------
def gerar_excel(df_ptax):
    conn = sqlite3.connect(CAMINHO_DB)
    df_hist = pd.read_sql_query(
        'SELECT indicador, serie_sgs, unidade, data_referencia, valor, data_execucao_rpa '
        'FROM historico_indicadores ORDER BY indicador, data_referencia DESC', conn)
    conn.close()

    df_resumo = (df_ptax.assign(indicador='Dolar_PTAX')
                 .groupby('indicador')['valor']
                 .agg(registros='count', menor='min', maior='max', media='mean').round(4).reset_index())

    with pd.ExcelWriter(CAMINHO_EXCEL, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='Resumo_PTAX_30d', index=False)
        df_hist.to_excel(writer, sheet_name='Historico', index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = 'A2'
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value or '')) for c in col) + 4


# ----------------------------------------------------------------------
# Execução
# ----------------------------------------------------------------------
def main():
    # 1) PTAX: sempre os últimos 30 dias
    df_ptax = buscar_periodo('Dolar_PTAX', DIAS_GRAFICO)

    # 2) Selic e IPCA: sempre consultados a cada execução
    d_dia, v_dia = buscar_ultimo('Selic_diaria')
    d_mes, v_mes = buscar_ultimo('Selic_mensal')
    d_ano, v_ano = buscar_ultimo('Selic_anual')
    selic = {'dia': v_dia, 'mes': v_mes, 'ano': v_ano, 'data_dia': d_dia, 'data_mes': d_mes, 'data_ano': d_ano}

    d_im, v_im = buscar_ultimo('IPCA_mensal')
    d_i12, v_i12 = buscar_ultimo('IPCA_12m')
    ipca = {'mes': v_im, '12m': v_i12, 'data_mes': d_im, 'data_12m': d_i12}

    # 3) Histórico no SQLite
    registros = [('Dolar_PTAX', r['data'], r['valor']) for _, r in df_ptax.iterrows()]
    for nome, (d, v) in {
        'Selic_diaria': (d_dia, v_dia), 'Selic_mensal': (d_mes, v_mes), 'Selic_anual': (d_ano, v_ano),
        'IPCA_mensal': (d_im, v_im), 'IPCA_12m': (d_i12, v_i12),
    }.items():
        if d is not None:
            registros.append((nome, d, v))
    salvar_historico(registros)

    # 4) Painel + Excel
    html = montar_html(df_ptax, selic, ipca)
    with open(CAMINHO_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    gerar_png(df_ptax, selic, ipca, CAMINHO_PNG)
    gerar_excel(df_ptax)

    print(f'Painel: {CAMINHO_HTML}\nImagem: {CAMINHO_PNG}\nExcel:  {CAMINHO_EXCEL}\nBanco:  {CAMINHO_DB}')

    try:  # exibe direto no Colab/Jupyter
        from IPython.display import HTML, display
        display(HTML(html))
    except ImportError:
        pass


if __name__ == '__main__':
    main()
