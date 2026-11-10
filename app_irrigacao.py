import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import math
import numpy as np
import io
import zipfile
import html
import os
import json
import unicodedata
import urllib.request
from datetime import datetime

st.set_page_config(
    page_title="Sistema de Irrigação Inteligente - Engenharia & Compliance", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("Sistema Integrado de Dimensionamento de Irrigação")
st.markdown("""
**Plataforma de Engenharia:** Conecta de ponta a ponta as necessidades da cultura, a física do solo, 
a hidráulica das tubulações, a captação de água de chuva e as instalações elétricas.
""")
st.markdown("---")

# ==============================================================================
# CAMINHOS DOS ARQUIVOS DE DADOS (os caminhos novos vêm primeiro; os antigos foram mantidos)
# Também procura na pasta do próprio script e na pasta de execução.
# ==============================================================================
PASTA_PROJETO = r"C:\Users\zoldy\Desktop\Projeto Irrigacao - Scripts"
PASTA_FACULDADE = r"C:\Users\zoldy\Desktop\Faculdade\Hidraulica e Irrigacao"
PASTA_TRABALHO = r"C:\Users\zoldy\Desktop\Trabalho hidraulicas"
PASTA_DESKTOP = r"C:\Users\zoldy\Desktop"
PASTA_DOWNLOADS = r"C:\Users\zoldy\Downloads"

def _pasta_do_script():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()

def _montar_caminhos(pastas, nomes):
    pastas_busca = [_pasta_do_script(), os.getcwd()] + list(pastas)
    return [os.path.join(p, n) for p in pastas_busca for n in nomes]

CAMINHOS_AGRO = _montar_caminhos(
    [PASTA_PROJETO, PASTA_TRABALHO, PASTA_DESKTOP],
    ["irrigacao_solo_cultura.xlsx", "irrigacao_solo_cultura (1).xlsx", "irrigacao_solo_cultura__1_.xlsx"]
)
CAMINHOS_MUNICIPIOS = _montar_caminhos([PASTA_PROJETO, PASTA_TRABALHO, PASTA_DESKTOP], ["municipios.csv"])
CAMINHOS_CHUVA = _montar_caminhos([PASTA_PROJETO, PASTA_TRABALHO, PASTA_DESKTOP], ["chuva_historica_municipios.csv"])
# Banco de bombas comerciais (curvas das tabelas de seleção dos fabricantes). Fica na mesma pasta do app ou nas pastas do projeto.
CAMINHOS_BOMBAS = _montar_caminhos([PASTA_PROJETO, PASTA_TRABALHO, PASTA_DESKTOP, PASTA_DOWNLOADS], ["bombas_comerciais.xlsx"])
# Planilha "Motor de Cálculo v7 Compliance": os cálculos dela foram incorporados ao código (abas 3, 6 e Laudo).
_caminho_motor_local = os.path.join(_pasta_do_script(), "Motor_de_Calculo_Irrigacao_v7_Compliance.xlsx")
CAMINHO_MOTOR_REF = _caminho_motor_local if os.path.exists(_caminho_motor_local) else os.path.join(PASTA_FACULDADE, "Motor_de_Calculo_Irrigacao_v7_Compliance.xlsx")

UF_POR_CODIGO = {
    11: 'RO', 12: 'AC', 13: 'AM', 14: 'RR', 15: 'PA', 16: 'AP', 17: 'TO',
    21: 'MA', 22: 'PI', 23: 'CE', 24: 'RN', 25: 'PB', 26: 'PE', 27: 'AL', 28: 'SE', 29: 'BA',
    31: 'MG', 32: 'ES', 33: 'RJ', 35: 'SP', 41: 'PR', 42: 'SC', 43: 'RS',
    50: 'MS', 51: 'MT', 52: 'GO', 53: 'DF'
}
MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
DIAS_MES = [31, 28.25, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

def _n(x, d=2):
    """Formata número no padrão brasileiro (1.234,56). Usado em textos, avisos e no laudo."""
    try:
        v = f"{float(x):,.{d}f}"
    except Exception:
        return str(x)
    return v.replace(",", "§").replace(".", ",").replace("§", ".")


def fator_christiansen(n_saidas, primeiro_a_meio=False, m=1.852):
    """Fator F de Christiansen para tubo com N saídas igualmente espaçadas (perda de carga com expoente m).
    Primeira saída a uma distância S da entrada: F = 1/(m+1) + 1/(2N) + √(m−1)/(6N²).
    Primeira saída a S/2: F = 2N/(2N−1) · [1/(m+1) + √(m−1)/(6N²)]. Para N = 1, F = 1."""
    N = max(1, int(n_saidas))
    if N == 1:
        return 1.0
    if primeiro_a_meio:
        return (2.0 * N / (2.0 * N - 1.0)) * (1.0 / (m + 1.0) + math.sqrt(m - 1.0) / (6.0 * N ** 2))
    return 1.0 / (m + 1.0) + 1.0 / (2.0 * N) + math.sqrt(m - 1.0) / (6.0 * N ** 2)


@st.cache_data(show_spinner=False)
def _ler_bombas(caminho, _mtime):
    """Lê a aba 'Bombas' da planilha. Cada curva vem como 'H:Q;H:Q' (altura em mca : vazão em m³/h)."""
    df = pd.read_excel(caminho, sheet_name="Bombas")
    col_pts = [c for c in df.columns if str(c).startswith("Pontos da curva")][0]
    def _f(v):
        try:
            x = float(str(v).replace(",", "."))
            return x if x == x else None
        except Exception:
            return None
    lista = []
    for _, r in df.iterrows():
        pts = []
        for p in str(r[col_pts]).split(";"):
            if ":" in p:
                a, b = p.split(":")[:2]
                if _f(a) is not None and _f(b) is not None:
                    pts.append((_f(a), _f(b)))
        if len(pts) < 3 or _f(r.get("Potência (cv)")) is None:
            continue
        pts.sort(key=lambda t: t[1])  # vazão crescente
        lista.append(dict(id=str(r["ID"]), marca=str(r["Marca"]), serie=str(r["Série"]), modelo=str(r["Modelo"]), cv=_f(r["Potência (cv)"]),
                          suc=str(r.get("Ø sucção (pol)") if str(r.get("Ø sucção (pol)")) != "nan" else ""),
                          rec=str(r.get("Ø recalque (pol)") if str(r.get("Ø recalque (pol)")) != "nan" else ""),
                          pmax=_f(r.get("Pressão máx. (mca)")), preco=_f(r.get("Preço (R$)")), npshr=_f(r.get("NPSH requerido (mca)")),
                          q=np.array([t[1] for t in pts]), h=np.array([t[0] for t in pts]), pag=r.get("Pág. PDF")))
    return lista

def carregar_bombas():
    for caminho in CAMINHOS_BOMBAS:
        try:
            if os.path.exists(caminho):
                lista = _ler_bombas(caminho, os.path.getmtime(caminho))
                if lista:
                    return lista, caminho
        except Exception:
            continue
    return [], None


def _normalizar(texto):
    t = unicodedata.normalize('NFD', str(texto))
    return ''.join(c for c in t if unicodedata.category(c) != 'Mn').lower().strip()

# ==============================================================================
# CARREGAMENTO DE DADOS HISTÓRICOS E AGRONÔMICOS
# ==============================================================================
@st.cache_data(show_spinner=False)
def carregar_dados_pluviometricos():
    caminhos = CAMINHOS_CHUVA + [
        r"C:\Users\zoldy\Desktop\Trabalho hidraulicas\chuva_historica_municipios.csv",
        r"C:\Users\zoldy\Desktop\chuva_historica_municipios.csv"
    ]
    for caminho in caminhos:
        for codificacao in ('utf-8-sig', 'latin-1'):
            try:
                df = pd.read_csv(caminho, encoding=codificacao)
                if not df.empty:
                    df['origem'] = 'arquivo'
                    return df
            except Exception:
                continue
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mock_data = {
        'uf': ['PI']*12 + ['MA']*12,
        'cidade': ['Teresina']*12 + ['Timon']*12,
        'mes': meses * 2,
        'chuva_historica_mm_mes_estimada': [250, 280, 300, 220, 100, 30, 15, 10, 20, 50, 100, 180,
                                            240, 270, 290, 210, 95, 25, 10, 5, 15, 45, 90, 170]
    }
    df_mock = pd.DataFrame(mock_data)
    df_mock['origem'] = 'exemplo'
    return df_mock

@st.cache_data(show_spinner=False)
def carregar_dados_agronomicos():
    caminhos = CAMINHOS_AGRO + [
        r"C:\Users\zoldy\Desktop\Trabalho hidraulicas\irrigacao_solo_cultura (1).xlsx",
        r"C:\Users\zoldy\Desktop\Trabalho hidraulicas\irrigacao_solo_cultura.xlsx",
        r"C:\Users\zoldy\Desktop\irrigacao_solo_cultura (1).xlsx",
        r"C:\Users\zoldy\Desktop\irrigacao_solo_cultura.xlsx"
    ]
    for caminho_excel in caminhos:
        try:
            df_matriz = pd.read_excel(caminho_excel, sheet_name="Matriz Solo x Cultura", skiprows=1)
            df_matriz.columns = df_matriz.iloc[0].astype(str)
            df_matriz = df_matriz[1:].dropna(subset=['Cultura']).reset_index(drop=True)
            
            df_solo = pd.read_excel(caminho_excel, sheet_name="Características do Solo", skiprows=1)
            df_solo.columns = df_solo.iloc[0].astype(str)
            df_solo = df_solo[1:].dropna(subset=['Característica']).reset_index(drop=True)
            return df_matriz, df_solo
        except Exception:
            continue
    return pd.DataFrame(), pd.DataFrame()

@st.cache_data(show_spinner=False)
def carregar_dados_agronomicos_extra():
    """Abas adicionais da planilha agronômica: Top 10 Nordeste, Infiltração (VIB Embrapa) e Fontes e Notas."""
    for caminho_excel in CAMINHOS_AGRO:
        try:
            df_top10 = pd.read_excel(caminho_excel, sheet_name="Top 10 Nordeste", skiprows=2).iloc[:, :5]
            df_top10 = df_top10.dropna(subset=['Cultura']).reset_index(drop=True)
            df_infil = pd.read_excel(caminho_excel, sheet_name="Infiltração", skiprows=3, nrows=3)
            df_fontes = pd.read_excel(caminho_excel, sheet_name="Fontes e Notas", header=None)
            fontes = [str(x) for x in df_fontes[0].dropna().tolist()]
            return df_top10, df_infil, fontes
        except Exception:
            continue
    return pd.DataFrame(), pd.DataFrame(), []

@st.cache_data(show_spinner=False)
def carregar_catalogo_municipios():
    """Todos os municípios do Brasil (IBGE) com UF e coordenadas, a partir do municipios.csv."""
    for caminho in CAMINHOS_MUNICIPIOS:
        df = None
        for codificacao in ('utf-8-sig', 'latin-1'):
            try:
                df = pd.read_csv(caminho, encoding=codificacao)
                break
            except Exception:
                df = None
        if df is None or df.empty:
            continue
        if not {'nome', 'codigo_uf', 'latitude', 'longitude'}.issubset(df.columns):
            continue
        df = df.copy()
        df['uf'] = df['codigo_uf'].map(UF_POR_CODIGO)
        df = df.dropna(subset=['uf']).rename(columns={'nome': 'cidade'})
        df['chave'] = df['uf'] + '|' + df['cidade'].map(_normalizar)
        return df[['codigo_ibge', 'cidade', 'uf', 'latitude', 'longitude', 'capital', 'chave']].reset_index(drop=True)
    return pd.DataFrame()

# ------------------------------------------------------------------------------
# Chuva para QUALQUER município: histórico (arquivo) -> NASA POWER (online) -> vizinho mais próximo
# ------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def montar_vizinhos_historico():
    hist = carregar_dados_pluviometricos()
    cat = carregar_catalogo_municipios()
    if hist.empty or cat.empty:
        return pd.DataFrame()
    h = hist[['uf', 'cidade']].drop_duplicates().copy()
    h['chave'] = h['uf'].astype(str).str.upper() + '|' + h['cidade'].map(_normalizar)
    h = h.rename(columns={'cidade': 'cidade_hist'})
    return h.merge(cat[['chave', 'latitude', 'longitude']], on='chave', how='inner').reset_index(drop=True)

@st.cache_data(show_spinner=False, ttl=60 * 60 * 24 * 30)
def buscar_chuva_nasa_power(lat, lon):
    """Climatologia mensal de precipitação (mm/dia -> mm/mês) no ponto. Levanta exceção se falhar (não vai para o cache)."""
    url = ("https://power.larc.nasa.gov/api/temporal/climatology/point?parameters=PRECTOTCORR&community=AG"
           f"&longitude={float(lon):.4f}&latitude={float(lat):.4f}&format=JSON")
    with urllib.request.urlopen(url, timeout=12) as resp:
        dados = json.loads(resp.read().decode('utf-8'))
    serie = dados['properties']['parameter']['PRECTOTCORR']
    chaves = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    valores = [float(serie[k]) for k in chaves]
    if any(v < 0 for v in valores):
        raise ValueError("Valor ausente (fill value) na resposta da NASA POWER")
    return [v * DIAS_MES[i] for i, v in enumerate(valores)]

def _vizinho_mais_proximo(lat, lon):
    viz = montar_vizinhos_historico()
    if viz.empty:
        return None
    lat1, lon1 = np.radians(float(lat)), np.radians(float(lon))
    lat2, lon2 = np.radians(viz['latitude'].astype(float).values), np.radians(viz['longitude'].astype(float).values)
    a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lon2 - lon1) / 2) ** 2
    dist_km = 6371.0 * 2 * np.arcsin(np.sqrt(a))
    i = int(np.argmin(dist_km))
    hist = carregar_dados_pluviometricos()
    sel = hist[(hist['uf'].astype(str).str.upper() == viz.loc[i, 'uf']) &
               (hist['cidade'].map(_normalizar) == _normalizar(viz.loc[i, 'cidade_hist']))].copy()
    return sel, viz.loc[i, 'cidade_hist'], viz.loc[i, 'uf'], float(dist_km[i])

def obter_chuva_municipio(uf, cidade, lat, lon, usar_online=True):
    """Retorna (DataFrame no formato do histórico, descrição da fonte, observação)."""
    hist = carregar_dados_pluviometricos()
    exato = hist[(hist['uf'].astype(str).str.upper() == uf) & (hist['cidade'].map(_normalizar) == _normalizar(cidade))].copy()
    if not exato.empty:
        origem = exato['origem'].iloc[0] if 'origem' in exato.columns else 'arquivo'
        if origem == 'arquivo':
            return exato, "Histórico municipal (arquivo CSV)", ""
        return exato, "Dados de exemplo embutidos (arquivo de chuva não encontrado)", \
            "O arquivo chuva_historica_municipios.csv não foi localizado; estes valores são apenas ilustrativos."
    nota = ""
    if usar_online:
        try:
            valores = buscar_chuva_nasa_power(lat, lon)
            df = pd.DataFrame({'uf': uf, 'cidade': cidade, 'mes': MESES_PT, 'chuva_historica_mm_mes_estimada': valores})
            return df, "NASA POWER (climatologia no ponto do município)", \
                "Município sem histórico no CSV: chuva estimada por climatologia de satélite/reanálise nas coordenadas do IBGE."
        except Exception as erro:
            nota = f"Consulta online indisponível ({type(erro).__name__}). "
    viz = _vizinho_mais_proximo(lat, lon)
    if viz is not None and not viz[0].empty:
        sel, nome_v, uf_v, dist = viz
        sel['uf'] = uf
        sel['cidade'] = cidade
        return sel, f"Estimativa pelo município mais próximo com histórico: {nome_v}/{uf_v} ({dist:.0f} km)", \
            nota + "Valor aproximado; prefira cadastrar o histórico real do município."
    return None, "Sem dados de chuva", nota + "Nenhuma fonte de chuva disponível para este município."

# ==============================================================================
# DADOS E FUNÇÕES DA ABA "ESQUEMÁTICO DA TUBULAÇÃO"
# Carrega a planilha dados_tubulacao_irrigacao.xlsx e define as funções de cálculo.
# Fontes dos dados: [1] Apostila Hidráulica UFPel 2018-2 (Tab. 1D, 1E, 1F; eq. 30)
#                   [2] Macintyre, Bombas e Instalações de Bombeamento (p. 654-655)
# ==============================================================================
CAMINHOS_TUBOS = _montar_caminhos([PASTA_DOWNLOADS, PASTA_PROJETO, PASTA_TRABALHO, PASTA_DESKTOP], ["dados_tubulacao_irrigacao.xlsx"])

# Valores padrão (usados só se a planilha não for encontrada) — mesmos da planilha
_MATERIAIS_PADRAO = {"PVC": 140.0, "PEAD / PEBD": 140.0, "Aço galvanizado": 125.0, "Alumínio": 120.0}
_DNS_PADRAO = [(13, 13.0), (19, 19.0), (25, 25.0), (32, 32.0), (38, 38.0), (50, 50.0), (63, 63.0), (75, 75.0),
               (100, 100.0), (125, 125.0), (150, 150.0), (200, 200.0), (250, 250.0), (300, 300.0), (350, 350.0)]
_PARAM_TUBO_PADRAO = {"K_HW": 10.646, "D_MIN_HW_MM": 50.0, "V_MIN": 0.6, "V_MAX": 2.0, "V_RECALQUE_ABS": 3.0,
                      "COMP_BARRA_M": 6.0, "SOBRA_CORTE_PCT": 5.0, "LEQ_LUVA_M": 0.0}
# Tabela 1F de [1]: comprimento equivalente em nº de diâmetros (Lf = n·D)
_ND_PADRAO = {"Cotovelo de 90°": 45.0, "Cotovelo de 45°": 20.0, "Tê, passagem direta": 20.0, "Tê, saída de lado": 50.0,
              "Registro de gaveta, aberto": 8.0, "Válvula de retenção": 100.0}

# Peça (rótulo da tela) -> (coluna da Tabela 1E, linha da Tabela 1F)
PECAS_TUBO = {
    "Cotovelo 90° (raio longo)": ("Cotovelo 90° raio longo", "Cotovelo de 90°"),
    "Cotovelo 90° (raio médio)": ("Cotovelo 90° raio médio", "Cotovelo de 90°"),
    "Cotovelo 45°": ("Cotovelo 45°", "Cotovelo de 45°"),
    "Tê passagem direta": ("Tê passagem direta", "Tê, passagem direta"),
    "Tê saída de lado": ("Tê saída de lado", "Tê, saída de lado"),
    "Registro de gaveta aberto": ("Válvula de gaveta aberta", "Registro de gaveta, aberto"),
    "Válvula de retenção leve": ("Válvula de retenção leve", "Válvula de retenção"),
}

@st.cache_data(show_spinner=False)
def carregar_dados_tubulacao():
    """Lê as abas Materiais, Diametros, Comp_Equivalentes, Comp_Equiv_nD e Parametros. Se faltar, usa os padrões."""
    for caminho in CAMINHOS_TUBOS:
        try:
            mat = pd.read_excel(caminho, sheet_name="Materiais")
            dia = pd.read_excel(caminho, sheet_name="Diametros")
            leq = pd.read_excel(caminho, sheet_name="Comp_Equivalentes")
            par = pd.read_excel(caminho, sheet_name="Parametros")
            nd = pd.read_excel(caminho, sheet_name="Comp_Equiv_nD")

            c_col = pd.to_numeric(mat.iloc[:, 1], errors="coerce")
            materiais = {str(n): float(c) for n, c in zip(mat.iloc[:, 0], c_col) if pd.notna(c)}

            dn = pd.to_numeric(dia.iloc[:, 0], errors="coerce")
            di = pd.to_numeric(dia.iloc[:, 2], errors="coerce")
            dns = [(int(a), float(b)) for a, b in zip(dn, di) if pd.notna(a) and pd.notna(b)]

            primeira = leq.columns[0]
            leq["_dn"] = pd.to_numeric(leq[primeira], errors="coerce")
            leq = leq.dropna(subset=["_dn"]).drop(columns=[primeira]).set_index("_dn").sort_index()
            leq.columns = [str(c).replace(" (m)", "") for c in leq.columns]
            leq = leq.apply(pd.to_numeric, errors="coerce")

            v_par = pd.to_numeric(par.iloc[:, 1], errors="coerce")
            parametros = {**_PARAM_TUBO_PADRAO, **{str(k): float(v) for k, v in zip(par.iloc[:, 0], v_par) if pd.notna(v)}}
            v_nd = pd.to_numeric(nd.iloc[:, 1], errors="coerce")
            tabela_nd = {**_ND_PADRAO, **{str(k): float(v) for k, v in zip(nd.iloc[:, 0], v_nd) if pd.notna(v)}}

            if materiais and dns:
                return {"materiais": materiais, "dns": dns, "leq": leq, "par": parametros, "nd": tabela_nd,
                        "origem": f"Planilha {os.path.basename(caminho)}"}
        except Exception:
            continue
    return {"materiais": dict(_MATERIAIS_PADRAO), "dns": list(_DNS_PADRAO), "leq": None,
            "par": dict(_PARAM_TUBO_PADRAO), "nd": dict(_ND_PADRAO),
            "origem": "Valores padrão embutidos (dados_tubulacao_irrigacao.xlsx não encontrada)"}

def leq_peca(dados, peca, di_mm):
    """Comprimento equivalente (m) de uma peça: Tabela 1E de [1] (interpolada no diâmetro); sem a planilha, Tabela 1F (n·D)."""
    coluna, chave_nd = PECAS_TUBO[peca]
    leq = dados["leq"]
    if leq is not None and coluna in leq.columns:
        s = leq[coluna].dropna()
        if len(s) >= 2:
            return float(np.interp(di_mm, s.index.values.astype(float), s.values.astype(float)))
    return dados["nd"].get(chave_nd, 0.0) * di_mm / 1000.0

def hazen_williams_hf(q_m3h, di_mm, c, comp_m, k=10.646):
    """Perda de carga contínua (mca), eq. (30) de [1]: hf = 10,646 · L · (Q/C)^1,852 / D^4,87 (SI)."""
    if q_m3h <= 0 or di_mm <= 0 or c <= 0 or comp_m <= 0:
        return 0.0
    return k * comp_m * ((q_m3h / 3600.0) / c) ** 1.852 / ((di_mm / 1000.0) ** 4.87)

def gerar_svg_planta_L(l1, l2, lado, n_cotovelos, n_tes):
    """Vista em planta: bomba -> trecho principal L1 (frente) -> desvio L2 para a Direita/Esquerda."""
    direita = (lado == "Direita")
    x0, y0 = (90 if direita else 330), 300
    k = 230.0 / max(l1, l2, 1.0)
    s1, s2 = max(40.0, l1 * k), max(40.0, l2 * k)
    sinal = 1 if direita else -1
    x1, y1 = x0 + sinal * s2, y0 - s1
    gotas = "".join(f'<circle cx="{x0 + sinal * d}" cy="{y1}" r="3" fill="#3498DB"/>' for d in range(20, int(s2), 22))
    seta = f'<polygon points="{x1},{y1 - 9} {x1 + sinal * 16},{y1} {x1},{y1 + 9}" fill="#27AE60"/>'
    texto_lado = "DIREITA →" if direita else "← ESQUERDA"
    xt = (x0 + x1) / 2
    svg = f"""
    <div style="display:flex;justify-content:center;margin:10px 0;">
    <svg width="100%" viewBox="0 0 420 340" style="max-width:460px;border:1px solid #ddd;background:#FDFEFE;border-radius:5px;" xmlns="http://www.w3.org/2000/svg">
      <text x="210" y="22" font-family="Arial" font-size="13" font-weight="bold" fill="#2C3E50" text-anchor="middle">Planta: trajeto em L (desvio à {lado})</text>
      <line x1="{x0}" y1="{y0}" x2="{x0}" y2="{y1}" stroke="#34495E" stroke-width="6"/>
      <line x1="{x0}" y1="{y1}" x2="{x1}" y2="{y1}" stroke="#7F8C8D" stroke-width="6"/>
      {gotas}{seta}
      <rect x="{x0 - 20}" y="{y0}" width="40" height="24" rx="5" fill="#34495E"/>
      <text x="{x0}" y="{y0 + 16}" font-family="Arial" font-size="10" font-weight="bold" fill="#FFFFFF" text-anchor="middle">BOMBA</text>
      <circle cx="{x0}" cy="{y1}" r="8" fill="#E67E22"/>
      <text x="{x0 + sinal * 14}" y="{y1 + 24}" font-family="Arial" font-size="10" fill="#E67E22" text-anchor="{'start' if direita else 'end'}">cotovelo ({n_cotovelos} un)</text>
      <text x="{x0 - sinal * 10}" y="{(y0 + y1) / 2}" font-family="Arial" font-size="12" font-weight="bold" fill="#34495E" text-anchor="{'end' if direita else 'start'}">L1 = {l1:.1f} m</text>
      <text x="{xt}" y="{y1 - 14}" font-family="Arial" font-size="12" font-weight="bold" fill="#566573" text-anchor="middle">L2 = {l2:.1f} m</text>
      <text x="{xt}" y="{y1 - 30}" font-family="Arial" font-size="11" fill="#27AE60" text-anchor="middle">{texto_lado}</text>
      <text x="12" y="52" font-family="Arial" font-size="10" fill="#7F8C8D">↑ Frente (eixo principal)</text>
      <text x="12" y="66" font-family="Arial" font-size="10" fill="#7F8C8D">Tês: {n_tes} un</text>
    </svg></div>
    """
    return "".join(linha.strip() for linha in svg.splitlines())

def gerar_svg_planta_multi(comprimentos, direcoes, n_cotovelos, n_tes):
    """Vista em planta com N trechos: sai da bomba para a frente; cada curva vira à Direita ou à Esquerda do sentido atual."""
    vx, vy = 0.0, -1.0
    ref = max(max(comprimentos), 1.0) * 0.12
    pts = [(0.0, 0.0)]
    dirs = []
    for i, c in enumerate(comprimentos):
        if i > 0:
            vx, vy = (-vy, vx) if direcoes[i - 1] == "Direita" else (vy, -vx)
        dirs.append((vx, vy))
        lv = max(c, ref)
        pts.append((pts[-1][0] + vx * lv, pts[-1][1] + vy * lv))
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    w, h = max(max(xs) - min(xs), 1.0), max(max(ys) - min(ys), 1.0)
    k = min(330.0 / w, 220.0 / h)
    ox, oy = 45 + (330 - w * k) / 2 - min(xs) * k, 70 + (220 - h * k) / 2 - min(ys) * k
    P = [(ox + x * k, oy + y * k) for x, y in pts]
    linhas = ""
    rotulos = ""
    for i in range(len(comprimentos)):
        (xa, ya), (xb, yb) = P[i], P[i + 1]
        cor = "#34495E" if i == 0 else "#7F8C8D"
        linhas += f'<line x1="{xa:.1f}" y1="{ya:.1f}" x2="{xb:.1f}" y2="{yb:.1f}" stroke="{cor}" stroke-width="6"/>'
        dx, dy = dirs[i]
        mx, my = (xa + xb) / 2 - dy * 14, (ya + yb) / 2 + dx * 14
        rotulos += f'<text x="{mx:.1f}" y="{my:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#566573" text-anchor="middle">L{i + 1} = {comprimentos[i]:.1f} m</text>'
    curvas = "".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="7" fill="#E67E22"/>' for x, y in P[1:-1])
    xe, ye = P[-1]
    dxe, dye = dirs[-1]
    seta = f'<polygon points="{xe + dxe * 14:.1f},{ye + dye * 14:.1f} {xe - dye * 8:.1f},{ye + dxe * 8:.1f} {xe + dye * 8:.1f},{ye - dxe * 8:.1f}" fill="#27AE60"/>'
    x0, y0 = P[0]
    svg = f"""
    <div style="display:flex;justify-content:center;margin:10px 0;">
    <svg width="100%" viewBox="0 0 420 340" style="max-width:460px;border:1px solid #ddd;background:#FDFEFE;border-radius:5px;" xmlns="http://www.w3.org/2000/svg">
      <text x="210" y="22" font-family="Arial" font-size="13" font-weight="bold" fill="#2C3E50" text-anchor="middle">Planta: {len(comprimentos)} trechos e {len(comprimentos) - 1} curva(s)</text>
      {linhas}{curvas}{seta}{rotulos}
      <rect x="{x0 - 20:.1f}" y="{y0 - 4:.1f}" width="40" height="24" rx="5" fill="#34495E"/>
      <text x="{x0:.1f}" y="{y0 + 12:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#FFFFFF" text-anchor="middle">BOMBA</text>
      <text x="12" y="44" font-family="Arial" font-size="10" fill="#7F8C8D">Cotovelos: {n_cotovelos} un · Tês: {n_tes} un</text>
      <text x="12" y="58" font-family="Arial" font-size="10" fill="#7F8C8D">Curvas à Direita/Esquerda do sentido de avanço</text>
    </svg></div>
    """
    return "".join(linha.strip() for linha in svg.splitlines())

def gerar_svg_perfil(l1, l2, dz, aclive, comprimentos=None):
    """Vista em perfil: mostra se o ponto final está acima (aclive) ou abaixo (declive) da bomba."""
    dy = 0.0 if dz <= 0 else max(12.0, min(110.0, dz * 6.0))
    if dz <= 0:
        ya = yb = 140.0
    elif aclive:
        ya, yb = 190.0, 190.0 - dy
    else:
        ya, yb = 80.0, 80.0 + dy
    xa, xb = 40.0, 380.0
    if comprimentos and len(comprimentos) > 1:
        total_c, acum, fracoes = max(sum(comprimentos), 0.001), 0.0, []
        for c in comprimentos[:-1]:
            acum += c
            fracoes.append(acum / total_c)
    else:
        fracoes = [l1 / max(l1 + l2, 0.001)]
    marcas = ""
    for j, f in enumerate(fracoes):
        xm = xa + (xb - xa) * f
        ym = ya + (yb - ya) * f
        marcas += f'<circle cx="{xm:.1f}" cy="{ym:.1f}" r="6" fill="#E67E22"/>'
        if len(fracoes) == 1:
            marcas += f'<text x="{xm:.1f}" y="{ym - 12:.1f}" font-family="Arial" font-size="10" fill="#E67E22" text-anchor="middle">fim de L1 / início de L2</text>'
        else:
            marcas += f'<text x="{xm:.1f}" y="{ym - 10:.1f}" font-family="Arial" font-size="9" fill="#E67E22" text-anchor="middle">curva {j + 1}</text>'
    rotulo = "sem desnível" if dz <= 0 else (f"ACLIVE: ponto final {dz:.1f} m acima da bomba" if aclive else f"DECLIVE: ponto final {dz:.1f} m abaixo da bomba")
    cor = "#C0392B" if (dz > 0 and aclive) else "#1E8449"
    svg = f"""
    <div style="display:flex;justify-content:center;margin:10px 0;">
    <svg width="100%" viewBox="0 0 420 250" style="max-width:460px;border:1px solid #ddd;background:#FDFEFE;border-radius:5px;" xmlns="http://www.w3.org/2000/svg">
      <text x="210" y="22" font-family="Arial" font-size="13" font-weight="bold" fill="#2C3E50" text-anchor="middle">Perfil topográfico (declividade uniforme)</text>
      <polygon points="{xa},{ya} {xb},{yb} {xb},235 {xa},235" fill="#8D6E63" opacity="0.35"/>
      <line x1="{xa}" y1="{ya}" x2="{xb}" y2="{yb}" stroke="#34495E" stroke-width="5"/>
      <rect x="{xa - 14}" y="{ya - 22}" width="40" height="20" rx="4" fill="#34495E"/>
      <text x="{xa + 6}" y="{ya - 8}" font-family="Arial" font-size="9" font-weight="bold" fill="#FFFFFF" text-anchor="middle">BOMBA</text>
      {marcas}
      <circle cx="{xb}" cy="{yb}" r="6" fill="#27AE60"/>
      <text x="{xb - 4}" y="{yb - 12}" font-family="Arial" font-size="10" fill="#1E8449" text-anchor="end">roça</text>
      <line x1="{xa}" y1="{ya}" x2="{xb}" y2="{ya}" stroke="#95A5A6" stroke-width="1" stroke-dasharray="5,4"/>
      <text x="210" y="244" font-family="Arial" font-size="11" font-weight="bold" fill="{cor}" text-anchor="middle">{rotulo}</text>
    </svg></div>
    """
    return "".join(linha.strip() for linha in svg.splitlines())

# ==============================================================================
# MOTOR VISUAL 2D/3D — geometria única compartilhada por todos os esquemáticos
# (trajeto da tubulação, setores, linhas laterais, casa, cisterna e bomba)
# ==============================================================================
PALETA_SETORES = ["#F4D03F", "#5DADE2", "#EC7063", "#AF7AC1", "#58D68D", "#F5B041", "#48C9B0", "#85929E"]
COR_SOLO = {"Arenoso": "#C9A66B", "Franco (médio)": "#7B5236", "Argiloso": "#8E4B3A"}
# (cor principal, cor escura, cor de brilho) por material de tubo
COR_TUBO = {
    "PVC": ("#2E86DE", "#1B4F8A", "#9CCBFF"),
    "PEAD / PEBD": ("#34495E", "#17202A", "#85929E"),
    "Aço galvanizado": ("#AEB6BF", "#5D6D7E", "#F2F4F4"),
    "Alumínio": ("#D5D8DC", "#7B7D7D", "#FDFEFE"),
}
CULTURAS_ARBOREAS = ("Manga", "Coco-da-baía", "Fruteiras", "Cacau", "Banana", "Uva")

def _cultura_arborea(cultura):
    return any(str(cultura).startswith(c) for c in CULTURAS_ARBOREAS)

def _hex_rgb(c):
    c = c.lstrip('#')
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))

def _misturar(c1, c2, t):
    a, b = _hex_rgb(c1), _hex_rgb(c2)
    return '#%02x%02x%02x' % tuple(int(a[i] * (1 - t) + b[i] * t) for i in range(3))

def _dist_polilinha(p, pts):
    melhor = 1e18
    for (ax, ay), (bx, by) in zip(pts[:-1], pts[1:]):
        dx, dy = bx - ax, by - ay
        l2 = dx * dx + dy * dy
        t = 0.0 if l2 == 0 else max(0.0, min(1.0, ((p[0] - ax) * dx + (p[1] - ay) * dy) / l2))
        melhor = min(melhor, math.hypot(p[0] - ax - t * dx, p[1] - ay - t * dy))
    return melhor

def calcular_layout_campo(comprimentos, direcoes, dz, aclive, comp_lateral, esp_linhas, esp_emissores,
                          n_linhas, n_setores, area_plantio, area_telhado, vol_cisterna_l):
    """Geometria em metros (x = frente, y = esquerda, z = altura). Origem na bomba."""
    comprimentos = [max(0.5, float(c)) for c in comprimentos]
    hx, hy = 1.0, 0.0
    pts, heads = [(0.0, 0.0)], []
    for i, c in enumerate(comprimentos):
        if i > 0:  # Direita = giro horário; Esquerda = anti-horário
            hx, hy = (hy, -hx) if direcoes[i - 1] == "Direita" else (-hy, hx)
        heads.append((hx, hy))
        pts.append((pts[-1][0] + hx * c, pts[-1][1] + hy * c))
    E, h = pts[-1], heads[-1]
    n = (-h[1], h[0])

    # Relevo: plano inclinado do ponto de partida (bomba) até o ponto final (E) com o desnível ΔZ
    corda = math.hypot(E[0], E[1])
    if corda >= 1.0:
        u = (E[0] / corda, E[1] / corda)
        s = (float(dz) if aclive else -float(dz)) / corda
    else:
        u, s = (1.0, 0.0), 0.0
    def z(x, y):
        return s * (x * u[0] + y * u[1])

    # Cisterna e casa (atrás da bomba)
    vol_m3 = max(float(vol_cisterna_l), 500.0) / 1000.0
    rc_real = min(4.0, max(0.8, math.sqrt(vol_m3 / (math.pi * 2.0))))
    # escala visual: em projetos grandes a cisterna cresce junto com o desenho (proporção preservada)
    ext0 = max(max(p[0] for p in pts) - min(p[0] for p in pts), max(p[1] for p in pts) - min(p[1] for p in pts), 12.0) + float(comp_lateral)
    esc_c = max(1.0, 0.7 * ext0 / 70.0)
    rc = min(6.0, rc_real * esc_c)
    hc = 2.0 * min(esc_c, 6.0 / rc_real)
    C = (-(rc + 2.5), 0.0)
    at = max(4.0, float(area_telhado))
    a_h = math.sqrt(at / 1.5)
    b_h = 1.5 * a_h
    frente = C[0] - rc - 4.0
    casa = dict(cx=frente - b_h / 2, cy=0.0, b=b_h, a=a_h, frente=frente)

    # Campo, linhas laterais e setores (mesma regra da aba 3)
    Lc = max(1.0, float(comp_lateral))
    esp_l, esp_e = max(0.2, float(esp_linhas)), max(0.1, float(esp_emissores))
    n_linhas, n_setores = max(1, int(n_linhas)), max(1, int(n_setores))
    Wp = max(float(area_plantio) / Lc, n_linhas * esp_l)
    t = [(j - (n_linhas - 1) / 2.0) * esp_l for j in range(n_linhas)]

    def cantos_campo(ho_):
        return [(E[0] + h[0] * ah + n[0] * an, E[1] + h[1] * ah + n[1] * an)
                for ah, an in ((0, -Wp / 2), (0, Wp / 2), (ho_ + Lc, Wp / 2), (ho_ + Lc, -Wp / 2))]
    pre = list(pts) + cantos_campo(2.0) + [(frente - b_h, -a_h / 2), (frente, a_h / 2), (C[0] - rc, -rc), (C[0] + rc, rc)]
    ext = max(max(p[0] for p in pre) - min(p[0] for p in pre), max(p[1] for p in pre) - min(p[1] for p in pre), 12.0)
    ho = max(2.0, 0.025 * ext)
    fs = max(1.0, ext / 70.0)

    npsec = max(1, math.ceil(n_linhas / n_setores))
    n_ef = math.ceil(n_linhas / npsec)
    setores = []
    for k in range(n_ef):
        a, b = k * npsec, min(n_linhas, (k + 1) * npsec)
        t0 = -Wp / 2 if k == 0 else (t[a - 1] + t[a]) / 2
        t1 = Wp / 2 if k == n_ef - 1 else (t[b - 1] + t[b]) / 2
        setores.append(dict(k=k + 1, linhas=list(range(a, b)), t0=t0, t1=t1, tc=(t[a] + t[b - 1]) / 2))
    n_em = max(1, int(Lc // esp_e))

    todos = list(pts) + cantos_campo(ho) + [(frente - b_h, -a_h / 2), (frente, a_h / 2),
                                           (C[0] - rc, -rc), (C[0] + rc, rc), (-1.9 * fs, -0.5 * fs), (0.3 * fs, 1.6 * fs)]
    bbox = (min(p[0] for p in todos), max(p[0] for p in todos), min(p[1] for p in todos), max(p[1] for p in todos))

    # Árvores decorativas (determinísticas) fora do trajeto, do campo e das construções
    rng = np.random.RandomState(11)
    m = 0.07 * ext
    decor = []
    for _ in range(300):
        if len(decor) >= 26:
            break
        px, py = rng.uniform(bbox[0] - m, bbox[1] + m), rng.uniform(bbox[2] - m, bbox[3] + m)
        r = rng.uniform(1.3, 2.2) * max(1.0, fs * 0.6)
        if _dist_polilinha((px, py), pts) < 4.0 + 0.03 * ext + r:
            continue
        ah, an = (px - E[0]) * h[0] + (py - E[1]) * h[1], (px - E[0]) * n[0] + (py - E[1]) * n[1]
        if -8 - r < ah < ho + Lc + 8 + r and abs(an) < Wp / 2 + 8 + r:
            continue
        if abs(px - casa['cx']) < b_h / 2 + 3 + r and abs(py - casa['cy']) < a_h / 2 + 3 + r:
            continue
        if math.hypot(px - C[0], py - C[1]) < rc + 3 + r or math.hypot(px, py) < 4 + r:
            continue
        decor.append((px, py, r))

    # z e distância acumulada em cada vértice do trajeto
    dist_acum, acc = [0.0], 0.0
    for c in comprimentos:
        acc += c
        dist_acum.append(acc)
    return dict(pts=pts, heads=heads, E=E, h=h, n=n, z=z, s=s, u=u, rc=rc, hc=hc, C=C, casa=casa, Lc=Lc, ho=ho, Wp=Wp, t=t,
                esp_l=esp_l, esp_e=esp_e, n_linhas=n_linhas, n_setores_pedidos=n_setores, setores=setores, n_em=n_em,
                ext=ext, fs=fs, bbox=bbox, decor=decor, comprimentos=comprimentos, dist_acum=dist_acum,
                z_pts=[z(x, y) for x, y in pts], dz=float(dz), aclive=bool(aclive))

def campo_xy(L, ah, an):
    """Coordenadas locais do campo (ah ao longo das laterais, an ao longo do cabeçote) -> mundo."""
    return (L['E'][0] + L['h'][0] * ah + L['n'][0] * an, L['E'][1] + L['h'][1] * ah + L['n'][1] * an)

# ------------------------------------------------------------------------------
# Construtor de malhas 3D (cilindros, esferas, caixas, telhado) — um Mesh3d por categoria
# ------------------------------------------------------------------------------
class _Malha:
    def __init__(self):
        self.v, self.f, self.o = [], [], 0

    def _add(self, V, F):
        V = np.asarray(V, float)
        self.v.append(V)
        self.f.append(np.asarray(F, int) + self.o)
        self.o += len(V)

    def cilindro(self, p0, p1, r, r1=None, seg=14, tampas=True):
        p0, p1 = np.array(p0, float), np.array(p1, float)
        d = p1 - p0
        L = np.linalg.norm(d)
        if L < 1e-9:
            return
        d /= L
        ref = np.array([0.0, 0.0, 1.0]) if abs(d[2]) < 0.9 else np.array([1.0, 0.0, 0.0])
        e1 = np.cross(d, ref)
        e1 /= np.linalg.norm(e1)
        e2 = np.cross(d, e1)
        ang = np.linspace(0, 2 * np.pi, seg, endpoint=False)
        anel = np.outer(np.cos(ang), e1) + np.outer(np.sin(ang), e2)
        V = np.vstack([p0 + r * anel, p1 + (r if r1 is None else r1) * anel])
        F = []
        for i in range(seg):
            j = (i + 1) % seg
            F += [[i, j, seg + i], [j, seg + j, seg + i]]
        if tampas:
            V = np.vstack([V, p0, p1])
            c0, c1 = 2 * seg, 2 * seg + 1
            for i in range(seg):
                j = (i + 1) % seg
                F += [[c0, j, i], [c1, seg + i, seg + j]]
        self._add(V, F)

    def esfera(self, c, rx, ry=None, rz=None, nl=7, nm=12):
        ry = rx if ry is None else ry
        rz = rx if rz is None else rz
        V, F = [], []
        for i in range(nl + 1):
            th = np.pi * i / nl
            for j in range(nm):
                ph = 2 * np.pi * j / nm
                V.append([c[0] + rx * np.sin(th) * np.cos(ph), c[1] + ry * np.sin(th) * np.sin(ph), c[2] + rz * np.cos(th)])
        for i in range(nl):
            for j in range(nm):
                a, b = i * nm + j, i * nm + (j + 1) % nm
                c2, d = (i + 1) * nm + j, (i + 1) * nm + (j + 1) % nm
                F += [[a, b, c2], [b, d, c2]]
        self._add(V, F)

    def _face(self, pts):
        k = len(pts)
        base = [list(p) for p in pts]
        F = [[0, i, i + 1] for i in range(1, k - 1)]
        self._add(base, F)

    def caixa(self, c, dx, dy, dz):
        x, y, z = c
        hx, hy, hz = dx / 2, dy / 2, dz / 2
        p = lambda sx, sy, sz: (x + sx * hx, y + sy * hy, z + sz * hz)
        for quad in (
            [p(-1, -1, -1), p(1, -1, -1), p(1, 1, -1), p(-1, 1, -1)], [p(-1, -1, 1), p(1, -1, 1), p(1, 1, 1), p(-1, 1, 1)],
            [p(-1, -1, -1), p(1, -1, -1), p(1, -1, 1), p(-1, -1, 1)], [p(-1, 1, -1), p(1, 1, -1), p(1, 1, 1), p(-1, 1, 1)],
            [p(-1, -1, -1), p(-1, 1, -1), p(-1, 1, 1), p(-1, -1, 1)], [p(1, -1, -1), p(1, 1, -1), p(1, 1, 1), p(1, -1, 1)]):
            self._face(quad)

    def telhado(self, cx, cy, z0, b, a, h, bei=0.5):
        x0, x1, y0, y1 = cx - b / 2 - bei, cx + b / 2 + bei, cy - a / 2 - bei, cy + a / 2 + bei
        zb, zt = z0 - 0.25, z0 + h
        self._face([(x0, y0, zb), (x1, y0, zb), (x1, cy, zt), (x0, cy, zt)])
        self._face([(x0, y1, zb), (x1, y1, zb), (x1, cy, zt), (x0, cy, zt)])
        self._face([(x0, y0, zb), (x0, y1, zb), (x0, cy, zt)])
        self._face([(x1, y0, zb), (x1, y1, zb), (x1, cy, zt)])

    def faixa(self, topo, zbase):
        """Faixa vertical (saia de terra) entre uma polilinha 3D e a cota zbase."""
        for (x0, y0, z0), (x1, y1, z1) in zip(topo[:-1], topo[1:]):
            self._face([(x0, y0, z0), (x1, y1, z1), (x1, y1, zbase), (x0, y0, zbase)])

    def trace(self, cor, nome="", opacidade=1.0, legenda=False):
        if not self.v:
            return None
        V, F = np.vstack(self.v), np.vstack(self.f)
        return go.Mesh3d(x=V[:, 0], y=V[:, 1], z=V[:, 2], i=F[:, 0], j=F[:, 1], k=F[:, 2], color=cor, opacity=opacidade,
                         name=nome, showlegend=bool(legenda and nome), hoverinfo="name" if nome else "skip",
                         flatshading=False,
                         lighting=dict(ambient=0.55, diffuse=0.85, specular=0.35, roughness=0.45, fresnel=0.15),
                         lightposition=dict(x=-200, y=-300, z=500))

CAMERAS_3D = {
    "Aérea (padrão)": dict(eye=dict(x=-0.95, y=-1.05, z=0.68)),
    "Topo (planta)": dict(eye=dict(x=0.0, y=0.0, z=2.6), up=dict(x=0, y=1, z=0)),
    "Lateral (perfil)": dict(eye=dict(x=0.0, y=-2.2, z=0.35)),
    "Ao nível do solo": dict(eye=dict(x=-0.75, y=-0.55, z=0.12)),
    "Foco na roça (setores)": None,
    "Foco na casa e bomba": None,
}

def gerar_cena_3d(L, cfg):
    """Cena 3D realista. cfg: metodo, cultura, tipo_solo, material, dn_mm, setor_ativo (0 = todos), camera, vol_cisterna,
    nivel_cisterna (0-1), pot_cv, area_plantio, area_telhado, exagero, arvores, barra_m."""
    z, ext = L['z'], L['ext']
    fs = max(L['fs'] * 1.35, 1.4)
    mat = cfg.get("material", "PVC")
    cor_t, cor_te, cor_tl = COR_TUBO.get(mat, COR_TUBO["PVC"])
    rv = max(cfg.get("dn_mm", 50.0) / 2000.0, ext / 300.0) * cfg.get("exagero", 2.0)
    solo = COR_SOLO.get(cfg.get("tipo_solo"), "#7B5236")
    metodo = cfg.get("metodo", "Gotejamento")
    arborea = cfg.get("arborea", False)
    ativo = int(cfg.get("setor_ativo", 0))
    zo = ext / 1200.0
    E, hh, nn = L['E'], L['h'], L['n']
    ho, Lc, Wp = L['ho'], L['Lc'], L['Wp']
    malhas = {}
    M = lambda nome: malhas.setdefault(nome, _Malha())
    def cp(ah, an, dz_=0.0):
        x, y = campo_xy(L, ah, an)
        return (x, y, z(x, y) + dz_)

    # ---------------- Terreno ----------------
    xmin, xmax, ymin, ymax = L['bbox']
    marg = 0.09 * ext
    xs, ys = np.linspace(xmin - marg, xmax + marg, 72), np.linspace(ymin - marg, ymax + marg, 72)
    X, Y = np.meshgrid(xs, ys)
    Z = z(X, Y)
    padrao = 0.5 + 0.5 * np.sin(X / (ext / 22.0)) * np.cos(Y / (ext / 27.0))
    terreno = go.Surface(x=X, y=Y, z=Z, surfacecolor=padrao, cmin=0, cmax=1, showscale=False,
                         colorscale=[[0, "#5C9C45"], [1, "#8CC863"]], hoverinfo="skip", name="Terreno",
                         lighting=dict(ambient=0.75, diffuse=0.6, specular=0.05, roughness=0.9))
    zbase = float(Z.min()) - 0.05 * ext
    bordas = [[(xs[i], ys[0], Z[0, i]) for i in range(len(xs))], [(xs[i], ys[-1], Z[-1, i]) for i in range(len(xs))],
              [(xs[0], ys[j], Z[j, 0]) for j in range(len(ys))], [(xs[-1], ys[j], Z[j, -1]) for j in range(len(ys))]]
    for b in bordas:
        M("saia").faixa(b, zbase)

    # ---------------- Tubulação principal (trajeto com curvas) ----------------
    P = [np.array([x, y, z(x, y) + rv]) for x, y in L['pts']]
    x_bomba = -0.45 * fs
    Pd = np.array([x_bomba, 0.0, z(x_bomba, 0.0) + rv])
    M("tubo").cilindro(Pd, P[0], rv)
    for i in range(len(P) - 1):
        M("tubo").cilindro(P[i], P[i + 1], rv)
    for i in range(1, len(P) - 1):                       # cotovelos
        M("conex").esfera(P[i], rv * 1.38)
    barra = cfg.get("barra_m", 6.0)
    for i in range(len(P) - 1):                          # luvas de união a cada barra
        d = P[i + 1] - P[i]
        l = np.linalg.norm(d)
        if l < 1e-6:
            continue
        d = d / l
        k = 1
        while k * barra < l - 0.3:
            M("conex").cilindro(P[i] + d * (k * barra - 0.5 * rv), P[i] + d * (k * barra + 0.5 * rv), rv * 1.25)
            k += 1
    seg0 = P[1] - P[0]
    l0 = np.linalg.norm(seg0)
    d0 = seg0 / l0
    pf = P[0] + d0 * min(max(3.0 * fs, 5 * rv), 0.30 * l0)       # filtro em Y
    pv = P[0] + d0 * min(max(6.0 * fs, 10 * rv), 0.60 * l0)      # registro geral
    M("filtro").cilindro(pf - d0 * 2.2 * rv, pf + d0 * 2.2 * rv, rv * 1.7)
    ramo = d0 * 0.5 + np.array([0, 0, 0.87])
    ramo = ramo / np.linalg.norm(ramo)
    M("filtro").cilindro(pf, pf + ramo * 5.5 * rv, rv * 1.35)
    M("conex").esfera(pf + ramo * 5.5 * rv, rv * 1.5)
    M("valv_geral").esfera(pv, rv * 1.7)
    M("valv_geral").cilindro(pv, pv + np.array([0, 0, 3.4 * rv]), rv * 0.32)
    perp = np.array([-d0[1], d0[0], 0.0])
    M("valv_geral").cilindro(pv + np.array([0, 0, 3.4 * rv]) - perp * 1.8 * rv, pv + np.array([0, 0, 3.4 * rv]) + perp * 1.8 * rv, rv * 0.28)

    # ---------------- Conjunto motobomba ----------------
    z0 = z(0.0, 0.0)
    zb = z0 + 0.14 * fs
    M("bomba_base").caixa((x_bomba, 0.65 * fs, z0 + 0.07 * fs), 1.0 * fs, 1.9 * fs, 0.14 * fs)
    M("bomba_carc").cilindro((x_bomba, 0, zb), (x_bomba, 0, zb + 0.62 * fs), 0.30 * fs, seg=20)
    M("bomba_motor").cilindro((x_bomba, 0.25 * fs, zb + 0.36 * fs), (x_bomba, 1.5 * fs, zb + 0.36 * fs), 0.30 * fs, seg=20)
    for yy in np.linspace(0.45, 1.35, 6):
        M("bomba_aleta").cilindro((x_bomba, yy * fs, zb + 0.36 * fs), (x_bomba, (yy + 0.06) * fs, zb + 0.36 * fs), 0.335 * fs, seg=20)
    zc = z(*L['C'])
    rc, hc = L['rc'], L['hc']
    Ps = np.array([L['C'][0] + rc, 0.0, z(L['C'][0] + rc, 0.0) + 0.45])
    Pe = np.array([x_bomba - 0.30 * fs, 0.0, z(x_bomba - 0.30 * fs, 0.0) + rv * 1.3])
    M("tubo").cilindro(Ps, Pe, rv * 1.25)
    M("conex").cilindro(Pe + np.array([0.0, 0, 0]), Pe + np.array([0.4 * rv, 0, 0]), rv * 1.7)

    # ---------------- Cisterna ----------------
    C0, C1 = L['C']
    nivel = min(1.0, max(0.05, cfg.get("nivel_cisterna", 0.6)))
    M("cisterna").cilindro((C0, C1, zc), (C0, C1, zc + hc), rc, seg=32)
    M("agua").cilindro((C0, C1, zc + 0.03), (C0, C1, zc + 0.03 + (hc - 0.1) * nivel), rc * 0.97, seg=32)
    M("cist_tampa").caixa((C0, C1, zc + hc + 0.06), 0.9, 0.9, 0.12)

    # ---------------- Casa, calha e condutor ----------------
    cs = L['casa']
    cantos = [(cs['cx'] + sx * cs['b'] / 2, cs['cy'] + sy * cs['a'] / 2) for sx in (-1, 1) for sy in (-1, 1)]
    zmin_c, zh = min(z(*c) for c in cantos), max(z(*c) for c in cantos)
    M("casa_base").caixa((cs['cx'], cs['cy'], (zmin_c + zh) / 2 - 0.1), cs['b'] + 0.3, cs['a'] + 0.3, zh - zmin_c + 0.4)
    M("casa").caixa((cs['cx'], cs['cy'], zh + 1.75), cs['b'], cs['a'], 3.5)
    M("telhado").telhado(cs['cx'], cs['cy'], zh + 3.5, cs['b'], cs['a'], 1.9)
    M("porta").caixa((cs['frente'] + 0.03, cs['cy'], zh + 1.05), 0.1, 1.0, 2.1)
    for sy in (-0.3, 0.3):
        M("janela").caixa((cs['frente'] + 0.03, cs['cy'] + sy * cs['a'], zh + 1.9), 0.1, 1.4, 1.1)
    xg = cs['frente'] + 0.35
    M("calha").cilindro((xg, -cs['a'] / 2, zh + 3.4), (xg, cs['a'] / 2, zh + 3.4), 0.13, seg=10)
    M("condutor").cilindro((xg, 0, zh + 3.4), (xg, 0, zh + 0.45), 0.09, seg=10)
    M("condutor").cilindro((xg, 0, zh + 0.45), (C0 - rc, 0, zc + 0.45), 0.09, seg=10)
    M("condutor").cilindro((xg + 1.0, 0, zh + 0.45), (xg + 1.0, 0, zh - 0.05), 0.07, seg=10)  # descarte da 1ª água
    M("rede").cilindro((C0, -rc - 5.0, zc + hc * 0.85), (C0, -rc * 0.97, zc + hc * 0.85), 0.07, seg=10)
    M("rede").esfera((C0, -rc - 2.5, zc + hc * 0.85), 0.2)

    # ---------------- Campo: cabeçote, setores, válvulas, laterais ----------------
    t, setores = L['t'], L['setores']
    M("tubo").cilindro(cp(0, -Wp / 2 + 0.2, rv), cp(0, Wp / 2 - 0.2, rv), rv * 0.9)
    M("conex").esfera(cp(0, 0, rv), rv * 1.45)
    traces = [terreno]
    hastes_x, hastes_y, hastes_z = [], [], []
    fileiras = dict(x=[], y=[], z=[])
    mangueiras = dict(x=[], y=[], z=[])
    emis = dict(x=[], y=[], z=[])
    gotas = dict(x=[], y=[], z=[])
    plantas = dict(x=[], y=[], z=[])
    rotulos = dict(x=[], y=[], z=[], t=[])
    for st_ in setores:
        k, tc = st_['k'], st_['tc']
        ligado = (ativo == 0 or ativo == k)
        M("tubo").cilindro(cp(0, tc, rv), cp(ho, tc, rv), rv * 0.8)
        pvs = cp(ho * 0.72, tc, rv)
        nome_v = "valv_aberta" if ligado else "valv_fechada"
        M(nome_v).esfera(pvs, rv * 1.6)
        M(nome_v).cilindro(pvs, (pvs[0], pvs[1], pvs[2] + 3.2 * rv), rv * 0.3)
        M(nome_v).cilindro((pvs[0] - nn[0] * 1.7 * rv, pvs[1] - nn[1] * 1.7 * rv, pvs[2] + 3.2 * rv),
                           (pvs[0] + nn[0] * 1.7 * rv, pvs[1] + nn[1] * 1.7 * rv, pvs[2] + 3.2 * rv), rv * 0.27)
        ta, tb = t[st_['linhas'][0]] - 0.4, t[st_['linhas'][-1]] + 0.4
        M("tubo").cilindro(cp(ho, ta, rv), cp(ho, tb, rv), rv * 0.7)
        # tapete do setor (solo + tom do setor; setor irrigando fica úmido)
        base = _misturar(solo, PALETA_SETORES[(k - 1) % 8], 0.55)
        if ligado:
            base = _misturar(base, "#1B4F72", 0.12)
        q = [cp(ho - 0.5, st_['t0'], zo), cp(ho - 0.5, st_['t1'], zo), cp(ho + Lc + 0.3, st_['t1'], zo), cp(ho + Lc + 0.3, st_['t0'], zo)]
        traces.append(go.Mesh3d(x=[p[0] for p in q], y=[p[1] for p in q], z=[p[2] for p in q], i=[0, 0], j=[1, 2], k=[2, 3],
                                color=base, name=f"Setor {k} ({len(st_['linhas'])} linha(s))", showlegend=True, hoverinfo="name",
                                lighting=dict(ambient=0.95, diffuse=0.3, specular=0.0)))
        contorno = q + [q[0]]
        traces.append(go.Scatter3d(x=[p[0] for p in contorno], y=[p[1] for p in contorno], z=[p[2] + zo for p in contorno],
                                   mode="lines", line=dict(color="white", width=4), showlegend=False, hoverinfo="skip"))
        pl = cp(ho + Lc / 2, tc, max(3.5, 0.08 * ext) + (k % 2) * 0.03 * ext)
        rotulos['x'].append(pl[0]); rotulos['y'].append(pl[1]); rotulos['z'].append(pl[2])
        rotulos['t'].append(f"<b>S{k}</b> · {len(st_['linhas'])} lin.<br>{'irrigando' if ligado else 'fechado'}")
        for j in st_['linhas']:
            a, b = cp(ho, t[j], 2 * zo), cp(ho + Lc, t[j], 2 * zo)
            for dic, off in ((fileiras, 0.0), (mangueiras, zo)):
                dic['x'] += [a[0], b[0], None]; dic['y'] += [a[1], b[1], None]; dic['z'] += [a[2] + off, b[2] + off, None]
            for e in range(L['n_em']):
                ah = ho + (e + 0.5) * L['esp_e']
                p = cp(ah, t[j], 3 * zo)
                emis['x'].append(p[0]); emis['y'].append(p[1]); emis['z'].append(p[2])
                if ligado and metodo == "Gotejamento":
                    gotas['x'].append(p[0]); gotas['y'].append(p[1]); gotas['z'].append(p[2] + zo)
                plantas['x'].append(p[0]); plantas['y'].append(p[1]); plantas['z'].append(p[2] + 0.12)
    # decimação para manter a cena leve
    def amostra(dic, maxn):
        n_ = len(dic['x'])
        if n_ <= maxn:
            return dic
        passo = int(math.ceil(n_ / maxn))
        return {k_: v[::passo] for k_, v in dic.items()}
    emis, gotas, plantas = amostra(emis, 2500), amostra(gotas, 2500), amostra(plantas, 1800)

    if arborea:
        sel = amostra(plantas, 60)
        tf = max(1.0, 0.6 * fs)
        for px, py, pz in zip(sel['x'], sel['y'], sel['z']):
            M("tronco").cilindro((px, py, pz), (px, py, pz + 1.3 * tf), 0.09 * tf, seg=8)
            M("copa_cult").esfera((px, py, pz + 2.0 * tf), 0.95 * tf, 0.95 * tf, 0.85 * tf, nl=6, nm=10)
    else:
        sel = amostra(plantas, 1200)
        traces.append(go.Scatter3d(x=sel['x'], y=sel['y'], z=sel['z'], mode="markers", name=f"Plantas ({cfg.get('cultura', '')})",
                                   marker=dict(size=4.5, color="#2F9E44", opacity=0.95), hoverinfo="skip", showlegend=True))
    traces.append(go.Scatter3d(x=fileiras['x'], y=fileiras['y'], z=fileiras['z'], mode="lines", name="Fileiras da cultura",
                               line=dict(color="#2E8B3D", width=9), hoverinfo="skip", showlegend=True))
    traces.append(go.Scatter3d(x=mangueiras['x'], y=mangueiras['y'], z=mangueiras['z'], mode="lines", name="Linhas laterais (PE)",
                               line=dict(color="#111111", width=3), hoverinfo="skip", showlegend=True))
    if metodo == "Gotejamento":
        traces.append(go.Scatter3d(x=emis['x'], y=emis['y'], z=emis['z'], mode="markers", name="Gotejadores",
                                   marker=dict(size=2.6, color="#F1C40F"), hoverinfo="skip", showlegend=True))
        if gotas['x']:
            traces.append(go.Scatter3d(x=gotas['x'], y=gotas['y'], z=gotas['z'], mode="markers", name="Água aplicada",
                                       marker=dict(size=3.4, color="#3498DB", opacity=0.8), hoverinfo="skip", showlegend=True))
    else:
        asp = amostra(emis, 48)
        hr = (1.2 if metodo == "Aspersão" else 0.45) * min(fs, 2.0)
        raio = min(0.45 * max(L['esp_e'], L['esp_l']), 8.0) if metodo == "Aspersão" else 1.6
        for px, py, pz in zip(asp['x'], asp['y'], asp['z']):
            M("aspersor").cilindro((px, py, pz), (px, py, pz + hr), 0.05 * min(fs, 2.0) + 0.02, seg=8)
            M("aspersor").esfera((px, py, pz + hr), 0.10 * min(fs, 2.0) + 0.04, nl=4, nm=8)
            ligado_local = True
            if ativo != 0:  # o jato só aparece no setor em operação
                ligado_local = False
                for st_ in setores:
                    if st_['k'] == ativo:
                        ah = (px - E[0]) * hh[0] + (py - E[1]) * hh[1]
                        an = (px - E[0]) * nn[0] + (py - E[1]) * nn[1]
                        ligado_local = st_['t0'] <= an <= st_['t1']
            if ligado_local:
                M("jato").cilindro((px, py, pz + 0.25), (px, py, pz + hr), raio, r1=0.05, seg=16, tampas=False)

    # ---------------- Árvores decorativas ----------------
    if cfg.get("arvores", True):
        for i, (px, py, r) in enumerate(L['decor']):
            pz = z(px, py)
            M("tronco").cilindro((px, py, pz), (px, py, pz + r * 1.15), r * 0.12, seg=8)
            M(f"copa{i % 3}").esfera((px, py, pz + r * 1.9), r, r, r * 0.9, nl=6, nm=10)

    # ---------------- Rótulos ----------------
    def rot(x, y, zz, txt):
        rotulos['x'].append(x); rotulos['y'].append(y); rotulos['z'].append(zz); rotulos['t'].append(txt)
    rot(cs['cx'], cs['cy'], zh + 7.5, f"<b>CASA</b><br>telhado {cfg.get('area_telhado', 0):,.0f} m²")
    rot(C0, C1 + rc + 4.0, zc + hc + 1.0, f"<b>CISTERNA</b><br>{cfg.get('vol_cisterna', 0):,.0f} L · {nivel * 100:.0f}%")
    rot(x_bomba, -3.2 * fs, z0 + 1.2 * fs, f"<b>BOMBA</b> {cfg.get('pot_cv', 0):.1f} CV")
    rot(pf[0] + perp[0] * 6 * rv, pf[1] + perp[1] * 6 * rv, pf[2] + 8 * rv, "<b>FILTRO</b>")
    rot(pv[0] - perp[0] * 6 * rv, pv[1] - perp[1] * 6 * rv, pv[2] + 11 * rv, "<b>REGISTRO</b>")
    i_long = int(np.argmax([np.linalg.norm(P[i + 1] - P[i]) for i in range(len(P) - 1)]))
    mp = (P[i_long] + P[i_long + 1]) / 2
    rot(mp[0], mp[1], mp[2] + 7 * rv, f"<b>ADUTORA {mat}</b><br>DN {cfg.get('dn_mm', 50):.0f} mm · {sum(L['comprimentos']):.0f} m")
    for i in range(1, len(P) - 1):
        rot(P[i][0], P[i][1], P[i][2] + 6 * rv, f"curva {i}")
    trace_rot = go.Scatter3d(x=rotulos['x'], y=rotulos['y'], z=rotulos['z'], text=rotulos['t'], mode="text",
                             textfont=dict(size=11, color="#1B2631"), showlegend=False, hoverinfo="skip")

    # ---------------- Malhas -> traces ----------------
    meta = {
        "saia": (solo, "", 1, False), "tubo": (cor_t, f"Tubo {mat} DN {cfg.get('dn_mm', 50):.0f} mm", 1, True),
        "conex": (cor_te, "Cotovelos, luvas e tês", 1, True), "filtro": ("#2C3E50", "Filtro", 1, True),
        "valv_geral": ("#E67E22", "Registro geral", 1, True), "valv_aberta": ("#27AE60", "Válvula de setor (aberta)", 1, True),
        "valv_fechada": ("#C0392B", "Válvula de setor (fechada)", 1, True), "bomba_base": ("#7F8C8D", "", 1, False),
        "bomba_carc": ("#5D6D7E", "Conjunto motobomba", 1, True), "bomba_motor": ("#1F618D", "", 1, False),
        "bomba_aleta": ("#2E86C1", "", 1, False), "cisterna": ("#D6EAF8", "Cisterna", 0.28, True),
        "agua": ("#2E86DE", "Água armazenada", 0.9, True), "cist_tampa": ("#7F8C8D", "", 1, False),
        "casa_base": ("#95A5A6", "", 1, False), "casa": ("#F5E6CA", "Casa", 1, True), "telhado": ("#B5503C", "Telhado", 1, False),
        "porta": ("#6E4B2A", "", 1, False), "janela": ("#85C1E9", "", 1, False), "calha": ("#7F8C8D", "", 1, False),
        "condutor": ("#ECF0F1", "Calha e condutor", 1, True), "rede": ("#5D8AA8", "Rede pública (reserva)", 1, True),
        "tronco": ("#6E4B2A", "", 1, False), "copa0": ("#3D7A2F", "", 1, False), "copa1": ("#4E8F3A", "", 1, False),
        "copa2": ("#2F6B2A", "", 1, False), "copa_cult": ("#2E9E3E", f"Cultura: {cfg.get('cultura', '')}", 1, True),
        "aspersor": ("#566573", "Aspersor / microaspersor", 1, True), "jato": ("#5DADE2", "Jato d'água", 0.35, True),
    }
    for nome, malha in malhas.items():
        cor, leg, op, mostrar = meta.get(nome, ("#999999", "", 1, False))
        tr = malha.trace(cor, leg, op, mostrar)
        if tr is not None:
            traces.append(tr)

    traces.append(trace_rot)   # rótulos por último (ficam sobre as malhas)
    fig = go.Figure(data=traces)
    # limites reais dos dados -> coordenadas normalizadas da câmera
    mn, mx = np.full(3, np.inf), np.full(3, -np.inf)
    for tr in fig.data:
        for eixo, nome_ in enumerate("xyz"):
            v = getattr(tr, nome_, None)
            if v is None:
                continue
            a_ = np.asarray([np.nan if q is None else q for q in np.ravel(v)], float)
            if a_.size and not np.all(np.isnan(a_)):
                mn[eixo], mx[eixo] = min(mn[eixo], np.nanmin(a_)), max(mx[eixo], np.nanmax(a_))
    mid, maxr = (mn + mx) / 2, float(max(mx - mn))
    def foco(p, off):
        c = [2.0 * (p[i] - mid[i]) / maxr for i in range(3)]
        return dict(center=dict(x=c[0], y=c[1], z=c[2] * 0.0 + c[2]), eye=dict(x=c[0] + off[0], y=c[1] + off[1], z=c[2] + off[2]))
    escolha = cfg.get("camera", "Aérea (padrão)")
    if escolha == "Foco na roça (setores)":
        pc = cp(ho + Lc / 2, 0.0, 0.0)
        cam = foco(pc, (-hh[0] * 0.62 - nn[0] * 0.42, -hh[1] * 0.62 - nn[1] * 0.42, 0.50))
    elif escolha == "Foco na casa e bomba":
        cam = foco((L['casa']['frente'] * 0.5, 0.0, zh), (-0.40, -0.62, 0.38))
    else:
        cam = dict(CAMERAS_3D.get(escolha, CAMERAS_3D["Aérea (padrão)"]))
    fig.update_layout(
        scene=dict(aspectmode="data", bgcolor="#D6EAF8", camera=cam,
                   xaxis=dict(visible=False), yaxis=dict(visible=False), zaxis=dict(visible=False)),
        paper_bgcolor="#D6EAF8", margin=dict(l=0, r=0, t=0, b=0), height=700,
        legend=dict(x=0.995, y=0.02, xanchor="right", yanchor="bottom", bgcolor="rgba(255,255,255,0.82)", font=dict(size=10)),
    )
    return fig

# ==============================================================================
# ESQUEMÁTICOS 2D REALISTAS (SVG) — usam a mesma geometria do 3D (calcular_layout_campo)
# ==============================================================================
def _svg_tubo(pts, w, estilo=None):
    """Tubo com aparência cilíndrica: contorno escuro, corpo e reflexo de luz."""
    c, ce, cl = estilo or COR_TUBO["PVC"]
    sp = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    base = 'fill="none" stroke-linecap="round" stroke-linejoin="round"'
    return (f'<polyline points="{sp}" {base} stroke="{ce}" stroke-width="{w:.1f}"/>'
            f'<polyline points="{sp}" {base} stroke="{c}" stroke-width="{w * 0.74:.1f}"/>'
            f'<polyline points="{sp}" {base} stroke="{cl}" stroke-width="{max(1.0, w * 0.16):.1f}" opacity="0.75" '
            f'transform="translate({-w * 0.13:.1f},{-w * 0.13:.1f})"/>')

def _fmt_br(v, casas=1):
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def _nice_scale(k, alvo_px=110):
    for m_ in (1, 2, 5, 10, 20, 25, 50, 100, 200, 250, 500):
        if m_ * k >= alvo_px * 0.6:
            return m_
    return 500

def _svg_nuvem(cx, cy, s=1.0, cor="#FFFFFF", op=0.95):
    return (f'<g opacity="{op}" fill="{cor}"><ellipse cx="{cx}" cy="{cy}" rx="{34 * s:.0f}" ry="{13 * s:.0f}"/>'
            f'<ellipse cx="{cx - 20 * s:.0f}" cy="{cy + 3 * s:.0f}" rx="{22 * s:.0f}" ry="{11 * s:.0f}"/>'
            f'<ellipse cx="{cx + 22 * s:.0f}" cy="{cy + 3 * s:.0f}" rx="{24 * s:.0f}" ry="{11 * s:.0f}"/></g>')

def _svg_sol(cx, cy, r=20):
    raios = "".join(f'<line x1="{cx + math.cos(a) * (r + 5):.1f}" y1="{cy + math.sin(a) * (r + 5):.1f}" '
                    f'x2="{cx + math.cos(a) * (r + 13):.1f}" y2="{cy + math.sin(a) * (r + 13):.1f}" stroke="#F4B400" stroke-width="3" stroke-linecap="round"/>'
                    for a in np.linspace(0, 2 * np.pi, 12, endpoint=False))
    return f'{raios}<circle cx="{cx}" cy="{cy}" r="{r}" fill="#FFD43B" stroke="#F4B400" stroke-width="2"/>'

def _svg_planta_icone(x, y, arborea, altura=16, cor="#2E9E3E"):
    if arborea:
        return (f'<rect x="{x - 1.5:.1f}" y="{y - altura * 0.55:.1f}" width="3" height="{altura * 0.55:.1f}" fill="#6E4B2A"/>'
                f'<circle cx="{x:.1f}" cy="{y - altura * 0.8:.1f}" r="{altura * 0.45:.1f}" fill="{cor}" stroke="#1E7B30" stroke-width="0.8"/>')
    return (f'<path d="M{x:.1f},{y:.1f} q-{altura * 0.35:.1f},-{altura * 0.5:.1f} -{altura * 0.1:.1f},-{altura:.1f} '
            f'q{altura * 0.12:.1f},{altura * 0.35:.1f} {altura * 0.1:.1f},{altura * 0.65:.1f} '
            f'q{altura * 0.18:.1f},-{altura * 0.5:.1f} {altura * 0.45:.1f},-{altura * 0.55:.1f} '
            f'q-{altura * 0.35:.1f},{altura * 0.2:.1f} -{altura * 0.45:.1f},{altura * 0.9:.1f}z" fill="{cor}" stroke="#1E7B30" stroke-width="0.7"/>')

# ------------------------------------------------------------------------------
# PLANTA BAIXA REALISTA (vista de cima): terreno, casa, cisterna, bomba, PVC, válvulas, setores e laterais
# ------------------------------------------------------------------------------
def gerar_svg_planta_realista(L, cfg, detalhe=True, uid="pl", W=940, H=640, foco="tudo"):
    mg, topo = 46, 34
    fs = max(L['fs'], 1.0)
    if foco == "roca":   # zoom na roça: setores, cabeçote, válvulas e laterais
        _c = [campo_xy(L, ah, an) for ah, an in ((-0.7 * L['Lc'], -L['Wp'] / 2 - 2), (-0.7 * L['Lc'], L['Wp'] / 2 + 2),
                                                 (L['ho'] + L['Lc'] + 1, L['Wp'] / 2 + 2), (L['ho'] + L['Lc'] + 1, -L['Wp'] / 2 - 2))]
        xmin, xmax = min(c[0] for c in _c), max(c[0] for c in _c)
        ymin, ymax = min(c[1] for c in _c), max(c[1] for c in _c)
    else:
        xmin, xmax, ymin, ymax = L['bbox']
    k = min((W - 2 * mg) / max(xmax - xmin, 1.0), (H - 2 * mg - topo) / max(ymax - ymin, 1.0))
    if foco != "roca":   # reserva espaço em pixels para casa, cisterna e bomba (ícones com tamanho mínimo)
        xmin = min(xmin, -0.45 * fs - 290.0 / k)
        k = min((W - 2 * mg) / max(xmax - xmin, 1.0), (H - 2 * mg - topo) / max(ymax - ymin, 1.0))
    ox = mg + ((W - 2 * mg) - (xmax - xmin) * k) / 2 - xmin * k
    oy = mg + topo + ((H - 2 * mg - topo) - (ymax - ymin) * k) / 2 + ymax * k
    Pp = lambda x, y: (ox + x * k, oy - y * k)
    Pc = lambda ah, an: Pp(*campo_xy(L, ah, an))
    mat = cfg.get("material", "PVC")
    est = COR_TUBO.get(mat, COR_TUBO["PVC"])
    solo = COR_SOLO.get(cfg.get("tipo_solo"), "#7B5236")
    metodo = cfg.get("metodo", "Gotejamento")
    arborea = cfg.get("arborea", False)
    ativo = int(cfg.get("setor_ativo", 0))
    ho, Lc, Wp, t, setores = L['ho'], L['Lc'], L['Wp'], L['t'], L['setores']
    o = []
    o.append(f'''<defs>
      <pattern id="gr{uid}" width="16" height="16" patternUnits="userSpaceOnUse"><rect width="16" height="16" fill="#8FCB6B"/>
        <path d="M2 13 l1 -5 M8 7 l1 -4 M12 14 l-1 -5 M5 4 l1 -3" stroke="#74B552" stroke-width="1.3" fill="none"/></pattern>
      <pattern id="tl{uid}" width="8" height="6" patternUnits="userSpaceOnUse"><rect width="8" height="6" fill="#C0563F"/>
        <path d="M0 5.5 H8 M4 0 V5.5" stroke="#9C402D" stroke-width="0.8" fill="none"/></pattern>
      <radialGradient id="tr{uid}" cx="38%" cy="32%"><stop offset="0" stop-color="#67B454"/><stop offset="1" stop-color="#2C6B28"/></radialGradient>
      <radialGradient id="ag{uid}" cx="38%" cy="32%"><stop offset="0" stop-color="#8CC4F5"/><stop offset="1" stop-color="#1F6FB0"/></radialGradient>
    </defs>''')
    o.append(f'<rect width="{W}" height="{H}" fill="url(#gr{uid})" rx="6"/>')
    for (x, y, r) in L['decor']:
        cx, cy = Pp(x, y)
        o.append(f'<circle cx="{cx + 3:.1f}" cy="{cy + 4:.1f}" r="{r * k:.1f}" fill="#000" opacity="0.18"/>'
                 f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r * k:.1f}" fill="url(#tr{uid})" stroke="#2C6B28" stroke-width="0.8"/>')

    o.append(f'<text x="{W / 2}" y="26" font-family="Arial" font-size="16" font-weight="bold" fill="#1B2631" text-anchor="middle" stroke="#fff" stroke-width="4" paint-order="stroke">'
             f'{"Detalhe da roça" if foco == "roca" else "Planta baixa"} — {metodo} · {len(setores)} setor(es) · {cfg.get("cultura", "")}</text>')
    # ---- campo / setores ----
    poly = lambda pts_: " ".join(f"{a:.1f},{b:.1f}" for a, b in pts_)
    campo_pts = [Pc(ho - 0.5, -Wp / 2), Pc(ho - 0.5, Wp / 2), Pc(ho + Lc + 0.3, Wp / 2), Pc(ho + Lc + 0.3, -Wp / 2)]
    o.append(f'<polygon points="{poly([(a + 2, b + 3) for a, b in campo_pts])}" fill="#000" opacity="0.18"/>')
    for st_ in setores:
        kk = st_['k']
        ligado = (ativo == 0 or ativo == kk)
        cor = _misturar(solo, PALETA_SETORES[(kk - 1) % 8], 0.55)
        if ligado:
            cor = _misturar(cor, "#1B4F72", 0.12)
        pts_s = [Pc(ho - 0.5, st_['t0']), Pc(ho - 0.5, st_['t1']), Pc(ho + Lc + 0.3, st_['t1']), Pc(ho + Lc + 0.3, st_['t0'])]
        o.append(f'<polygon points="{poly(pts_s)}" fill="{cor}" stroke="#FFFFFF" stroke-width="2.2" stroke-dasharray="7,4"/>')
    # fileiras, laterais, emissores
    n_emis = 0
    for st_ in setores:
        ligado = (ativo == 0 or ativo == st_['k'])
        for j in st_['linhas']:
            a, b = Pc(ho, t[j]), Pc(ho + Lc, t[j])
            wrow = max(2.0, min(L['esp_l'] * k * 0.5, 14.0))
            o.append(f'<line x1="{a[0]:.1f}" y1="{a[1]:.1f}" x2="{b[0]:.1f}" y2="{b[1]:.1f}" stroke="#2E8B3D" stroke-width="{wrow:.1f}" stroke-linecap="round" opacity="0.8"/>')
            o.append(f'<line x1="{a[0]:.1f}" y1="{a[1]:.1f}" x2="{b[0]:.1f}" y2="{b[1]:.1f}" stroke="#111" stroke-width="1.3"/>')
            if detalhe and metodo == "Gotejamento" and L['esp_e'] * k >= 3.0:
                for e in range(L['n_em']):
                    if n_emis > 2600:
                        break
                    p = Pc(ho + (e + 0.5) * L['esp_e'], t[j])
                    n_emis += 1
                    o.append(f'<circle cx="{p[0]:.1f}" cy="{p[1]:.1f}" r="1.5" fill="{"#2E86DE" if ligado else "#F1C40F"}"/>')
    if metodo != "Gotejamento":
        raio = min(0.45 * max(L['esp_e'], L['esp_l']), 8.0) if metodo == "Aspersão" else 1.6
        cont = 0
        for st_ in setores:
            ligado = (ativo == 0 or ativo == st_['k'])
            for j in st_['linhas']:
                for e in range(L['n_em']):
                    if cont > 80:
                        break
                    p = Pc(ho + (e + 0.5) * L['esp_e'], t[j])
                    cont += 1
                    if ligado:
                        o.append(f'<circle cx="{p[0]:.1f}" cy="{p[1]:.1f}" r="{raio * k:.1f}" fill="#5DADE2" opacity="0.22" stroke="#2E86DE" stroke-width="1" stroke-dasharray="4,3"/>')
                    o.append(f'<circle cx="{p[0]:.1f}" cy="{p[1]:.1f}" r="3.2" fill="#566573" stroke="#fff" stroke-width="1"/>')
    elif arborea:
        sel = [(j, e) for st_ in setores for j in st_['linhas'] for e in range(0, L['n_em'], max(1, int(round(2.5 / max(L['esp_e'], 0.1)))))]
        for (j, e) in sel[:150]:
            p = Pc(ho + (e + 0.5) * L['esp_e'], t[j])
            o.append(f'<circle cx="{p[0]:.1f}" cy="{p[1]:.1f}" r="{max(3.0, 1.3 * k):.1f}" fill="url(#tr{uid})" opacity="0.9"/>')

    # ---- cabeçote, sub-mains, válvulas ----
    wm = 7.0
    o.append(_svg_tubo([Pc(0, -Wp / 2 + 0.2), Pc(0, Wp / 2 - 0.2)], wm, est))
    for st_ in setores:
        kk = st_['k']
        ligado = (ativo == 0 or ativo == kk)
        o.append(_svg_tubo([Pc(0, st_['tc']), Pc(ho, st_['tc'])], 5.0, est))
        o.append(_svg_tubo([Pc(ho, t[st_['linhas'][0]] - 0.4), Pc(ho, t[st_['linhas'][-1]] + 0.4)], 5.0, est))
        vx, vy = Pc(ho * 0.72, st_['tc'])
        cv = "#27AE60" if ligado else "#C0392B"
        o.append(f'<rect x="{vx - 9:.1f}" y="{vy - 9:.1f}" width="18" height="18" rx="4" fill="{cv}" stroke="#fff" stroke-width="1.5"/>'
                 f'<text x="{vx:.1f}" y="{vy + 4:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#fff" text-anchor="middle">V{kk}</text>')
        lx, ly = Pc(ho + Lc / 2, st_['tc'])
        wpx = abs(st_['t1'] - st_['t0']) * k
        compacto = wpx < 100
        rw = max(34.0, min(96.0, wpx - 6))
        txt1 = f"S{kk}" if compacto else f"SETOR {kk}"
        txt2 = f"{len(st_['linhas'])} lin." if compacto else f"{len(st_['linhas'])} linha(s) · {'irrigando' if ligado else 'fechado'}"
        o.append(f'<rect x="{lx - rw / 2:.1f}" y="{ly - 15:.1f}" width="{rw:.1f}" height="30" rx="6" fill="#fff" opacity="0.85"/>'
                 f'<text x="{lx:.1f}" y="{ly - 2:.1f}" font-family="Arial" font-size="12" font-weight="bold" fill="#1B2631" text-anchor="middle">{txt1}</text>'
                 f'<text x="{lx:.1f}" y="{ly + 11:.1f}" font-family="Arial" font-size="{9 if compacto else 10}" fill="#34495E" text-anchor="middle">{txt2}</text>')

    # ---- casa e cisterna ----
    cs = L['casa']
    p_bomba = Pp(-0.45 * fs, 0.0)
    rcp = max(L['rc'] * k, 22.0)
    cc = (p_bomba[0] - (rcp + 54.0), p_bomba[1])
    casa_w, casa_h = max(cs['b'] * k, 120.0), max(cs['a'] * k, 76.0)
    casa_dir = cc[0] - rcp - max(4.0 * k, 40.0)
    c1, c2 = (casa_dir - casa_w, p_bomba[1] - casa_h / 2), (casa_dir, p_bomba[1] + casa_h / 2)
    o.append(f'<rect x="{c1[0] + 4:.1f}" y="{c1[1] + 5:.1f}" width="{c2[0] - c1[0]:.1f}" height="{c2[1] - c1[1]:.1f}" fill="#000" opacity="0.2"/>'
             f'<rect x="{c1[0]:.1f}" y="{c1[1]:.1f}" width="{c2[0] - c1[0]:.1f}" height="{c2[1] - c1[1]:.1f}" fill="url(#tl{uid})" stroke="#7B2D1E" stroke-width="2"/>'
             f'<line x1="{c1[0]:.1f}" y1="{(c1[1] + c2[1]) / 2:.1f}" x2="{c2[0]:.1f}" y2="{(c1[1] + c2[1]) / 2:.1f}" stroke="#7B2D1E" stroke-width="3"/>'
             f'<text x="{(c1[0] + c2[0]) / 2:.1f}" y="{(c1[1] + c2[1]) / 2 - 7:.1f}" font-family="Arial" font-size="12" font-weight="bold" fill="#fff" text-anchor="middle">CASA</text>'
             f'<text x="{(c1[0] + c2[0]) / 2:.1f}" y="{(c1[1] + c2[1]) / 2 + 17:.1f}" font-family="Arial" font-size="10" fill="#fff" text-anchor="middle">telhado {_fmt_br(cfg.get("area_telhado", 0), 0)} m²</text>')
    o.append(_svg_tubo([(casa_dir, p_bomba[1]), (cc[0] - rcp, p_bomba[1])], 4.0, ("#ECF0F1", "#7F8C8D", "#FFFFFF")))
    o.append(f'<circle cx="{cc[0] + 3:.1f}" cy="{cc[1] + 4:.1f}" r="{rcp:.1f}" fill="#000" opacity="0.2"/>'
             f'<circle cx="{cc[0]:.1f}" cy="{cc[1]:.1f}" r="{rcp:.1f}" fill="url(#ag{uid})" stroke="#566573" stroke-width="3"/>'
             f'<circle cx="{cc[0]:.1f}" cy="{cc[1]:.1f}" r="{rcp * 0.78:.1f}" fill="none" stroke="#fff" stroke-width="1" opacity="0.6"/>'
             f'<text x="{cc[0]:.1f}" y="{cc[1] - rcp - 7:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#1B2631" text-anchor="middle">CISTERNA</text>'
             f'<text x="{cc[0]:.1f}" y="{cc[1] + 4:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#fff" text-anchor="middle">{_fmt_br(cfg.get("vol_cisterna", 0), 0)} L</text>')

    # ---- tubulação principal em PVC ----
    pts_px = [Pp(x, y) for x, y in L['pts']]
    o.append(_svg_tubo([(cc[0] + rcp, p_bomba[1]), p_bomba], 5.0, est))
    o.append(_svg_tubo([p_bomba] + pts_px, 9.0, est))
    barra = cfg.get("barra_m", 6.0)
    for i in range(len(pts_px) - 1):
        (ax, ay), (bx, by) = pts_px[i], pts_px[i + 1]
        seg = math.hypot(bx - ax, by - ay)
        if seg < 1:
            continue
        ux, uy = (bx - ax) / seg, (by - ay) / seg
        if barra * k >= 9:
            d = barra * k
            while d < seg - 3:
                px_, py_ = ax + ux * d, ay + uy * d
                o.append(f'<line x1="{px_ - uy * 6:.1f}" y1="{py_ + ux * 6:.1f}" x2="{px_ + uy * 6:.1f}" y2="{py_ - ux * 6:.1f}" stroke="{est[1]}" stroke-width="2.6"/>')
                d += barra * k
        mx, my = (ax + bx) / 2 - uy * 16, (ay + by) / 2 + ux * 16
        o.append(f'<text x="{mx:.1f}" y="{my:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#1B2631" text-anchor="middle" '
                 f'stroke="#fff" stroke-width="3" paint-order="stroke">L{i + 1} = {_fmt_br(L["comprimentos"][i])} m</text>')
    for i in range(1, len(pts_px) - 1):
        o.append(f'<circle cx="{pts_px[i][0]:.1f}" cy="{pts_px[i][1]:.1f}" r="8" fill="{est[1]}" stroke="#fff" stroke-width="2"/>'
                 f'<text x="{pts_px[i][0]:.1f}" y="{pts_px[i][1] + 3.5:.1f}" font-family="Arial" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">{i}</text>')
    o.append(f'<circle cx="{pts_px[-1][0]:.1f}" cy="{pts_px[-1][1]:.1f}" r="9" fill="{est[1]}" stroke="#fff" stroke-width="2"/>')
    # filtro, registro e bomba
    (ax, ay), (bx, by) = pts_px[0], pts_px[1]
    seg0 = max(math.hypot(bx - ax, by - ay), 1.0)
    ux, uy = (bx - ax) / seg0, (by - ay) / seg0
    dist_f = min(max(3.0 * fs * k, 26.0), 0.30 * seg0)
    dist_v = min(max(6.0 * fs * k, 52.0), 0.60 * seg0)
    fx, fy = ax + ux * dist_f, ay + uy * dist_f
    vx, vy = ax + ux * dist_v, ay + uy * dist_v
    o.append(f'<circle cx="{fx:.1f}" cy="{fy:.1f}" r="9" fill="#2C3E50" stroke="#fff" stroke-width="2"/>'
             f'<text x="{fx + uy * 20:.1f}" y="{fy - ux * 20 + 3:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#1B2631" text-anchor="middle" stroke="#fff" stroke-width="3" paint-order="stroke">filtro</text>'
             f'<circle cx="{vx:.1f}" cy="{vy:.1f}" r="7" fill="#E67E22" stroke="#fff" stroke-width="2"/>'
             f'<text x="{vx - uy * 20:.1f}" y="{vy + ux * 20 + 3:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#1B2631" text-anchor="middle" stroke="#fff" stroke-width="3" paint-order="stroke">registro</text>')
    o.append(f'<rect x="{p_bomba[0] - 22:.1f}" y="{p_bomba[1] - 11:.1f}" width="30" height="22" rx="4" fill="#1F618D" stroke="#fff" stroke-width="2"/>'
             f'<circle cx="{p_bomba[0] + 6:.1f}" cy="{p_bomba[1]:.1f}" r="11" fill="#5D6D7E" stroke="#fff" stroke-width="2"/>'
             f'<text x="{p_bomba[0] - 4:.1f}" y="{p_bomba[1] + 26:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#1B2631" text-anchor="middle" stroke="#fff" stroke-width="3" paint-order="stroke">BOMBA {cfg.get("pot_cv", 0):.1f} CV</text>')

    # ---- rosa dos ventos, escala e legenda ----
    o.append(f'<g transform="translate({W - 52},{topo + 38})"><circle r="24" fill="#fff" opacity="0.85" stroke="#566573"/>'
             f'<polygon points="0,-19 6,4 0,0 -6,4" fill="#C0392B"/><polygon points="0,19 6,-4 0,0 -6,-4" fill="#7F8C8D"/>'
             f'<text y="-27" font-family="Arial" font-size="12" font-weight="bold" fill="#1B2631" text-anchor="middle">N</text></g>')
    esc = _nice_scale(k)
    o.append(f'<g transform="translate({W / 2 - 40:.0f},{H - 24})"><rect x="-6" y="-16" width="{esc * k + 56:.0f}" height="26" rx="4" fill="#fff" opacity="0.85"/>'
             f'<line x1="0" y1="0" x2="{esc * k:.1f}" y2="0" stroke="#1B2631" stroke-width="3"/>'
             f'<line x1="0" y1="-5" x2="0" y2="5" stroke="#1B2631" stroke-width="2"/><line x1="{esc * k:.1f}" y1="-5" x2="{esc * k:.1f}" y2="5" stroke="#1B2631" stroke-width="2"/>'
             f'<text x="{esc * k + 8:.1f}" y="4" font-family="Arial" font-size="11" fill="#1B2631">{esc} m</text></g>')
    leg = [(est[0], f"Tubo {mat} (adutora / cabeçote)", "tubo"), ("#111111", "Linha lateral (PE)", "lin"), ("#27AE60", "Válvula aberta", "v"),
           ("#C0392B", "Válvula fechada", "v"), ("#2E8B3D", "Fileira da cultura", "lin")]
    lx0, ly0 = 22, H - 20 - 17 * len(leg)
    o.append(f'<rect x="{lx0 - 8}" y="{ly0 - 14}" width="196" height="{17 * len(leg) + 14}" rx="6" fill="#fff" opacity="0.88"/>')
    for i, (c_, txt, tp) in enumerate(leg):
        yy = ly0 + i * 17
        if tp == "v":
            o.append(f'<rect x="{lx0}" y="{yy - 8}" width="14" height="12" rx="3" fill="{c_}"/>')
        else:
            o.append(f'<line x1="{lx0}" y1="{yy - 2}" x2="{lx0 + 14}" y2="{yy - 2}" stroke="{c_}" stroke-width="{6 if tp == "tubo" else 3}" stroke-linecap="round"/>')
        o.append(f'<text x="{lx0 + 22}" y="{yy + 2}" font-family="Arial" font-size="10.5" fill="#1B2631">{txt}</text>')
    corpo = "".join(o)
    return (f'<div style="display:flex;justify-content:center;margin:8px 0;"><svg width="100%" viewBox="0 0 {W} {H}" '
            f'style="max-width:{W}px;border:1px solid #cfd8dc;border-radius:8px;" xmlns="http://www.w3.org/2000/svg">{corpo}</svg></div>').replace("\n", "")

# ------------------------------------------------------------------------------
# PERFIL TOPOGRÁFICO REALISTA (vista lateral ao longo do trajeto)
# ------------------------------------------------------------------------------
def gerar_svg_perfil_realista(L, cfg, uid="pf", W=940, H=450):
    est = COR_TUBO.get(cfg.get("material", "PVC"), COR_TUBO["PVC"])
    arborea = cfg.get("arborea", False)
    metodo = cfg.get("metodo", "Gotejamento")
    dist, zs = L['dist_acum'], L['z_pts']
    ho, Lc = L['ho'], L['Lc']
    slope_last = L['s'] * (L['h'][0] * L['u'][0] + L['h'][1] * L['u'][1])
    s_end = dist[-1] + ho + Lc
    z_end = zs[-1] + slope_last * (ho + Lc)
    pad_px, mgl, mgr = 330.0, 24, 44
    kx = (W - mgl - mgr - pad_px) / max(s_end, 1.0)
    pad = pad_px / kx
    X = [-pad] + dist + [s_end]
    Z = [zs[0]] + zs + [z_end]
    zmin, zmax = min(Z), max(Z)
    r = zmax - zmin
    ky = kx if r < 1e-6 else max(kx, min(kx * 18.0, 110.0 / r))
    yb = 300.0
    fx = lambda sv: mgl + pad_px + sv * kx
    fy = lambda zv: yb - (zv - zmin) * ky
    solo = COR_SOLO.get(cfg.get("tipo_solo"), "#7B5236")
    gp = [(fx(a), fy(b)) for a, b in zip(X, Z)]
    gs = " ".join(f"{a:.1f},{b:.1f}" for a, b in gp)
    bot = H - 18
    o = [f'''<defs><linearGradient id="sk{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7EC0EE"/><stop offset="1" stop-color="#EAF6FF"/></linearGradient>
      <linearGradient id="sl{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="{_misturar(solo, '#000000', 0.05)}"/><stop offset="1" stop-color="{_misturar(solo, '#000000', 0.45)}"/></linearGradient>
      <linearGradient id="aw{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7FB8F0"/><stop offset="1" stop-color="#1F6FB0"/></linearGradient></defs>''']
    o.append(f'<rect width="{W}" height="{H}" fill="url(#sk{uid})" rx="6"/>')
    o.append(f'<text x="{W / 2}" y="26" font-family="Arial" font-size="16" font-weight="bold" fill="#1B2631" text-anchor="middle">Perfil topográfico do trajeto — {"ACLIVE" if L["aclive"] else "DECLIVE"} de {_fmt_br(L["dz"])} m</text>')
    o.append(_svg_sol(W - 80, 70, 18) + _svg_nuvem(180, 70, 1.0) + _svg_nuvem(470, 52, 0.8) + _svg_nuvem(720, 84, 0.9))
    o.append(f'<path d="M0,{yb - 40} Q{W * 0.2},{yb - 95} {W * 0.4},{yb - 55} T{W * 0.75},{yb - 70} T{W},{yb - 45} V{yb} H0Z" fill="#A9D68B" opacity="0.55"/>')
    o.append(f'<polygon points="{gs} {fx(s_end):.1f},{bot} {fx(-pad):.1f},{bot}" fill="url(#sl{uid})"/>')
    o.append(f'<polyline points="{" ".join(f"{a:.1f},{b + 20:.1f}" for a, b in gp)}" fill="none" stroke="{_misturar(solo, "#FFFFFF", 0.25)}" stroke-width="2" opacity="0.5" stroke-dasharray="10,6"/>')
    rng = np.random.RandomState(5)
    for _ in range(26):
        px_ = rng.uniform(fx(-pad) + 10, fx(s_end) - 10)
        base_y = np.interp(px_, [g[0] for g in gp], [g[1] for g in gp])
        py_ = rng.uniform(base_y + 26, bot - 8)
        o.append(f'<ellipse cx="{px_:.1f}" cy="{py_:.1f}" rx="{rng.uniform(3, 7):.1f}" ry="{rng.uniform(2, 4.5):.1f}" fill="#8D8D8D" opacity="0.55"/>')
    o.append(f'<polyline points="{gs}" fill="none" stroke="#4C9A2F" stroke-width="9" stroke-linejoin="round"/>'
             f'<polyline points="{gs}" fill="none" stroke="#7CC152" stroke-width="4" stroke-linejoin="round"/>')

    # casa + cisterna + bomba (à esquerda, fora de escala)
    gy0 = fy(zs[0])
    hx0 = fx(-pad) + 8
    o.append(f'<rect x="{hx0 + 8:.1f}" y="{gy0 - 54:.1f}" width="92" height="52" fill="#F5E6CA" stroke="#7B6A4C" stroke-width="2"/>'
             f'<polygon points="{hx0:.1f},{gy0 - 54:.1f} {hx0 + 54:.1f},{gy0 - 92:.1f} {hx0 + 108:.1f},{gy0 - 54:.1f}" fill="#B5503C" stroke="#7B2D1E" stroke-width="2"/>'
             f'<rect x="{hx0 + 40:.1f}" y="{gy0 - 32:.1f}" width="16" height="30" fill="#6E4B2A"/><rect x="{hx0 + 68:.1f}" y="{gy0 - 42:.1f}" width="22" height="18" fill="#AED6F1" stroke="#7B6A4C"/>'
             f'<text x="{hx0 + 54:.1f}" y="{gy0 - 100:.1f}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B2631" text-anchor="middle">CASA · {_fmt_br(cfg.get("area_telhado", 0), 0)} m²</text>')
    cx0 = hx0 + 128
    niv = min(1.0, max(0.05, cfg.get("nivel_cisterna", 0.6)))
    o.append(f'<rect x="{cx0:.1f}" y="{gy0 - 50:.1f}" width="56" height="50" rx="4" fill="#EAF2F8" stroke="#566573" stroke-width="3"/>'
             f'<rect x="{cx0 + 3:.1f}" y="{gy0 - 3 - 44 * niv:.1f}" width="50" height="{44 * niv:.1f}" fill="url(#aw{uid})"/>'
             f'<text x="{cx0 + 28:.1f}" y="{gy0 - 56:.1f}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B2631" text-anchor="middle">CISTERNA</text>'
             f'<text x="{cx0 + 28:.1f}" y="{gy0 - 66:.1f}" font-family="Arial" font-size="9.5" fill="#1B2631" text-anchor="middle">{_fmt_br(cfg.get("vol_cisterna", 0), 0)} L</text>')
    bx0 = cx0 + 66
    o.append(f'<rect x="{bx0:.1f}" y="{gy0 - 8:.1f}" width="46" height="8" fill="#7F8C8D"/>'
             f'<rect x="{bx0:.1f}" y="{gy0 - 27:.1f}" width="28" height="19" rx="4" fill="#1F618D" stroke="#fff"/>'
             f'<circle cx="{bx0 + 34:.1f}" cy="{gy0 - 17:.1f}" r="10" fill="#5D6D7E" stroke="#fff" stroke-width="1.5"/>'
             f'<text x="{bx0 + 20:.1f}" y="{gy0 - 33:.1f}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B2631" text-anchor="middle">BOMBA {cfg.get("pot_cv", 0):.1f} CV</text>')
    o.append(_svg_tubo([(cx0 + 56, gy0 - 5), (bx0 + 6, gy0 - 5)], 5.0, est))

    # tubulação sobre o terreno + curvas
    off = 6.0
    pp = [(fx(a), fy(b) - off) for a, b in zip(dist, zs)]
    o.append(_svg_tubo([(bx0 + 34, gy0 - 17), (bx0 + 52, gy0 - off - 4), pp[0]] + pp[1:], 8.0, est))
    for i in range(1, len(pp) - 1):
        o.append(f'<circle cx="{pp[i][0]:.1f}" cy="{pp[i][1]:.1f}" r="8" fill="{est[1]}" stroke="#fff" stroke-width="2"/>'
                 f'<text x="{pp[i][0]:.1f}" y="{pp[i][1] + 3.5:.1f}" font-family="Arial" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">{i}</text>'
                 f'<text x="{pp[i][0]:.1f}" y="{pp[i][1] - 16:.1f}" font-family="Arial" font-size="10" fill="#1B2631" text-anchor="middle">curva {i}</text>')
    for i in range(len(pp) - 1):
        mx, my = (pp[i][0] + pp[i + 1][0]) / 2, (pp[i][1] + pp[i + 1][1]) / 2
        o.append(f'<text x="{mx:.1f}" y="{my + 34:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#FFFFFF" text-anchor="middle">L{i + 1} = {_fmt_br(L["comprimentos"][i])} m</text>')
    # filtro e registro
    seg_len = max(pp[1][0] - pp[0][0], 1.0)
    fxp = pp[0][0] + min(46.0, 0.3 * seg_len)
    vxp = pp[0][0] + min(92.0, 0.6 * seg_len)
    yy_ = lambda xp: np.interp(xp, [q[0] for q in pp], [q[1] for q in pp])
    o.append(f'<rect x="{fxp - 7:.1f}" y="{yy_(fxp) - 22:.1f}" width="14" height="24" rx="4" fill="#2C3E50" stroke="#fff"/>'
             f'<text x="{fxp:.1f}" y="{yy_(fxp) - 27:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#1B2631" text-anchor="middle">filtro</text>'
             f'<circle cx="{vxp:.1f}" cy="{yy_(vxp):.1f}" r="7" fill="#E67E22" stroke="#fff" stroke-width="1.5"/>'
             f'<line x1="{vxp:.1f}" y1="{yy_(vxp) - 7:.1f}" x2="{vxp:.1f}" y2="{yy_(vxp) - 17:.1f}" stroke="#E67E22" stroke-width="3"/>'
             f'<text x="{vxp:.1f}" y="{yy_(vxp) - 22:.1f}" font-family="Arial" font-size="10" font-weight="bold" fill="#1B2631" text-anchor="middle">registro</text>')

    # roça (setores) no fim do trajeto
    xr0, xr1 = fx(dist[-1] + ho), fx(s_end)
    ncol = max(3, int((xr1 - xr0) / 13))
    for i in range(ncol):
        xp = xr0 + (i + 0.5) * (xr1 - xr0) / ncol
        yp = np.interp(xp, [g[0] for g in gp], [g[1] for g in gp])
        o.append(_svg_planta_icone(xp, yp, arborea, 22 if arborea else 15))
        if metodo != "Gotejamento" and i % 3 == 1:
            o.append(f'<line x1="{xp:.1f}" y1="{yp:.1f}" x2="{xp:.1f}" y2="{yp - 22:.1f}" stroke="#566573" stroke-width="2"/>'
                     f'<path d="M{xp - 12:.1f},{yp - 12:.1f} Q{xp:.1f},{yp - 38:.1f} {xp + 12:.1f},{yp - 12:.1f}" fill="#5DADE2" opacity="0.35" stroke="#2E86DE" stroke-width="1"/>')
    n_set = len(L['setores'])
    o.append(f'<text x="{(xr0 + xr1) / 2:.1f}" y="{np.interp((xr0 + xr1) / 2, [g[0] for g in gp], [g[1] for g in gp]) - 64:.1f}" font-family="Arial" font-size="12" font-weight="bold" fill="#1B2631" text-anchor="middle">'
             f'ROÇA · {n_set} setor(es) · {L["n_linhas"]} linha(s)</text>')
    # cota do desnível
    xd = fx(dist[-1]) + (xr1 - fx(dist[-1])) * 0.5 if False else fx(s_end) + 14
    y0d, y1d = fy(zs[0]), fy(z_end)
    o.append(f'<line x1="{fx(0):.1f}" y1="{y0d:.1f}" x2="{xd + 8:.1f}" y2="{y0d:.1f}" stroke="#C0392B" stroke-width="1.3" stroke-dasharray="6,4"/>'
             f'<line x1="{xd:.1f}" y1="{y0d:.1f}" x2="{xd:.1f}" y2="{y1d:.1f}" stroke="#C0392B" stroke-width="2"/>'
             f'<text x="{xd - 5:.1f}" y="{(y0d + y1d) / 2:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#C0392B" text-anchor="end" '
             f'stroke="#fff" stroke-width="3" paint-order="stroke">ΔZ = {_fmt_br(L["dz"])} m</text>')
    o.append(f'<text x="{W - 12}" y="{H - 6}" font-family="Arial" font-size="9.5" fill="#1B2631" text-anchor="end">Escala vertical exagerada {ky / kx:.1f}×; casa, cisterna e bomba fora de escala</text>')
    corpo = "".join(o)
    return (f'<div style="display:flex;justify-content:center;margin:8px 0;"><svg width="100%" viewBox="0 0 {W} {H}" '
            f'style="max-width:{W}px;border:1px solid #cfd8dc;border-radius:8px;" xmlns="http://www.w3.org/2000/svg">{corpo}</svg></div>').replace("\n", "")

# ------------------------------------------------------------------------------
# PERFIL DO SOLO E BALANÇO HÍDRICO (corte do solo com raízes, CC / PMP, ETc e lâmina)
# ------------------------------------------------------------------------------
def gerar_svg_perfil_solo(cfg, uid="ps", W=940, H=560):
    solo_nome = cfg.get("tipo_solo", "Franco (médio)")
    cor = COR_SOLO.get(solo_nome, "#7B5236")
    prof = float(cfg.get("prof_raiz", 30.0))
    cad = float(cfg.get("cad_mm_cm", 1.5))
    etc, lam = float(cfg.get("etc", 5.0)), float(cfg.get("lamina_bruta", 5.5))
    turno = int(cfg.get("turno", 3))
    util = float(cfg.get("reserva_util", cad * prof * 0.5))
    arborea = cfg.get("arborea", False)
    metodo = cfg.get("metodo", "Gotejamento")
    cultura = cfg.get("cultura", "")
    topo = 235.0
    prof_max = max(100.0, prof + 25.0)
    ky = (H - 30 - topo) / prof_max
    ys = lambda cm: topo + cm * ky
    xs0, xs1 = 40, 600
    o = [f'''<defs><linearGradient id="sk{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7EC0EE"/><stop offset="1" stop-color="#EAF6FF"/></linearGradient>
      <linearGradient id="ha{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="{_misturar(cor, '#000000', 0.30)}"/><stop offset="1" stop-color="{_misturar(cor, '#000000', 0.12)}"/></linearGradient>
      <linearGradient id="hb{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="{cor}"/><stop offset="1" stop-color="{_misturar(cor, '#FFFFFF', 0.12)}"/></linearGradient>
      <linearGradient id="hc{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="{_misturar(cor, '#FFFFFF', 0.2)}"/><stop offset="1" stop-color="{_misturar(cor, '#FFFFFF', 0.38)}"/></linearGradient>
      <linearGradient id="ag{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7FB8F0"/><stop offset="1" stop-color="#2471A3"/></linearGradient></defs>''']
    o.append(f'<rect width="{W}" height="{H}" fill="url(#sk{uid})" rx="6"/>')
    o.append(_svg_sol(W - 70, 58, 18) + _svg_nuvem(150, 50, 1.0) + _svg_nuvem(400, 38, 0.8) + _svg_nuvem(650, 62, 0.9))
    o.append(f'<text x="{W / 2}" y="24" font-family="Arial" font-size="16" font-weight="bold" fill="#1B2631" text-anchor="middle">Perfil do solo, raízes e balanço hídrico — {cultura} · solo {solo_nome}</text>')
    # horizontes do solo
    o.append(f'<rect x="{xs0}" y="{topo}" width="{xs1 - xs0}" height="{ys(20) - topo:.1f}" fill="url(#ha{uid})"/>'
             f'<rect x="{xs0}" y="{ys(20):.1f}" width="{xs1 - xs0}" height="{ys(60) - ys(20):.1f}" fill="url(#hb{uid})"/>'
             f'<rect x="{xs0}" y="{ys(60):.1f}" width="{xs1 - xs0}" height="{H - 30 - ys(60):.1f}" fill="url(#hc{uid})"/>')
    rng = np.random.RandomState(3)
    n_pt = {"Arenoso": 220, "Franco (médio)": 120, "Argiloso": 60}.get(solo_nome, 120)
    for _ in range(n_pt):
        px_, py_ = rng.uniform(xs0 + 4, xs1 - 4), rng.uniform(topo + 6, H - 34)
        rr = rng.uniform(0.8, 2.0) if solo_nome != "Argiloso" else rng.uniform(0.6, 1.2)
        o.append(f'<circle cx="{px_:.1f}" cy="{py_:.1f}" r="{rr:.1f}" fill="#000" opacity="0.18"/>')
    for _ in range(7):
        o.append(f'<ellipse cx="{rng.uniform(xs0 + 20, xs1 - 20):.1f}" cy="{rng.uniform(ys(55), H - 40):.1f}" rx="{rng.uniform(5, 10):.1f}" ry="{rng.uniform(3, 6):.1f}" fill="#8D8D8D" opacity="0.6"/>')
    o.append(f'<rect x="{xs0}" y="{topo - 6}" width="{xs1 - xs0}" height="8" fill="#5FAE3F"/><rect x="{xs0}" y="{topo - 6}" width="{xs1 - xs0}" height="3" fill="#86C95E"/>')
    for cm in range(0, int(prof_max) + 1, 10):
        o.append(f'<line x1="{xs0 - 6}" y1="{ys(cm):.1f}" x2="{xs0}" y2="{ys(cm):.1f}" stroke="#1B2631" stroke-width="1"/>'
                 f'<text x="{xs0 - 9}" y="{ys(cm) + 3.5:.1f}" font-family="Arial" font-size="9.5" fill="#1B2631" text-anchor="end">{cm}</text>')
    o.append(f'<text x="14" y="{topo + 8}" font-family="Arial" font-size="9.5" fill="#1B2631" transform="rotate(-90 14,{topo + 8}) translate(-60,0)">profundidade (cm)</text>')

    # zona radicular
    pcx = 270.0
    yprof = ys(prof)
    o.append(f'<rect x="{pcx - 150}" y="{topo}" width="300" height="{yprof - topo:.1f}" fill="#2E86DE" opacity="0.13" stroke="#2E86DE" stroke-width="1.5" stroke-dasharray="7,4"/>')
    raizes = ""
    for fr, dx in ((1.0, 0), (0.78, -22), (0.78, 22), (0.55, -48), (0.55, 48), (0.4, -75), (0.4, 75), (0.25, -105), (0.25, 105)):
        yy = topo + (yprof - topo) * fr
        raizes += f'<path d="M{pcx},{topo + 2} C{pcx + dx * 0.25},{topo + (yy - topo) * 0.4} {pcx + dx * 0.9},{topo + (yy - topo) * 0.6} {pcx + dx},{yy:.1f}" fill="none" stroke="#F3E3C2" stroke-width="{2.6 if dx == 0 else 1.6}" stroke-linecap="round"/>'
    o.append(raizes)
    # planta(s)
    if arborea:
        o.append(_svg_planta_icone(pcx, topo - 4, True, 118))
    else:
        for dx in (-70, -35, 0, 35, 70):
            o.append(_svg_planta_icone(pcx + dx, topo - 4, False, 46 if dx == 0 else 38))
    # emissor e água
    if metodo == "Gotejamento":
        o.append(f'<line x1="{pcx - 130}" y1="{topo - 14}" x2="{pcx + 130}" y2="{topo - 14}" stroke="#111" stroke-width="3"/>'
                 f'<rect x="{pcx - 4}" y="{topo - 19}" width="8" height="8" rx="2" fill="#F1C40F" stroke="#111"/>')
        for i, dy in enumerate((6, 20, 34)):
            o.append(f'<path d="M{pcx},{topo - 8 + dy - 7} q-3,5 0,8 q3,-3 0,-8z" fill="#3498DB"/>')
        o.append(f'<ellipse cx="{pcx}" cy="{topo + 26}" rx="48" ry="38" fill="#2E86DE" opacity="0.28" stroke="#2E86DE" stroke-dasharray="4,3"/>')
    else:
        o.append(f'<line x1="{pcx + 70}" y1="{topo}" x2="{pcx + 70}" y2="{topo - 62}" stroke="#566573" stroke-width="4"/><circle cx="{pcx + 70}" cy="{topo - 64}" r="6" fill="#566573"/>'
                 f'<path d="M{pcx + 70},{topo - 64} Q{pcx + 10},{topo - 118} {pcx - 60},{topo - 12} M{pcx + 70},{topo - 64} Q{pcx + 130},{topo - 118} {pcx + 200},{topo - 12}" fill="none" stroke="#3498DB" stroke-width="2.5" stroke-dasharray="5,4"/>')
        o.append(f'<rect x="{pcx - 150}" y="{topo}" width="300" height="{ys(min(prof, 22)) - topo:.1f}" fill="#2E86DE" opacity="0.22"/>')
    # ETc (evapotranspiração)
    for dx in (-100, 100):
        o.append(f'<path d="M{pcx + dx},{topo - 135} C{pcx + dx - 10},{topo - 100} {pcx + dx + 10},{topo - 80} {pcx + dx},{topo - 40}" fill="none" stroke="#1B2631" stroke-width="2" stroke-dasharray="6,4"/>'
                 f'<polygon points="{pcx + dx - 6},{topo - 128} {pcx + dx},{topo - 142} {pcx + dx + 6},{topo - 128}" fill="#1B2631"/>')
    o.append(f'<rect x="{pcx - 150}" y="{topo - 176}" width="300" height="24" rx="5" fill="#FFE600" stroke="#1B2631"/>'
             f'<text x="{pcx}" y="{topo - 159}" font-family="Arial" font-size="13" font-weight="bold" fill="#1B2631" text-anchor="middle">EVAPOTRANSPIRAÇÃO · ETc = {_fmt_br(etc, 2)} mm/dia</text>')
    # chaves e rótulos laterais
    o.append(f'<line x1="{xs1 - 8}" y1="{topo}" x2="{xs1 - 8}" y2="{yprof:.1f}" stroke="#1B2631" stroke-width="2"/>'
             f'<line x1="{xs1 - 14}" y1="{topo}" x2="{xs1 - 2}" y2="{topo}" stroke="#1B2631" stroke-width="2"/><line x1="{xs1 - 14}" y1="{yprof:.1f}" x2="{xs1 - 2}" y2="{yprof:.1f}" stroke="#1B2631" stroke-width="2"/>'
             f'<text x="{xs1 - 14}" y="{(topo + yprof) / 2:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#fff" text-anchor="end" stroke="#000" stroke-width="2.5" paint-order="stroke">raiz efetiva {_fmt_br(prof, 0)} cm</text>')
    for cm, nome in ((10, "horizonte A"), (40, "horizonte B"), (min(prof_max - 8, 80), "horizonte C")):
        o.append(f'<text x="{xs0 + 8}" y="{ys(cm):.1f}" font-family="Arial" font-size="10.5" font-style="italic" fill="#fff" stroke="#000" stroke-width="2" paint-order="stroke">{nome}</text>')

    # painel do balanço hídrico (CC / 50% / PMP)
    bx, bw = 660, 52
    total = cad * prof
    h_total = 200.0
    y_cc = 150.0
    y_pmp = y_cc + h_total
    y50 = y_cc + h_total * 0.5
    o.append(f'<rect x="{bx - 40}" y="{topo - 150}" width="270" height="{H - topo + 120:.0f}" rx="8" fill="#fff" opacity="0.88" stroke="#cfd8dc"/>')
    o.append(f'<text x="{bx + 95}" y="{topo - 128}" font-family="Arial" font-size="12" font-weight="bold" fill="#1B2631" text-anchor="middle">BALANÇO HÍDRICO DA ZONA RADICULAR</text>')
    o.append(f'<rect x="{bx}" y="{y_cc}" width="{bw}" height="{h_total:.1f}" fill="#fff" stroke="#1B2631" stroke-width="2"/>'
             f'<rect x="{bx}" y="{y50:.1f}" width="{bw}" height="{y_pmp - y50:.1f}" fill="#B9C2CB" opacity="0.7"/>'
             f'<rect x="{bx}" y="{y_cc}" width="{bw}" height="{y50 - y_cc:.1f}" fill="url(#ag{uid})"/>')
    o.append(f'<line x1="{bx - 12}" y1="{y_cc}" x2="{bx + bw + 14}" y2="{y_cc}" stroke="#1B4F8A" stroke-width="2.5"/>'
             f'<line x1="{bx - 12}" y1="{y50:.1f}" x2="{bx + bw + 14}" y2="{y50:.1f}" stroke="#E67E22" stroke-width="2.5" stroke-dasharray="6,3"/>'
             f'<line x1="{bx - 12}" y1="{y_pmp:.1f}" x2="{bx + bw + 14}" y2="{y_pmp:.1f}" stroke="#C0392B" stroke-width="2.5"/>')
    for yy, ttl, sub, cc in ((y_cc, "CC · capacidade de campo", "solo saturado útil", "#1B4F8A"),
                             (y50, "Nível de manejo (50% da CAD)", "hora de irrigar", "#E67E22"),
                             (y_pmp, "PMP · ponto de murcha", "planta não extrai mais", "#C0392B")):
        o.append(f'<text x="{bx + bw + 20}" y="{yy + 1:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="{cc}">{ttl}</text>'
                 f'<text x="{bx + bw + 20}" y="{yy + 14:.1f}" font-family="Arial" font-size="9.5" fill="#566573">{sub}</text>')
    o.append(f'<text x="{bx + bw / 2}" y="{(y_cc + y50) / 2 + 4:.1f}" font-family="Arial" font-size="11" font-weight="bold" fill="#fff" text-anchor="middle">{_fmt_br(util, 1)} mm</text>'
             f'<text x="{bx + bw / 2}" y="{y_cc - 8}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B4F8A" text-anchor="middle">CAD {_fmt_br(total, 1)} mm</text>')
    resumo = [("Reserva útil (50%)", f"{_fmt_br(util, 1)} mm"), ("Turno de rega", f"a cada {turno} dia(s)"),
              ("Lâmina bruta aplicada", f"{_fmt_br(lam, 2)} mm/dia"), ("CAD do solo", f"{_fmt_br(cad, 2)} mm/cm")]
    for i, (a, b) in enumerate(resumo):
        yy = H - 98 + i * 20
        o.append(f'<text x="{bx - 26}" y="{yy}" font-family="Arial" font-size="11.5" fill="#566573">{a}:</text>'
                 f'<text x="{bx + 214}" y="{yy}" font-family="Arial" font-size="11.5" font-weight="bold" fill="#1B2631" text-anchor="end">{b}</text>')
    corpo = "".join(o)
    return (f'<div style="display:flex;justify-content:center;margin:8px 0;"><svg width="100%" viewBox="0 0 {W} {H}" '
            f'style="max-width:{W}px;border:1px solid #cfd8dc;border-radius:8px;" xmlns="http://www.w3.org/2000/svg">{corpo}</svg></div>').replace("\n", "")

# ------------------------------------------------------------------------------
# ESQUEMA DO SISTEMA DE CAPTAÇÃO (da casa à roça), com os valores reais do projeto
# ------------------------------------------------------------------------------
def gerar_svg_cisterna_realista(cfg, uid="ci", W=960, H=540):
    est = COR_TUBO.get(cfg.get("material", "PVC"), COR_TUBO["PVC"])
    niv = min(1.0, max(0.05, cfg.get("nivel_cisterna", 0.6)))
    backup = cfg.get("nivel_backup", 20) / 100.0
    n_set = int(cfg.get("n_setores", 1))
    metodo = cfg.get("metodo", "Gotejamento")
    arborea = cfg.get("arborea", False)
    gy = 300.0
    o = [f'''<defs><linearGradient id="sk{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7EC0EE"/><stop offset="1" stop-color="#EAF6FF"/></linearGradient>
      <linearGradient id="so{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8B5E3C"/><stop offset="1" stop-color="#5A3A22"/></linearGradient>
      <linearGradient id="aw{uid}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7FB8F0"/><stop offset="1" stop-color="#1F6FB0"/></linearGradient>
      <pattern id="tl{uid}" width="9" height="7" patternUnits="userSpaceOnUse"><rect width="9" height="7" fill="#C0563F"/><path d="M0 6.5 H9 M4.5 0 V6.5" stroke="#9C402D" stroke-width="0.9" fill="none"/></pattern></defs>''']
    o.append(f'<rect width="{W}" height="{H}" fill="url(#sk{uid})" rx="6"/>')
    o.append(f'<text x="{W / 2}" y="24" font-family="Arial" font-size="16" font-weight="bold" fill="#1B2631" text-anchor="middle">Esquema do sistema: da casa até a roça (cisterna e conjunto motobomba)</text>')
    o.append(_svg_sol(W - 60, 66, 18) + _svg_nuvem(110, 62, 1.0, "#E5E8EC") + _svg_nuvem(330, 48, 0.85, "#E5E8EC") + _svg_nuvem(560, 70, 0.9, "#E5E8EC"))
    for x_ in (70, 100, 130, 160, 190, 220, 350, 380, 410, 580, 610):  # chuva
        o.append(f'<line x1="{x_}" y1="92" x2="{x_ - 6}" y2="110" stroke="#3498DB" stroke-width="2" opacity="0.8"/>')
    o.append(f'<path d="M0,{gy - 20} Q200,{gy - 50} 400,{gy - 25} T{W},{gy - 30} V{gy} H0Z" fill="#A9D68B" opacity="0.5"/>')
    o.append(f'<rect x="0" y="{gy}" width="{W}" height="{H - gy}" fill="url(#so{uid})"/><rect x="0" y="{gy - 5}" width="{W}" height="9" fill="#5FAE3F"/><rect x="0" y="{gy - 5}" width="{W}" height="3" fill="#86C95E"/>')
    rng = np.random.RandomState(9)
    for _ in range(40):
        o.append(f'<ellipse cx="{rng.uniform(5, W - 5):.0f}" cy="{rng.uniform(gy + 14, H - 8):.0f}" rx="{rng.uniform(3, 8):.1f}" ry="{rng.uniform(2, 4.5):.1f}" fill="#8D8D8D" opacity="0.45"/>')
    # casa
    o.append(f'<rect x="38" y="{gy - 112}" width="196" height="108" fill="#F5E6CA" stroke="#7B6A4C" stroke-width="3"/>'
             f'<polygon points="20,{gy - 112} 136,{gy - 192} 252,{gy - 112}" fill="url(#tl{uid})" stroke="#7B2D1E" stroke-width="3"/>'
             f'<rect x="120" y="{gy - 66}" width="32" height="62" fill="#6E4B2A" stroke="#4A3320" stroke-width="2"/><circle cx="145" cy="{gy - 34}" r="2.5" fill="#F4D03F"/>'
             f'<rect x="58" y="{gy - 90}" width="42" height="38" fill="#AED6F1" stroke="#7B6A4C" stroke-width="2"/><line x1="79" y1="{gy - 90}" x2="79" y2="{gy - 52}" stroke="#7B6A4C"/>'
             f'<rect x="176" y="{gy - 90}" width="42" height="38" fill="#AED6F1" stroke="#7B6A4C" stroke-width="2"/><line x1="197" y1="{gy - 90}" x2="197" y2="{gy - 52}" stroke="#7B6A4C"/>'
             f'<text x="136" y="{gy - 144}" font-family="Arial" font-size="12" font-weight="bold" fill="#fff" text-anchor="middle">TELHADO</text>'
             f'<text x="136" y="{gy - 130}" font-family="Arial" font-size="11" fill="#fff" text-anchor="middle">{_fmt_br(cfg.get("area_telhado", 0), 0)} m² · runoff {_fmt_br(cfg.get("runoff", 0.85), 2)}</text>')
    # calha, condutor, descarte, filtro (PVC branco)
    branco = ("#ECF0F1", "#7F8C8D", "#FFFFFF")
    o.append(_svg_tubo([(246, gy - 112), (270, gy - 112), (270, gy - 22), (330, gy - 22), (330, gy + 36), (350, gy + 36)], 9.0, branco))
    o.append(f'<text x="276" y="{gy - 120}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B2631">2 Calha {cfg.get("diam_calha", "")} e condutor {cfg.get("diam_condutor", "")}</text>')
    o.append(_svg_tubo([(270, gy - 70), (310, gy - 70), (310, gy - 6)], 8.0, branco))
    o.append(f'<circle cx="310" cy="{gy - 4}" r="6" fill="#95A5A6" stroke="#1B2631"/>'
             f'<text x="318" y="{gy - 56}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B2631">3 Descarte da 1ª água</text>'
             f'<text x="318" y="{gy - 43}" font-family="Arial" font-size="10" fill="#1B2631">{_fmt_br(cfg.get("vol_ff", 0), 0)} L</text>')
    o.append(f'<rect x="262" y="{gy - 62}" width="16" height="30" rx="4" fill="#F5CBA7" stroke="#7F8C8D" stroke-width="2"/>'
             f'<text x="256" y="{gy - 42}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#1B2631" text-anchor="end" stroke="#F5E6CA" stroke-width="3" paint-order="stroke">4 Filtro</text>')
    # cisterna (corte)
    cx0, cx1, cy0, cy1 = 350, 600, gy + 14, gy + 196
    hw = cy1 - cy0
    o.append(f'<rect x="{cx0 - 6}" y="{cy0 - 4}" width="{cx1 - cx0 + 12}" height="{hw + 10}" fill="#7F8C8D" stroke="#566573" stroke-width="3" rx="4"/>'
             f'<rect x="{cx0}" y="{cy0}" width="{cx1 - cx0}" height="{hw}" fill="#EAF2F8"/>'
             f'<rect x="{cx0}" y="{cy1 - hw * niv:.1f}" width="{cx1 - cx0}" height="{hw * niv:.1f}" fill="url(#aw{uid})" opacity="0.92"/>'
             f'<path d="M{cx0},{cy1 - hw * niv:.1f} q30,-6 62,0 t62,0 t62,0 t64,0" fill="none" stroke="#fff" stroke-width="2" opacity="0.7"/>'
             f'<rect x="{cx0 + 100}" y="{gy - 10}" width="64" height="26" fill="#95A5A6" stroke="#566573" stroke-width="2"/>'
             f'<text x="{cx0 + 132}" y="{gy - 15}" font-family="Arial" font-size="9.5" fill="#1B2631" text-anchor="middle">tampa de inspeção</text>')
    ybk = cy1 - hw * backup
    o.append(f'<line x1="{cx0 - 10}" y1="{ybk:.1f}" x2="{cx1 + 10}" y2="{ybk:.1f}" stroke="#E74C3C" stroke-width="2.4" stroke-dasharray="8,4"/>'
             f'<text x="{cx0 + 8}" y="{ybk - 5:.1f}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#C0392B" stroke="#fff" stroke-width="3" paint-order="stroke">nível crítico {cfg.get("nivel_backup", 20)}% (abre a rede pública)</text>')
    o.append(f'<text x="{(cx0 + cx1) / 2}" y="{cy0 + 38:.1f}" font-family="Arial" font-size="14" font-weight="bold" fill="#1B2631" text-anchor="middle" stroke="#fff" stroke-width="4" paint-order="stroke">5 CISTERNA · {_fmt_br(cfg.get("vol_cisterna", 0), 0)} L</text>'
             f'<text x="{(cx0 + cx1) / 2}" y="{cy0 + 56:.1f}" font-family="Arial" font-size="11" fill="#1B2631" text-anchor="middle" stroke="#fff" stroke-width="4" paint-order="stroke">água limpa captada: {_fmt_br(cfg.get("vol_liq", 0), 0)} L ({niv * 100:.0f}% da cisterna)</text>')
    # rede pública com boia (entra pela lateral esquerda da cisterna, abaixo do solo)
    yr = cy0 + 108
    o.append(_svg_tubo([(194, yr), (cx0 - 6, yr), (cx0 + 14, yr)], 8.0, ("#7FA7C9", "#34607F", "#CFE3F5")))
    o.append(f'<circle cx="{cx0 + 18}" cy="{yr}" r="7" fill="#E67E22" stroke="#fff" stroke-width="1.5"/>'
             f'<rect x="36" y="{yr - 13}" width="158" height="26" rx="5" fill="#fff" stroke="#34607F"/>'
             f'<text x="115" y="{yr + 4}" font-family="Arial" font-size="10.5" font-weight="bold" fill="#34607F" text-anchor="middle">REDE PÚBLICA · {_fmt_br(cfg.get("vazao_rede", 0), 0)} L/h</text>')
    # extravasor
    o.append(_svg_tubo([(cx0 - 6, cy0 + 22), (cx0 - 36, cy0 + 22), (cx0 - 36, cy0 + 60)], 6.0, branco))
    o.append(f'<text x="{cx0 - 42}" y="{cy0 + 76}" font-family="Arial" font-size="10" fill="#fff" text-anchor="end">extravasor</text>')
    # bomba (sucção afogada) e recalque
    sy = cy1 - 26
    o.append(_svg_tubo([(cx1 + 4, sy), (650, sy), (650, gy - 16)], 8.0, est))
    o.append(f'<rect x="632" y="{gy - 6}" width="96" height="8" fill="#7F8C8D"/>'
             f'<rect x="664" y="{gy - 44}" width="62" height="38" rx="6" fill="#1F618D" stroke="#fff" stroke-width="2"/>'
             f'<circle cx="650" cy="{gy - 24}" r="17" fill="#5D6D7E" stroke="#fff" stroke-width="2"/>'
             f'<text x="700" y="{gy + 28}" font-family="Arial" font-size="11.5" font-weight="bold" fill="#fff" text-anchor="middle" stroke="#000" stroke-width="3" paint-order="stroke">6 BOMBA {cfg.get("pot_cv", 0):.1f} CV</text>'
             f'<text x="650" y="{cy1 + 20}" font-family="Arial" font-size="10" fill="#fff" text-anchor="middle">sucção afogada</text>')
    ry = gy - 24
    o.append(_svg_tubo([(650, gy - 41), (650, gy - 70), (760, gy - 70), (760, gy - 16)], 8.0, est))
    o.append(f'<rect x="700" y="{gy - 82}" width="16" height="24" rx="4" fill="#2C3E50" stroke="#fff"/><circle cx="738" cy="{gy - 70}" r="6" fill="#E67E22" stroke="#fff"/>'
             f'<text x="708" y="{gy - 88}" font-family="Arial" font-size="9.5" font-weight="bold" fill="#1B2631" text-anchor="middle">filtro</text>'
             f'<text x="706" y="{gy - 100}" font-family="Arial" font-size="11" font-weight="bold" fill="#1B2631" text-anchor="middle">7 Recalque ({cfg.get("material", "PVC")} DN {cfg.get("dn_mm", 50):.0f})</text>')
    # roça com setores
    fx0, fx1 = 772, W - 14
    wsec = (fx1 - fx0) / max(n_set, 1)
    o.append(f'<rect x="{fx0 - 4}" y="{gy - 8}" width="{fx1 - fx0 + 8}" height="20" fill="#6D4C41" rx="3"/>')
    for i in range(n_set):
        cs_ = _misturar("#6D4C41", PALETA_SETORES[i % 8], 0.5)
        x0_ = fx0 + i * wsec
        o.append(f'<rect x="{x0_:.1f}" y="{gy - 8}" width="{wsec:.1f}" height="20" fill="{cs_}" stroke="#fff" stroke-width="1.5" stroke-dasharray="5,3"/>'
                 f'<line x1="{x0_ + wsec / 2:.1f}" y1="{gy - 16}" x2="{x0_ + wsec / 2:.1f}" y2="{gy - 40}" stroke="{est[1]}" stroke-width="3"/>'
                 f'<circle cx="{x0_ + wsec / 2:.1f}" cy="{gy - 46}" r="9" fill="#27AE60" stroke="#fff" stroke-width="1.5"/>'
                 f'<text x="{x0_ + wsec / 2:.1f}" y="{gy - 42.5}" font-family="Arial" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">V{i + 1}</text>')
        for j in range(3 if wsec > 40 else 2):
            xp = x0_ + (j + 0.7) * wsec / (3 if wsec > 40 else 2) - 8
            o.append(_svg_planta_icone(xp + 8, gy - 8, arborea, 24 if arborea else 17))
            if metodo != "Gotejamento":
                o.append(f'<path d="M{xp - 4:.1f},{gy - 64} Q{xp + 8:.1f},{gy - 82} {xp + 20:.1f},{gy - 64}" fill="#5DADE2" opacity="0.4" stroke="#2E86DE"/>')
    o.append(_svg_tubo([(760, gy - 16), (fx0, gy - 16), (fx1 - 4, gy - 16)], 6.0, est))
    o.append(f'<text x="{(fx0 + fx1) / 2:.1f}" y="{gy - 118}" font-family="Arial" font-size="13" font-weight="bold" fill="#1B2631" text-anchor="middle">8 ROÇA</text>'
             f'<text x="{(fx0 + fx1) / 2:.1f}" y="{gy - 103}" font-family="Arial" font-size="11" fill="#1B2631" text-anchor="middle">{n_set} setor(es) · {cfg.get("cultura", "")}</text>')
    # balanço no rodapé
    faixa = (f'Chuva de cálculo {_fmt_br(cfg.get("precip", 0), 1)} mm  →  captado {_fmt_br(cfg.get("vol_bruto", 0), 0)} L  −  descarte {_fmt_br(cfg.get("vol_ff", 0), 0)} L  =  '
             f'aproveitável {_fmt_br(cfg.get("vol_liq", 0), 0)} L  ·  autonomia {_fmt_br(cfg.get("autonomia", 0), 1)} dia(s)')
    o.append(f'<rect x="20" y="{H - 38}" width="{W - 40}" height="26" rx="6" fill="#fff" opacity="0.9"/>'
             f'<text x="{W / 2}" y="{H - 20}" font-family="Arial" font-size="12" font-weight="bold" fill="#1B2631" text-anchor="middle">{faixa}</text>')
    corpo = "".join(o)
    return (f'<div style="display:flex;justify-content:center;margin:8px 0;"><svg width="100%" viewBox="0 0 {W} {H}" '
            f'style="max-width:{W}px;border:1px solid #cfd8dc;border-radius:8px;" xmlns="http://www.w3.org/2000/svg">{corpo}</svg></div>').replace("\n", "")

# ==============================================================================
# PROCESSADOR HIDRÁULICO (núcleo numérico, sem dependência de Streamlit)
#  - curva característica da bomba H(Q), rendimento, potência e NPSHr
#  - curva do sistema e ponto de operação (bomba x sistema)
#  - solver de rede por Newton (algoritmo do gradiente global, Todini e Pilati):
#    tubos (Hazen-Williams), perdas localizadas, bomba por curva, reservatórios
#    de nível fixo, demandas fixas e demandas dependentes da pressão (emissores)
# Unidades: vazão em m³/h nas entradas; internamente m³/s. Cargas em mca.
# ==============================================================================
HW_N = 1.852


def proc_hw_r(L, D_mm, C):
    """Coeficiente r de h = r·Q^1,852 (Q em m³/s), Hazen-Williams SI: 10,67·L/(C^1,852·D^4,87)."""
    D = D_mm / 1000.0
    if L <= 0 or D <= 0 or C <= 0:
        return 0.0
    return 10.67 * L / ((C ** HW_N) * (D ** 4.87))


def proc_coefs_bomba(h0, pontos):
    """H(Q) = h0 + a1·Q + a2·Q² (Q em m³/h).
    1 ponto de catálogo (Q, H): parábola H = h0 − c·Q² (a1 = 0), sempre decrescente.
    2 pontos de catálogo: parábola exata por (0,h0) e pelos dois pontos."""
    if len(pontos) == 1:
        q1, h1 = pontos[0]
        return float(h0), 0.0, float(-(h0 - h1) / (q1 ** 2))
    (q1, h1), (q2, h2) = pontos[:2]
    A = np.array([[q1, q1 ** 2], [q2, q2 ** 2]], dtype=float)
    a1, a2 = np.linalg.solve(A, np.array([h1 - h0, h2 - h0], dtype=float))
    return float(h0), float(a1), float(a2)


def proc_q_max(cf):
    """Vazão em que H = 0 (fim da curva) ou None se a curva não corta o eixo."""
    h0, a1, a2 = cf
    if abs(a2) < 1e-12:
        return None
    disc = a1 * a1 - 4 * a2 * h0
    if disc < 0:
        return None
    raizes = [(-a1 + math.sqrt(disc)) / (2 * a2), (-a1 - math.sqrt(disc)) / (2 * a2)]
    pos = [r for r in raizes if r > 0]
    return min(pos) if pos else None


def proc_curva_valida(cf):
    """Admissível se for decrescente em [0, Qmax]."""
    h0, a1, a2 = cf
    qm = proc_q_max(cf)
    return bool(qm) and h0 > 0 and a1 <= 1e-9 and (a1 + 2 * a2 * qm) <= 1e-9


def ajustar_curva_catalogo(q, h, h0=None):
    """Ajusta H(Q) = c0 + a1·Q + a2·Q² aos pontos tabelados do catálogo (mínimos quadrados).
    Se a curva livre não for decrescente, usa H = c0 + a2·Q². Devolve (coeficientes, erro médio em mca) ou None."""
    q = np.asarray(q, dtype=float); h = np.asarray(h, dtype=float)
    Q = np.append(q, 0.0) if h0 else q
    Hh = np.append(h, float(h0)) if h0 else h
    try:
        a2, a1, c0 = np.polyfit(Q, Hh, 2)
        cf = (float(c0), float(a1), float(a2))
        if not proc_curva_valida(cf):
            A = np.vstack([np.ones_like(Q), Q ** 2]).T
            (c0, a2), *_ = np.linalg.lstsq(A, Hh, rcond=None)
            cf = (float(c0), 0.0, float(a2))
        if not proc_curva_valida(cf):
            return None
        erro = float(np.sqrt(np.mean((np.array([cf[0] + cf[1] * x + cf[2] * x * x for x in q]) - h) ** 2)))
        return cf, erro
    except Exception:
        return None


def proc_h_bomba(Q, cf):
    h0, a1, a2 = cf
    return h0 + a1 * Q + a2 * Q * Q


def proc_eta(Q, q_bep, eta_bep):
    """Rendimento parabólico com máximo no ponto de melhor eficiência: η = ηb·(2x − x²), x = Q/Qb."""
    x = Q / q_bep if q_bep > 0 else 0.0
    return max(0.05, eta_bep * (2 * x - x * x))


def proc_npshr(Q, q_bep, npshr_bep):
    """NPSHr aproximado: cresce com o quadrado da vazão (30 % do valor no BEP a vazão nula). Aproximação — use a curva do fabricante."""
    x = Q / q_bep if q_bep > 0 else 0.0
    return npshr_bep * (0.3 + 0.7 * x * x)


def proc_pot_eixo_kw(Q_m3h, H, eta):
    return 9.81 * (Q_m3h / 3600.0) * H / max(eta, 1e-3)


def proc_h_sistema(Q, s):
    """Carga que o sistema exige na vazão Q (m³/h): estática + sucção + adutora + filtro + pressão da lateral."""
    qr = max(s["q_ref"], 1e-9)
    q_s = Q / 3600.0
    hf_rec = proc_hw_r(s["L_rec"], s["D_rec"], s["C_rec"]) * (q_s ** HW_N) * (1.0 + s["env"])
    hf_suc = s["hf_suc"] * (Q / qr) ** HW_N
    hf_fil = s["hf_fil"] * (Q / qr) ** 2
    hin = s["hin"] * (max(Q, 0.0) / qr) ** (1.0 / max(s["x"], 0.05))
    return s["h_est"] + hf_suc + hf_rec + hf_fil + hin


def proc_ponto_operacao(cf, qmax, s):
    """Interseção bomba x sistema por bisecção. Retorna Q (m³/h) ou None se a bomba não vence a carga estática."""
    f = lambda q: proc_h_bomba(q, cf) - proc_h_sistema(q, s)
    lo, hi = 1e-6, float(qmax)
    if f(lo) <= 0:
        return None
    if f(hi) >= 0:
        return hi
    for _ in range(100):
        mid = 0.5 * (lo + hi)
        if f(mid) > 0:
            lo = mid
        else:
            hi = mid
    return 0.5 * (lo + hi)


# ------------------------------------------------------------------------------
# Solver de rede
# ------------------------------------------------------------------------------
def proc_resolver_rede(nos, elementos, itmax=200, tol=1e-8):
    """
    nos: lista de dict(id, z, H_fixo (None = nó livre), d_nom (m³/h), p_nom (mca, 0 = demanda fixa), x)
    elementos: lista de dict(id, de, para, tipo ('tubo'|'perda'|'bomba'), L, D, C, hf_nom, q_ref, cf=(h0,a1,a2))
        tubo : h = r·Q^1,852 (Hazen-Williams, L já com comprimentos equivalentes)
        perda: h = k·Q²  com k = hf_nom/(q_ref/3600)²
        bomba: ganho H = h0 + a1·Q + a2·Q² (Q em m³/h)
    Retorna dict com cargas, vazões, perdas, convergência.
    """
    idx = {n["id"]: i for i, n in enumerate(nos)}
    livres = [n for n in nos if n["H_fixo"] is None]
    jx = {n["id"]: k for k, n in enumerate(livres)}
    nj, npipe = len(livres), len(elementos)
    if npipe == 0 or nj == 0:
        return {"ok": False, "msg": "Rede sem trechos ou sem nós livres."}

    B = np.zeros((npipe, nj))
    B0 = np.zeros(npipe)
    r = np.zeros(npipe)
    ex = np.full(npipe, HW_N)
    is_p = np.zeros(npipe, dtype=bool)
    cfs = [None] * npipe
    for p, e in enumerate(elementos):
        for end, sgn in ((e["de"], 1.0), (e["para"], -1.0)):
            if end not in idx:
                return {"ok": False, "msg": f"Nó '{end}' do trecho '{e['id']}' não existe."}
            if end in jx:
                B[p, jx[end]] = sgn
            else:
                B0[p] += sgn * nos[idx[end]]["H_fixo"]
        t = e["tipo"]
        if t == "tubo":
            r[p] = proc_hw_r(e["L"], e["D"], e["C"])
        elif t == "perda":
            qr = max(e.get("q_ref", 1.0), 1e-9) / 3600.0
            r[p] = e["hf_nom"] / (qr ** 2)
            ex[p] = 2.0
        elif t == "bomba":
            is_p[p] = True
            cfs[p] = e["cf"]
        else:
            return {"ok": False, "msg": f"Tipo desconhecido: {t}"}
        if not is_p[p]:
            r[p] = max(r[p], 1e-6)  # evita singularidade em trechos de comprimento ou perda nulos

    z = np.array([n["z"] for n in livres], dtype=float)
    dn = np.array([n["d_nom"] for n in livres], dtype=float) / 3600.0
    pn = np.array([n["p_nom"] for n in livres], dtype=float)
    xe = np.array([n["x"] for n in livres], dtype=float)

    def demanda(H):
        P = H - z
        d = dn.copy()
        dd = np.zeros(nj)
        pdd = (pn > 0) & (xe > 0)
        Pp = np.maximum(P, 1e-3)
        ratio = Pp / np.where(pn > 0, pn, 1.0)
        d = np.where(pdd, dn * np.where(P > 0, ratio, 0.0) ** xe, dn)
        dd = np.where(pdd, dn * xe / np.where(pn > 0, pn, 1.0) * ratio ** (xe - 1.0), 0.0)
        return d, dd

    def f_e(Q):
        f = np.zeros(npipe)
        df = np.zeros(npipe)
        aq = np.maximum(np.abs(Q), 1e-7)
        loss = r * (aq ** (ex - 1.0)) * Q
        dloss = ex * r * aq ** (ex - 1.0)
        f = -loss
        df = -dloss
        for p in np.where(is_p)[0]:
            h0, a1, a2 = cfs[p]
            q3 = Q[p] * 3600.0
            f[p] = h0 + a1 * q3 + a2 * q3 * q3
            df[p] = (a1 + 2 * a2 * q3) * 3600.0
        return f, df

    def residuos(Q, H):
        d, _ = demanda(H)
        f, _ = f_e(Q)
        return np.concatenate([B @ H + B0 + f, B.T @ Q + d])

    H = np.full(nj, max([n["H_fixo"] for n in nos if n["H_fixo"] is not None] + [float(z.max())]))
    Q = np.full(npipe, 1e-3)
    for p in np.where(is_p)[0]:
        Q[p] = 1e-3
    res = residuos(Q, H)
    nres = float(np.linalg.norm(res, np.inf))
    ok, it = False, 0
    for it in range(1, itmax + 1):
        d, dd = demanda(H)
        f, df = f_e(Q)
        J = np.zeros((npipe + nj, npipe + nj))
        J[:npipe, :npipe] = np.diag(df)
        J[:npipe, npipe:] = B
        J[npipe:, :npipe] = B.T
        J[npipe:, npipe:] = np.diag(dd)
        try:
            dx = np.linalg.solve(J, -res)
        except np.linalg.LinAlgError:
            return {"ok": False, "msg": "Sistema singular: confira se todos os nós estão conectados à fonte."}
        passo = 1.0
        while passo > 1e-4:
            Qn, Hn = Q + passo * dx[:npipe], H + passo * dx[npipe:]
            rn = residuos(Qn, Hn)
            nn = float(np.linalg.norm(rn, np.inf))
            if nn < nres or nn < tol:
                break
            passo *= 0.5
        Q, H, res, nres = Qn, Hn, rn, nn
        if nres < tol:
            ok = True
            break
    d, _ = demanda(H)
    f, _ = f_e(Q)
    Hn_all = {n["id"]: (n["H_fixo"] if n["H_fixo"] is not None else float(H[jx[n["id"]]])) for n in nos}
    dem_real = {n["id"]: (float(d[jx[n["id"]]]) * 3600.0 if n["id"] in jx else 0.0) for n in nos}
    hf = {}
    for p, e in enumerate(elementos):
        if is_p[p]:
            hf[e["id"]] = -float(f[p])  # negativo = ganho da bomba
        else:
            hf[e["id"]] = float(-f[p])
    return {"ok": ok, "iter": it, "residuo": nres, "H": Hn_all, "Q_m3h": {e["id"]: float(Q[p]) * 3600.0 for p, e in enumerate(elementos)},
            "hf": hf, "demanda_m3h": dem_real, "msg": "" if ok else "Não convergiu."}


df_chuvas_completo = carregar_dados_pluviometricos()
df_matriz_culturas, df_solo_carac = carregar_dados_agronomicos()
df_top10_ne, df_infiltracao_embrapa, fontes_embrapa = carregar_dados_agronomicos_extra()
df_catalogo = carregar_catalogo_municipios()


# ==============================================================================
# PROJETO PRÉ-CARREGADO (back-test do artigo da Revista FT) — botão na aba 11 ou ?preset=artigo na URL
# ==============================================================================
PRESET_ARTIGO = {
    "metodo": "Aspersão", "area": 10000.0, "dist_ad": 10.0, "desn_rec": 0.0, "alt_suc": 0.0, "ov_vib_on": True, "ov_vib": 12.0,
    "q_emissor": 2720.0, "s_emissor": 18.0, "s_linhas": 18.0,
    "num_setores": 2, "comp_lat": 50.0, "man_nem": True, "nem_manual": 3, "man_nlin": True, "nlin_manual": 10,
    "pnom": 25.0, "ha": 2.0, "c_rug": 150.0, "d_lat": 40.0, "diam_adutora": 100.0, "conex": 0.0, "l_princ": 100.0, "margem": 0.0,
    "hf_fil": 0.0, "hf_suc": 0.0, "ploc": 4.0, "etam3": 92.0, "rend_b": 65, "bt_lsuc": 5.0,
}


def aplicar_preset_artigo():
    for k_, v_ in PRESET_ARTIGO.items():
        st.session_state[k_] = v_
    st.session_state["_met_prev"] = "Aspersão"
    st.session_state["_area_prev"] = 10000.0


def limpar_preset_artigo():
    for k_ in list(PRESET_ARTIGO.keys()) + ["_met_prev", "_area_prev", "diam_succao_mm"]:
        st.session_state.pop(k_, None)


if st.query_params.get("preset") == "artigo" and not st.session_state.get("_preset_ok"):
    aplicar_preset_artigo()
    st.session_state["_preset_ok"] = True

# ==============================================================================
# SIDEBAR: DADOS DE ENTRADA REAIS 
# ==============================================================================
st.sidebar.header("1. Localização e Clima")
dados_cidade = None
fonte_chuva = ""
nota_fonte_chuva = ""

if not df_catalogo.empty:
    # Base completa de municípios (IBGE): todos os 5.570 municípios disponíveis
    estados = sorted(df_catalogo['uf'].unique())
    idx_uf = estados.index('PI') if 'PI' in estados else 0
    uf_selecionado = st.sidebar.selectbox("Estado (UF):", estados, index=idx_uf, help="Dado de localização. Filtra a base de municípios do IBGE.")
    cidades = sorted(df_catalogo[df_catalogo['uf'] == uf_selecionado]['cidade'].unique())
    idx_cid = cidades.index('Teresina') if 'Teresina' in cidades else 0
    cidade_input = st.sidebar.selectbox("Município:", cidades, index=idx_cid, help="Define o histórico pluviométrico que será usado nos cálculos de captação de chuva.")
    linha_mun = df_catalogo[(df_catalogo['uf'] == uf_selecionado) & (df_catalogo['cidade'] == cidade_input)].iloc[0]
    usar_chuva_online = st.sidebar.checkbox(
        "Buscar chuva online (NASA POWER) se o município não tiver histórico", value=True,
        help="Se desmarcado, municípios sem histórico no CSV usam o município mais próximo que tenha dados."
    )
    dados_cidade, fonte_chuva, nota_fonte_chuva = obter_chuva_municipio(
        uf_selecionado, cidade_input, linha_mun['latitude'], linha_mun['longitude'], usar_chuva_online
    )
    st.sidebar.caption(f"Coordenadas IBGE: {linha_mun['latitude']:.4f}, {linha_mun['longitude']:.4f}")
elif not df_chuvas_completo.empty:
    estados = sorted(df_chuvas_completo['uf'].unique())
    uf_selecionado = st.sidebar.selectbox("Estado (UF):", estados, index=0, help="Dado de localização. Filtra a base de dados meteorológicos.")
    cidades = sorted(df_chuvas_completo[df_chuvas_completo['uf'] == uf_selecionado]['cidade'].unique())
    cidade_input = st.sidebar.selectbox("Município:", cidades, index=0, help="Define o histórico pluviométrico oficial que será usado nos cálculos de captação de chuva.")
    dados_cidade = df_chuvas_completo[(df_chuvas_completo['uf'] == uf_selecionado) & (df_chuvas_completo['cidade'] == cidade_input)].copy()
    fonte_chuva = "Histórico municipal (arquivo CSV)"

if dados_cidade is not None and not dados_cidade.empty:
    meses_map = {"Janeiro": "Jan", "Fevereiro": "Fev", "Março": "Mar", "Abril": "Abr", "Maio": "Mai", "Junho": "Jun", 
                 "Julho": "Jul", "Agosto": "Ago", "Setembro": "Set", "Outubro": "Out", "Novembro": "Nov", "Dezembro": "Dez"}
    dados_cidade['Mês'] = dados_cidade['mes'].map(meses_map)
    dados_cidade = dados_cidade.rename(columns={'chuva_historica_mm_mes_estimada': 'Chuva (mm)'})
    df_clima = dados_cidade[['Mês', 'Chuva (mm)']]
    
    chuva_total_anual = float(dados_cidade['Chuva (mm)'].sum())
    chuva_media_mensal = float(dados_cidade['Chuva (mm)'].mean())
    idx_min = dados_cidade['Chuva (mm)'].idxmin()
    idx_max = dados_cidade['Chuva (mm)'].idxmax()
    mes_mais_seco = dados_cidade.loc[idx_min, 'Mês']
    chuva_mes_seco = float(dados_cidade.loc[idx_min, 'Chuva (mm)'])
    mes_mais_chuvoso = dados_cidade.loc[idx_max, 'Mês']
    chuva_mes_chuvoso = float(dados_cidade.loc[idx_max, 'Chuva (mm)'])
    
    st.sidebar.success(f"**{cidade_input} / {uf_selecionado}**\n\n• Média Mensal: **{chuva_media_mensal:.1f} mm/mês**\n• Total Anual: **{chuva_total_anual:.1f} mm/ano**\n\n• Fonte: {fonte_chuva}")
    if nota_fonte_chuva:
        st.sidebar.warning(nota_fonte_chuva)
else:
    if df_catalogo.empty and df_chuvas_completo.empty:
        cidade_input = "Teresina"
        uf_selecionado = "PI"
    df_clima = None
    chuva_total_anual = 1200.0
    chuva_media_mensal = 100.0
    mes_mais_seco = "Ago"
    chuva_mes_seco = 6.5
    mes_mais_chuvoso = "Mar"
    chuva_mes_chuvoso = 260.0
    st.sidebar.warning(nota_fonte_chuva or "Sem dados de chuva para o município; usando valores padrão.")

col_amb1, col_amb2 = st.sidebar.columns(2)
altitude_local = col_amb1.number_input("Altitude (m):", min_value=0.0, value=120.0, step=20.0, help="Dado topográfico. Usado para calcular a pressão atmosférica exata da região, o que afeta diretamente o cálculo de cavitação (NPSH) da bomba.")
temp_ambiente = col_amb2.number_input("Temp. Máx (°C):", min_value=15.0, value=35.0, step=1.0, help="Dado climático. O calor dilata e enfraquece o PVC. Usado para calcular o fator de perda de resistência (derating térmico) da tubulação.")

st.sidebar.header("2. Terreno e Cultura")
if not df_matriz_culturas.empty:
    lista_culturas = df_matriz_culturas['Cultura'].tolist()
else:
    lista_culturas = ["Alface", "Tomate", "Milho", "Feijão-caupi", "Uva (videira)", "Melancia"]

cultura = st.sidebar.selectbox("Cultura Agrícola:", lista_culturas, index=0, help="Dado agronômico. Define o coeficiente da cultura (Kc), que determina a necessidade diária de água da planta.")
tipo_solo = st.sidebar.selectbox("Textura do Solo:", ["Arenoso", "Franco (médio)", "Argiloso"], index=1, help="Dado pedológico (obtido por laboratório ou análise tátil). Define o quão rápido o solo consegue absorver a água (VIB).")
profundidade_raiz = st.sidebar.slider("Profundidade Efetiva da Raiz (cm):", 10.0, 100.0, 30.0, 5.0, help="Dado agronômico. Determina o volume útil de terra que precisa ser molhado.")

area_plantio = st.sidebar.number_input("Área Total da Plantação (m²):", min_value=10.0, value=500.0, step=50.0, help="Dado de campo. Usado para calcular a demanda hídrica total diária.", key="area")

if not df_matriz_culturas.empty:
    try:
        recomendacao_excel = df_matriz_culturas.loc[df_matriz_culturas['Cultura'] == cultura, tipo_solo].values[0]
    except Exception:
        recomendacao_excel = "Gotejamento localizado sugerido."
else:
    recomendacao_excel = "Gotejamento localizado sugerido."

def inferir_metodo(texto):
    t = str(texto).lower()
    if 'microaspers' in t: return 'Microaspersão'
    elif 'aspers' in t or 'pivô' in t: return 'Aspersão'
    return 'Gotejamento'

metodo_sugerido = inferir_metodo(recomendacao_excel)
opcoes_metodo = ["Gotejamento", "Microaspersão", "Aspersão"]
idx_sugerido = opcoes_metodo.index(metodo_sugerido) if metodo_sugerido in opcoes_metodo else 0
metodo_selecionado = st.sidebar.selectbox("Método de Irrigação:", opcoes_metodo, index=idx_sugerido, help="Decisão de projeto. Impacta diretamente na eficiência de aplicação (ex: gotejamento tem eficiência de 90%, aspersão 75%).", key="metodo")

# Base da informação (Embrapa x regra geral) para a combinação cultura/solo selecionada
base_info_cultura = ""
grupo_cultura = ""
if not df_matriz_culturas.empty and cultura in df_matriz_culturas['Cultura'].values:
    _linha_cultura = df_matriz_culturas[df_matriz_culturas['Cultura'] == cultura].iloc[0]
    base_info_cultura = str(_linha_cultura.get('Base da informação', ''))
    grupo_cultura = str(_linha_cultura.get('Grupo', ''))
    st.sidebar.caption(f"Método sugerido pela matriz: **{metodo_sugerido}** · Grupo: {grupo_cultura} · Base: {base_info_cultura}")

st.sidebar.header("3. Captação de Água de Chuva")
area_telhado = st.sidebar.number_input("Área de Telhado (m²):", min_value=0.0, value=150.0, step=10.0, help="Dado arquitetônico. Área plana de captação que receberá a chuva.")
tipo_telhado = st.sidebar.selectbox(
    "Tipo de Cobertura / Telha:",
    ["Telha Cerâmica ou Fibrocimento (Runoff 0.85)", 
     "Telha Metálica ou Aluzinco (Runoff 0.90)", 
     "Laje de Concreto Impermeabilizada (Runoff 0.80)"],
    help="Dado construtivo. Determina o coeficiente de escoamento (Runoff). Telhas lisas perdem menos água por absorção."
)
dict_runoff = {
    "Telha Cerâmica ou Fibrocimento (Runoff 0.85)": 0.85,
    "Telha Metálica ou Aluzinco (Runoff 0.90)": 0.90,
    "Laje de Concreto Impermeabilizada (Runoff 0.80)": 0.80
}
c_runoff = dict_runoff[tipo_telhado]

st.sidebar.header("4. Topografia e Rede")
distancia_adutora = st.sidebar.number_input("Distância do Reservatório até a Roça (m):", min_value=1.0, value=80.0, step=5.0, help="Dado topográfico. Comprimento linear da tubulação principal. Afeta diretamente a perda de carga (atrito) do sistema.", key="dist_ad")
desnivel_recalque = st.sidebar.number_input("Desnível da Bomba até a Roça - Z (m):", min_value=0.0, value=8.0, step=1.0, help="Dado topográfico. Diferença de altura. É a força bruta contra a gravidade que a bomba precisa fazer.", key="desn_rec")
altura_succao = st.sidebar.number_input("Altura de Sucção da Bomba (m):", min_value=0.0, value=2.0, step=0.5, help="Distância vertical do nível da água no reservatório até o eixo da bomba. Criticamente usado no cálculo de NPSH.", key="alt_suc")
tarifa_agua = st.sidebar.number_input("Tarifa Concessionária (R$/m³):", min_value=0.0, value=5.50, step=0.5, help="Dado tarifário local (fatura de água). Usado para calcular o ROI e o balanço financeiro.")

st.sidebar.header("5. Visualização 2D / 3D")
setor_visual = st.sidebar.selectbox("Setor em operação (esquemáticos e 3D):", ["Todos os setores", "Setor 1", "Setor 2", "Setor 3", "Setor 4", "Setor 5", "Setor 6"], index=0,
                                    help="Só afeta o desenho: a válvula do setor escolhido fica aberta (verde), as demais fecham (vermelho) e só o setor em operação aparece molhado.")
setor_ativo_vis = 0 if setor_visual.startswith("Todos") else int(setor_visual.split()[-1])

st.sidebar.header("6. Entrada Livre de Dados")
with st.sidebar.expander("Sobrescrever valores automáticos", expanded=False):
    st.caption("Por padrão a ferramenta calcula ou sugere cada valor. Marque somente o que você quer substituir, por exemplo para reproduzir um projeto já existente (back-test) com os dados dele.")
    ov_eto_on = st.checkbox("Informar ETo (mm/dia)", key="ov_eto_on")
    ov_eto = st.number_input("ETo (mm/dia):", min_value=0.5, max_value=15.0, value=5.5, step=0.1, key="ov_eto", disabled=not ov_eto_on)
    ov_kc_on = st.checkbox("Informar Kc da cultura", key="ov_kc_on")
    ov_kc = st.number_input("Kc:", min_value=0.1, max_value=1.6, value=1.0, step=0.05, key="ov_kc", disabled=not ov_kc_on)
    ov_ef_on = st.checkbox("Informar eficiência de aplicação (%)", key="ov_ef_on")
    ov_ef = st.number_input("Eficiência (%):", min_value=30.0, max_value=99.0, value=75.0, step=1.0, key="ov_ef", disabled=not ov_ef_on)
    ov_vib_on = st.checkbox("Informar VIB do solo (mm/h)", key="ov_vib_on")
    ov_vib = st.number_input("VIB (mm/h):", min_value=0.5, max_value=300.0, value=12.0, step=0.5, key="ov_vib", disabled=not ov_vib_on)
    ov_cad_on = st.checkbox("Informar CAD do solo (mm por cm de solo)", key="ov_cad_on")
    ov_cad = st.number_input("CAD (mm/cm):", min_value=0.3, max_value=4.0, value=1.5, step=0.1, key="ov_cad", disabled=not ov_cad_on)
    ov_turno_on = st.checkbox("Informar turno de rega (dias)", key="ov_turno_on")
    ov_turno = st.number_input("Turno de rega (dias):", min_value=1, max_value=30, value=1, step=1, key="ov_turno", disabled=not ov_turno_on)

# ==============================================================================
# BANCO DE DADOS AGRONÔMICOS E COERÊNCIA FÍSICA AUTOMÁTICA
# ==============================================================================
banco_culturas_fao = {
    'Milho': {'kc': 1.10, 'eficiencia': 0.85}, 'Feijão-caupi': {'kc': 1.05, 'eficiencia': 0.85},
    'Uva (videira)': {'kc': 0.80, 'eficiencia': 0.90}, 'Tomate': {'kc': 1.15, 'eficiencia': 0.90},
    'Melão': {'kc': 1.00, 'eficiencia': 0.90}, 'Melancia': {'kc': 1.00, 'eficiencia': 0.90},
    'Cebola': {'kc': 1.05, 'eficiencia': 0.90}, 'Hortaliças folhosas (alface, couve)': {'kc': 0.95, 'eficiencia': 0.85},
    'Fruteiras (citros, mamão, goiaba)': {'kc': 0.75, 'eficiencia': 0.90}, 'Arroz': {'kc': 1.15, 'eficiencia': 0.60},
    'Soja': {'kc': 1.15, 'eficiencia': 0.85}, 'Algodão': {'kc': 1.15, 'eficiencia': 0.85},
    'Cana-de-açúcar': {'kc': 1.20, 'eficiencia': 0.80}, 'Cacau': {'kc': 1.00, 'eficiencia': 0.90},
    'Mandioca': {'kc': 0.80, 'eficiencia': 0.85}, 'Banana': {'kc': 1.10, 'eficiencia': 0.90},
    'Manga': {'kc': 0.85, 'eficiencia': 0.90}, 'Coco-da-baía': {'kc': 0.90, 'eficiencia': 0.90}
}
banco_cad_solos = {"Arenoso": 1.0, "Franco (médio)": 1.5, "Argiloso": 2.0} # mm/cm
banco_vib_solos = {"Arenoso": 25.0, "Franco (médio)": 12.0, "Argiloso": 5.0} # mm/h 
# VIB de referência Embrapa (planilha "Infiltração"): limite inferior conservador de cada grupo hidrológico
# Arenoso = grupo A (> 190 mm/h) | Franco = grupo B (40 a 190 mm/h) | Argiloso = grupo D (< 3 mm/h)
banco_vib_embrapa_conservador = {"Arenoso": 190.0, "Franco (médio)": 40.0, "Argiloso": 3.0} # mm/h

dados_fao = banco_culturas_fao.get(cultura, {'kc': 1.0, 'eficiencia': 0.85})
kc_cultura = float(ov_kc) if ov_kc_on else dados_fao['kc']
eto_media = float(ov_eto) if ov_eto_on else 5.5 # mm/dia 
etc_cultura = eto_media * kc_cultura
cad_solo = float(ov_cad) if ov_cad_on else banco_cad_solos[tipo_solo]
vib_solo_padrao = float(ov_vib) if ov_vib_on else banco_vib_solos[tipo_solo]

if metodo_selecionado == "Gotejamento":
    eficiencia_sistema = 0.90
    qa_sugerido = 2.0 # L/h
    se_sugerido = 0.5 # m
    sl_sugerido = 1.0 # m
    pn_linha_padrao = 30.0 # mca
elif metodo_selecionado == "Microaspersão":
    eficiencia_sistema = 0.85
    qa_sugerido = 35.0 # L/h
    se_sugerido = 2.0 # m
    sl_sugerido = 3.0 # m
    pn_linha_padrao = 40.0 # mca
else:
    eficiencia_sistema = 0.75
    qa_sugerido = 450.0 # L/h
    se_sugerido = 12.0 # m
    sl_sugerido = 12.0 # m
    pn_linha_padrao = 50.0 # mca

if ov_ef_on:
    eficiencia_sistema = float(ov_ef) / 100.0
lamina_bruta_diaria = etc_cultura / eficiencia_sistema
consumo_diario_total = lamina_bruta_diaria * area_plantio # Litros/dia
reserva_util_solo = cad_solo * profundidade_raiz * 0.5 # mm
turno_rega = int(ov_turno) if ov_turno_on else (max(1, math.floor(reserva_util_solo / etc_cultura)) if etc_cultura > 0 else 1)

lado_campo_quadrado = math.sqrt(area_plantio)
comp_linha_lateral = round(lado_campo_quadrado, 1) 
largura_campo = area_plantio / comp_linha_lateral if comp_linha_lateral > 0 else 1

# ==============================================================================
# FUNÇÕES DE DESENHO VETORIAL (SVG)
# ==============================================================================
def gerar_svg_lamina(lamina_mm):
    altura_agua = min(80, max(10, lamina_mm * 5))
    svg = f"""
    <div style="display: flex; justify-content: center; margin: 20px 0;">
        <svg width="400" height="220" viewBox="0 0 400 220" style="border: 1px solid #ddd; background-color: #FDFEFE; border-radius: 5px;">
            <!-- Solo -->
            <rect x="50" y="140" width="300" height="80" fill="#8D6E63" />
            <!-- Raiz e Planta -->
            <path d="M200,140 Q180,180 200,210 M200,140 Q220,170 190,200" fill="none" stroke="#5D4037" stroke-width="3"/>
            <path d="M200,140 Q170,100 190,60 Q200,80 200,140 Z" fill="#27AE60"/>
            <path d="M200,140 Q230,110 210,70 Q200,90 200,140 Z" fill="#2ECC71"/>
            <!-- Camada de Água (Lâmina Bruta) -->
            <rect x="50" y="{140 - altura_agua}" width="300" height="{altura_agua}" fill="#3498DB" opacity="0.6" />
            <!-- Cotas e Linhas -->
            <line x1="30" y1="{140 - altura_agua}" x2="45" y2="{140 - altura_agua}" stroke="#333" stroke-width="1"/>
            <line x1="30" y1="140" x2="45" y2="140" stroke="#333" stroke-width="1"/>
            <line x1="37" y1="{140 - altura_agua}" x2="37" y2="140" stroke="#333" stroke-width="1"/>
            <text x="5" y="{140 - (altura_agua/2) + 5}" font-family="Arial" font-size="12" font-weight="bold" fill="#333">{lamina_mm:.1f} mm</text>
            <text x="60" y="30" font-family="Arial" font-size="12" font-weight="bold" fill="#2C3E50">Representação visual da Lâmina Bruta</text>
            <text x="60" y="45" font-family="Arial" font-size="10" fill="#7F8C8D">1 mm de altura equivale a 1 Litro de água espalhado em 1 m²</text>
        </svg>
    </div>
    """
    return svg
def gerar_svg_cisterna():
    svg = """
    <div style="display: flex; justify-content: center; margin: 20px 0;">
    <svg width="100%" viewBox="0 0 800 480" style="max-width: 900px; border: 1px solid #ddd; background-color: #FDFEFE; border-radius: 5px;" xmlns="http://www.w3.org/2000/svg">
        <text x="400" y="22" font-family="Arial" font-size="15" font-weight="bold" fill="#2C3E50" text-anchor="middle">Esquema do Sistema: da Casa até a Roça (Cisterna Enterrada)</text>

        <!-- Subsolo -->
        <rect x="0" y="250" width="800" height="230" fill="#8D6E63" opacity="0.35"/>
        <line x1="0" y1="250" x2="800" y2="250" stroke="#5D4037" stroke-width="3"/>
        <text x="8" y="268" font-family="Arial" font-size="10" font-weight="bold" fill="#5D4037">NÍVEL DO SOLO</text>
        <text x="8" y="470" font-family="Arial" font-size="10" font-weight="bold" fill="#5D4037">SUBSOLO</text>

        <!-- Chuva -->
        <ellipse cx="110" cy="30" rx="40" ry="12" fill="#BDC3C7"/>
        <ellipse cx="88" cy="35" rx="25" ry="10" fill="#BDC3C7"/>
        <ellipse cx="135" cy="35" rx="25" ry="10" fill="#BDC3C7"/>
        <line x1="85" y1="48" x2="79" y2="64" stroke="#3498DB" stroke-width="2"/>
        <line x1="110" y1="48" x2="104" y2="64" stroke="#3498DB" stroke-width="2"/>
        <line x1="135" y1="48" x2="129" y2="64" stroke="#3498DB" stroke-width="2"/>

        <!-- Casa -->
        <rect x="30" y="140" width="160" height="110" fill="#ECF0F1" stroke="#2C3E50" stroke-width="3"/>
        <polygon points="15,140 110,75 205,140" fill="#C0392B" stroke="#2C3E50" stroke-width="3"/>
        <rect x="85" y="190" width="30" height="60" fill="#8D6E63" stroke="#2C3E50" stroke-width="2"/>
        <rect x="45" y="165" width="28" height="28" fill="#AED6F1" stroke="#2C3E50" stroke-width="2"/>
        <rect x="147" y="165" width="28" height="28" fill="#AED6F1" stroke="#2C3E50" stroke-width="2"/>
        <text x="110" y="240" font-family="Arial" font-size="10" font-weight="bold" fill="#2C3E50" text-anchor="middle">CASA</text>
        <text x="110" y="100" font-family="Arial" font-size="10" font-weight="bold" fill="#FFFFFF" text-anchor="middle">1. Telhado</text>

        <!-- Calha e condutor vertical (gravidade) -->
        <line x1="200" y1="138" x2="215" y2="138" stroke="#7F8C8D" stroke-width="6"/>
        <line x1="215" y1="138" x2="215" y2="300" stroke="#7F8C8D" stroke-width="6"/>
        <text x="225" y="128" font-family="Arial" font-size="10" font-weight="bold" fill="#2C3E50">2. Calha e condutor</text>

        <!-- Descarte da primeira água (derivação vertical) -->
        <line x1="215" y1="155" x2="275" y2="155" stroke="#7F8C8D" stroke-width="5"/>
        <rect x="268" y="155" width="14" height="90" fill="#BDC3C7" stroke="#7F8C8D" stroke-width="2"/>
        <line x1="275" y1="245" x2="275" y2="258" stroke="#7F8C8D" stroke-width="3"/>
        <text x="292" y="200" font-family="Arial" font-size="10" font-weight="bold" fill="#2C3E50">3. Descarte</text>
        <text x="292" y="213" font-family="Arial" font-size="10" font-weight="bold" fill="#2C3E50">da 1ª água</text>

        <!-- Filtro vertical -->
        <rect x="198" y="176" width="34" height="56" fill="#F5CBA7" stroke="#7F8C8D" stroke-width="2" rx="4"/>
        <text x="215" y="204" font-family="Arial" font-size="8" font-weight="bold" fill="#2C3E50" text-anchor="middle" transform="rotate(-90, 215, 204)">4. FILTRO</text>

        <!-- Cisterna enterrada -->
        <rect x="170" y="275" width="240" height="180" fill="#FDFEFE" stroke="#2C3E50" stroke-width="4" rx="4"/>
        <rect x="172" y="320" width="236" height="133" fill="#3498DB" opacity="0.5"/>
        <rect x="290" y="263" width="40" height="12" fill="#95A5A6" stroke="#2C3E50" stroke-width="2"/>
        <text x="310" y="258" font-family="Arial" font-size="9" fill="#2C3E50" text-anchor="middle">Tampa de inspeção</text>
        <text x="290" y="440" font-family="Arial" font-size="12" font-weight="bold" fill="#1B4F72" text-anchor="middle">5. CISTERNA</text>

        <!-- Nível crítico -->
        <line x1="165" y1="410" x2="415" y2="410" stroke="#E74C3C" stroke-width="2" stroke-dasharray="8,4"/>
        <text x="330" y="404" font-family="Arial" font-size="10" font-weight="bold" fill="#E74C3C">Nível crítico (abre a rede pública)</text>

        <!-- Extravasor (ladrão) -->
        <line x1="170" y1="295" x2="135" y2="295" stroke="#7F8C8D" stroke-width="4"/>
        <line x1="135" y1="295" x2="135" y2="340" stroke="#7F8C8D" stroke-width="4"/>
        <text x="128" y="355" font-family="Arial" font-size="10" fill="#7F8C8D" text-anchor="middle">Extravasor</text>
        <text x="128" y="367" font-family="Arial" font-size="10" fill="#7F8C8D" text-anchor="middle">(ladrão)</text>

        <!-- Sucção afogada e bomba na lateral -->
        <line x1="410" y1="345" x2="440" y2="345" stroke="#34495E" stroke-width="6"/>
        <rect x="440" y="322" width="50" height="46" fill="#34495E" rx="6"/>
        <text x="465" y="349" font-family="Arial" font-size="10" font-weight="bold" fill="#FFFFFF" text-anchor="middle">BOMBA</text>
        <text x="465" y="384" font-family="Arial" font-size="10" fill="#34495E" text-anchor="middle">Sucção afogada</text>

        <!-- Recalque até a roça -->
        <line x1="465" y1="322" x2="465" y2="228" stroke="#27AE60" stroke-width="6"/>
        <line x1="465" y1="228" x2="715" y2="228" stroke="#27AE60" stroke-width="6"/>
        <polygon points="715,218 735,228 715,238" fill="#27AE60"/>
        <text x="590" y="218" font-family="Arial" font-size="11" font-weight="bold" fill="#1E8449" text-anchor="middle">Recalque</text>

        <!-- Roça -->
        <rect x="738" y="232" width="55" height="18" fill="#6D4C41"/>
        <path d="M748,232 Q744,214 752,206 Q756,220 748,232 Z" fill="#27AE60"/>
        <path d="M766,232 Q762,210 772,200 Q776,218 766,232 Z" fill="#2ECC71"/>
        <path d="M783,232 Q779,216 787,208 Q791,222 783,232 Z" fill="#27AE60"/>
        <text x="765" y="196" font-family="Arial" font-size="11" font-weight="bold" fill="#2C3E50" text-anchor="middle">ROÇA</text>
    </svg>
    </div>
    """
    return "".join(linha.strip() for linha in svg.splitlines())


# ==============================================================================
# CONFIGURAÇÃO DAS ABAS 
# ==============================================================================
tab_captacao, tab_planta, tab_hidraulica, tab_tubulacao, tab_processador, tab_eletrica, tab_planta_3d, tab_laudo, tab_referencias, tab_economia, tab_backtest = st.tabs([
    "1. Água de Chuva",
    "2. Planta e Solo",
    "3. Tubulações e Bomba",
    "4. Esquemático da Tubulação",
    "5. Processador Hidráulico",
    "6. Instalação Elétrica",
    "7. Planta Baixa e 3D",
    "8. Memorial e Laudo",
    "9. Referências Técnicas",
    "10. Balanço Econômico",
    "11. Back-test (Artigo × Ferramenta)"
])

# ------------------------------------------------------------------------------
# 1. ABA: ÁGUA DE CHUVA
# ------------------------------------------------------------------------------
with tab_captacao:
    st.subheader("Sistema de Captação e Aproveitamento de Água Pluvial")
    
    opcao_regime = st.radio(
        "Cenário de Dimensionamento Pluvial (Base Oficial 100% Ancorada):",
        [f"Média Mensal do Município ({chuva_media_mensal:.1f} mm)",
         f"Mês Mais Seco / Estiagem ({mes_mais_seco}: {chuva_mes_seco:.1f} mm)",
         f"Mês Mais Chuvoso ({mes_mais_chuvoso}: {chuva_mes_chuvoso:.1f} mm)"],
        index=0,
        horizontal=True
    )
    if "Média Mensal" in opcao_regime:
        precipitacao_referencia = chuva_media_mensal
        tipo_cenario_txt = "Média Mensal Histórica"
    elif "Mais Seco" in opcao_regime:
        precipitacao_referencia = chuva_mes_seco
        tipo_cenario_txt = f"Estiagem Crítica ({mes_mais_seco})"
    else:
        precipitacao_referencia = chuva_mes_chuvoso
        tipo_cenario_txt = f"Pico de Chuva ({mes_mais_chuvoso})"

    col_c1, col_c2, col_c3 = st.columns(3)
    col_c1.metric("Precipitação de Cálculo", f"{precipitacao_referencia:.1f} mm", tipo_cenario_txt)
    taxa_ff = col_c2.number_input("Descarte da Primeira Água (L/m² telhado):", value=2.0, min_value=1.0, max_value=4.0, step=0.5, help="Define o volume descartado para lavar sujeiras do telhado antes do armazenamento.")
    vol_cisterna_existente = col_c3.number_input("Volume da Cisterna Prevista (Litros):", value=10000.0, step=1000.0, help="Volume nominal comercial da caixa d'água pretendida.")
    
    col_c4, col_c5 = st.columns(2)
    nivel_boia_backup = col_c4.slider("Nível Mínimo da Cisterna para abrir Rede Pública (%):", 10, 40, 20, step=5, help="Garante que a roça nunca fique sem água se a chuva faltar.")
    vazao_rede_concessionaria = col_c5.number_input("Vazão de saída da Torneira/Rede Pública para a cisterna (L/h):", value=1000.0, min_value=0.0, step=100.0, help="Informe a vazão medida da torneira da rua que abastece a cisterna na falta de chuva. Aparece no esquemático, na entrada lateral esquerda da cisterna.")
    
    # Cálculos Pluviais
    vol_bruto_captado = area_telhado * precipitacao_referencia * c_runoff
    vol_first_flush = area_telhado * taxa_ff
    vol_liquido_aproveitavel = max(0.0, vol_bruto_captado - vol_first_flush)
    
    autonomia_dias = (vol_liquido_aproveitavel / consumo_diario_total) if consumo_diario_total > 0 else 0
    vol_reserva_backup = vol_cisterna_existente * (nivel_boia_backup / 100.0)
    tempo_encher_cisterna_horas = ((vol_cisterna_existente - vol_reserva_backup) / vazao_rede_concessionaria) if vazao_rede_concessionaria > 0 else 0
    
    # --- Cálculos adicionais do Motor de Cálculo (aba "4. Fonte Dupla e Pluvial") ---
    demanda_ciclo_litros = consumo_diario_total * turno_rega  # Demanda hídrica por ciclo de irrigação (L)
    autonomia_ciclos = (vol_liquido_aproveitavel / demanda_ciclo_litros) if demanda_ciclo_litros > 0 else 0
    excede_cisterna = vol_liquido_aproveitavel > vol_cisterna_existente
    status_transbordo = "RISCO: Necessita Ladrão Ativo" if excede_cisterna else "CAPACIDADE OK"
    demanda_horaria_media = consumo_diario_total / 24.0  # L/h médios que a rede pública precisa repor
    rede_lenta = vazao_rede_concessionaria < demanda_horaria_media
    status_falta_agua = "ALERTA: Rede pública lenta para o ciclo" if rede_lenta else "SEGURO"
    reposicao_lenta = tempo_encher_cisterna_horas > (turno_rega * 24.0)
    
    if area_telhado <= 50: diam_calha, diam_condutor = "100 mm", "75 mm"
    elif area_telhado <= 120: diam_calha, diam_condutor = "150 mm", "100 mm"
    else: diam_calha, diam_condutor = "200 mm", "150 mm"
    
    st.markdown("#### Balanço do Mês de Referência")
    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Chuva Total Captada", f"{vol_bruto_captado:,.0f} L")
    mc2.metric("Descarte First-Flush", f"{vol_first_flush:,.0f} L")
    mc3.metric("Água Limpa Aproveitável", f"{vol_liquido_aproveitavel:,.0f} L")
    mc4.metric("Autonomia de Irrigação", f"{autonomia_dias:.1f} dias")
    
    st.markdown("#### Dupla Fonte, Reposição e Transbordo")
    mc5, mc6, mc7, mc8 = st.columns(4)
    mc5.metric("Demanda por Ciclo de Irrigação", f"{demanda_ciclo_litros:,.0f} L", f"Turno de {turno_rega} dia(s)")
    mc6.metric("Autonomia Pluvial", f"{autonomia_ciclos:.1f} ciclos")
    mc7.metric("Volume de Backup (abre rede pública)", f"{vol_reserva_backup:,.0f} L")
    mc8.metric("Tempo de Reposição pela Rede", f"{tempo_encher_cisterna_horas:.1f} h")
    
    if excede_cisterna:
        st.warning(f"Status de Transbordo: {status_transbordo}. O volume líquido do período ({vol_liquido_aproveitavel:,.0f} L) supera a cisterna ({vol_cisterna_existente:,.0f} L).")
    else:
        st.success(f"Status de Transbordo: {status_transbordo}. O volume líquido cabe na cisterna prevista.")
    if rede_lenta:
        st.warning(f"Risco de Falta d'Água: {status_falta_agua}. A rede ({vazao_rede_concessionaria:,.0f} L/h) não cobre a demanda média diária ({demanda_horaria_media:,.0f} L/h).")
    else:
        st.success(f"Risco de Falta d'Água: {status_falta_agua}. A rede pública ({vazao_rede_concessionaria:,.0f} L/h) cobre a demanda média diária ({demanda_horaria_media:,.0f} L/h).")
    if reposicao_lenta:
        st.warning(f"A reposição da cisterna pela rede ({tempo_encher_cisterna_horas:.1f} h) é mais lenta que o turno de rega ({turno_rega * 24} h).")
    
    st.markdown("### Esquema do Reservatório Híbrido")
    # O desenho realista (casa, cisterna, bomba e roça com setores) é inserido aqui logo após a aba 4,
    # porque usa a bomba (aba 3) e o tubo (aba 4). A função antiga gerar_svg_cisterna() foi mantida no código.

# ------------------------------------------------------------------------------
# 2. ABA: PLANTA E SOLO 
# ------------------------------------------------------------------------------
with tab_planta:
    st.subheader(f"Demanda de Água da Cultura: {cultura} em Solo {tipo_solo}")
    
    col_ag1, col_ag2, col_ag3 = st.columns(3)
    if st.session_state.get("_met_prev") != metodo_selecionado:
        for _k in ("q_emissor", "s_emissor", "s_linhas"):
            st.session_state.pop(_k, None)
        st.session_state["_met_prev"] = metodo_selecionado
    vazao_emissor_campo = col_ag1.number_input("Vazão de Cada Emissor (L/h):", value=qa_sugerido, step=0.5, help="Dado de catálogo. A quantidade exata de água que sai de um único furo/aspersor.", key="q_emissor")
    espacamento_emissores = col_ag2.number_input("Espaçamento entre Emissores na Linha (m):", value=se_sugerido, step=0.1, help="Dado de campo/projeto agronômico.", key="s_emissor")
    espacamento_linhas = col_ag3.number_input("Espaçamento entre Fileiras/Linhas (m):", value=sl_sugerido, step=0.1, help="Distância entre as mangueiras de irrigação na roça.", key="s_linhas")
    
    escala_vib = st.radio(
        "Referência de VIB para a verificação de infiltração:",
        ["Classes de irrigação (padrão do projeto)", "Embrapa conservador (limite inferior do grupo hidrológico)"],
        index=0, horizontal=True,
        help="As duas escalas servem a fins diferentes (ver aba Referências). A padrão usa as classes de irrigação (Bernardo); a conservadora usa o limite inferior do grupo hidrológico medido pela Embrapa."
    )
    if ov_vib_on:
        vib_solo = float(ov_vib)  # VIB informada pelo usuário tem prioridade sobre as tabelas
    elif escala_vib.startswith("Embrapa"):
        vib_solo = float(banco_vib_embrapa_conservador[tipo_solo])
    else:
        vib_solo = float(vib_solo_padrao)
    taxa_aplicacao_real = vazao_emissor_campo / (espacamento_emissores * espacamento_linhas) if (espacamento_emissores * espacamento_linhas) > 0 else 0
    risco_escorrimento = taxa_aplicacao_real > vib_solo
    
    st.markdown("#### Indicadores de Manejo")
    ma1, ma2, ma3, ma4 = st.columns(4)
    ma1.metric("Necessidade da Cultura (ETc)", f"{etc_cultura:.2f} mm/dia")
    ma2.metric("Lâmina Bruta Aplicada", f"{lamina_bruta_diaria:.2f} mm/dia")
    ma3.metric("Consumo Diário Total da Roça", f"{consumo_diario_total:,.0f} Litros/dia")
    ma4.metric("Turno de Rega Máximo", f"A cada {turno_rega} dia(s)")
    
    cfg_solo = dict(cultura=cultura, tipo_solo=tipo_solo, prof_raiz=profundidade_raiz, cad_mm_cm=cad_solo, etc=etc_cultura,
                    lamina_bruta=lamina_bruta_diaria, turno=turno_rega, reserva_util=reserva_util_solo,
                    arborea=_cultura_arborea(cultura), metodo=metodo_selecionado)
    st.markdown("#### Perfil do solo, raízes e balanço hídrico")
    st.markdown(gerar_svg_perfil_solo(cfg_solo), unsafe_allow_html=True)
    with st.expander("Representação simples da lâmina bruta (versão anterior)"):
        st.markdown(gerar_svg_lamina(lamina_bruta_diaria), unsafe_allow_html=True)
    
    ma5, ma6, ma7 = st.columns(3)
    ma5.metric("Taxa de Aplicação Real (Ta)", f"{taxa_aplicacao_real:.2f} mm/h")
    ma6.metric("VIB Estática", f"{vib_solo:.1f} mm/h", f"Solo {tipo_solo}")
    status_txt = "RISCO: Escorrimento Superficial" if risco_escorrimento else "SEGURO: Solo absorve 100%"
    ma7.metric("Segurança da Infiltração", status_txt)
    
    if risco_escorrimento:
        st.error(f"Incompatibilidade Física Solo x Emissor: A taxa de água ({taxa_aplicacao_real:.2f} mm/h) é maior que a velocidade de infiltração do seu solo ({vib_solo:.1f} mm/h). Haverá poças e escorrimento de água.")
    else:
        st.success(f"Taxa de Aplicação Aprovada: A água penetra suavemente no solo sem escorrer superficialmente.")
    
    # --- Demanda líquida x bruta (Motor de Cálculo, aba "5. Agro e Elétrica") ---
    demanda_liquida_total = etc_cultura * area_plantio  # L/dia (1 mm = 1 L/m²)
    demanda_bruta_corrigida = demanda_liquida_total / eficiencia_sistema
    md1, md2, md3 = st.columns(3)
    md1.metric("Demanda Hídrica Líquida da Roça", f"{demanda_liquida_total:,.0f} L/dia")
    md2.metric("Eficiência de Aplicação (Ea)", f"{eficiencia_sistema * 100:.0f} %", metodo_selecionado)
    md3.metric("Demanda Bruta Corrigida (c/ perdas)", f"{demanda_bruta_corrigida:,.0f} L/dia")
    
    # --- Ficha agronômica com a base Embrapa (planilha irrigacao_solo_cultura) ---
    with st.expander("Ficha Agronômica da Cultura e do Solo (base Embrapa)", expanded=False):
        if not df_matriz_culturas.empty and cultura in df_matriz_culturas['Cultura'].values:
            linha_cult = df_matriz_culturas[df_matriz_culturas['Cultura'] == cultura].iloc[0]
            base_txt = str(linha_cult.get('Base da informação', '-'))
            st.markdown(f"**Grupo:** {linha_cult.get('Grupo', '-')}  |  **Base da informação:** {base_txt}")
            if 'Embrapa' in base_txt:
                st.caption("Esta cultura tem base direta em publicações da Embrapa (métodos); a associação com a textura combina essas fontes com regras agronômicas gerais.")
            else:
                st.caption("Regra agronômica geral, sem base direta nos trechos consultados da Embrapa. Validar com engenheiro agrônomo.")
            st.markdown(f"**Sistema indicado para solo {tipo_solo}:** {recomendacao_excel}")
            tabela_recs = pd.DataFrame({
                'Textura do solo': ["Arenoso", "Franco (médio)", "Argiloso"],
                'Sistema indicado': [str(linha_cult.get(t, '-')) for t in ["Arenoso", "Franco (médio)", "Argiloso"]]
            })
            st.dataframe(tabela_recs, hide_index=True, use_container_width=True)
        else:
            st.info("Planilha agronômica não encontrada: ficha da cultura indisponível.")
        
        if not df_top10_ne.empty:
            linha_top = df_top10_ne[df_top10_ne['Cultura'] == cultura]
            if not linha_top.empty:
                lt = linha_top.iloc[0]
                st.markdown(f"**Relevância no Nordeste:** {int(lt['Ordem (aprox.)'])}º lugar (ordem aproximada) · {lt['Principais áreas no NE']}")
                st.caption(f"{lt['Dado econômico (fonte)']} — {lt['Observação']}")
            else:
                st.caption("Esta cultura não está entre as 10 de maior relevância econômica no Nordeste (IBGE/PAM e Sudene).")
        
        if not df_solo_carac.empty and tipo_solo in df_solo_carac.columns:
            st.markdown(f"**Comportamento hídrico do solo {tipo_solo}**")
            tabela_solo = df_solo_carac[['Característica', tipo_solo]].rename(columns={tipo_solo: f'Solo {tipo_solo}'})
            st.dataframe(tabela_solo, hide_index=True, use_container_width=True)
        
        if not df_infiltracao_embrapa.empty:
            linha_inf = df_infiltracao_embrapa[df_infiltracao_embrapa['Textura'] == tipo_solo]
            if not linha_inf.empty:
                li = linha_inf.iloc[0]
                st.markdown(f"**Velocidade de Infiltração Básica (Embrapa) — grupo hidrológico {li['Grupo hidrológico']}**")
                st.markdown(f"- Faixa de referência: {li['Faixa de referência (mm/h)']}")
                st.markdown(f"- Medições de campo: {li['Medições de campo (Embrapa)']}")
                st.markdown(f"- Observação: {li['Observação']}")
                st.caption("Em projeto, medir a VIB no local. Solos compactados ou degradados caem para grupos de menor infiltração.")

# ------------------------------------------------------------------------------
# 3. ABA: TUBULAÇÕES E BOMBA 
# ------------------------------------------------------------------------------
with tab_hidraulica:
    st.subheader("Dimensionamento Hidráulico Completo e Conjunto Motobomba")
    
    col_set1, col_set2 = st.columns(2)
    num_setores = col_set1.selectbox("Divisão da Roça em Setores de Irrigação:", [1, 2, 3, 4, 5, 6], index=0, help="Projeto hidráulico. Dividir a roça reduz o tamanho da bomba necessária.", key="num_setores")
    if st.session_state.get("_area_prev") != area_plantio:
        st.session_state.pop("comp_lat", None)
        st.session_state["_area_prev"] = area_plantio
    comp_linha_digitado = col_set2.number_input("Comprimento de Cada Linha Lateral (m):", value=float(comp_linha_lateral), step=1.0, help="Medida real da mangueira de irrigação instalada no solo.", key="comp_lat")
    
    num_emissores_por_linha = math.floor(comp_linha_digitado / espacamento_emissores) if espacamento_emissores > 0 else 1
    largura_real_campo = area_plantio / comp_linha_digitado if comp_linha_digitado > 0 else 1
    num_total_linhas_campo = max(1, math.floor(largura_real_campo / espacamento_linhas)) if espacamento_linhas > 0 else 1
    with st.expander("Entrada livre: geometria do campo (substitui o cálculo automático)", expanded=False):
        st.caption(f"Automático: {num_emissores_por_linha} emissor(es) por lateral (comprimento ÷ espaçamento, arredondado para baixo) e {num_total_linhas_campo} lateral(is) no campo "
                   f"(largura = área ÷ comprimento = {largura_real_campo:.1f} m). Em aspersão, o número de aspersores do projeto pode diferir desse arredondamento.")
        gm1, gm2 = st.columns(2)
        man_nem = gm1.checkbox("Informar o nº de emissores/aspersores por lateral", key="man_nem")
        nem_manual = gm1.number_input("Emissores por lateral:", min_value=1, value=max(1, int(num_emissores_por_linha)), step=1, disabled=not man_nem, key="nem_manual")
        man_nlin = gm2.checkbox("Informar o nº total de linhas laterais no campo", key="man_nlin")
        nlin_manual = gm2.number_input("Linhas laterais no campo:", min_value=1, value=max(1, int(num_total_linhas_campo)), step=1, disabled=not man_nlin, key="nlin_manual")
        if man_nem:
            num_emissores_por_linha = int(nem_manual)
        if man_nlin:
            num_total_linhas_campo = int(nlin_manual)
    num_linhas_por_setor = max(1, math.ceil(num_total_linhas_campo / num_setores))
    
    vazao_linha_lateral_m3h = (num_emissores_por_linha * vazao_emissor_campo) / 1000.0
    vazao_setor_operando_m3h = vazao_linha_lateral_m3h * num_linhas_por_setor
    
    with st.expander("1. Linhas Laterais no Campo", expanded=True):
        st.markdown("**Por que este cálculo existe?** Para medir quanta energia a água perde esfregando no interior da mangueira (Perda de Carga). O objetivo é garantir que a água saia com a mesma força na primeira e na última planta.")
        
        col_l1, col_l2, col_l3 = st.columns(3)
        diam_linha_int_mm = col_l1.number_input("Diâmetro Interno da Linha Lateral (mm):", value=16.0, step=1.0, help="Dado de catálogo do fabricante do tubo.", key="d_lat")
        pressao_nominal_emissor = col_l2.number_input("Pressão Nominal de Serviço do Emissor (mca):", value=10.0, step=1.0, help="Pressão recomendada no catálogo do fabricante para garantir a vazão informada.", key="pnom")
        pn_classe_tubo_lateral = col_l3.number_input("Classe de Pressão do Tubo Lateral - PN (mca):", value=pn_linha_padrao, step=5.0, help="Pressão máxima que a mangueira suporta de fábrica.")
        
        # Parâmetros do Motor de Cálculo (aba "1. Linhas Laterais"), agora editáveis (valores padrão inalterados)
        col_l4, col_l5, col_l6 = st.columns(3)
        modo_f = col_l4.selectbox("Fator F de Christiansen:", ["Automático (pelo nº de saídas)", "Manual"], key="modo_f", help="O automático calcula F pelo número de emissores da lateral (N) e pela posição do 1º emissor. Manual: digite o valor de uma tabela ou do projeto que você está reproduzindo.")
        c_rugosidade_pvc = col_l5.number_input("Coeficiente de Rugosidade C (Hazen-Williams):", value=140.0, min_value=80.0, max_value=160.0, step=5.0, help="Adimensional. PVC/PE novo: 140 a 150. Também usado na adutora.", key="c_rug")
        z_lateral = col_l6.number_input("Desnível Topográfico da Linha Lateral - Z (m):", value=0.0, step=0.5, help="Positivo para aclive, negativo para declive. Entra na pressão de entrada como 0,5·Z.")

        col_l7, col_l8, col_l9 = st.columns(3)
        pos_primeiro = col_l7.selectbox("Posição do 1º emissor na lateral:", ["A uma distância S da entrada", "A S/2 da entrada"], key="pos1", help="S é o espaçamento entre emissores. Muda o fator F quando há poucas saídas.")
        altura_riser = col_l8.number_input("Altura do tubo do aspersor sobre a lateral (m):", min_value=0.0, value=0.0, step=0.1, key="ha", help="Soma à pressão inicial da lateral (Ha). Em gotejamento, deixe 0.")
        fator_f_manual = col_l9.number_input("F manual:", min_value=0.20, max_value=1.00, value=0.36, step=0.01, key="fman", disabled=modo_f.startswith("Auto"))
        fator_f = fator_christiansen(num_emissores_por_linha, pos_primeiro.startswith("A S/2")) if modo_f.startswith("Auto") else float(fator_f_manual)
        st.caption(f"F aplicado = {fator_f:.3f} (N = {num_emissores_por_linha} saída(s) por lateral).")
        
        q_linha_m3s = vazao_linha_lateral_m3h / 3600.0
        area_secao_linha = math.pi * ((diam_linha_int_mm / 1000.0) ** 2) / 4.0
        vel_linha_lateral = (q_linha_m3s / area_secao_linha) if area_secao_linha > 0 else 0
        
        hf_linha_continua = 10.67 * comp_linha_digitado * ((q_linha_m3s / c_rugosidade_pvc) ** 1.852) / ((diam_linha_int_mm / 1000.0) ** 4.87) if diam_linha_int_mm > 0 else 0
        hf_lateral_corrigida = hf_linha_continua * fator_f
        
        pressao_inicio_lateral = pressao_nominal_emissor + (0.75 * hf_lateral_corrigida) + (0.5 * z_lateral) + altura_riser
        derating_termico = 1.0 - ((temp_ambiente - 25.0) * 0.02) if temp_ambiente > 25.0 else 1.0
        derating_termico = max(0.4, derating_termico)
        pn_real_lateral = pn_classe_tubo_lateral * derating_termico
        seguro_pressao_lateral = pressao_inicio_lateral <= pn_real_lateral
        
        vazao_flushing_m3h = 0.5 * area_secao_linha * 3600.0
        limite_var_pressao = pressao_nominal_emissor * 0.20
        atende_iso9261 = hf_lateral_corrigida <= limite_var_pressao
        
        # Verificação de flushing vs vazão normal (Motor de Cálculo)
        flushing_excede = vazao_flushing_m3h > vazao_linha_lateral_m3h
        status_flushing = "Exige mais água para limpar do que opera" if flushing_excede else "OK"
        status_pressao_lateral = "SEGURO" if seguro_pressao_lateral else "RISCO ALTO: Hin > PN Real"
        
        hl1, hl2, hl3, hl4 = st.columns(4)
        hl1.metric("Perda de Carga Corrigida", f"{hf_lateral_corrigida:.2f} mca")
        hl2.metric("Pressão Entrada da Linha", f"{pressao_inicio_lateral:.2f} mca")
        hl3.metric("Classe Real (Amolecimento térmico)", f"{pn_real_lateral:.1f} mca")
        hl4.metric("Autolimpeza (Flushing)", f"{vazao_flushing_m3h:.3f} m³/h")
        
        hl5, hl6, hl7, hl8 = st.columns(4)
        hl5.metric("Nº de Emissores na Linha (N)", f"{num_emissores_por_linha}")
        hl6.metric("Vazão Total da Linha (Q)", f"{vazao_linha_lateral_m3h:.3f} m³/h")
        hl7.metric("Velocidade na Linha (V)", f"{vel_linha_lateral:.2f} m/s")
        hl8.metric("Derating Térmico", f"{derating_termico:.2f}", f"{temp_ambiente:.0f} °C")
        
        if seguro_pressao_lateral:
            st.success(f"Status de Pressão: {status_pressao_lateral}. A pressão de entrada ({pressao_inicio_lateral:.1f} mca) está abaixo da classe real do tubo ({pn_real_lateral:.1f} mca).")
        else:
            st.error(f"Status de Pressão: {status_pressao_lateral}. Pressão de entrada ({pressao_inicio_lateral:.1f} mca) acima da classe real do tubo ({pn_real_lateral:.1f} mca).")
        if flushing_excede:
            st.warning(f"Verificação de Flushing: {status_flushing}. A vazão para velocidade de 0,5 m/s ({vazao_flushing_m3h:.3f} m³/h) é maior que a vazão normal da linha ({vazao_linha_lateral_m3h:.3f} m³/h); prever válvula de descarga/final de linha.")
        else:
            st.success("Verificação de Flushing: OK. A vazão normal da linha já garante a velocidade mínima de autolimpeza.")

    with st.expander("2. Tubulação Principal Adutora e Aríete", expanded=True):
        st.markdown("**Por que este cálculo existe?** Para garantir que a tubulação principal não rache ou exploda se uma válvula fechar de repente (Golpe de Aríete), interrompendo o movimento da água subitamente.")
        
        _q_ad_sug = vazao_setor_operando_m3h / 3600.0
        _d_calc_mm = math.sqrt(4.0 * _q_ad_sug / (math.pi * 1.5)) * 1000.0 if _q_ad_sug > 0 else 25.0
        _d_comerciais = [25.0, 32.0, 40.0, 50.0, 60.0, 75.0, 85.0, 100.0, 125.0, 150.0, 200.0]
        _d_sugerido = next((d_ for d_ in _d_comerciais if d_ >= _d_calc_mm), _d_comerciais[-1])
        st.caption(f"Diâmetro sugerido: para a vazão de {_n(vazao_setor_operando_m3h, 2)} m³/h e velocidade-alvo de 1,5 m/s, o diâmetro interno calculado é {_n(_d_calc_mm, 1)} mm, "
                   f"o que leva ao diâmetro comercial aproximado de {_n(_d_sugerido, 0)} mm (faixa aceitável de velocidade: 0,6 a 2,2 m/s).")
        st.button(f"Usar o diâmetro sugerido ({_n(_d_sugerido, 0)} mm)", key="btn_d_sug", on_click=lambda v_=float(_d_sugerido): st.session_state.update({"diam_adutora": v_}))
        col_ad1, col_ad2, col_ad3 = st.columns(3)
        diam_adutora_int_mm = col_ad1.number_input("Diâmetro Interno da Adutora (mm):", value=float(_d_sugerido), step=5.0, key="diam_adutora", help="Medida interna do tubo escolhido. O valor digitado é mantido quando outros dados mudam; use o botão acima para voltar ao sugerido.")
        comp_equiv_adutora = col_ad2.number_input("Perdas em Conexões / Curvas (m):", value=10.0, step=2.0, help="Cada curva ou joelho adiciona uma resistência equivalente a metros a mais de tubo.", key="conex")
        pn_adutora = col_ad3.number_input("Classe de Pressão da Adutora - PN (mca):", value=40.0, step=5.0, help="Dado de catálogo do fabricante do tubo principal.")
        
        # Parâmetros do Motor de Cálculo (aba "2. Rede Hidráulica"), agora editáveis (valores padrão inalterados)
        col_ad4, col_ad5 = st.columns(2)
        margem_envelhecimento = col_ad4.number_input("Margem de Segurança p/ Envelhecimento (%):", value=15.0, min_value=0.0, max_value=50.0, step=5.0, help="Acréscimo na perda de carga para considerar o envelhecimento (incrustação) da tubulação.", key="margem")
        celeridade_onda_pvc = col_ad5.number_input("Celeridade da Onda - a (m/s):", value=350.0, min_value=200.0, max_value=500.0, step=25.0, help="Velocidade da onda de pressão. Aproximadamente 300 a 400 m/s para PVC/PE.")
        
        comp_principal_area = st.number_input("Linha principal dentro da área irrigada (m), no mesmo diâmetro da adutora:", min_value=0.0, value=0.0, step=5.0, key="l_princ",
                                              help="Trecho que percorre a roça levando água até as laterais. É somado ao comprimento da adutora (a vazão é considerada integral, o que é conservador).")
        comp_virtual_adutora = distancia_adutora + comp_equiv_adutora + comp_principal_area
        q_adutora_m3s = vazao_setor_operando_m3h / 3600.0
        area_secao_adutora = math.pi * ((diam_adutora_int_mm / 1000.0) ** 2) / 4.0
        vel_adutora = (q_adutora_m3s / area_secao_adutora) if area_secao_adutora > 0 else 0
        
        hf_adutora_teorica = 10.67 * comp_virtual_adutora * ((q_adutora_m3s / c_rugosidade_pvc) ** 1.852) / ((diam_adutora_int_mm / 1000.0) ** 4.87) if diam_adutora_int_mm > 0 else 0
        hf_adutora_ajustada = hf_adutora_teorica * (1.0 + (margem_envelhecimento / 100.0))
        
        sobrepressao_ariete = (celeridade_onda_pvc * vel_adutora) / 9.81
        ph_ariete = st.container()  # preenchido após o cálculo da AMT (precisa da sucção e do filtro)

    with st.expander("3. Dimensionamento da Bomba e Cavitação", expanded=True):
        st.markdown("**Por que este cálculo existe?** Define o tamanho do motor que você precisa comprar (CV) e calcula se a bomba não sofrerá Cavitação (fervura fria da água que destrói o rotor metálico).")
        
        col_b1, col_b2, col_b3 = st.columns(3)
        hf_filtro_areia = col_b1.number_input("Perda no Cabeçal de Filtragem (mca):", value=8.0, step=1.0, help="Estimativa da força perdida quando a água passa por um filtro de areia cheio de sujeira.", key="hf_fil")
        rendimento_bomba = col_b2.slider("Rendimento Hidráulico da Bomba (%):", 40, 85, 70, step=5, help="Eficiência nominal da bomba (retirada do catálogo do fabricante).", key="rend_b")
        fator_folga_motor = col_b3.slider("Folga de Potência do Motor (%):", 5, 25, 15, step=5, help="Margem de segurança padrão para não forçar o motor no limite.")
        
        col_b4, col_b5, col_b6 = st.columns(3)
        npsh_requerido_catalogo = col_b4.number_input("NPSH Requerido da Bomba (mca):", value=3.5, step=0.5, help="Dado Crítico do Fabricante! Mostra a pressão mínima na sucção para não destruir a bomba.")
        diam_succao_mm = col_b5.number_input("Diâmetro Tubo de Sucção (mm):", value=max(50.0, diam_adutora_int_mm * 1.25), step=5.0, help="Diâmetro do tubo que mergulha no poço/caixa. Jamais pode ser mais fino que a adutora.")
        hf_perda_succao = col_b6.number_input("Perda no Tubo de Sucção (mca):", value=1.2, step=0.1, help="Perda de energia específica só no pequeno trecho de sucção.", key="hf_suc")
        
        col_b7, col_b8, _col_b9 = st.columns(3)
        perda_loc_pct = col_b7.number_input("Perdas localizadas (% do subtotal) — opcional:", min_value=0.0, max_value=30.0, value=0.0, step=0.5, key="ploc",
                                            help="Acréscimo percentual sobre o subtotal das perdas (prática de 3 a 5 %). 0 = não aplicar (as conexões da adutora já entram pelo comprimento equivalente).")
        eta_motor_pct = col_b8.number_input("Rendimento do motor (%) — opcional:", min_value=50.0, max_value=100.0, value=100.0, step=1.0, key="etam3",
                                            help="Usado só na 'potência instalada' (potência da bomba ÷ rendimento do motor). 100 = não considerar.")
        subtotal_amt = desnivel_recalque + altura_succao + hf_perda_succao + hf_adutora_ajustada + hf_filtro_areia + pressao_inicio_lateral
        hf_localizada_pct = subtotal_amt * perda_loc_pct / 100.0
        amt_total = subtotal_amt + hf_localizada_pct  # inclui altura e perda na sucção
        potencia_cv_teorica = (vazao_setor_operando_m3h * amt_total) / (270.0 * (rendimento_bomba / 100.0))
        potencia_cv_comercial = potencia_cv_teorica * (1.0 + (fator_folga_motor / 100.0))
        # Motor comercial padronizado (próximo valor de catálogo igual ou acima do calculado)
        _MOTORES_CV = [0.25, 0.33, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 6.0, 7.5, 10.0, 12.5, 15.0, 20.0, 25.0, 30.0, 40.0, 50.0, 60.0, 75.0, 100.0]
        motor_comercial_cv = next((m for m in _MOTORES_CV if m >= potencia_cv_comercial - 1e-9), math.ceil(potencia_cv_comercial))
        
        # Potência hidráulica pura (sem rendimento), conforme Motor de Cálculo (aba "3. Bombeamento")
        potencia_hidraulica_pura_cv = (vazao_setor_operando_m3h * amt_total) / 270.0
        
        patm_local = 10.33 - (altitude_local * 0.001)
        pvapor = 0.32 
        npsh_disponivel = patm_local - altura_succao - hf_perda_succao - pvapor
        seguro_cavitacao = npsh_disponivel > npsh_requerido_catalogo
        regra_ouro_diametros = diam_succao_mm >= diam_adutora_int_mm
        
        hb1, hb2, hb3, hb4 = st.columns(4)
        hb1.metric("Altura Manométrica (AMT)", f"{amt_total:.2f} mca")
        hb2.metric("Potência Hidráulica", f"{potencia_cv_teorica:.2f} CV")
        hb3.metric("Bomba Comercial Sugerida", f"~ {potencia_cv_comercial:.2f} CV", help=f"Valor calculado com a folga do motor. Motor comercial de catálogo imediatamente acima: {motor_comercial_cv:g} CV (esse valor alimenta a aba 6).")
        st.caption(f"Motor comercial padronizado a adotar: {motor_comercial_cv:g} CV (potência calculada com folga = {potencia_cv_comercial:.2f} CV). A aba 6 usa esse valor.")
        hb4.metric("NPSH Disponível (NPSHd)", f"{npsh_disponivel:.2f} mca")
        
        potencia_cv_instalada = potencia_cv_teorica / (eta_motor_pct / 100.0)
        if perda_loc_pct > 0 or eta_motor_pct < 100:
            st.caption(f"Subtotal das perdas e desníveis = {_n(subtotal_amt, 2)} mca; perdas localizadas ({_n(perda_loc_pct, 1)} %) = {_n(hf_localizada_pct, 2)} mca; "
                       f"potência instalada (bomba ÷ motor) = {_n(potencia_cv_instalada, 2)} CV.")
        hb5, hb6, hb7, hb8 = st.columns(4)
        hb5.metric("Potência Hidráulica Teórica (Q·AMT/270)", f"{potencia_hidraulica_pura_cv:.2f} CV")
        hb6.metric("Potência Efetiva do Motor (sem fator)", f"{potencia_cv_teorica:.2f} CV")
        hb7.metric("Pressão Atmosférica Local", f"{patm_local:.2f} mca")
        hb8.metric("Vazão do Setor (Q)", f"{vazao_setor_operando_m3h:.2f} m³/h")
        
        if seguro_cavitacao:
            st.success(f"Status de Cavitação: SEGURO. NPSHd ({npsh_disponivel:.2f} mca) > NPSHr ({npsh_requerido_catalogo:.1f} mca).")
        else:
            st.error(f"Status de Cavitação: RISCO ALTO: Cavitação iminente. NPSHd ({npsh_disponivel:.2f} mca) ≤ NPSHr ({npsh_requerido_catalogo:.1f} mca).")
        if regra_ouro_diametros:
            st.success("Verificação de Diâmetros (Regra de Ouro): OK. Sucção maior ou igual ao recalque.")
        else:
            st.error("Verificação de Diâmetros (Regra de Ouro): ERRO TÉCNICO: Sucção menor que Recalque.")

        # ---------- Bombas comerciais do catálogo que atendem ao ponto de projeto
        bomba_cat_sel = None
        st.markdown("#### Bombas comerciais do catálogo que atendem ao projeto")
        _bombas_db, _cam_bombas = carregar_bombas()
        if not _bombas_db:
            st.info("Banco de bombas não encontrado. Coloque o arquivo **bombas_comerciais.xlsx** na mesma pasta do app (ou na pasta do projeto, Desktop ou Downloads) para listar as bombas comerciais. "
                    "Sem ele, a aba 5 usa uma curva estimada.")
        else:
            _qp, _hp = float(vazao_setor_operando_m3h), float(amt_total)
            st.caption(f"Ponto de projeto: {_n(_qp, 2)} m³/h a {_n(_hp, 1)} mca. Uma bomba atende quando a vazão está dentro da faixa tabelada do catálogo e a altura da curva nessa vazão "
                       f"é maior ou igual à altura exigida. Fonte: {os.path.basename(_cam_bombas)} ({len(_bombas_db)} bombas, vazão para sucção de 0 m.c.a.).")
            _cand, _prox = [], []
            for _b in _bombas_db:
                _q0, _q1 = float(_b["q"][0]), float(_b["q"][-1])
                _hb = float(np.interp(_qp, _b["q"], _b["h"]))
                if _q0 <= _qp <= _q1 and _hb >= _hp and _b["cv"] >= 0.95 * float(potencia_cv_teorica):
                    _rz = _qp / (0.5 * (_q0 + _q1))
                    _cand.append(((0 if 0.7 <= _rz <= 1.2 else 1, _hb - _hp), _b, _hb, _rz))
                else:
                    _qc = min(max(_qp, _q0), _q1); _hc = float(np.interp(_qc, _b["q"], _b["h"]))
                    _prox.append((max(0.0, _hp - _hc) / max(_hp, 1e-9) + abs(_qp - _qc) / max(_qp, 1e-9), _b, _hc))
            _cand.sort(key=lambda t: (t[0][0], t[0][1], t[1]["cv"]))
            _prox.sort(key=lambda t: t[0])
            def _linha_bomba(_b, _hb, _extra):
                return {"Modelo": _b["modelo"], "Potência (cv)": _b["cv"], "Conexões (sucção x recalque, pol)": f"{_b['suc']} x {_b['rec']}" if _b["suc"] else "—",
                        "Altura da bomba na vazão do projeto (mca)": round(_hb, 1), _extra[0]: _extra[1],
                        "Faixa de vazão tabelada (m³/h)": f"{_b['q'][0]:g} a {_b['q'][-1]:g}",
                        "Preço (R$)": (_n(_b["preco"], 2) if _b["preco"] else "sem preço")}
            if _cand:
                st.success(f"{len(_cand)} bomba(s) do catálogo atendem ao ponto de projeto (vazão na faixa tabelada, altura suficiente e motor maior ou igual à potência no eixo necessária). "
                           "Primeiro aparecem as que operam no meio da faixa tabelada (70 a 120%), da menor para a maior folga de altura.")
                _df_c = pd.DataFrame([dict(_linha_bomba(_b, _hb, ("Folga de altura (mca)", round(_m[1], 1))), **{"Posição na faixa (Q/Qmeio)": f"{_rz * 100:.0f}%"}) for _m, _b, _hb, _rz in _cand[:15]])
                st.dataframe(_df_c, hide_index=True, use_container_width=True)
                _opts = {f"{_b['modelo']} · {_b['cv']:g} cv · folga {_m[1]:.1f} mca": _b for _m, _b, _hb, _rz in _cand[:30]}
                _rot_sem = "Curva estimada (sem bomba do catálogo)"
                _lista_opts = [_rot_sem] + list(_opts.keys())
                if st.session_state.get("bomba_cat_lbl") not in _lista_opts:
                    st.session_state.pop("bomba_cat_lbl", None)
                _esc = st.selectbox("Usar a curva real de uma bomba do catálogo na aba 5 (ponto de operação):", _lista_opts, key="bomba_cat_lbl",
                                    help="A aba 5 passa a usar a curva tabelada da bomba escolhida e a aba 6 usa a potência do motor dela.")
                if _esc != _rot_sem:
                    _b = _opts[_esc]
                    _aj = ajustar_curva_catalogo(_b["q"], _b["h"], _b["pmax"] if (_b["pmax"] and _b["pmax"] >= float(np.max(_b["h"]))) else None)
                    if _aj is None:
                        st.warning("Não foi possível ajustar uma curva decrescente a esta bomba. Escolha outra ou use a curva estimada.")
                    else:
                        bomba_cat_sel = dict(_b, cf=_aj[0], rms=_aj[1], q_bep=0.5 * (float(_b["q"][0]) + float(_b["q"][-1])))
                        motor_comercial_cv = float(_b["cv"])
                        st.caption(f"Selecionada: {_b['modelo']} ({_b['cv']:g} cv). A aba 5 usará a curva dela e a aba 6 o motor de {_b['cv']:g} cv. "
                                   f"Preço: {'R$ ' + _n(_b['preco'], 2) if _b['preco'] else 'não informado (preencha a coluna Preço na planilha)'}.")
            else:
                st.warning("Nenhuma bomba do banco atende a este ponto de projeto. Pode faltar bomba maior ou multiestágio no banco. As mais próximas (não atendem) estão abaixo.")
                _df_p = pd.DataFrame([_linha_bomba(_b, _hc, ("Situação", "não atende")) for _s, _b, _hc in _prox[:8]])
                st.dataframe(_df_p, hide_index=True, use_container_width=True)
                st.session_state.pop("bomba_cat_lbl", None)


    with ph_ariete:
        pressao_saida_bomba_op = amt_total - altura_succao - hf_perda_succao  # carga na saída da bomba em operação (sem sucção)
        pressao_maxima_transiente = pressao_saida_bomba_op + sobrepressao_ariete
        seguro_ariete = pressao_maxima_transiente <= pn_adutora
        
        st.caption("Pico = pressão na saída da bomba em operação (desnível + perdas na adutora e no filtro + pressão da lateral) + sobrepressão de Joukowsky.")
        ha1, ha2, ha3, ha4 = st.columns(4)
        ha1.metric("Velocidade da Água", f"{vel_adutora:.2f} m/s")
        ha2.metric("Perda na Adutora", f"{hf_adutora_ajustada:.2f} mca")
        ha3.metric("Tranco de Aríete (ΔH)", f"+ {sobrepressao_ariete:.1f} mca")
        ha4.metric("Pressão Máxima de Pico", f"{pressao_maxima_transiente:.1f} mca")
        
        ha5, ha6, ha7 = st.columns(3)
        ha5.metric("Comprimento Total Virtual (Lt)", f"{comp_virtual_adutora:.1f} m")
        ha6.metric("Perda de Carga Teórica", f"{hf_adutora_teorica:.2f} mca")
        ha7.metric("Perda Ajustada (envelhecimento)", f"{hf_adutora_ajustada:.2f} mca", f"+{margem_envelhecimento:.0f}%")
        
        if seguro_ariete:
            st.success(f"Status de Risco de Rompimento (Aríete): SEGURO. Pico de {pressao_maxima_transiente:.1f} mca dentro da classe PN {pn_adutora:.0f} mca.")
        else:
            st.error(f"Status de Risco de Rompimento (Aríete): RISCO ALTO: Transiente > PN. Pico de {pressao_maxima_transiente:.1f} mca acima da classe PN {pn_adutora:.0f} mca.")

    with st.expander("4. Qualidade da Água e Risco de Entupimento dos Emissores", expanded=False):
        st.markdown("**Por que este cálculo existe?** Em irrigação localizada o entupimento é a principal causa de falha. A classificação abaixo segue Nakayama e Bucks (1991), "
                    "a partir da análise da água (laboratório). Os valores padrão representam água de boa qualidade; **substitua pelos da sua análise**.")
        cq1, cq2, cq3, cq4 = st.columns(4)
        q_ss = cq1.number_input("Sólidos suspensos (mg/L):", min_value=0.0, value=30.0, step=5.0, key="qa_ss")
        q_ph = cq2.number_input("pH:", min_value=3.0, max_value=11.0, value=6.8, step=0.1, key="qa_ph")
        q_fe = cq3.number_input("Ferro total (mg/L):", min_value=0.0, value=0.1, step=0.1, key="qa_fe")
        q_dz = cq4.number_input("Dureza total (mg/L CaCO₃):", min_value=0.0, value=100.0, step=10.0, key="qa_dz")
        cq5, cq6, cq7, cq8 = st.columns(4)
        q_mn = cq5.number_input("Manganês (mg/L):", min_value=0.0, value=0.05, step=0.05, key="qa_mn")
        q_h2s = cq6.number_input("Sulfeto de hidrogênio (mg/L):", min_value=0.0, value=0.1, step=0.1, key="qa_h2s")
        q_sd = cq7.number_input("Sólidos dissolvidos (mg/L):", min_value=0.0, value=300.0, step=50.0, key="qa_sd")
        q_orif = cq8.number_input("Diâmetro do orifício do emissor (mm):", min_value=0.2, value=1.0, step=0.1, key="qa_orif", help="Dado do catálogo do emissor. Define a abertura máxima do filtro.")
        _faixas = [("Sólidos suspensos (físico)", q_ss, 50, 100, "mg/L"), ("pH (químico)", q_ph, 7.0, 8.0, ""), ("Sólidos dissolvidos (químico)", q_sd, 500, 2000, "mg/L"),
                   ("Manganês (químico)", q_mn, 0.1, 1.5, "mg/L"), ("Ferro (químico)", q_fe, 0.2, 1.5, "mg/L"), ("Sulfeto de hidrogênio (químico)", q_h2s, 0.2, 2.0, "mg/L"),
                   ("Dureza (químico)", q_dz, 150, 300, "mg/L")]
        _niveis = ["Baixo", "Moderado", "Severo"]
        linhas_qa, nivel_max = [], 0
        for nome_, val_, lim1_, lim2_, un_ in _faixas:
            niv_ = 0 if val_ < lim1_ else (1 if val_ <= lim2_ else 2)
            nivel_max = max(nivel_max, niv_)
            linhas_qa.append({"Parâmetro": nome_, "Valor": f"{_n(val_, 2)} {un_}".strip(), "Baixo": f"< {_n(lim1_, 1)}", "Moderado": f"{_n(lim1_, 1)} a {_n(lim2_, 1)}", "Severo": f"> {_n(lim2_, 1)}", "Risco": _niveis[niv_]})
        risco_entupimento = _niveis[nivel_max]
        parametros_criticos_qa = [l_["Parâmetro"] for l_ in linhas_qa if l_["Risco"] == risco_entupimento] if nivel_max > 0 else []
        _micra = q_orif * 1000.0 / 10.0
        _malhas = [(80, 177), (100, 149), (120, 125), (150, 105), (200, 74), (250, 63)]
        malha_sugerida = next((m_ for m_, ab_ in _malhas if ab_ <= _micra), 250)
        st.dataframe(pd.DataFrame(linhas_qa), hide_index=True, use_container_width=True)
        qm1, qm2 = st.columns(2)
        qm1.metric("Risco de entupimento (pior parâmetro)", risco_entupimento)
        qm2.metric("Filtragem mínima sugerida", f"~{malha_sugerida} mesh", f"abertura ≤ {_n(_micra, 0)} µm (1/10 do orifício)", delta_color="off")
        if nivel_max == 0:
            st.success("Risco baixo: filtro de disco ou tela na malha sugerida e lavagem periódica das linhas atendem.")
        elif nivel_max == 1:
            st.warning("Risco moderado (" + "; ".join(parametros_criticos_qa) + "): use filtro de areia seguido de disco ou tela, lavagem frequente das laterais e acompanhe a análise da água a cada safra. "
                       "Para ferro e dureza, avalie injeção de ácido ou cloro.")
        else:
            st.error("Risco severo (" + "; ".join(parametros_criticos_qa) + "): faça pré-tratamento (decantação, aeração ou oxidação para ferro e manganês, cloração), filtragem em dois estágios e lavagem frequente. "
                     "Sem isso, a vida útil dos emissores cai muito.")
        st.caption("Regra prática: a abertura do filtro deve ser no máximo 1/10 do diâmetro do orifício do emissor; confirme com o fabricante. A conversão para mesh é aproximada."
                   + (" Em aspersão os bocais são maiores e o risco de entupimento é menor, mas a filtragem ainda protege bomba e aspersores." if metodo_selecionado == "Aspersão" else ""))

# ------------------------------------------------------------------------------
# 4. ABA: ESQUEMÁTICO E DIMENSIONAMENTO DA TUBULAÇÃO (estudo de tubulações)
# ------------------------------------------------------------------------------
with tab_tubulacao:
    st.subheader("Esquemático e Dimensionamento da Tubulação")
    dados_tubo = carregar_dados_tubulacao()
    par_t = dados_tubo["par"]
    st.caption(f"Base de dados da tubulação: {dados_tubo['origem']}. Perda contínua por Hazen-Williams e localizada pelo método dos comprimentos equivalentes (Apostila UFPel 2018-2, eq. 30 e Tab. 1D/1E).")

    # ---- Entradas: material e diâmetro ----
    st.markdown("#### 1. Material e Diâmetro")
    ct1, ct2, ct3, ct4 = st.columns(4)
    material_tubo = ct1.selectbox("Material da tubulação:", list(dados_tubo["materiais"].keys()), help="O coeficiente C vem da planilha de tubulações (Tab. 1D da apostila). PEAD e alumínio têm C adotado fora das apostilas: validar com o fabricante.")
    c_hw_tubo = ct2.number_input("Coeficiente C (Hazen-Williams):", value=float(dados_tubo["materiais"][material_tubo]), min_value=60.0, max_value=160.0, step=5.0, key=f"c_hw_{material_tubo}", help="Pode ser alterado (tubo usado ou incrustado tem C menor).")
    lista_dn = [d[0] for d in dados_tubo["dns"]]
    dn_tubo = ct3.selectbox("Diâmetro nominal (mm):", lista_dn, index=lista_dn.index(50) if 50 in lista_dn else 0)
    di_padrao = float(dict(dados_tubo["dns"])[dn_tubo])
    di_tubo_mm = ct4.number_input("Diâmetro interno real (mm):", value=di_padrao, min_value=5.0, step=1.0, key=f"di_tubo_{dn_tubo}", help="Por padrão igual ao nominal. Ajuste para o diâmetro interno do catálogo do tubo escolhido.")

    # ---- Entradas: geometria ----
    st.markdown("#### 2. Geometria, Orientação e Topografia")
    n_trechos = int(st.number_input("Número de trechos retos (cada mudança de direção = 1 curva de 90°):", value=2, min_value=2, max_value=8, step=1, help="2 trechos = 1 curva (L1 e L2). Cada trecho adicional acrescenta uma curva e um comprimento Ln."))
    cg1, cg2, cg3 = st.columns(3)
    l1_tubo = cg1.number_input("Trecho principal L1 (m) - avanço frontal:", value=max(1.0, float(distancia_adutora)), min_value=1.0, step=5.0, help="Comprimento do trecho que sai da bomba em frente. Padrão: distância da barra lateral.")
    lado_tubo = cg2.radio("Desvio do trecho lateral:", ["Direita", "Esquerda"], horizontal=True, help="Lado para o qual a 1ª curva vira (início de L2), olhando da bomba para a frente (sentido de L1).")
    l2_tubo = cg3.number_input("Trecho lateral L2 (m):", value=max(1.0, float(comp_linha_digitado)), min_value=1.0, step=1.0, help="Comprimento do trecho após a curva. Padrão: comprimento da linha lateral da aba 3.")

    comprimentos_trechos = [l1_tubo, l2_tubo]
    direcoes_curvas = [lado_tubo]  # direção da curva que antecede cada trecho a partir do 2º
    if n_trechos > 2:
        st.caption("Trechos adicionais: cada curva vira à Direita ou à Esquerda do sentido em que a tubulação vinha avançando.")
    for i_tr in range(3, n_trechos + 1):
        ce1, ce2 = st.columns(2)
        comprimentos_trechos.append(ce1.number_input(f"Trecho L{i_tr} (m):", value=max(1.0, float(l2_tubo)), min_value=1.0, step=1.0, key=f"l_trecho_{i_tr}"))
        direcoes_curvas.append(ce2.radio(f"Curva antes do trecho L{i_tr}:", ["Direita", "Esquerda"], horizontal=True, key=f"lado_trecho_{i_tr}"))
    n_curvas = n_trechos - 1

    cz1, cz2, cz3 = st.columns(3)
    dz_tubo = cz1.number_input("Desnível ΔZ (m):", value=float(desnivel_recalque), min_value=0.0, step=0.5, help="Diferença de cota entre a origem (bomba/cisterna) e o ponto final.")
    sentido_dz = cz2.radio("Sentido do desnível:", ["Aclive (ponto final mais alto)", "Declive (ponto final mais baixo)"], horizontal=True)
    aclive_tubo = sentido_dz.startswith("Aclive")
    q_tubo_m3h = cz3.number_input("Vazão de projeto Q (m³/h):", value=max(0.01, round(float(vazao_setor_operando_m3h), 3)), min_value=0.01, step=0.1, format="%.3f", help="Padrão: vazão do setor operando calculada na aba 3.")

    # ---- Entradas: conexões ----
    st.markdown("#### 3. Conexões e Acessórios")
    cx1, cx2, cx3, cx4 = st.columns(4)
    tipo_cotovelo = cx1.selectbox("Tipo de cotovelo 90°:", ["Cotovelo 90° (raio médio)", "Cotovelo 90° (raio longo)"])
    n_cot_extra = int(cx1.number_input("Cotovelos 90° adicionais (un):", value=0, min_value=0, step=1, help="As curvas do trajeto já contam como cotovelos 90° automaticamente. Use este campo só para cotovelos extras (desvios de obstáculos, saídas etc.)."))
    n_cotovelos = n_curvas + n_cot_extra
    n_tes_direta = int(cx2.number_input("Tês - passagem direta (un):", value=0, min_value=0, step=1))
    n_tes_lado = int(cx3.number_input("Tês - saída de lado (un):", value=0, min_value=0, step=1))
    n_luvas = int(cx4.number_input("Luvas (un):", value=0, min_value=0, step=1, help="Luva não consta na Tab. 1E; perda considerada desprezível (parâmetro LEQ_LUVA_M da planilha)."))
    with st.expander("Acessórios adicionais (opcional)"):
        ca1, ca2 = st.columns(2)
        n_registros = int(ca1.number_input("Registros de gaveta abertos (un):", value=0, min_value=0, step=1))
        n_retencao = int(ca2.number_input("Válvulas de retenção leve (un):", value=0, min_value=0, step=1))

    cp1, cp2 = st.columns(2)
    p_final_mca = cp1.number_input("Pressão de serviço exigida no ponto final (mca):", value=max(0.0, round(float(pressao_inicio_lateral), 1)), min_value=0.0, step=1.0, help="Pressão que precisa chegar ao fim do trecho. Padrão: pressão de entrada da linha lateral (aba 3).")
    pn_tubo_mca = cp2.number_input("Classe de pressão do tubo - PN (mca):", value=max(5.0, float(pn_adutora)), min_value=5.0, step=5.0, help="Dado de catálogo. Padrão: PN da adutora (aba 3).")

    # ---- Cálculos ----
    l_real = sum(comprimentos_trechos)
    leq_cot = n_cotovelos * leq_peca(dados_tubo, tipo_cotovelo, di_tubo_mm)
    leq_te_dir = n_tes_direta * leq_peca(dados_tubo, "Tê passagem direta", di_tubo_mm)
    leq_te_lado = n_tes_lado * leq_peca(dados_tubo, "Tê saída de lado", di_tubo_mm)
    leq_luva = n_luvas * par_t["LEQ_LUVA_M"]
    leq_reg = n_registros * leq_peca(dados_tubo, "Registro de gaveta aberto", di_tubo_mm)
    leq_ret = n_retencao * leq_peca(dados_tubo, "Válvula de retenção leve", di_tubo_mm)
    leq_total = leq_cot + leq_te_dir + leq_te_lado + leq_luva + leq_reg + leq_ret
    l_virtual = l_real + leq_total

    k_hw = par_t["K_HW"]
    hf_continua_t = hazen_williams_hf(q_tubo_m3h, di_tubo_mm, c_hw_tubo, l_real, k_hw)
    hf_localizada_t = hazen_williams_hf(q_tubo_m3h, di_tubo_mm, c_hw_tubo, leq_total, k_hw)
    hf_total_t = hf_continua_t + hf_localizada_t

    area_tubo_m2 = math.pi * ((di_tubo_mm / 1000.0) ** 2) / 4.0
    vel_tubo = (q_tubo_m3h / 3600.0) / area_tubo_m2 if area_tubo_m2 > 0 else 0.0

    # Carga geométrica: aclive soma à exigência da bomba (+ΔZ a vencer); declive abate (−ΔZ a favor)
    dz_assinado = dz_tubo if aclive_tubo else -dz_tubo
    h_bomba_exigida = p_final_mca + hf_total_t + dz_assinado
    h_bomba_exigida_ref = max(h_bomba_exigida, 0.0)

    usa_derating = material_tubo in ("PVC", "PEAD / PEBD")
    pn_real_tubo = pn_tubo_mca * (derating_termico if usa_derating else 1.0)

    # ---- Esquemático (geometria única compartilhada com a planta baixa e o 3D) ----
    n_linhas_campo = num_total_linhas_campo
    layout_campo = calcular_layout_campo(comprimentos_trechos, direcoes_curvas, dz_tubo, aclive_tubo, comp_linha_digitado,
                                         espacamento_linhas, espacamento_emissores, n_linhas_campo, num_setores,
                                         area_plantio, area_telhado, vol_cisterna_existente)
    n_setores_desenho = len(layout_campo['setores'])
    setor_ativo_efetivo = setor_ativo_vis if setor_ativo_vis <= n_setores_desenho else 0
    cfg_visual = dict(
        metodo=metodo_selecionado, cultura=cultura, tipo_solo=tipo_solo, material=material_tubo, dn_mm=float(dn_tubo),
        setor_ativo=setor_ativo_efetivo, vol_cisterna=vol_cisterna_existente,
        nivel_cisterna=min(1.0, vol_liquido_aproveitavel / vol_cisterna_existente) if vol_cisterna_existente > 0 else 0.5,
        nivel_backup=nivel_boia_backup, pot_cv=potencia_cv_comercial, area_plantio=area_plantio, area_telhado=area_telhado,
        arborea=_cultura_arborea(cultura), barra_m=par_t["COMP_BARRA_M"], n_setores=n_setores_desenho, runoff=c_runoff,
        diam_calha=diam_calha, diam_condutor=diam_condutor, vol_ff=vol_first_flush, vol_liq=vol_liquido_aproveitavel,
        vazao_rede=vazao_rede_concessionaria, precip=precipitacao_referencia, vol_bruto=vol_bruto_captado, autonomia=autonomia_dias,
    )
    st.markdown("#### 4. Esquemático Visual")
    st.caption("Mostra a casa, a cisterna, a bomba, o trajeto em PVC com as curvas, as válvulas e os setores. Altere o setor em operação na barra lateral (5. Visualização).")
    st.markdown(gerar_svg_planta_realista(layout_campo, cfg_visual, detalhe=False, uid="pt4"), unsafe_allow_html=True)
    st.markdown(gerar_svg_perfil_realista(layout_campo, cfg_visual, uid="pf4"), unsafe_allow_html=True)
    with st.expander("Esquemático simplificado (versão anterior)"):
        esq1, esq2 = st.columns(2)
        esq1.markdown(gerar_svg_planta_multi(comprimentos_trechos, direcoes_curvas, n_cotovelos, n_tes_direta + n_tes_lado), unsafe_allow_html=True)
        esq2.markdown(gerar_svg_perfil(l1_tubo, l2_tubo, dz_tubo, aclive_tubo, comprimentos_trechos), unsafe_allow_html=True)

    # ---- Resultados ----
    st.markdown("#### 5. Perdas de Carga e Pressão Exigida na Bomba")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric(f"Comprimento Real ({n_trechos} trechos)", f"{l_real:.1f} m", f"{n_curvas} curva(s)")
    r2.metric("Comprimento Equivalente (conexões)", f"{leq_total:.2f} m")
    r3.metric("Comprimento Virtual Total", f"{l_virtual:.2f} m")
    r4.metric("Velocidade (V)", f"{vel_tubo:.2f} m/s")

    r5, r6, r7, r8 = st.columns(4)
    r5.metric("Perda Contínua (hf)", f"{hf_continua_t:.2f} mca")
    r6.metric("Perda Localizada", f"{hf_localizada_t:.2f} mca")
    r7.metric("Perda Total", f"{hf_total_t:.2f} mca")
    r8.metric("Carga Geométrica", f"{dz_assinado:+.1f} m", "Aclive: bomba vence a altura" if aclive_tubo else "Declive: gravidade ajuda")

    r9, r10, r11 = st.columns(3)
    r9.metric("Pressão Exigida na Bomba", f"{h_bomba_exigida_ref:.2f} mca", f"{h_bomba_exigida_ref * 0.0980665:.2f} bar")
    r10.metric("Classe Real do Tubo", f"{pn_real_tubo:.1f} mca", f"derating {derating_termico:.2f}" if usa_derating else "metal: sem derating")
    r11.metric("Coeficiente C Adotado", f"{c_hw_tubo:.0f}", material_tubo)

    # ---- Alertas ----
    if di_tubo_mm < par_t["D_MIN_HW_MM"]:
        st.warning(f"Hazen-Williams é válida para D ≥ {par_t['D_MIN_HW_MM']:.0f} mm (Apostila UFPel, p. 21). Com D = {di_tubo_mm:.0f} mm o resultado é apenas uma estimativa; para laterais finas confira com a fórmula Universal ou o catálogo do fabricante.")
    if vel_tubo < par_t["V_MIN"]:
        st.warning(f"Velocidade baixa: {vel_tubo:.2f} m/s está abaixo de {par_t['V_MIN']:.1f} m/s. Risco de deposição de sedimentos e acúmulo de ar. Considere um diâmetro menor.")
    elif vel_tubo <= par_t["V_MAX"]:
        st.success(f"Velocidade OK: {vel_tubo:.2f} m/s dentro da faixa {par_t['V_MIN']:.1f} a {par_t['V_MAX']:.1f} m/s.")
    elif vel_tubo <= par_t["V_RECALQUE_ABS"]:
        st.warning(f"Velocidade alta: {vel_tubo:.2f} m/s acima de {par_t['V_MAX']:.1f} m/s. Perdas e golpe de aríete aumentam; considere um diâmetro maior.")
    else:
        st.error(f"Velocidade excessiva: {vel_tubo:.2f} m/s acima do máximo absoluto de {par_t['V_RECALQUE_ABS']:.1f} m/s para recalque (Apostila UFPel, p. 65). Aumente o diâmetro.")

    if h_bomba_exigida > pn_real_tubo:
        st.error(f"Pressão excessiva: a bomba precisa entregar {h_bomba_exigida_ref:.1f} mca, acima da classe real do tubo ({pn_real_tubo:.1f} mca). Escolha um tubo de PN maior ou reduza as perdas.")
    elif h_bomba_exigida > 0.8 * pn_real_tubo:
        st.warning(f"Pressão elevada: {h_bomba_exigida_ref:.1f} mca já passa de 80% da classe real do tubo ({pn_real_tubo:.1f} mca). Sem folga para o golpe de aríete (veja a aba 3).")
    else:
        st.success(f"Pressão compatível: {h_bomba_exigida_ref:.1f} mca exigidos na bomba, dentro da classe real do tubo ({pn_real_tubo:.1f} mca).")

    if h_bomba_exigida < 0:
        st.info("O declive é tão favorável que a gravidade sozinha supera as perdas e a pressão de serviço. Verifique se é necessário válvula redutora de pressão ou quebra-pressão.")

    if npsh_disponivel > npsh_requerido_catalogo:
        st.success(f"Cavitação: SEGURO. NPSHd ({npsh_disponivel:.2f} mca) > NPSHr ({npsh_requerido_catalogo:.1f} mca) (dados da aba 3).")
    else:
        st.error(f"Risco de cavitação: NPSHd ({npsh_disponivel:.2f} mca) ≤ NPSHr ({npsh_requerido_catalogo:.1f} mca) (dados da aba 3). Reduza a altura de sucção ou a perda na sucção.")

    # ---- Lista de materiais ----
    st.markdown("#### 6. Lista de Materiais (BOM)")
    barra_m = par_t["COMP_BARRA_M"]
    sobra_pct = par_t["SOBRA_CORTE_PCT"]
    barras_trechos = [math.ceil(c * (1 + sobra_pct / 100.0) / barra_m) for c in comprimentos_trechos]
    luvas_emenda = sum(max(0, b - 1) for b in barras_trechos)
    espec = f"{material_tubo}, DN {dn_tubo} mm"
    linhas_bom = [
        *[{"Item": f"Barras de tubo - trecho L{i + 1}", "Especificação": f"{espec}, barra de {barra_m:.0f} m", "Qtd": barras_trechos[i], "Unid": "barra", "Observação": f"{comprimentos_trechos[i]:.1f} m + {sobra_pct:.0f}% de sobra"} for i in range(len(comprimentos_trechos))],
        {"Item": "Cotovelos 90°", "Especificação": f"{espec} ({tipo_cotovelo})", "Qtd": n_cotovelos, "Unid": "un", "Observação": f"{n_curvas} curva(s) do trajeto + {n_cot_extra} adicional(is)"},
        {"Item": "Tês - passagem direta", "Especificação": espec, "Qtd": n_tes_direta, "Unid": "un", "Observação": ""},
        {"Item": "Tês - saída de lado", "Especificação": espec, "Qtd": n_tes_lado, "Unid": "un", "Observação": ""},
        {"Item": "Luvas (informadas)", "Especificação": espec, "Qtd": n_luvas, "Unid": "un", "Observação": "Perda localizada desprezada"},
        {"Item": "Luvas de emenda entre barras (sugestão)", "Especificação": espec, "Qtd": luvas_emenda, "Unid": "un", "Observação": "1 por junta entre barras; confirme o tipo de união do material"},
        {"Item": "Registros de gaveta", "Especificação": espec, "Qtd": n_registros, "Unid": "un", "Observação": ""},
        {"Item": "Válvulas de retenção leve", "Especificação": espec, "Qtd": n_retencao, "Unid": "un", "Observação": ""},
    ]
    df_bom = pd.DataFrame([l for l in linhas_bom if l["Qtd"] > 0])
    st.dataframe(df_bom, hide_index=True, use_container_width=True)
    st.caption("A lista cobre tubos e conexões do trajeto (todos os trechos). Adesivo, vedação, suportes, válvulas de descarga e demais acessórios não estão incluídos.")
    st.download_button("Baixar lista de materiais (CSV)", data=df_bom.to_csv(index=False).encode("utf-8-sig"),
                       file_name="Lista_de_Materiais_Tubulacao.csv", mime="text/csv", use_container_width=True)

    with st.expander("Memória de cálculo desta aba"):
        st.markdown(
            f"- Hazen-Williams: hf = {k_hw} · L · (Q/C)^1,852 / D^4,87, com Q em m³/s e D em m (Apostila UFPel 2018-2, eq. 30).\n"
            f"- Comprimento virtual: Lv = ΣLn ({' + '.join(f'{c:.1f}' for c in comprimentos_trechos)}) + ΣLeq = {l_real:.1f} + {leq_total:.2f} = {l_virtual:.2f} m (Tab. 1E interpolada em D = {di_tubo_mm:.0f} mm).\n"
            f"- Pressão exigida na bomba = pressão de serviço + perda total + carga geométrica = {p_final_mca:.1f} + {hf_total_t:.2f} + ({dz_assinado:+.1f}) = {h_bomba_exigida:.2f} mca.\n"
            f"- Convenção: aclive soma ΔZ à exigência da bomba; declive subtrai (a gravidade ajuda). Declividade considerada uniforme ao longo de todos os trechos.\n"
            f"- Esta aba é independente da AMT da aba 3 (que usa a adutora, o filtro e a lateral). Use-a para estudar o trajeto isoladamente."
        )

# ------------------------------------------------------------------------------
# 5. ABA: INSTALAÇÃO ELÉTRICA
# ------------------------------------------------------------------------------
with tab_eletrica:
    st.subheader("Dimensionamento Elétrico do Alimentador da Bomba")
    
    col_e1, col_e2, col_e3 = st.columns(3)
    pot_motor_adotada = col_e1.number_input("Potência Nominal do Motor (CV):", value=float(motor_comercial_cv), min_value=0.25, step=0.25, help="Padrão: motor comercial imediatamente acima da potência calculada na aba 3. Pode ser alterado para o motor realmente adquirido.")
    tensao_circuito = col_e2.selectbox("Tensão da Rede Elétrica:", ["Monofásica 220V", "Trifásica 220V", "Trifásica 380V"], help="Padrão fornecido pela concessionária elétrica.")
    distancia_quadro_cabo = col_e3.number_input("Distância do Quadro Elétrico até a Bomba (m):", value=60.0, step=5.0, help="Medida física para calcular perda de tensão (resistência do cabo de cobre).")
    
    col_e4, _col_e5, _col_e6 = st.columns(3)
    queda_max_perm = col_e4.number_input("Queda de Tensão Máxima Permitida (%):", value=4.0, min_value=1.0, max_value=8.0, step=0.5, help="Geralmente 4% a 5% (NBR 5410).")
    
    pot_watts = pot_motor_adotada * 736.0
    cos_phi = 0.85
    rend_motor = 0.80
    rho_cobre = 0.0172
    
    if "Monofásica 220V" in tensao_circuito:
        v_rede = 220.0
        corrente_nominal_in = pot_watts / (v_rede * cos_phi * rend_motor)
        fator_circ = 2.0
    elif "Trifásica 220V" in tensao_circuito:
        v_rede = 220.0
        corrente_nominal_in = pot_watts / (math.sqrt(3) * v_rede * cos_phi * rend_motor)
        fator_circ = math.sqrt(3)
    else:
        v_rede = 380.0
        corrente_nominal_in = pot_watts / (math.sqrt(3) * v_rede * cos_phi * rend_motor)
        fator_circ = math.sqrt(3)
        
    delta_v_volts_limite = v_rede * (queda_max_perm / 100.0)
    secao_teorica_mm2 = (fator_circ * rho_cobre * distancia_quadro_cabo * corrente_nominal_in) / delta_v_volts_limite
    
    bitolas_comerciais = [1.5, 2.5, 4.0, 6.0, 10.0, 16.0, 25.0, 35.0, 50.0, 70.0, 95.0, 120.0]
    bitola_escolhida = bitolas_comerciais[-1]
    for b in bitolas_comerciais:
        if b >= secao_teorica_mm2:
            bitola_escolhida = b
            break
            
    queda_efetiva_v = (fator_circ * rho_cobre * distancia_quadro_cabo * corrente_nominal_in) / bitola_escolhida
    queda_efetiva_perc = (queda_efetiva_v / v_rede) * 100.0
    
    st.markdown("#### Especificação do Cabo e Proteção")
    me1, me2, me3, me4 = st.columns(4)
    me1.metric("Corrente Elétrica (In)", f"{corrente_nominal_in:.1f} A")
    me2.metric("Seção Teórica Mínima", f"{secao_teorica_mm2:.2f} mm²")
    me3.metric("Cabo Comercial Recomendado", f"{bitola_escolhida:.1f} mm²")
    me4.metric("Queda de Tensão Efetiva", f"{queda_efetiva_perc:.2f} %")
    
    if secao_teorica_mm2 > bitolas_comerciais[-1]:
        st.error(f"A seção teórica ({secao_teorica_mm2:.1f} mm²) excede a maior bitola da tabela ({bitolas_comerciais[-1]:.0f} mm²). Reduza a distância, aumente a tensão ou use cabos em paralelo.")
    elif queda_efetiva_perc <= queda_max_perm:
        st.success(f"Queda de tensão de {queda_efetiva_perc:.2f}% dentro do limite de {queda_max_perm:.1f}% (NBR 5410). Cabo de cobre de {_n(bitola_escolhida, 1)} mm².")
    else:
        st.error(f"Queda de tensão de {queda_efetiva_perc:.2f}% acima do limite de {queda_max_perm:.1f}%.")

# ------------------------------------------------------------------------------
# 1 (continuação). ABA: ÁGUA DE CHUVA — esquema realista (precisa da bomba e do tubo, calculados nas abas 3 e 4)
# ------------------------------------------------------------------------------
with tab_captacao:
    st.markdown(gerar_svg_cisterna_realista(cfg_visual, uid="ci1"), unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 6. ABA: PLANTA BAIXA E SIMULAÇÃO 3D (realista e interligada)
# ------------------------------------------------------------------------------
with tab_planta_3d:
    st.subheader("Planta Baixa e Maquete Digital 3D")
    st.caption("Tudo usa a mesma geometria: trajeto e curvas (aba 4), setores, linhas e emissores (aba 3), casa e cisterna (aba 1), bomba (aba 3), solo e cultura (aba 2).")
    vv1, vv2, vv3, vv4 = st.columns(4)
    vv1.metric("Setores no desenho", f"{n_setores_desenho}", f"pedidos: {num_setores}" if n_setores_desenho != num_setores else None, delta_color="off")
    vv2.metric("Linhas por setor", f"{num_linhas_por_setor}")
    vv3.metric("Linhas laterais", f"{layout_campo['n_linhas']}")
    vv4.metric("Emissores no total", f"{layout_campo['n_linhas'] * layout_campo['n_em']:,}".replace(",", "."))
    if n_setores_desenho < num_setores:
        st.info(f"Você pediu {num_setores} setores, mas com {layout_campo['n_linhas']} linha(s) lateral(is) só é possível formar {n_setores_desenho} setor(es) com linhas (mesma regra da aba 3).")
    if setor_ativo_vis > n_setores_desenho:
        st.warning(f"O Setor {setor_ativo_vis} não existe neste projeto ({n_setores_desenho} setor(es)); mostrando todos em operação.")

    st.markdown("#### Planta baixa")
    st.markdown(gerar_svg_planta_realista(layout_campo, cfg_visual, detalhe=True, uid="pb6"), unsafe_allow_html=True)
    st.markdown("#### Detalhe da roça: setores, válvulas e linhas laterais")
    st.markdown(gerar_svg_planta_realista(layout_campo, cfg_visual, detalhe=True, uid="rz6", foco="roca"), unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Maquete 3D interativa")
    c3a, c3b, c3c = st.columns([3, 2, 1])
    camera_3d = c3a.radio("Câmera:", list(CAMERAS_3D.keys()), index=0, horizontal=True, help="Gire com o mouse, aproxime com a rolagem. As visões de foco já enquadram a roça ou a casa e a bomba.")
    exagero_tubo = c3b.slider("Espessura visual dos tubos:", 1.0, 4.0, 2.0, 0.5, help="Em escala real os tubos seriam finos demais para enxergar; este controle só afeta o desenho.")
    mostrar_arvores = c3c.checkbox("Árvores", value=True)
    fig_3d = gerar_cena_3d(layout_campo, dict(cfg_visual, camera=camera_3d, exagero=exagero_tubo, arvores=mostrar_arvores))
    st.plotly_chart(fig_3d, use_container_width=True, key="cena_3d_realista")
    st.caption("Espessura dos tubos exagerada para visualização; distâncias, desnível, setores, linhas e emissores seguem os dados do projeto. "
               "Setores verdes = válvula aberta (irrigando), vermelhos = fechada.")

    with st.expander("Esquemáticos simplificados (versão anterior)", expanded=False):
    
        def gerar_planta_baixa_svg(distancia, area, comp_linha, cultura, metodo):
            cor_fundo, cor_tubo_prin, cor_tubo_lat, cor_agua, cor_texto = "#FDFEFE", "#34495E", "#7F8C8D", "#3498DB", "#2C3E50"
            cor_filtro, cor_tanque = "#BDC3C7", "#95A5A6"
            linhas_html = ""
            y_pos = [240, 290, 340, 390, 440]
            for i, y in enumerate(y_pos):
                x_start = 300 + (i * 20)
                x_end = x_start + 450
                linhas_html += f'<line x1="{x_start}" y1="{y}" x2="{x_end}" y2="{y}" stroke="{cor_tubo_lat}" stroke-width="3" />'
                for xg in range(x_start + 30, x_end - 20, 50):
                    linhas_html += f'<circle cx="{xg}" cy="{y}" r="2.5" fill="{cor_agua}" />'
        
            svg = f"""
            <div style="display: flex; justify-content: center; padding: 20px; background-color: {cor_fundo}; border-radius: 10px; border: 1px solid #ddd; overflow-x: auto;">
                <svg width="100%" height="auto" viewBox="-150 -20 1200 750" style="max-width: 900px;" xmlns="http://www.w3.org/2000/svg">
                    <text x="450" y="30" font-family="Arial" font-size="18" fill="{cor_texto}" font-weight="bold" text-anchor="middle">Arranjo Isométrico do Projeto: {metodo}</text>
                    <g transform="rotate(-15, 450, 275) translate(80, 150)">
                        <rect x="50" y="60" width="30" height="70" fill="{cor_filtro}" rx="5" stroke="#7F8C8D" stroke-width="2"/>
                        <text x="65" y="50" font-family="Arial" font-size="11" fill="{cor_texto}" text-anchor="middle">Filtro Areia</text>
                        <line x1="-70" y1="120" x2="260" y2="120" stroke="{cor_tubo_prin}" stroke-width="6" />
                        <line x1="260" y1="120" x2="260" y2="440" stroke="{cor_tubo_prin}" stroke-width="6" />
                        <text x="250" y="280" font-family="Arial" font-size="12" fill="{cor_texto}" font-weight="bold" transform="rotate(-90, 250, 280)">Linha Principal</text>
                        <g>{linhas_html}</g>
                    </g>
                </svg>
            </div>
            """
            return svg.replace('\n', '')

        st.markdown(gerar_planta_baixa_svg(distancia_adutora, area_plantio, comp_linha_digitado, cultura, metodo_selecionado), unsafe_allow_html=True)
    
        st.markdown("---")
        def gerar_simulacao_3d(distancia, desnivel, area, comp_linha, solo, metodo):
            lado = math.sqrt(area)
            fig = go.Figure()
            fig.add_trace(go.Scatter3d(
                x=[0, distancia], y=[0, 0], z=[0, desnivel], mode='lines+markers',
                line=dict(color='#34495E', width=8), marker=dict(size=[12, 8], color=['#3498DB', '#E74C3C']),
                name='Adutora Principal'
            ))
        
            y_lines = np.linspace(-lado/2, lado/2, 5)
            for y in y_lines:
                fig.add_trace(go.Scatter3d(
                    x=[distancia, distancia + comp_linha], y=[y, y], z=[desnivel, desnivel],
                    mode='lines', line=dict(color='#27AE60', width=4), showlegend=False
                ))
            
            fig.update_layout(
                title=f"Maquete 3D Interativa",
                scene=dict(xaxis_title='Comprimento (m)', yaxis_title='Largura (m)', zaxis_title='Elevação (m)'),
                margin=dict(l=0, r=0, b=0, t=40)
            )
            return fig

        fig_3d = gerar_simulacao_3d(distancia_adutora, desnivel_recalque, area_plantio, comp_linha_digitado, tipo_solo, metodo_selecionado)
        st.plotly_chart(fig_3d, use_container_width=True)


# ------------------------------------------------------------------------------
# 5. ABA: PROCESSADOR HIDRÁULICO (bomba x sistema + solver de rede)
# ------------------------------------------------------------------------------
proc_resumo = {"ok": False}

with tab_processador:
    import hashlib
    from plotly.subplots import make_subplots

    st.subheader("Processador Hidráulico: Bomba x Sistema e Solver de Rede")
    _sis0 = dict(q_ref=float(vazao_setor_operando_m3h), L_rec=float(comp_virtual_adutora), D_rec=float(diam_adutora_int_mm), C_rec=float(c_rugosidade_pvc),
                 env=float(margem_envelhecimento) / 100.0, hf_suc=float(hf_perda_succao), hf_fil=float(hf_filtro_areia), hin=float(pressao_inicio_lateral),
                 x=0.5, h_est=float(desnivel_recalque + altura_succao))
    amt_req = proc_h_sistema(float(vazao_setor_operando_m3h), _sis0)  # carga exigida na vazão de projeto, já com sucção
    st.caption("Em vez de impor a vazão, o processador encontra onde a bomba e a tubulação realmente se equilibram (ponto de operação) e resolve a rede inteira "
               "(cisterna, sucção, bomba, adutora, filtro, setores) com o método de Newton (gradiente global), incluindo malhas e emissores que variam a vazão com a pressão.")

    # ---------------------------------------------------------------- 1. BOMBA
    with st.expander("1. Curva da Bomba e Parâmetros de Operação", expanded=True):
        st.markdown("Informe os pontos do **catálogo do fabricante**. Se deixar os valores padrão, o sistema usa uma curva típica estimada a partir do seu projeto "
                    "(altura nominal = carga exigida pelo sistema na vazão de projeto; altura a vazão zero = 1,3 x essa carga), que serve para estudo, mas **deve ser trocada pela curva real antes da compra**.")
        curva_manual = st.checkbox("Informar os pontos do catálogo do fabricante", value=False, key="pb_manual",
                                   help="Desmarcado: a bomba é estimada a partir do projeto e acompanha qualquer mudança nos dados. Marcado: você digita os valores do catálogo.")
        q_bep, h_nom_b = max(0.1, float(vazao_setor_operando_m3h)), max(1.0, float(amt_req))
        h0_b, eta_bep = round(1.3 * h_nom_b, 1), float(rendimento_bomba) / 100.0
        pontos_b = [(q_bep, h_nom_b)]
        if curva_manual and bomba_cat_sel is None:
            pb1, pb2, pb3, pb4 = st.columns(4)
            q_bep = pb1.number_input("Vazão nominal / melhor eficiência (m³/h):", min_value=0.1, value=round(q_bep, 2), step=0.5)
            h_nom_b = pb2.number_input("Altura no ponto nominal (mca):", min_value=1.0, value=round(h_nom_b, 1), step=1.0)
            h0_b = pb3.number_input("Altura com registro fechado (mca):", min_value=1.0, value=round(max(1.3 * h_nom_b, h_nom_b + 1.0), 1), step=1.0,
                                    help="Também chamada de shutoff. Está no catálogo; típico de 1,2 a 1,4 vezes a altura nominal.")
            eta_bep = pb4.number_input("Rendimento no ponto nominal (%):", min_value=20.0, max_value=90.0, value=float(rendimento_bomba), step=1.0) / 100.0
            pontos_b = [(q_bep, h_nom_b)]
            if st.checkbox("Informar um segundo ponto do catálogo (curva mais fiel)", value=False, key="pb_2p"):
                pq1, pq2 = st.columns(2)
                q2_b = pq1.number_input("2º ponto: vazão (m³/h):", min_value=0.1, value=round(1.3 * q_bep, 2), step=0.5)
                h2_b = pq2.number_input("2º ponto: altura (mca):", min_value=0.0, value=round(max(0.0, h_nom_b - (h0_b - h_nom_b) * 0.69), 1), step=1.0)
                pontos_b.append((q2_b, h2_b))
        elif bomba_cat_sel is None:
            st.caption(f"Curva estimada: ponto nominal {_n(q_bep, 2)} m³/h a {_n(h_nom_b, 1)} mca, {_n(h0_b, 1)} mca com registro fechado e rendimento de {eta_bep * 100:.0f}%.")
        if bomba_cat_sel is not None:
            q_bep = bomba_cat_sel["q_bep"]; h0_b = bomba_cat_sel["cf"][0]
            h_nom_b = proc_h_bomba(q_bep, bomba_cat_sel["cf"]); pontos_b = [(q_bep, h_nom_b)]
            st.success(f"Curva real do catálogo: {bomba_cat_sel['modelo']} ({bomba_cat_sel['cv']:g} cv), ajustada a {len(bomba_cat_sel['q'])} pontos tabelados "
                       f"(erro médio do ajuste {_n(bomba_cat_sel['rms'], 2)} mca). O ponto de melhor eficiência não consta nas tabelas de seleção e foi estimado no meio da faixa de vazão tabelada "
                       f"({_n(q_bep, 2)} m³/h); o rendimento vem do campo da aba 3. Para validar, use a curva completa do fabricante.")
        curva_real = curva_manual or (bomba_cat_sel is not None)
        pe1, pe2, pe3 = st.columns(3)
        eta_motor = pe1.number_input("Rendimento do motor elétrico (%):", min_value=50.0, max_value=98.0, value=85.0, step=1.0, key="pb_etam") / 100.0
        tarifa_kwh = pe2.number_input("Tarifa de energia (R$/kWh):", min_value=0.0, value=0.90, step=0.05, key="pb_tar", help="Valor ilustrativo. Use o da sua fatura de energia.")
        x_emissor = pe3.number_input("Expoente do emissor x (q = k·H^x):", min_value=0.05, max_value=1.0, value=0.50, step=0.05, key="pb_x",
                                     help="0,5 para orifício/gotejador turbulento; perto de 0,1 para autocompensante; 1,0 para regime laminar. Define quanto a vazão varia com a pressão.")

    with st.expander("Topografia da adutora (ponto alto intermediário)", expanded=False):
        tem_crista = st.checkbox("A adutora passa por um ponto mais alto que o destino", value=False, key="pb_crista",
                                 help="Num ponto alto a pressão cai e o ar se acumula. O processador calcula a pressão ali.")
        cota_crista, dist_crista = float(desnivel_recalque), float(distancia_adutora) * 0.5
        if tem_crista:
            cc1, cc2 = st.columns(2)
            cota_crista = cc1.number_input("Cota do ponto alto (m acima do eixo da bomba):", min_value=0.0, value=round(float(desnivel_recalque) + 3.0, 1), step=0.5)
            dist_crista = cc2.number_input("Distância da bomba até o ponto alto (m):", min_value=1.0, max_value=float(max(distancia_adutora, 2.0)), value=round(float(distancia_adutora) * 0.5, 1), step=1.0)

    cf_b = bomba_cat_sel["cf"] if bomba_cat_sel is not None else proc_coefs_bomba(h0_b, pontos_b)
    qmax_b = proc_q_max(cf_b)
    curva_ok = proc_curva_valida(cf_b) and qmax_b is not None and all(q_ < qmax_b for q_, _ in pontos_b) and all(h0_b > h_ for _, h_ in pontos_b)
    if not curva_ok:
        st.error("A curva da bomba informada não é válida: a altura deve ser maior com o registro fechado do que nos pontos do catálogo e deve cair quando a vazão sobe. Revise a altura com registro fechado e o 2º ponto.")
    else:
        sis = dict(q_ref=float(vazao_setor_operando_m3h), L_rec=float(comp_virtual_adutora), D_rec=float(diam_adutora_int_mm), C_rec=float(c_rugosidade_pvc),
                   env=float(margem_envelhecimento) / 100.0, hf_suc=float(hf_perda_succao), hf_fil=float(hf_filtro_areia), hin=float(pressao_inicio_lateral),
                   x=float(x_emissor), h_est=float(desnivel_recalque + altura_succao))

        def ponto_completo(q, s_):
            h = proc_h_bomba(q, cf_b)
            eta = proc_eta(q, q_bep, eta_bep)
            p_eixo = proc_pot_eixo_kw(q, h, eta)
            hf_s = s_["hf_suc"] * (q / max(s_["q_ref"], 1e-9)) ** HW_N
            npshd = patm_local - altura_succao - hf_s - pvapor
            npshr = proc_npshr(q, q_bep, npsh_requerido_catalogo)
            p_desc = h - altura_succao - hf_s
            area = math.pi * (s_["D_rec"] / 1000.0) ** 2 / 4.0
            v = (q / 3600.0) / area if area > 0 else 0.0
            return dict(Q=q, H=h, eta=eta, p_eixo=p_eixo, p_cv=p_eixo / 0.7355, p_elet=p_eixo / eta_motor, npshd=npshd, npshr=npshr,
                        marg=npshd - npshr, p_desc=p_desc, v=v, dH=celeridade_onda_pvc * v / 9.81)

        q_op = proc_ponto_operacao(cf_b, qmax_b, sis)
        if q_op is None:
            st.error(f"A bomba não vence a carga estática do sistema ({_n(sis['h_est'], 1)} mca) com a altura de {_n(h0_b, 1)} mca de fechamento. Escolha uma bomba mais forte ou reduza o desnível.")
        else:
            op = ponto_completo(q_op, sis)
            p_shut = h0_b - altura_succao
            p_pico = op["p_desc"] + op["dH"]
            razao_bep = q_op / q_bep
            faixa_ok = 0.70 <= razao_bep <= 1.20
            marg_min = max(0.5, 0.2 * op["npshr"])
            npsh_ok = op["marg"] >= marg_min
            desvio_q = (q_op - vazao_setor_operando_m3h) / vazao_setor_operando_m3h * 100.0
            desvio_txt = ("+" if desvio_q >= 0 else "") + _n(desvio_q, 1) + "%"
            vazao_ok = abs(desvio_q) <= 10.0
            pn_ok = (p_shut <= pn_adutora) and (p_pico <= pn_adutora)
            vel_ok = 0.6 <= op["v"] <= 2.2
            hf_adut_op = proc_hw_r(sis["L_rec"], sis["D_rec"], sis["C_rec"]) * ((op["Q"] / 3600.0) ** HW_N) * (1.0 + sis["env"])
            p_crista = op["p_desc"] - hf_adut_op * min(1.0, dist_crista / max(float(distancia_adutora), 1e-6)) - cota_crista
            motor_ok = potencia_cv_comercial >= op["p_cv"] * 1.0

            # ------------------------------------------------ 2. PONTO DE OPERAÇÃO
            st.markdown("#### 2. Ponto de Operação (onde a bomba e o sistema se equilibram)")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Vazão de operação", f"{_n(op['Q'], 2)} m³/h", f"{desvio_txt} vs projeto ({_n(vazao_setor_operando_m3h, 2)})")
            m2.metric("Altura de operação", f"{_n(op['H'], 1)} mca", f"AMT da aba 3: {_n(amt_total, 1)} mca", delta_color="off")
            m3.metric("Rendimento da bomba", f"{op['eta'] * 100:.0f}%")
            m4.metric("Potência no eixo", f"{_n(op['p_cv'], 2)} cv", f"{_n(op['p_eixo'], 2)} kW")
            m5, m6, m7, m8 = st.columns(4)
            m5.metric("Posição na curva (Q/Qbep)", f"{razao_bep * 100:.0f}%", "faixa 70 a 120%", delta_color="off")
            m6.metric("NPSH disponível", f"{_n(op['npshd'], 2)} mca", f"requerido: {_n(op['npshr'], 2)} mca", delta_color="off")
            m7.metric("Pressão na saída da bomba", f"{_n(op['p_desc'], 1)} mca", f"fechada: {_n(p_shut, 1)} mca", delta_color="off")
            m8.metric("Velocidade na adutora", f"{_n(op['v'], 2)} m/s")

            if bomba_cat_sel is not None and not (float(bomba_cat_sel["q"][0]) <= q_op <= float(bomba_cat_sel["q"][-1])):
                st.warning(f"O ponto de operação ({_n(q_op, 2)} m³/h) fica fora da faixa tabelada desta bomba ({_n(bomba_cat_sel['q'][0], 2)} a {_n(bomba_cat_sel['q'][-1], 2)} m³/h). "
                           "O resultado depende da extrapolação da curva: confirme com o fabricante.")
            qs = np.linspace(0.02 * qmax_b, qmax_b, 160)
            fig_ps = make_subplots(specs=[[{"secondary_y": True}]])
            fig_ps.add_vrect(x0=0.7 * q_bep, x1=1.2 * q_bep, fillcolor="#27AE60", opacity=0.10, line_width=0, annotation_text="faixa recomendada", annotation_position="top left")
            fig_ps.add_trace(go.Scatter(x=qs, y=[proc_h_bomba(q, cf_b) for q in qs], name="Curva da bomba", line=dict(color="#1F77B4", width=3)), secondary_y=False)
            if bomba_cat_sel is not None:
                fig_ps.add_trace(go.Scatter(x=bomba_cat_sel["q"], y=bomba_cat_sel["h"], mode="markers", name="Pontos do catálogo",
                                            marker=dict(color="#1F77B4", size=7, symbol="diamond")), secondary_y=False)
            qs_sis = np.linspace(0.0, qmax_b, 160)
            fig_ps.add_trace(go.Scatter(x=qs_sis, y=[proc_h_sistema(q, sis) for q in qs_sis], name="Curva do sistema", line=dict(color="#E67E22", width=3)), secondary_y=False)
            fig_ps.add_trace(go.Scatter(x=qs, y=[proc_eta(q, q_bep, eta_bep) * 100 for q in qs], name="Rendimento (%)", line=dict(color="#27AE60", width=2, dash="dash")), secondary_y=True)
            fig_ps.add_trace(go.Scatter(x=[op["Q"]], y=[op["H"]], mode="markers+text", text=["operação"], textposition="top right", name="Ponto de operação",
                                        marker=dict(size=14, color="#C0392B", line=dict(color="white", width=2))), secondary_y=False)
            fig_ps.add_trace(go.Scatter(x=[vazao_setor_operando_m3h], y=[amt_total], mode="markers+text", text=["projeto (aba 3)"], textposition="bottom right", name="Ponto de projeto",
                                        marker=dict(size=11, symbol="diamond", color="#2C3E50")), secondary_y=False)
            fig_ps.update_layout(height=470, margin=dict(l=10, r=10, t=30, b=10), legend=dict(orientation="h", y=-0.2), xaxis_title="Vazão (m³/h)")
            fig_ps.update_yaxes(title_text="Altura (mca)", secondary_y=False, rangemode="tozero")
            fig_ps.update_yaxes(title_text="Rendimento (%)", secondary_y=True, range=[0, 100], tickvals=[0, 20, 40, 60, 80, 100], showgrid=False)
            st.plotly_chart(fig_ps, use_container_width=True)

            fig_np = go.Figure()
            fig_np.add_trace(go.Scatter(x=qs, y=[ponto_completo(q, sis)["npshd"] for q in qs], name="NPSH disponível", line=dict(color="#1F77B4", width=3)))
            fig_np.add_trace(go.Scatter(x=qs, y=[proc_npshr(q, q_bep, npsh_requerido_catalogo) for q in qs], name="NPSH requerido (estimado)", line=dict(color="#C0392B", width=3, dash="dash")))
            fig_np.add_trace(go.Scatter(x=[op["Q"]], y=[op["npshd"]], mode="markers", name="Operação", marker=dict(size=12, color="#C0392B", line=dict(color="white", width=2))))
            fig_np.update_layout(height=320, margin=dict(l=10, r=10, t=30, b=10), xaxis_title="Vazão (m³/h)", yaxis_title="NPSH (mca)", legend=dict(orientation="h", y=-0.3),
                                 title="Cavitação ao longo da curva: o risco aumenta quando a vazão sobe")
            st.plotly_chart(fig_np, use_container_width=True)

            # ------------------------------------------------ 3. COMPARATIVO DE SETORIZAÇÃO
            st.markdown("#### 3. Comparativo de Setorização com a Mesma Bomba")
            st.caption("Quanto mais setores, menor a vazão simultânea e a potência, porém mais horas de bombeamento por rega. A tabela mostra o equilíbrio real da bomba escolhida em cada divisão.")
            linhas_cmp = []
            custo_atual, kwh_mes_atual = None, None
            for n_s in (1, 2, 3, 4, 5, 6):
                per_set = max(1, math.ceil(num_total_linhas_campo / n_s))
                q_ref_n = vazao_linha_lateral_m3h * per_set
                s_n = dict(sis, q_ref=q_ref_n)
                q_n = proc_ponto_operacao(cf_b, qmax_b, s_n)
                if q_n is None:
                    linhas_cmp.append({"Setores": n_s, "Linhas/setor": per_set, "Q projeto (m³/h)": q_ref_n, "Situação": "Bomba não vence a carga estática"})
                    continue
                o_n = ponto_completo(q_n, s_n)
                taxa_n = taxa_aplicacao_real * (q_n / q_ref_n) if q_ref_n > 0 else 0.0
                h_set = (lamina_bruta_diaria * turno_rega / taxa_n) if taxa_n > 0 else 0.0
                h_evento = h_set * n_s
                kwh_ev = o_n["p_elet"] * h_evento
                custo_mes = kwh_ev * (30.0 / max(1, turno_rega)) * tarifa_kwh
                if n_s == num_setores:
                    custo_atual, kwh_mes_atual = custo_mes, kwh_ev * (30.0 / max(1, turno_rega))
                sit = "OK"
                if not (0.7 <= q_n / q_bep <= 1.2):
                    sit = "Fora da faixa da bomba"
                elif o_n["marg"] < max(0.5, 0.2 * o_n["npshr"]):
                    sit = "NPSH insuficiente"
                elif abs(q_n - q_ref_n) / q_ref_n > 0.10:
                    sit = "Vazão diverge do projeto"
                linhas_cmp.append({"Setores": n_s, "Linhas/setor": per_set, "Q projeto (m³/h)": round(q_ref_n, 2), "Q operação (m³/h)": round(q_n, 2),
                                   "Altura (mca)": round(o_n["H"], 1), "Rendimento (%)": round(o_n["eta"] * 100), "Potência eixo (kW)": round(o_n["p_eixo"], 2),
                                   "Horas por rega": round(h_evento, 1), "kWh por rega": round(kwh_ev, 1), "Custo de energia (R$/mês)": round(custo_mes, 2), "Situação": sit,
                                   "Atual": "◄" if n_s == num_setores else ""})
            df_cmp = pd.DataFrame(linhas_cmp)
            st.dataframe(df_cmp, hide_index=True, use_container_width=True)
            if "Custo de energia (R$/mês)" in df_cmp.columns:
                fig_cmp = go.Figure(go.Bar(x=[f"{int(r['Setores'])} setor(es)" for _, r in df_cmp.iterrows() if pd.notna(r.get("Custo de energia (R$/mês)"))],
                                           y=[r["Custo de energia (R$/mês)"] for _, r in df_cmp.iterrows() if pd.notna(r.get("Custo de energia (R$/mês)"))], marker_color="#1F77B4"))
                fig_cmp.update_layout(height=280, margin=dict(l=10, r=10, t=30, b=10), yaxis_title="R$/mês", title="Custo mensal de energia por divisão da roça")
                st.plotly_chart(fig_cmp, use_container_width=True)

            # ------------------------------------------------ 4. REDE
            st.markdown("#### 4. Solver de Rede (todos os trechos, nós e setores)")
            st.caption("A tabela é montada automaticamente a partir do projeto e pode ser editada: altere comprimentos, diâmetros ou adicione trechos paralelos e malhas. "
                       "Cotas em relação ao eixo da bomba (z = 0). Perda localizada: h = k·Q², com a perda informada na vazão de projeto.")
            n_set_r = int(num_setores)
            setores_ids = [f"S{k}" for k in range(1, n_set_r + 1)]
            abertos = st.multiselect("Setores abertos (em irrigação):", setores_ids, default=setores_ids[:1], key="proc_abertos")
            nos_rows = [
                {"Nó": "Cisterna", "Cota z (m)": -float(altura_succao), "Carga fixa (mca)": -float(altura_succao), "Demanda (m³/h)": 0.0, "P nominal (mca)": 0.0, "Expoente x": 0.0},
                {"Nó": "Succao", "Cota z (m)": 0.0, "Carga fixa (mca)": None, "Demanda (m³/h)": 0.0, "P nominal (mca)": 0.0, "Expoente x": 0.0},
                {"Nó": "Recalque", "Cota z (m)": 0.0, "Carga fixa (mca)": None, "Demanda (m³/h)": 0.0, "P nominal (mca)": 0.0, "Expoente x": 0.0},
                {"Nó": "Filtro_ent", "Cota z (m)": float(desnivel_recalque), "Carga fixa (mca)": None, "Demanda (m³/h)": 0.0, "P nominal (mca)": 0.0, "Expoente x": 0.0},
                {"Nó": "Cabecote", "Cota z (m)": float(desnivel_recalque), "Carga fixa (mca)": None, "Demanda (m³/h)": 0.0, "P nominal (mca)": 0.0, "Expoente x": 0.0},
            ]
            for k in range(1, n_set_r + 1):
                linhas_k = max(0, min(num_linhas_por_setor, num_total_linhas_campo - (k - 1) * num_linhas_por_setor))
                nos_rows.append({"Nó": f"S{k}", "Cota z (m)": float(desnivel_recalque), "Carga fixa (mca)": None, "Demanda (m³/h)": round(vazao_linha_lateral_m3h * linhas_k, 4),
                                 "P nominal (mca)": round(float(pressao_inicio_lateral), 2), "Expoente x": float(x_emissor)})
            L_man = round(max(1.0, num_linhas_por_setor * espacamento_linhas), 1)
            tr_rows = [
                {"Trecho": "succao", "De": "Cisterna", "Para": "Succao", "Tipo": "Perda localizada", "L (m)": 0.0, "D (mm)": float(diam_succao_mm), "C": float(c_rugosidade_pvc), "Perda @Q projeto (mca)": float(hf_perda_succao)},
                {"Trecho": "bomba", "De": "Succao", "Para": "Recalque", "Tipo": "Bomba", "L (m)": 0.0, "D (mm)": 0.0, "C": 0.0, "Perda @Q projeto (mca)": 0.0},
                {"Trecho": "adutora", "De": "Recalque", "Para": "Filtro_ent", "Tipo": "Tubo", "L (m)": round(float(comp_virtual_adutora) * (1 + margem_envelhecimento / 100.0), 1),
                 "D (mm)": float(diam_adutora_int_mm), "C": float(c_rugosidade_pvc), "Perda @Q projeto (mca)": 0.0},
                {"Trecho": "filtro", "De": "Filtro_ent", "Para": "Cabecote", "Tipo": "Perda localizada", "L (m)": 0.0, "D (mm)": 0.0, "C": 0.0, "Perda @Q projeto (mca)": float(hf_filtro_areia)},
            ]
            anterior = "Cabecote"
            for k in range(1, n_set_r + 1):
                tr_rows.append({"Trecho": f"deriv_S{k}", "De": anterior, "Para": f"S{k}", "Tipo": "Tubo", "L (m)": L_man, "D (mm)": float(diam_adutora_int_mm),
                                "C": float(c_rugosidade_pvc), "Perda @Q projeto (mca)": 0.0})
                anterior = f"S{k}"
            df_nos0, df_tr0 = pd.DataFrame(nos_rows), pd.DataFrame(tr_rows)
            chave = hashlib.md5((df_nos0.to_csv() + df_tr0.to_csv()).encode()).hexdigest()[:10]
            ed1, ed2 = st.columns([1, 1.25])
            with ed1:
                st.markdown("**Nós**")
                df_nos = st.data_editor(df_nos0, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"ed_nos_{chave}")
            with ed2:
                st.markdown("**Trechos**")
                df_tr = st.data_editor(df_tr0, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"ed_tr_{chave}",
                                       column_config={"Tipo": st.column_config.SelectboxColumn("Tipo", options=["Tubo", "Perda localizada", "Bomba"], required=True)})
            try:
                nos_s, el_s = [], []
                for _, r_ in df_nos.dropna(subset=["Nó"]).iterrows():
                    fx = r_.get("Carga fixa (mca)")
                    fechado = (str(r_["Nó"]) in setores_ids) and (str(r_["Nó"]) not in abertos)
                    nos_s.append(dict(id=str(r_["Nó"]), z=float(r_["Cota z (m)"]), H_fixo=None if pd.isna(fx) else float(fx),
                                      d_nom=0.0 if fechado else float(r_["Demanda (m³/h)"] or 0.0), p_nom=float(r_["P nominal (mca)"] or 0.0), x=float(r_["Expoente x"] or 0.0)))
                tipos = {"Tubo": "tubo", "Perda localizada": "perda", "Bomba": "bomba"}
                for _, r_ in df_tr.dropna(subset=["Trecho"]).iterrows():
                    el_s.append(dict(id=str(r_["Trecho"]), de=str(r_["De"]), para=str(r_["Para"]), tipo=tipos[r_["Tipo"]], L=float(r_["L (m)"] or 0.0), D=float(r_["D (mm)"] or 0.0),
                                     C=float(r_["C"] or 0.0), hf_nom=float(r_["Perda @Q projeto (mca)"] or 0.0), q_ref=float(vazao_setor_operando_m3h), cf=cf_b))
                rede = proc_resolver_rede(nos_s, el_s)
            except Exception as ex_:
                rede = {"ok": False, "msg": f"Entrada inválida na tabela: {ex_}"}

            if not rede.get("ok"):
                st.warning(f"O solver não convergiu: {rede.get('msg', '')} Verifique se todos os nós estão ligados à cisterna e se os nomes de 'De' e 'Para' existem na tabela de nós.")
            else:
                q_pump_rede = rede["Q_m3h"].get("bomba", float("nan"))
                z_no = {n["id"]: n["z"] for n in nos_s}
                dem_aberta = sum(rede["demanda_m3h"][s_] for s_ in abertos if s_ in rede["demanda_m3h"])
                pr_set = {s_: rede["H"][s_] - z_no[s_] for s_ in abertos if s_ in rede["H"]}
                p_min_set = min(pr_set.values()) if pr_set else float("nan")
                rr1, rr2, rr3, rr4 = st.columns(4)
                rr1.metric("Vazão da bomba (rede)", f"{_n(q_pump_rede, 2)} m³/h", f"curva simples, 1 setor: {_n(op['Q'], 2)}", delta_color="off")
                rr2.metric("Vazão entregue aos setores abertos", f"{_n(dem_aberta, 2)} m³/h")
                rr3.metric("Menor pressão de entrada nos setores", f"{_n(p_min_set, 1)} mca", f"necessária: {_n(pressao_inicio_lateral, 1)} mca", delta_color="off")
                no_suc = next((n_["id"] for n_ in nos_s if n_["id"].lower().startswith("suc")), None)
                npshd_rede = (patm_local + (rede["H"][no_suc] - z_no[no_suc]) - pvapor) if no_suc else float("nan")
                npshr_rede = proc_npshr(q_pump_rede, q_bep, npsh_requerido_catalogo)
                rr4.metric("NPSH disponível pela rede", f"{_n(npshd_rede, 2)} mca", f"requerido: {_n(npshr_rede, 2)} mca", delta_color="off")
                st.caption(f"Convergiu em {rede['iter']} iterações (resíduo {rede['residuo']:.1e}). A curva simples considera um único setor e não inclui o trecho de derivação até os setores; "
                           "com vários setores abertos ao mesmo tempo, vale o resultado da rede.")
                if len(abertos) != 1:
                    st.info(f"{len(abertos)} setor(es) abertos ao mesmo tempo: a demanda total é maior que a de projeto (1 setor). Confira abaixo se a pressão nos setores continua suficiente.")

                linhas_n = []
                for n in nos_s:
                    pr = rede["H"][n["id"]] - n["z"]
                    obs = ""
                    if n["id"] in setores_ids and n["id"] in abertos:
                        obs = "OK" if pr >= 0.95 * pressao_inicio_lateral else "Pressão abaixo da necessária"
                    elif pr < 0:
                        obs = "Pressão negativa (vácuo)" if n["id"].lower().startswith("suc") else "Pressão negativa: risco de colapso e entrada de ar"
                    linhas_n.append({"Nó": n["id"], "Cota (m)": round(n["z"], 2), "Carga (mca)": round(rede["H"][n["id"]], 2), "Pressão (mca)": round(pr, 2),
                                     "Vazão consumida (m³/h)": round(rede["demanda_m3h"][n["id"]], 3), "Observação": obs})
                st.markdown("**Resultados nos nós**")
                st.dataframe(pd.DataFrame(linhas_n), hide_index=True, use_container_width=True)

                linhas_t = []
                for e in el_s:
                    qh = rede["Q_m3h"][e["id"]]
                    if e["tipo"] == "tubo" and e["D"] > 0:
                        area = math.pi * (e["D"] / 1000.0) ** 2 / 4.0
                        v = abs(qh) / 3600.0 / area
                        obs = "" if abs(qh) < 1e-3 else ("OK" if 0.6 <= v <= 2.2 else ("Velocidade baixa" if v < 0.6 else "Velocidade alta"))
                        linhas_t.append({"Trecho": e["id"], "Tipo": "Tubo", "Vazão (m³/h)": round(qh, 3), "Velocidade (m/s)": round(v, 2),
                                         "Perda (mca)": round(rede["hf"][e["id"]], 3), "Gradiente (m/100 m)": round(rede["hf"][e["id"]] / e["L"] * 100, 2) if e["L"] > 0 else None, "Observação": obs})
                    elif e["tipo"] == "bomba":
                        linhas_t.append({"Trecho": e["id"], "Tipo": "Bomba", "Vazão (m³/h)": round(qh, 3), "Perda (mca)": round(rede["hf"][e["id"]], 3), "Observação": "ganho de carga"})
                    else:
                        linhas_t.append({"Trecho": e["id"], "Tipo": "Perda localizada", "Vazão (m³/h)": round(qh, 3), "Perda (mca)": round(rede["hf"][e["id"]], 3)})
                st.markdown("**Resultados nos trechos**")
                st.dataframe(pd.DataFrame(linhas_t), hide_index=True, use_container_width=True)

                # ---- Linha piezométrica ao longo do caminho cisterna -> setor aberto mais distante
                try:
                    fonte = next(n["id"] for n in nos_s if n["H_fixo"] is not None)
                    adj = {}
                    for e in el_s:
                        adj.setdefault(e["de"], []).append((e["para"], e))
                        adj.setdefault(e["para"], []).append((e["de"], e))
                    pai = {fonte: None}
                    fila = [fonte]
                    while fila:
                        u = fila.pop(0)
                        for v_, e_ in adj.get(u, []):
                            if v_ not in pai:
                                pai[v_] = (u, e_)
                                fila.append(v_)
                    alvo = None
                    for s_ in abertos:
                        if s_ in pai:
                            alvo = s_
                    if alvo:
                        cam = []
                        no_ = alvo
                        while no_ is not None:
                            cam.append(no_)
                            no_ = pai[no_][0] if pai[no_] else None
                        cam = cam[::-1]
                        xs, zs, hs = [0.0], [z_no[cam[0]]], [rede["H"][cam[0]]]
                        for a_, b_ in zip(cam[:-1], cam[1:]):
                            e_ = pai[b_][1]
                            xs.append(xs[-1] + (e_["L"] if e_["tipo"] == "tubo" else 0.0))
                            zs.append(z_no[b_])
                            hs.append(rede["H"][b_])
                        fig_lp = go.Figure()
                        fig_lp.add_trace(go.Scatter(x=xs, y=zs, name="Terreno / eixo", fill="tozeroy", line=dict(color="#8D6E63", width=2), fillcolor="rgba(141,110,99,0.25)"))
                        fig_lp.add_trace(go.Scatter(x=xs, y=hs, name="Linha piezométrica", line=dict(color="#1F77B4", width=3)))
                        fig_lp.add_trace(go.Scatter(x=xs, y=[z + pn_adutora for z in zs], name=f"Limite de pressão (PN {pn_adutora:.0f} mca)", line=dict(color="#C0392B", width=2, dash="dash")))
                        for xx, nn in zip(xs, cam):
                            fig_lp.add_annotation(x=xx, y=rede["H"][nn], text=nn, showarrow=False, yshift=14, font=dict(size=10))
                        fig_lp.update_layout(height=380, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="Distância ao longo do caminho (m)", yaxis_title="Cota / carga (m)",
                                             title=f"Linha piezométrica: cisterna até {alvo}", legend=dict(orientation="h", y=-0.25))
                        st.plotly_chart(fig_lp, use_container_width=True)
                except Exception:
                    pass

            # ------------------------------------------------ 5. PARECER
            st.markdown("#### 5. Parecer Automático do Processador")
            marg_txt = f"{_n(op['marg'], 2)} mca (mínimo recomendado {_n(marg_min, 2)} mca)"
            parecer_rows = [
                {"Verificação": "Vazão de operação x vazão de projeto", "Resultado": f"{_n(op['Q'], 2)} m³/h contra {_n(vazao_setor_operando_m3h, 2)} m³/h ({desvio_txt})",
                 "Situação": "OK" if vazao_ok else ("Bomba entrega menos que o projeto" if desvio_q < 0 else "Bomba entrega mais que o projeto: sobrepressão e consumo excessivo")},
                {"Verificação": "Faixa de operação da bomba (70 a 120% do ponto de melhor eficiência)", "Resultado": f"{razao_bep * 100:.0f}% do ponto nominal", "Situação": "OK" if faixa_ok else "Fora da faixa: vibração, recirculação e desgaste"},
                {"Verificação": "NPSH no ponto de operação", "Resultado": marg_txt, "Situação": "OK" if npsh_ok else "Margem insuficiente: risco de cavitação"},
                {"Verificação": "Pressão com registro fechado e aríete x PN da adutora", "Resultado": f"Fechada {_n(p_shut, 1)} mca; operação + aríete {_n(p_pico, 1)} mca; PN {_n(pn_adutora, 0)} mca",
                 "Situação": "OK" if pn_ok else "Acima da classe de pressão do tubo"},
                {"Verificação": "Velocidade na adutora no ponto de operação", "Resultado": f"{_n(op['v'], 2)} m/s", "Situação": "OK" if vel_ok else "Fora de 0,6 a 2,2 m/s"},
                {"Verificação": "Motor comercial x potência no eixo", "Resultado": f"Motor {_n(potencia_cv_comercial, 1)} cv; eixo {_n(op['p_cv'], 2)} cv", "Situação": "OK" if motor_ok else "Motor subdimensionado"},
            ]
            if tem_crista:
                parecer_rows.append({"Verificação": "Pressão no ponto alto da adutora", "Resultado": f"{_n(p_crista, 1)} mca a {_n(cota_crista, 1)} m de cota",
                                     "Situação": ("OK: instale ventosa de tríplice função no ponto alto (boa prática)" if p_crista >= 2.0 else
                                                  ("Pressão muito baixa: ventosa obrigatória e revisar traçado" if p_crista >= 0 else "Pressão negativa: risco de colapso do tubo e entrada de ar; revise o traçado"))})
            if rede.get("ok"):
                parecer_rows.append({"Verificação": "Pressão de entrada nos setores abertos (rede)", "Resultado": f"Mínima de {_n(p_min_set, 1)} mca; necessária {_n(pressao_inicio_lateral, 1)} mca",
                                     "Situação": "OK" if p_min_set >= 0.95 * pressao_inicio_lateral else "Abaixo da necessária: abra menos setores ao mesmo tempo ou use bomba maior"})
                marg_rede = npshd_rede - npshr_rede
                parecer_rows.append({"Verificação": "NPSH pela rede (pressão real na sucção)", "Resultado": f"Disponível {_n(npshd_rede, 2)} mca; requerido {_n(npshr_rede, 2)} mca",
                                     "Situação": "OK" if marg_rede >= max(0.5, 0.2 * npshr_rede) else "Margem insuficiente: risco de cavitação"})
            st.dataframe(pd.DataFrame(parecer_rows), hide_index=True, use_container_width=True)
            if not curva_real:
                st.warning("A curva da bomba é uma estimativa típica. Marque \"Informar os pontos do catálogo do fabricante\" e digite os valores reais para validar este parecer.")
            if amt_req - amt_total > 0.5:
                st.info(f"Carga exigida pelo sistema na vazão de projeto: {_n(amt_req, 1)} mca, contra {_n(amt_total, 1)} mca de AMT na aba 3. "
                        f"A diferença ({_n(amt_req - amt_total, 1)} mca) é a altura de sucção ({_n(altura_succao, 1)} m) mais a perda na sucção ({_n(hf_perda_succao, 1)} mca), que a aba 3 não soma "
                        f"(se o desnível informado foi medido a partir do eixo da bomba). Dimensione a bomba pelo valor do processador.")

            proc_resumo.update(dict(
                ok=True, curva_real=curva_real, q_op=op["Q"], h_op=op["H"], eta=op["eta"], p_cv=op["p_cv"], razao_bep=razao_bep, faixa_ok=faixa_ok,
                marg=op["marg"], marg_min=marg_min, npsh_ok=npsh_ok, p_shut=p_shut, p_pico=p_pico, pn_ok=pn_ok, vazao_ok=vazao_ok, desvio_q=desvio_q, desvio_txt=desvio_txt,
                tem_crista=tem_crista, p_crista=p_crista, cota_crista=cota_crista, custo_energia_mes=custo_atual, kwh_mes=kwh_mes_atual, tarifa_kwh=tarifa_kwh, parecer=parecer_rows))

    with st.expander("Memória de cálculo e método do processador"):
        st.markdown(r"""
**Curva da bomba:** $H(Q)=H_0+a_1Q+a_2Q^2$ ajustada ao ponto nominal (e a um 2º ponto, se informado); com um só ponto, $H=H_0-cQ^2$, sempre decrescente.
**Rendimento:** parábola com máximo no ponto nominal, $\eta=\eta_b\,(2x-x^2)$, com $x=Q/Q_b$. **NPSH requerido:** estimado como $NPSH_r=NPSH_{r,b}\,(0{,}3+0{,}7x^2)$; use a curva do fabricante quando houver.
**Curva do sistema:** $H_{sis}(Q)=H_{est}+h_{suc}+h_{adut}+h_{filtro}+H_{in}$, com perda contínua de Hazen-Williams $h=10{,}67\,L\,Q^{1{,}852}/(C^{1{,}852}D^{4{,}87})$,
perdas localizadas quadráticas e pressão de entrada da lateral $H_{in}=H_{in,0}\,(Q/Q_0)^{1/x}$ (emissor $q=kH^x$). O ponto de operação é a interseção com a curva da bomba, obtida por bisecção.
**Solver de rede:** método de Newton aplicado ao sistema completo de equações de energia (cada trecho) e continuidade (cada nó), conhecido como algoritmo do gradiente global (Todini e Pilati, 1988), com amortecimento por busca linear.
Os setores são nós de demanda dependente da pressão. **Validação:** o solver reproduz a solução analítica de tubo único e em paralelo e coincide com o método de Hardy-Cross em malha.
**Limites:** regime permanente (não calcula transitórios dinâmicos; o aríete segue a fórmula de Joukowsky), água a temperatura ambiente, rendimento e NPSHr aproximados quando não há curva de catálogo.
""")


# ------------------------------------------------------------------------------
# 6. ABA: MEMORIAL TÉCNICO E LAUDO EXECUTIVO DE ENGENHARIA (COM EXPORTAÇÃO)
# ------------------------------------------------------------------------------
with tab_laudo:
    st.subheader("Memorial de Cálculo Analítico & Laudo Técnico de Executabilidade")
    
    exige_outorga = vazao_setor_operando_m3h > 5.0
    status_outorga_txt = "EXIGE PROCESSO DE OUTORGA" if exige_outorga else "DISPENSA DE OUTORGA (Uso Insignificante)"
    
    criterios = [
        {"item": "1. Infiltração Solo x Água", "norma": "Princípios Agronômicos", "resultado": f"Ta = {taxa_aplicacao_real:.2f} mm/h vs VIB = {vib_solo:.1f} mm/h", "aprovado": not risco_escorrimento, "obs": "Sem escorrimento superficial"},
        {"item": "2. Uniformidade da Linha", "norma": "Norma Hidráulica", "resultado": f"ΔP = {hf_lateral_corrigida:.2f} mca (Max: {limite_var_pressao:.2f} mca)", "aprovado": atende_iso9261, "obs": "Variação de pressão ≤ 20%"},
        {"item": "3. Pressão Térmica Lateral", "norma": "Norma Hidráulica", "resultado": f"Hin = {pressao_inicio_lateral:.1f} mca ≤ PN Real = {pn_real_lateral:.1f} mca", "aprovado": seguro_pressao_lateral, "obs": "Tubo resiste ao calor local"},
        {"item": "4. Autolimpeza das Linhas", "norma": "Norma Hidráulica", "resultado": f"Vazão de Flushing = {vazao_flushing_m3h:.3f} m³/h", "aprovado": True, "obs": "Garante velocidade mínima"},
        {"item": "5. Risco de Golpe de Aríete", "norma": "Equação Fluido-dinâmica", "resultado": f"Pico = {pressao_maxima_transiente:.1f} mca ≤ PN {pn_adutora:.0f} mca", "aprovado": seguro_ariete, "obs": "Adutora resiste a fechamento brusco"},
        {"item": "6. Velocidade na Adutora", "norma": "Critério de Engenharia", "resultado": f"V = {vel_adutora:.2f} m/s", "aprovado": (0.6 <= vel_adutora <= 2.2), "obs": "Escoamento sem golpe nem atrito excessivo"},
        {"item": "7. Proteção contra Cavitação", "norma": "Critério Mecânico", "resultado": f"NPSHd = {npsh_disponivel:.2f} mca > NPSHr = {npsh_requerido_catalogo:.1f} mca", "aprovado": seguro_cavitacao, "obs": "Folga de sucção garantida"},
        {"item": "8. Regra de Ouro da Sucção", "norma": "Boa Prática", "resultado": f"Dsuc ({diam_succao_mm:.0f}mm) ≥ Drec ({diam_adutora_int_mm:.0f}mm)", "aprovado": regra_ouro_diametros, "obs": "Sucção dimensionada"},
        {"item": "9. Queda de Tensão Elétrica", "norma": "Norma Elétrica", "resultado": f"Queda = {queda_efetiva_perc:.2f}% (Max: {queda_max_perm:.1f}%)", "aprovado": (queda_efetiva_perc <= queda_max_perm), "obs": f"Cabo de cobre {_n(bitola_escolhida, 1)} mm²"},
        {"item": "10. Outorga de Água", "norma": "Legislação Ambiental", "resultado": f"Vazão de Captação = {vazao_setor_operando_m3h:.2f} m³/h", "aprovado": not exige_outorga, "obs": status_outorga_txt}
    ]
    
    eh_aspersao = metodo_selecionado == "Aspersão"
    if eh_aspersao:
        criterios[1]["norma"] = "Variação de pressão ≤ 20 % da pressão de serviço"
        criterios[3].update({"norma": "Não aplicável", "resultado": "Não se aplica à aspersão", "obs": "Critério de linhas de gotejamento"})
    total_aprovados = sum(1 for c in criterios if c["aprovado"])
    status_geral_projeto = "PROJETO TOTALMENTE EXECUTÁVEL E APROVADO" if total_aprovados >= 9 else "PROJETO EXECUTÁVEL COM RESSALVAS TÉCNICAS"
    if total_aprovados < 7:
        status_geral_projeto = "PROJETO REPROVADO - NECESSITA AJUSTES DE DIMENSIONAMENTO"
        
    st.markdown(f"### Veredito Geral: **{status_geral_projeto}** ({total_aprovados} de 10 critérios conformes)")
    
    # --- Matriz de Compliance Normativa (Motor de Cálculo, aba "6. Compliance e Normas") ---
    st.markdown("#### Matriz de Compliance e Certificação Normativa")
    status_iso9261 = "CONFORME" if (seguro_pressao_lateral and atende_iso9261) else "RESSALVA"
    status_nbr5410 = f"OK - Seção do cabo: {_n(bitola_escolhida, 1)} mm²" if queda_efetiva_perc <= queda_max_perm else f"RESSALVA - Queda de {queda_efetiva_perc:.2f}% (Seção: {_n(bitola_escolhida, 1)} mm²)"
    status_nbr15527 = f"CONFORME - First Flush de {vol_first_flush:,.0f} L descartados; {status_transbordo}"
    compliance = [
        {"norma": "ABNT NBR ISO 9261 (Emissores e Uniformidade)", "status": status_iso9261, "criterio": "Validação de Pressão (Hin ≤ PN Real) e variação de pressão ≤ 20%"},
        {"norma": "ABNT NBR 5410 (Instalações Elétricas BT)", "status": status_nbr5410, "criterio": "Queda de tensão máxima respeitada"},
        {"norma": "ABNT NBR 15527 (Aproveitamento de Água Pluvial)", "status": status_nbr15527, "criterio": "First Flush calculado e Ladrão validado"},
        {"norma": f"Outorga de Uso de Água (órgão estadual de {uf_selecionado} / ANA)", "status": status_outorga_txt, "criterio": "Baseado na Vazão Total (Q) > 5 m³/h"},
    ]
    if eh_aspersao:
        compliance[0] = {"norma": "Uniformidade de pressão na lateral (aspersão)", "status": status_iso9261,
                         "criterio": "Pressão inicial ≤ PN real do tubo e variação de pressão ≤ 20 % da pressão de serviço (Bernardo et al.)"}
    st.dataframe(
        pd.DataFrame(compliance).rename(columns={"norma": "Norma / Legislação", "status": "Status do Projeto", "criterio": "Critério de Validação"}),
        hide_index=True, use_container_width=True
    )
    
    # --- Verificações complementares do Motor de Cálculo ---
    st.markdown("#### Verificações Complementares do Motor de Cálculo")
    verificacoes_extra = [
        {"Verificação": "Flushing vs Vazão Normal da Linha", "Resultado": f"Flushing = {vazao_flushing_m3h:.3f} m³/h vs Q linha = {vazao_linha_lateral_m3h:.3f} m³/h", "Status": status_flushing},
        {"Verificação": "Transbordo da Cisterna (Ladrão)", "Resultado": f"Volume líquido = {vol_liquido_aproveitavel:,.0f} L vs Cisterna = {vol_cisterna_existente:,.0f} L", "Status": status_transbordo},
        {"Verificação": "Risco de Falta d'Água (Rede Pública)", "Resultado": f"Rede = {vazao_rede_concessionaria:,.0f} L/h vs Demanda média = {demanda_horaria_media:,.0f} L/h", "Status": status_falta_agua},
        {"Verificação": "Tempo de Reposição via Rede Pública", "Resultado": f"{tempo_encher_cisterna_horas:.1f} h vs Turno de rega = {turno_rega * 24} h", "Status": "ALERTA: Reposição lenta" if reposicao_lenta else "OK"},
    ]
    if eh_aspersao:
        verificacoes_extra = [v_ for v_ in verificacoes_extra if not v_["Verificação"].startswith("Flushing")]
    verificacoes_extra.append({"Verificação": "Qualidade da água e risco de entupimento", "Resultado": f"Risco {risco_entupimento.lower()}" + (f" ({'; '.join(parametros_criticos_qa)})" if parametros_criticos_qa else "") + f"; filtragem sugerida ~{malha_sugerida} mesh",
                               "Status": "OK" if risco_entupimento == "Baixo" else f"ALERTA: risco {risco_entupimento.lower()}, exige filtragem e manejo reforçados"})
    if proc_resumo.get("ok"):
        _aviso_curva = "" if proc_resumo["curva_real"] else " (curva da bomba estimada: confirmar com o catálogo)"
        verificacoes_extra += [
            {"Verificação": "Ponto de operação bomba x sistema", "Resultado": f"Q = {_n(proc_resumo['q_op'], 2)} m³/h a H = {_n(proc_resumo['h_op'], 1)} mca (projeto: {_n(vazao_setor_operando_m3h, 2)} m³/h, {proc_resumo['desvio_txt']})",
             "Status": ("OK" if proc_resumo["vazao_ok"] else "ALERTA: vazão diverge do projeto") + _aviso_curva},
            {"Verificação": "Faixa de operação da bomba (70 a 120% do BEP)", "Resultado": f"{proc_resumo['razao_bep'] * 100:.0f}% do ponto nominal; rendimento {proc_resumo['eta'] * 100:.0f}%",
             "Status": ("OK" if proc_resumo["faixa_ok"] else "ALERTA: fora da faixa") + _aviso_curva},
            {"Verificação": "NPSH no ponto de operação", "Resultado": f"Margem de {_n(proc_resumo['marg'], 2)} mca (mínimo {_n(proc_resumo['marg_min'], 2)} mca)",
             "Status": ("OK" if proc_resumo["npsh_ok"] else "ALERTA: margem insuficiente") + _aviso_curva},
            {"Verificação": "Pressão máxima (fechada e aríete) x PN da adutora", "Resultado": f"Fechada {_n(proc_resumo['p_shut'], 1)} mca; pico {_n(proc_resumo['p_pico'], 1)} mca; PN {_n(pn_adutora, 0)} mca",
             "Status": ("OK" if proc_resumo["pn_ok"] else "ALERTA: acima da classe de pressão") + _aviso_curva},
        ]
        if proc_resumo.get("tem_crista"):
            verificacoes_extra.append({"Verificação": "Pressão no ponto alto da adutora", "Resultado": f"{_n(proc_resumo['p_crista'], 1)} mca a {_n(proc_resumo['cota_crista'], 1)} m de cota",
                                       "Status": "OK (instalar ventosa)" if proc_resumo["p_crista"] >= 2.0 else "ALERTA: pressão muito baixa ou negativa no ponto alto"})
    st.dataframe(pd.DataFrame(verificacoes_extra), hide_index=True, use_container_width=True)
    
    # ==========================================================================
    # RELATÓRIO TÉCNICO AO CLIENTE (DOCX gerado sem dependências externas)
    # Estrutura inspirada na ABNT NBR 10719 (relatório técnico) e formatação
    # NBR 14724: A4, margens 3/3/2/2 cm, fonte 12, entrelinha 1,5, só texto preto.
    # ==========================================================================
    st.markdown("#### Dados para o Relatório ao Cliente")
    col_r1, col_r2 = st.columns(2)
    nome_cliente = col_r1.text_input("Nome do cliente:", value="", key="laudo_cliente")
    nome_propriedade = col_r2.text_input("Propriedade / local da obra:", value="", key="laudo_propriedade")
    col_r3, col_r4, col_r5 = st.columns(3)
    resp_tecnico = col_r3.text_input("Responsável técnico:", value="", key="laudo_resp")
    crea_numero = col_r4.text_input("CREA / registro:", value="", key="laudo_crea")
    art_numero = col_r5.text_input("Nº da ART:", value="", key="laudo_art")

    def montar_analises():
        """Para cada um dos 10 critérios: referência, análise (por que aprovou / por que reprovou)
        e, se reprovado, medidas de adequação com os números reais do projeto."""
        A = []

        def add(i, ref, analise, recs=None, complemento=None):
            c = criterios[i]
            A.append({
                "num": i + 1, "titulo": c["item"].split(". ", 1)[1], "ref": ref,
                "resultado": c["resultado"], "aprovado": bool(c["aprovado"]),
                "analise": analise, "recs": (recs or []) if not c["aprovado"] else [],
                "complemento": complemento,
            })

        # 1 ---- Infiltração
        q_max = vib_solo * espacamento_emissores * espacamento_linhas
        add(0, "Bernardo et al. (Manual de Irrigação); valores de VIB por classe textural (Embrapa Solos).",
            (f"A taxa de aplicação do sistema é de {_n(taxa_aplicacao_real)} mm/h, enquanto o solo ({tipo_solo}) "
             f"infiltra, no mínimo, {_n(vib_solo, 1)} mm/h (velocidade de infiltração básica, VIB). Como a água chega "
             f"mais devagar do que o solo consegue absorver (folga de {_n(vib_solo - taxa_aplicacao_real)} mm/h), ela "
             f"entra no solo no próprio local de aplicação, sem formar poças nem escorrer pela superfície. Isso evita "
             f"perda de água, erosão e irrigação desigual, e por isso o critério está aprovado.")
            if criterios[0]["aprovado"] else
            (f"A taxa de aplicação do sistema ({_n(taxa_aplicacao_real)} mm/h) é maior do que a velocidade de infiltração "
             f"básica do solo ({tipo_solo}, VIB = {_n(vib_solo, 1)} mm/h), um excesso de {_n(taxa_aplicacao_real - vib_solo)} mm/h. "
             f"Nessa condição a água chega mais rápido do que o solo absorve, formando poças e escorrimento superficial, "
             f"com desperdício de água, risco de erosão e distribuição irregular."),
            [f"Reduzir a vazão de cada emissor para no máximo {_n(q_max, 1)} L/h (valor atual: {_n(vazao_emissor_campo, 1)} L/h), "
             f"mantendo os espaçamentos de {_n(espacamento_emissores)} m x {_n(espacamento_linhas)} m.",
             "Ou aumentar a área atendida por emissor (maior espaçamento entre emissores ou entre linhas), reduzindo a taxa em mm/h.",
             "Dividir a irrigação em pulsos curtos, com intervalos para infiltração, e proteger o solo com cobertura morta (palhada).",
             ])

        # 2 ---- Uniformidade
        add(1, ("Critério de variação de pressão de 20 % da pressão de serviço (Bernardo et al.); fator de Christiansen para múltiplas saídas." if eh_aspersao else "ABNT NBR ISO 9261:2006 (emissores); critério de variação de pressão de 20 % (Keller e Karmeli; Bernardo et al.); fator de Christiansen para múltiplas saídas."),
            (f"Ao longo da linha lateral de {_n(comp_linha_digitado, 0)} m (diâmetro interno de {_n(diam_linha_int_mm, 0)} mm), a "
             f"perda de pressão por atrito, já corrigida pelo fator de Christiansen para tubos com várias saídas, é de "
             f"{_n(hf_lateral_corrigida)} mca. Esse valor é menor do que o limite de {_n(limite_var_pressao)} mca (20 % da pressão "
             f"de serviço do emissor, de {_n(pressao_nominal_emissor, 1)} mca). Assim, o último emissor da linha recebe quase a mesma "
             f"pressão do primeiro e entrega praticamente a mesma vazão, o que garante que as plantas recebam água de forma uniforme.")
            if criterios[1]["aprovado"] else
            (f"A perda de pressão ao longo da linha lateral é de {_n(hf_lateral_corrigida)} mca, acima do limite de "
             f"{_n(limite_var_pressao)} mca (20 % da pressão de serviço de {_n(pressao_nominal_emissor, 1)} mca). Os emissores do "
             f"fim da linha recebem pressão menor e, por isso, vazão menor que a dos emissores do início, resultando em "
             f"irrigação desuniforme."),
            [f"Reduzir o comprimento da linha lateral (atual: {_n(comp_linha_digitado, 0)} m) ou alimentá-la pelo centro, dividindo o percurso da água.",
             f"Aumentar o diâmetro interno da linha lateral (atual: {_n(diam_linha_int_mm, 0)} mm); a perda de pressão cai de forma muito acentuada com o diâmetro.",
             f"Adotar emissores autocompensantes, que mantêm a vazão numa faixa ampla de pressão, ou reduzir o número de emissores por linha.",
             f"Meta: perda de pressão na lateral inferior a {_n(limite_var_pressao)} mca (hoje {_n(hf_lateral_corrigida)} mca).",
             ])

        # 3 ---- Pressão x tubo (térmico)
        pn_req = pressao_inicio_lateral / derating_termico if derating_termico > 0 else pressao_inicio_lateral
        add(2, "Catálogo do fabricante do tubo (classe de pressão PN e redução da resistência do PVC/PE com a temperatura).",
            (f"A pressão no início da lateral é de {_n(pressao_inicio_lateral, 1)} mca. O tubo escolhido tem classe PN de "
             f"{_n(pn_classe_tubo_lateral, 1)} mca a 25 °C; com a temperatura máxima local de {_n(temp_ambiente, 0)} °C o plástico "
             f"perde resistência e a pressão suportada cai para {_n(pn_real_lateral, 1)} mca (fator de {_n(derating_termico * 100, 0)} %). "
             f"Como a pressão de trabalho é menor que esse valor, o tubo trabalha com folga mesmo nos dias mais quentes, "
             f"sem risco de estufamento ou rompimento. Por isso o critério está aprovado.")
            if criterios[2]["aprovado"] else
            (f"A pressão no início da lateral ({_n(pressao_inicio_lateral, 1)} mca) é maior do que a pressão que o tubo suporta na "
             f"temperatura local ({_n(pn_real_lateral, 1)} mca, resultado de PN {_n(pn_classe_tubo_lateral, 1)} mca com redução térmica de "
             f"{_n(derating_termico * 100, 0)} % a {_n(temp_ambiente, 0)} °C). Há risco de estufamento, vazamentos e rompimento da linha."),
            [f"Instalar regulador de pressão na entrada de cada setor para limitar a pressão a, no máximo, {_n(pn_real_lateral, 1)} mca.",
             f"Ou adotar tubo de classe superior, com PN real (já com redução térmica) de pelo menos {_n(pressao_inicio_lateral, 1)} mca, isto é, PN nominal próximo de {_n(pn_req, 1)} mca.",
             "Reduzir a exposição ao sol (enterrar ou cobrir as linhas) para diminuir a temperatura de operação do tubo.",
             ])

        # 4 ---- Flushing
        add(3, ("Não aplicável à aspersão." if eh_aspersao else "Boa prática de projeto de irrigação localizada (Bernardo et al.): linhas laterais devem poder ser lavadas periodicamente."),
            ("O teste de autolimpeza (lavagem das linhas) vale para tubos de gotejamento, cujos orifícios pequenos acumulam sedimentos. Em aspersão, com bocais de maior diâmetro, ele não se aplica "
             "e o critério é considerado atendido sem avaliação.") if eh_aspersao else
            (f"Para limpar as linhas laterais é necessária uma vazão de lavagem (flushing) de {_n(vazao_flushing_m3h, 3)} m³/h, que cria a velocidade "
             f"mínima capaz de arrastar sedimentos e resíduos até o fim da linha. A vazão normal de operação da linha é de {_n(vazao_linha_lateral_m3h, 3)} m³/h. "
             + ("A vazão de lavagem é compatível com a capacidade do sistema, de modo que a limpeza periódica pode ser feita sem sobrecarregar a bomba. "
                if not flushing_excede else
                "A vazão de lavagem é superior à vazão de operação, mas a limpeza pode ser feita lavando uma linha ou um setor por vez. ")
             + "Por isso o critério está aprovado."),
            None,
            ("Como a lavagem exige mais água do que a operação normal, recomenda-se executá-la por partes (uma linha ou setor por vez), "
             "com a extremidade das laterais dotadas de válvula ou final de linha de fácil abertura.") if (flushing_excede and not eh_aspersao) else None)

        # 5 ---- Golpe de aríete
        base_est = pressao_maxima_transiente - sobrepressao_ariete
        vmax_ar = ((pn_adutora - base_est) * 9.81 / celeridade_onda_pvc) if (pn_adutora > base_est and celeridade_onda_pvc > 0) else 0.0
        d_ar = diam_adutora_int_mm * math.sqrt(vel_adutora / vmax_ar) if vmax_ar > 0 else None
        recs5 = [
            "Aumentar o tempo de fechamento das válvulas e adotar partida e parada progressivas da bomba (soft-starter ou inversor de frequência).",
            "Instalar válvula antecipadora de onda, válvula de alívio ou câmara de ar (ventosas de tríplice função) nos pontos altos e logo após a bomba.",
        ]
        if d_ar:
            recs5.append(f"Reduzir a velocidade na adutora para no máximo {_n(vmax_ar)} m/s, o que corresponde a um diâmetro interno de cerca de {_n(d_ar, 0)} mm (atual: {_n(diam_adutora_int_mm, 0)} mm).")
        recs5.append(f"Ou adotar tubo de classe de pressão superior à atual (PN {_n(pn_adutora, 0)} mca), maior que o pico de {_n(pressao_maxima_transiente, 1)} mca.")
        add(4, "Equação de Joukowsky–Allievi: sobrepressão = celeridade x velocidade / g.",
            (f"Se a bomba parar ou uma válvula fechar de repente, a água em movimento gera uma onda de pressão. Com velocidade de {_n(vel_adutora)} m/s "
             f"e celeridade de onda de {_n(celeridade_onda_pvc, 0)} m/s, a sobrepressão calculada é de {_n(sobrepressao_ariete, 1)} mca. Somada à pressão de operação na saída da bomba, "
             f"o pico chega a {_n(pressao_maxima_transiente, 1)} mca, abaixo da classe de pressão da adutora (PN {_n(pn_adutora, 0)} mca). Isso significa que "
             f"o tubo suporta o pior caso de fechamento brusco, e por isso o critério está aprovado.")
            if criterios[4]["aprovado"] else
            (f"A velocidade de {_n(vel_adutora)} m/s e a celeridade de {_n(celeridade_onda_pvc, 0)} m/s produzem sobrepressão de {_n(sobrepressao_ariete, 1)} mca num fechamento "
             f"brusco. O pico de pressão chega a {_n(pressao_maxima_transiente, 1)} mca, acima da classe de pressão da adutora (PN {_n(pn_adutora, 0)} mca), "
             f"com risco de ruptura de tubos e conexões."),
            recs5)

        # 6 ---- Velocidade
        ok6 = criterios[5]["aprovado"]
        if vel_adutora < 0.6:
            d6 = diam_adutora_int_mm * math.sqrt(vel_adutora / 1.0) if vel_adutora > 0 else None
            txt6 = (f"A velocidade na adutora é de {_n(vel_adutora)} m/s, abaixo do mínimo de 0,6 m/s. Em velocidade tão baixa a areia e o lodo "
                    f"tendem a se depositar no fundo do tubo, e o diâmetro está superdimensionado (custo maior sem ganho).")
            rec6 = [f"Reduzir o diâmetro interno da adutora para cerca de {_n(d6, 0)} mm (atual: {_n(diam_adutora_int_mm, 0)} mm), o que leva a velocidade para perto de 1,0 m/s." if d6 else
                    "Reduzir o diâmetro da adutora ou aumentar a vazão de operação.",
                    "Conferir a perda de carga e o golpe de aríete após a alteração do diâmetro (critérios 5 e 9)."]
        else:
            d6 = diam_adutora_int_mm * math.sqrt(vel_adutora / 1.5)
            txt6 = (f"A velocidade na adutora é de {_n(vel_adutora)} m/s, acima do máximo de 2,2 m/s. Velocidades altas aumentam a perda de pressão "
                    f"por atrito, o consumo de energia, o ruído e a intensidade do golpe de aríete.")
            rec6 = [f"Aumentar o diâmetro interno da adutora para cerca de {_n(d6, 0)} mm (atual: {_n(diam_adutora_int_mm, 0)} mm), levando a velocidade para perto de 1,5 m/s.",
                    "Conferir o custo da tubulação e a nova perda de carga na altura manométrica da bomba."]
        add(5, "Critério usual de projeto de adutoras: velocidade entre 0,6 m/s (evita sedimentação) e 2,2 m/s (limita perda de carga e transitórios).",
            (f"A velocidade da água na adutora é de {_n(vel_adutora)} m/s, dentro da faixa recomendada de 0,6 a 2,2 m/s. Abaixo de 0,6 m/s haveria depósito "
             f"de sedimentos e acima de 2,2 m/s a perda de pressão por atrito e a sobrepressão em manobras cresceriam muito. "
             f"Dentro da faixa, o escoamento é estável e econômico, e por isso o critério está aprovado.") if ok6 else txt6,
            rec6)

        # 7 ---- Cavitação
        defic7 = npsh_requerido_catalogo - npsh_disponivel
        add(6, "ABNT NBR 12214 (projeto de sistema de bombeamento): NPSH disponível deve superar o requerido pela bomba, com margem de segurança.",
            (f"O NPSH disponível na instalação é de {_n(npsh_disponivel)} mca (pressão atmosférica local de {_n(patm_local, 2)} mca, menos {_n(altura_succao, 1)} m de altura de "
             f"sucção, menos {_n(hf_perda_succao, 2)} mca de perda na sucção e menos a pressão de vapor da água). O fabricante exige {_n(npsh_requerido_catalogo, 1)} mca. "
             f"A folga de {_n(npsh_disponivel - npsh_requerido_catalogo)} mca impede a formação de bolhas de vapor no rotor, evitando cavitação, ruído, vibração e desgaste da bomba. "
             f"Por isso o critério está aprovado.")
            if criterios[6]["aprovado"] else
            (f"O NPSH disponível ({_n(npsh_disponivel)} mca) não supera o NPSH requerido pela bomba ({_n(npsh_requerido_catalogo, 1)} mca) com a margem necessária "
             f"(déficit de {_n(max(defic7, 0))} mca). Há risco de cavitação, que causa ruído, vibração, queda de vazão e destruição do rotor."),
            [f"Reduzir a altura de sucção (atual: {_n(altura_succao, 1)} m) em pelo menos {_n(max(defic7, 0) + 0.5, 1)} m, ou instalar a bomba afogada (abaixo do nível da água).",
             f"Diminuir a perda na sucção (atual: {_n(hf_perda_succao, 2)} mca): aumentar o diâmetro, encurtar o trecho e reduzir curvas, registros e válvulas.",
             "Escolher bomba com NPSH requerido menor, conforme curva do fabricante.",
             "Evitar água quente na sucção, pois a pressão de vapor aumenta com a temperatura.",
             ])

        # 8 ---- Regra de ouro
        add(7, "ABNT NBR 12214 (limites de velocidade e diâmetros na sucção); boa prática de instalações de bombeamento.",
            (f"O diâmetro da tubulação de sucção ({_n(diam_succao_mm, 0)} mm) é igual ou maior que o da tubulação de recalque ({_n(diam_adutora_int_mm, 0)} mm). "
             f"Com isso a velocidade na sucção é menor que no recalque, a perda de pressão na entrada da bomba fica pequena e o NPSH disponível é preservado. "
             f"Por isso o critério está aprovado.")
            if criterios[7]["aprovado"] else
            (f"O diâmetro da sucção ({_n(diam_succao_mm, 0)} mm) é menor que o do recalque ({_n(diam_adutora_int_mm, 0)} mm). A água entra na bomba em velocidade "
             f"maior do que sai, aumentando a perda de pressão na sucção e o risco de cavitação."),
            [f"Adotar tubulação de sucção com diâmetro de pelo menos {_n(diam_adutora_int_mm, 0)} mm (de preferência uma bitola comercial acima).",
             "Instalar redução excêntrica na entrada da bomba, com a parte plana voltada para cima, evitando bolsas de ar.",
             ])

        # 9 ---- Queda de tensão
        add(8, "ABNT NBR 5410:2004, item 6.2.7: queda de tensão máxima nos circuitos terminais (4 %).",
            (f"O motor da bomba demanda corrente de {_n(corrente_nominal_in, 1)} A no circuito {tensao_circuito}, a {_n(distancia_quadro_cabo, 0)} m do quadro. "
             f"Com cabo de cobre de {_n(bitola_escolhida, 1)} mm² (seção mínima teórica de {_n(secao_teorica_mm2, 2)} mm²), a queda de tensão calculada é de "
             f"{_n(queda_efetiva_perc)} %, inferior ao limite de {_n(queda_max_perm, 1)} %. Dessa forma o motor recebe tensão adequada na partida e em regime, "
             f"sem superaquecimento do cabo ou perda de rendimento, e por isso o critério está aprovado.")
            if criterios[8]["aprovado"] else
            (f"Com cabo de cobre de {_n(bitola_escolhida, 1)} mm², corrente de {_n(corrente_nominal_in, 1)} A e distância de {_n(distancia_quadro_cabo, 0)} m, a queda de tensão "
             f"é de {_n(queda_efetiva_perc)} %, acima do limite de {_n(queda_max_perm, 1)} %. O motor pode ter dificuldade de partida, aquecer e perder rendimento."),
            [f"Aumentar a seção do condutor para, no mínimo, {_n(secao_teorica_mm2, 2)} mm², adotando a próxima bitola comercial acima da atual ({_n(bitola_escolhida, 1)} mm²).",
             f"Aproximar o quadro de comando da bomba (distância atual: {_n(distancia_quadro_cabo, 0)} m).",
             "Avaliar alimentação em tensão mais alta (por exemplo, trifásica 380 V), que reduz a corrente e a queda de tensão.",
             ])

        # 10 ---- Outorga
        add(9, "Lei nº 9.433/1997 (Política Nacional de Recursos Hídricos); regras da ANA e do órgão gestor estadual.",
            (f"A vazão de captação do projeto é de {_n(vazao_setor_operando_m3h)} m³/h, abaixo da referência de 5 m³/h adotada neste estudo para enquadrar o uso como "
             f"insignificante. Nessa condição, em geral, o uso dispensa outorga. Os limites exatos são definidos por cada órgão gestor (ANA e órgão estadual de {uf_selecionado}); "
             f"recomenda-se confirmá-los antes da implantação. Por isso o critério está aprovado.")
            if criterios[9]["aprovado"] else
            (f"A vazão de captação do projeto é de {_n(vazao_setor_operando_m3h)} m³/h, acima da referência de 5 m³/h adotada neste estudo para uso insignificante. "
             f"O uso da água provavelmente depende de outorga de direito de uso de recursos hídricos. Isto não é falha de dimensionamento, mas exigência legal "
             f"que deve ser cumprida antes de operar."),
            [f"Protocolar o pedido de outorga junto ao órgão gestor competente (ANA ou órgão estadual de {uf_selecionado}), anexando este relatório e o projeto.",
             "Alternativamente, avaliar a redução da vazão de captação (por exemplo, operando menos setores ao mesmo tempo) ou o uso de água de chuva armazenada, caso o cenário permita.",
             "Manter o comprovante de outorga junto ao projeto para fiscalização.",
             ])
        return A

    def gerar_docx_laudo(d, analises, compliance_rows, extras_rows):
        W = 9071  # largura útil em DXA (A4, margens 3 cm / 2 cm)
        esc = lambda v: html.escape(str(v), quote=False)

        def runs(text, bold=False, italic=False, size=None):
            out = []
            for i, part in enumerate(str(text).split("**")):
                if part == "":
                    continue
                b = bold or (i % 2 == 1)
                rpr = ("<w:b/>" if b else "") + ("<w:i/>" if italic else "") + (f'<w:sz w:val="{size}"/>' if size else "")
                out.append(f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(part)}</w:t></w:r>')
            return "".join(out)

        def P(text="", style=None, align=None, first=0, left=0, hang=0, bold=False, italic=False,
              size=None, before=None, after=None, line=None, keep=False, pbb=False):
            ppr = ""
            if style: ppr += f'<w:pStyle w:val="{style}"/>'
            if keep: ppr += "<w:keepNext/>"
            if pbb: ppr += "<w:pageBreakBefore/>"
            sp = ""
            if before is not None: sp += f' w:before="{before}"'
            if after is not None: sp += f' w:after="{after}"'
            if line is not None: sp += f' w:line="{line}" w:lineRule="auto"'
            if sp: ppr += f"<w:spacing{sp}/>"
            ind = ""
            if left: ind += f' w:left="{left}"'
            if hang: ind += f' w:hanging="{hang}"'
            if first: ind += f' w:firstLine="{first}"'
            if ind: ppr += f"<w:ind{ind}/>"
            if align: ppr += f'<w:jc w:val="{align}"/>'
            return f"<w:p><w:pPr>{ppr}</w:pPr>{runs(text, bold, italic, size)}</w:p>"

        def H1(t): return P(t, style="Titulo1")
        def H2(t): return P(t, style="Titulo2")
        def corpo(t): return P(t, first=709)
        def simples(t, **k): return P(t, **k)

        def tabela(rows, widths, header=None, bold_first=False):
            """Tabela sem nenhuma borda e sem sombreamento; cabeçalho apenas em negrito."""
            grid = "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
            nil = "".join(f'<w:{s} w:val="nil"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
            x = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{nil}</w:tblBorders>'
                 f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                 f'<w:tblGrid>{grid}</w:tblGrid>')

            def linha(cells, bold=False, hdr=False):
                trpr = "<w:trPr><w:cantSplit/>" + ("<w:tblHeader/>" if hdr else "") + "</w:trPr>"
                tcs = ""
                for j, (c, w) in enumerate(zip(cells, widths)):
                    b = bold or (bold_first and j == 0)
                    tcs += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
                            f'{P(c, align="left", bold=b, size=22, before=40, after=40, line=240)}</w:tc>')
                return f"<w:tr>{trpr}{tcs}</w:tr>"

            if header: x += linha(header, bold=True, hdr=True)
            for r in rows: x += linha(r)
            return x + "</w:tbl>"

        meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
        hoje = datetime.now()
        data_ext = f"{hoje.day} de {meses[hoje.month - 1]} de {hoje.year}"
        nd = "Não informado"
        cliente = d["cliente"] or nd

        corpo_xml = []
        add_ = corpo_xml.append

        # ---------------- CAPA ----------------
        add_(P("RELATÓRIO TÉCNICO DE VALIDAÇÃO", align="center", bold=True, size=32, before=2600, after=120))
        add_(P("PROJETO DE IRRIGAÇÃO", align="center", bold=True, size=28, after=120))
        add_(P("Memorial de cálculo e laudo de executabilidade", align="center", size=24, after=1400))
        add_(P(f"Cliente: {cliente}", align="center", after=60))
        add_(P(f"Local: {d['propriedade'] or nd} – {d['cidade']}/{d['uf']}", align="center", after=60))
        add_(P(f"Cultura: {d['cultura']}", align="center", after=60))
        add_(P(f"Responsável técnico: {d['resp'] or nd}", align="center", after=3200))
        add_(P(f"{d['cidade']} – {d['uf']}", align="center", after=60))
        add_(P(f"{hoje.year}", align="center"))

        # ---------------- 1 OBJETO ----------------
        add_(P("1 OBJETO E FINALIDADE", style="Titulo1", pbb=True))
        add_(corpo(f"Este relatório apresenta ao cliente {cliente} a validação técnica do projeto de irrigação da cultura de {d['cultura']}, "
                   f"em {d['cidade']}/{d['uf']}. O objetivo é informar, em linguagem objetiva, se o sistema dimensionado é tecnicamente executável e se atende às normas "
                   f"da ABNT e aos critérios de engenharia aplicáveis."))
        add_(corpo("Foram verificados dez critérios que cobrem o solo, a tubulação, a bomba, a parte elétrica e a legislação de uso da água. Para cada critério são "
                   "apresentados o resultado calculado, a referência normativa, a justificativa da aprovação e, quando houver não conformidade, as medidas "
                   "recomendadas para adequação do projeto."))

        # ---------------- 2 DADOS ----------------
        add_(H1("2 DADOS DO PROJETO"))
        kv = [["Cliente", cliente], ["Propriedade / local", d["propriedade"] or nd], ["Município / UF", f"{d['cidade']} / {d['uf']}"],
              ["Cultura", d["cultura"]], ["Método de irrigação", d["metodo"]], ["Tipo de solo", d["solo"]], ["Área irrigada", d["area"]],
              ["Vazão de captação", d["vazao"]], ["Altura manométrica total", d["amt"]], ["Potência da bomba", d["potencia"]],
              ["Responsável técnico", d["resp"] or nd], ["Registro profissional (CREA)", d["crea"] or nd], ["ART nº", d["art"] or nd], ["Data", data_ext]]
        add_(tabela(kv, [3200, W - 3200], bold_first=True))

        # ---------------- 3 PARECER ----------------
        n_ok = d["total_aprovados"]
        nao_conf = [a for a in analises if not a["aprovado"]]
        add_(H1("3 PARECER GERAL"))
        add_(P(f"**{d['status_geral']}**", first=709))
        if len(nao_conf) == 0:
            add_(corpo(f"Dos 10 critérios verificados, todos foram atendidos. O projeto é considerado tecnicamente executável nas condições informadas ao sistema."))
        elif n_ok >= 9:
            add_(corpo(f"Dos 10 critérios verificados, {n_ok} foram atendidos. O projeto é considerado tecnicamente executável, mas o item não conforme deve ser tratado "
                       f"conforme as recomendações da seção 4 antes da implantação ou da operação."))
        elif n_ok >= 7:
            add_(corpo(f"Dos 10 critérios verificados, {n_ok} foram atendidos. O projeto é executável com ressalvas: os itens não conformes devem ser corrigidos "
                       f"conforme as recomendações da seção 4 antes da implantação."))
        else:
            add_(corpo(f"Dos 10 critérios verificados, apenas {n_ok} foram atendidos. O projeto não deve ser executado na forma atual; é necessário ajustar o "
                       f"dimensionamento conforme as recomendações da seção 4 e refazer a verificação."))
        add_(P("A Tabela 1 resume o resultado de cada critério.", first=709, keep=True))
        add_(P("Tabela 1 – Resumo da verificação dos critérios", bold=True, size=22, before=120, after=60, line=240, keep=True))
        resumo = [[f"{a['num']}. {a['titulo']}", a["resultado"], "CONFORME" if a["aprovado"] else "NÃO CONFORME"] for a in analises]
        add_(tabela(resumo, [2800, 4171, 2100], header=["Critério", "Resultado calculado", "Situação"]))
        add_(P("Fonte: elaboração própria, com base no motor de cálculo do projeto.", size=20, before=60, after=120, line=240))

        # ---------------- 4 ANÁLISE ----------------
        add_(H1("4 ANÁLISE DETALHADA DOS CRITÉRIOS"))
        add_(corpo("Cada subseção a seguir explica por que o critério foi aprovado ou reprovado e, nos casos de não conformidade, indica como adequar o projeto."))
        for a in analises:
            add_(H2(f"4.{a['num']} {a['titulo']}"))
            add_(P(f"**Referência:** {a['ref']}", keep=True))
            add_(P(f"**Resultado calculado:** {a['resultado']}", keep=True))
            add_(P(f"**Situação:** {'CONFORME' if a['aprovado'] else 'NÃO CONFORME'}", keep=True))
            add_(P(f"**Análise:** {a['analise']}", first=0))
            if a["complemento"]:
                add_(P(f"**Observação:** {a['complemento']}"))
            if a["recs"]:
                add_(P("**Medidas recomendadas para adequação:**", keep=True))
                for k, r in enumerate(a["recs"]):
                    add_(P(f"{'abcdefgh'[k]}) {r}", left=1134, hang=425))

        # ---------------- 5 CONFORMIDADE ----------------
        add_(H1("5 CONFORMIDADE NORMATIVA"))
        add_(corpo("A Tabela 2 relaciona as normas da ABNT e a legislação consideradas e a situação do projeto frente a cada uma."))
        add_(P("Tabela 2 – Conformidade com normas e legislação", bold=True, size=22, before=120, after=60, line=240, keep=True))
        add_(tabela([[c["norma"], c["status"], c["criterio"]] for c in compliance_rows], [3000, 2600, 3471], header=["Norma / legislação", "Situação", "Critério de validação"]))
        add_(P("Fonte: elaboração própria.", size=20, before=60, after=120, line=240))

        # ---------------- 6 COMPLEMENTARES ----------------
        add_(H1("6 VERIFICAÇÕES COMPLEMENTARES"))
        add_(corpo("Além dos dez critérios, o sistema verifica o funcionamento conjunto da captação de água de chuva, da cisterna e da rede pública."))
        add_(P("Tabela 3 – Verificações complementares", bold=True, size=22, before=120, after=60, line=240, keep=True))
        add_(tabela([[e["Verificação"], e["Resultado"], e["Status"]] for e in extras_rows], [2700, 4271, 2100], header=["Verificação", "Resultado", "Situação"]))
        add_(P("Fonte: elaboração própria.", size=20, before=60, after=120, line=240))

        # ---------------- 7 CONCLUSÃO ----------------
        add_(H1("7 CONCLUSÃO"))
        if nao_conf:
            lista = "; ".join(f"{a['num']} ({a['titulo']})" for a in nao_conf)
            add_(corpo(f"O projeto apresenta não conformidade nos seguintes critérios: {lista}. Após aplicar as medidas recomendadas na seção 4, o dimensionamento "
                       f"deve ser recalculado e este relatório, emitido novamente, para confirmar a conformidade."))
        else:
            add_(corpo("Todos os critérios verificados foram atendidos. O dimensionamento é coerente entre si (solo, linhas laterais, adutora, bomba e instalação elétrica) "
                       "e compatível com as normas consideradas, estando o projeto apto à execução nas condições informadas."))
        add_(corpo("Este relatório baseia-se nos dados informados ao sistema e nos modelos de cálculo adotados. Ele não substitui o projeto executivo, a anotação de "
                   "responsabilidade técnica (ART), a vistoria de campo nem as exigências dos órgãos ambientais e das concessionárias. Os dados de campo (vazões, "
                   "pressões, características do solo e da água) devem ser confirmados na implantação."))
        add_(P(f"{d['cidade']}/{d['uf']}, {data_ext}.", before=360, after=720, keep=True))
        add_(P("______________________________________", align="center", keep=True, line=240))
        add_(P(d["resp"] or "Responsável técnico", align="center", bold=True, keep=True, line=240))
        add_(P(f"CREA: {d['crea'] or nd}   |   ART nº {d['art'] or nd}", align="center", line=240))

        # ---------------- REFERÊNCIAS ----------------
        add_(P("REFERÊNCIAS", style="Titulo1", pbb=True, align="center"))
        refs = [
            "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 5410**: instalações elétricas de baixa tensão. Rio de Janeiro: ABNT, 2004.",
            "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 10719**: informação e documentação: relatório técnico e/ou científico: apresentação. Rio de Janeiro: ABNT, 2015.",
            "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 12214**: projeto de sistema de bombeamento de água para abastecimento público. Rio de Janeiro: ABNT, 2020.",
            "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 15527**: água de chuva: aproveitamento de coberturas em áreas urbanas para fins não potáveis: requisitos. Rio de Janeiro: ABNT, 2019.",
            "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR ISO 9261**: equipamentos de irrigação agrícola: emissores e tubos emissores: especificação e métodos de ensaio. Rio de Janeiro: ABNT, 2006.",
            "BERNARDO, S.; SOARES, A. A.; MANTOVANI, E. C. **Manual de irrigação**. 8. ed. Viçosa: UFV, 2006.",
            "BRASIL. **Lei nº 6.496, de 7 de dezembro de 1977**. Institui a Anotação de Responsabilidade Técnica na prestação de serviços de engenharia, de arquitetura e agronomia.",
            "BRASIL. **Lei nº 9.433, de 8 de janeiro de 1997**. Institui a Política Nacional de Recursos Hídricos.",
            "CHRISTIANSEN, J. E. **Irrigation by sprinkling**. Berkeley: University of California, 1942. (Bulletin, 670).",
            "KELLER, J.; KARMELI, D. **Trickle irrigation design**. Glendora: Rain Bird Sprinkler Manufacturing Corporation, 1975.",
        ]
        for r in refs:
            add_(P(r, align="left", line=240, after=240))

        sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:headerReference w:type="first" r:id="rId3"/>'
                '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1701" w:right="1134" w:bottom="1134" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/>'
                '<w:titlePg/></w:sectPr>')
        NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        doc_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(corpo_xml)}{sect}</w:body></w:document>'

        styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Titulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="240"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="000000"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Titulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="000000"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Cabecalho"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
</w:styles>'''

        hdr_page = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}><w:p><w:pPr><w:pStyle w:val="Cabecalho"/></w:pPr>'
                    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>')
        hdr_empty = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}><w:p><w:pPr><w:pStyle w:val="Cabecalho"/></w:pPr></w:p></w:hdr>'

        ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
              '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
              '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
              '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/></Relationships>')

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/document.xml", doc_xml)
            z.writestr("word/styles.xml", styles_xml)
            z.writestr("word/header1.xml", hdr_page)
            z.writestr("word/header2.xml", hdr_empty)
            z.writestr("word/_rels/document.xml.rels", doc_rels)
        return buf.getvalue()

    analises_laudo = montar_analises()
    dados_docx = {
        "cliente": nome_cliente.strip(), "propriedade": nome_propriedade.strip(), "resp": resp_tecnico.strip(),
        "crea": crea_numero.strip(), "art": art_numero.strip(),
        "cidade": cidade_input, "uf": uf_selecionado, "cultura": str(cultura), "metodo": str(metodo_selecionado), "solo": str(tipo_solo),
        "area": f"{_n(area_plantio, 0)} m²",
        "vazao": f"{_n(vazao_setor_operando_m3h)} m³/h", "amt": f"{_n(amt_total, 1)} mca", "potencia": f"{_n(potencia_cv_comercial, 1)} cv",
        "status_geral": status_geral_projeto, "total_aprovados": total_aprovados,
    }

    st.download_button(
        label="Baixar Relatório Técnico ao Cliente (DOCX)",
        data=gerar_docx_laudo(dados_docx, analises_laudo, compliance, verificacoes_extra),
        file_name="Relatorio_Tecnico_Validacao_Irrigacao.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# ------------------------------------------------------------------------------
# 7. ABA: REFERÊNCIAS TÉCNICAS
# ------------------------------------------------------------------------------
with tab_referencias:
    st.subheader("Base Normativa e Bibliografia Técnica")
    st.markdown("""
    O motor de cálculo desta plataforma utiliza metodologias de engenharia consolidadas e normas técnicas brasileiras vigentes.
    
    **1. Agronomia e Solo**
    * **Boletim FAO 56 (Organização das Nações Unidas):** Determinação da Evapotranspiração da Cultura (ETc) e coeficientes de cultivo (Kc).
    * **Embrapa Solos:** Tabela referencial de Velocidade de Infiltração Básica (VIB) e Capacidade de Água Disponível (CAD) por textura de solo.
    
    **2. Hidráulica de Precisão**
    * **ABNT NBR ISO 9261:** Equipamentos de irrigação agrícola — Emissores e tubos emissores (Limitação de variação de pressão em linhas laterais).
    * **Equação de Christiansen (Fator F):** Correção da perda de carga em tubos com múltiplas saídas.
    * **Fórmula de Allievi:** Dimensionamento do transiente hidráulico (Golpe de Aríete) em adutoras.
    * **Cálculo de NPSH:** Verificação do *Net Positive Suction Head* para evitar destruição do rotor por cavitação.
    * **Hazen-Williams:** perda de carga contínua em tubos pressurizados, usada na adutora e no solver de rede.
    * **Algoritmo do Gradiente Global (Todini e Pilati, 1988):** solução simultânea das equações de energia e continuidade em redes malhadas, com bomba por curva e emissores dependentes da pressão (aba Processador Hidráulico).
    
    **3. Água e Eletricidade**
    * **ABNT NBR 15527:** Aproveitamento de água de chuva para fins não potáveis (Determinação do coeficiente Runoff e descarte First-Flush).
    * **ABNT NBR 5410:** Instalações elétricas de baixa tensão (Dimensionamento de condutores de cobre por critério de limite de queda de tensão de 4%).
    """)
    
    st.markdown("""
    **4. Bases de Dados Utilizadas**
    * **IBGE (municípios):** lista dos 5.570 municípios com código IBGE, UF e coordenadas (arquivo `municipios.csv`).
    * **Histórico de chuva municipal:** arquivo `chuva_historica_municipios.csv`; para municípios sem histórico, climatologia mensal de precipitação da **NASA POWER** (consulta online nas coordenadas do município) ou, offline, o município mais próximo com histórico.
    * **Hazen-Williams:** perda de carga contínua em tubulações de PVC/PE, com fator F de Christiansen nas laterais.
    * **Planilha `dados_tubulacao_irrigacao.xlsx` (aba 4):** coeficientes C por material, diâmetros comerciais, comprimentos equivalentes de conexões e parâmetros de velocidade/lista de materiais.
        * **[1] GUEDES, H. A. S. Hidráulica.** Apostila, UFPel, 2018-2: eq. (30) de Hazen-Williams e limite D ≥ 50 mm (p. 21); comprimentos equivalentes (p. 25); velocidades no recalque (p. 65); Tab. 1D (valores de C), 1E (comprimentos equivalentes em metros) e 1F (em diâmetros), p. 141-143.
        * **[2] MACINTYRE, A. J. Bombas e Instalações de Bombeamento.** LTC: fórmula de Williams-Hazen e valores de C (p. 654); comprimentos virtuais (p. 655).
        * Valores sem base nas apostilas (C de PEAD/PEBD e alumínio, velocidades de 0,6 a 2,0 m/s, barra de 6 m) vêm do pedido do projeto e devem ser validados com o catálogo do fabricante.
    """)
    
    st.markdown("**5. Classes de VIB (Embrapa / ESALQ-USP)**")
    if not df_infiltracao_embrapa.empty:
        st.dataframe(
            df_infiltracao_embrapa[['Textura', 'Grupo hidrológico', 'VIB mín. (mm/h)', 'VIB máx. (mm/h)', 'Faixa de referência (mm/h)']],
            hide_index=True, use_container_width=True
        )
    col_vib1, col_vib2 = st.columns(2)
    col_vib1.markdown("**Grupos hidrológicos (conservação do solo; Embrapa Meio-Norte)**")
    col_vib1.dataframe(pd.DataFrame({'Grupo': ['D', 'C', 'B', 'A'], 'VIB (mm/h)': ['Menos de 3', '3 a 40', '40 a 190', 'Acima de 190'],
                                     'Textura associada': ['Argilosa', 'Franco-argilosa', 'Franca / média', 'Arenosa']}), hide_index=True, use_container_width=True)
    col_vib2.markdown("**Classes para irrigação (Bernardo; ESALQ/USP)**")
    col_vib2.dataframe(pd.DataFrame({'Classe': ['Muito alta', 'Alta', 'Média', 'Baixa'], 'VIB (mm/h)': ['Acima de 30', '15 a 30', '5 a 15', 'Menos de 5']}), hide_index=True, use_container_width=True)
    st.caption("As duas escalas servem a fins diferentes: os grupos A a D orientam conservação do solo e escoamento; as classes de irrigação orientam a intensidade máxima de aplicação em aspersão. Não são diretamente comparáveis.")
    
    if fontes_embrapa:
        st.markdown("**6. Fontes Consultadas (Embrapa e outras) e Premissas**")
        for linha_fonte in fontes_embrapa[1:]:
            st.markdown("* " + linha_fonte.replace('$', r'\$'))
    else:
        st.info("Planilha agronômica não encontrada: lista de fontes indisponível.")

# ------------------------------------------------------------------------------
# 8. ABA: BALANÇO ECONÔMICO
# ------------------------------------------------------------------------------
with tab_economia:
    st.subheader("Viabilidade Econômica e Economia de Água")
    
    if df_clima is not None and not df_clima.empty:
        consumo_mensal_litros = consumo_diario_total * 30.0
        resultados_mensais = []
        economia_anual_reais = 0.0
        agua_anual_poupada_litros = 0.0
        
        for _, row in df_clima.iterrows():
            mes_nome = row["Mês"]
            chuva_m = row["Chuva (mm)"]
            captado_m = chuva_m * area_telhado * c_runoff
            usado_chuva = min(captado_m, consumo_mensal_litros)
            usado_rede = max(0.0, consumo_mensal_litros - usado_chuva)
            econ_mes = (usado_chuva / 1000.0) * tarifa_agua
            
            economia_anual_reais += econ_mes
            agua_anual_poupada_litros += usado_chuva
            resultados_mensais.append({
                "Mês": mes_nome, "Chuva (mm)": chuva_m, 
                "Chuva Usada (L)": usado_chuva, "Rede Pública Usada (L)": usado_rede,
                "Economia (R$)": econ_mes
            })
            
        df_eco = pd.DataFrame(resultados_mensais)
        
        ec1, ec2, ec3 = st.columns(3)
        ec1.metric("Água Total Poupada no Ano", f"{agua_anual_poupada_litros:,.0f} Litros")
        ec2.metric("Tarifa Concessionária", f"R$ {tarifa_agua:.2f} / m³")
        ec3.metric("Economia Anual Estimada", f"R$ {economia_anual_reais:,.2f} / ano")
        
        if proc_resumo.get("ok") and proc_resumo.get("custo_energia_mes") is not None:
            st.markdown("#### Custo de Energia do Bombeamento")
            en1, en2, en3 = st.columns(3)
            en1.metric("Energia por mês", f"{_n(proc_resumo['kwh_mes'], 1)} kWh")
            en2.metric("Custo de energia por mês", f"R$ {_n(proc_resumo['custo_energia_mes'], 2)}")
            en3.metric("Custo de energia por ano", f"R$ {_n(proc_resumo['custo_energia_mes'] * 12, 2)}")
            st.caption(f"Estimativa do Processador Hidráulico (aba 5): potência no ponto de operação, horas de rega e tarifa de R$ {_n(proc_resumo['tarifa_kwh'], 2)}/kWh. "
                       "É um custo de operação da irrigação, independente de a água vir da chuva ou da rede, e por isso não é descontado da economia de água acima.")
        df_grafico = pd.melt(df_eco, id_vars=['Mês'], value_vars=['Chuva Usada (L)', 'Rede Pública Usada (L)'], var_name='Origem da Água', value_name='Volume (L)')
        fig_balanco = px.bar(df_grafico, x="Mês", y="Volume (L)", color="Origem da Água", 
                             color_discrete_map={"Chuva Usada (L)": "#27AE60", "Rede Pública Usada (L)": "#E74C3C"},
                             title="Balanço Mensal: Origem da Água para a Irrigação")
        st.plotly_chart(fig_balanco, use_container_width=True)


# ------------------------------------------------------------------------------
# 11. ABA: BACK-TEST (projeto publicado x ferramenta) com exportação em XLSX
# ------------------------------------------------------------------------------
def _gerar_xlsx_backtest(df_cmp, entradas, notas, titulo):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo"
    fino = Side(style="thin", color="999999")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Diferença (%) = (Ferramenta − Referência) / Referência × 100. Situação: OK até 2 %, Atenção até 10 %, Divergente acima de 10 %."
    ws["A2"].font = Font(italic=True, size=9)
    cab = ["Grupo", "Parâmetro", "Unidade", "Referência (artigo)", "Ferramenta", "Diferença (%)", "Situação", "Observação"]
    for j, c in enumerate(cab, 1):
        cel = ws.cell(row=4, column=j, value=c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="1F3A5F")
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = borda
    r = 5
    for _, row in df_cmp.iterrows():
        ref, fer = row["Referência"], row["Ferramenta"]
        ws.cell(row=r, column=1, value=row["Grupo"])
        ws.cell(row=r, column=2, value=row["Parâmetro"])
        ws.cell(row=r, column=3, value=row["Unidade"])
        ws.cell(row=r, column=4, value=None if pd.isna(ref) else float(ref))
        ws.cell(row=r, column=5, value=None if pd.isna(fer) else float(fer))
        if (not pd.isna(ref)) and (not pd.isna(fer)) and float(ref) != 0:
            ws.cell(row=r, column=6, value=f"=(E{r}-D{r})/D{r}*100")
            ws.cell(row=r, column=7, value=f'=IF(ABS(F{r})<=2,"OK",IF(ABS(F{r})<=10,"Atenção","Divergente"))')
        else:
            ws.cell(row=r, column=6, value="n/d")
            ws.cell(row=r, column=7, value="Sem referência")
        ws.cell(row=r, column=8, value=row["Observação"])
        for j in range(1, 9):
            c_ = ws.cell(row=r, column=j)
            c_.border = borda
            c_.alignment = Alignment(vertical="top", wrap_text=(j in (2, 8)))
        for j in (4, 5):
            ws.cell(row=r, column=j).number_format = "#,##0.00"
        ws.cell(row=r, column=6).number_format = "0.0"
        r += 1
    for j, w in enumerate([14, 44, 10, 16, 14, 13, 14, 70], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A5"
    from openpyxl.formatting.rule import CellIsRule
    rng = f"G5:G{r - 1}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", bgColor="C6EFCE")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Atenção"'], fill=PatternFill("solid", bgColor="FFEB9C")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Divergente"'], fill=PatternFill("solid", bgColor="FFC7CE")))
    w2 = wb.create_sheet("Entradas usadas")
    w2.append(["Entrada", "Valor"])
    for c in w2[1]:
        c.font = Font(bold=True)
    for k, v in entradas:
        w2.append([k, v])
    w2.column_dimensions["A"].width = 58
    w2.column_dimensions["B"].width = 26
    w3 = wb.create_sheet("Notas")
    for linha in notas:
        w3.append([linha])
    w3.column_dimensions["A"].width = 150
    for row_ in w3.iter_rows():
        for c in row_:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


with tab_backtest:
    st.subheader("Back-test: projeto publicado × ferramenta")
    st.markdown(
        "Compara cada valor de um projeto já publicado (coluna **Referência**, editável) com o valor calculado pela ferramenta **com os dados que estão nas abas ao lado**. "
        "Vem preenchido com o artigo *Projeto de um sistema de irrigação para atender a demanda hídrica do milho forrageiro* (Rev. FT, IFPA Castanhal-PA, DOI 10.5281/zenodo.7831318). "
        "Para reproduzir o artigo, preencha as abas conforme o quadro de entradas abaixo e, na aba 3, zere o filtro, a sucção e a margem de envelhecimento.")

    bb1, bb2 = st.columns(2)
    bb1.button("Carregar o projeto do artigo (preenche todos os campos)", key="btn_preset", on_click=aplicar_preset_artigo, use_container_width=True, type="primary")
    bb2.button("Voltar aos valores padrão da ferramenta", key="btn_limpar", on_click=limpar_preset_artigo, use_container_width=True)
    st.caption("O botão preenche as barras e abas 2 e 3 com os dados do artigo e a tabela abaixo atualiza na hora. Também dá para abrir o app já carregado acrescentando ?preset=artigo ao endereço.")
    with st.expander("Entradas para reproduzir o artigo (quadro de configuração)", expanded=False):
        st.markdown(
            "- **Barra lateral:** método Aspersão; área 10.000 m²; distância reservatório–roça 10 m; desnível 0 m; altura de sucção 0 m; entrada livre da VIB = 12 mm/h.\n"
            "- **Aba 2:** vazão do emissor 2.720 L/h; espaçamentos 18 × 18 m.\n"
            "- **Aba 3:** 2 setores; geometria manual = 3 aspersores por lateral e 10 laterais; comprimento da lateral 50 m; pressão nominal 25 mca; "
            "altura do aspersor 2 m; F automático; C = 150; diâmetro interno da lateral 40 mm; adutora 100 mm; conexões 0 m; linha principal na área 100 m; "
            "envelhecimento 0 %; filtro 0 mca; perda na sucção 0 mca; **perdas localizadas 4 %**; rendimento da bomba 65 %; **rendimento do motor 92 %**.")

    bt1, bt2, bt3, bt4 = st.columns(4)
    comp_succao_bt = bt1.number_input("Comprimento da sucção (m):", min_value=0.0, value=5.0, step=1.0, key="bt_lsuc")
    preco_100 = bt2.number_input("Preço do tubo 100 mm (R$/vara 6 m):", min_value=0.0, value=164.91, step=1.0, key="bt_p100")
    preco_lat = bt3.number_input("Preço do tubo da lateral (R$/vara 6 m):", min_value=0.0, value=63.15, step=1.0, key="bt_plat")
    preco_asp = bt4.number_input("Preço do aspersor (R$/un):", min_value=0.0, value=30.0, step=1.0, key="bt_pasp")
    preco_bomba = st.number_input("Preço da bomba (R$):", min_value=0.0, value=5433.91, step=50.0, key="bt_pbomba")

    # ---- valores da ferramenta ----
    _q_lat_m3h = vazao_linha_lateral_m3h
    _hf_adm = 0.20 * pressao_nominal_emissor

    def _hf_lat_cont(d_mm):
        return 10.67 * comp_linha_digitado * ((q_linha_m3s / c_rugosidade_pvc) ** 1.852) / ((d_mm / 1000.0) ** 4.87) * fator_f

    _lo, _hi = 5.0, 300.0
    for _ in range(80):
        _mid = 0.5 * (_lo + _hi)
        if _hf_lat_cont(_mid) > _hf_adm:
            _lo = _mid
        else:
            _hi = _mid
    d_lat_req = 0.5 * (_lo + _hi) if _hf_lat_cont(300.0) <= _hf_adm else float("nan")
    _comerc_lat = [16.0, 20.0, 25.0, 32.0, 40.0, 50.0, 60.0, 75.0, 85.0, 100.0]
    d_lat_com = next((d_ for d_ in _comerc_lat if d_ >= d_lat_req), float("nan")) if not math.isnan(d_lat_req) else float("nan")
    pin_artigo_metodo = pressao_nominal_emissor + altura_riser + 0.75 * _hf_adm + 0.5 * z_lateral
    hf_princ_somente_area = hazen_williams_hf(vazao_setor_operando_m3h, diam_adutora_int_mm, c_rugosidade_pvc, comp_principal_area, k=10.646) if comp_principal_area > 0 else float("nan")
    n_total_asp = num_total_linhas_campo * num_emissores_por_linha
    tubos_lat = math.ceil(num_total_linhas_campo * comp_linha_digitado / 6.0 - 1e-9)
    tubos_princ = math.ceil(comp_principal_area / 6.0 - 1e-9) if comp_principal_area > 0 else 0
    tubos_100_total = tubos_princ + math.ceil(distancia_adutora / 6.0 - 1e-9) + math.ceil(comp_succao_bt / 6.0 - 1e-9)
    orcamento_tool = tubos_100_total * preco_100 + tubos_lat * preco_lat + n_total_asp * preco_asp + preco_bomba
    _nan = float("nan")

    linhas_bt = [
        # grupo, parâmetro, unidade, ref (artigo), ferramenta, observação
        ("Entrada", "Área irrigada", "m²", 10000.0, float(area_plantio), "Dado de entrada."),
        ("Entrada", "Vazão de cada aspersor", "m³/h", 2.72, vazao_emissor_campo / 1000.0, "O artigo usa 2,72 nos cálculos e 2,73 no texto do aspersor."),
        ("Entrada", "VIB do solo", "mm/h", 12.0, float(vib_solo), "Entrada livre da VIB."),
        ("Entrada", "Taxa de aplicação (Ia = q/(Ea·El))", "mm/h", 8.9, float(taxa_aplicacao_real), "O artigo publica 8,9, mas 2,72/(18×18) = 8,40 mm/h. A ferramenta calcula pela fórmula; em ambos os casos Ia < VIB."),
        ("Geometria", "Aspersores por lateral", "un", 3.0, float(num_emissores_por_linha), "50/18 = 2,77, adotado 3."),
        ("Geometria", "Linhas laterais no campo", "un", 10.0, float(num_total_linhas_campo), "5 de cada lado da principal."),
        ("Geometria", "Total de aspersores", "un", 30.0, float(n_total_asp), ""),
        ("Vazão", "Vazão do sistema (setor operando)", "m³/h", 40.8, float(vazao_setor_operando_m3h), "15 aspersores simultâneos (um lado da principal por vez)."),
        ("Lateral", "Fator F de Christiansen (3 saídas)", "-", 0.535, float(fator_f), "Pequena diferença de tabela (0,534 pela fórmula exata)."),
        ("Lateral", "hf admissível (20 % de Ps)", "mca", 5.0, float(_hf_adm), "Critério de Bernardo et al."),
        ("Lateral", "Diâmetro calculado da lateral", "mm", 33.0, float(d_lat_req), "A ferramenta resolve a equação continuamente; o artigo arredonda o resultado."),
        ("Lateral", "Diâmetro comercial da lateral", "mm", 40.0, float(d_lat_com), "Menor diâmetro comercial ≥ calculado (lista da ferramenta)."),
        ("Lateral", "hf real da lateral no diâmetro informado", "mca", _nan, float(hf_lateral_corrigida), "O artigo não recalcula o hf com o diâmetro comercial; usa o admissível (5 mca)."),
        ("Pressão", "Pressão na entrada da lateral — método do artigo (hf admissível)", "mca", 30.75, float(pin_artigo_metodo), "Pin = Ps + Aa + 0,75·hf_adm + 0,5·ΔZ. Mesmo critério do artigo."),
        ("Pressão", "Pressão na entrada da lateral — método da ferramenta (hf real)", "mca", 30.75, float(pressao_inicio_lateral), "Usa o hf real do diâmetro informado. Fica abaixo do valor do artigo porque o tubo comercial é maior que o mínimo."),
        ("Principal", "Diâmetro por velocidade de 1,5 m/s", "mm", 97.0, float(_d_calc_mm), "D = √(4Q/πV)."),
        ("Principal", "Diâmetro comercial adotado", "mm", 100.0, float(diam_adutora_int_mm), ""),
        ("Principal", "hf da linha principal (100 m, vazão integral)", "mca", 1.8, float(hf_princ_somente_area), "Artigo: 10,646·(Q/C)^1,852/D^4,87·100."),
        ("Bomba", "Subtotal (Pin + hf principal + desníveis + hf recalque/sucção)", "mca", 32.55, float(subtotal_amt), "Na ferramenta o subtotal inclui também o trecho de recalque (distância informada), que o artigo despreza."),
        ("Bomba", "Perdas localizadas", "mca", 1.30, float(hf_localizada_pct), "4 % do subtotal."),
        ("Bomba", "Altura manométrica total (Hman)", "mca", 33.85, float(amt_total), ""),
        ("Bomba", "Hman recalculado pelo método do artigo (Pin do artigo + hf principal, + % localizadas)", "mca", 33.85, float((pin_artigo_metodo + (hf_princ_somente_area if comp_principal_area > 0 else 0.0)) * (1.0 + perda_loc_pct / 100.0)), "Mesma sequência do artigo (sem recalque/sucção): mostra que a ferramenta reproduz a conta dele."),
        ("Bomba", "Potência instalada pelo método do artigo", "cv", 8.5, float(vazao_setor_operando_m3h * ((pin_artigo_metodo + (hf_princ_somente_area if comp_principal_area > 0 else 0.0)) * (1.0 + perda_loc_pct / 100.0)) / (270.0 * (rendimento_bomba / 100.0) * (eta_motor_pct / 100.0))), "Q·Hman/(270·ηb·ηm), equivalente a Q(L/s)·Hman/(75·ηb·ηm)."),
        ("Bomba", "Potência instalada (Q·Hman/(75·ηm·ηb))", "cv", 8.5, float(potencia_cv_instalada), "Artigo: ηm = 92 %, ηb = 65 %. 1 cv = 75 kgf·m/s."),
        ("Materiais", "Tubos de 40 mm (varas de 6 m)", "un", 84.0, float(tubos_lat), "10 laterais × 50 m ÷ 6 m."),
        ("Materiais", "Tubos de 100 mm — principal", "un", 17.0, float(tubos_princ), "100 m ÷ 6 m."),
        ("Materiais", "Tubos de 100 mm — total (principal + recalque + sucção)", "un", 20.0, float(tubos_100_total), ""),
        ("Materiais", "Aspersores", "un", 30.0, float(n_total_asp), ""),
        ("Orçamento", "Orçamento inicial", "R$", 14936.71, float(orcamento_tool), "Quantidades da ferramenta × preços digitados acima (preços de 2023 do artigo)."),
    ]
    df_ref0 = pd.DataFrame({"Parâmetro": [l_[1] for l_ in linhas_bt], "Unidade": [l_[2] for l_ in linhas_bt], "Referência": [l_[3] for l_ in linhas_bt]})
    st.markdown("**Coluna Referência (editável)** — troque pelos valores de outro projeto, se quiser reaproveitar o comparativo.")
    df_ref_ed = st.data_editor(df_ref0, hide_index=True, use_container_width=True, disabled=["Parâmetro", "Unidade"], key="bt_editor",
                               column_config={"Referência": st.column_config.NumberColumn("Referência (artigo)", format="%.3f")})
    df_cmp = pd.DataFrame({
        "Grupo": [l_[0] for l_ in linhas_bt],
        "Parâmetro": [l_[1] for l_ in linhas_bt],
        "Unidade": [l_[2] for l_ in linhas_bt],
        "Referência": pd.to_numeric(df_ref_ed["Referência"], errors="coerce").values,
        "Ferramenta": [l_[4] for l_ in linhas_bt],
        "Observação": [l_[5] for l_ in linhas_bt],
    })
    df_cmp["Diferença (%)"] = np.where(df_cmp["Referência"].notna() & (df_cmp["Referência"] != 0) & df_cmp["Ferramenta"].notna(),
                                       (df_cmp["Ferramenta"] - df_cmp["Referência"]) / df_cmp["Referência"].replace(0, np.nan) * 100.0, np.nan)
    df_cmp["Situação"] = df_cmp["Diferença (%)"].apply(lambda v: "Sem referência" if pd.isna(v) else ("OK" if abs(v) <= 2 else ("Atenção" if abs(v) <= 10 else "Divergente")))

    n_ok = int((df_cmp["Situação"] == "OK").sum())
    n_at = int((df_cmp["Situação"] == "Atenção").sum())
    n_dv = int((df_cmp["Situação"] == "Divergente").sum())
    m1, m2, m3 = st.columns(3)
    m1.metric("Itens dentro de 2 %", n_ok)
    m2.metric("Itens entre 2 e 10 %", n_at)
    m3.metric("Itens acima de 10 %", n_dv)

    st.markdown("**Resultado do comparativo**")
    st.dataframe(df_cmp[["Grupo", "Parâmetro", "Unidade", "Referência", "Ferramenta", "Diferença (%)", "Situação", "Observação"]], hide_index=True, use_container_width=True,
                 column_config={"Referência": st.column_config.NumberColumn(format="%.3f"), "Ferramenta": st.column_config.NumberColumn(format="%.3f"),
                                "Diferença (%)": st.column_config.NumberColumn(format="%.1f")})

    notas_bt = [
        "COMO LER ESTE COMPARATIVO",
        "A coluna Referência traz os valores publicados no artigo (Rev. FT, IFPA Castanhal-PA, DOI 10.5281/zenodo.7831318). A coluna Ferramenta traz os valores calculados com as entradas da planilha 'Entradas usadas'.",
        "",
        "DIFERENÇAS DE MÉTODO (não são erros da ferramenta)",
        "1) Pressão na entrada da lateral: o artigo usa o hf ADMISSÍVEL (5 mca); a ferramenta usa o hf REAL do diâmetro informado. As duas versões aparecem no quadro.",
        "2) Recalque e sucção: o artigo despreza as perdas e os desníveis; a ferramenta calcula o hf do trecho de recalque pelo comprimento informado.",
        "3) Fator F: 0,535 (tabela) contra o valor da fórmula de Christiansen.",
        "4) Coeficiente de Hazen-Williams: o artigo usa 10,646 e a lateral da ferramenta usa 10,67 (diferença de 0,2 %).",
        "",
        "INCONSISTÊNCIAS DO PRÓPRIO ARTIGO",
        "a) Ia publicado = 8,9 mm/h, mas q/(Ea·El) = 2,72/(18×18) = 8,40 mm/h (8,43 com 2,73 m³/h).",
        "b) Vazão do aspersor aparece como 2,72 e como 2,73 m³/h.",
        "c) Vazão do sistema aparece como 0,0113 m³/s em um trecho e 0,0133 m³/s em outro.",
        "d) O texto diz 17 tubos de 100 mm; o orçamento usa 20 (17 + 2 de recalque + 1 de sucção).",
        "e) O diâmetro da lateral é calculado com 33 mm e adotado 40 mm, mas o hf e a pressão não são recalculados com 40 mm.",
        "f) A pressão na entrada usa o hf admissível (5 mca) e não o hf real.",
    ]
    entradas_bt = [
        ("Área (m²)", float(area_plantio)), ("Método", str(metodo_selecionado)), ("Vazão do emissor (L/h)", float(vazao_emissor_campo)),
        ("Espaçamento entre emissores (m)", float(espacamento_emissores)), ("Espaçamento entre linhas (m)", float(espacamento_linhas)),
        ("Nº de setores", int(num_setores)), ("Emissores por lateral", int(num_emissores_por_linha)), ("Linhas laterais no campo", int(num_total_linhas_campo)),
        ("Comprimento da lateral (m)", float(comp_linha_digitado)), ("Pressão nominal do emissor (mca)", float(pressao_nominal_emissor)),
        ("Altura do aspersor (m)", float(altura_riser)), ("Coeficiente C", float(c_rugosidade_pvc)), ("F aplicado", float(fator_f)),
        ("Diâmetro interno da lateral (mm)", float(diam_linha_int_mm)), ("Diâmetro interno da adutora (mm)", float(diam_adutora_int_mm)),
        ("Distância reservatório–roça (m)", float(distancia_adutora)), ("Linha principal na área (m)", float(comp_principal_area)),
        ("Conexões da adutora (m eq.)", float(comp_equiv_adutora)), ("Margem de envelhecimento (%)", float(margem_envelhecimento)),
        ("Desnível do recalque (m)", float(desnivel_recalque)), ("Altura de sucção (m)", float(altura_succao)), ("Perda na sucção (mca)", float(hf_perda_succao)),
        ("Perda no filtro (mca)", float(hf_filtro_areia)), ("Perdas localizadas (% do subtotal)", float(perda_loc_pct)),
        ("Rendimento da bomba (%)", float(rendimento_bomba)), ("Rendimento do motor (%)", float(eta_motor_pct)), ("Comprimento da sucção (m)", float(comp_succao_bt)),
        ("Preço tubo 100 mm (R$)", float(preco_100)), ("Preço tubo da lateral (R$)", float(preco_lat)), ("Preço do aspersor (R$)", float(preco_asp)), ("Preço da bomba (R$)", float(preco_bomba)),
    ]
    xlsx_bytes = _gerar_xlsx_backtest(df_cmp, entradas_bt, notas_bt, "Back-test: artigo publicado × ferramenta de irrigação")
    st.download_button("Baixar comparativo (XLSX)", data=xlsx_bytes, file_name="Backtest_Comparativo_Artigo_x_Ferramenta.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with st.expander("Notas e inconsistências do artigo", expanded=False):
        for l_ in notas_bt:
            st.markdown(l_ if l_ else " ")