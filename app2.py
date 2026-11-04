import calendar

import base64

from pathlib import Path

from html import escape

from PIL import Image, ImageDraw, ImageFont
from io import BytesIO



import pandas as pd

import plotly.express as px

import streamlit as st



st.set_page_config(page_title="Dashboard de Ativações", page_icon="📊", layout="wide")

st.warning("🧪 AMBIENTE DE DEMONSTRAÇÃO — Todos os registros e valores apresentados são fictícios.")



# ======================== VISUAL ========================

st.markdown("""

<style>

.stApp {

    background: #0B1120;

    color: #E2E8F0;

}



[data-testid="stHeader"] {

    background: #0B1120;

}



[data-testid="stSidebar"] {

    background: #111C30;

}



h1, h2, h3 {

    color: #F8FAFC !important;

    font-weight: 700 !important;

}



[data-testid="stCaptionContainer"] p {

    color: #94A3B8 !important;

}



[data-testid="stPlotlyChart"] {

    background: #151F32;

    border: 1px solid #29364D;

    border-radius: 22px;

    padding: 12px;

    overflow: hidden;

}



[data-testid="stDataFrame"] {

    border: 1px solid #29364D;

    border-radius: 16px;

    overflow: hidden;

}



[data-testid="stVerticalBlock"] {

    gap: 1rem;

}

</style>

""", unsafe_allow_html=True)



# Caminhos portáveis para execução local e Streamlit Community Cloud.
PASTA_PROJETO = Path(__file__).resolve().parent
PASTA_DADOS = PASTA_PROJETO / "dados"

ARQUIVO_OPERADORAS = PASTA_DADOS / "operadoras.xlsx"

ARQUIVO_HISTORICO = PASTA_DADOS / "historico_operadoras.xlsx"

PASTA_STARLINK = PASTA_DADOS / "starlink"

ARQUIVO_METAS = PASTA_DADOS / "Meta_26.xlsx"

ARQUIVOS_STARLINK = sorted(PASTA_STARLINK.glob("*.csv"))





for arquivo in (ARQUIVO_OPERADORAS, ARQUIVO_HISTORICO):

    if not arquivo.exists():

        st.error(f"Arquivo não encontrado: {arquivo}\nConfira se a pasta dados está ao lado do app.py.")

        st.stop()



if not ARQUIVOS_STARLINK:

    st.error("Nenhum CSV encontrado na pasta dados/starlink.")

    st.stop()





@st.cache_data

def carregar_excel(caminho):

    return pd.read_excel(caminho)



@st.cache_data

def carregar_csv(caminho):

    return pd.read_csv(caminho)

@st.cache_data
def carregar_metas(caminho):
    df = pd.read_excel(caminho, sheet_name="Metas", header=None)

    meses = {
        "jan": 1, "fev": 2, "mar": 3,
        "abr": 4, "mai": 5, "jun": 6,
        "jul": 7, "ago": 8, "set": 9,
        "out": 10, "nov": 11, "dez": 12
    }

    metas = {}

    for i in range(len(df) - 1):
        for coluna in range(df.shape[1]):
            nome = str(df.iat[i, coluna]).strip().lower()

            if nome in meses:
                valor = pd.to_numeric(
                    df.iat[i + 1, coluna],
                    errors="coerce"
                )

                if pd.notna(valor):
                    metas[meses[nome]] = float(valor)

    return metas



@st.cache_data

def consolidar_starlink(arquivos):



    bases = []



    for arquivo in sorted(arquivos):



        df = pd.read_csv(

            arquivo,

            dtype={"service_line_number": "string"}

        )



        df.columns = df.columns.str.strip()



        if "service_line_number" not in df.columns:

            raise ValueError(

                f"Coluna service_line_number ausente: {arquivo}"

            )



        df["_arquivo"] = Path(arquivo).name

        bases.append(df)



    if not bases:

        return pd.DataFrame()



    starlink = pd.concat(bases, ignore_index=True)



    # Padronizar campos vazios

    starlink = starlink.replace(

        r"^\s*$", pd.NA, regex=True

    )



    starlink["service_line_number"] = (

        starlink["service_line_number"]

        .astype("string")

        .str.strip()

    )



    # Separar registros sem identificador

    sem_id = starlink[

        starlink["service_line_number"].isna()

    ].copy()



    com_id = starlink[

        starlink["service_line_number"].notna()

    ].copy()



    # Contabilizar duplicidades

    duplicados = com_id.duplicated(

        subset="service_line_number"

    ).sum()



    # Priorizar o último valor preenchido de cada coluna

    def ultimo_preenchido(serie):

        valores = serie.dropna()

        return valores.iloc[-1] if not valores.empty else pd.NA



    consolidado = (

        com_id.groupby(

            "service_line_number",

            as_index=False,

            sort=False

        )

        .agg(ultimo_preenchido)

    )



    consolidado = consolidado.drop(

        columns=["_arquivo"], errors="ignore"

    )



    return consolidado



def formatar_reais(valor):

    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")



def normalizar_texto(serie):

    return (serie.fillna("").astype("string").str.strip().str.casefold().str.replace(r"\s+", " ", regex=True))



def aplicar_estilo_moderno(fig, barras=False, moeda=False):



    fig.update_layout(

        template="plotly_dark",

        paper_bgcolor="#151F32",

        plot_bgcolor="#151F32",

        font=dict(

            family="Arial, sans-serif",

            size=12,

            color="#E2E8F0"

        ),

        title=dict(

            x=0.03,

            xanchor="left",

            font=dict(size=15, color="#FFFFFF")

        ),

        legend=dict(

            orientation="h",

            yanchor="top",

            y=-0.25,

            xanchor="left",

            x=0,

            font=dict(size=11, color="#CBD5E1"),

            bgcolor="rgba(0,0,0,0)"

        ),

        margin=dict(l=35, r=20, t=70, b=105),

        xaxis=dict(

            showgrid=False,

            tickfont=dict(size=10, color="#CBD5E1"),

            automargin=True

        ),

        yaxis=dict(

            showgrid=True,

            gridcolor="#2B3B52",

            tickfont=dict(size=10, color="#CBD5E1"),

            automargin=True

        ),

        hoverlabel=dict(

            bgcolor="#263B59",

            bordercolor="#425C80",

            font=dict(color="#FFFFFF", size=12)

        ),

        bargap=0.30,

        bargroupgap=0.12

    )



    if moeda:

        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")



    if barras:

        fig.update_traces(

            marker_line_width=0,

            textposition="auto",

            textfont=dict(size=11, color="#FFFFFF"),

            cliponaxis=False

        )



        try:

            fig.update_traces(marker_cornerradius=8)

        except (TypeError, ValueError):

            pass



    return fig











#melhorando a visualização do MoM e YoY dentro do card

def indicador_colorido(nome, valor):



    if valor is None or pd.isna(valor):

        cor = "#94A3B8"

        texto = "N/D"

    elif valor > 0:

        cor = "#10B981"

        texto = f"+{valor:.2f}%"

    elif valor < 0:

        cor = "#F87171"

        texto = f"{valor:.2f}%"

    else:

        cor = "#FBBF24"

        texto = "0,00%"



    texto = texto.replace(".", ",")



    return (

    f'<span style="color:#94A3B8;font-weight:600;">'

    f'{nome}</span> '

    f'<span style="color:{cor};font-weight:700;">'

    f'{texto}</span>'

)



def exibir_card(titulo, valor, detalhe=None):



    titulo = escape(str(titulo))

    valor = escape(str(valor))

    detalhe = str(detalhe) if detalhe else ""



    html = f"""

    <div style="

        background: linear-gradient(135deg, #19263D, #111B2D);

        border: 1px solid #30415D;

        border-radius: 18px;

        padding: 14px 18px;

        height: 112px;

        box-sizing: border-box;

        font-family: Arial, sans-serif;

        display: flex;

        flex-direction: column;

        justify-content: center;

        overflow: hidden;

    ">

        <div style="

            color: #A8BAD3;

            font-size: 12px;

            font-weight: 600;

            margin-bottom: 6px;

        ">{titulo}</div>



        <div style="

            color: #FFFFFF;

            font-size: 24px;

            font-weight: 750;

            line-height: 1.2;

            overflow-wrap: anywhere;

        ">{valor}</div>



        <div style="

            color: #94A3B8;

            font-size: 11px;

            margin-top: 6px;

        ">{detalhe}</div>

    </div>

    """



    st.html(html)

def gerar_relatorio_whatsapp(
    data_inicio,
    data_fim,
    meta_mensal,
    comissao_realizada,
    receita_star,
    receita_total,
    projecao_comissao,
    projecao_star,
    projecao_total,
    ativ_varejo,
    ativ_starlink
):
    largura, altura = 1080, 1350

    img = Image.new("RGB", (largura, altura), "#0B1120")
    draw = ImageDraw.Draw(img)

    def fonte(tamanho, negrito=False):
        arquivo = (
            "C:/Windows/Fonts/arialbd.ttf"
            if negrito else
            "C:/Windows/Fonts/arial.ttf"
        )
        try:
            return ImageFont.truetype(arquivo, tamanho)
        except OSError:
            return ImageFont.load_default()

    def texto(x, y, conteudo, tamanho=25,
              cor="#FFFFFF", negrito=False):
        draw.text(
            (x, y),
            str(conteudo),
            font=fonte(tamanho, negrito),
            fill=cor
        )

    def moeda(valor):
        return (
            f"R$ {valor:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    def percentual(valor):
        return f"{valor:.2f}%".replace(".", ",")

    def card(x, y, titulo, valor, largura_card=490):
        draw.rounded_rectangle(
            (x, y, x + largura_card, y + 135),
            radius=18,
            fill="#19263D",
            outline="#30415D",
            width=2
        )
        texto(x + 22, y + 22, titulo, 22, "#A8BAD3")
        texto(x + 22, y + 65, valor, 33, "#FFFFFF", True)

    meta_valida = meta_mensal is not None and meta_mensal > 0

    atingimento = (
        receita_total / meta_mensal * 100
        if meta_valida else None
    )

    atingimento_projetado = (
        projecao_total / meta_mensal * 100
        if meta_valida else None
    )

    # Cabeçalho
    texto(45, 40, "ALLIED", 52, "#FFFFFF", True)
    texto(
        45, 112,
        "INTELIGÊNCIA COMERCIAL | VAREJO DIGITAL",
        22, "#94A3B8"
    )
    texto(
        45, 160,
        f"Resultado acumulado: {data_inicio:%d/%m/%Y} "
        f"a {data_fim:%d/%m/%Y}",
        25, "#FFFFFF", True
    )

    draw.line((45, 215, 1035, 215), fill="#30415D", width=3)

    # Indicadores de metas
    texto(45, 250, "ACOMPANHAMENTO DE META", 29, "#FFFFFF", True)

    card(
        45, 310, "Meta mensal",
        moeda(meta_mensal) if meta_valida else "Não informada"
    )
    card(545, 310, "Receita realizada", moeda(receita_total))

    card(
        45, 465, "Atingimento realizado",
        percentual(atingimento) if meta_valida else "N/D"
    )
    card(
        545, 465, "Atingimento projetado",
        percentual(atingimento_projetado)
        if meta_valida else "N/D"
    )

    # Barra de atingimento
    texto(45, 640, "PROGRESSO DA META", 25, "#FFFFFF", True)

    draw.rounded_rectangle(
        (45, 690, 1035, 720),
        radius=12,
        fill="#30415D"
    )

    if meta_valida:
        progresso = min(max(atingimento / 100, 0), 1)
        if progresso > 0:
            draw.rounded_rectangle(
                (45, 690, 45 + int(990 * progresso), 720),
                radius=12,
                fill="#10B981"
            )

    # Resultado por operação
    texto(45, 775, "RESULTADO POR OPERAÇÃO", 29, "#FFFFFF", True)

    card(
        45, 835,
        "Receita Operadoras",
        moeda(comissao_realizada)
    )
    card(
        545, 835,
        "Receita Starlink",
        moeda(receita_star)
    )

    # Ativações
    card(
        45, 990,
        "Ativações Operadoras",
        f"{ativ_varejo:,}".replace(",", ".")
    )
    card(
        545, 990,
        "Ativações Starlink",
        f"{ativ_starlink:,}".replace(",", ".")
    )

    # Projeções consolidadas
    texto(
        45, 1170,
        "PROJEÇÃO DE FECHAMENTO",
        26, "#FFFFFF", True
    )

    texto(
        45, 1220,
        moeda(projecao_total),
        43, "#10B981", True
    )

    texto(
        45, 1295,
        "Fonte: Dashboard Allied | Operadoras + Starlink",
        19, "#94A3B8"
    )

    buffer = BytesIO()
    img.save(buffer, format="PNG", optimize=True)
    buffer.seek(0)

    return buffer.getvalue()

## cabelhado dashboard



def exibir_cabecalho(data_atualizacao=None):



    caminho_logo = Path(__file__).resolve().parent / "assets" / "logo_allied.png"



    logo_base64 = base64.b64encode(caminho_logo.read_bytes()).decode("utf-8") if caminho_logo.exists() else ""



    if data_atualizacao is not None and pd.notna(data_atualizacao):

        data_texto = pd.to_datetime(
            data_atualizacao
        ).strftime("%d/%m/%Y")

    else:

        data_texto = "Não disponível"



    st.html(f"""

    <div style="

        background: linear-gradient(110deg, #111B2D, #192B46);

        border: 1px solid #30415D;

        border-radius: 16px;

        padding: 18px 28px;

        margin-bottom: 22px;

        min-height: 90px;

        display: flex;

        align-items: center;

        justify-content: space-between;

        gap: 20px;

        font-family: Arial, sans-serif;

    ">

        <div style="

            display: flex;

            align-items: center;

            gap: 14px;

        ">

            <img

                src="data:image/png;base64,{logo_base64}"

                style="

                    width: 48px;

                    height: 48px;

                    object-fit: contain;

                "

            >



            <div>

                <div style="

                    color: #FFFFFF;

                    font-size: 25px;

                    font-weight: 800;

                    letter-spacing: 2px;

                ">ALLIED</div>



                <div style="

                    color: #94A3B8;

                    font-size: 11px;

                    margin-top: 5px;

                ">

                    INTELIGÊNCIA COMERCIAL | VAREJO DIGITAL

                </div>

            </div>

        </div>



        <div style="text-align:right;">

            <div style="

                color: #94A3B8;

                font-size: 11px;

            ">ATUALIZADO ATÉ</div>



            <div style="

                color: #FFFFFF;

                font-size: 13px;

                font-weight: 700;

                margin-top: 5px;

            ">{data_texto}</div>

        </div>

    </div>

    """)


def adicionar_tabela_ao_png(imagem_original, tabela):
    from PIL import Image, ImageDraw, ImageFont
    from io import BytesIO

    def fonte(tamanho, negrito=False):
        arquivo = (
            "C:/Windows/Fonts/arialbd.ttf"
            if negrito else
            "C:/Windows/Fonts/arial.ttf"
        )
        try:
            return ImageFont.truetype(arquivo, tamanho)
        except OSError:
            return ImageFont.load_default()

    def moeda(valor):
        return (
            f"R$ {valor:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    def inteiro(valor):
        return f"{int(valor):,}".replace(",", ".")

    # Abrir imagem executiva existente
    original = Image.open(
        BytesIO(imagem_original)
    ).convert("RGB")

    largura = original.width

    # Tabela em duas linhas por varejo para facilitar
    # a leitura dos valores no celular
    altura_titulo = 110
    altura_cabecalho = 70
    altura_linha = 106
    margem_final = 55

    altura_tabela = (
        altura_titulo
        + altura_cabecalho
        + len(tabela) * altura_linha
        + margem_final
    )

    imagem = Image.new(
        "RGB",
        (largura, original.height + altura_tabela),
        "#0B1120"
    )

    imagem.paste(original, (0, 0))
    draw = ImageDraw.Draw(imagem)

    inicio = original.height

    draw.text(
        (45, inicio + 30),
        "RESULTADO CONSOLIDADO POR VAREJO",
        font=fonte(28, True),
        fill="#FFFFFF"
    )

    y = inicio + altura_titulo

    # Cabeçalho
    draw.rounded_rectangle(
        (35, y, largura - 35, y + 58),
        radius=12,
        fill="#263B59"
    )

    draw.text(
        (55, y + 16),
        "VAREJO / OPERAÇÃO",
        font=fonte(22, True),
        fill="#FFFFFF"
    )

    draw.text(
        (largura - 55, y + 16),
        "ATIVAÇÕES | REALIZADO | PROJEÇÃO",
        font=fonte(19, True),
        fill="#FFFFFF",
        anchor="ra"
    )

    y += altura_cabecalho

    for _, linha in tabela.iterrows():
        total = linha["Varejo"] == "TOTAL"

        cor_fundo = (
            "#174B3C" if total
            else "#19263D"
        )

        draw.rounded_rectangle(
            (35, y, largura - 35, y + 96),
            radius=12,
            fill=cor_fundo
        )

        draw.text(
            (55, y + 10),
            str(linha["Varejo"]),
            font=fonte(25, True),
            fill="#FFFFFF"
        )

        draw.text(
            (55, y + 53),
            f"Ativações: {inteiro(linha['Ativacoes'])}",
            font=fonte(20),
            fill="#CBD5E1"
        )

        draw.text(
            (largura - 55, y + 39),
            f"Realizado: {moeda(linha['Comissao_Realizada'])}",
            font=fonte(20, True),
            fill="#FFFFFF",
            anchor="ra"
        )

        draw.text(
            (largura - 55, y + 69),
            f"Projeção: {moeda(linha['Projecao_Receita'])}",
            font=fonte(19),
            fill="#A7F3D0" if total else "#A8BAD3",
            anchor="ra"
        )

        y += altura_linha

    buffer = BytesIO()
    imagem.save(buffer, format="PNG", optimize=True)

    return buffer.getvalue()






def variacao(atual, anterior):

    if pd.isna(anterior) or anterior == 0:

        return None

    return (atual - anterior) / abs(anterior) * 100



def texto_variacao(valor):

    return f"{valor:+.2f}%" if valor is not None and pd.notna(valor) else "N/D"



# ======================== CARREGAR E PREPARAR ========================

operadoras = carregar_excel(str(ARQUIVO_OPERADORAS))

historico = carregar_excel(str(ARQUIVO_HISTORICO))

starlink = consolidar_starlink(

    tuple(str(arquivo) for arquivo in ARQUIVOS_STARLINK)

)

if ARQUIVO_METAS.exists():
    metas_mensais = carregar_metas(str(ARQUIVO_METAS))
else:
    metas_mensais = {}
    st.warning("Arquivo de metas não encontrado.")



# Padroniza o nome da coluna de preço, inclusive quando há espaços extras.
for base in (operadoras, historico):
    base.columns = base.columns.astype(str).str.strip()
    if "Valor" not in base.columns and "valor" in base.columns:
        base.rename(columns={"valor": "Valor"}, inplace=True)

colunas_operadoras = {"NomeRede", "TipoPlano", "DataPedido", "Valor"}

colunas_starlink = {"affiliate_partner_oe", "activated_date"}

if not colunas_operadoras.issubset(operadoras.columns):

    st.error(f"Colunas ausentes em operadoras.xlsx: {sorted(colunas_operadoras - set(operadoras.columns))}")

    st.stop()

if not colunas_operadoras.issubset(historico.columns):

    st.error(f"Colunas ausentes em historico_operadoras.xlsx: {sorted(colunas_operadoras - set(historico.columns))}")

    st.stop()

if not colunas_starlink.issubset(starlink.columns):

    st.error(f"Colunas ausentes em starlink.csv: {sorted(colunas_starlink - set(starlink.columns))}")

    st.stop()



# Valores fixos preservados para redes não abrangidas pela nova regra.
regras_comissao_fixa = {
    "magazine luiza - pré": 1.00,
    "lasa": 8.30,
    "samsung - allied": 40.00,
    "magazine luiza": 13.5
}

percentuais_comissao = {
    
    "gbarbosa": 0.14,
    "motorola": 0.12,
    "nagem": 0.12,
}


def converter_valor_plano(serie):
    """Converte valores numéricos ou textos monetários BR para reais."""
    texto = serie.astype("string").str.strip().str.replace("R$", "", regex=False)
    texto = texto.str.replace(r"\s+", "", regex=True)
    # Se houver vírgula, ela é o separador decimal (ex.: 1.234,56).
    com_virgula = texto.str.contains(",", na=False)
    texto.loc[com_virgula] = (
        texto.loc[com_virgula]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(texto, errors="coerce")


def preparar_operadoras(df):
    df = df.copy()
    df["NomeRede"] = df["NomeRede"].astype("string").str.strip()
    df["TipoPlano"] = df["TipoPlano"].astype("string").str.strip()
    df["DataPedido"] = pd.to_datetime(df["DataPedido"], errors="coerce", dayfirst=True)

    nome = df["NomeRede"].str.casefold()
    pre = nome.isin(["magalu", "magazine luiza"]) & df["TipoPlano"].str.casefold().eq("pre")
    df.loc[pre, "NomeRede"] = "Magazine Luiza - Pré"
    df.loc[df["NomeRede"].str.casefold().eq("magalu"), "NomeRede"] = "Magazine Luiza"

    chave = normalizar_texto(df["NomeRede"])
    df["ValorPlano"] = converter_valor_plano(df["Valor"])
    df["PercentualComissao"] = chave.map(percentuais_comissao).astype("float64")

    # As faixas são mensais por rede, considerando toda a base do mês,
    # independentemente do intervalo selecionado no filtro do dashboard.
    mes = df["DataPedido"].dt.to_period("M")
    quantidade_mes = df.groupby([chave, mes], dropna=False)["NomeRede"].transform("size")
    df.loc[chave.eq("guaibim"), "PercentualComissao"] = (
        quantidade_mes.loc[chave.eq("guaibim")].ge(2001).map({True: 0.12, False: 0.10})
    )
    df.loc[chave.eq("bemol"), "PercentualComissao"] = (
        quantidade_mes.loc[chave.eq("bemol")].ge(3001).map({True: 0.14, False: 0.16})
    )

    # Comissão percentual por ativação; valor ausente permanece sem comissão
    # calculada para permitir auditoria (não é tratado como plano gratuito).
    df["ComissaoAllied"] = df["ValorPlano"] * df["PercentualComissao"]

    # Preservar comissões fixas das demais redes.
    fixa = chave.map(regras_comissao_fixa)
    df.loc[fixa.notna(), "ComissaoAllied"] = fixa[fixa.notna()]
    df.loc[chave.str.contains("samsung", na=False), "ComissaoAllied"] = 40.0
    return df


def preparar_starlink(df):

    df = df.copy()



    df["activated_date"] = pd.to_datetime(

        df["activated_date"],

        errors="coerce",

        utc=True

    ).dt.tz_localize(None)



    # Padronizar os nomes dos planos

    planos_validos = {

    "roam-100gb": "Viagem 100GB",

    "roam - 100gb": "Viagem 100GB",

    "residential 100 mbps": "Residencial 100 Mbps",

    "residential max": "Residencial Max",

    "roam - unlimited": "Viagem Ilimitado"

}



    plano = (

        df["activated_plan"]

        .fillna("")

        .astype(str)

        .str.strip()

        .str.lower()

        .str.replace("–", "-", regex=False)

        .str.replace("—", "-", regex=False)

        .str.replace(r"\s+", " ", regex=True)

    )



    df["Plano"] = plano.map(planos_validos)



    # Somente planos autorizados

    df = df[df["Plano"].notna()].copy()



    # Identificar a rede

    original = (

        df["affiliate_partner_oe"]

        .fillna("")

        .astype(str)

        .str.strip()

    )



    chave = original.str.lower().str.replace(

        r"\s+", "", regex=True

    )



    de_para = {

        "varejo_digital": "Mobcom",

        "varejodigital": "Mobcom",

        "varejo": "Mobcom",

        "mobcom": "Mobcom",

        "": "Mobcom",

        "nan": "Mobcom",

        "havan": "Havan",

        "magalu": "Magazine Luiza",

        "magazineluiza": "Magazine Luiza"

    }



    df["Rede"] = chave.map(de_para).fillna(original)



    # Comissão Allied: R$ 200 para todas as redes
    df["Comissão Allied"] = 200.0



    # Regras de repasse ao parceiro

    df["Repasse Parceiro"] = df["Rede"].eq("Mobcom").map({

        True: 0.0,

        False: 100.0

    })



    df["Receita"] = df["Comissão Allied"]

    df["Repasse"] = df["Repasse Parceiro"]



    return df



operadoras = preparar_operadoras(operadoras)

historico = preparar_operadoras(historico)



# Aplicar filtros e regras comerciais
starlink = preparar_starlink(starlink)


# ======================== FILTROS ========================

st.sidebar.header("Filtros")

data_operadoras = pd.to_datetime(

    operadoras["DataPedido"], errors="coerce", utc=True

).dt.tz_localize(None).max()



data_starlink = pd.to_datetime(

    starlink["activated_date"], errors="coerce", utc=True

).dt.tz_localize(None).max()



datas_validas = [

    data for data in [data_operadoras, data_starlink]

    if pd.notna(data)

]



max_data = max(

    datas_validas,

    default=pd.Timestamp("2026-10-23")



)



exibir_cabecalho(max_data)



# ======================== FILTRO DINÂMICO D-1 ========================

# Data atual
hoje = pd.Timestamp.now(tz="America/Sao_Paulo").date()

# Primeiro dia do mês vigente
primeiro_dia_mes = hoje.replace(day=1)

# Data de ontem (D-1)
ontem = hoje - pd.Timedelta(days=1)

# Tratamento para o primeiro dia do mês
# Quando hoje é dia 1, considera o mês anterior completo
if hoje.day == 1:
    data_inicio_padrao = ontem.replace(day=1)
else:
    data_inicio_padrao = primeiro_dia_mes

# Filtros de data
data_inicio = st.sidebar.date_input(
    "Data inicial",
    value=data_inicio_padrao
)

data_fim = st.sidebar.date_input(
    "Data final",
    value=ontem
)

if data_inicio > data_fim:

    st.sidebar.error("A data inicial deve ser anterior ou igual à data final.")

    st.stop()

operadoras_f = operadoras[operadoras["DataPedido"].dt.date.between(data_inicio, data_fim)].copy()


historico_datas = historico.dropna(subset=["DataPedido"]).copy()

starlink_f = starlink[starlink["activated_date"].dt.date.between(data_inicio, data_fim)].copy()

# ======================== PROJEÇÕES E COMPARATIVOS ========================

datas_validas = operadoras_f["DataPedido"].dropna()

if datas_validas.empty:

    dia_ref, percentual_esperado, comissao_realizada, projecao_comissao = 0, 0.0, 0.0, 0.0

else:

    dia_ref = min(datas_validas.max().day, 30)

    percentual_esperado = 0.60 * dia_ref / 15 if dia_ref <= 15 else 0.60 + 0.40 * (dia_ref - 15) / 15

    comissao_realizada = pd.to_numeric(operadoras_f["ComissaoAllied"], errors="coerce").sum()

    projecao_comissao = comissao_realizada / percentual_esperado if percentual_esperado else 0.0

if datas_validas.empty:

    comissao_realizada = pd.to_numeric(operadoras_f["ComissaoAllied"], errors="coerce").sum()



# Histórico: compara o mesmo intervalo selecionado, deslocado um mês e um ano.

inicio_atual = pd.Timestamp(data_inicio)

fim_atual = datas_validas.max().normalize() if not datas_validas.empty else pd.Timestamp(data_fim)

inicio_mom, fim_mom = inicio_atual - pd.DateOffset(months=1), fim_atual - pd.DateOffset(months=1)

inicio_yoy, fim_yoy = inicio_atual - pd.DateOffset(years=1), fim_atual - pd.DateOffset(years=1)

def dados_periodo(inicio, fim):

    return historico_datas[historico_datas["DataPedido"].between(inicio, fim)]

hist_mom, hist_yoy = dados_periodo(inicio_mom, fim_mom), dados_periodo(inicio_yoy, fim_yoy)

ativ_varejo = len(operadoras_f)

ativ_mom, ativ_yoy = len(hist_mom), len(hist_yoy)

com_mom = pd.to_numeric(hist_mom["ComissaoAllied"], errors="coerce").sum()

com_yoy = pd.to_numeric(hist_yoy["ComissaoAllied"], errors="coerce").sum()

var_ativ_mom, var_ativ_yoy = variacao(ativ_varejo, ativ_mom), variacao(ativ_varejo, ativ_yoy)

var_com_mom, var_com_yoy = variacao(comissao_realizada, com_mom), variacao(comissao_realizada, com_yoy)



# ================= PROJEÇÃO STARLINK - ÚLTIMOS 7 DIAS =================



# Garantir que todas as ativações tenham uma rede
starlink_f["Rede"] = (
    starlink_f["Rede"]
    .fillna("Mobcom")
    .replace("", "Mobcom")
)

# Consolidar todas as ativações por rede
resumo_star = (
    starlink_f.groupby("Rede", dropna=False)
    .agg(
        Ativações=("Rede", "size"),
        Receita_Realizada=("Receita", "sum"),
        Repasse_Parceiro=("Repasse", "sum")
    )
    .reset_index()
)

# Conferência entre visão geral e gráfico
total_geral = len(starlink_f)
total_grafico = resumo_star["Ativações"].sum()

if total_geral != total_grafico:
    st.error(
        f"Divergência Starlink: "
        f"Visão geral = {total_geral} | "
        f"Gráfico = {total_grafico}"
    )


resumo_star["Projecao_Receita"] = resumo_star["Receita_Realizada"]

resumo_star["Projecao_Ativacoes"] = resumo_star["Ativações"].astype(float)



datas_star = starlink_f["activated_date"].dropna()



if not datas_star.empty:



    data_ref_star = datas_star.max().normalize()
    inicio_mes_star = data_ref_star.replace(day=1)
    fim_mes_star = data_ref_star + pd.offsets.MonthEnd(0)

    # Considera o encerramento do mês
    mes_encerrado = pd.Timestamp(data_fim) >= fim_mes_star

    if mes_encerrado:
        dias_restantes = 0



    inicio_7_dias = max(

        inicio_mes_star,

        data_ref_star - pd.Timedelta(days=6)

    )



    dias_analisados = (data_ref_star - inicio_7_dias).days + 1

    dias_restantes = (
        0 if mes_encerrado
        else max((fim_mes_star - data_ref_star).days, 0)
        )



    ultimos_7 = starlink_f[

        starlink_f["activated_date"].between(

            inicio_7_dias,

            data_ref_star + pd.Timedelta(days=1),

            inclusive="left"

        )

    ]



    resumo_7 = (

        ultimos_7.groupby("Rede")

        .agg(

            Ativacoes_7=("Rede", "size"),

            Receita_7=("Receita", "sum")

        )

        .reset_index()

    )



    resumo_star = resumo_star.merge(

        resumo_7,

        on="Rede",

        how="left"

    )



    resumo_star[["Ativacoes_7", "Receita_7"]] = (

        resumo_star[["Ativacoes_7", "Receita_7"]].fillna(0)

    )



# Projeção de ativações futuras
ativacoes_futuras = (
    resumo_star["Ativacoes_7"]
    / dias_analisados
    * dias_restantes
).clip(lower=0).round().astype(int)

# Projeção total de ativações
resumo_star["Projecao_Ativacoes"] = (
    resumo_star["Ativações"] + ativacoes_futuras
)

# Comissão Allied: R$ 200 para todas as redes
resumo_star["Comissao_Unitária"] = 200.0

# Projeção de receita Allied
resumo_star["Projecao_Receita"] = (
    resumo_star["Receita_Realizada"]
    + ativacoes_futuras * resumo_star["Comissao_Unitária"]
)

# Validação: projeção não pode ser inferior ao realizado
assert (
    resumo_star["Projecao_Receita"]
    >= resumo_star["Receita_Realizada"]
).all(), "Erro: projeção Starlink inferior à receita realizada."

# Validação de fechamento do mês
if dias_restantes == 0:
    resumo_star["Projecao_Ativacoes"] = resumo_star["Ativações"]
    resumo_star["Projecao_Receita"] = resumo_star["Receita_Realizada"]

resumo_star = resumo_star.drop(

        columns=["Ativacoes_7", "Receita_7"]

    )



receita_star = resumo_star["Receita_Realizada"].sum()

projecao_star = resumo_star["Projecao_Receita"].sum()



# Receita total realizada - Varejo + Starlink

receita_total = comissao_realizada + receita_star



# Projeção total - Varejo + Starlink

projecao_total = projecao_comissao + projecao_star

# ================= TABELA CONSOLIDADA WHATSAPP =================

# OPERADORAS
tabela_operadoras = (
    operadoras_f
    .assign(
        Varejo=operadoras_f["NomeRede"].fillna("Sem Rede")
    )
    .groupby("Varejo", as_index=False)
    .agg(
        Ativacoes=("Varejo", "size"),
        Comissao_Realizada=("ComissaoAllied", "sum")
    )
)

tabela_operadoras["Projecao_Receita"] = (
    tabela_operadoras["Comissao_Realizada"] / percentual_esperado
    if percentual_esperado > 0 else 0.0
)

# STARLINK
tabela_starlink = resumo_star[
    ["Rede", "Ativações", "Receita_Realizada", "Projecao_Receita"]
].copy()

tabela_starlink = tabela_starlink.rename(
    columns={
        "Rede": "Varejo",
        "Ativações": "Ativacoes",
        "Receita_Realizada": "Comissao_Realizada"
    }
)

# Identificar as operações Starlink
tabela_starlink["Varejo"] = (
    tabela_starlink["Varejo"]
    .fillna("Mobcom")
    .astype(str)
    .str.strip()
    + " – Starlink"
)

# Consolidar as duas operações
tabela_whatsapp = pd.concat(
    [tabela_operadoras, tabela_starlink],
    ignore_index=True
)

# Agrupar eventuais nomes repetidos
tabela_whatsapp = (
    tabela_whatsapp
    .groupby("Varejo", as_index=False)
    .agg(
        Ativacoes=("Ativacoes", "sum"),
        Comissao_Realizada=("Comissao_Realizada", "sum"),
        Projecao_Receita=("Projecao_Receita", "sum")
    )
)

# Ordenar do maior para o menor
tabela_whatsapp = tabela_whatsapp.sort_values(
    ["Comissao_Realizada", "Varejo"],
    ascending=[False, True]
).reset_index(drop=True)

# Total geral
linha_total = pd.DataFrame([{
    "Varejo": "TOTAL",
    "Ativacoes": tabela_whatsapp["Ativacoes"].sum(),
    "Comissao_Realizada": tabela_whatsapp["Comissao_Realizada"].sum(),
    "Projecao_Receita": tabela_whatsapp["Projecao_Receita"].sum()
}])

tabela_whatsapp = pd.concat(
    [tabela_whatsapp, linha_total],
    ignore_index=True
)

# ======================== CARDS ========================



# ======================== META VAREJO DIGITAL ========================

st.subheader("Acompanhamento de Meta — Varejo Digital")

# Identificar o mês do filtro
mes_referencia = data_fim.month
ano_referencia = data_fim.year

# O arquivo atual contém metas de 2026
meta_mensal = (
    metas_mensais.get(mes_referencia)
    if ano_referencia == 2026 else None
)

# Validar se o filtro pertence a um único mês
mesmo_mes = (
    data_inicio.year == data_fim.year
    and data_inicio.month == data_fim.month
)

# Validar se o período começa no primeiro dia do mês
periodo_acumulado = data_inicio.day == 1

if not mesmo_mes:
    st.warning(
        "Selecione datas dentro do mesmo mês para acompanhar "
        "a meta mensal."
    )

elif not periodo_acumulado:
    st.warning(
        "Para comparar o realizado com a meta mensal, "
        "selecione o período a partir do dia 1."
    )

elif meta_mensal is None or meta_mensal <= 0:
    st.info("Meta não cadastrada para o mês selecionado.")

else:
    # Valores já calculados no dashboard
    realizado_meta = float(receita_total)
    projetado_meta = float(projecao_total)

    # Percentuais
    atingimento_realizado = (
        realizado_meta / meta_mensal * 100
    )

    atingimento_projetado = (
        projetado_meta / meta_mensal * 100
    )

    # Diferenças
    valor_faltante = max(
        meta_mensal - realizado_meta, 0
    )

    diferenca_projetada = (
        projetado_meta - meta_mensal
    )

    # Primeira linha
    m1, m2, m3 = st.columns(3)

    with m1:
        exibir_card(
            "Meta Mensal",
            formatar_reais(meta_mensal)
        )

    with m2:
        exibir_card(
            "Receita Realizada",
            formatar_reais(realizado_meta)
        )

    with m3:
            exibir_card(
                "Valor Restante para Meta",
                formatar_reais(valor_faltante)
            )

    

    # Segunda linha
    m4, m5, m6 = st.columns(3)

    with m4:
        exibir_card(
            "Atingimento da Meta",
            f"{atingimento_realizado:.2f}%".replace(".", ",")
        )

    with m5:
        exibir_card(
            "Projeção de Receita",
            formatar_reais(projetado_meta)
        )

    with m6:
        exibir_card(
            "Atingimento Projetado",
            f"{atingimento_projetado:.2f}%".replace(".", ",")
        )

    

    

    # Resultado projetado
    if diferenca_projetada >= 0:
        st.success(
            "Projeção acima da meta em "
            f"{formatar_reais(diferenca_projetada)}"
        )
    else:
        st.warning(
            "Projeção abaixo da meta em "
            f"{formatar_reais(abs(diferenca_projetada))}"
        )

    st.caption(
        f"Período: {data_inicio:%d/%m/%Y} "
        f"até {data_fim:%d/%m/%Y} | "
        "Receita Allied de Operadoras + Starlink"
    )

st.subheader("Visão geral")

c1, c2, c3 = st.columns(3)

with c1:

    exibir_card(

        "Ativações - Varejo",

        f"{ativ_varejo:,}".replace(",", "."),

        indicador_colorido("MoM", var_ativ_mom)

        + "  |  "

        + indicador_colorido("YoY", var_ativ_yoy)

    )



with c2:

    exibir_card(

        "Receita Realizada - Varejo",

        formatar_reais(comissao_realizada),

        indicador_colorido("MoM", var_com_mom)

        + "  |  "

        + indicador_colorido("YoY", var_com_yoy)

    )

with c3: exibir_card("Projeção Receita - Varejo", formatar_reais(projecao_comissao))

c4, c5, c6 = st.columns(3)

with c4: exibir_card("Ativações - Starlink", f"{len(starlink_f):,}".replace(",", "."))

with c5: exibir_card("Receita Realizada - Starlink", formatar_reais(receita_star))

with c6: exibir_card("Projeção Receita - Starlink", formatar_reais(projecao_star))

c7, c8, c9 = st.columns(3)

with c7: exibir_card("Ativações - Varejo + Starlink", f"{ativ_varejo + len(starlink_f):,}".replace(",", "."))

with c8: exibir_card("Receita Realizada - Varejo + Starlink", formatar_reais(receita_total))

with c9: exibir_card("Projeção Receita - Varejo + Starlink", formatar_reais(projecao_total))

if dia_ref:

    st.caption(f"Projeção Varejo baseada no dia {dia_ref}, com {percentual_esperado:.2%} do volume mensal esperado.")



# ======================== RELATÓRIO POR REDE ========================

st.subheader("Relatório Diário - Varejo")

def resumo_por_periodo(df, sufixo):

    x = df.copy()

    x["Rede"] = x["NomeRede"].fillna("Sem Rede")

    x["_chave"] = normalizar_texto(x["Rede"])

    return x.groupby("_chave").agg(**{f"Ativacoes_{sufixo}":("NomeRede", "size"), f"Comissao_{sufixo}":("ComissaoAllied", "sum")}).reset_index(), x

resumo_varejo = operadoras_f.assign(Rede=operadoras_f["NomeRede"].fillna("Sem Rede"), _chave=lambda x: normalizar_texto(x["NomeRede"]))

resumo_varejo = resumo_varejo.groupby(["Rede", "_chave"], dropna=False).agg(Ativacoes=("NomeRede", "size"), Comissao_Allied=("ComissaoAllied", "sum")).reset_index()

hm, dm = resumo_por_periodo(hist_mom, "MoM")

hy, dy = resumo_por_periodo(hist_yoy, "YoY")

resumo_varejo = resumo_varejo.merge(hm, on="_chave", how="left").merge(hy, on="_chave", how="left")

for atual, anterior, nome in [("Ativacoes", "Ativacoes_MoM", "MoM Ativações (%)"), ("Ativacoes", "Ativacoes_YoY", "YoY Ativações (%)"), ("Comissao_Allied", "Comissao_MoM", "MoM Comissão Allied (%)"), ("Comissao_Allied", "Comissao_YoY", "YoY Comissão Allied (%)")]:

    resumo_varejo[nome] = resumo_varejo.apply(lambda r: variacao(r[atual], r[anterior]) if pd.notna(r[anterior]) else None, axis=1)

resumo_varejo["Projecao_Allied"] = resumo_varejo["Comissao_Allied"] / percentual_esperado if percentual_esperado else 0

resumo_varejo = resumo_varejo.sort_values("Comissao_Allied", ascending=False)

linha_total = {"Rede":"TOTAL", "Ativacoes":ativ_varejo, "Comissao_Allied":comissao_realizada, "Projecao_Allied":resumo_varejo["Projecao_Allied"].sum(), "MoM Ativações (%)":variacao(ativ_varejo, ativ_mom), "YoY Ativações (%)":variacao(ativ_varejo, ativ_yoy), "MoM Comissão Allied (%)":variacao(comissao_realizada, com_mom), "YoY Comissão Allied (%)":variacao(comissao_realizada, com_yoy)}

resumo_varejo = resumo_varejo.drop(columns=["_chave", "Ativacoes_MoM", "Comissao_MoM", "Ativacoes_YoY", "Comissao_YoY"], errors="ignore")

resumo_varejo = pd.concat([resumo_varejo, pd.DataFrame([linha_total])], ignore_index=True)

resumo_varejo = resumo_varejo[["Rede", "Ativacoes", "MoM Ativações (%)", "YoY Ativações (%)", "Comissao_Allied", "MoM Comissão Allied (%)", "YoY Comissão Allied (%)", "Projecao_Allied"]]

def colorir_percentual(v):

    if pd.isna(v): return "color:#808080"

    if abs(v) < 1e-6: return "color:#D4A017;font-weight:bold"

    return "color:#16803D;font-weight:bold" if v > 0 else "color:#D32F2F;font-weight:bold"

cols_pct = ["MoM Ativações (%)", "YoY Ativações (%)", "MoM Comissão Allied (%)", "YoY Comissão Allied (%)"]

# Formatação do Relatório Diário - Varejo
estilo = (
    resumo_varejo.style
    .map(colorir_percentual, subset=cols_pct)
    .apply(
        lambda linha: [
            "font-weight: bold; background-color: #24344D"
            if linha["Rede"] == "TOTAL" else ""
            for _ in linha
        ],
        axis=1
    )
    .format(
        {
            "Ativacoes": "{:,.0f}",
            "Comissao_Allied": "R$ {:,.2f}",
            "Projecao_Allied": "R$ {:,.2f}",
            **{coluna: "{:+.2f}%" for coluna in cols_pct}
        },
        na_rep="N/D",
        decimal=",",
        thousands="."
    )
)

st.dataframe(
    estilo,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Rede": "Rede",
        "Ativacoes": "Ativações",
        "Comissao_Allied": "Comissão Allied Realizada",
        "Projecao_Allied": "Projeção Comissão Allied"
    }
)


# ======================== GRÁFICOS EM DUAS COLUNAS ========================

st.subheader("Análise de Operadoras")

g1, g2 = st.columns(2, gap="large")

with g1:

    comissoes_rede = (

        operadoras_f.assign(

            NomeRede=operadoras_f["NomeRede"].fillna("Sem Rede")

        )

        .groupby("NomeRede")["ComissaoAllied"]

        .sum()

        .reset_index()

    )



    comissoes_rede.columns = ["Rede", "Comissão Allied Realizada"]



    comissoes_rede["Projeção Comissão Allied"] = (

        comissoes_rede["Comissão Allied Realizada"] / percentual_esperado

        if percentual_esperado else 0

    )



    comissoes_rede = comissoes_rede.sort_values(

        "Comissão Allied Realizada", ascending=False

    )



    dados_com = comissoes_rede.melt(

        id_vars="Rede",

        value_vars=["Comissão Allied Realizada", "Projeção Comissão Allied"],

        var_name="Indicador",

        value_name="Receita (R$)"

    )



    fig_com = px.bar(

        dados_com,

        x="Rede",

        y="Receita (R$)",

        color="Indicador",

        barmode="group",

        text="Receita (R$)",

        title="Comissão Allied realizada x projeção",

        color_discrete_map={

            "Comissão Allied Realizada": "#2F6BCE",

            "Projeção Comissão Allied": "#0615A3"

        },

        category_orders={"Rede": comissoes_rede["Rede"].tolist()}

    )



    fig_com.update_traces(

        texttemplate="R$ %{y:,.0f}",

        textposition="outside",

        hovertemplate="<b>%{x}</b><br>%{fullData.name}<br>R$ %{y:,.2f}<extra></extra>"

    )



    aplicar_estilo_moderno(fig_com, barras=True, moeda=True)



    st.plotly_chart(

        fig_com,

        use_container_width=True,

        key="grafico_comissoes_varejo"

    )





with g2:

    ativ_rede = (

        operadoras_f.assign(

            NomeRede=operadoras_f["NomeRede"].fillna("Sem Rede")

        )

        .groupby("NomeRede")

        .size()

        .reset_index(name="Ativações Realizadas")

        .rename(columns={"NomeRede": "Rede"})

    )



    ativ_rede["Projeção de Ativações"] = (

        ativ_rede["Ativações Realizadas"] / percentual_esperado

        if percentual_esperado else 0

    )



    ativ_rede = ativ_rede.sort_values(

        "Ativações Realizadas", ascending=False

    )



    dados_ativ = ativ_rede.melt(

        id_vars="Rede",

        value_vars=["Ativações Realizadas", "Projeção de Ativações"],

        var_name="Indicador",

        value_name="Ativação"

    )



    fig_ativ = px.bar(

        dados_ativ,

        x="Rede",

        y="Ativação",

        color="Indicador",

        barmode="group",

        text="Ativação",

        title="Ativações realizadas x projeção",

        color_discrete_map={

            "Ativações Realizadas": "#2F6BCE",

            "Projeção de Ativações": "#0615A3"

        },

        category_orders={"Rede": ativ_rede["Rede"].tolist()}

    )



    fig_ativ.update_traces(

        texttemplate="%{y:,.0f}",

        textposition="outside"

    )



    aplicar_estilo_moderno(fig_ativ, barras=True)



    st.plotly_chart(

        fig_ativ,

        use_container_width=True,

        key="grafico_ativacoes_varejo"

    )



# ======================== CONFERÊNCIA E EVOLUÇÃO DIÁRIA ========================

# ================= CONFERÊNCIA DE COMISSÕES STARLINK =================

st.subheader("Conferência de comissões Starlink")

resumo_comissoes = (
    starlink_f.groupby("Rede", dropna=False)
    .agg(
        Ativações=("Rede", "size"),
        Comissão_Allied=("Comissão Allied", "sum"),
        Repasse_Parceiro=("Repasse Parceiro", "sum")
    )
    .reset_index()
    .sort_values("Comissão_Allied", ascending=False)
)

# Adicionar linha TOTAL
linha_total_starlink = pd.DataFrame([{
    "Rede": "TOTAL",
    "Ativações": resumo_comissoes["Ativações"].sum(),
    "Comissão_Allied": resumo_comissoes["Comissão_Allied"].sum(),
    "Repasse_Parceiro": resumo_comissoes["Repasse_Parceiro"].sum()
}])

resumo_comissoes = pd.concat(
    [resumo_comissoes, linha_total_starlink],
    ignore_index=True
)

# Formatação da tabela
estilo_starlink = resumo_comissoes.style.format(
    {
        "Ativações": "{:,.0f}",
        "Comissão_Allied": "R$ {:,.2f}",
        "Repasse_Parceiro": "R$ {:,.2f}"
    },
    decimal=",",
    thousands="."
).apply(
    lambda linha: [
        "font-weight: bold; background-color: #24344D"
        if linha["Rede"] == "TOTAL" else ""
        for _ in linha
    ],
    axis=1
)

st.dataframe(
    estilo_starlink,
    use_container_width=True,
    hide_index=True
)

# ================= RANKING DE PLANOS STARLINK =================



st.subheader("Ranking de Planos Mais Contratados - Starlink")



ranking_planos = starlink_f.copy()



# Plano efetivamente ativado

ranking_planos["Plano"] = (

    ranking_planos["activated_plan"]

    .fillna("Não informado")

    .astype(str)

    .str.strip()

    .replace("", "Não informado")

)



# Valor do plano

ranking_planos["Valor_Plano"] = pd.to_numeric(

    ranking_planos["activated_plan_cost"],

    errors="coerce"

).fillna(0)



# Agrupamento por plano

ranking_planos = (

    ranking_planos.groupby("Plano", as_index=False)

    .agg(

        Ativacoes=("Plano", "size"),

        Receita_Allied=("Comissão Allied", "sum"),

        Valor_Total_Planos=("Valor_Plano", "sum")

    )

    .sort_values("Ativacoes", ascending=False)

    .reset_index(drop=True)

)



# Participação percentual

total_ativacoes = ranking_planos["Ativacoes"].sum()



ranking_planos["Participacao"] = (

    ranking_planos["Ativacoes"] / total_ativacoes * 100

    if total_ativacoes else 0

)



# Posição no ranking

ranking_planos.insert(

    0, "Ranking", range(1, len(ranking_planos) + 1)

)



# Exibição

st.dataframe(

    ranking_planos[

        ["Ranking", "Plano", "Ativacoes", "Participacao"]

    ],

    use_container_width=True,

    hide_index=True,

    column_config={

        "Ranking": st.column_config.NumberColumn(

            "Posição", format="%d"

        ),

        "Plano": "Plano contratado",

        "Ativacoes": st.column_config.NumberColumn(

            "Ativações", format="%d"

        ),

        "Participacao": st.column_config.ProgressColumn(

            "Participação (%)",

            min_value=0,

            max_value=100,

            format="%.1f%%"

        )

    }

)



st.subheader("Análise de Starlink")

g3, g4 = st.columns(2, gap="large")

with g3:

    dados_receita = resumo_star.melt(

        id_vars="Rede",

        value_vars=[

            "Receita_Realizada",

            "Repasse_Parceiro",

            "Projecao_Receita"

        ],

        var_name="Indicador",

        value_name="Receita (R$)"

    )



    dados_receita["Indicador"] = dados_receita["Indicador"].map({

        "Receita_Realizada": "Receita realizada",

        "Repasse_Parceiro": "Repasse ao parceiro",

        "Projecao_Receita": "Projeção de receita"

    })



    fig_receita = px.bar(

        dados_receita,

        x="Rede",

        y="Receita (R$)",

        color="Indicador",

        barmode="group",

        text="Receita (R$)",

        title="Receita x repasse x projeção",

        color_discrete_map={

            "Receita realizada": "#2F6BCE",

            "Repasse ao parceiro": "#39B9DA",

            "Projeção de receita": "#0615A3"

        },

        category_orders={

            "Rede": ["Mobcom", "Havan", "Magazine Luiza"],

            "Indicador": [

                "Receita realizada",

                "Repasse ao parceiro",

                "Projeção de receita"

            ]

        }

    )



    fig_receita.update_traces(

        texttemplate="R$ %{y:,.0f}",

        textposition="outside",

        hovertemplate=(

            "<b>%{x}</b><br>"

            "%{fullData.name}<br>"

            "R$ %{y:,.2f}<extra></extra>"

        )

    )



    aplicar_estilo_moderno(

        fig_receita,

        barras=True,

        moeda=True

    )



    st.plotly_chart(

        fig_receita,

        use_container_width=True,

        key="grafico_receita_starlink"

    )





with g4:

    dados_ativ_star = resumo_star[

        ["Rede", "Ativações", "Projecao_Ativacoes"]

    ].copy()



    dados_ativ_star = dados_ativ_star.rename(

        columns={

            "Projecao_Ativacoes": "Projeção de Ativações"

        }

    )



    dados_ativ_star = dados_ativ_star.sort_values(

        "Ativações",

        ascending=False

    )



    dados_star = dados_ativ_star.melt(

        id_vars="Rede",

        value_vars=[

            "Ativações",

            "Projeção de Ativações"

        ],

        var_name="Indicador",

        value_name="Ativação"

    )



    fig_star = px.bar(

        dados_star,

        x="Rede",

        y="Ativação",

        color="Indicador",

        barmode="group",

        text="Ativação",

        title="Ativações Starlink realizadas x projeção",

        color_discrete_map={

            "Ativações": "#2F6BCE",

            "Projeção de Ativações": "#0615A3"

        },

        category_orders={

            "Rede": dados_ativ_star["Rede"].tolist()

        }

    )



    fig_star.update_traces(

        texttemplate="%{y:,.0f}",

        textposition="outside"

    )



    aplicar_estilo_moderno(

        fig_star,

        barras=True

    )



    st.plotly_chart(

        fig_star,

        use_container_width=True,

        key="grafico_ativacoes_starlink"

    )







e1, e2 = st.columns(2, gap="large")

with e1:

    st.subheader("Evolução diária - Operadoras")

    diario = operadoras_f.dropna(subset=["DataPedido"]).assign(Data=lambda x: x["DataPedido"].dt.date).groupby("Data").size().reset_index(name="Ativações")

    fig_dia = px.line(diario, x="Data", y="Ativações", markers=True, title="Ativações por dia", color_discrete_sequence=["#2F6BCE"])

    fig_dia.update_traces(line=dict(width=3), marker=dict(size=7))

    aplicar_estilo_moderno(fig_dia)

    st.plotly_chart(fig_dia, use_container_width=True, key="evolucao_operadoras")

with e2:

    st.subheader("Evolução diária - Starlink")

    diario_star = starlink_f.dropna(subset=["activated_date"]).assign(Data=lambda x: x["activated_date"].dt.date).groupby("Data").size().reset_index(name="Ativações")

    fig_dia_star = px.line(diario_star, x="Data", y="Ativações", markers=True, title="Ativações por dia", color_discrete_sequence=["#16A34A"])

    fig_dia_star.update_traces(line=dict(width=3), marker=dict(size=7))

    aplicar_estilo_moderno(fig_dia_star)

    st.plotly_chart(fig_dia_star, use_container_width=True, key="evolucao_starlink")



with st.expander("Consultar dados de Operadoras"):

    st.dataframe(operadoras_f, use_container_width=True, hide_index=True)

with st.expander("Consultar dados de Starlink"):

    st.dataframe(starlink_f, use_container_width=True, hide_index=True)

# ================= RELATÓRIO EXECUTIVO WHATSAPP =================

st.divider()

st.subheader("📱 Relatório Executivo — WhatsApp")

st.caption(
    "Gere uma imagem com os principais indicadores do Varejo Digital "
    "para compartilhar no WhatsApp."
)

# Meta do mês selecionado
mes_referencia = data_fim.month
ano_referencia = data_fim.year

meta_relatorio = (
    metas_mensais.get(mes_referencia)
    if ano_referencia == 2026 else None
)

# Identifica alterações no período selecionado
periodo_relatorio = (
    data_inicio.isoformat(),
    data_fim.isoformat()
)

if st.session_state.get("periodo_relatorio") != periodo_relatorio:
    st.session_state["periodo_relatorio"] = periodo_relatorio
    st.session_state.pop("imagem_relatorio", None)

# Botão para gerar a imagem
if st.button(
    "📸 Gerar imagem do relatório",
    type="primary",
    use_container_width=True,
    key="gerar_relatorio_whatsapp"
):
    try:
        with st.spinner("Gerando imagem do relatório..."):

            imagem_relatorio = gerar_relatorio_whatsapp(
                data_inicio=data_inicio,
                data_fim=data_fim,
                meta_mensal=meta_relatorio,
                comissao_realizada=comissao_realizada,
                receita_star=receita_star,
                receita_total=receita_total,
                projecao_comissao=projecao_comissao,
                projecao_star=projecao_star,
                projecao_total=projecao_total,
                ativ_varejo=ativ_varejo,
                ativ_starlink=len(starlink_f)
            )
            # Acrescentar tabela consolidada ao PNG
            imagem_relatorio = adicionar_tabela_ao_png(
                imagem_relatorio,
                tabela_whatsapp
            )
            st.session_state["imagem_relatorio"] = imagem_relatorio

        st.success("Imagem gerada com sucesso!")

    except Exception as erro:
        st.error(f"Erro ao gerar o relatório: {erro}")

# Exibir prévia e download
if "imagem_relatorio" in st.session_state:

    st.markdown("### Prévia do relatório")

    st.image(
        st.session_state["imagem_relatorio"],
        caption=(
            f"Relatório Allied | "
            f"{data_inicio:%d/%m/%Y} a {data_fim:%d/%m/%Y}"
        ),
        use_container_width=True
    )

    st.download_button(
        label="📥 Baixar imagem para WhatsApp",
        data=st.session_state["imagem_relatorio"],
        file_name=(
            f"Relatorio_Allied_"
            f"{data_fim:%d_%m_%Y}.png"
        ),
        mime="image/png",
        use_container_width=True,
        key="baixar_relatorio_whatsapp"
    )



 # RANKING PLANOS STARLINK
 # py -m streamlit run verificar_bases.py/app2.py
