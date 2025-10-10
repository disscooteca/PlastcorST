import streamlit as st
import os
from PIL import Image
import pandas as pd
import plotly.express as px
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
import datetime
from fpdf.enums import Align
import time
from datetime import datetime, date
import threading
import time
import json


## Configuração da página ##
# Configuração da página
st.set_page_config(
    page_title="PLASTCOR",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

pagina = st.sidebar.selectbox("Escolha a página", ["Home", "Ordem de Serviço", "Produção", #"Fechar Ordem de Serviço", "Ver ordens de Serviço",
                                                   "Quadro de Funcionários", "Falta", "Estampas"])

scopes = [
    "https://spreadsheets.google.com/feeds", 
    "https://www.googleapis.com/auth/drive",
]

if 'gcp_service_account_json' not in st.secrets:
    st.error("JSON da service account não encontrado nos secrets!")
    st.stop()

else:
    service_account_json = st.secrets["gcp_service_account_json"]
    service_account_info = json.loads(service_account_json)
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        keyfile_dict=service_account_info,
        scopes=scopes
    )

client = gspread.authorize(creds) #Acessando sheets

planilha_completa = client.open(
    title="Plastcor_Estampas", 
    folder_id="1CEp2KbtqQnOx3beI2aseUxjPZ1AWe8J6"
    )

planilhaProducao = planilha_completa.get_worksheet(0) #Obtendo planilha Real
planilhaQuadro = planilha_completa.get_worksheet(1) #Obtendo planilha Real
planilhaFalta = planilha_completa.get_worksheet(2) #Obtendo planilha Real
planilhaOS = planilha_completa.get_worksheet(3) #Obtendo planilha Real

dados_Producao_completo = planilhaProducao.get_all_records()
dados_Quadro_completo = planilhaQuadro.get_all_records() #Obtendo todos os dados da planilha
dados_Falta_completo = planilhaFalta.get_all_records() #Obtendo todos os dados da planilha
dados_os_completo = planilhaOS.get_all_records() #Obtendo todos os dados da planilha

## Tamanhos ##
tamanhos = [
    "P", "M", "G", "GG"
]

## Tratamento de dados ##
dados_Producao_completo = pd.DataFrame(dados_Producao_completo) #Colocando dados no formato de dataframe
dados_Quadro_completo = pd.DataFrame(dados_Quadro_completo) #Colocando dados no formato de dataframe
dados_Falta_completo = pd.DataFrame(dados_Falta_completo) #Colocando dados no formato de dataframe
dados_os_completo = pd.DataFrame(dados_os_completo) #Colocando dados no formato de dataframe

## Equipes ##
equipes = dados_Quadro_completo["SETOR"].tolist()

# Pasta onde as imagens estão armazenadas
PASTA_ESTAMPAS = "Estampas"

extensoes_validas = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')

# Lista apenas arquivos com as extensões especificadas
arquivos_imagens = [
    f for f in os.listdir(PASTA_ESTAMPAS)
    if f.lower().endswith(extensoes_validas)
]  

# Lista apenas arquivos com prefixo "b_" e extensões válidas
arquivos_bordado = [
    f for f in os.listdir(PASTA_ESTAMPAS)
    if f.lower().startswith("b_") and f.lower().endswith(extensoes_validas)
]

#Nomes das estampas
nomes_estampas = [os.path.splitext(f)[0] for f in arquivos_imagens]

#Estampas com bordado
estampas_bordado = [os.path.splitext(f)[0] for f in arquivos_bordado]

if 'rotation' not in st.session_state:
    st.session_state['rotation'] = "a"

class ContadorSegundos:
    def __init__(self):
        self.contador = 0
        self.executando = False
        self.thread = None
    
    def iniciar(self):
        self.executando = True
        self.thread = threading.Thread(target=self._contar)
        self.thread.daemon = True  # Thread morre quando programa principal termina
        self.thread.start()
    
    def _contar(self):
        while self.executando:
            time.sleep(1)
            self.contador += 1
            print(f"Contador: {self.contador} segundos")
    
    def parar(self):
        self.executando = False
        if self.thread:
            self.thread.join()
    
    def obter_valor(self):
        return self.contador

## Funções ##
def mostrar_planilha(planilha):
    st.dataframe(planilha)

def buscar_imagem(nome_base):
    pasta = os.path.join(os.path.dirname(__file__), "Estampas")
    for ext in [".jpg", ".png", '.jpeg', '.gif', '.bmp', '.webp']:
        caminho = os.path.join(pasta, f"{nome_base}{ext}")
        if os.path.exists(caminho):
            return caminho
    return None

def create():
    st.header("Informe os dados da OS")

    if not dados_os_completo.empty:

        codigo = dados_os_completo.iloc[-1]["Código OS"] + 1

    else:
        codigo = 1

    codigo = int(codigo)

    # Data de hoje
    hoje = date.today()

    # Formata para dd/mm/yyyy
    data_carimbo = hoje.strftime("%d/%m/%Y")
    

    tipoImpressao = ["Imprimir 1 OS em uma folha", "Imprimir 2 OS em uma folha", "Imprimir 3 OS em uma folha"]
    imprimir = st.radio("Impressão",tipoImpressao)

    if imprimir == "Imprimir 1 OS em uma folha":
    
        st.markdown("---")
        estampa = st.selectbox("Qual modelo de estampa?", nomes_estampas)

        try:
            imagem_estampa = Image.open(f"Estampas\{estampa}.jpg")
            st.image(imagem_estampa)

        except:
            try:
                imagem_estampa = Image.open(f"Estampas\{estampa}.png")
                st.image(imagem_estampa)

            except:
                try:
                    imagem_estampa = Image.open(f"Estampas\{estampa}.webp")
                    st.image(imagem_estampa)

                except:
                    try:
                        imagem_estampa = Image.open(f"Estampas\{estampa}.jpeg")
                        st.image(imagem_estampa)

                    except:
                        st.warning("modelo não encontrado para visualização")

        with st.form("AbrirOS"):
            
            data_entrega = st.date_input("Qual data da entrega prevista?", value="today", min_value="today")
            data_entrega = data_entrega.strftime("%d/%m/%Y")
            
            #st.markdown("---")
            #quantidade = st.number_input("Qual quantidade de camisas?", min_value = 1) #Voltar com quantidade?

            st.markdown("---")
            tamanho = st.radio("Qual tamanho do modelo?", tamanhos)

            st.markdown("---")
            cliente = st.text_input("Qual o nome do cliente?")

            st.markdown("---")
            equipes_filtradas = [e for e in equipes if e != "BORDADO"]
            equipe = st.radio("Qual equipe responsável?", equipes_filtradas)

            if estampa in estampas_bordado:
                bordado = True
            else: 
                bordado = False

            st.markdown("---")
            observacao = st.text_input("Observação?", max_chars=200)

            st.markdown("---")

            submitted = st.form_submit_button("Criar OS!")
                
            # Só executa quando o botão for clicado
            if submitted:

                # Pega todos os valores da primeira coluna
                valores = planilhaOS.col_values(1)

                # Descobre a primeira linha vazia
                linha_vazia = len(valores) + 1  # +1 porque col_values não conta a próxima linha vazia

                #cria nova linha
                nova_linha = [[str(codigo), str(data_carimbo), str(data_entrega), str(estampa),
                    #str(bordado), #str(quantidade),
                    str(tamanho), str(cliente), str(equipe), str(observacao)]]

                planilhaOS.update(f"A{linha_vazia}:K{linha_vazia}", nova_linha)

                if imprimir:
                    pdf = FPDF("landscape", "mm", "A5")
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)

                    pdf.cell(130, 8, "ORDEM DE SERVIÇO", ln=True, align=Align.C)
                    pdf.set_font("Arial", style="B", size=14)
                    pdf.cell(130, 8, f"Cliente: {cliente}", ln=True, border=1)  # Texto em negrito

                    # Volta para a fonte normal (sem negrito)
                    pdf.set_font("Arial", style="", size=12)
                    pdf.cell(130, 8, f"Equipe: {equipe}", ln=True, border=True)
                    #pdf.cell(130, 8, f"Código: {codigo}", ln=True, border=True)
                    pdf.cell(130, 8, f"Estampa: {estampa}", ln=True, border=True)
                    #pdf.cell(130, 8, f"Bordado: {bordado}", ln=True, border=True)
                    #pdf.cell(130, 8, f"Quantidade: {quantidade}", ln=True, border=True)
                    pdf.cell(130, 8, f"Quantidade: ", ln=True, border=True)
                    pdf.cell(130, 8, f"Tamanho: {tamanho}", ln=True, border=True)
                    #pdf.cell(130, 8, f"Data Carimbo: {data_carimbo}", ln=True, border=True)
                    pdf.cell(130, 8, f"Data Entrega Prevista: {data_entrega}", ln=True, border=True)
                    pdf.multi_cell(
                        130, 8,                # largura e altura da linha
                        f"Observação: {observacao}", 
                        border=1,              # desenha borda
                        align= Align.L             # alinhamento do texto
                    )

                    # Posição para começar a coluna da direita
                    x_imagem = 130  # posição horizontal
                    y_imagem = 18   # posição vertical
                    largura_imagem = 70
                    altura_imagem = 100

                    # Desenha a borda grande
                    pdf.rect(x_imagem, y_imagem, largura_imagem, altura_imagem)    

                    ## Colocar imagem na OS ##
                    # imagem_path = buscar_imagem("b_flores")
                    # if imagem_path:
                    try:
                        pdf.image(f"Estampas\\{estampa}.jpg", x=x_imagem+2, y=y_imagem+2, w=largura_imagem-4)

                    except:
                        try:
                            pdf.image(f"Estampas\\{estampa}.png", x=x_imagem+2, y=y_imagem+2, w=largura_imagem-4)

                        except:
                            try:
                                pdf.image(f"Estampas\\{estampa}.webp", x=x_imagem+2, y=y_imagem+2, w=largura_imagem-4)

                            except:
                                try:
                                    pdf.image(f"Estampas\\{estampa}.jpeg", x=x_imagem+2, y=y_imagem+2, w=largura_imagem-4)

                                except:
                                    st.warning("modelo não encontrado para visualização")
                    # else:
                    #     st.write("Imagem não encontrada")    

                    pasta_base = os.path.dirname(__file__)

                    # Caminho relativo para a pasta "OS"
                    caminho_pasta = os.path.join(pasta_base, "OS")

                    # Salva o PDF na pasta "OS"
                    pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente}.pdf"))

                st.toast('Ordem de seviço criada!', icon='🎉')
    
    elif imprimir == "Imprimir 2 OS em uma folha":
        st.markdown("---")
        estampa1 = st.selectbox("Qual modelo de estampa 1?", nomes_estampas)

        try:
            imagem_estampa1 = Image.open(f"Estampas\{estampa1}.jpg")
            st.image(imagem_estampa1)

        except:
            try:
                imagem_estampa1 = Image.open(f"Estampas\{estampa1}.png")
                st.image(imagem_estampa1)

            except:
                try:
                    imagem_estampa1 = Image.open(f"Estampas\{estampa1}.webp")
                    st.image(imagem_estampa1)

                except:
                    try:
                        imagem_estampa1 = Image.open(f"Estampas\{estampa1}.jpeg")
                        st.image(imagem_estampa1)

                    except:
                        st.warning("modelo não encontrado para visualização")

        st.markdown("---")
        estampa2 = st.selectbox("Qual modelo de estampa 2?", nomes_estampas)

        try:
            imagem_estampa2 = Image.open(f"Estampas\{estampa2}.jpg")
            st.image(imagem_estampa2)

        except:
            try:
                imagem_estampa2 = Image.open(f"Estampas\{estampa2}.png")
                st.image(imagem_estampa2)

            except:
                try:
                    imagem_estampa2 = Image.open(f"Estampas\{estampa2}.webp")
                    st.image(imagem_estampa2)

                except:
                    try:
                        imagem_estampa2 = Image.open(f"Estampas\{estampa2}.jpeg")
                        st.image(imagem_estampa2)

                    except:
                        st.warning("modelo não encontrado para visualização")

        with st.form("AbrirOS"):
            
            data_entrega1 = st.date_input("Qual data da entrega prevista pedido 1?", value="today", min_value="today")
            data_entrega1 = data_entrega1.strftime("%d/%m/%Y")

            tamanho1 = st.radio("Qual tamanho do modelo 1?", tamanhos)

            cliente1 = st.text_input("Qual o nome do cliente 1?", value="")

            equipes_filtradas = [e for e in equipes if e != "BORDADO"]
            equipe1 = st.radio("Qual equipe responsável pelo pedido 1?", equipes_filtradas)

            observacao1 = st.text_input("Observação modelo 1?", max_chars=200)

            st.markdown("---")

            data_entrega2 = st.date_input("Qual data da entrega prevista pedido 2?", value="today", min_value="today")
            data_entrega2 = data_entrega2.strftime("%d/%m/%Y")
            
            #st.markdown("---")
            #quantidade = st.number_input("Qual quantidade de camisas?", min_value = 1) #Voltar com quantidade?

            tamanho2 = st.radio("Qual tamanho do modelo 2?", tamanhos)
            
            cliente2 = st.text_input("Qual o nome do cliente 2?", value=cliente1)
            
            equipe2 = st.radio("Qual equipe responsável pelo pedido 2?", equipes_filtradas)

            # if estampa in estampas_bordado:
            #     bordado = True
            # else: 
            #     bordado = False
            
            observacao2 = st.text_input("Observação modelo 2?", max_chars=200)

            submitted = st.form_submit_button("Criar OS!")
                
            # Só executa quando o botão para clicado
            if submitted:

                # Pega todos os valores da primeira coluna
                valores = planilhaOS.col_values(1)

                # Descobre a primeira linha vazia
                linha_vazia = len(valores) + 1  # +1 porque col_values não conta a próxima linha vazia

                #cria nova linha
                nova_linha1 = [[str(codigo), str(data_carimbo), str(data_entrega1), str(estampa1),
                    #str(bordado), #str(quantidade),
                    str(tamanho1), str(cliente1), str(equipe1), str(observacao1)]]
                
                nova_linha2 = [[str(codigo+1), str(data_carimbo), str(data_entrega2), str(estampa2),
                    #str(bordado), #str(quantidade),
                    str(tamanho2), str(cliente2), str(equipe2), str(observacao2)]]

                planilhaOS.update(f"A{linha_vazia}:K{linha_vazia}", nova_linha1)
                planilhaOS.update(f"A{linha_vazia+1}:K{linha_vazia+1}", nova_linha2)

                if imprimir:
                    pdf = FPDF("portrait", "mm", "A4")  # Alterado para portrait para melhor aproveitamento do espaço
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)

                    # PRIMEIRA OS (PARTE SUPERIOR)
                    pdf.cell(130, 8, "ORDEM DE SERVIÇO", ln=True, align=Align.C)
                    pdf.set_font("Arial", style="B", size=14)
                    pdf.cell(130, 8, f"Cliente: {cliente1}", ln=True, border=1)

                    pdf.set_font("Arial", style="", size=12)
                    pdf.cell(130, 7, f"Equipe: {equipe1}", ln=True, border=True)
                    pdf.cell(130, 7, f"Estampa: {estampa1}", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Tamanho: {tamanho1}", ln=True, border=True)
                    pdf.cell(130, 7, f"Quantidade: ", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Data Entrega: {data_entrega1}",ln=True, border=True)
                    #pdf.cell(95, 8, f"Código: {codigo}", ln=True, border=True)

                    pdf.multi_cell(130, 7, f"Observação: {observacao1}", border=1, ln=True, align=Align.L)
                    
                    # Área para imagem da primeira OS
                    x_imagem1 = 140
                    y_imagem1 = 18
                    largura_imagem = 50
                    altura_imagem = 75

                    pdf.rect(x_imagem1, y_imagem1, largura_imagem, altura_imagem)

                    try:
                        pdf.image(f"Estampas\\{estampa1}.jpg", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                    except:
                        try:
                            pdf.image(f"Estampas\\{estampa1}.png", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                        except:
                            try:
                                pdf.image(f"Estampas\\{estampa1}.webp", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                            except:
                                try:
                                    pdf.image(f"Estampas\\{estampa1}.jpeg", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                                except:
                                    pass

                    # Observação da primeira OS
                    pdf.set_xy(10, 95)
                    

                    # LINHA DIVISÓRIA (GUIA DE CORTE) - ALTERAÇÃO PRINCIPAL
                    pdf.set_draw_color(0, 0, 0)  # Cor preta
                    pdf.set_line_width(0.5)  # Espessura da linha
                    pdf.line(10, 148, 200, 148)  # Linha horizontal no meio da página
                    
                    # Texto indicativo de corte
                    pdf.set_xy(85, 145)
                    pdf.set_font("Arial", style="I", size=8)
                    pdf.cell(40, 5, "--- LINHA DE CORTE ---", align=Align.C)

                    # SEGUNDA OS (PARTE INFERIOR) - ALTERAÇÃO PRINCIPAL
                    pdf.set_font("Arial", size=12)
                    pdf.set_xy(10, 155)
                    pdf.cell(130, 8, "ORDEM DE SERVIÇO", ln=True, align=Align.C)
                    
                    pdf.set_font("Arial", style="B", size=14)
                    pdf.cell(130, 8, f"Cliente: {cliente2}", ln=True, border=1)

                    pdf.set_font("Arial", style="", size=12)
                    pdf.cell(130, 7, f"Equipe: {equipe2}", ln=True, border=True)
                    pdf.cell(130, 7, f"Estampa: {estampa2}", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Tamanho: {tamanho2}", ln=True, border=True)
                    pdf.cell(130, 7, f"Quantidade: ", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Data Entrega: {data_entrega2}", ln=True, border=True)
                    #pdf.cell(95, 8, f"Código: {codigo}", ln=True, border=True)

                    pdf.multi_cell(130, 8, f"Observação: {observacao2}", border=1, ln=True, align=Align.L)

                    # Área para imagem da segunda OS
                    x_imagem2 = 140
                    y_imagem2 = 173
                    pdf.rect(x_imagem2, y_imagem2, largura_imagem, altura_imagem)

                    try:
                        pdf.image(f"Estampas\\{estampa2}.jpg", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                    except:
                        try:
                            pdf.image(f"Estampas\\{estampa2}.png", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                        except:
                            try:
                                pdf.image(f"Estampas\\{estampa2}.webp", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                            except:
                                try:
                                    pdf.image(f"Estampas\\{estampa2}.jpeg", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                                except:
                                    pass

                    # Observação da segunda OS
                    pdf.set_xy(10, 250)
                    

                    pasta_base = os.path.dirname(__file__)
                    caminho_pasta = os.path.join(pasta_base, "OS")

                    if cliente1 == cliente2:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente2}.pdf"))
                    else:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente1}_{cliente2}.pdf"))

                st.toast('Ordem de serviço criada!', icon='🎉')

    elif imprimir ==  "Imprimir 3 OS em uma folha":
        st.markdown("---")
        estampa1 = st.selectbox("Qual modelo de estampa 1?", nomes_estampas)

        try:
            imagem_estampa1 = Image.open(f"Estampas\{estampa1}.jpg")
            st.image(imagem_estampa1)

        except:
            try:
                imagem_estampa1 = Image.open(f"Estampas\{estampa1}.png")
                st.image(imagem_estampa1)

            except:
                try:
                    imagem_estampa1 = Image.open(f"Estampas\{estampa1}.webp")
                    st.image(imagem_estampa1)

                except:
                    try:
                        imagem_estampa1 = Image.open(f"Estampas\{estampa1}.jpeg")
                        st.image(imagem_estampa1)

                    except:
                        st.warning("modelo não encontrado para visualização")

        st.markdown("---")
        estampa2 = st.selectbox("Qual modelo de estampa 2?", nomes_estampas)

        try:
            imagem_estampa2 = Image.open(f"Estampas\{estampa2}.jpg")
            st.image(imagem_estampa2)

        except:
            try:
                imagem_estampa2 = Image.open(f"Estampas\{estampa2}.png")
                st.image(imagem_estampa2)

            except:
                try:
                    imagem_estampa2 = Image.open(f"Estampas\{estampa2}.webp")
                    st.image(imagem_estampa2)

                except:
                    try:
                        imagem_estampa2 = Image.open(f"Estampas\{estampa2}.jpeg")
                        st.image(imagem_estampa2)

                    except:
                        st.warning("modelo não encontrado para visualização")

        st.markdown("---")
        estampa3 = st.selectbox("Qual modelo de estampa 3?", nomes_estampas)

        try:
            imagem_estampa3 = Image.open(f"Estampas\{estampa3}.jpg")
            st.image(imagem_estampa3)

        except:
            try:
                imagem_estampa3 = Image.open(f"Estampas\{estampa3}.png")
                st.image(imagem_estampa3)

            except:
                try:
                    imagem_estampa3 = Image.open(f"Estampas\{estampa3}.webp")
                    st.image(imagem_estampa3)

                except:
                    try:
                        imagem_estampa3 = Image.open(f"Estampas\{estampa3}.jpeg")
                        st.image(imagem_estampa3)

                    except:
                        st.warning("modelo não encontrado para visualização")

        with st.form("AbrirOS"):
            
            data_entrega1 = st.date_input("Qual data da entrega prevista pedido 1?", value="today", min_value="today")
            data_entrega1 = data_entrega1.strftime("%d/%m/%Y")

            tamanho1 = st.radio("Qual tamanho do modelo 1?", tamanhos)

            cliente1 = st.text_input("Qual o nome do cliente 1?", value="")

            equipes_filtradas = [e for e in equipes if e != "BORDADO"]
            equipe1 = st.radio("Qual equipe responsável pelo pedido 1?", equipes_filtradas)

            observacao1 = st.text_input("Observação modelo 1?", max_chars=200)

            st.markdown("---")

            data_entrega2 = st.date_input("Qual data da entrega prevista pedido 2?", value="today", min_value="today")
            data_entrega2 = data_entrega2.strftime("%d/%m/%Y")
            
            #st.markdown("---")
            #quantidade = st.number_input("Qual quantidade de camisas?", min_value = 1) #Voltar com quantidade?

            tamanho2 = st.radio("Qual tamanho do modelo 2?", tamanhos)
            
            cliente2 = st.text_input("Qual o nome do cliente 2?", value=cliente1)
            
            equipe2 = st.radio("Qual equipe responsável pelo pedido 2?", equipes_filtradas)

            # if estampa in estampas_bordado:
            #     bordado = True
            # else: 
            #     bordado = False
            
            observacao2 = st.text_input("Observação modelo 2?", max_chars=200)

            st.markdown("---")

            data_entrega3 = st.date_input("Qual data da entrega prevista pedido 3?", value="today", min_value="today")
            data_entrega3 = data_entrega3.strftime("%d/%m/%Y")
            
            #st.markdown("---")
            #quantidade = st.number_input("Qual quantidade de camisas?", min_value = 1) #Voltar com quantidade?

            tamanho3 = st.radio("Qual tamanho do modelo 3?", tamanhos)
            
            cliente3 = st.text_input("Qual o nome do cliente 3?", value=cliente1)
            
            equipe3 = st.radio("Qual equipe responsável pelo pedido 3?", equipes_filtradas)

            # if estampa in estampas_bordado:
            #     bordado = True
            # else: 
            #     bordado = False
            
            observacao3 = st.text_input("Observação modelo 3?", max_chars=200)

            submitted = st.form_submit_button("Criar OS!")
                
            # Só executa quando o botão para clicado
            if submitted:

                # Pega todos os valores da primeira coluna
                valores = planilhaOS.col_values(1)

                # Descobre a primeira linha vazia
                linha_vazia = len(valores) + 1  # +1 porque col_values não conta a próxima linha vazia

                #cria nova linha
                nova_linha1 = [[str(codigo), str(data_carimbo), str(data_entrega1), str(estampa1),
                    #str(bordado), #str(quantidade),
                    str(tamanho1), str(cliente1), str(equipe1), str(observacao1)]]
                
                nova_linha2 = [[str(codigo+1), str(data_carimbo), str(data_entrega2), str(estampa2),
                    #str(bordado), #str(quantidade),
                    str(tamanho2), str(cliente2), str(equipe2), str(observacao2)]]
                
                nova_linha3 = [[str(codigo+2), str(data_carimbo), str(data_entrega3), str(estampa3),
                    #str(bordado), #str(quantidade),
                    str(tamanho3), str(cliente3), str(equipe3), str(observacao3)]]

                planilhaOS.update(f"A{linha_vazia}:K{linha_vazia}", nova_linha1)
                planilhaOS.update(f"A{linha_vazia+1}:K{linha_vazia+1}", nova_linha2)
                planilhaOS.update(f"A{linha_vazia+2}:K{linha_vazia+2}", nova_linha3)

                if imprimir:
                    pdf = FPDF("portrait", "mm", "A4")  # Alterado para portrait para melhor aproveitamento do espaço
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)

                    # PRIMEIRA OS (PARTE SUPERIOR)
                    pdf.cell(130, 8, "ORDEM DE SERVIÇO", ln=True, align=Align.C)
                    pdf.set_font("Arial", style="B", size=14)
                    pdf.cell(130, 8, f"Cliente: {cliente1}", ln=True, border=1)

                    pdf.set_font("Arial", style="", size=12)
                    pdf.cell(130, 7, f"Equipe: {equipe1}", ln=True, border=True)
                    pdf.cell(130, 7, f"Estampa: {estampa1}", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Tamanho: {tamanho1}", ln=True, border=True)
                    pdf.cell(130, 7, f"Quantidade: ", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Data Entrega: {data_entrega1}",ln=True, border=True)
                    #pdf.cell(95, 8, f"Código: {codigo}", ln=True, border=True)

                    pdf.multi_cell(130, 7, f"Observação: {observacao1}", border=1, ln=True, align=Align.L)
                    
                    # Área para imagem da primeira OS
                    x_imagem1 = 140
                    y_imagem1 = 18
                    largura_imagem = 50
                    altura_imagem = 75

                    pdf.rect(x_imagem1, y_imagem1, largura_imagem, altura_imagem)

                    try:
                        pdf.image(f"Estampas\\{estampa1}.jpg", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                    except:
                        try:
                            pdf.image(f"Estampas\\{estampa1}.png", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                        except:
                            try:
                                pdf.image(f"Estampas\\{estampa1}.webp", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                            except:
                                try:
                                    pdf.image(f"Estampas\\{estampa1}.jpeg", x=x_imagem1+2, y=y_imagem1+2, w=largura_imagem-4)
                                except:
                                    pass                   

                    # LINHA DIVISÓRIA (GUIA DE CORTE) - ALTERAÇÃO PRINCIPAL
                    pdf.set_draw_color(0, 0, 0)  # Cor preta
                    pdf.set_line_width(0.5)  # Espessura da linha
                    pdf.line(10, 98, 200, 98)  # Linha horizontal no meio da página
                    
                    # Texto indicativo de corte
                    pdf.set_xy(85, 95)
                    pdf.set_font("Arial", style="I", size=8)
                    pdf.cell(40, 5, "--- LINHA DE CORTE ---", align=Align.C)

                    # SEGUNDA OS
                    pdf.set_font("Arial", size=12)
                    pdf.set_xy(10, 105)
                    pdf.cell(130, 8, "ORDEM DE SERVIÇO", ln=True, align=Align.C)
                    
                    pdf.set_font("Arial", style="B", size=14)
                    pdf.cell(130, 8, f"Cliente: {cliente2}", ln=True, border=1)

                    pdf.set_font("Arial", style="", size=12)
                    pdf.cell(130, 7, f"Equipe: {equipe2}", ln=True, border=True)
                    pdf.cell(130, 7, f"Estampa: {estampa2}", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Tamanho: {tamanho2}", ln=True, border=True)
                    pdf.cell(130, 7, f"Quantidade: ", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Data Entrega: {data_entrega2}", ln=True, border=True)
                    #pdf.cell(95, 8, f"Código: {codigo}", ln=True, border=True)

                    pdf.multi_cell(130, 8, f"Observação: {observacao2}", border=1, ln=True, align=Align.L)

                    # Área para imagem da segunda OS
                    x_imagem2 = 140
                    y_imagem2 = 103
                    pdf.rect(x_imagem2, y_imagem2, largura_imagem, altura_imagem)

                    try:
                        pdf.image(f"Estampas\\{estampa2}.jpg", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                    except:
                        try:
                            pdf.image(f"Estampas\\{estampa2}.png", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                        except:
                            try:
                                pdf.image(f"Estampas\\{estampa2}.webp", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                            except:
                                try:
                                    pdf.image(f"Estampas\\{estampa2}.jpeg", x=x_imagem2+2, y=y_imagem2+2, w=largura_imagem-4)
                                except:
                                    pass

                    # LINHA DIVISÓRIA (GUIA DE CORTE) - ALTERAÇÃO PRINCIPAL
                    pdf.set_draw_color(0, 0, 0)  # Cor preta
                    pdf.set_line_width(0.5)  # Espessura da linha
                    pdf.line(10, 188, 200, 188)  # Linha horizontal no meio da página
                    
                    # Texto indicativo de corte
                    pdf.set_xy(85, 185)
                    pdf.set_font("Arial", style="I", size=8)
                    pdf.cell(40, 5, "--- LINHA DE CORTE ---", align=Align.C)

                    # TERCEIRA OS
                    pdf.set_font("Arial", size=12)
                    pdf.set_xy(10, 195)
                    pdf.cell(130, 8, "ORDEM DE SERVIÇO", ln=True, align=Align.C)
                    
                    pdf.set_font("Arial", style="B", size=14)
                    pdf.cell(130, 8, f"Cliente: {cliente3}", ln=True, border=1)

                    pdf.set_font("Arial", style="", size=12)
                    pdf.cell(130, 7, f"Equipe: {equipe3}", ln=True, border=True)
                    pdf.cell(130, 7, f"Estampa: {estampa3}", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Tamanho: {tamanho3}", ln=True, border=True)
                    pdf.cell(130, 7, f"Quantidade: ", ln=True, border=True)
                    
                    pdf.cell(130, 7, f"Data Entrega: {data_entrega3}", ln=True, border=True)
                    #pdf.cell(95, 8, f"Código: {codigo}", ln=True, border=True)

                    pdf.multi_cell(130, 8, f"Observação: {observacao3}", border=1, ln=True, align=Align.L)

                    # Área para imagem da segunda OS
                    x_imagem3 = 140
                    y_imagem3 = 203
                    pdf.rect(x_imagem3, y_imagem3, largura_imagem, altura_imagem)

                    try:
                        pdf.image(f"Estampas\\{estampa3}.jpg", x=x_imagem3+2, y=y_imagem3+2, w=largura_imagem-4)
                    except:
                        try:
                            pdf.image(f"Estampas\\{estampa3}.png", x=x_imagem3+2, y=y_imagem3+2, w=largura_imagem-4)
                        except:
                            try:
                                pdf.image(f"Estampas\\{estampa3}.webp", x=x_imagem3+2, y=y_imagem3+2, w=largura_imagem-4)
                            except:
                                try:
                                    pdf.image(f"Estampas\\{estampa3}.jpeg", x=x_imagem3+2, y=y_imagem3+2, w=largura_imagem-4)
                                except:
                                    pass
                    

                    pasta_base = os.path.dirname(__file__)
                    caminho_pasta = os.path.join(pasta_base, "OS")

                    if cliente1 == cliente2:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente2}.pdf"))
                    elif cliente1 == cliente2 and cliente1 != cliente3:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente1}_{cliente2}.pdf"))
                    elif cliente3 == cliente2 and cliente1 != cliente3:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente3}_{cliente2}.pdf"))
                    elif cliente1 == cliente3 and cliente2 != cliente3:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente1}_{cliente3}.pdf"))
                    else:
                        pdf.output(os.path.join(caminho_pasta, f"OS_{codigo}_{cliente1}_{cliente2}_{cliente3}.pdf"))

                st.toast('Ordem de serviço criada!', icon='🎉')


# Gerar dados fictícios de produção para o último mês
def gerar_dados_producao(periodo):
    if st.session_state['rotation'] == "a":
    
        dados_periodo = dados_Producao_completo.copy()
        dados_periodo['DATA_DT'] = pd.to_datetime(dados_periodo['DATA'], format='%d/%m/%Y')
        dados_periodo['MES'] = dados_periodo['DATA_DT'].dt.month
        dados_periodo['ANO'] = dados_periodo['DATA_DT'].dt.year
        dados_periodo['MES_ANO'] = dados_periodo['MES'].astype(str) + '/' + dados_periodo['ANO'].astype(str)
        dados_filtrados = dados_periodo[dados_periodo['MES_ANO'] == periodo]
        
        # Dados agrupados
        producao_por_setor = dados_filtrados.groupby('SETOR')['TOTAL'].sum().reset_index()
        producao_diaria_setor = dados_filtrados.groupby(['DATA_DT', 'SETOR'])['TOTAL'].sum().reset_index().sort_values('DATA_DT')

        st.subheader(f"Produção por Setor - Mês {periodo}") 

        fig_pizza = px.pie(
            producao_por_setor,
            names='SETOR',
            values='TOTAL',
            title=f"Distribuição da Produção - Mês {periodo}",
            hole=0.4
        )
        st.plotly_chart(fig_pizza, use_container_width=True)

        i = 0

        for setor in dados_filtrados['SETOR'].unique():
            dados_setor = producao_diaria_setor[producao_diaria_setor['SETOR'] == setor]

            fig_linha = px.line(
                dados_setor,
                x='DATA_DT',
                y='TOTAL',
                title=f"Produção - {setor}",
                markers = True,
                labels={"TOTAL": "Unidades Produzidas", "DATA_DT": "Data"},
                color_discrete_sequence=[px.colors.qualitative.G10[i % len(px.colors.qualitative.G10)]]
            )
            fig_linha.update_layout(xaxis_tickformat='%d/%m')
            st.plotly_chart(fig_linha, use_container_width=True)
            i+=1

        # Opção 3: Gráfico de pizza por subsetor dentro de cada setor
        st.subheader(f"Distribuição por Subsetor - Mês {periodo}")

        for setor in dados_filtrados['SETOR'].unique():
            dados_setor = dados_filtrados[dados_filtrados['SETOR'] == setor]
            producao_subsetor = dados_setor.groupby('SUBSETOR')['TOTAL'].sum().reset_index()
            
            fig_pizza = px.bar(
                producao_subsetor,
                y='TOTAL',
                x='SUBSETOR',
                title=f"Distribuição por Subsetor - {setor}",
                color="SUBSETOR",
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            st.plotly_chart(fig_pizza, use_container_width=True)
    
    else:
        dados_periodo = dados_Producao_completo.copy()
        dados_periodo['DATA_DT'] = pd.to_datetime(dados_periodo['DATA'], format='%d/%m/%Y')
        dados_periodo['MES'] = dados_periodo['DATA_DT'].dt.month
        dados_periodo['ANO'] = dados_periodo['DATA_DT'].dt.year
        dados_periodo['MES_ANO'] = dados_periodo['MES'].astype(str) + '/' + dados_periodo['ANO'].astype(str)
        dados_filtrados = dados_periodo[dados_periodo['MES_ANO'] == periodo]
        
        # Dados agrupados
        producao_por_setor = dados_filtrados.groupby('SETOR')['TOTAL'].sum().reset_index()
        producao_diaria_setor = dados_filtrados.groupby(['DATA_DT', 'SETOR'])['TOTAL'].sum().reset_index().sort_values('DATA_DT')

        @st.fragment
        def rotacion():
            delay = 5
            repeticao = 10
            texto_placeholder1 = st.empty()
            texto_placeholder2 = st.empty()

            # Atualiza o display
            with texto_placeholder1.container():
                fig_pizza = px.pie(
                    producao_por_setor,
                    names='SETOR',
                    values='TOTAL',
                    title=f"Distribuição da Produção - Mês {periodo}",
                    hole=0.4,
                    height=320  # Altura compacta
                )
                
                fig_pizza.update_layout(
                    margin=dict(l=10, r=10, t=30, b=10),
                    title_x=0.5,
                    showlegend=True,  # Controlar se mostra legenda
                    legend=dict(
                        orientation="h",  # Legenda horizontal ocupa menos espaço vertical
                        yanchor="bottom",
                        y=-0.1,
                        xanchor="center",
                        x=0.5
                    )
                )
                
                st.plotly_chart(fig_pizza, use_container_width=True)

            for k in range(repeticao):
                i = 0
                for setor in dados_filtrados['SETOR'].unique():
    
                    with texto_placeholder2.container():
                        dados_setor = producao_diaria_setor[producao_diaria_setor['SETOR'] == setor]

                        fig_linha = px.line(
                            dados_setor,
                            x='DATA_DT',
                            y='TOTAL',
                            title=f"Produção - {setor}",
                            markers = True,
                            labels={"TOTAL": "Unidades Produzidas", "DATA_DT": "Data"},
                            color_discrete_sequence=[px.colors.qualitative.G10[i % len(px.colors.qualitative.G10)]],
                            height=530
                        )
                        fig_linha.update_layout(xaxis_tickformat='%d/%m')
                        st.plotly_chart(fig_linha, use_container_width=True, key=f"chart_{k}.{i}.a")
                        i += 1

                                    
                    for j in range(delay):
                
                        time.sleep(1)

                i = 0
                for setor in dados_filtrados['SETOR'].unique():
                    
                    with texto_placeholder2.container():
                        dados_setor = dados_filtrados[dados_filtrados['SETOR'] == setor]
                        producao_subsetor = dados_setor.groupby('SUBSETOR')['TOTAL'].sum().reset_index()
                        
                        fig_pizza = px.bar(
                            producao_subsetor,
                            y='TOTAL',
                            x='SUBSETOR',
                            title=f"Distribuição por Subsetor - {setor}",
                            color="SUBSETOR",
                            color_discrete_sequence=px.colors.qualitative.Bold,
                            height= 530
                        )
                        st.plotly_chart(fig_pizza, use_container_width=True, key=f"chart_{k}.{i}.b")
                        i += 1

                    for j in range(delay):
                
                        time.sleep(1)
            
            with texto_placeholder2.container():
                st.warning("Resete o código!")

        rotacion()
        

if pagina == "Home":
    ##Modo Apresentação##
    if st.sidebar.button("🖥️ Modo Apresentação"):
        st.session_state.modo_apresentacao = not st.session_state.get('modo_apresentacao', False)

    st.sidebar.write("- - -")

    # Inicializar estado da apresentação se não existir
    if 'modo_apresentacao' not in st.session_state:
        st.session_state.modo_apresentacao = False

    # Se o modo apresentação estiver ativo, ocultar outros elementos
    if st.session_state.modo_apresentacao:

        st.session_state['rotation'] = "b"
        dados_periodo = dados_Producao_completo.copy()
        dados_periodo['DATA_DT'] = pd.to_datetime(dados_periodo['DATA'], format='%d/%m/%Y')
        dados_periodo['MES'] = dados_periodo['DATA_DT'].dt.month
        dados_periodo['ANO'] = dados_periodo['DATA_DT'].dt.year
        meses_anos = dados_periodo.groupby(['MES', 'ANO']).size().reset_index()
        meses_anos['MES_ANO'] = meses_anos['MES'].astype(str) + '/' + meses_anos['ANO'].astype(str)
        meses_anos = meses_anos.sort_values(['ANO', 'MES'], ascending=[False, False])

        meses = meses_anos['MES_ANO'].unique().tolist()
        col1, col2, col3 = st.columns([4, 2, 1])

        with col1:
            periodo_selecionado = st.selectbox("Mês", meses)

        with col3:
            st.write("")


            if st.button("🚪 Sair da Apresentação", type="primary", use_container_width=True):
                st.session_state.modo_apresentacao = False
                st.rerun()
        
       

        # Ocultar sidebar e outros elementos
        st.markdown("""
            <style>
            .stSidebar {display: none;}
            header {display: none;}
            .stApp > div:first-child {padding-top: 0;}
            </style>
        """, unsafe_allow_html=True)

        # # Gerar dados de produção (reutilizando a função existente)
        gerar_dados_producao(periodo_selecionado)
                    
        # Não renderizar o resto da aplicação
        st.stop()    

    ##Modo normal##
    # CSS personalizado que se adapta ao tema
    st.markdown("""
        <style>
        /* Estilos que se adaptam ao tema */
        .company-header {
            font-size: 3.5rem;
            font-weight: 700;
            text-align: center;
            margin-bottom: 0.5rem;
            padding: 2rem;
            background-image: url('Estamparia2.png');
            background-size: cover;
            background-position: center;
            border-radius: 10px;
            color: white;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.7);
        }
        .company-tagline {
            font-size: 1.5rem;
            text-align: center;
            margin-bottom: 2rem;
        }
        .section-header {
            font-size: 2rem;
            border-bottom: 2px solid;
            padding-bottom: 0.5rem;
            margin: 2rem 0 1.5rem 0;
        }
        .service-card {
            border-radius: 10px;
            padding: 1.5rem;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            height: 100%;
            transition: transform 0.3s;
        }
        .service-card:hover {
            transform: translateY(-5px);
        }
        .contact-info {
            padding: 2rem;
            border-radius: 10px;
            margin-top: 2rem;
        }
        .footer {
            text-align: center;
            padding: 1.5rem;
            margin-top: 3rem;
            font-size: 0.9rem;
        }
        
        /* Ajustes específicos para tema claro */
        [data-theme="light"] {
            --text-color: #31333F;
            --bg-color: #f8f9fa;
            --card-bg: white;
            --primary-color: #1a73e8;
            --secondary-color: #5f6368;
            --gradient: linear-gradient(to right, #ffffff, #f1f3f5);
        }
        
        /* Ajustes específicos para tema escuro */
        [data-theme="dark"] {
            --text-color: #FFFFFF;
            --bg-color: #0E1117;
            --card-bg: #262730;
            --primary-color: #3eb0f7;
            --secondary-color: #AAAAAA;
            --gradient: linear-gradient(to right, #0E1117, #1a1a2e);
        }
        
        /* Aplicação das variáveis CSS */
        .main {
            background-color: var(--bg-color);
        }
        .stApp {
            background: var(--gradient);
        }
        .company-tagline {
            color: var(--secondary-color);
        }
        .section-header {
            color: var(--primary-color);
            border-bottom-color: var(--primary-color);
        }
        .service-card {
            background-color: var(--card-bg);
            color: var(--text-color);
        }
        .contact-info {
            background-color: var(--primary-color);
            color: white;
        }
        .footer {
            color: var(--secondary-color);
        }
        </style>
    """, unsafe_allow_html=True)

    # Header da empresa com imagem de fundo
    st.markdown('<h1 class="company-header">PLASTCOR ESTAMPARIA</h1>', unsafe_allow_html=True)
    st.markdown('<p class="company-tagline">Transformando ideias em estampas de alta qualidade</p>', unsafe_allow_html=True)

    # Imagem de destaque
    st.image("Estamparia2.png", use_container_width=True)

    # # Seção de serviços
    # st.markdown('<h2 class="section-header">Avisos</h2>', unsafe_allow_html=True)

    # col1, col2, col3 = st.columns(3)

    # with col1:
    #     st.markdown("""
    #         <div class="service-card">
    #             <h3>Festa de Final de Ano</h3>
    #             <p>🎉 Teremos uma festa de final de ano no dia 20/12/25, a partir das 11 da manhã.</p>
    #         </div>
    #     """, unsafe_allow_html=True)
    #     with st.expander("Saiba mais."):
    #         st.markdown(f"A comemoração será na chácara Pirimpi, https://maps.app.goo.gl/kwjY1C19LsVavca57. Tragam suas famílias")

    # with col2:
    #     st.markdown("""
    #         <div class="service-card">
    #             <h3>🎄 Feriado dia 25/12/25 e 01/01/26</h3>
    #             <p>Informamos que a empresa não funcionará nesses dias.</p>
    #         </div>
    #     """, unsafe_allow_html=True)

    # with col3:
    #     st.markdown("""
    #         <div class="service-card">
    #             <h3>Reunião de final de ano</h3>
    #             <p>Reunião com todos os funcionários dia 19/12/25.</p>
    #         </div>
    #     """, unsafe_allow_html=True)
    #     with st.expander("Saiba mais."):
    #         st.markdown(f"Reunião será feita no galpão 1, ás 16h, todos os funcionários estão convocados.")

    ##Plotando dados falsos ##
    st.markdown('<h2 class="section-header">DASHBOARD - Produção</h2>', unsafe_allow_html=True)

    st.session_state['rotation'] = "a"
    
    dados_periodo = dados_Producao_completo.copy()
    dados_periodo['DATA_DT'] = pd.to_datetime(dados_periodo['DATA'], format='%d/%m/%Y')
    dados_periodo['MES'] = dados_periodo['DATA_DT'].dt.month
    dados_periodo['ANO'] = dados_periodo['DATA_DT'].dt.year
    meses_anos = dados_periodo.groupby(['MES', 'ANO']).size().reset_index()
    meses_anos['MES_ANO'] = meses_anos['MES'].astype(str) + '/' + meses_anos['ANO'].astype(str)
    meses_anos = meses_anos.sort_values(['ANO', 'MES'], ascending=[False, False])

    meses = meses_anos['MES_ANO'].unique().tolist()
    
    periodo_selecionado = st.sidebar.selectbox("Mês", meses)

    gerar_dados_producao(periodo_selecionado)

    
if pagina ==  "Produção":

    st.sidebar.write("- - -")

    qp = st.sidebar.radio("Escolha:", ["Lançar Produção", "Editar Informações", "Produção Individual"])

    if qp == "Lançar Produção":

        setorProducao = st.selectbox("Informe o setor do lançamento", dados_Quadro_completo["SETOR"].unique())
        horaextra = st.radio("Hora extra?", ["Não", "Sim"])

        with st.form("Lançamento Produção"):
            proddata = st.date_input("Data do lançamento", format="DD/MM/YYYY")

            subsetores_filtrados = dados_Quadro_completo[dados_Quadro_completo["SETOR"] == setorProducao]["SUBSETOR"].unique()

            subsetorProducao = st.selectbox(
                "Informe o subsetor do lançamento:", 
                subsetores_filtrados
            )
            
            producaoSetor = st.number_input("Qual a quantidade produzida na Jornada de Trabalho Regular?",step=1)

            if horaextra == "Não":
                producaoHESetor = 0
                total = producaoSetor
            elif horaextra == "Sim":
                producaoHESetor = st.number_input("Qual a quantidade produzida em hora extra?",step=1)
                total = producaoSetor + producaoHESetor

            observacao = st.text_input("Observações:", max_chars=50)

            proddata = proddata.strftime("%d/%m/%Y")  # Formato DD/MM/YYYY

            valores_linha = [
                proddata,
                setorProducao,
                subsetorProducao,
                producaoSetor,
                horaextra,
                producaoHESetor,
                total,
                observacao,
            ]

            submitted = st.form_submit_button("Submeter")

            if submitted:

                st.toast(f"{valores_linha}")

                if dados_Producao_completo is not None and len(dados_Producao_completo) > 0:
                    ultima_linha = dados_Producao_completo.index[-1] + 3
                else:
                    ultima_linha = 2  # primeira linha de dados
            
                planilhaProducao.batch_update([{
                            'range': f'A{ultima_linha}:H{ultima_linha}',
                            'values': [valores_linha]
                        }])
                
                st.toast("Submissão feita com sucesso", icon=":material/thumb_up:")

    elif qp == "Editar Informações":

        dataProd = st.selectbox("Informe a data lançada:", dados_Producao_completo["DATA"].unique())
        setorProd = st.selectbox("Selecione o setor:", dados_Producao_completo[dados_Producao_completo["DATA"] == dataProd]["SETOR"].unique())
        subsetores_filtrados = dados_Quadro_completo[dados_Quadro_completo["SETOR"] == setorProd]["SUBSETOR"].unique()

        subsetorProducao = st.selectbox(
            "Informe o subsetor do lançamento:", 
            subsetores_filtrados
        )
        
        indiceEdicao = dados_Producao_completo.index[(dados_Producao_completo['DATA'] == dataProd) & (dados_Producao_completo['SUBSETOR'] == subsetorProducao)].tolist()
        linhaEdicao = indiceEdicao[0] + 2

        with st.form("Editar Info"):            

            data_datetime = datetime.strptime(dataProd, "%d/%m/%Y")
            editdata = st.date_input("Alteração na data lançada?", format="DD/MM/YYYY", value= data_datetime)
            editdata = editdata.strftime("%d/%m/%Y")  # Formato DD/MM/YYYY
            
            producaoAnterior = dados_Producao_completo.iloc[indiceEdicao[0]]['PRODUCAO']
            producaoSetor = st.number_input("Alteração na quantidade produzida?",step=1,value=producaoAnterior)
            horaextra = st.radio("Hora extra?", ["Não", "Sim"])

            producaoHEAnterior = dados_Producao_completo.iloc[indiceEdicao[0]]['PRODUCAO HORA EXTRA']
            producaoHESetor = st.number_input("Alteração na quantidade produzida em hora extra?",step=1,value=producaoHEAnterior)
            total = producaoSetor + producaoHESetor
            observacaoanterior = dados_Producao_completo.iloc[indiceEdicao[0]]['OBSERVACOES']
            observacao = st.text_input("Alteração nas observações?", max_chars=50, value=observacaoanterior)

            valores_linha = [
                editdata,
                setorProd,
                subsetorProducao,
                producaoSetor,
                horaextra,
                producaoHESetor,
                total,
                observacao
            ]

            submitted = st.form_submit_button("Submeter")

            if submitted:

                st.toast(f"{valores_linha}")
            
                planilhaProducao.batch_update([{
                            'range': f'A{linhaEdicao}:H{linhaEdicao}',
                            'values': [valores_linha]
                        }])
                
                st.toast("Submissão feita com sucesso", icon=":material/thumb_up:")
    
    if qp == "Produção Individual":
        setorInd = st.selectbox("Informe o setor do funcionário:", dados_Quadro_completo["SETOR"].unique())
        subsetorInd = st.selectbox("Informe o subsetor do funcionário:", dados_Quadro_completo[dados_Quadro_completo["SETOR"] == setorInd]["SUBSETOR"].unique())
        dados_periodo = dados_Producao_completo.copy()
        dados_periodo['DATA_DT'] = pd.to_datetime(dados_periodo['DATA'], format='%d/%m/%Y')
        dados_periodo['MES'] = dados_periodo['DATA_DT'].dt.month
        dados_periodo['ANO'] = dados_periodo['DATA_DT'].dt.year
        meses_anos = dados_periodo.groupby(['MES', 'ANO']).size().reset_index()
        meses_anos['MES_ANO'] = meses_anos['MES'].astype(str) + '/' + meses_anos['ANO'].astype(str)
        dados_periodo['MES_ANO'] = dados_periodo['MES'].astype(str) + '/' + dados_periodo['ANO'].astype(str)
        meses_anos = meses_anos.sort_values(['ANO', 'MES'], ascending=[False, False])

        meses = meses_anos['MES_ANO'].unique().tolist()

        nomeInd = st.selectbox("Selecione o funcionário:", dados_Quadro_completo[dados_Quadro_completo["SUBSETOR"] == subsetorInd]["NOME"].unique())
        
        periodo_selecionado = st.selectbox("Mês", meses)

        if nomeInd in dados_Falta_completo["NOME"].values:
            
            filtro = dados_periodo[dados_periodo["MES_ANO"] == periodo_selecionado]
            filtro = filtro[filtro["SUBSETOR"] == subsetorInd]
            
            filtrofalta = filtrofalta = dados_Falta_completo[
                (dados_Falta_completo["NOME"] == nomeInd) & 
                (dados_Falta_completo["ABONIR?"] == "Não")
            ]
            datas_falta = filtrofalta['DATA'].unique()
            produção_valida = filtro[~filtro['DATA'].isin(datas_falta)]

            somaInd = produção_valida['PRODUCAO'].sum()

            somaequipe = filtro["PRODUCAO"].sum()
            
            st.write("- - -")
            st.header(f"Produção {nomeInd}:  {somaInd}")
            st.subheader(f"A produção da equipe: {somaequipe}.")

        else:
            filtro = dados_periodo[dados_periodo["MES_ANO"] == periodo_selecionado]
            filtro = filtro[filtro["SUBSETOR"] == subsetorInd]

            somaequipe = filtro["PRODUCAO"].sum()

            
            st.write("- - -")
            st.header(f"Produção {nomeInd}:  {somaequipe}")
            st.subheader(f"A produção da equipe: {somaequipe}.")


if pagina ==  "Estampas":

    st.sidebar.write("- - -")

    #Select bar para mostrar estampa
    visual = st.sidebar.selectbox(
        "Escolha um modelo de estampas", 
        arquivos_imagens)

    st.title("Visualizador de Estampas") #Título da aplicação
    imagem = Image.open(f"Estampas\{visual}")
    st.image(imagem, caption=visual)

if pagina == "Fechar Ordem de Serviço":
    codigo_procurado = st.number_input("Qual ordem de serviço deseja fechar?", min_value=None, step=1)
    if not dados_os_completo.empty:

        indice = dados_os_completo.index[dados_os_completo['Código OS'] == codigo_procurado].tolist()
        
        if st.button("Fechar Ordem!"):
            if indice:
                primeira_linha = indice[0]+2  # pega o primeiro índice encontrado
                valor_atual = planilhaOS.cell(primeira_linha, 4).value
                if valor_atual == "Aberto":
                    planilhaOS.update_cell(primeira_linha, 4, "Fechado")
                    st.toast(f"Código '{codigo_procurado}' fechado com sucesso!")
                    st.write("Dê reload na página para visualização")
                else:
                    st.warning(f"Código '{codigo_procurado}' não está aberto e não foi atualizado.")
                
                
            else:
                st.warning(f"Código '{codigo_procurado}' não foi encontrado.")

        st.header("Ordens de serviço Abertas")
        
        dados_os_abertos = dados_os_completo.loc[dados_os_completo['Status'] == 'Aberto', :] #filtro de "aberto"
        mostrar_planilha(dados_os_abertos)

    else:
        st.write("Não há ordens de serviço registradas")

#Visualização dos dados
if pagina == "Ver ordens de Serviço":
    st.header("Ordens de serviço")
    mostrar_planilha(dados_os_completo)

if pagina == "Ordem de Serviço":
    create()

if pagina == "Quadro de Funcionários":
    st.sidebar.write("- - -")

    qf = st.sidebar.radio("Escolha:", ["Visualizar Quadro", "Adicionar novo funcionário", "Editar Informações"])
    if qf == "Visualizar Quadro":
        st.subheader("Quadro de funcionários")
        st.dataframe(dados_Quadro_completo)

    elif qf == "Adicionar novo funcionário":
        if dados_Quadro_completo is not None and len(dados_Quadro_completo) > 0:
            ultima_linha = dados_Quadro_completo.index[-1] + 3
        else:
            ultima_linha = 2  # primeira linha de dados

        with st.form("Novo funcionário"):

            nomeFuncionario = st.text_input("Nome do novo funcionário")
            setorFuncionario = st.selectbox("Informe o setor desse funcionário", dados_Quadro_completo["SETOR"].unique(), accept_new_options=True)
            subsetorFuncionario = st.selectbox("Informe o subsetor desse funcionário:", dados_Quadro_completo["SUBSETOR"].unique(), accept_new_options=True)
            cargoFuncionario = st.selectbox("Informe o cargo desse funcionário", dados_Quadro_completo["CARGO"].unique(), accept_new_options=True)
            statusFuncionario = "Verdadeiro"

            valores_linha = [
                nomeFuncionario,
                setorFuncionario,
                subsetorFuncionario,
                cargoFuncionario,
                statusFuncionario,
            ]

            submitted = st.form_submit_button("Submeter")

            if submitted:
                st.toast(f"{valores_linha}")
            
                planilhaQuadro.batch_update([{
                            'range': f'A{ultima_linha}:E{ultima_linha}',
                            'values': [valores_linha]
                        }])
                
                st.toast("Submissão feita com sucesso", icon=":material/thumb_up:")
                
    elif qf == "Editar Informações":
        

        setorInfo = st.selectbox("Informe o setor desse funcionário:", dados_Quadro_completo["SETOR"].unique())
        nomeInfo = st.selectbox("Selecione o nome:", dados_Quadro_completo[dados_Quadro_completo["SETOR"] == setorInfo])

        indiceEdicao = dados_Quadro_completo.index[dados_Quadro_completo['NOME'] == nomeInfo].tolist()
        linhaEdicao = indiceEdicao[0] + 2

        with st.form("Edit funcionário"):

            setorFuncionario = st.selectbox("Informe o setor desse funcionário:", dados_Quadro_completo["SETOR"].unique(), accept_new_options=True)
            subsetorFuncionario = st.selectbox("Informe o subsetor desse funcionário:", dados_Quadro_completo["SUBSETOR"].unique(), accept_new_options=True)
            cargoFuncionario = st.selectbox("Informe o cargo desse funcionário:", dados_Quadro_completo["CARGO"].unique(), accept_new_options=True)
            statusFuncionario = st.selectbox("O funcionário continua Ativo?", ['Verdadeiro', 'Falso'])

            valores_linha = [
                nomeInfo,
                setorFuncionario,
                subsetorFuncionario,
                cargoFuncionario,
                statusFuncionario,
            ]

            submitted = st.form_submit_button("Submeter")

            if submitted:

                st.toast(f"{valores_linha}")
            
                planilhaQuadro.batch_update([{
                            'range': f'A{linhaEdicao}:E{linhaEdicao}',
                            'values': [valores_linha]
                        }])
                
                st.toast("Submissão feita com sucesso", icon=":material/thumb_up:")

if pagina == "Falta":

    st.sidebar.write("- - -")
    
    st.subheader("Lançamento de Falta.")
    qf = st.sidebar.radio("Escolha:", ["Lançar Falta", "Editar Informações"])

    if qf == "Lançar Falta":
        setorInfo = st.selectbox("Informe o setor do funcionário:", dados_Quadro_completo["SETOR"].unique())

        with st.form("Lançamento Falta"):
            faltadata = st.date_input("Dia da falta", format="DD/MM/YYYY")
            nomeInfo = st.selectbox("Selecione o nome:", dados_Quadro_completo[dados_Quadro_completo["SETOR"] == setorInfo])
            turno = st.selectbox("Informe o turno:", ["Matutino", "Vespertino", "Dia inteiro"])
            atestado = st.selectbox("Apresentou justificaiva (abonar falta)?", ["Sim", "Não"])
            observacao = st.text_input("Observações:", max_chars=50)

            faltadata = faltadata.strftime("%d/%m/%Y")  # Formato DD/MM/YYYY

            valores_linha = [
                faltadata,
                nomeInfo,
                turno,
                atestado,
                observacao,
            ]

            submitted = st.form_submit_button("Submeter")

            if submitted:

                st.toast(f"{valores_linha}")

                if dados_Falta_completo is not None and len(dados_Falta_completo) > 0:
                    ultima_linha = dados_Falta_completo.index[-1] + 3
                else:
                    ultima_linha = 2  # primeira linha de dados
            
                planilhaFalta.batch_update([{
                            'range': f'A{ultima_linha}:E{ultima_linha}',
                            'values': [valores_linha]
                        }])
                
                st.toast("Submissão feita com sucesso", icon=":material/thumb_up:")

    elif qf == "Editar Informações":

        dados_periodo = dados_Falta_completo.copy()
        dados_periodo['DATA_DT'] = pd.to_datetime(dados_periodo['DATA'], format='%d/%m/%Y')
        dados_periodo['MES'] = dados_periodo['DATA_DT'].dt.month
        dados_periodo['ANO'] = dados_periodo['DATA_DT'].dt.year
        meses_anos = dados_periodo.groupby(['MES', 'ANO']).size().reset_index()
        meses_anos['MES_ANO'] = meses_anos['MES'].astype(str) + '/' + meses_anos['ANO'].astype(str)
        dados_periodo['MES_ANO'] = dados_periodo['MES'].astype(str) + '/' + dados_periodo['ANO'].astype(str)
        meses_anos = meses_anos.sort_values(['ANO', 'MES'], ascending=[False, False])

        meses = meses_anos['MES_ANO'].unique().tolist()
        
        periodo_selecionado = st.selectbox("Informe o mês da ausência registrada desse funcionário:", meses)
                
        dataInfo = st.selectbox("Qual a data que deseja editar?", dados_periodo[dados_periodo["MES_ANO"] == periodo_selecionado])
        nomeInfo = st.selectbox("Selecione o nome:", dados_Falta_completo[dados_Falta_completo["DATA"] == dataInfo]["NOME"].unique())

        indiceEdicao = dados_Falta_completo.index[(dados_Falta_completo['DATA'] == dataInfo) & (dados_Falta_completo['NOME'] == nomeInfo)].tolist()
        linhaEdicao = indiceEdicao[0] + 2

        nomes_disponiveis = dados_Quadro_completo["NOME"].unique()
        try:
            indice_pre_selecionado = list(nomes_disponiveis).index(nomeInfo)
        except ValueError:
            indice_pre_selecionado = 0  # Fallback se o nome não for encontrado

        with st.form("Editar Info"):

            data_datetime = datetime.strptime(dataInfo, "%d/%m/%Y")
            editdata = st.date_input("Dia da falta?", format="DD/MM/YYYY", value= data_datetime)
            editdata = editdata.strftime("%d/%m/%Y")  # Formato DD/MM/YYYY

            editFuncionario = st.selectbox("Informe o funcionário:", dados_Quadro_completo["NOME"].unique(), index=indice_pre_selecionado)
            editTurno = st.selectbox("Informe o turno:", ["Matutino", "Vespertino", "Dia inteiro"])
            editatestado = st.selectbox("Apresentou justificaiva (abonar falta)?", ["Sim", "Não"])
            editobservacao = st.text_input("Observações:", max_chars=50)

            valores_linha = [
                editdata,
                editFuncionario,
                editTurno,
                editatestado,
                editobservacao
            ]

            submitted = st.form_submit_button("Submeter")

            if submitted:

                st.toast(f"{valores_linha}")
            
                planilhaFalta.batch_update([{
                            'range': f'A{linhaEdicao}:E{linhaEdicao}',
                            'values': [valores_linha]
                        }])
                
                st.toast("Submissão feita com sucesso", icon=":material/thumb_up:")
