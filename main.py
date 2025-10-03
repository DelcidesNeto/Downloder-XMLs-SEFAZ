import undetected_chromedriver as uc
import threading
from undetected_chromedriver.webelement import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from pyautogui import press
from time import sleep
from collections import namedtuple
import json, calendar, configparser, subprocess, re, pytz, os, shutil
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog
from collections import namedtuple
import threading
from datetime import datetime, date








ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
class App(ctk.CTk):
    def resource_path(self, relative_path):
        import sys
        """Pega o caminho absoluto do recurso, funciona no .exe e no .py"""
        try:
            # PyInstaller cria uma pasta temporária e guarda em _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def __init__(self):
        super().__init__()
        self.processo_rodando = False
        self.title("Download de XMLs SEFAZ")
        self.iconbitmap(self.resource_path('xml.ico'))
        self.geometry("700x750")

        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(padx=20, pady=20, fill="both", expand=True)

        titulo = ctk.CTkLabel(frame, text="Configurações de Download", font=("Arial", 18, "bold"))
        titulo.pack(pady=10)

        # ----------- DATAS ----------- #
        data_frame = ctk.CTkFrame(frame)
        data_frame.pack(pady=10, fill="x", padx=10)

        ctk.CTkLabel(data_frame, text="Mês/Ano Inicial:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.data_inicial = ctk.CTkEntry(data_frame, placeholder_text="MM/AAAA")
        self.data_inicial.grid(row=0, column=1, padx=5, pady=5)
        self.data_inicial.bind("<KeyRelease>", lambda e: self.formatar_data(self.data_inicial))

        ctk.CTkLabel(data_frame, text="Mês/Ano Final:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.data_final = ctk.CTkEntry(data_frame, placeholder_text="MM/AAAA")
        self.data_final.grid(row=1, column=1, padx=5, pady=5)
        self.data_final.bind("<KeyRelease>", lambda e: self.formatar_data(self.data_final))

        # ----------- MODELO ----------- #
        modelo_frame = ctk.CTkFrame(frame)
        modelo_frame.pack(pady=10, fill="x", padx=10)

        ctk.CTkLabel(modelo_frame, text="Modelo:").pack(anchor="w", padx=5, pady=5)

        self.modelo_opcao = ctk.StringVar(value="55_65")
        self.modelo_menu = ctk.CTkOptionMenu(modelo_frame, variable=self.modelo_opcao, values=["55_65", "55", "57", "65"])
        self.modelo_menu.pack(padx=10, pady=5, fill="x")

        # ----------- SALVAR ----------- #
        salvar_frame = ctk.CTkFrame(frame)
        salvar_frame.pack(pady=10, fill="x", padx=10)

        ctk.CTkLabel(salvar_frame, text="Local de Salvamento:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.path_salvar = ctk.CTkEntry(salvar_frame, placeholder_text="Selecione uma pasta")
        self.path_salvar.grid(row=0, column=1, padx=5, pady=5, sticky="we")

        btn_pasta = ctk.CTkButton(salvar_frame, text="Procurar", command=self.selecionar_pasta)
        btn_pasta.grid(row=0, column=2, padx=5, pady=5)
        salvar_frame.columnconfigure(1, weight=1)

        # ----------- CERTIFICADO / CNPJ ----------- #
        cert_frame = ctk.CTkFrame(frame)
        cert_frame.pack(pady=10, fill="x", padx=10)

        ctk.CTkLabel(cert_frame, text="Certificado Digital:").pack(anchor="w", padx=5, pady=5)
        self.certificado_var = ctk.BooleanVar(value=True)

        ctk.CTkRadioButton(cert_frame, text="Selecionar automaticamente", variable=self.certificado_var, value=True, command=self.toggle_cnpj).pack(anchor="w", padx=10)
        ctk.CTkRadioButton(cert_frame, text="Informar CPF/CNPJ manualmente", variable=self.certificado_var, value=False, command=self.toggle_cnpj).pack(anchor="w", padx=10)

        self.cnpj_entry = ctk.CTkEntry(cert_frame, state="disabled")
        self.cnpj_entry.pack(pady=5, padx=20, fill="x")
        self.cnpj_entry.bind("<KeyRelease>", self.formatar_cnpj_cpf)

        # ----------- BOTÃO INICIAR ----------- #
        frame_final = ctk.CTkFrame(frame)
        frame_final.pack(pady=10, fill='x')

        btn_limpar_links = ctk.CTkButton(frame_final, text="Limpar Links", font=("Arial", 14, "bold"), command=self.zerarJson)
        btn_limpar_links.grid(row=0, column=0, padx=5, sticky="we")
        btn_iniciar = ctk.CTkButton(frame_final, text="Iniciar Busca + download", font=("Arial", 14, "bold"), command=self.iniciar_download)
        btn_iniciar.grid(row=0, column=1, padx=5, sticky="we")
        btn_redownload = ctk.CTkButton(frame_final, text="Baixar Xmls", font=("Arial", 14, "bold"), command=self.reDownload)
        btn_redownload.grid(row=0, column=2, padx=5, sticky="we")
        frame_final.columnconfigure((0, 1), weight=1)

        # ------------ Botão cancelar ----------#
        self.botao_cancelar = ctk.CTkButton(frame, text='Cancelar', font=('Arial', 14, 'bold'), fg_color='red', hover_color='#b01200', command=self.cancelarBusca, state='disabled')
        self.botao_cancelar.pack(pady=10)

        # ------------------ Logs --------------#
        log_frame = ctk.CTkFrame(frame)
        log_frame.pack(pady=10, fill="x", padx=10)
        ctk.CTkLabel(log_frame, text='Área de Logs:').grid(row=0)
        self.logs = ctk.CTkTextbox(log_frame, state='disabled', height=100)
        self.logs.grid(row=1, padx=5, sticky='we')
        log_frame.columnconfigure(0, weight=1)








        # ----------- FUNÇÕES ----------- #








    def adicionarLogApp(self, log: str=''):
        self.logs.configure(state='normal')
        self.logs.insert('end', log+'\n')
        self.logs.see('end')
        self.logs.configure(state='disabled')
    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.path_salvar.delete(0, "end")
            self.path_salvar.insert(0, pasta)

    def toggle_cnpj(self):
        if not self.certificado_var.get():
            self.cnpj_entry.configure(state="normal", fg_color=("gray70", "gray30"))
        else:
            self.cnpj_entry.delete(0, 'end')
            self.cnpj_entry.configure(state="disabled", fg_color=("white", "gray20"))

    def formatar_data(self, entry_widget):
        valor = entry_widget.get().replace("/", "")
        if len(valor) > 6:
            valor = valor[:6]
        if len(valor) > 2:
            valor = valor[:2] + "/" + valor[2:]
        entry_widget.delete(0, "end")
        entry_widget.insert(0, valor)
    def cancelarBusca(self):
        self.botao_cancelar.configure(state='disabled')
        self.processo_rodando = False
        self.navegador.quit()
        
    def show_message(self, titulo, texto):
        # Cria a janela modal
        popup = ctk.CTkToplevel(self)
        popup.title(titulo)
        popup.geometry("400x150")
        popup.resizable(False, False)

        largura_janela = self.winfo_width()
        altura_janela = self.winfo_height()
        x_janela = self.winfo_x()
        y_janela = self.winfo_y()

        # Largura e altura do popup
        largura_popup = popup.winfo_width()
        altura_popup = popup.winfo_height()

        # Calcula posição para centralizar
        x = x_janela + (largura_janela // 2) - (largura_popup // 2)
        y = y_janela + (altura_janela // 2) - (altura_popup // 2)

        # Aplica a posição
        popup.geometry(f"+{x}+{y}")

        # Label com o texto
        label = ctk.CTkLabel(popup, text=texto)
        label.pack(pady=30)

        # Botão OK
        btn = ctk.CTkButton(popup, text="OK", command=popup.destroy)
        btn.pack(pady=10)

        # Bloqueia interação com a janela principal até fechar o popup
        popup.grab_set()


    def formatar_cnpj_cpf(self, event):
        entry = event.widget
        texto = ''.join(filter(str.isdigit, entry.get()))

        if len(texto) <= 11:  # CPF
            texto = texto[:11]  # limita a 11 dígitos
            formato = ''
            for i, d in enumerate(texto):
                if i in [3, 6]:
                    formato += '.'
                elif i == 9:
                    formato += '-'
                formato += d
        else:  # CNPJ
            texto = texto[:14]  # limita a 14 dígitos
            formato = ''
            for i, d in enumerate(texto):
                if i in [2, 5]:
                    formato += '.'
                elif i == 8:
                    formato += '/'
                elif i == 12:
                    formato += '-'
                formato += d

        entry.delete(0, 'end')
        entry.insert(0, formato)


    def iniciar_download(self):
        if not self.processo_rodando:
            dataInicial = self.data_inicial.get()
            dataFinal = self.data_final.get()
            modelo = self.modelo_opcao.get()
            pathDeDownload = self.path_salvar.get()
            autoSelectCert = self.certificado_var.get()
            cpfCnpj = self.cnpj_entry.get().replace('.', '').replace('-', '').replace('/', '')
            if dataInicial == '':
                self.show_message('Campos Faltando', 'Preencha a data inicial!')
            elif dataFinal == '':
                self.show_message('Campos Faltando', 'Preencha a data final!')
            elif pathDeDownload == '':
                self.show_message('Campos Faltando', 'Preencha o local de download!')
            elif not autoSelectCert and cpfCnpj == '':
                self.show_message('Campos Faltando', 'Preencha o campo de CPF/CNPJ!')
            else:
                MyRecord = namedtuple('Record', ['MesAnoInicial', 'MesAnoFinal', 'PathDeDownload', 'AutoSelectCert', 'ModeloDoDocumento', 'CpfCnpj'])
                pathDeDownload += f'/{cpfCnpj}'
                record = MyRecord(MesAnoInicial=dataInicial, MesAnoFinal=dataFinal, PathDeDownload=pathDeDownload, AutoSelectCert=autoSelectCert, ModeloDoDocumento=modelo, CpfCnpj=cpfCnpj)
                self.thread = threading.Thread(target=self.buscarXmls, args=(record,), daemon=True)
                self.thread.start()
        else:
            self.show_message('Erro', 'Você já iniciou um processo, cancele ele para iniciar este!')


    def EsperarParaApertarTab(self):
        # Dá um tempo para o popup de certificado aparecer
        sleep(2)
        # Pressiona tab algumas vezes e depois Enter para selecionar o certificado
        for c in range(0, 2):
            press('tab')
            sleep(0.5)
        press('enter')

    def confirmation(self, titulo, mensagem):
        resposta = {"value": None}  # Usamos um dict para poder modificar dentro do popup

        popup = ctk.CTkToplevel()
        popup.title(titulo)
        popup.geometry("450x150")
        popup.resizable(False, False)
        
        largura_janela = self.winfo_width()
        altura_janela = self.winfo_height()
        x_janela = self.winfo_x()
        y_janela = self.winfo_y()

        # Largura e altura do popup
        largura_popup = popup.winfo_width()
        altura_popup = popup.winfo_height()

        # Calcula posição para centralizar
        x = x_janela + (largura_janela // 2) - (largura_popup // 2)
        y = y_janela + (altura_janela // 2) - (altura_popup // 2)

        # Aplica a posição
        popup.geometry(f"+{x}+{y}")

        ctk.CTkLabel(popup, text=mensagem).pack(pady=20)

        def sim():
            resposta["value"] = True
            popup.destroy()

        def nao():
            resposta["value"] = False
            popup.destroy()

        btn_frame = ctk.CTkFrame(popup)
        btn_frame.pack(pady=10, fill="x")

        ctk.CTkButton(btn_frame, text="Sim", command=sim).pack(side="left", expand=True, padx=10)
        ctk.CTkButton(btn_frame, text="Não", command=nao).pack(side="right", expand=True, padx=10)

        popup.grab_set()  # Bloqueia a interação com a janela principal até fechar o popup
        popup.wait_window()  # Espera o popup fechar

        return resposta["value"]
    

    def zerarJson(self, criarArquivo=False):
        if criarArquivo or self.confirmation('Remover Links', 'Essa ação não pode ser desfeita, deseja apagar os links de download?'):
            jsonInicial = {'links': []}
            with open('links.json','wt+') as arqivoJson:
                json.dump(jsonInicial, arqivoJson, indent=4)
        


    def adicionarLinks(self, link):
        try:
            with open(f'{os.getcwd()}\links.json', 'r', encoding='utf-8') as arquivoJson:
                linksJson = json.load(arquivoJson)
                linksJson['links'].append(link)
                with open('links.json','wt+') as links:
                    json.dump(linksJson, links, indent=4)
        except:
            self.zerarJson(criarArquivo=True) # Chama a criação do json, que em tese também o zera


    def GetMes(self, AMes: str):
        pos = AMes.find('/')
        return int(AMes[:pos])


    def GetAno(self, AAno: str):
        pos = AAno.find('/')
        return int(AAno[pos+1:])


    def calcular_diferenca_em_meses(self, mes_ano_inicial, mes_ano_final):
        # Calcula a diferença entre os meses
        ano_inicial = self.GetAno(mes_ano_inicial)
        mes_inicial = self.GetMes(mes_ano_inicial)
        ano_final = self.GetAno(mes_ano_final)
        mes_final = self.GetMes(mes_ano_final)

        # Calculando a diferença em meses
        diff_anos = ano_final - ano_inicial
        diff_meses = mes_final - mes_inicial
        return diff_anos * 12 + diff_meses

    def GetLastDay(self, mes_ano):
        """
        Recebe uma string no formato 'MM/YYYY' e retorna o último dia daquele mês.
        Exemplo: '03/2025' retorna 31 (porque março tem 31 dias)
        """
        try:
            # Separar mês e ano
            mes, ano = mes_ano.split('/')
            mes = int(mes)
            ano = int(ano)
            # Usar calendar para obter o último dia do mês
            ultimo_dia = calendar.monthrange(ano, mes)[1]
            if ultimo_dia < 10:
                ultimo_dia = f'0{ultimo_dia}'
            return ultimo_dia
        except Exception as e:
            print(f"Erro ao calcular último dia: {e}")
            return None
        

    def GetDadosArqIni():
        # Cria uma instância do ConfigParser
        config = configparser.ConfigParser()

        # Lê o arquivo .ini
        config.read('config.ini')

        # Acessa uma seção e lê uma chave
        MesAnoInicial = config['Dados-Para-Download-xmls']['mes/ano-inicial']
        MesAnoFinal = config['Dados-Para-Download-xmls']['mes/ano-final']
        PathDeDownload = config['Dados-Para-Download-xmls']['path-dos-xmls']
        SelecionarCertificadoAutomaticamente = config['Dados-Para-Download-xmls']['selecionar-certificado-automaticamente']
        ModeloDoDocumento = config['Dados-Para-Download-xmls']['modelo-do-documento']
        AutoSelectCert = False
        if SelecionarCertificadoAutomaticamente == 'True':
            AutoSelectCert = True
        MyRecord = namedtuple('Record', ['MesAnoInicial', 'MesAnoFinal', 'PathDeDownload', 'AutoSelectCert', 'ModeloDoDocumento'])
        return MyRecord(MesAnoInicial=MesAnoInicial, MesAnoFinal=MesAnoFinal, PathDeDownload=PathDeDownload, AutoSelectCert=AutoSelectCert, ModeloDoDocumento=ModeloDoDocumento)


    def fazerPesquisa(self, navegador: uc.Chrome, data: str, modeloDoDocumento: str):
        while True:
            Status = ''
            try:
                navegador.find_element(By.ID, 'cmpDataInicial').click()
            except:
                navegador.get('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica') #Caso ele não encontre o label da data, significa que está na página de erro ainda
                sleep(1)
                navegador.find_element(By.ID, 'cmpDataInicial').click()
            sleep(1)
            navegador.find_element(By.ID, 'cmpDataInicial').send_keys(f'01/{data}')
            sleep(1)
            navegador.find_element(By.ID, 'cmpDataFinal').click()
            sleep(1)
            navegador.find_element(By.ID, 'cmpDataFinal').send_keys(f'{self.GetLastDay(data)}/{data}')
            sleep(1)
            Select(navegador.find_element(By.ID, 'cmpModelo')).select_by_value(modeloDoDocumento)
            try:
                WebDriverWait(navegador, 10).until(
                    lambda driver: driver.find_element(By.NAME, "g-recaptcha-response").get_attribute("value") != ""
                )
            except:
                navegador.refresh()
                continue
            navegador.find_element(By.ID, 'btnPesquisar').click()
            sleep(1)
            if navegador.find_element(By.ID, 'message-containter').is_displayed():
                Status = navegador.find_element(By.ID, 'message-containter').get_attribute('innerText')
            return Status


    def PegarCnpjAuto(self):
        try:
            # Listar certificados instalados
            resultado = subprocess.check_output(
                'certutil -store -user My', 
                shell=True, 
                text=True
            )

            # Procurar por padrões de CNPJ (14 dígitos) no resultado
            cnpjs = re.findall(r'\b\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}\b', resultado)
            if cnpjs:
                return cnpjs[0]  # Retorna o primeiro CNPJ encontrado
            else:
                # Procurar por padrão numérico que pode ser um CNPJ sem formatação
                cnpjs_sem_formato = re.findall(r'\b\d{14}\b', resultado)
                if cnpjs_sem_formato:
                    cnpj = cnpjs_sem_formato[0]
                    # Formatar o CNPJ
                    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}".replace(".", "").replace("/", "").replace("-", "")
                
            print("Nenhum CNPJ encontrado nos certificados")
            return None
        except Exception as e:
            print(f"Erro ao buscar CNPJ: {e}")
            return None


    def PegarCnpjManual(CertificadoNome):
        try:
            # Listar certificados instalados
            resultado = subprocess.check_output(
                'certutil -store -user My', 
                shell=True, 
                text=True
            )

            # Encontra a seção correspondente ao certificado desejado
            certificados = resultado.split("\n")
            capturando = False
            certificado_dados = ""

            for linha in certificados:
                if CertificadoNome.lower() in linha.lower():
                    capturando = True  # Começa a capturar as informações do certificado
                if capturando:
                    certificado_dados += linha + "\n"
                if capturando and "Certificado" in linha:
                    break  # Para a captura quando termina o bloco do certificado

            # Procurar por CNPJ formatado
            cnpj = re.search(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", certificado_dados)
            if cnpj:
                return cnpj.group()

            # Procurar por CNPJ sem formatação
            cnpj_sem_formato = re.search(r"\b\d{14}\b", certificado_dados)
            if cnpj_sem_formato:
                cnpj = cnpj_sem_formato.group()
                return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}".replace(".", "").replace("/", "").replace("-", "")
            
            print("Nenhum CNPJ encontrado no certificado")
            return None
        except Exception as e:
            print(f"Erro ao buscar CNPJ: {e}")
            return None


    def AdicionarLog(self, log=''):
        with open(f'{os.getcwd()}\log.txt', 'at') as Arquivo:
            Arquivo.write(f'{log}\n')


    def realizarDownloadXmls(self, pathDownload: str):
        with open('links.json', 'rt', encoding='utf-8') as arquivoJson:
            linksJson = json.load(arquivoJson)['links']
        for link in linksJson:
            self.navegador.get(link)
        os.makedirs(pathDownload, exist_ok=True)
        self.awaitDownload(pathDownload)
        self.botao_cancelar.configure(state='disabled')
        self.processo_rodando = False
        self.navegador.quit()
        

    def reDownload(self):
        def procedure(cpfCnpj: str, autoSelectCert: bool, pathDeDownload: str):
            opcoes = uc.ChromeOptions()
            if autoSelectCert:
                CNPJ = self.PegarCnpjAuto()
            else:
                CNPJ = cpfCnpj
            # prefs = {
            #     "download.default_directory": pathDeDownload,   # muda o path do download
            #     "download.prompt_for_download": False,         # não perguntar onde salvar
            #     "download.directory_upgrade": True,            # sobrescreve se necessário
            #     "safebrowsing.enabled": True,                  # permite downloads "não seguros"
            #     "profile.default_content_settings.popups": 0
            # }
            # opcoes.add_experimental_option("prefs", prefs)
            DataLog = datetime.today().astimezone(pytz.timezone('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M:%S')
            self.processo_rodando = True
            self.adicionarLogApp(f'{DataLog} - Iniciando o download dos links armazenados...')
            self.navegador = uc.Chrome(headless=False, use_subprocess=True, options=opcoes)
            self.botao_cancelar.configure(state='normal')
            
            if autoSelectCert:
                threading.Thread(target=self.EsperarParaApertarTab).start()
            
            self.navegador.get('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica')
            CnpjSite = self.navegador.find_element(By.XPATH, '//*[@id="cmpCnpj"]').get_attribute('innerText')
            if CNPJ != CnpjSite:
                self.AdicionarLog(f'{DataLog} - CPF/CNPJ do certificado: {CNPJ}, CPF/CNPJ do site: {CnpjSite}| Os dados não batem, então a busca não será realizada...')
                self.adicionarLogApp(f'{DataLog} - CPF/CNPJ do certificado: {CNPJ}, CPF/CNPJ do site: {CnpjSite}| Os dados não batem, então a busca não será realizada...')
                self.botao_cancelar.configure(state='disabled')
                self.processo_rodando = False
                self.navegador.quit()
            os.makedirs(pathDeDownload, exist_ok=True)
            cpfCnpj = pathDeDownload[pathDeDownload.rfind('/')+1:]
            with open('links.json', 'rt', encoding='utf-8') as arquivoJson:
                linksJson = json.load(arquivoJson)['links']
                for link in linksJson:
                    if cpfCnpj in link:
                        self.navegador.get(link)
                    
            self.awaitDownload(pathDeDownload)
            self.adicionarLogApp(f'{DataLog} - Verifique a sua pasta de downloads localizada em {pathDeDownload}')
            self.processo_rodando = False
            self.navegador.quit()
            self.botao_cancelar.configure(state='disabled')
            


        if not self.processo_rodando:
            autoSelectCert = self.certificado_var.get()
            cpfCnpj = self.cnpj_entry.get().replace('.', '').replace('-', '').replace('/', '')
            pathDeDownload = self.path_salvar.get()+'/'+cpfCnpj
            if pathDeDownload == '':
                self.show_message('Campos Faltando', 'Preencha o local de download!')
            elif not autoSelectCert and cpfCnpj == '':
                self.show_message('Campos Faltando', 'Preencha o campo de CPF/CNPJ!')
            else:
                self.thread = threading.Thread(target=procedure, args=(cpfCnpj, autoSelectCert, pathDeDownload), daemon=True)
                self.thread.start()
        else:
            self.show_message('Erro', 'Você já iniciou um processo, cancele ele para iniciar este!')

    # def reDownload(self):
    #     def procedure(cpfCnpj: str, autoSelectCert: bool, pathDeDownload: str):
    #         opcoes = uc.ChromeOptions()
    #         if autoSelectCert:
    #             CNPJ = self.PegarCnpjAuto()
    #         else:
    #             CNPJ = cpfCnpj
    #         prefs = {
    #             "download.default_directory": pathDeDownload,   # muda o path do download
    #             "download.prompt_for_download": False,         # não perguntar onde salvar
    #             "download.directory_upgrade": True,            # sobrescreve se necessário
    #             "safebrowsing.enabled": True,                  # permite downloads "não seguros"
    #             "profile.default_content_settings.popups": 0,
    #         }
    #         opcoes.add_experimental_option("prefs", prefs)
            

    #         self.navegador = uc.Chrome(headless=False, use_subprocess=True, options=opcoes)
    #         if autoSelectCert:
    #             threading.Thread(target=self.EsperarParaApertarTab).start()
    #         while True:
    #             self.navegador.get('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica')
    #             try:
    #                 WebDriverWait(self.navegador, 10).until(
    #                     lambda driver: driver.find_element(By.NAME, "g-recaptcha-response").get_attribute("value") != ""
    #                 )
    #                 self.navegador.find_element(By.ID, 'btnHistoricoDownload').click()
    #                 break
                    
    #             except:
    #                 self.navegador.refresh()
    #                 continue
    #         for linha in self.navegador.find_element(By.TAG_NAME, 'table').find_element(By.TAG_NAME, 'tbody').find_elements(By.TAG_NAME, 'tr'):
    #             id = linha.find_element(By.CLASS_NAME, 'col-arquivo').get_attribute('innerText')
    #             linha.find_element(By.CLASS_NAME, 'col-acoes').find_element(By.ID, id).click()      
    #         self.awaitDownload(pathDeDownload)
    #         self.navegador.quit()

        
    #     pathDeDownload = self.path_salvar.get()
    #     autoSelectCert = self.certificado_var.get()
    #     cpfCnpj = self.cnpj_entry.get().replace('.', '').replace('-', '').replace('/', '')
    #     if pathDeDownload == '':
    #         self.show_message('Campos Faltando', 'Preencha o local de download!')
    #     # elif not autoSelectCert and cpfCnpj == '':
    #     #     self.show_message('Campos Faltando', 'Preencha o campo de CPF/CNPJ!')
    #     else:
    #         self.thread = threading.Thread(target=procedure, args=(cpfCnpj, autoSelectCert, pathDeDownload), daemon=True)
    #         self.thread.start()

    def moverArquivos(self, pathDownloadUserSystem: str, pathDeDownload: str):
        arquivos = os.listdir(pathDownloadUserSystem)
        cpfCnpj = pathDeDownload[pathDeDownload.rfind('/')+1:]
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pathDownloadUserSystem, arquivo)
            data_criacao = datetime.fromtimestamp(os.path.getmtime(caminho_arquivo)).date()
            if data_criacao == date.today() and cpfCnpj in caminho_arquivo:
                shutil.move(caminho_arquivo, pathDeDownload)


    def awaitDownload(self, pathDeDownload: str):
        pathDownloadUserSystem = os.path.join(os.path.expanduser('~'), 'Downloads')
        sleep(2)
        while True:
            esperar = False
            sleep(1)
            arquivos = os.listdir(pathDownloadUserSystem)
            for arquivo in arquivos:
                caminho = os.path.join(pathDownloadUserSystem, arquivo)
                data_criacao = datetime.fromtimestamp(os.path.getmtime(caminho)).date()
                if data_criacao == date.today():
                    if '.crdownload' in arquivo:
                        esperar = True
                        break
            if esperar == False:
                break
        self.moverArquivos(pathDownloadUserSystem, pathDeDownload)
        


    def buscarXmls(self, dadosRecord):
        opcoes = uc.ChromeOptions()
        if (dadosRecord.AutoSelectCert):
            CNPJ = self.PegarCnpjAuto()
        else:
            CNPJ = dadosRecord.CpfCnpj
        # prefs = {
        #     "download.default_directory": dadosRecord.PathDeDownload+f'/{CNPJ}',   # muda o path do download
        #     "download.prompt_for_download": False,         # não perguntar onde salvar
        #     "download.directory_upgrade": True,            # sobrescreve se necessário
        #     "safebrowsing.enabled": True,                  # permite downloads "não seguros"
        #     "profile.default_content_settings.popups": 0
        # }
        # opcoes.add_experimental_option("prefs", prefs)
        self.processo_rodando = True
        self.navegador = uc.Chrome(headless=False, use_subprocess=True, options=opcoes) # version_main=140
        self.botao_cancelar.configure(state='normal')

        if dadosRecord.AutoSelectCert:
            threading.Thread(target=self.EsperarParaApertarTab).start()


        self.navegador.get('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica')


        CnpjSite = self.navegador.find_element(By.XPATH, '//*[@id="cmpCnpj"]').get_attribute('innerText')

        if CNPJ != CnpjSite:
            self.AdicionarLog(f'CPF/CNPJ do certificado: {CNPJ}, CPF/CNPJ do site: {CnpjSite}| Os dados não batem, então a busca não será realizada...')
            self.adicionarLogApp(f'CPF/CNPJ do certificado: {CNPJ}, CPF/CNPJ do site: {CnpjSite}| Os dados não batem, então a busca não será realizada...')
            self.botao_cancelar.configure(state='disabled')
            self.processo_rodando = False
            self.navegador.quit()

        else:
            anoInical = self.GetAno(dadosRecord.MesAnoInicial)
            mesInicial = self.GetMes(dadosRecord.MesAnoInicial)
            anoFinal = self.GetAno(dadosRecord.MesAnoFinal)
            mesFinal =  self.GetMes(dadosRecord.MesAnoFinal)
            quantidadeDeMeses = self.calcular_diferenca_em_meses(dadosRecord.MesAnoInicial, dadosRecord.MesAnoFinal)+1
            primeiroLoop = True
            for c in range(0, quantidadeDeMeses):
                CnpjSite = self.navegador.find_element(By.XPATH, '//*[@id="cmpCnpj"]').get_attribute('innerText')
                encontrouDocumentos = True
                DataLog = datetime.today().astimezone(pytz.timezone('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M:%S')
                primeiroLoop = False
                Data = f'0{mesInicial}/{anoInical}' if mesInicial < 10 else f'{mesInicial}/{anoInical}'
                if CNPJ != CnpjSite:
                    self.AdicionarLog(f'{DataLog} - O CPF/CNPJ do certificado não confere com o CPF/CNPJ do site, caso queira continuar de onde parou configure o mes/ano-inicial = {Data}')
                    self.adicionarLogApp(f'{DataLog} - O CPF/CNPJ do certificado não confere com o CPF/CNPJ do site, caso queira continuar de onde parou configure o mes/ano-inicial = {Data}')
                    self.navegador.quit()
                    exit()
                Erro1 = self.fazerPesquisa(navegador=self.navegador, data=Data, modeloDoDocumento=dadosRecord.ModeloDoDocumento)
                if Erro1 != '':
                    Erro2 = self.fazerPesquisa(navegador=self.navegador, data=Data, modeloDoDocumento=dadosRecord.ModeloDoDocumento)
                    if Erro2 != '':
                        Erro3 = self.fazerPesquisa(navegador=self.navegador, data=Data, modeloDoDocumento=dadosRecord.ModeloDoDocumento)
                        if Erro3 != '':
                            self.AdicionarLog(f'{DataLog} - Erro ao pesquisar a data {Data}, mês sem resultados para o CNPJ {CNPJ}!')
                            self.adicionarLogApp(f'{DataLog} - Erro ao pesquisar a data {Data}, mês sem resultados para o CNPJ {CNPJ}!')
                            self.navegador.get('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica')
                            encontrouDocumentos = False
                if encontrouDocumentos:
                    self.navegador.find_element(By.XPATH, '//*[@id="content"]/div[2]/div/button').click()
                    sleep(1)
                    self.navegador.find_element(By.ID, 'dnwld-all-btn-ok').click()
                    sleep(1)
                    linkBase = 'https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica/resultado/download/historico/arquivo/'
                    restoLink = self.navegador.find_element(By.TAG_NAME, 'table').find_element(By.TAG_NAME, 'tbody').find_elements(By.TAG_NAME, 'tr')[0].find_element(By.CLASS_NAME, 'col-arquivo').get_attribute('innerText')
                    linkDownload = linkBase+restoLink
                    self.adicionarLinks(linkDownload)
                    self.AdicionarLog(f'{DataLog} - Link de download de {Data} obtido com sucesso!')
                    self.adicionarLogApp(f'{DataLog} - Link de download de {Data} obtido com sucesso!')
                    self.navegador.get('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica')
                if not primeiroLoop:
                    if mesInicial == mesFinal and anoInical == anoFinal:
                        self.adicionarLogApp(f"{DataLog} - Fim das pesquisas, preparando-se para realizar os downloads...")
                        self.adicionarLogApp(f'{DataLog} - O sistema irá aguardar 1 minuto antes de iniciar o download por questões de segurança')
                        sleep(60)
                        self.adicionarLogApp(f'{DataLog} - Downloads iniciados, aguarde até a conclusão...')
                        self.realizarDownloadXmls(pathDownload=dadosRecord.PathDeDownload)
                        self.adicionarLogApp(f'f{DataLog} - Verifique a sua pasta de downloads localizada em {dadosRecord.PathDeDownload}')
                if mesInicial < 13:
                    mesInicial += 1
                if mesInicial == 13:
                    mesInicial = 1
                    anoInical += 1


if __name__ == "__main__":
    app = App()
    app.mainloop()
