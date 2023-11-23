import tkinter
import tkinter.messagebox
from tkinter.messagebox import showerror, showinfo, askquestion
import customtkinter
from tkinter.filedialog import askdirectory, askopenfilename
import fitz
from re import search
from os import listdir, path
from num2words import num2words
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import subprocess

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")
#configurações de aparencia acima!

class App(customtkinter.CTk): #criação da interface do sistema.
    def __init__(self):
        super().__init__()

        self.title("Emissor e extrator")
        self.geometry(f"{880}x{480}")

        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Extrator de NF3E", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="EDP", command=self.extrairEDP)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=20)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="ENERGISA", command=self.extrairENERG)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=20)
        self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, text="CEMIG", command=self.extrairCEMIG)
        self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=20)

        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text='Sair', text_color=("gray10", "#DCE4EE"), command=self.close)
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.radiobutton_frame = customtkinter.CTkFrame(self)
        self.radiobutton_frame.grid(row=0, column=1, padx=(30, 0), pady=20, sticky="nsw")
        self.varfrm1 = tkinter.StringVar(value='')
        self.label_radio_group1 = customtkinter.CTkLabel(master=self.radiobutton_frame, text="Emissor de recibo", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.label_radio_group1.grid(row=0, column=1, columnspan=1, padx=(5, 15), pady=10, sticky="")
        self.label_radio_group1 = customtkinter.CTkLabel(master=self.radiobutton_frame, text="Marque o pagador:", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_radio_group1.grid(row=1, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, text="Empresa teste 1", variable=self.varfrm1, value='emp1')
        self.radio_button_1.grid(row=2, column=1, pady=10, padx=20, sticky="")
        self.radio_button_2 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, text="Empresa teste 2", variable=self.varfrm1, value='emp2')
        self.radio_button_2.grid(row=3, column=1, pady=10, padx=20, sticky="")

        self.varfrm2 = tkinter.StringVar(value='')
        self.label_radio_group2 = customtkinter.CTkLabel(master=self.radiobutton_frame, text="Marque o recebedor:", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_radio_group2.grid(row=4, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.radio_button_3 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, text="Empresa teste 1", variable=self.varfrm2, value='emp1')
        self.radio_button_3.grid(row=5, column=1, pady=10, padx=20, sticky="")
        self.radio_button_4 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, text="Sócio teste 1", variable=self.varfrm2, value='soc1')
        self.radio_button_4.grid(row=6, column=1, pady=10, padx=20, sticky="")
        self.radio_button_5 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, text="Sócio teste 2", variable=self.varfrm2, value='soc2')
        self.radio_button_5.grid(row=7, column=1, pady=10, padx=20, sticky="")

        self.radiobutton_frame1 = customtkinter.CTkFrame(self)
        self.radiobutton_frame1.grid(row=0, column=2, padx=10, pady=20, sticky="nsw")

        self.varfrm3 = tkinter.StringVar(value='')
        self.label_radio_group3 = customtkinter.CTkLabel(master=self.radiobutton_frame1, text="Será um adiantamento?", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_radio_group3.grid(row=7, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.radio_button_6 = customtkinter.CTkRadioButton(master=self.radiobutton_frame1, text="Sim", variable=self.varfrm3, value='S')
        self.radio_button_6.grid(row=8, column=1, pady=10, padx=20, sticky="n")
        self.radio_button_7 = customtkinter.CTkRadioButton(master=self.radiobutton_frame1, text="Não", variable=self.varfrm3, value='N')
        self.radio_button_7.grid(row=9, column=1, pady=10, padx=20, sticky="n")

        self.label_valorr = customtkinter.CTkLabel(master=self.radiobutton_frame1, text="Valor:", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_valorr.grid(row=10, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.valorr = customtkinter.CTkEntry(self.radiobutton_frame1, placeholder_text="Valor no formato: 1500.50")
        self.valorr.grid(row=11, column=1, padx=10, pady=0, sticky="nsew")

        self.label_data = customtkinter.CTkLabel(master=self.radiobutton_frame1, text="Data:", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_data.grid(row=12, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.data = customtkinter.CTkEntry(self.radiobutton_frame1, placeholder_text="Data no formato: 310323")
        self.data.grid(row=13, column=1, padx=10, pady=0, sticky="nsew")
        self.label_espaco = customtkinter.CTkLabel(master=self.radiobutton_frame1, text="\n")
        self.label_espaco.grid(row=14, column=1, padx=0, pady=0, sticky="nsew")
        self.botao_gerar = customtkinter.CTkButton(self.radiobutton_frame1, text="Gerar", command=self.gerar)
        self.botao_gerar.grid(row=15, column=1, padx=10, pady=10)

        self.extratorpgtos_frame = customtkinter.CTkFrame(self)
        self.extratorpgtos_frame.grid(row=0, column=3, padx=20, pady=20, sticky="nsew")

        self.label_radio_group1 = customtkinter.CTkLabel(master=self.extratorpgtos_frame, text="Extrator de pgtos", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.label_radio_group1.grid(row=0, column=1, columnspan=1, padx=10, pady=10, sticky="")

        self.label_data1 = customtkinter.CTkLabel(master=self.extratorpgtos_frame, text="Data do pagamento:", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_data1.grid(row=1, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.dt = customtkinter.CTkEntry(self.extratorpgtos_frame, placeholder_text="Data no formato: 310323")
        self.dt.grid(row=2, column=1, padx=10, pady=0, sticky="nsew")

        self.label_ref = customtkinter.CTkLabel(master=self.extratorpgtos_frame, text="Mês de referência:", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.label_ref.grid(row=3, column=1, columnspan=1, padx=10, pady=10, sticky="")
        self.ref = customtkinter.CTkEntry(self.extratorpgtos_frame, placeholder_text="Mês no formato: 032023")
        self.ref.grid(row=4, column=1, padx=10, pady=0, sticky="nsew")

        self.label_espaco1 = customtkinter.CTkLabel(master=self.extratorpgtos_frame, text="\n")
        self.label_espaco1.grid(row=5, column=1, padx=0, pady=0, sticky="nsew")
        self.label_espaco2 = customtkinter.CTkLabel(master=self.extratorpgtos_frame, text="\n")
        self.label_espaco2.grid(row=6, column=1, padx=0, pady=0, sticky="nsew")
        self.botao_gerar = customtkinter.CTkButton(self.extratorpgtos_frame, text="Buscar e Gerar", command=self.pgtos)
        self.botao_gerar.grid(row=7, column=1, padx=10, pady=0)
        self.label_espaco4 = customtkinter.CTkLabel(master=self.extratorpgtos_frame, text="\n")
        self.label_espaco4.grid(row=8, column=1, padx=0, pady=0, sticky="nsew")
    # definições das funções a serem utilizadas abaixo:
    def extrairEDP(self): #extrator das faturas de energia da fornecedora EDP.
        try:
            buscar = askdirectory() #serão escritos dois txt, o primeiro é o layout do SPED para importar as faturas para o sistema fiscal do Alterdata, o segundo é para importar as duplicatas ref. as faturas.
            with open("Arquivos gerados\EDPextrator.txt", "w") as arquivo1:
                arquivo1.write("|0000|CÓDIGO EMPRESA|0|DATA INICIAL|DATA FINAL|NOME DA MINHA EMPRESA|CNPJ||UF|INSC. ESTADUAL|CÓD IBGE - MUNICIPIO|||A|1|\n|0001|0|\n|0005|NOME EMPRESA|CEP|RUA|NUM||BAIRRO|TEL|||\n|0100|NOME CONTADOR|CPF|CRC||CEP|RUA|NUM||BAIRRO|TEL||E-MAIL|CÓD IBGE - MUNICIPIO|\n|0150|F28152650000171|EDP ESPIRITO SANTO DISTRIBUIÇÃO DE ENERGIA|01058|28152650000171||080250165|3205309||RUA FLORENTINO FALLER|80|1º, 2º E 3º ANDAR - SALA 101, 102, 201, 202, 301 E 302|ENSEADA DO SUA|\n|0990|6|\n|B001|1|\n|B990|2|\n|C001|0|\n")
            with open("Arquivos gerados\DuplEDP.txt", "w") as arquivo2:
                arquivo2.write("Numero da Nota; Data escrituracao; CPF/CNPJ do participante; Numero da Duplicata; Data de Vencimento da Duplicata; Valor da Duplicata; Observacao\n")
            for file in listdir(buscar): #aqui ele vai percorrer todos os pdf que estão na pasta selecionada.
                if search("pdf", file):
                    documento = path.join(buscar, file)
                    doc = fitz.open(documento)
                    pagina = doc[0]
                    texto = pagina.get_text() #vai extratir o conteudo dos pdf como texto.
                    arquivo = open("Arquivos gerados\EDPextrator.txt", 'a')
                    duplicatas = open("Arquivos gerados\DuplEDP.txt", 'a')
                    # aqui abaixo começa a procurar pelas informações que precisa para preencher os campos do layout.
                    posiniNum = texto.find('Conta de Energia Elétrica nº') + 29
                    posfinNum = texto.find('\n', posiniNum)
                    numero = texto[posiniNum:posfinNum].strip()
                    numero = numero.replace('.', '')

                    datafind = texto.find('Emissão')
                    posiniEmi = texto.find('\n', datafind) + 1
                    posfinEmi = texto.find('\n', posiniEmi)
                    data1 = texto[posiniEmi:posfinEmi].strip()
                    data = data1.replace('/', '')

                    obs = ''
                    proxfat = texto.find('SERÁ COBRADO NA PRÓXIMA FATURA')
                    if proxfat != -1:
                        venc = data1
                        obs = 'O VALOR SERA COBRADO NA PROXIMA FATURA'
                    else:
                        datavenc = texto.find('Data de Vencimento')
                        if datavenc != -1:
                            posiniVenc = texto.find('\n', datavenc) + 1
                            posfinVenc = texto.find('\n', posiniVenc)
                            venc = texto[posiniVenc:posfinVenc].strip()
                        else:
                            posiniVenc = texto.find('Conta de Energia Elétrica nº') + 40
                            posiniVenc1 = texto.find('/', posiniVenc) - 2
                            posfinVenc = texto.find('', posiniVenc1) + 10
                            venc = texto[posiniVenc1:posfinVenc]

                    tipo = ''
                    trifasico = texto.find('TRIFÁSICO')
                    bifasico = texto.find('BIFÁSICO')
                    monofasico = texto.find('MONOFÁSICO')
                    if trifasico != -1:
                        tipo = '3'
                    if bifasico != -1:
                        tipo = '2'
                    if monofasico != -1:
                        tipo = '1'

                    tensao = ''
                    residencial = texto.find('RESIDENCIAL')
                    rural = texto.find('RURAL')
                    comercial = texto.find('COMERCIAL')
                    industrial = texto.find('INDUSTRIAL')
                    if residencial != -1:
                        tensao = '07'
                    if rural != -1:
                        tensao = '09'
                    if comercial != -1:
                        tensao = '12'
                    if industrial != -1:
                        tensao = '04'

                    posiniVR = texto.find('Fornecimento de energia elétrica') + 33
                    posfinVR = texto.find(',', posiniVR)
                    valorr = texto[posiniVR:posfinVR + 3]
                    valorr = valorr.replace('.', '').replace(',', '.')
                    valor_base = float(valorr)

                    bonus = texto.find('Bônus de Itaipu Lei 10.438/02')
                    if bonus != -1:
                        posiniDESC = texto.find('Bônus de Itaipu Lei 10.438/02') + 30
                        posfinDESC = texto.find(',', posiniDESC)
                        valord = texto[posiniDESC:posfinDESC + 3]
                        valord = valord.replace('.', '').replace(',', '.')
                        valor_desc = float(valord)
                    else:
                        valor_desc = 0

                    devolucao = texto.find('Devolução')
                    if devolucao != -1:
                        posiniDev = texto.find('\n', devolucao) + 1
                        posfinDev = texto.find(',', posiniDev) + 3
                        devol = texto[posiniDev:posfinDev].strip()
                        devol = devol.replace('.', '').replace(',', '.')
                        valor_devol = float(devol)
                    else:
                        valor_devol = 0

                    iluminacao = texto.find('Contribuição de Ilum. Pública - Lei Municipal')
                    if iluminacao != -1:
                        posiniILU = texto.find('Contribuição de Ilum. Pública - Lei Municipal')
                        posiniILU1 = texto.find('\n', posiniILU) + 1
                        posfinILU = texto.find(',', posiniILU1)
                        valori = texto[posiniILU1:posfinILU + 3].strip()
                        valori = valori.replace('.', '').replace(',', '.')
                        valor_ilum = float(valori)
                    else:
                        valor_ilum = 0

                    interrup = texto.find('DMIC - Duração Max Interrup em')
                    if interrup != -1:
                        posiniINT = texto.find('DMIC - Duração Max Interrup em') + 37
                        posfinINT = texto.find(',', posiniINT)
                        valorint = texto[posiniINT:posfinINT + 3]
                        valorint = valorint.replace('.', '').replace(',', '.')
                        valor_interrup = float(valorint)
                    else:
                        valor_interrup = 0

                    valor = (valor_base + valor_ilum - valor_desc - valor_interrup - valor_devol)
                    valor = f'{valor:.2f}'
                    valor = str(valor)
                    valor = valor.replace('.', ',')
                    # aqui abaixo ele escreve as informações nos dois txt conforme o layout necessitado.
                    arquivo.write(f'|C500|0|1|F28152650000171|06|00|||02|{numero}|{data}|{data}|{valor}|0|{valor}|0|0|0|0,00|0,00|0|0||0|0|{tipo}|{tensao}||||||52||||||||\n')
                    arquivo.write(f'|C590|090|1255|0|{valor}|0,00|0,00|0|0|0||\n')

                    duplicatas.write(f'"{numero}"; "{data1}"; "28152650000171"; "01"; "{venc}"; "{valor}"; "{obs}"\n')
                    duplicatas.close()

            arquivo = open("Arquivos gerados\EDPextrator.txt", 'a')
            arquivo.write('\n')
            arquivo.close()

        except:
            showerror(title='ERRO', message='Erro ao extrair PDF!')
            return
        showinfo(title='Extração', message='Arquivo PDF extraído com sucesso!')

    def extrairENERG(self): #extrator das faturas de energia da fornecedora ENERGISA.
        try:
            buscar = askdirectory() #serão escritos dois txt, o primeiro é o layout do SPED para importar as faturas para o sistema fiscal do Alterdata, o segundo é para importar as duplicatas ref. as faturas.
            with open("Arquivos gerados\ENERGextrator.txt", "w") as arquivo1:
                arquivo1.write("|0000|CÓDIGO EMPRESA|0|DATA INICIAL|DATA FINAL|NOME DA MINHA EMPRESA|CNPJ||UF|INSC. ESTADUAL|CÓD IBGE - MUNICIPIO|||A|1|\n|0001|0|\n|0005|NOME EMPRESA|CEP|RUA|NUM||BAIRRO|TEL|||\n|0100|NOME CONTADOR|CPF|CRC||CEP|RUA|NUM||BAIRRO|TEL||E-MAIL|CÓD IBGE - MUNICIPIO|\n|0150|F19527639000158|ENERGISA MINAS GERAIS - DISTRIBUIDORA DE ENERGIA|01058|19527639000158||1530560230000|3115300||AV MANOEL INACIO PEIXOTO|1200||DISTRITO INDUSTRIAL|\n|0990|6|\n|B001|1|\n|B990|2|\n|C001|0|\n")
            with open("Arquivos gerados\DuplENERG.txt", "w") as arquivo2:
                arquivo2.write("Numero da Nota; Data escrituracao; CPF/CNPJ do participante; Numero da Duplicata; Data de Vencimento da Duplicata; Valor da Duplicata; Observacao\n")
            for file in listdir(buscar): #aqui ele vai percorrer todos os pdf que estão na pasta selecionada.
                if search("pdf", file):
                    documento = path.join(buscar, file)
                    doc = fitz.open(documento)
                    pagina = doc[0]
                    texto = pagina.get_text() #vai extratir o conteudo dos pdf como texto.
                    arquivo = open("Arquivos gerados\ENERGextrator.txt", 'a')
                    duplicatas = open("Arquivos gerados\DuplENERG.txt", 'a')
                    # aqui abaixo começa a procurar pelas informações que precisa para preencher os campos do layout.
                    modelo_1 = texto.find("DANF3E") #faturas da ENERGISA tem dois layouts, então aqui ele testa se é o primeiro layout.
                    if modelo_1 != -1:
                        posiniNum = texto.find('NOTA FISCAL Nº:') + 16
                        posfinNum = texto.find('NOTA FISCAL Nº:') + 27
                        numero = texto[posiniNum:posfinNum]
                        numero = numero.replace('.', '')

                        posiniEmi = texto.find('EMISSÃO:') + 8
                        posfinEmi = texto.find('EMISSÃO:') + 18
                        data1 = texto[posiniEmi:posfinEmi]
                        data = data1.replace('/', '')

                        modelo = texto.find('CONSIDERAR ESTA NOTA FISCAL QUITADA SOMENTE APÓS O EFETIVO DÉBITO')
                        if modelo != -1:
                            posiniVenc = texto.find('VENCIMENTO') + 11
                            posfinVenc = texto.find('VENCIMENTO') + 21
                            venc = texto[posiniVenc:posfinVenc].strip()
                        else:
                            modelo1 = texto.find('EMITIDO EM CONTINGÊNCIA Pendente de Autorização')
                            if modelo1 != -1:
                                posiniVenc = texto.find('RICMS/MG - 2002') + 64
                                posfinVenc = texto.find('RICMS/MG - 2002') + 74
                                venc = texto[posiniVenc:posfinVenc].strip()
                            else:
                                posiniVenc = texto.find('RICMS/MG - 2002') + 16
                                posfinVenc = texto.find('RICMS/MG - 2002') + 26
                                venc = texto[posiniVenc:posfinVenc].strip()

                        posinichave = texto.find('EMISSÃO:') + 19
                        posfinchave = texto.find('', posinichave) + 54
                        chave = texto[posinichave:posfinchave]
                        chave = chave.replace(' ', '').replace('\n', '')

                        posiniVR = texto.find('% Alíq.') + 8
                        posfinVR = texto.find(',', posiniVR)
                        valorr = texto[posiniVR:posfinVR + 3]
                        vrilum = 0
                        iluminacao = texto.find('CONTRIBUICAO ILUM PUBLICA')
                        if iluminacao != -1:
                            posiniVRILU = texto.find(f'{valorr}', posfinVR)
                            posfinVRILU1 = texto.find(',', posiniVRILU - 4)
                            posiniVRILU1 = texto.find('', posfinVRILU1 - 2)
                            vrilum1 = texto[posiniVRILU1:posfinVRILU1 + 3]
                            vrilum = vrilum1
                            vrilum = vrilum.replace('.', '').replace(',', '.')
                            vrilum = float(vrilum)
                        if iluminacao == -1:
                            vrilum = 0
                        valorr = valorr.replace('.', '').replace(',', '.')
                        valor_base = float(valorr)

                        valor = (valor_base + vrilum)
                        valor = f'{valor:.2f}'
                        valor = str(valor)
                        valor = valor.replace('.', ',')

                        arquivo.write(f'|C500|0|1|F19527639000158|66|00|002||01|{numero}|{data}|{data}|{valor}|0|{valor}|0|0|0|0,00|0,00|0|0||0|0|||{chave}|||||52||||||||\n')
                        arquivo.write(f'|C590|090|2255|0|{valor}|0,00|0,00|0|0|0||\n')

                        duplicatas.write(f'"{numero}"; "{data1}"; "19527639000158"; "01"; "{venc}"; "{valor}"; ""\n')
                        duplicatas.close()
                    else: #>>>>>>>>>>>>>aqui é o segundo layout!<<<<<<<<<<<<<<<<<<<
                        num = texto.find('SÉRIE')
                        posfinnum = texto.find('\n', num - 6)
                        posininum = texto.find('\n', posfinnum - 14)
                        numero = texto[posininum:posfinnum].strip()

                        chv = texto.find('Chave de Acesso')
                        posinichv1 = texto.find('\n', chv)
                        posinichv = texto.find('\n', posinichv1 + 1)
                        posfinchv = texto.find('\n', posinichv + 1)
                        chave = texto[posinichv:posfinchv].strip()
                        chave = chave.replace(' ', '')

                        emi = texto.find('Valor')
                        posfinemi = texto.find('\n', emi - 3)
                        posiniemi = texto.find('\n', posfinemi - 13)
                        emissao1 = texto[posiniemi:posfinemi].strip()
                        emissao = emissao1.replace('/', '')

                        venc = texto.find('Faturamento pela média/mínimo')
                        posinivenc = texto.find('\n', venc)
                        posfinvenc = texto.find('\n', posinivenc + 1)
                        vencimento = texto[posinivenc:posfinvenc].strip()

                        base = texto.find('SÉRIE')
                        posinivr = texto.find('\n', base)
                        posfinvr = texto.find('\n', posinivr + 1)
                        vr_base = texto[posinivr:posfinvr].strip()
                        ilum = texto.find('ILUM PUBLICA')
                        if ilum != -1:
                            posiniilu1 = texto.find(f'{vr_base}', posfinvr + 1)
                            posiniilu2 = texto.find('\n', posiniilu1 + 1)
                            posfinilu = texto.find('\n', posiniilu2 + 1)
                            vr_ilum = texto[posiniilu2:posfinilu].strip()
                            vr_ilum = vr_ilum.replace(',', '.')
                            vr_ilum = float(vr_ilum)
                        else:
                            vr_ilum = 0

                        vr_base = vr_base.replace('.', '')
                        vr_base = vr_base.replace(',', '.')
                        vr_base = float(vr_base)

                        valor = (vr_base + vr_ilum)
                        valor = f'{valor:.2f}'
                        valor = str(valor)
                        valor = valor.replace('.', ',')
                        # aqui abaixo ele escreve as informações nos dois txt conforme o layout necessitado.
                        arquivo.write(f'|C500|0|1|F19527639000158|66|00|002||01|{numero}|{emissao}|{emissao}|{valor}|0|{valor}|0|0|0|0,00|0,00|0|0||0|0|||{chave}|||||52||||||||\n')
                        arquivo.write(f'|C590|090|2255|0|{valor}|0,00|0,00|0|0|0||\n')

                        duplicatas.write(f'"{numero}"; "{emissao1}"; "19527639000158"; "01"; "{vencimento}"; "{valor}"; ""\n')
                        duplicatas.close()

            arquivo = open("Arquivos gerados\ENERGextrator.txt", 'a')
            arquivo.write('\n')
            arquivo.close()

        except:
            showerror(title='ERRO', message='Erro ao extrair PDF!')
            return
        showinfo(title='Extração', message='Arquivo PDF extraído com sucesso!')

    def extrairCEMIG(self): #extrator das faturas de energia da fornecedora CEMIG.
        try:
            buscar = askdirectory() #serão escritos dois txt, o primeiro é o layout do SPED para importar as faturas para o sistema fiscal do Alterdata, o segundo é para importar as duplicatas ref. as faturas.
            with open("Arquivos gerados\CEMIGextrator.txt", "w") as arquivo1:
                arquivo1.write("|0000|CÓDIGO EMPRESA|0|DATA INICIAL|DATA FINAL|NOME DA MINHA EMPRESA|CNPJ||UF|INSC. ESTADUAL|CÓD IBGE - MUNICIPIO|||A|1|\n|0001|0|\n|0005|NOME EMPRESA|CEP|RUA|NUM||BAIRRO|TEL|||\n|0100|NOME CONTADOR|CPF|CRC||CEP|RUA|NUM||BAIRRO|TEL||E-MAIL|CÓD IBGE - MUNICIPIO|\n|0150|F06981180000116|CEMIG DISTRIBUICAO S.A|01058|06981180000116||0623221360087|3106200||AV BARBACENA|1200|17 ANDAR - ALA A1|SANTO AGOSTINHO|\n|0990|6|\n|B001|1|\n|B990|2|\n|C001|0|\n")
            with open("Arquivos gerados\DuplCEMIG.txt", "w") as arquivo2:
                arquivo2.write("Numero da Nota; Data escrituracao; CPF/CNPJ do participante; Numero da Duplicata; Data de Vencimento da Duplicata; Valor da Duplicata; Observacao\n")
            for file in listdir(buscar): #aqui ele vai percorrer todos os pdf que estão na pasta selecionada.
                if search("pdf", file):
                    documento = path.join(buscar, file)
                    doc = fitz.open(documento)
                    pagina = doc[0]
                    texto = pagina.get_text() #vai extratir o conteudo dos pdf como texto.
                    arquivo = open("Arquivos gerados\CEMIGextrator.txt", 'a')
                    duplicatas = open("Arquivos gerados\DuplCEMIG.txt", 'a')
                    #aqui abaixo começa a procurar pelas informações que precisa para preencher os campos do layout.
                    posiniNum = texto.find('NOTA FISCAL Nº ') + 15
                    posfinNum = texto.find('-', posiniNum) - 1
                    numero = texto[posiniNum:posfinNum]

                    posiniEmi = texto.find('Data de emissão:') + 17
                    posfinEmi = texto.find('Data de emissão:') + 27
                    data1 = texto[posiniEmi:posfinEmi]
                    data = data1.replace('/', '')

                    posiniVenc = texto.find('Vencimento')
                    posiniVenc1 = texto.find('/', posiniVenc) - 2
                    posfinVenc = texto.find('', posiniVenc1) + 10
                    venc = texto[posiniVenc1:posfinVenc]

                    custo = texto.find('Custo de Disponibilidade')
                    if custo != -1:
                        posiniCT = texto.find('Custo de Disponibilidade') + 34
                        posfinCT = texto.find(',', posiniCT)
                        valorr = texto[posiniCT:posfinCT + 3]
                    else:
                        posiniVR = texto.find('kWh') + 30
                        posfinVR = texto.find(',', posiniVR)
                        valorr = texto[posiniVR:posfinVR + 3].strip()
                    vrilum = 0
                    iluminacao = texto.find('Contrib Ilum Publica Municipal')
                    if iluminacao != -1:
                        posiniVRILU = texto.find('Contrib Ilum Publica Municipal') + 30
                        posfinVRILU = texto.find(',', posiniVRILU)
                        vrilum1 = texto[posiniVRILU:posfinVRILU + 3]
                        vrilum = vrilum1
                        vrilum = vrilum.replace('.', '').replace(',', '.')
                        vrilum = float(vrilum)
                    if iluminacao == -1:
                        vrilum = 0
                    valorr = valorr.replace('.', '').replace(',', '.')
                    valor_base = float(valorr)

                    bonus = texto.find('Bônus Itaipu')
                    if bonus != -1:
                        posiniDESC = texto.find(',', bonus) - 5
                        posfinDESC = texto.find(',', posiniDESC)
                        valord = texto[posiniDESC:posfinDESC + 3].strip()
                        valord = valord.replace('.', '').replace(',', '.').replace('-', '')
                        valor_desc = float(valord)
                    else:
                        valor_desc = 0

                    comp = texto.find('Compensação FIC')
                    if comp != -1:
                        posiniComp = texto.find('\n', comp) + 1
                        posfinComp = texto.find(',', posiniComp) + 3
                        vrcomp = texto[posiniComp:posfinComp]
                        vrcomp = vrcomp.replace('-', ' ').replace(',', '.').strip()
                        valor_comp = float(vrcomp)
                    else:
                        valor_comp = 0

                    posinichave = texto.find('chave de acesso:') + 17
                    posfinchave = texto.find('', posinichave) + 44
                    chave = texto[posinichave:posfinchave]

                    valor = (valor_base + vrilum - valor_desc - valor_comp)
                    valor = f'{valor:.2f}'
                    valor = str(valor)
                    valor = valor.replace('.', ',')
                    #aqui abaixo ele escreve as informações nos dois txt conforme o layout necessitado.
                    arquivo.write(f'|C500|0|1|F06981180000116|66|00|000||01|{numero}|{data}|{data}|{valor}|0|{valor}|0|0|0|0,00|0,00|0|0||0|0|||{chave}|||||52||||||||\n')
                    arquivo.write(f'|C590|090|2255|0|{valor}|0,00|0,00|0|0|0||\n')

                    duplicatas.write(f'"{numero}"; "{data1}"; "06981180000116"; "01"; "{venc}"; "{valor}"; ""\n')
                    duplicatas.close()

            arquivo = open("Arquivos gerados\CEMIGextrator.txt", 'a')
            arquivo.write('\n')
            arquivo.close()

        except:
            showerror(title='ERRO', message='Erro ao extrair PDF!')
            return
        showinfo(title='Extração', message='Arquivo PDF extraído com sucesso!')

    def close(self):
        self.quit()

    def gerar(self):
        try:
            empresa1 = ('Empresa teste 1', '99.999.999/9999-99', 'Rua A, nº 1, Fundos Sala, Centro', #dados da empres 1
                  'CEP: 00.000-000 Cidade/UF', 'emp1')
            empresa2 = ('Empresa teste 2', '88.888.888/8888-88', 'Rua B, nº 2, Térreo, Centro', #dados da empresa 2
                  'CEP: 00.000-000 Cidade/UF', 'emp2')
            socio1 = ('Sócio teste 1', '000.000.000-00', '', '', 'soc1') #dados sócio 1
            socio2 = ('Sócio teste 2', '999.999.999-99', '', '', 'soc2') #dados sócio 2
            adi = ('um adiantamento da distribuição de lucros.', 'retirada de parte dos lucros acumulados.')
            tipo = ('pessoa jurídica, portadora do C.N.P.J nº', ' pessoa física,  portador do  C.P.F nº ')
            mes = ''
            pessoa = ''
            pg = self.varfrm1.get() #vai buscar a opção do pagador marcada no radio button.
            rec = self.varfrm2.get() #vai buscar a opção do recebedor marcada no radio button.
            ad = self.varfrm3.get() #vai buscar a opção de adiantamento marcada no radio button.
            vr = float(self.valorr.get()) #vai buscar o valor digitado no entry.
            if self.data.get() == '': #caso nao informar a data no entry, ele seta a data como nenhum, para não da erro na busca e sim na criação do pdf.
                dt = None
            else:
                dt = self.data.get() #vai buscar o valor informado no entry para a data.
            extenso = num2words(vr, lang='pt_BR', to="currency") #aqui pega o valor e escreve por extenso.
            if pg == 'emp1' and rec == 'soc1': #as condições para atribuir os dados das listas de quem foi o pagador e recebedor.
                pg = empresa1
                rec = socio1
                pessoa = tipo[1]
            if pg == 'emp1' and rec == 'soc2':
                pg = empresa1
                rec = socio2
                pessoa = tipo[1]
            if pg == 'emp2' and rec == 'emp1':
                pg = empresa2
                rec = empresa1
                pessoa = tipo[0]
            if pg == 'emp2' and rec == 'soc1':
                pg = empresa2
                rec = socio1
                pessoa = tipo[1]
            if pg == 'emp2' and rec == 'soc2':
                pg = empresa2
                rec = socio2
                pessoa = tipo[1]
            if ad == 'S': #verificando a opção atribuida para adiantamento : SIM ou NÃO.
                ad = adi[0]
            if ad == 'N':
                ad = adi[1]
            if dt[2:4] == '01': #transformando o mês da data digitado com número para o nome.
                mes = 'Janeiro'
            if dt[2:4] == '02':
                mes = 'Fevereiro'
            if dt[2:4] == '03':
                mes = 'Março'
            if dt[2:4] == '04':
                mes = 'Abril'
            if dt[2:4] == '05':
                mes = 'Maio'
            if dt[2:4] == '06':
                mes = 'Junho'
            if dt[2:4] == '07':
                mes = 'Julho'
            if dt[2:4] == '08':
                mes = 'Agosto'
            if dt[2:4] == '09':
                mes = 'Setembro'
            if dt[2:4] == '10':
                mes = 'Outubro'
            if dt[2:4] == '11':
                mes = 'Novembro'
            if dt[2:4] == '12':
                mes = 'Dezembro'
            valor = f'{vr:_.2f}'
            valor = valor.replace('.', ',').replace('_', '.') #atribuindo virgula nas casas decimais e ponto para milhar.
            pdf = canvas.Canvas(f'./PDFs gerados/{pg[4]} x {rec[4]} - {dt[0:2]}-{dt[2:4]}-{dt[4:6]} - {valor}.pdf', pagesize=A4) #abrindo o arquivo pdf a ser escrito, o nome do PDF será de acordo com as informações coletadas: pagador x recebedor - data - valor.pdf
            if pg == empresa2 and rec == empresa1: #aqui um ponto para melhoria também, como não soube fazer para centralizar as escritas, considerando tamanho de nomes diferentes das empresas e sócios, criei um layout para cada condição.
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(195, 800, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(178, 785, f'{pg[2]}')
                pdf.drawString(231, 770, f'{pg[3]}')
                pdf.drawString(233, 755, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 700, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 670, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(50, 635, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 620, f'sócio-administrador da empresa {pg[0]}, DECLARA, para todos os')
                pdf.drawString(30, 605, f'fins legais e de direitos que recebeu nesta data da referida empresa, a quantia de R$ {valor}')
                pdf.drawString(30, 590, f'({extenso}) correspondente a {ad}')
                pdf.drawString(50, 575, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 550, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(200, 500, '_________________________')
                pdf.drawString(215, 485, f'{rec[0]}')
                pdf.drawString(228, 470, 'Sócio Administrador')
                pdf.drawString(218, 455, f'CPF: {rec[1]}')
                pdf.drawString(30, 415, '---------------------------------------------------------------------------------------------------------------------------------------')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(195, 375, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(178, 360, f'{pg[2]}')
                pdf.drawString(231, 345, f'{pg[3]}')
                pdf.drawString(233, 330, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 275, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 245, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(50, 210, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 195, f'sócio-administrador da empresa {pg[0]}, DECLARA, para todos os')
                pdf.drawString(30, 180, f'fins legais e de direitos que recebeu nesta data da referida empresa, a quantia de R$ {valor}')
                pdf.drawString(30, 165, f'({extenso}) correspondente a {ad}')
                pdf.drawString(50, 150, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 125, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(200, 75, '__________________________')
                pdf.drawString(215, 60, f'{rec[0]}')
                pdf.drawString(228, 45, 'Sócio Administrador')
                pdf.drawString(218, 30, f'CPF: {rec[1]}')
            if pg == empresa2 and rec == socio1:
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(195, 800, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(178, 785, f'{pg[2]}')
                pdf.drawString(231, 770, f'{pg[3]}')
                pdf.drawString(233, 755, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 700, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 670, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 650, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 635, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 620, f'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 605, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 590, f'correspondente a {ad}')
                pdf.drawString(60, 575, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 550, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 500, '_________________________')
                pdf.drawString(228, 485, f'{rec[0]}')
                pdf.drawString(235, 470, 'Sócio Administrador')
                pdf.drawString(233, 455, f'CPF: {rec[1]}')
                pdf.drawString(30, 415, '---------------------------------------------------------------------------------------------------------------------------------------')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(195, 375, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(178, 360, f'{pg[2]}')
                pdf.drawString(231, 345, f'{pg[3]}')
                pdf.drawString(233, 330, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 275, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 245, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 225, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 210, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 195, 'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 180, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 165, f'correspondente a {ad}')
                pdf.drawString(60, 150, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 125, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 75, '__________________________')
                pdf.drawString(228, 60, f'{rec[0]}')
                pdf.drawString(235, 45, 'Sócio Administrador')
                pdf.drawString(233, 30, f'CPF: {rec[1]}')
            if pg == empresa2 and rec == socio2:
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(195, 800, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(178, 785, f'{pg[2]}')
                pdf.drawString(231, 770, f'{pg[3]}')
                pdf.drawString(233, 755, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 700, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 670, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 650, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 635, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 620, f'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 605, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 590, f'correspondente a {ad}')
                pdf.drawString(60, 575, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 550, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 500, '________________________')
                pdf.drawString(232, 485, f'{rec[0]}')
                pdf.drawString(235, 470, 'Sócio Administrador')
                pdf.drawString(233, 455, f'CPF: {rec[1]}')
                pdf.drawString(30, 415, '---------------------------------------------------------------------------------------------------------------------------------------')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(195, 375, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(178, 360, f'{pg[2]}')
                pdf.drawString(231, 345, f'{pg[3]}')
                pdf.drawString(233, 330, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 275, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 245, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 225, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 210, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 195, 'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 180, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 165, f'correspondente a {ad}')
                pdf.drawString(60, 150, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 125, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 75, '_________________________')
                pdf.drawString(232, 60, f'{rec[0]}')
                pdf.drawString(235, 45, 'Sócio Administrador')
                pdf.drawString(233, 30, f'CPF: {rec[1]}')
            if pg == empresa1 and rec == socio1:
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(218, 800, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(185, 785, f'{pg[2]}')
                pdf.drawString(233, 770, f'{pg[3]}')
                pdf.drawString(236, 755, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 700, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 670, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 650, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 635, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 620, 'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 605, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 590, f'correspondente a {ad}')
                pdf.drawString(60, 575, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 550, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 500, '_________________________')
                pdf.drawString(228, 485, f'{rec[0]}')
                pdf.drawString(235, 470, 'Sócio Administrador')
                pdf.drawString(233, 455, f'CPF: {rec[1]}')
                pdf.drawString(30, 415, '---------------------------------------------------------------------------------------------------------------------------------------')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(218, 375, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(185, 360, f'{pg[2]}')
                pdf.drawString(233, 345, f'{pg[3]}')
                pdf.drawString(236, 330, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 275, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 245, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 225, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 210, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 195, 'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 180, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 165, f'correspondente a {ad}')
                pdf.drawString(60, 150, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 125, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 75, '__________________________')
                pdf.drawString(228, 60, f'{rec[0]}')
                pdf.drawString(235, 45, 'Sócio Administrador')
                pdf.drawString(233, 30, f'CPF: {rec[1]}')
            if pg == empresa1 and rec == socio2:
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(218, 800, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(185, 785, f'{pg[2]}')
                pdf.drawString(233, 770, f'{pg[3]}')
                pdf.drawString(236, 755, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 700, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 670, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 650, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 635, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 620, 'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 605, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 590, f'correspondente a {ad}')
                pdf.drawString(60, 575, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 550, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 500, '________________________')
                pdf.drawString(232, 485, f'{rec[0]}')
                pdf.drawString(235, 470, 'Sócio Administrador')
                pdf.drawString(233, 455, f'CPF: {rec[1]}')
                pdf.drawString(30, 415, '---------------------------------------------------------------------------------------------------------------------------------------')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(218, 375, f'{pg[0]}')
                pdf.setFont('Helvetica', 10)
                pdf.drawString(185, 360, f'{pg[2]}')
                pdf.drawString(233, 345, f'{pg[3]}')
                pdf.drawString(236, 330, f'CNPJ: {pg[1]}')
                pdf.setFont('Helvetica-Bold', 12)
                pdf.drawString(180, 275, 'RECIBO DE DISTRIBUIÇÃO DE LUCROS')
                pdf.drawString(470, 245, f'R$: {valor}')
                pdf.setFont('Helvetica', 12)
                pdf.drawString(60, 225, f'A, {rec[0]}, {pessoa} {rec[1]}')
                pdf.drawString(30, 210, f'sócio-administrador da empresa {pg[0]}, acima identificado, DECLARA')
                pdf.drawString(30, 195, 'para todos os fins legais e de direitos que recebeu nesta data da referida empresa a quantia de')
                pdf.drawString(30, 180, f'R$ {valor} ({extenso})')
                pdf.drawString(30, 165, f'correspondente a {ad}')
                pdf.drawString(60, 150, 'Para maior clareza e por ser expressão da verdade, firmo o presente recibo e assino.')
                pdf.drawString(30, 125, f'Cidade/UF, {dt[0:2]} de {mes} de 20{dt[4:6]}.')
                pdf.drawString(210, 75, '_________________________')
                pdf.drawString(232, 60, f'{rec[0]}')
                pdf.drawString(235, 45, 'Sócio Administrador')
                pdf.drawString(233, 30, f'CPF: {rec[1]}')
            pdf.save()
        except:
            showerror(title='ERRO', message='Erro ao criar arquivo PDF!')
            return
        showinfo(title='PDF Gerado', message='Arquivo PDF gerado com sucesso!')
        abrir = askquestion(title='Abrir', message='Deseja abrir o PDF gerado?')
        if abrir == 'yes':
            subprocess.call(f'./PDFs gerados/{pg[4]} x {rec[4]} - {dt[0:2]}-{dt[2:4]}-{dt[4:6]} - {valor}.pdf', shell=True)

    def pgtos(self):
        try:
            buscar = askopenfilename()
            data = self.dt.get()
            mes = self.ref.get()
            with open("Arquivos gerados\pgtos.txt", "w") as arquivo: #cria ou abre um txt chamado pgtos e escreve a primeira linha conforme abaixo:
                arquivo.write('"CodLancamentoAutomatico"; "ContaDebito"; "ContaCredito"; "DataLancamento"; "ValorLancamento"; "CodHistorico"; "ComplementoHistorico"\n')
            with fitz.open(buscar) as pdf: #vai abrir o explorer para selecionar a pasta onde está o pdf do relatório de pagamentos de funcionários
                for pagina in pdf:         #esse é um relatório específico emitido no sistema Alterdata, módulo do departamento pessoal.
                    texto = pagina.get_text()
                    arquivo = open("Arquivos gerados\pgtos.txt", 'a') #abre o txt pgtos para começar a escrever o conteudo conforme layout e informações abaixo:
                    ini = texto.find('RELATÓRIO DE LÍQUIDO GERAL') #após ele achar o titulo do relatorio, irá percorrer algumas quebras de linhas para achar o 1º nome.
                    if ini != -1:
                        iniNM = texto.find('\n', ini + 1)
                        meiNM = texto.find('\n', iniNM + 1)
                        finNM = texto.find('\n', meiNM + 1)
                        nome = texto[meiNM:finNM].strip().upper()
                        iniVR = texto.find('\n', finNM)
                        finVR = texto.find(',', iniVR)
                        valor = texto[iniVR:finVR + 3].strip()
                        arquivo.write(f'"00631"; ""; ""; "{data[0:2]}/{data[2:4]}/{data[4:6]}"; "{valor}"; "1001"; "{nome} REF. {mes[0:2]}/{mes[2:6]}"\n') #escrevendo no txt a 1ª informação.
                        for funcionarios in range(0, 42): #aqui ele vai percorrer os demais nomes em cada página (são 42 nomes após o 1º em cada página no relatório)
                            iniNM = texto.find('\n', finVR + 1) #vale acrescentar aqui que na ultima pagina do relatorio não chega a ter doso os 43 nomes, então ele pega até o ultimo e depois
                            meiNM = texto.find('\n', iniNM + 1) #começa a repetir os da primeira, ai eu entreno no txt e excluo os repetidos que é facil de visualizar. Esse é um dos pontos
                            finNM = texto.find('\n', meiNM + 1) # que precisa de melhorias.
                            nome = texto[meiNM:finNM].strip().upper()
                            iniVR = texto.find('\n', finNM)
                            finVR = texto.find(',', iniVR)
                            valor = texto[iniVR:finVR + 3].strip()
                            if nome == 'ELIAS RODRIGUES CAMPOS': #aqui é uma condição de uma nome específico pois ele é sócio e não funcionário, então o codigo do lançamento muda de 00631 para 00471.
                                arquivo.write(f'"00471"; ""; ""; "{data[0:2]}/{data[2:4]}/{data[4:6]}"; "{valor}"; "1001"; "{nome} REF. {mes[0:2]}/{mes[2:6]}"\n')
                            if nome == 'MAYCON JONES RODRIGUES TRINDADE': #aqui é uma condição de uma nome específico pois ele é autonomo e não funcionário, então o codigo do lançamento muda de 00631 para 00682.
                                arquivo.write(f'"00682"; ""; ""; "{data[0:2]}/{data[2:4]}/{data[4:6]}"; "{valor}"; "1001"; "{nome} REF. {mes[0:2]}/{mes[2:6]}"\n')
                            else: #e aqui a ultima condição, caso não for os nomes específicos acima, ele escreve como funcionário mesmo... codigo do lançamento 00631.
                                arquivo.write(f'"00631"; ""; ""; "{data[0:2]}/{data[2:4]}/{data[4:6]}"; "{valor}"; "1001"; "{nome} REF. {mes[0:2]}/{mes[2:6]}"\n')
                    else:
                        break
            arquivo.close()
        except:
            showerror(title='ERRO', message='Erro ao extrair PDF!')
            return
        showinfo(title='Extração', message='Arquivo PDF extraído com sucesso!')

if __name__ == "__main__":
    app = App()
    app.mainloop()
