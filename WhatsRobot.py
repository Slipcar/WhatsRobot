from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pymongo import MongoClient, errors
from urllib.parse import quote, unquote
from datetime import datetime
import qrcode
import PySimpleGUI as sg
import openpyxl as excel
import PIL.Image
import pyautogui
import pyperclip
import base64
import time
import cv2
import io

nome = ''
imagem = None
naoLocalizado = []
grupoNome = ''
dataHora = datetime.now()
dataHoraTexto = dataHora.strftime(r'%d-%m-%Y %H-%M')
options = webdriver.ChromeOptions()
# options.set_headless()
driver = webdriver.Chrome(options=options, executable_path=r'./driver/chromedriver.exe')
driver.get('https://web.whatsapp.com/')

class TelaRobo:
    def __init__(self):
        # layout
        layout = [
            [sg.Text('Usuário', size=(6,1)),sg.Input(key='usuario', size=(17,1))],
            [sg.Text('Senha', size=(6,1)),sg.Input(key='senha', password_char='*', size=(17,1))],
            [sg.Button('Entrar', size=(10,1), bind_return_key=True), sg.Button('Sair', size=(10,1))]
        ]

        self.janela1 = sg.Window('Login', icon="Logo.ico").layout(layout)
        self.janela2 = False

    # Extrair dados da janela
    def Iniciar(self):
        while True:
            try:          
                self.event, self.values = self.janela1.Read()
                self.usuario = self.values['usuario']
                self.senha = self.values['senha']
                # print(self.event, self.values)
            except:
                break

            if self.event == sg.WIN_CLOSED or self.event == 'Exit' or self.event == 'Sair':
                self.janela1.hide()
                driver.quit()
                break
            if self.event == 'Entrar':
                if not self.usuario or not self.senha:
                    sg.popup('Atenção!','Preencha os campos em branco e tente novamente.', icon="Logo.ico")
                else:
                    try:
                        liberar = consultaUsuario(self.usuario, self.senha)
                        if liberar == 'sim':
                            self.janela1.close()
                            self.janela2 = True

                            if self.janela2:
                                file_list_column = [
                                    [sg.Text("Carregar Contatos Excel", pad=((3,255),0)), sg.Text('Olá,', pad=(0,10)), sg.Text(nome.split(' ')[0], text_color='black', pad=(1,0), key='user')],
                                    [sg.In(key='excel', size=(55,1), enable_events=True), sg.FileBrowse('Carregar Excel')],
                                    [sg.Text('Mensagem para enviar')],
                                    [sg.Multiline(size=(68, 10), key='codMsg',
                                            tooltip='Mensagem para enviar')],
                                    [sg.Listbox(
                                        values=[], enable_events=True, size=(30, 15), key="contatos",
                                    ),  sg.Image(filename='qrcode.png', size=(256,256),key="QRImage"), 
                                    ],
                                    [sg.Button('Iniciar', size=(28, 1), key="mensagem", 
                                    disabled=False),sg.Button('Recarregar QRCode', size=(31,1))]
                                    # , sg.Button('Carregar WhatsApp', (28, 1), key="Wapp")
                                ]

                                image_viewer_column = [
                                    [sg.Text("Escolha uma imagem para enviar")],
                                    [sg.In(key='imagem-carregar', enable_events=True),
                                    sg.FileBrowse('Carregar Imagem')],
                                    # [sg.Checkbox('Enviar Imagem', key='enviar-imagem')],
                                    [sg.Image(filename=None, size=(400, 390),
                                            enable_events=True, key="imagem")],
                                ]

                                layout = [
                                    [
                                        sg.Column(file_list_column),
                                        sg.VSeparator(),
                                        sg.Column(image_viewer_column, element_justification='c')
                                    ]
                                ]
                                # janela
                                self.janela2 = sg.Window("Whats Robot", layout, icon="Logo.ico")

                                while True:
                                    try:
                                        self.event, self.values = self.janela2.Read()
                                        self.excel = self.values['excel']
                                        self.imagem = self.values['imagem-carregar']
                                        # self.enviarImg = self.values['enviar-imagem']
                                        self.codMsg = self.values['codMsg']
                                        
                                        # print(self.event, self.values, self.encodeMsg)
                                    except:
                                        pass
                                    if self.event == sg.WIN_CLOSED or self.event == 'Exit':
                                        driver.quit()
                                        break  
                                    if self.event == "excel":
                                        try:
                                            # print(self.excel)
                                            contatos = lerXLS(self.excel)
                                            self.janela2['contatos'].update(contatos)
                                        except:
                                            sg.popup("Arquivo inválido, tente novamente.", icon='logo.ico')
                                            pass
                                    if self.event == "imagem-carregar":
                                        global imagem
                                        imagem = self.imagem

                                        try:
                                            self.janela2["imagem"].update(
                                                data=convert_to_bytes(self.imagem, resize=(400, 400))
                                            )
                                        except Exception as e:
                                            sg.popup('Erro ao carregar Imagem', icon="Logo.ico")
                                    if self.event == "Recarregar QRCode":
                                        verificarQRCode(self.janela2)
                                    if self.event == "mensagem":
                                        if driver: 
                                            self.janela2.hide()
                                            popUp = sg.popup_yes_no('WhatsRobot', 'WhatsApp Está aberto e carregado?', icon="Logo.ico")
                                            if popUp == "Sim":
                                                try:
                                                    if "user":
                                                        enviarMensagem(contatos,self.codMsg, self.imagem)
                                                        escreverXLS(naoLocalizado)
                                                        break
                                                    else:
                                                        sg.popup('Atenção!','Não existe usuário logado.', icon='logo.ico')
                                                        break
                                                except Exception as e:
                                                    sg.popup('Erro!','Erro ao carregar contatos.', e, icon="Logo.ico")
                                            else:
                                                driver.refresh()
                                                time.sleep(0.5)
                                                self.janela2.UnHide()
                                                verificarQRCode(self.janela2)
                        else:    
                            popup = sg.popup('Aviso!','Usuário não autorizado, entre em contato com o suporte.', icon="Logo.ico")       
                    except errors.ServerSelectionTimeoutError as e:
                        sg.popup('Erro de Usuário','Não foi possível verificar o usuário.', e, icon="Logo.ico")


def capturarQRCode():
    try:
        imgQR = driver.find_element_by_xpath("//div[contains(@class,'_1yHR2')]")
        imgQR2 = imgQR.get_attribute('data-ref')
        qrCode2 = gerarQRCode(imgQR2)
    except:
        pass


def gerarQRCode(QrCode):
    try:
        arquivo = 'qrcode.png'
        img = qrcode.make(QrCode)
        img2 = img.resize((256,256))
        img2.save(arquivo)
    except:
        sg.popup('Erro ao gerar QRCode', icon='Logo.icon')


def verificarQRCode(janela):
    try:
        refresh = driver.find_element_by_xpath("//span[contains(@data-icon,'refresh-large')]")
        refresh.click()
        time.sleep(0.5)
        capturarQRCode()  
    except: 
        capturarQRCode()  
        try:
            janela['QRImage'].update(
                data=convert_to_bytes('qrcode.png', resize=(256,256)))
        except:      
            sg.popup('Erro ao recarregar QRCode', icon='Logo.ico')


def lerXLS(arquivoExcel):
    arquivo = arquivoExcel
    lista = []
    file = excel.load_workbook(arquivo)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        # print(contact)
        lista.append(contact)
    return lista


def escreverXLS(listaNomes):
    file = excel.Workbook()
    cont = 1
    for cell in listaNomes:
        sheet = file.active
        col = 'A'
        comp = col + str(cont)
        sheet[comp] = cell
        # print(sheet[comp])
        # print(cell)
        cont += 1
    file.save("NaoLocalizados_" + dataHoraTexto + ".xlsx")


def envia_media(fileToSend):
        """ Envia media """
        try:
            # Clica no botão adicionar
            # driver.find_element_by_xpath("//span[contains(@data-icon,'clip')]").click()
            driver.find_element_by_css_selector("span[data-icon='clip']").click()
            # Seleciona input
            attach = driver.find_element_by_css_selector("input[type='file']")
            # Adiciona arquivo
            attach.send_keys(fileToSend)
            time.sleep(2)
            # Seleciona botão enviar
            # send = driver.find_element_by_xpath("//*[@id='app']/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div")
            # Clica no botão enviar
            # send.click()
        except Exception as e:
            print("Erro ao enviar media", e)


# def copiarImagem():
#     try:
#         my_img = cv2.imread(imagem, cv2.IMREAD_ANYCOLOR)
#         # print(my_img)
#         cv2.imshow("My image", my_img)
#         pyautogui.hotkey('ctrl', 'c')
#         cv2.waitKey(0)
#         cv2.destroyAllWindows()
#     except Exception as e:
#         sg.popup('Erro ao carregar imagem','Formato de imagens suportadas: .png, .jpg ou .jpeg',icon='Logo.icon')


def enviarMensagem(grupos, mensagem, bool):
    sg.popup_animated(
        'hand.gif', 'Aguarde a conclusão do processo de envio.')
    whats = pyautogui.locateOnScreen('whats.png', confidence = 0.9)
    pyautogui.click(whats) 
    for grupo in grupos:
        try:   
            grupoNome = grupo
            firstName = grupoNome.split(" ")[0]
            # Iniciando pesquisa...
            grupo = driver.find_element_by_xpath("//div[contains(@class,'copyable-text selectable-text')]")
            time.sleep(1)
            grupo.clear()
            grupo.click()
            time.sleep(1)
            grupo.send_keys(grupoNome)
            time.sleep(0.5)
            # Verificando se o contato existe...
            grupo = driver.find_element_by_xpath(
                f"//span[@title='{grupoNome}']")
            time.sleep(2)
            grupo.click()
            chatBox = driver.find_elements_by_xpath(
                "//div[contains(@class,'copyable-text selectable-text')]")
            time.sleep(2)
            chatBox[1].click()

            if bool:
                pyperclip.copy(mensagem)
                pyperclip.paste()
                chatBox[1].send_keys(f"Olá *{firstName}*," + Keys.SHIFT + Keys.ENTER)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                envia_media(imagem)
                time.sleep(0.5)
                pyautogui.press('enter')
                # botao_enviar = driver.find_element_by_xpath("//span[contains(@data-icon,'send')]")
                # botao_enviar.click()
            else:
                pyperclip.copy(mensagem)
                pyperclip.paste()
                chatBox[1].send_keys(f"Olá *{firstName}*," + Keys.SHIFT + Keys.ENTER)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                pyautogui.press('enter')
                # botao_enviar = driver.find_element_by_xpath("//span[contains(@data-icon,'send')]")
                # botao_enviar.click()
        except:
            naoLocalizado.append(grupoNome)
            time.sleep(2)
            # print("Contato não localizado: " + grupoNome)
            pass
        # else:
        #     print("Envio Feito para: " + grupoNome)

    sg.popup_animated(image_source=None)
    time.sleep(5)
    # print(driver)
    driver.quit()
    sg.popup('Envio Completo!',
             icon='logo.ico', no_titlebar=True, image='ok.png', grab_anywhere=True)


def convert_to_bytes(file_or_bytes, resize=None):
    '''
    Will convert into bytes and optionally resize an image that is a file or a base64 bytes object.
    Turns into  PNG format in the process so that can be displayed by tkinter
    :param file_or_bytes: either a string filename or a bytes base64 image object
    :type file_or_bytes:  (Union[str, bytes])
    :param resize:  optional new size
    :type resize: (Tuple[int, int] or None)
    :return: (bytes) a byte-string object
    :rtype: (bytes)
    '''
    if isinstance(file_or_bytes, str):
        img = PIL.Image.open(file_or_bytes)
    else:
        try:
            img = PIL.Image.open(io.BytesIO(base64.b64decode(file_or_bytes)))
        except:
            dataBytesIO = io.BytesIO(file_or_bytes)
            img = PIL.Image.open(dataBytesIO)

    cur_width, cur_height = img.size
    if resize:
        new_width, new_height = resize
        scale = min(new_height/cur_height, new_width/cur_width)
        img = img.resize(
            (int(cur_width*scale), int(cur_height*scale)), PIL.Image.ANTIALIAS)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    del img
    return bio.getvalue()

def consultaUsuario(usuario, senha):
    global nome
    url = "mongodb://Clientes:Clientes2k21@clusterslip-shard-00-00.4ykgn.gcp.mongodb.net:27017,clusterslip-shard-00-01.4ykgn.gcp.mongodb.net:27017,clusterslip-shard-00-02.4ykgn.gcp.mongodb.net:27017/login?ssl=true&replicaSet=ClusterSlip-shard-0&authSource=admin&retryWrites=true&w=majority"
    try:
        client = MongoClient(url)

        db = client['login']
        coll = db['usuario']

        query = {"user":"{}".format(usuario), "senha":"{}".format(senha)}
        result = coll.find_one(query)
        ativo = result.get('ativo')
        nome = result.get('nome')

        return ativo
    except Exception as e:
        pass
        # sg.popup('Erro','Erro ao conectar com o banco entre em contato com o suporte.', e, icon="Logo.ico")

tela = TelaRobo()
tela.Iniciar()