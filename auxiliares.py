import os
import time
import pypyodbc as pyodbc
import sensiveis as senhas

import sensiveis


def caminhoprojeto(subpasta=''):
    caminho = os.path.dirname(os.path.abspath(__file__))
    if len(subpasta) > 0:
        if os.path.isdir(caminho+'\\'+subpasta):
            return caminho + '\\' + subpasta
        else:
            return ''
    else:
        if os.path.isdir(caminho):
            return caminho
        else:
            return ''


def ultimoarquivo(caminho, extensao):
    lista_arquivos = os.listdir(caminho)
    ultimadata = 0
    ultimoatualizado = ''
    for arquivo in lista_arquivos:
        if extensao.upper() in arquivo.upper():
            if os.path.getmtime(caminho+'/'+arquivo) > ultimadata:
                ultimoatualizado = caminho+'/'+arquivo

    return ultimoatualizado


def renomeararquivo(nomeantigo, novonome):
    if os.path.isfile(to_raw(novonome)):
        os.remove(to_raw(novonome))
    time.sleep(0.5)
    os.rename(to_raw(nomeantigo), to_raw(novonome))


def to_raw(string):
    return fr"{string}"


class Banco:
    def __init__(self, caminho):
        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=%s;Pwd=%s' % caminho % senhas.senhabanco
        print(constr)
        conxn = pyodbc.connect(constr)
        self.cursor = conxn.cursor()

    def consultar(self, sql):
        self.cursor.execute(sql)
        resultado = self.cursor.fetchall()
        return resultado

    def fecharbanco(self):
        self.cursor.close()


def quantidade_cores(caminho):
    from PIL import Image
    #from matplotlib.pyplot import imshow

    if os.path.isfile(caminho):
        image = Image.open(caminho)
        #Corta a imagem pra tirar as bordas inúteis
        image = image.crop((1, 0, 130, 24))
        #imshow(image)
        #Pega as cores e a quantidade de pixels (pelo que entendi) da mesma cor e coloca em uma lista
        #Calcula o número de pixels na tela (quant pixels eixo "x" * quant pixels eixo "y") para retornar
        #o máximo de cores possíveis (pior caso cada pixel de uma cor diferente), se tiver mais cores do que o definido
        #ao chamar o getcolors retorna None
        cores = image.convert('RGB').getcolors(image.size[0]*image.size[1])
        if cores is not None:
            #Retorna a quantidade de elementos da lista de cores
            return len(cores)
        else:
            return 0


def retornorvalorlista(lista, indice):
    import numpy as np

    lst = np.array(lista)
    return lst[indice]


def escreverlistaexcelog(caminho, lista):
    import pandas as pd

    dicionario = pd.DataFrame.from_dict(lista, orient='columns')
    #writer = dicionario.ExcelWriter(caminho, engine='xlsxwriter')
    writer = pd.ExcelWriter(path=caminho, engine='xlsxwriter')
    dicionario.to_excel(writer, sheet_name='LOG', index=False)
    #dicionario.to_excel(writer, 'Teste')
    writer.save()


def acertardataatual():
    from datetime import datetime

    textodata = datetime.now()
    return textodata.strftime('%Y_%m_%d_%H_%M_%S')


def caminhospadroes(caminho):
    import ctypes.wintypes
    # CSIDL	                        Decimal	Hex	    Shell	Description
    # CSIDL_ADMINTOOLS	            48	    0x30	5.0	    The file system directory that is used to store administrative tools for an individual user.
    # CSIDL_ALTSTARTUP	            29	    0x1D	 	    The file system directory that corresponds to the user's nonlocalized Startup program group.
    # CSIDL_APPDATA	                26	    0x1A	4.71	The file system directory that serves as a common repository for application-specific data.
    # CSIDL_BITBUCKET	            10	    0x0A	 	    The virtual folder containing the objects in the user's Recycle Bin.
    # CSIDL_CDBURN_AREA	            59	    0x3B	6.0	    The file system directory acting as a staging area for files waiting to be written to CD.
    # CSIDL_COMMON_ADMINTOOLS	    47	    0x2F	5.0	    The file system directory containing administrative tools for all users of the computer.
    # CSIDL_COMMON_ALTSTARTUP	    30	    0x1E	        NT-based only	The file system directory that corresponds to the nonlocalized Startup program group for all users.
    # CSIDL_COMMON_APPDATA	        35	    0x23	5.0	    The file system directory containing application data for all users.
    # CSIDL_COMMON_DESKTOPDIRECTORY	25	    0x19	        NT-based only	The file system directory that contains files and folders that appear on the desktop for all users.
    # CSIDL_COMMON_DOCUMENTS	    46	    0x2E	 	    The file system directory that contains documents that are common to all users.
    # CSIDL_COMMON_FAVORITES	    31	    0x1F	        NT-based only	The file system directory that serves as a common repository for favorite items common to all users.
    # CSIDL_COMMON_MUSIC	        53	    0x35	6.0	    The file system directory that serves as a repository for music files common to all users.
    # CSIDL_COMMON_PICTURES	        54	    0x36	6.0	    The file system directory that serves as a repository for image files common to all users.
    # CSIDL_COMMON_PROGRAMS	        23	    0x17	        NT-based only	The file system directory that contains the directories for the common program groups that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTMENU	    22	    0x16	        NT-based only	The file system directory that contains the programs and folders that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTUP	        24	    0x18	        NT-based only	The file system directory that contains the programs that appear in the Startup folder for all users.
    # CSIDL_COMMON_TEMPLATES	    45	    0x2D	        NT-based only	The file system directory that contains the templates that are available to all users.
    # CSIDL_COMMON_VIDEO	        55	    0x37	6.0	    The file system directory that serves as a repository for video files common to all users.
    # CSIDL_COMPUTERSNEARME	        61	    0x3D	6.0	    The folder representing other machines in your workgroup.
    # CSIDL_CONNECTIONS	            49	    0x31	6.0	    The virtual folder representing Network Connections, containing network and dial-up connections.
    # CSIDL_CONTROLS	            3	    0x03	 	    The virtual folder containing icons for the Control Panel applications.
    # CSIDL_COOKIES	                33	    0x21	 	    The file system directory that serves as a common repository for Internet cookies.
    # CSIDL_DESKTOP	                0	    0x00	 	    The virtual folder representing the Windows desktop, the root of the shell namespace.
    # CSIDL_DESKTOPDIRECTORY	    16	    0x10	 	    The file system directory used to physically store file objects on the desktop.
    # CSIDL_DRIVES	                17	    0x11	 	    The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
    # CSIDL_FAVORITES	            6	    0x06	 	    The file system directory that serves as a common repository for the user's favorite items.
    # CSIDL_FONTS	                20	    0x14	 	    A virtual folder containing fonts.
    # CSIDL_HISTORY	                34	    0x22	 	    The file system directory that serves as a common repository for Internet history items.
    # CSIDL_INTERNET	            1	    0x01	 	    A viritual folder for Internet Explorer.
    # CSIDL_INTERNET_CACHE	        32	    0x20	4.72	The file system directory that serves as a common repository for temporary Internet files.
    # CSIDL_LOCAL_APPDATA	        28	    0x1C	5.0	    The file system directory that serves as a data repository for local (nonroaming) applications.
    # CSIDL_MYDOCUMENTS	            5	    0x05	6.0	    The virtual folder representing the My Documents desktop item.
    # CSIDL_MYMUSIC	                13	    0x0D	6.0	    The file system directory that serves as a common repository for music files.
    # CSIDL_MYPICTURES	            39	    0x27	5.0	    The file system directory that serves as a common repository for image files.
    # CSIDL_MYVIDEO	                14	    0x0E	6.0	    The file system directory that serves as a common repository for video files.
    # CSIDL_NETHOOD	                19	    0x13	 	    A file system directory containing the link objects that may exist in the My Network Places virtual folder.
    # CSIDL_NETWORK	                18	    0x12	 	    A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
    # CSIDL_PERSONAL	            5	    0x05	 	    The file system directory used to physically store a user's common repository of documents. (From shell version 6.0 onwards, CSIDL_PERSONAL is equivalent to CSIDL_MYDOCUMENTS, which is a virtual folder.)
    # CSIDL_PHOTOALBUMS	            69	    0x45	Vista	The virtual folder used to store photo albums.
    # CSIDL_PLAYLISTS	            63	    0x3F	Vista	The virtual folder used to store play albums.
    # CSIDL_PRINTERS	            4	    0x04	 	    The virtual folder containing installed printers.
    # CSIDL_PRINTHOOD	            27	    0x1B	 	    The file system directory that contains the link objects that can exist in the Printers virtual folder.
    # CSIDL_PROFILE	                40	    0x28	5.0	    The user's profile folder.
    # CSIDL_PROGRAM_FILES	        38	    0x26	5.0	    The Program Files folder.
    # CSIDL_PROGRAM_FILESX86	    42	    0x2A	5.0	    The Program Files folder for 32-bit programs on 64-bit systems.
    # CSIDL_PROGRAM_FILES_COMMON	43	    0x2B	5.0	    A folder for components that are shared across applications.
    # CSIDL_PROGRAM_FILES_COMMONX86	44	    0x2C	5.0	    A folder for 32-bit components that are shared across applications on 64-bit systems.
    # CSIDL_PROGRAMS	            2	    0x02	 	    The file system directory that contains the user's program groups (which are themselves file system directories).
    # CSIDL_RECENT	                8	    0x08	 	    The file system directory that contains shortcuts to the user's most recently used documents.
    # CSIDL_RESOURCES	            56	    0x38	6.0	    The file system directory that contains resource data.
    # CSIDL_RESOURCES_LOCALIZED	    57	    0x39	6.0	    The file system directory that contains localized resource data.
    # CSIDL_SAMPLE_MUSIC	        64	    0x40	Vista	The file system directory that contains sample music.
    # CSIDL_SAMPLE_PLAYLISTS	    65	    0x41	Vista	The file system directory that contains sample playlists.
    # CSIDL_SAMPLE_PICTURES	        66	    0x42	Vista	The file system directory that contains sample pictures.
    # CSIDL_SAMPLE_VIDEOS	        67	    0x43	Vista	The file system directory that contains sample videos.
    # CSIDL_SENDTO	                9	    0x09	 	    The file system directory that contains Send To menu items.
    # CSIDL_STARTMENU	            11	    0x0B	 	    The file system directory containing Start menu items.
    # CSIDL_STARTUP	                7	    0x07	 	    The file system directory that corresponds to the user's Startup program group.
    # CSIDL_SYSTEM	                37	    0x25	5.0	    The Windows System folder.
    # CSIDL_SYSTEMX86	            41	    0x29	5.0	    The Windows 32-bit System folder on 64-bit systems.
    # CSIDL_TEMPLATES	            21	    0x15	 	    The file system directory that serves as a common repository for document templates.
    # CSIDL_WINDOWS	                36	    0x24	5.0	    The Windows directory or SYSROOT.

    csidl_personal = caminho  # Caminho padrão
    shgfp_type_current = 0  # Para não pegar a pasta padrão e sim a definida como documentos

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.Shell32.SHGetFolderPathW(None, csidl_personal, None, shgfp_type_current, buf)

    return buf.value


def caminhoselecionado(tipojanela=1, titulojanela='Selecione o caminho/arquivo:',
                       tipoarquivos=('Todos os Arquivos', '*.*'), caminhoini=caminhospadroes(5), arquivoinicial=''):
    import tkinter as tk
    from tkinter import filedialog

    'Cria a janela raiz'
    root = tk.Tk()
    root.withdraw()

    if tipojanela == 1:
        retorno = filedialog.askopenfilename(title=titulojanela,
                                             initialdir=caminhoini, filetypes=tipoarquivos, initialfile=arquivoinicial)
        if retorno is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None

    elif tipojanela == 2:
        name = filedialog.asksaveasfile(mode='w', defaultextension='.txt',
                                        filetypes=tipoarquivos,
                                        initialdir=caminhoini,
                                        title=titulojanela, initialfile=arquivoinicial)
        if name is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None
        text2save = str(name.name)
        name.write('')
        retorno = text2save

    elif tipojanela == 3:
        name = filedialog.askdirectory(initialdir=caminhoini, title=titulojanela)
        if name is None:  # askdirectory return `None` if dialog closed with "cancel".
            return None
        text2save = name
        retorno = text2save

    else:
        return

    return retorno


def resolvercaptcha(caminho, site, identificacaocaixa, caixacaptcha, identicacaobotao, botao):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as ec
    from selenium.common.exceptions import TimeoutException

    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key()

    captcha_text = solver.solve_and_return_solution(caminho)

    if len(str(captcha_text)) != 0:
        captcha = site.verificarobjetoexiste(identificacaocaixa, caixacaptcha)
        captcha.send_keys(captcha_text)
        confirmar = site.verificarobjetoexiste(identicacaobotao, botao)
        confirmar.click()
        try:
            alert = WebDriverWait(site.navegador, 2).until(ec.alert_is_present())
            texto = alert.text
            texto = texto.strip()
            alert.accept()
            if texto == 'CÓDIGO digitado Inválido!':
                solver.report_incorrect_image_captcha()
                resposta = False
            else:
                resposta = True
                if os.path.isfile(caminho):
                    os.remove(caminho)

        except TimeoutException:
            resposta = True
            if os.path.isfile(caminho):
                os.remove(caminho)

    else:
        print("Erro solução captcha:" + solver.error_code)
        resposta = False

    return resposta