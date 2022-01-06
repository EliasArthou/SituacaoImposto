from selenium import webdriver
import io
from PIL import Image
import os
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from anticaptchaofficial.imagecaptcha import *
from selenium.webdriver.support.ui import Select
import auxiliares as aux
from bs4 import BeautifulSoup
import messagebox


class TratarSite:
    def __init__(self, url, nomeperfil):
        self.url = url
        self.perfil = nomeperfil
        self.navegador = None
        self.options = None

    def abrirnavegador(self):

        self.navegador = self.configuraprofilechrome()

        if self.navegador is not None:
            self.navegador.get(self.url)
            time.sleep(1)
            #Testa se a página carregou (ainda tem que fazer um teste e condição quando ele apresenta um texto de erro de carregamento)
            #==========================================================================================================================
            resultadolimpo = ''
            corposite = BeautifulSoup(self.navegador.page_source, 'html.parser')
            for string in corposite.strings:
                resultadolimpo = resultadolimpo + ' ' + string

            if len(resultadolimpo) == 0:
                messagebox.msgbox(f'Site com problema de carregamento!', messagebox.MB_OK, 'Site fora do ar')
                self.navegador = -1
            # ==========================================================================================================================
            time.sleep(1)
            return self.navegador

    def configuraprofilechrome(self):
        self.options = webdriver.ChromeOptions()
        if aux.caminhoprojeto('Profile') != '':
            self.options.add_argument("user-data-dir=" + aux.caminhoprojeto('Profile'))
            self.options.add_argument("--start-maximized")
            self.options.add_argument("---printing")
            self.options.add_argument("--disable-print-preview")
            #Forma invisível
            #self.options.add_argument("--headless")

            if aux.caminhoprojeto('Downloads') != '':
                self.options.add_experimental_option('prefs', {
                "profile.name": self.perfil,
                "download.default_directory": aux.caminhoprojeto('Downloads'),  #Change default directory for downloads
                "download.prompt_for_download": False,  #To auto download the file
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True  #It will not show PDF directly in chrome
                })
        return webdriver.Chrome(options=self.options)

    def verificarobjetoexiste(self, identificador, endereco, valorselecao='', itemunico=True):
        if self.navegador is not None:
            try:
                if itemunico:
                    if len(valorselecao) == 0:
                        elemento = self.navegador.find_element(getattr(By, identificador), endereco)
                    else:
                        elemento = Select(self.navegador.find_element(getattr(By, identificador), endereco))
                        elemento.select_by_value(valorselecao)
                else:
                    if len(valorselecao) == 0:
                        elemento = self.navegador.find_elements(getattr(By, identificador), endereco)
                    else:
                        elemento = Select(self.navegador.find_elements(getattr(By, identificador), endereco))
                        elemento.select_by_value(valorselecao)

                return elemento

            except NoSuchElementException:
                return None

    def baixarimagem(self, identificador, endereco, caminho):
        achouimagem = False
        salvouimagem = False
        if self.navegador is not None:
            if os.path.isfile(caminho):
                os.remove(caminho)
            elemento = self.verificarobjetoexiste(identificador, endereco)
            if elemento is not None:
                achouimagem = True
                image = elemento.screenshot_as_png
                imagestream = io.BytesIO(image)
                im = Image.open(imagestream)
                im.save(caminho)
                if os.path.isfile(caminho):
                    #Verifica se a imagem veio toda "preta" (quando o site não carrega o CAPTCHA)
                    if aux.quantidade_cores(caminho) > 1:
                        salvouimagem = True
                    else:
                        os.remove(caminho)
                        salvouimagem = False
                else:
                    salvouimagem = False

        return achouimagem, salvouimagem

    def esperadownloads(self, caminho, timeout, nfiles=None):
        """
        Wait for downloads to finish with a specified timeout.
        Args
        ----
        caminho : str
            The path to the folder where the files will be downloaded.
        timeout : int
            How many seconds to wait until timing out.
        nfiles : int, defaults to None
            If provided, also wait for the expected number of files.
        """
        time.sleep(1)
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < timeout:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(caminho)
            if nfiles and len(files) != nfiles:
                dl_wait = True

            for fname in files:
                if fname.endswith('.crdownload'):
                    dl_wait = True

            seconds += 1

        time.sleep(1)
        return seconds

    def num_abas(self):
        if self.navegador is not None:
            return len(self.navegador.window_handles)

    def fecharsite(self):
        if self.navegador is not None and hasattr(self.navegador, 'quit'):
            self.navegador.quit()



