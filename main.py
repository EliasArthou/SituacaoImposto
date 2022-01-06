import os
import web
import auxiliares as aux
import messagebox
import sys
import sensiveis as senha


caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                      tipoarquivos=[('Banco WMartins', '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                      caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')

if len(caminhobanco) == 0:
    messagebox.msgbox(f'Selecione o caminho do Banco de Dados!', messagebox.MB_OK, 'Erro Banco')
    sys.exit()

bd = aux.Banco(caminhobanco)

resultado = bd.consultar('SELECT * FROM [Lista Codigos IPTUs Completa]')
bd.fecharbanco()
valorano = messagebox.inputbox('Ano de Extração', 'Digite o Ano de Extração:')

if valorano is not None:
    if not valorano.isdigit():
        messagebox.msgbox(f'Digite um valor válido (precisa ser numérico)!', messagebox.MB_OK, 'Ano de Extração')
        sys.exit()
else:
    messagebox.msgbox(f'Digite o ano de extração!', messagebox.MB_OK, 'Ano de Extração')
    sys.exit()

resolveucaptcha = False
mensagemerro = None
pastadownload = aux.caminhoprojeto()+'\\'+'Downloads'
listachaves = ['Número Cliente', 'Arquivo', 'Incrição Imobiliária', 'Exercício', 'Número Guia', 'Tipo Guia',
               'Data de Emissão', 'Quantidade de Parcelas', 'Valor da Guia']
listaexcel = []
site = web.TratarSite(senha.site, senha.nomeprofile)
site.abrirnavegador()
dadosiptu = []
limiteserrosseguidos = 10
errosseguidos = 0


try:
    for linha in resultado:
        if not os.path.isfile(pastadownload + '\\' + linha['codigo'] + '_' + linha['iptu'] + '.pdf'):
            while not resolveucaptcha:
                deuerro = True
                codigocliente = linha['codigo']
                if site.url != senha.site:
                    site.fecharsite()
                    site = web.TratarSite(senha.site, senha.nomeprofile)
                    site.abrirnavegador()

                if site is not None and site.navegador != -1:
                    inscricao = site.verificarobjetoexiste('NAME', 'inscricao')
                    if inscricao is not None:
                        inscricao.clear()
                        inscricao.send_keys(linha['iptu'])
                        ano = site.verificarobjetoexiste('NAME', 'exercicio', valorano)
                        caminho = aux.caminhoprojeto() + '\\' + 'captcha.png'
                        achouimagem, salvouimagem = site.baixarimagem('XPATH', '//*[@id="img"]', caminho)
                        if achouimagem and salvouimagem:
                            resolveucaptcha = aux.resolvercaptcha(caminho, site, 'ID', 'texto_imagem', 'NAME', 'btenviar')
                            if resolveucaptcha:
                                botaogerar = site.verificarobjetoexiste('NAME', 'btConsultar')
                                if botaogerar is not None:
                                    botaogerar.click()
                                    site.esperadownloads(pastadownload, 10)
                                    baixado = aux.ultimoarquivo(pastadownload, '.pdf')
                                    if 'PCRJ_IPTU' not in baixado:
                                        baixado = ''

                                    if len(baixado) > 0:
                                        aux.renomeararquivo(baixado, pastadownload + '/' + codigocliente + '_' +
                                                            linha['iptu'] + '.pdf')
                                        listaitens = site.verificarobjetoexiste('CLASS_NAME', 'TDCentro',
                                                                                itemunico=False)
                                        textotratado = ''
                                        if listaitens is not None:
                                            dadosiptu.append(codigocliente)
                                            if os.path.isfile(pastadownload + '/' + codigocliente + '_'
                                                              + linha['iptu'] + '.pdf'):
                                                dadosiptu.append(pastadownload + '/' + codigocliente + '_'
                                                                 + linha['iptu'] + '.pdf')
                                                deuerro = False
                                            else:
                                                dadosiptu.append('')

                                            for item in listaitens:
                                                dadosiptu.append(item.text)

                                            listaexcel.append(dict(zip(listachaves, dadosiptu)))
                                            dadosiptu = []
                                    else:
                                        resolveucaptcha = False
                                        for arquivo in os.listdir(pastadownload):
                                            if 'PCRJ_IPTU' in arquivo:
                                                os.remove(arquivo)

                                        if site.num_abas() > 1:
                                            site.fecharsite()
                                            site = web.TratarSite(senha.site, senha.nomeprofile)
                                            site.abrirnavegador()
                                            achouimagem = False
                                            salvouimagem = False
                                            resolveucaptcha = False
                                            errosseguidos += 1
                                            if site.navegador == -1 or errosseguidos >= limiteserrosseguidos:
                                                break
                                            continue
                                else:
                                    mensagemerro = site.verificarobjetoexiste('XPATH', '/html/body/div/div/center/div/table[1]/tbody/tr[3]/td/font/b')
                                    if mensagemerro is not None:
                                        mensagemerro = None
                                        resolveucaptcha = False

                                botaopesquisa = site.verificarobjetoexiste('NAME', 'bt')
                                if botaopesquisa is not None:
                                    botaopesquisa.click()
                                    inscricao = None
                                    ano = None
                                    botaogerar = None
                                    botaopesquisa = None
                                    achouimagem = False
                                    salvouimagem = False
                                    codigocliente = ''

                                else:
                                    site.fecharsite()
                                    site = web.TratarSite(senha.site, senha.nomeprofile)
                                    site.abrirnavegador()
                                    achouimagem = False
                                    salvouimagem = False
                                    resolveucaptcha = False

                            else:
                                achouimagem = False
                                salvouimagem = False

                            if achouimagem and not salvouimagem:
                                site.fecharsite()
                                site = web.TratarSite(senha.site, senha.nomeprofile)
                                site.abrirnavegador()
                                achouimagem = False
                                salvouimagem = False
                                resolveucaptcha = False

                if deuerro:
                    errosseguidos += 1
                else:
                    errosseguidos = 0

                if site.navegador == -1 or errosseguidos >= limiteserrosseguidos:
                    break

            resolveucaptcha = False
        if site.navegador == -1 or errosseguidos >= limiteserrosseguidos:
            break
        #messagebox.msgbox(f'O tempo decorrido foi de: {"{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes),
        # int(seconds))}', messagebox.MB_OK, 'Tempo Decorrido')
finally:
    if site is not None:
        site.fecharsite()

    if errosseguidos >= limiteserrosseguidos:
        messagebox.msgbox(f'Site com instabilidade! Tentar novamente mais tarde!', messagebox.MB_OK, '%s erros seguidos' % limiteserrosseguidos)

    if len(listaexcel) > 0:
        aux.escreverlistaexcelog('Log_' + aux.acertardataatual() + '.xlsx', listaexcel)
