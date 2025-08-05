import lista_des_dir
from updateBlock import update_block
import datetime
import win32com.client
import os
import time

def revisar_carimbo(folder_in, dados_revisao):

    t1 = f'{datetime.datetime.now():%X}'

    hora = datetime.datetime.now().strftime('%H h-%M m')
    data = datetime.datetime.now().strftime('%d-%m-%Y')

    path = r'L:\IFR - Infraestrutura\04_PROJETOS INFRA\PROGRAMA EP\Automatização de Documentos\Revisão de Carimbo\SAIDA'

    path = os.path.join(path, data)
    path = os.path.join(path, hora)

    if not os.path.exists(path):
        os.makedirs(path)

    folder_out = path
    
    lista_desenhos = []

    app_cad = dados_revisao['app_cad']

    log_saida = ''

    log_erro = ''

    lista_erros = []

    app = None

    delay = 0

    if app_cad == 'zwcad':

        try:

            app = win32com.client.dynamic.Dispatch("ZWCAD.Application")
            delay = 0
            app.Visible = True

        except Exception as e:

            log_saida += str(e)


    elif app_cad =='autocad':

        try:

            app = win32com.client.dynamic.Dispatch("AutoCAD.Application")
            delay = 1
            app.Visible = True
            time.sleep(3)


        except Exception as e:

            log_saida += str(e)

    if app == None:

        return 'APLICATIVO CAD NÃO ENCONTRADO'
    
    try:

        if len(app.Documents) != 0:
            if app.Documents[0].WindowTitle == 'Drawing1.dwg':
                app.Documents[0].Close()
            else:
                app = None
                return 'FECHAR DOCUMENTOS ABERTOS ANTES DE RODAR O PROGRAMA'
            
        
            
    except Exception as e:

        log_saida += str(e)

    lista_desenhos = lista_des_dir.list_dir_func(folder_in)

    n_desenhos = len(lista_desenhos)

    lista_desenhos_bloco_nencontrado = []

    n_bloco  = dados_revisao['n_bloco']

    args_revisao  = [

    dados_revisao['REV'],

    dados_revisao['T.E.'],

    dados_revisao['DESCRIÇÃO'],

    dados_revisao['PROJ'],

    dados_revisao['DES.'],

    dados_revisao['VER.'],

    dados_revisao['A.A.R.'],

    dados_revisao['AUT.'],

    dados_revisao['DATA'],

    dados_revisao['padrao_bloco']]

    n_bloco = dados_revisao['n_bloco']

    cont_err_cham = 0

    space = 'PaperSpace'

    while len(lista_desenhos) != 0:

        if cont_err_cham > 20:

            print('LIMITE DE ERROS ATINGIDO, DESENHOS RESTANTES:')

            log_erro = log_erro + '<br> LIMITE DE ERROS ATINGIDO, DESENHOS RESTANTES:<br>'
            
            for i in lista_desenhos:
                log_erro += i + '<br>'
                print(i)

            lista_desenhos = []

        for desenho in lista_desenhos:

            try:
                log_saida = update_block(folder_in, folder_out, desenho, app, args_revisao,
                            n_bloco, space, delay)
                lista_desenhos.remove(desenho)

                if log_saida == None:
                    log_saida = ''

            except Exception as e:

                if 'BLOCO NÃO FOI ENCONTRADO' in str(e):
                    lista_desenhos_bloco_nencontrado.append(desenho)
                    lista_desenhos.remove(desenho)
 
                print('ERRO NO DESENHO:  ' + desenho + ' Tentando Novamente')
                print(e)

                if str(e) in lista_erros:
                    pass
                else:
                    lista_erros.append(str(e))

                cont_err_cham += 1
                print(cont_err_cham)


    try:
        for document in app.Documents:
            document.Close(SaveChanges=False)
    except Exception as e:
        log_saida += str(e)

    t2 = datetime.datetime.now()

    log_saida += 'Concluído ' + str(n_desenhos) + ' desenhos em ' + f'{datetime.datetime.now():%X}'
    log_saida += '<br>'

    if len(lista_desenhos_bloco_nencontrado) > 0:

        log_saida += 'DESENHOS EM QUE O BLOCO NÃO FOI ENCONTRADO:'

        for i in lista_desenhos_bloco_nencontrado:

            log_saida +='<br>' + i

    else:
        log_saida += '<br>'
        log_saida += 'NÃO HOUVE DESENHO ONDE O BLOCO NÃO FOI ENCONTRADO'

    if log_erro != '':
        log_saida = 'Concluído em ' + f'{datetime.datetime.now():%X} <br>'
        log_saida += log_erro

        log_saida += '<br>Erros ocorrentes: <br>'

        for erro in lista_erros:

            log_saida += '-' + erro + '<br>'
    else:
        pass
    try:
        app.Quit()
    except:
        pass
    
    return log_saida














class revisador_carimbo:

    def __init__(self, folder_in, dados_revisao, path_saida=r'L:\IFR - Infraestrutura\04_PROJETOS INFRA\
                 PROGRAMA EP\Automatização de Documentos\Revisão de Carimbo\SAIDA', limite_erros = 20, space = 'PaperSpace'):

        self.folder_in = folder_in

        self.dados_revisao = dados_revisao

        self.folder_out = path_saida


        self.utilizar_app_cad = dados_revisao['app_cad']

        self.app = None

        self.limite_erros = limite_erros

        self.log_saida = ''

        self.log_erros = ''
        
        self.tarefa_atual = ''

        self.lista_desenhos = self.get_lista_desenhos()

        self.lista_desenhos_bloco_nencontrado = []

        self.set_app()


        self.args_revisao =  [

            dados_revisao['REV'],

            dados_revisao['T.E.'],

            dados_revisao['DESCRIÇÃO'],

            dados_revisao['PROJ'],

            dados_revisao['DES.'],

            dados_revisao['VER.'],

            dados_revisao['A.A.R.'],

            dados_revisao['AUT.'],

            dados_revisao['DATA'],

            dados_revisao['padrao_bloco'],

            dados_revisao['n_bloco'],

            space

            ]

    def executar_revisao(self):

        pass


    def get_lista_desenhos(self):

        return lista_des_dir.list_dir_func(self.folder_in)


    def revisar_desenho(self, desenho):

        status = False

        while not status:

            try:

                status = update_block(self.folder_in, self.folder_out, desenho, self.app, self.args_revisao,
                                self.args_revisao[-2], self.args_revisao[-1]) #TODO Alterar update bloco para gerir self.args_revisao
                
            except Exception as e:

                if 'BLOCO NÃO FOI ENCONTRADO' in str(e):

                    self.lista_desenhos_bloco_nencontrado.append(desenho)

                    continue

                if str(e) not in self.lista_erros:
                    
                    self.append(str(e))

                cont_err_cham += 1

                for document in self.app.Documents:
                    document.Close(SaveChanges=False)
                



    def set_app(self):

        if self.utilizar_app_cad == 'zwcad':

            try:

                self.app = win32com.client.dynamic.Dispatch("ZWCAD.Application")

            except Exception as e:

                self.log_erros += 'ERRO SET_APP: ' + str(e) + '<br>'


        elif self.utilizar_app_cad =='autocad':

            try:

                self.app = win32com.client.dynamic.Dispatch("AutoCAD.Application")

            except Exception as e:

                self.log_erros += 'ERRO SET_APP: ' + str(e) + '<br>'

        if self.app == None:

            self.log_erros += 'ERRO SET_APP: APLICATIVO CAD NÃO ENCONTRADO<br>'

        else:

            if len(self.app.Documents) != 0:

                if self.app.Documents[0].WindowTitle == 'Drawing1.dwg':

                    self.app.Documents[0].Close()

                else:

                    self.app = None

                    self.log_erros += 'ERRO SET_APP: FECHAR TODOS OS ARQUIVOS ABERTOS NO CAD ANTES DE EXECUTAR<br>'



    def visibilidade_cad(self, vis):

        if self.app != None:

            self.app.Visible = vis
