from nicegui import run,ui
import datetime
import tkinter as tk
from tkinter import filedialog
import revisar_carimbo_func
from revisar_carimbo_func import revisador_carimbo
import threading
import time

class gui_revisao_carimbo:

    def __init__(self):

        data_hoje = datetime.datetime.now().strftime('%d/%m/%Y')
 
        dark = ui.dark_mode()

        dark.enable()

        self.lista_dados_revisao = {}

        self.lista_dados_revisao['REV'] = ''

        self.lista_dados_revisao['T.E.'] = ''

        self.lista_dados_revisao['DESCRIÇÃO'] = ''

        self.lista_dados_revisao['PROJ'] = ''

        self.lista_dados_revisao['DES.'] = ''

        self.lista_dados_revisao['VER.'] = ''

        self.lista_dados_revisao['A.A.R.'] = ''

        self.lista_dados_revisao['AUT.'] = ''

        self.lista_dados_revisao['DATA'] = ''

        self.lista_dados_revisao['n_bloco'] = 'A1'

        self.lista_dados_revisao['padrao_bloco'] = 'vale'

        self.lista_dados_revisao['app_cad'] = 'zwcad'

        self.lista_dados_entrada = {}

        self.pasta_entrada = ''

        self.args_revisao = ''

        self.state_exec = False



        with ui.card():

            ui.label('SUBIR REVISÃO DE CARIMBO (versão beta 0.1)').style('color:#6E93D6; font-size:150%')

            ui.space()

            ui.label('Programa CAD').props('inline')

            with ui.row():

                self.app_cad = ui.radio(['ZWCad', 'AutoCad'], value='ZWCad', on_change=self.update_lista_dados_revisao).props('inline')


            ui.label('Padrão Carimbo').props('inline')

            with ui.row():

                self.carimbo_padrao = ui.radio(['Vale', 'MRS'], value='Vale', on_change=self.update_lista_dados_revisao).props('inline')

            ui.label('Nome do Bloco - Dados de entrada').props('inline')

            self.n_bloco =  ui.input('Nome do Bloco', value='A1', on_change=self.update_lista_dados_revisao).style('width: 100px;')

            ui.label('Tabela de Revisão - Dados de entrada').props('inline')


            with ui.row():

                self.lista_dados_entrada['REV'] = ui.input('REV', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['T.E.'] = ui.input('T.E.', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['DESCRIÇÃO'] = ui.input('DESCRIÇÃO', on_change=self.update_lista_dados_revisao).style('width: 200px;')

                self.lista_dados_entrada['PROJ'] = ui.input('PROJ.', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['DES.'] = ui.input('DES.', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['VER.'] = ui.input('VER.', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['A.A.R.'] = ui.input('A.A.R.', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['AUT.'] = ui.input('AUT.', on_change=self.update_lista_dados_revisao).style('width: 50px;')

                self.lista_dados_entrada['DATA'] = ui.input('DATA', value=data_hoje, on_change=self.update_lista_dados_revisao).style('width: 80px;')



            with ui.row():


                ui.button('selecionar pasta com arquivos', on_click=lambda: self.browser_folder())

                self.spinner_select = ui.spinner(size='lg')

                self.spinner_select.visible = False


            # ui.button('Info', on_click=self.dialog_info.open)

            with ui.row():

                ui.button('Executar', on_click= lambda: self.chamar_revisar_carimbo())

                self.spinner_executar = ui.spinner(size='lg')

                self.spinner_executar.visible = False


            self.log_exec = ui.html()

            self.timer_exec = ''

        ui.run()


    def get_pasta_entrada(self):

        return self.pasta_entrada
    
    def get_lista_dados_revisao(self):

        return self.lista_dados_revisao
    
    def update_lista_dados_revisao(self):

        for values in self.lista_dados_entrada.keys():

            self.lista_dados_revisao[values] = self.lista_dados_entrada[values].value
        
        self.lista_dados_revisao['n_bloco'] = self.n_bloco.value

        self.lista_dados_revisao['padrao_bloco'] = self.carimbo_padrao.value.lower()

        self.lista_dados_revisao['app_cad'] = self.app_cad.value.lower()


        if self.lista_dados_revisao['padrao_bloco'] == 'mrs':

        
            self.lista_dados_entrada['DES.'].visible = False

            self.lista_dados_entrada['VER.'].visible = False

            self.lista_dados_entrada['A.A.R.'].visible = False

            self.lista_dados_entrada['AUT.'].visible = False
        
        elif self.lista_dados_revisao['padrao_bloco'] == 'vale':


            self.lista_dados_entrada['DES.'].visible = True

            self.lista_dados_entrada['VER.'].visible = True

            self.lista_dados_entrada['A.A.R.'].visible = True

            self.lista_dados_entrada['AUT.'].visible = True
        


    def todo(self):

        pass


    async def browser_folder(self):

        time.sleep(0.1)

        root = tk.Tk()

        root.withdraw()

        self.spinner_select.set_visibility(True)
        
        root.lift()

        root.attributes('-topmost', True)

        root.after_idle(root.attributes, '-topmost', False)

        root.update()

        self.pasta_entrada = await run.cpu_bound(filedialog.askdirectory)

        self.spinner_select.set_visibility(False)

        root.destroy()


    async def chamar_revisar_carimbo(self):

        if self.state_exec == True:
            pass

        else:
            self.state_exec = True

            if self.timer_exec == '':
                pass
            else:
                self.timer_exec.cancel()

            if self.pasta_entrada == '':

                self.log_exec.set_content('ERRO = PASTA DE ENTRADA DE ARQUIVOS NÃO SELECIONADA')

                return None

            self.log_exec.set_content('')

            self.update_lista_dados_revisao()

            time_inicio =f'Início Execução {datetime.datetime.now():%X}'

            self.log_exec.set_content(time_inicio)

            self.timer_exec = ui.timer(1.0, lambda: self.log_exec.set_content(time_inicio + '<br><br>' + f'Status da revisão: <br>Executando em {datetime.datetime.now():%X}'))

            self.spinner_executar.set_visibility(True)

            log = await run.cpu_bound(revisar_carimbo_func.revisar_carimbo, self.get_pasta_entrada(), self.get_lista_dados_revisao())

            if log != None:

                self.timer_exec.cancel()

                self.log_exec.set_content(time_inicio + '<br><br>Status da revisão: <br>' + log)

            self.spinner_executar.set_visibility(False)

            self.state_exec = False




    async def chamar_revisar_carimbo_com_classe(self):

        if self.state_exec == True:
            pass

        else:

            self.state_exec = True

            if self.timer_exec == '':
                pass
            else:
                self.timer_exec.cancel()


            if self.pasta_entrada == '':

                self.log_exec.set_content('ERRO = PASTA DE ENTRADA DE ARQUIVOS NÃO SELECIONADA')

                return None

            else:

                self.log_exec.set_content('')

                self.update_lista_dados_revisao()

                time_inicio =f'Início Execução {datetime.datetime.now():%X}'

                self.log_exec.set_content(time_inicio)

                self.timer_exec = ui.timer(1.0, lambda: self.log_exec.set_content(time_inicio + '<br><br>' + f'Status da revisão: <br>Executando em {datetime.datetime.now():%X}'))

                self.spinner_executar.set_visibility(True)

                self.revisador_carimbo = revisador_carimbo(self.pasta_entrada, self.lista_dados_revisao)

                log = await run.cpu_bound(self.revisador_carimbo.executar_revisao)

            if log != None:

                self.timer_exec.cancel()

                self.log_exec.set_content(time_inicio + '<br><br>Status da revisão: <br>' + log)

            self.spinner_executar.set_visibility(False)

            self.state_exec = False
