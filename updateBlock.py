import os
import time


'''
Função auxiliar para verificar se uma string contém números

'''
def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)


'''
Função auxiliar para subir letra da revisão

'''

def subirletra(revisao):

    revisao = revisao.replace(' ','')

    alfabeto = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    if revisao in alfabeto:

        return alfabeto[alfabeto.index(revisao)+1]
    
    else:

        return revisao
    
'''
Função que faz alterações nas tags do bloco de revisão para o padrão Vale
Estrutura:

REV-TIPO DE EMISSÃO-DESCRIÇÃO-PROJETISTA-DESENHISTA-VERIFICADOR-APROVADOR-AUTORIZADOR-DATA

'''

def subir_revisao_vale(tags_bloco, args_revisao):

    # Separar os dados de revisão (começam a partir da posição 13 no caso do bloco da Vale)

    dados_revisao = [i.TextString for i in tags_bloco[13:]]

    # Pegar todas as revisões (separadas de 9 em 9 no bloco da Vale)

    revisoes = dados_revisao[::9]

    # Encontrar a última revisão e seu index no array dos atributos

    ultima_revisao = 'A'

    for i in revisoes:

        if i != '':

            ultima_revisao = i

    index_ultima_revisao = revisoes.index(ultima_revisao)
    
    index_ultima_revisao = index_ultima_revisao*9 + 13

    index_nova_revisao = index_ultima_revisao + 9


    # Subir a última revisão

    if has_numbers(ultima_revisao):

        ultima_revisao = str(int(ultima_revisao)+1)

    else:

        ultima_revisao = subirletra(ultima_revisao)


    # Adicionar dados da nova revisão
    
    rev = args_revisao[0]

    te = args_revisao[1]

    descricao = args_revisao[2]

    proj = args_revisao[3]

    des = args_revisao[4]

    ver = args_revisao[5]

    aar = args_revisao[6]

    aut = args_revisao[7]

    data_emissao = args_revisao[8]

    # Revisão no carimbo

    if rev == '':
        tags_bloco[12].TextString = ultima_revisao
    else:
        tags_bloco[12].TextString = rev

    # Revisão na tabela de revisão    
    if rev == '':
        tags_bloco[index_nova_revisao].TextString = ultima_revisao
    else:
        tags_bloco[index_nova_revisao].TextString = rev

    # Tipo de emissão na tabela de revisão 
    if te == '':
        tags_bloco[index_nova_revisao + 1].TextString = tags_bloco[index_ultima_revisao + 1].TextString
    else:
        tags_bloco[index_nova_revisao + 1].TextString = te

    # Descrição da emissão na tabela de revisão 
    if descricao == '':
        tags_bloco[index_nova_revisao + 2].TextString = tags_bloco[index_ultima_revisao + 2].TextString
    else:
        tags_bloco[index_nova_revisao + 2].TextString = descricao

    # Sigla do projetista na tabela de revisão 
    if proj == '':
        tags_bloco[index_nova_revisao + 3].TextString = tags_bloco[index_ultima_revisao + 3].TextString
    else:
        tags_bloco[index_nova_revisao + 3].TextString = proj

    # Sigla do desenhista na tabela de revisão 
    if des == '':
        tags_bloco[index_nova_revisao + 4].TextString = tags_bloco[index_ultima_revisao + 4].TextString
    else:
        tags_bloco[index_nova_revisao + 4].TextString = des

    # Sigla do verificador na tabela de revisão 
    if ver == '':
        tags_bloco[index_nova_revisao + 5].TextString = tags_bloco[index_ultima_revisao + 5].TextString
    else:
        tags_bloco[index_nova_revisao + 5].TextString = ver

    # Siglado aprovador na tabela de revisão 
    if aar == '':
        tags_bloco[index_nova_revisao + 6].TextString = tags_bloco[index_ultima_revisao + 6].TextString
    else:
        tags_bloco[index_nova_revisao + 6].TextString = aar

    # Sigla do autorizador na tabela de revisão 
    if aut == '':
        tags_bloco[index_nova_revisao + 7].TextString = tags_bloco[index_ultima_revisao + 7].TextString
    else:
        tags_bloco[index_nova_revisao + 7].TextString = aut

    # Data de emissão na tabela de revisão 
    tags_bloco[index_nova_revisao + 8].TextString = data_emissao

    # Scale factor do atributo de data de emissão
    tags_bloco[index_nova_revisao + 8].ScaleFactor = 1



'''
Função que faz alterações nas tags do bloco de revisão para o padrão MRS
Estrutura:

REV-TIPO DE EMISSÃO-DESCRIÇÃO-PROJETISTA-DATA
'''


def subir_revisao_mrs(tags_bloco, args_revisao):

    
    # Separar os dados de revisão (começam a partir da posição 13 no caso do bloco da MRS)

    dados_revisao = [i.TextString for i in tags_bloco[13:]]

    # Pegar todas as revisões (separadas de 5 em 5 no bloco da Vale)

    revisoes = dados_revisao[::5]

    # Encontrar a última revisão e seu index no array dos atributos

    ultima_revisao = '0'

    for i in revisoes:

        if i != '':

            ultima_revisao = i

    index_ultima_revisao = revisoes.index(ultima_revisao)
    
    index_ultima_revisao = index_ultima_revisao*5 + 13

    index_nova_revisao = index_ultima_revisao + 5


    # Subir a última revisão
    if has_numbers(ultima_revisao):

        ultima_revisao = str(int(ultima_revisao)+1)

    else:

        ultima_revisao = subirletra(ultima_revisao)

    # Adicionar dados da nova revisão
    
    rev = args_revisao[0]

    te = args_revisao[1]

    descricao = args_revisao[2]

    proj = args_revisao[3]

    data_emissao = args_revisao[8]

    # Revisão na tabela de revisão    

    if rev == '':
        tags_bloco[index_nova_revisao].TextString = ultima_revisao
    else:
        tags_bloco[index_nova_revisao].TextString = rev

    # Data de emissão na tabela de revisão
    if len(data_emissao) == 10:
        tags_bloco[index_nova_revisao + 1].TextString = data_emissao[:-4] + data_emissao[-2:]
    else:
        tags_bloco[index_nova_revisao + 1].TextString = data_emissao

    # Tipo de emissão na tabela de revisão 
    if te == '':
        tags_bloco[index_nova_revisao + 2].TextString = tags_bloco[index_ultima_revisao + 2].TextString
    else:
        tags_bloco[index_nova_revisao + 2].TextString = te

    # Sigla do projetista na tabela de revisão 
    if proj == '':
        tags_bloco[index_nova_revisao + 3].TextString = tags_bloco[index_ultima_revisao + 3].TextString
    else:
        tags_bloco[index_nova_revisao + 3].TextString = proj

    # Descrição da emissão na tabela de revisão 
    if descricao == '':
        tags_bloco[index_nova_revisao + 4].TextString = tags_bloco[index_ultima_revisao + 4].TextString
    else:
        tags_bloco[index_nova_revisao + 4].TextString = descricao



'''
Função para alteração de atributos em um desenho de autocad

Parameters

folder_in : str
    Caminho para o desenho que será alterado

folder_out : str
    Caminho para o local que o desenho alterado será salvo

file : str
    Nome do desenho que será alterado
    
zcadapp : str
    Objeto tipo COM que contém uma aplicação ZWCad aberta
    
data_emissao : str
    Data de emissão do documento

n_bloco : str
    Nome do bloco que terá atributos alterados
    
space : str
    string que define se o bloco a ser substituido está no model space ou no paperspace

'''

def update_block(folder_in, folder_out, file, zcadapp, args_revisao, n_bloco, space):

    # Caminho para os arquivos de entrada

    drawing_in = os.path.join(folder_in, file)

    # Abrindo os arquivos para manipulação

    doc = zcadapp.Documents.Open(drawing_in)

    time.sleep(2)



    # Verificar o estado do ZWcad para evitar erros de chamado

    # state = zcadapp.GetZcadState()

    # while not state.IsQuiescent:
    #     time.sleep(2)

    # Selecionar o space que será feita a procura pelo bloco

    if space == 'PaperSpace':
        time.sleep(2)
        entities = zcadapp.ActiveDocument.PaperSpace
    elif space == 'ModelSpace':
        entities = zcadapp.ActiveDocument.ModelSpace
    else:
        raise ValueError('Erro na variável de entrada: space')
    
    # Procurar o bloco no desenho

    achou = False

    for entity in entities:
        name = entity.EntityName

        if name == 'AcDbBlockReference':
            if entity.EffectiveName == n_bloco:
                tags_bloco =  entity.GetAttributes()
                achou = True
                break

    if not achou:
        
        raise ValueError('BLOCO NÃO FOI ENCONTRADO')


    #SUBIR REVISÃO
    if achou:

        padrao_bloco = args_revisao[-1]

        if padrao_bloco == 'vale':
            subir_revisao_vale(tags_bloco, args_revisao)

        elif padrao_bloco == 'mrs':
            subir_revisao_mrs(tags_bloco, args_revisao)


    doc.PurgeAll()

    drawing_out = os.path.join(folder_out, file)

    if padrao_bloco == 'vale':

        file_split = file.split('=')

        file = file_split[0] + '=' + str(int(file_split[-1][0])+1) + '.' + file_split[-1].split('.')[-1]


    elif padrao_bloco == 'mrs':

        pass
    
    drawing_out = os.path.join(folder_out, file)

    # Salvar desenho
    time.sleep(2)
    doc.SaveAs(drawing_out)

    # Fechar o desenho
    for document in zcadapp.Documents:
        document.Close(SaveChanges=False)

    return True
