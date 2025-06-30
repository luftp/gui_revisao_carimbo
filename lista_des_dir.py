import os

def list_dir_comrev_func(folder_in):

    dic_rev = {}
    list_rev = []
    lista_desenhos = []

    for f in os.listdir(folder_in):
        if f.endswith('.dwg') or f.endswith('.DWG'):

            chave = 'DE-'+f.split('-')[-3] + '-' +f.split('-')[-2]

            if chave not in dic_rev.keys():
                dic_rev[chave] = []
                dic_rev[chave].append(f)
            else:
                dic_rev[chave].append(f)

    for desenho in dic_rev.keys():
        lista_desenhos.append(dic_rev[desenho][-1])

    return lista_desenhos


def list_dir_func(folder_in):

    lista_desenhos = []

    for f in os.listdir(folder_in):
        if f.endswith('.dwg') or f.endswith('.DWG'):
            lista_desenhos.append(f)

    return lista_desenhos

