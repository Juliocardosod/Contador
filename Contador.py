from configparser import ConfigParser
import os

cfg = ConfigParser()
cfg.read_file(open(os.path.join(os.path.dirname(__file__),"config.ini")))

#Adquire parametros do arquivo config
diretorio = cfg.get('DEFAULT','DIRETORIO')
codigos = (cfg.get('DEFAULT','CODIGOS')).split(',')
metodos = (cfg.get('DEFAULT','METODOS')).split(',')

def coletaArquivo():
    lista = []
    try:
        for filename in os.listdir(diretorio):
            # filestamp = os.stat(os.path.join(diretorio, filename)).st_mtime
            lista.append(os.path.join(diretorio, filename))
    except Exception as ex:
        print(f'Erro: {ex}')
    finally:
        return lista

def buscaMetodo(arquivos): #Busca por métodos e códigos
    for arquivo in arquivos:
        print('Nome do arquivo: '+os.path.basename(arquivo)) #Imprime o nome do arquivo
        for palavra in metodos:
            for item in codigos:
                arq = open(arquivo, "r")
                contador_linhas = 0
                palavraEncontrada = 0
                for linha in arq:
                    contador_linhas = contador_linhas + 1
                    if item in linha and palavra in linha:
                        palavraEncontrada = palavraEncontrada+1
                   
                print(f'Método {palavra} Código {item}: {palavraEncontrada} \n')
                arq.close()

def buscaCodigo(arquivos): #Busca por códigos
    for arquivo in arquivos:
        print('Nome do arquivo: '+os.path.basename(arquivo)) #Imprime o nome do arquivo
        for item in codigos:
            arq = open(arquivo, "r")
            contador_linhas = 0
            palavraEncontrada = 0
            for linha in arq:
                contador_linhas = contador_linhas + 1
                if item in linha:
                    palavraEncontrada = palavraEncontrada+1
                
            print(f'Código {item}: {palavraEncontrada}')
            arq.close()

listaArquivos = coletaArquivo()
# buscaMetodo(listaArquivos)
buscaCodigo(listaArquivos)    