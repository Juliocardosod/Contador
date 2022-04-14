from ast import Global
from configparser import ConfigParser
import os
import time
from xmlrpc.client import DateTime
import xlsxwriter

versão = 'V1.00'

cfg = ConfigParser()
cfg.read_file(open(os.path.join(os.path.dirname(__file__),"config.ini")))

#Adquire parametros do arquivo config
diretorio = cfg.get('DEFAULT','DIRETORIO')
codigos = ['200','400','401']#(cfg.get('DEFAULT','CODIGOS')).split(',')
metodos = (cfg.get('DEFAULT','METODOS')).split(',')


dir_list = os.listdir(diretorio)
# print(dir_list)

def coletaArquivo(dir):
    lista = []
    local = f'{diretorio}//{dir}'
    try:
        for filename in os.listdir(local):
            # filestamp = os.stat(os.path.join(diretorio, filename)).st_mtime
            lista.append(os.path.join(local, filename))
    except Exception as ex:
        print(f'Erro: {ex}')
    finally:
        return lista

def busca(): #Busca por métodos e códigos
    #CRIA ARQUIVO
    workbook = xlsxwriter.Workbook(f'{diretorio}//Relatório por cód.xlsx')

    for dir in dir_list: #Loop diretório / pasta
        c200 = 0
        c400 = 0
        c401 = 0
        row = 2 #Contador de inhas da planilha

        sheet = workbook.add_worksheet(dir)
        #CABEÇALHO
        sheet.write(f'A1', 'Data')
        sheet.write(f'B1', 'Método')
        sheet.write(f'C1', 'Ok Cód 200')
        sheet.write(f'E1', 'Erro Cód 400')
        sheet.write(f'G1', 'Falha Cód 401')
        sheet.write(f'I1', 'Total')

        percent_fmt = workbook.add_format({'num_format': '0.00%'})

        arquivos = coletaArquivo(dir)
        for arquivo in arquivos: #Loop arquivo / linha
            # print('Nome do arquivo: '+os.path.basename(arquivo)) #Imprime o nome do arquivo
            
            #DATA ARQUIVO
            ti_m = os.path.getmtime(arquivo)
            m_ti = time.ctime(ti_m)
            t_obj = time.strptime(m_ti)
            T_stamp = time.strftime("%Y-%m-%d", t_obj)
            
            for item in codigos: #Códigos
                for palavra in metodos: #Métodos
                # for item in codigos: #Códigos
                    arq = open(arquivo, "r")
                    contador_linhas = 0
                    palavraEncontrada = 0
                    try:
                        for linha in arq:
                            contador_linhas = contador_linhas + 1
                            if item in linha and palavra in linha:
                                palavraEncontrada = palavraEncontrada+1
                    except Exception as ex:
                        print(f'Erro na linha {contador_linhas} no arquivo: {os.path.basename(arquivo)} da pasta {dir}')

                #Salva quantidades nas variaveis        
                    if item == '200':
                        c200 = palavraEncontrada
                    if item == '400':
                        c400 = palavraEncontrada
                    if item == '401':
                        c401 = palavraEncontrada

            # print(f'Código {item}: {palavraEncontrada} \n')

            arq.close()
            #INSERE DADOS
            sheet.write(f'A{row}', T_stamp)
            sheet.write(f'B{row}', 'Geral')
            sheet.write(f'C{row}', c200)
            if c200 == 0:
                sheet.write(f'D{row}', 0)
            else:
                sheet.write(f'D{row}', f'=C{row}/I{row}')
            sheet.write(f'E{row}', c400)
            if c400 == 0:
                sheet.write(f'F{row}', 0)
            else:
                sheet.write(f'F{row}', f'=E{row}/I{row}')
            sheet.write(f'G{row}', c401)
            if c401 == 0:
                sheet.write(f'H{row}', 0)
            else:
                sheet.write(f'H{row}', f'=G{row}/I{row}')
            sheet.write(f'I{row}', f'=C{row}+E{row}+G{row}')

            row = row+1
        sheet.set_column(f'D:D', None, percent_fmt)
        sheet.set_column(f'F:F', None, percent_fmt)
        sheet.set_column(f'H:H', None, percent_fmt)  

    workbook.close()

def buscaMetodo(): #Busca por métodos e códigos
    #CRIA ARQUIVO
    workbook = xlsxwriter.Workbook('Relatório por método.xlsx')

    for dir in dir_list:
        c200 = 0
        c400 = 0
        c401 = 0
        row = 2 #Contador de inhas da planilha

        sheet = workbook.add_worksheet(dir)
        #CABEÇALHO
        sheet.write(f'A1', 'Data')
        sheet.write(f'B1', 'Método')
        sheet.write(f'C1', 'Ok Cód 200')
        sheet.write(f'E1', 'Erro Cód 400')
        sheet.write(f'G1', 'Falha Cód 401')
        sheet.write(f'I1', 'Total')

        percent_fmt = workbook.add_format({'num_format': '0.00%'})

        arquivos = coletaArquivo(dir)
        for arquivo in arquivos:
            # print('Nome do arquivo: '+os.path.basename(arquivo)) #Imprime o nome do arquivo
            
            #DATA ARQUIVO
            ti_m = os.path.getmtime(arquivo)
            m_ti = time.ctime(ti_m)
            t_obj = time.strptime(m_ti)
            T_stamp = time.strftime("%Y-%m-%d", t_obj)
            
            for palavra in metodos: #Métodos
                for item in codigos: #Códigos
                    arq = open(arquivo, "r")
                    contador_linhas = 0
                    palavraEncontrada = 0
                    try:
                        for linha in arq:
                            contador_linhas = contador_linhas + 1
                            if item in linha and palavra in linha:
                                palavraEncontrada = palavraEncontrada+1
                    except Exception as ex:
                        print(f'Erro na linha {contador_linhas} no arquivo: {os.path.basename(arquivo)} da pasta {dir}')
                
                #Salva quantidades nas variaveis        
                    if item == '200':
                        c200 = palavraEncontrada
                    if item == '400':
                        c400 = palavraEncontrada
                    if item == '401':
                        c401 = palavraEncontrada

                    # print(f'Método {palavra} Código {item}: {palavraEncontrada} \n')
                    arq.close()
                #INSERE DADOS
                sheet.write(f'A{row}', T_stamp)
                sheet.write(f'B{row}', f'{palavra}')
                sheet.write(f'C{row}', c200)
                if c200 == 0:
                    sheet.write(f'D{row}', 0)
                else:
                    sheet.write(f'D{row}', f'=C{row}/I{row}')
                sheet.write(f'E{row}', c400)
                if c400 == 0:
                    sheet.write(f'F{row}', 0)
                else:
                    sheet.write(f'F{row}', f'=E{row}/I{row}')
                sheet.write(f'G{row}', c401)
                if c401 == 0:
                    sheet.write(f'H{row}', 0)
                else:
                    sheet.write(f'H{row}', f'=G{row}/I{row}')
                sheet.write(f'I{row}', f'=C{row}+E{row}+G{row}')

                row = row+1
        sheet.set_column(f'D:D', None, percent_fmt)
        sheet.set_column(f'F:F', None, percent_fmt)
        sheet.set_column(f'H:H', None, percent_fmt)  

    workbook.close()

def buscaCodigo(): #Busca por códigos
    workbook = xlsxwriter.Workbook('sample.xlsx')
    for dir in dir_list:
        c200 = 0
        c400 = 0
        c401 = 0
        row = 2 #Contador de inhas da planilha
        
        
        #CRIA ARQUIVO
        
        sheet = workbook.add_worksheet(dir)
        #CABEÇALHO
        sheet.write(f'A1', 'Data')
        sheet.write(f'B1', 'Método')
        sheet.write(f'C1', 'Ok Cód 200')
        sheet.write(f'E1', 'Erro Cód 400')
        sheet.write(f'G1', 'Falha Cód 401')
        sheet.write(f'I1', 'Total')

        percent_fmt = workbook.add_format({'num_format': '0.00%'})

        arquivos = coletaArquivo(dir)
        for arquivo in arquivos:
            # print('Nome do arquivo: '+os.path.basename(arquivo)) #Imprime o nome do arquivo
            
            #DATA ARQUIVO
            ti_m = os.path.getmtime(arquivo)
            m_ti = time.ctime(ti_m)
            t_obj = time.strptime(m_ti)
            # dataArquivo = DateTime(os.stat(os.path.join(diretorio, arquivo)).st_mtime)
            T_stamp = time.strftime("%Y-%m-%d", t_obj)
            # print(T_stamp)
            
            for item in codigos:
                arq = open(arquivo, "r")
                contador_linhas = 0
                palavraEncontrada = 0
                try:
                    for linha in arq:
                        contador_linhas = contador_linhas + 1
                        if item in linha:
                            palavraEncontrada = palavraEncontrada+1
                except Exception as ex:
                        print(f'Erro na linha {contador_linhas} no arquivo: {os.path.basename(arquivo)} da pasta {dir}')
                
                #Salva quantidades nas variaveis        
                if item == '200':
                    c200 = palavraEncontrada
                if item == '400':
                    c400 = palavraEncontrada
                if item == '401':
                    c401 = palavraEncontrada
                
                # print(f'Código {item}: {palavraEncontrada}')
                arq.close()

            #INSERE DADOS
            sheet.write(f'A{row}', T_stamp)
            sheet.write(f'B{row}', 'Geral')
            sheet.write(f'C{row}', c200)
            if c200 == 0:
                sheet.write(f'D{row}', 0)
            else:
                sheet.write(f'D{row}', f'=C{row}/I{row}')
            sheet.write(f'E{row}', c400)
            if c400 == 0:
                sheet.write(f'F{row}', 0)
            else:
                sheet.write(f'F{row}', f'=E{row}/I{row}')
            sheet.write(f'G{row}', c401)
            if c401 == 0:
                sheet.write(f'H{row}', 0)
            else:
                sheet.write(f'H{row}', f'=G{row}/I{row}')
            sheet.write(f'I{row}', f'=C{row}+E{row}+G{row}')
            
            

            # salvaPlanilha('sample.xlsx', row, 'tabela', T_stamp, c200, c400, c401 )
            row = row + 1
        #DEFINE FORMATO PORCENTAGEM
        sheet.set_column(f'D:D', None, percent_fmt)
        sheet.set_column(f'F:F', None, percent_fmt)
        sheet.set_column(f'H:H', None, percent_fmt)  

    workbook.close()


# buscaMetodo()#Busca por método

# buscaCodigo()Busca por código
busca() #Busca por código V2    