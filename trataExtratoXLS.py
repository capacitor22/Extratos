import re, logging, xlrd, csv
# import os, PyPDF2
from xlutils.copy import copy

# from tabulate import tabulate
# import tabula 
import pdfplumber


def processaBradesco(arquivo='./'):

    #Criando um logger para o modulo
    trataExtratos_logger = logging.getLogger('root.trataExtratos')    
    trataExtratos_logger.info('Processando extrato do Bradesco')

    # Hardcoded para teste
    if (arquivo == './'):
        planilha = arquivo + 'Bradesco.xls'
    else:
        planilha = arquivo

    workbook = xlrd.open_workbook(planilha)
    sheet = workbook.sheet_by_index(0)
    
    linhas = []
    i = 0
    linhaInicial = 0
    linhaFinal = 0

    # Marcando o inicio e o fim da parte relevante do extrato
    for linha in range(0,sheet.nrows):
        if (sheet.row_values(linha)[5] == 'Saldo (R$)'):
            linhaInicial = i 
        if (sheet.row_values(linha)[0].startswith('Os dados acima têm como base')):
            linhaFinal = i
        i += 1

    # Povoando a lista linhas
    for linha in range(linhaInicial,linhaFinal):
        if not ((sheet.row_values(linha)[1].startswith('Total')) or (sheet.row_values(linha)[1].startswith('SALDO'))):
            if (sheet.row_values(linha)[0] == ''):
                linhas[len(linhas)-1].append(sheet.row_values(linha)[1])
            elif (sheet.row_values(linha)[0] == 'Data'):
                linhas.append(sheet.row_values(linha))
                linhas[len(linhas)-1].append('Cred/Deb')
                linhas[len(linhas)-1].append('Observação Adicional')
            else:
                linhas.append(sheet.row_values(linha))
                if (sheet.row_values(linha)[3] == ''):
                    # linhas[len(linhas)-1].append(sheet.row_values(linha)[4])
                    valorSemPonto = sheet.row_values(linha)[4].replace('.', '')
                    linhas[len(linhas)-1].append(valorSemPonto)
                    
                else:
                    # linhas[len(linhas)-1].append(sheet.row_values(linha)[3])
                    valorSemPonto = sheet.row_values(linha)[3].replace('.', '')
                    linhas[len(linhas)-1].append(valorSemPonto)

    # print(f'LINHAS = {linhas}')
    # print(f'TYPE OF LINHAS = {type(linhas)}')
    return (linhas)

def processaItau(arquivo='./'):
    #Criando um logger para o modulo
    trataExtratos_logger = logging.getLogger('root.trataExtratos')    
    trataExtratos_logger.info('Processando extrato do Itaú. Arquivo %s', arquivo)

    # Hardcoded para teste
    if (arquivo == './'):
        planilha = arquivo + 'Itau.xls'
    else:
        planilha = arquivo

    workbook = xlrd.open_workbook(planilha)
    sheet = workbook.sheet_by_index(0)

    linhas = []
    i = 0
    linhaInicial = 0
    linhaFinal = 0

    # Marcando o inicio e o fim da parte relevante do extrato
    for linha in range(0,sheet.nrows):
        if (sheet.row_values(linha)[1] == 'lançamento'):
            linhaInicial = i + 1
        if (sheet.row_values(linha)[0] == 'lançamentos futuros'):
            linhaFinal = i
        i += 1

    # Povoando a lista linhas
    for linha in range(linhaInicial,linhaFinal):
        if not (sheet.row_values(linha)[1].startswith('SALDO ')):
            linhas.append(sheet.row_values(linha))
        
    # print(f'LINHAS = {linhas}')
    # print(f'TYPE OF LINHAS = {type(linhas)}')
    return (linhas)

def processaBB_csv(arquivo='./'):
    #Criando um logger para o modulo
    trataExtratos_logger = logging.getLogger('root.trataExtratos')    
    trataExtratos_logger.info('Processando extrato do Banco do Brasil. Arquivo CSV %s', arquivo)

        # Hardcoded para teste
    if (arquivo == './'):
        arqcsv = arquivo + 'BB.csv'
        # print('entrei aqui e csv =' + arqcsv)
    else:
        arqcsv = arquivo

    linhas = []

    with open(arqcsv, mode='r') as arquivo_csv:
        leitor_csv = csv.reader(arquivo_csv)

        for linha in leitor_csv:
            linhas.append(linha)
    
    # print(linhas)
    return (linhas)

def processaBB(arquivo='./'):
    #Criando um logger para o modulo
    trataExtratos_logger = logging.getLogger('root.trataExtratos')    
    trataExtratos_logger.info('Processando extrato do Banco do Brasil. Arquivo PDF %s', arquivo)

        # Hardcoded para teste
    if (arquivo == './'):
        pdf = arquivo + 'BB.pdf'
    else:
        pdf = arquivo

    print('PDFPLUMBER')
    with pdfplumber.open(pdf) as pdfpdf:
        pagina1 = pdfpdf.pages[0]
        textoPDFPagina1 = pagina1.extract_text()
    print(textoPDFPagina1)
    print('--------------------')
    cabecalho = 'Dia Histórico Valor'
    nova_entrada = re.compile("(^[0-9]*/[0-9]*/[0-9]*)\s(.*)")
    nova_entrada_so_data = re.compile("^[0-9]*/[0-9]*/[0-9]*")
    nova_entrada_final = re.compile("(.*) ([0-9.]*,[0-9]*\s[\(\)\-\+]*)")
    nova_entrada_final_so_valor = re.compile("([0-9.]*,[0-9]*\s[\(\)\-\+]*)")

    entrar = False

    linhas = []
    linha = ['','','','']

    for line in textoPDFPagina1.split('\n'):
        if line == cabecalho:
            print('cabecalho = ' + cabecalho)
            entrar = True
        if line == 'Informações Adicionais':
            print('FIM DO PROCESSAMENTO')
            entrar = False
        if entrar:
            if nova_entrada.match(line):
                if linha[0] != '' and linha[1] != '' and linha[2] != '':                 
                    linhas.append(linha)
                    linha = ['','','','']
                line_re = nova_entrada.search(line) 
                print(r'(0) Data Nova Entrada: ' + line_re.group(1))
                print(r'(1) Descr1 Nova Entrada: ' + line_re.group(2))
                linha[0] = line_re.group(1) # Data
                linha[1] = line_re.group(2) # Histórico
            elif nova_entrada_so_data.match(line):
                if linha[0] != '' and linha[1] != '' and linha[2] != '':                 
                    linhas.append(linha)
                    linha = ['','','',''] 
                print(r'(0) Data Só Data: ' + line)
                linha[0] = line # Data
            elif nova_entrada_final.match(line):
                line_re = nova_entrada_final.search(line) 
                print(r'(1) Descr1 Nova Entrada Final: ' + line_re.group(1))
                print(r'(2) Valor Nova Entrada Final: ' + line_re.group(2))
                linha[1] = line_re.group(1) # Histórico
                linha[2] = line_re.group(2) # Valor
            elif nova_entrada_final_so_valor.match(line):
                print(r'(2) Valor Final so valor: ' + line)
                linha[2] = line # Valor
            else:
                print(r'(3) Descricao adicional: ' + line)
                if line == cabecalho:
                    linha = ['Dia', 'Histórico', 'Valor', 'Descrição Adicional']
                else:
                    linha[3] = line # Descricao Adicional

                linhas.append(linha)
                linha = ['','','','']


    print('================================')
    
    for line in linhas:
        print(line)

    return (linhas)



        
def processaOutros(arquivo='./'):

    print("-----------------------\n")
    print("Iniciando processamento para Outros...\n")
    print("-----------------------\n")

    #Criando um logger para o modulo
    trataExtratos_logger = logging.getLogger('root.trataExtratos')    

    trataExtratos_logger.info('Processando outros arquivos. Arquivo %s', arquivo)


if __name__ == '__main__':
    processaBB_csv()
    # main('Rebuild')
