# Importando bibliotecas que serão utilizadas
from pandas import read_excel
from time import sleep
import pyautogui


# criando menu principal personalizado com as opções de funções.
print('-' * 56)
print (f"{'NOME DA EMPRESA': >32}")
print('-' * 56)
print(f"{'MENU PRINCIPAL': ^48}")
print('-'* 56)
print('[ 1 ] - Entrada de NFS-e')
print('[ 2 ] - Lançamento de Retenções NFS-e')
print('[ 3 ] - Sair')
options = int(input('Escolha a opção desejada: '))
sleep(7)


# Criando variáveis para armazenar as imagens que serão utilizadas no processo de 
# automação de cliques do mouse e preenchimentos de campos com teclado.      
buttom_new = (r'.\image\botaonovo.PNG')
buttom_new_ = (r'.\image\botaonovo_.PNG')
button_record = (r'.\image\botaogravar.PNG')
buttom_contab = (r'.\image\btcontabilidade.PNG')
buttom_complem = (r'.\image\btcomplement.png')
buttom_perid = (r'.\image\periodomenor.png')
buttom_parc = (r'.\image\parcela.png')
buttom_entra = (r'.\image\entrada.png')
buttom_canc = (r'.\image\cancela.png')
buttom_semp = (r'.\image\semparcela.png')


# utilizando as variáveis para obter as coordenadas de cliques com a função
# locateCenterOnScreen da biblioteca Pyautogui.
click_new = pyautogui.locateCenterOnScreen(buttom_new)
click_new_= pyautogui.locateCenterOnScreen(buttom_new_)
to_record = pyautogui.locateCenterOnScreen(button_record)
click_cont = pyautogui.locateCenterOnScreen(buttom_contab)
click_comp = pyautogui.locateCenterOnScreen(buttom_complem)
click_perid = pyautogui.locateCenterOnScreen(buttom_perid)
click_parc = pyautogui.locateCenterOnScreen(buttom_parc)
click_entr = pyautogui.locateCenterOnScreen(buttom_entra)
click_canc = pyautogui.locateCenterOnScreen(buttom_canc)
click_semp = pyautogui.locateCenterOnScreen(buttom_semp)


# criando estrutura condicional para clique na coordenada correta
# quando a imagem pode estar alterada na abertura da tela de input de dados 
# no sistema da Domínio.
if options == 1:
    if click_new == 'botaonovo.png':
        pyautogui.click(click_new)
    else:
        click_new_ == 'botaonovo_.png'
        pyautogui.click(click_new_)
        sleep(2)


    # utilizando a função read_excel para leitura do arquivo no formato Excel
    # junto com seu parâmetro sheet_name para leitura da aba correta.
    df = read_excel(r'.\filesxls\notasServicosSP_cubatao.xlsx', sheet_name='Ago')


    # selecionando colunas do DataFrame para realizar a iteração paralela abaixo
    cnpj = df['cnpj'].astype(str)
    nNota = df['numeroNota']
    dtemissao = df['dataEmissao']
    acumula = df['acumulador']
    cfops = df['cfop']
    vlnfs = df['valorNotas']


    # estrutura de repetição para seleção correta do campo para preenchimento com
    # aguardo de 3 segundos evitando travamento do sistema.    
    for cn, nn, de, ac, cf, vn in zip(cnpj, nNota, dtemissao, acumula, cfops, vlnfs):
        pyautogui.write("39")
        sleep(1)
        
        # estrutura de repetição com iteração paralela com as colunas do Dataframe    
        for i in range(4):
            pyautogui.press('TAB')
        sleep(3)
        
    
        print('Incluindo Nota fiscal de serviço nº. ' + nn)
        if  len(cn) == 14:
            pyautogui.write(cn, interval=0.1)
        else:
            pyautogui.write('0' + cn, interval=0.1)            
        sleep(0.5)
        

        # estrutra de repetição para preencher no campo correta o número da nota fiscal emitida 
        # pelo Fornecedor.
        for i in range(2):
            pyautogui.press('ENTER')
        sleep(0.5)
        pyautogui.typewrite(str(nn), interval=0.1)
        sleep(1)
                        
        
        # estrutura de repetição para preencher no campo correto a Data de Emissão da nota 
        # emitida pelo Fornecedor.        
        for i in range(3):
            pyautogui.press('ENTER')
        sleep(1)
        pyautogui.write(str(de))
        pyautogui.press('Enter')
        sleep(1)
       
        
        # preenchimento da Data de Saida da nota de serviço emitida no seu campo correto
        # dentro do sistema de Escrita Fiscal.       
        pyautogui.write(str(de))
        pyautogui.press('ENTER')
        sleep(1)
        
        
        # preenchimento do código "Acumulador" de acordo com a nota emitida (com retenção de ISS ou sem). 
        # OBS: este processo por enquanto esta manual. O usuário deve preencher este código manualmente
        # no DataFrame(planilha Excel) antes de iniciar o script.
        pyautogui.write(str(ac), interval=0.1)
        pyautogui.press('ENTER')
        sleep(3)


        # preenchimento do código "CFOP" de acordo com a nota emitida.
        # OBS: este processo por enquanto esta manual. O usuário deve preencher este código manualmente
        # no DataFrame(planilha Excel) antes de iniciar o script.
        pyautogui.write(str(cf), interval=0.1)
        pyautogui.press('ENTER')
        sleep(3)
        
        
        # preenchimento do valor total da nota fiscal emitida pelo Fornecedor, substituíndo o carácter "."(ponto) 
        # pelo carácter ","(vírgula), padrão do sistema da Domínio e padrão brasileiro para valores monetários.
        pyautogui.write(str(vn).replace('.', ','), interval=0.1)
        sleep(1)


        # utilizando a função hotkey do pyautogui para pressionar a tecla "Gravar" com o atalho
        # Alt+g do teclado, com aguardo de 5 segundos para não travar o sistema.
        for i in range(4):
            pyautogui.press('ENTER')
            
        pyautogui.hotkey('Alt', 'g')
        sleep(2)


        # utilizando a função hotkey do pyautogui para pressionar a tecla "Gravar" com o atalho
        # Alt+g do teclado, com aguardo de 5 segundos para não travar o sistema.
        pyautogui.hotkey('Alt', 'g')
        sleep(2)
        
        
        # utilizando a função hotkey do pyautogui para pressionar a tecla "Gravar" com o atalho
        # Alt+g do teclado, com aguardo de 5 segundos para não travar o sistema.
        pyautogui.hotkey('Alt', 'g')
        sleep(2)

        
        # finalizando o lançamento.
        pyautogui.press('s')
        sleep(12)
