# Importando bibliotecas que serão utilizadas

import pandas as pd
import pyautogui
import time 


# Criando variáveis para armazenar as imagens que serão utilizadas no processo de 
# automação de cliques do mouse e preenchimentos de campos com teclado.

time.sleep(5)
buttom_new = 'botaonovo.png'
buttom_new_ = 'botaonovo_.png'
button_record = 'botaogravar.png'
buttom_contab = 'btcontabilidade.png'
buttom_complem = 'btcomplement.png'
buttom_perid = 'periodomenor.png'
buttom_parc = 'parcela.png'
buttom_entra = 'entrada.png'
buttom_canc = 'cancela.png'
buttom_semp = 'semparcela.png'

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

if click_new == 'botaonovo.png':
    pyautogui.click(click_new)
else:
    pyautogui.click(click_new_)
time.sleep(1)

# utilizando a função read_excel para leitura do arquivo no formato Excel
# junto com seu parâmetro sheet_name para leitura da aba correta. 

df = pd.read_excel('notasServicosSP_cubatao.xlsx', sheet_name='Mar')

# selecionando colunas do DataFrame para realizar a iteração paralela
# abaixo.

cnpj = df['cnpj'].astype(str)
nNota = df['numeroNota']
dtemissao = df['dataEmissao']
acumula = df['acumulador']
cfops = df['cfop']
vlnfs = df['valorNotas']

# estrutura de repetição com iteração paralela com as colunas do Dataframe

for cn, nn, de, ac, cf, vn in zip(cnpj, nNota, dtemissao, acumula, cfops, vlnfs):
    pyautogui.write("39")
    time.sleep(1)
    
    # estrutura de repetição para seleção correta do campo para preenchimento com
    # aguardo de 3 segundos evitando travamento do sistema.

    for i in range(4):
        pyautogui.press('TAB')
    time.sleep(3)
    print(cn)
    
    # estrutura condicional para preenchimento do campo "Fornecedor" com o número
    # de dígitos obrigatórios no sistema.
    # no sistema de Escrita Fiscal.

    if  len(cn) == 14:
        pyautogui.write(cn, interval=0.1)
    else:
        pyautogui.write("0" + cn, interval=0.1)
        time.sleep(0.5)

    # estrutra de repetição para preencher no campo correta o número da nota fiscal emitida 
    # pelo Fornecedor.   
    
    for i in range(2):
        pyautogui.press('ENTER')
    time.sleep(0.5)
    pyautogui.typewrite(str(nn), interval=0.1)
    time.sleep(1)

    # estrutura de repetição para preencher no campo correto a Data de Emissão da nota 
    # emitida pelo Fornecedor. 

    for i in range(3):
        pyautogui.press('ENTER')
    time.sleep(1)
    pyautogui.write(str(de))
    time.sleep(1)
    pyautogui.press('Enter')
    
    # preenchimento da Data de Saida da nota de serviço emitida no seu campo correto
    # dentro do sistema de Escrita Fiscal. 

    pyautogui.write(str(de))
    pyautogui.press('ENTER')
    time.sleep(1)
    
    
    # preenchimento do código "Acumulador" de acordo com a nota emitida (com retenção de ISS ou sem). 
    # OBS: este processo por enquanto esta manual. O usuário deve preencher este código manualmente
    # no DataFrame(planilha Excel) antes de iniciar o script.
    
    pyautogui.write(str(ac), interval=0.1)
    pyautogui.press('ENTER')
    time.sleep(3)
    
    # preenchimento do código "CFOP" de acordo com a nota emitida.
    # OBS: este processo por enquanto esta manual. O usuário deve preencher este código manualmente
    # no DataFrame(planilha Excel) antes de iniciar o script.

    pyautogui.write(str(cf), interval=0.4)
    pyautogui.press('ENTER')
    time.sleep(3)
    
    
    # preenchimento do valor total da nota fiscal emitida pelo Fornecedor, substituíndo o carácter "."(ponto) 
    # pelo carácter ","(vírgula), padrão do sistema da Domínio e padrão brasileiro para valores monetários.

    pyautogui.write(str(vn).replace('.', ','), interval=0.2)
    time.sleep(4)
    
    
    # estrutura de repetição para preenchimento automático dos demais campos que formam 
    # o valor total da nota fiscal emitida pelo Fornecedor.
    #   
    for i in range(4):
        pyautogui.press('ENTER')

    # clique na aba "Contabilidade" obedecendo o fluxo natural que o sistema decorre com
    # aguardo de 4 segundos para não travar o sistema.

    pyautogui.click(buttom_contab)
    time.sleep(4)
    
    # preenchimento das contas contábeis que serão utilizadas para posterior integração contábil
    # OBS: estas contas são de responsabilidade do usuário e de acordo com o seu plano de contas.
    # Caso os acumuladores já estejam parametrizados com as contas contábeis utilizadas para contabilização,
    # este trecho de código pode ser descartado sua utilização. Neste caso, as contas não estavam parametrizadas.

    pyautogui.press('ENTER')
    time.sleep(1)
    pyautogui.write('411')
    pyautogui.press('Enter')
    
    
    for i in range(3):
        pyautogui.press('ENTER')
        time.sleep(2)

    pyautogui.write('10000')
    pyautogui.press('Enter')
    
    
    for i in range(4):
        pyautogui.press('ENTER')
        time.sleep(2)

    pyautogui.write('478')
    pyautogui.press('ENTER')


    for i in range(4):
        pyautogui.press('ENTER')
        time.sleep(2)

    pyautogui.write('477')
    pyautogui.press('ENTER')
    
    
    for i in range(4):
        pyautogui.press('ENTER')
        time.sleep(2)

    pyautogui.write('509')
    pyautogui.press('ENTER')
    
    
    for i in range(4):
        pyautogui.press('ENTER')
        time.sleep(2)

    # clique na aba "Complementar" obedecendo o fluxo natural que o sistema decorre com
    # aguardo de 5 segundos para não travar o sistema.
    
    pyautogui.click(buttom_complem)
    time.sleep(5)
    
    # utilizando a função hotkey do pyautogui para pressionar a tecla "Gravar" com o atalho
    # Alt+g do teclado, com aguardo de 5 segundos para não travar o sistema.
    
    pyautogui.hotkey('Alt', 'g')
    time.sleep(5)

    # utilizando a função hotkey do pyautogui para pressionar a tecla "Gravar" com o atalho
    # Alt+g do teclado, com aguardo de 5 segundos para não travar o sistema.

    pyautogui.hotkey('Alt', 'g')
    time.sleep(5)

    # utilizando a função press para pressionar a tecla SIM para confirmar gravação do lançamento
    # da nota fiscal emitida pelo Fornecedor com aguardo de 12 segundos para evitar travar o sistema
    # e a tela retornar à sua aba principal.

    pyautogui.press('s')
    time.sleep(12)
