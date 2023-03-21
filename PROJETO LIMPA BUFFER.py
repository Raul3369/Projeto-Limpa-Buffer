# IMPORTANDO AS BIBLIOTECAS

from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
import glob
import os
import pandas as pd
from openpyxl import *
from datetime import date,time,datetime
import time
import win32com.client as win32

def get_files():
    # Testando e abrindo o navegador, com a msg que o Chrome está sendo controlado por um software
    driver = webdriver.Edge(r"C:\Program Files\MICROSOFT EDGE WEBDRIVER\teste novo\msedgedriver.exe")

    # Entrando no site do GPX, colocar o usuário, senha e clicar para entrar
    driver.get("http://sgpautomacaocontingencia.br154.corpintra.net/Eixos/Home/Login?ReturnUrl=%2fEixos%2fHome%2fIndex")
    time.sleep (3)
    # Inspecionando o campo que quero digitar o usuário, irei utilizar o 'name'
    usuario = driver.find_element(By.NAME,'usuario')
    usuario.send_keys ("aradael")

    # A cima o passo a passo e embaixo o jeito resumido, sem criar váriaveis
    driver.find_element(By.NAME, 'senha').send_keys ("Leonardo@023")

    # Digitar ENTER
    pyautogui.hotkey('ENTER')

    # comando para clicar em algum box da internet, através do elemento 
    driver.find_element(By.ID, 'tipoeixo').click()
    time.sleep (3)
    driver.find_element(By.XPATH, '//*[@id="wrapper"]/nav/div/div[2]/ul/li[1]/ul/li[3]/a').click()
    time.sleep (3)
    driver.find_element(By.XPATH, '//*[@id="lvl1Nav"]/ul/li[5]/a').click()
    time.sleep (3)
    driver.find_element(By.XPATH, '//*[@id="mapacubodianteiro_wrapper"]/div/a[1]').click()
    time.sleep (4)
    driver.find_element(By.CLASS_NAME, 'dt-buttons').click() #botão que baixa o excel


    time.sleep(50)

    pyautogui.hotkey('alt','f4')

    pyautogui.hotkey('ENTER')

    time.sleep(3)

    driver.find_element(By.XPATH, '//*[@id="myModal"]/div/div/div[1]/button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="tipoeixo"]').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="wrapper"]/nav/div/div[2]/ul/li[2]/ul/li[5]/a/strong').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="lvl1Nav"]/ul/li[5]/a').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="mapacubotraseiro_wrapper"]/div/a[1]/span').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="listacubotraseiroexcel_wrapper"]/div[1]/a/span').click()

    time.sleep(50)
    
    pyautogui.hotkey('alt','f4')

    pyautogui.hotkey('ENTER')
    
    time.sleep(3)

def analyzes_traseiro(traseiro): #o traseiro é o max file
     
    # Ler o arquivo baixado
    df2 = pd.read_excel(traseiro,header = [1])

    # Criando a coluna "Observacoes", no excel Puffer_CuboTraseiro.xlsx, para adicionar informações sobre o cubo
    df2["Observacoes"] = ""

    # ARQUIVO PRODUTOS QUE ESTÃO NA LINHA 

    # Lendo o arquivo TLLPontosNPsEIXO.txt
    df = pd.read_csv (r"R:\FTP\PLV\TLLPontosNPsEIXO.txt",delimiter = "|",header = None)

    # Tira as linhas que contém valor vazio e atualiza o dataframe
    df2_new = df2.dropna().reset_index(drop=True) # função reset_index usada porque sem ela o index fica fora de ordem, pois o pandas pula o index do valor que era NaN (0,1,3...)

    # Pega o número de linhas do dataframe atualizado
    lenght = df2_new.shape[0]

    # Preenchendo as linhas do dataframe df2 com os cubos que estão sobrando
    for np in range (0,lenght):
        NP_PAI = df2_new.loc[np,"NP Pai"] # armazena na variável NP_PAI o NP pai da linha np do dataframe df2
        NP_PAI_df = df.loc[df[1] == int(NP_PAI)] # cria um dataframe apenas com as linhas em que o NP Pai da coluna 1 é igual ao NP_PAI
        if NP_PAI_df.empty:
            df2_new.loc[df2_new["NP Pai"] == NP_PAI,"Observacoes"] = "cubo sobrando" # coloca a observação na coluna "Observacoes" do dataframe df2_new
        else:
            lenght_NP_PAI_df = NP_PAI_df.shape[0] # pega o número de linhas desse novo dataframe
            for i in range (0,lenght_NP_PAI_df):
                status = NP_PAI_df[2].iloc[i] # armazena na variável "status" o status do cubo
                status = status[:14] # o comando [:14] pega apenas os primeiros 14 digítos da string, assim removendo os espaços depois da string
                if status == "Final de Linha":
                    df2_new.loc[df2_new["NP Pai"] == NP_PAI,"Observacoes"] = "cubo sobrando" # coloca a observação na coluna "Observacoes" do dataframe df2_new
                
    # Filtra quais linhas tem observação e cria um dataframe, depois cria uma coluna nova "Email"
    cubo_sobrando_df = df2_new[df2_new["Observacoes"] == "cubo sobrando"]
    cubo_sobrando_df = cubo_sobrando_df.reset_index(drop=True) # esse comando acerta a ordem dos index, pois sem ele o index fica igual ao do dataframe original "df2_new"
    df2_new["Email"] = ""

    # Faz um loop nas linhas do dataframe x, colocando a data e hora no campo Email
    lenght_cubo_sobrando_df = cubo_sobrando_df.shape[0]
    for i in range(0,lenght_cubo_sobrando_df):
        NP = cubo_sobrando_df.loc[i,"NP Pai"]
        date = get_date()
        tempo = get_time()
        date_time = date + "  " + tempo
        cubo_sobrando_df.loc[cubo_sobrando_df["NP Pai"] == NP,"Email"] = date_time

    aux = cubo_sobrando_df

    # Apagando as linhas do dataframe X que já estão no csv
    df_csv = pd.read_csv(r"C:\Users\RCAMINA\Desktop\CUBOS_SOBRANDO.csv")

    for i in range(0,lenght_cubo_sobrando_df):
        NP = cubo_sobrando_df.loc[i,"NP Pai"]
        df_cubo = df_csv.loc[df_csv["NP Pai"] == NP]
        if not df_cubo.empty:   # se tiver vazio é porque o NP sem cubo não está na tabela do CSV, logo deve ser adicionado na tabela
                cubo_sobrando_df = cubo_sobrando_df.drop(i) # deleta a linha que possui o NP que já está no csv
            
    cubo_sobrando_df = cubo_sobrando_df = cubo_sobrando_df.reset_index(drop=True)

    # Aqui fazemos o append do dataframe no csv
    cubo_sobrando_df.to_csv(r"C:\Users\RCAMINA\Desktop\CUBOS_SOBRANDO.csv", mode='a', index=False, header=False)

    # Apagando as linhas do csv que os cubos já foram tirados do buffer
    df_csv = pd.read_csv(r"C:\Users\RCAMINA\Desktop\CUBOS_SOBRANDO.csv")

    lenght_df_csv = df_csv.shape[0]

    for i in range(0,lenght_df_csv):
        NP = df_csv.loc[i,"NP Pai"]
        df_cubo = aux.loc[aux["NP Pai"] == NP]
        if df_cubo.empty:   # se estiver vazio é porque o cubo ja foi tirado do buffer
            df_csv = df_csv.drop(i) # deleta a linha do cubo que estava sobrando
                
    # Aqui fazemos o append do dataframe no csv
    df_csv.to_csv(r"C:\Users\RCAMINA\Desktop\CUBOS_SOBRANDO.csv", mode='w', index=False, header=True)

    # Lendo o csv
    df_csv = pd.read_csv(r"C:\Users\RCAMINA\Desktop\CUBOS_SOBRANDO.csv")
    
    return(df_csv)

def analyzes_dianteiro(dianteiro): # teria que ser o penúltimo arquivo baixado
     
    # Ler o arquivo baixado
    df3 = pd.read_excel(dianteiro,header = [1])

    # Criando a coluna "Observacoes", no excel Puffer_CuboTraseiro.xlsx, para adicionar informações sobre o cubo
    df3["Observacoes"] = ""

    # ARQUIVO PRODUTOS QUE ESTÃO NA LINHA 

    # Lendo o arquivo TLLPontosNPsEIXO.txt
    df = pd.read_csv (r"R:\FTP\PLV\TLLPontosNPsEIXO.txt",delimiter = "|",header = None)

    # Tira as linhas que contém valor vazio e atualiza o dataframe
    df3_new = df3.dropna().reset_index(drop=True) # função reset_index usada porque sem ela o index fica fora de ordem, pois o pandas pula o index do valor que era NaN (0,1,3...)

    # Pega o número de linhas do dataframe atualizado
    lenght = df3_new.shape[0]

    # Preenchendo as linhas do dataframe df2 com os cubos que estão sobrando
    for np in range (0,lenght):
        NP_PAI = df3_new.loc[np,"NP Pai"] # armazena na variável NP_PAI o NP pai da linha np do dataframe df2
        NP_PAI_df = df.loc[df[1] == int(NP_PAI)] # cria um dataframe apenas com as linhas em que o NP Pai da coluna 1 é igual ao NP_PAI
        if NP_PAI_df.empty:
            df3_new.loc[df3_new["NP Pai"] == NP_PAI,"Observacoes"] = "cubo sobrando" # coloca a observação na coluna "Observacoes" do dataframe df2_new
        else:
            lenght_NP_PAI_df = NP_PAI_df.shape[0] # pega o número de linhas desse novo dataframe
            for i in range (0,lenght_NP_PAI_df):
                status = NP_PAI_df[2].iloc[i] # armazena na variável "status" o status do cubo
                status = status[:14] # o comando [:14] pega apenas os primeiros 14 digítos da string, assim removendo os espaços depois da string
                if status == "Final de Linha":
                    df3_new.loc[df3_new["NP Pai"] == NP_PAI,"Observacoes"] = "cubo sobrando" # coloca a observação na coluna "Observacoes" do dataframe df2_new
                
    # Filtra quais linhas tem observação e cria um dataframe, depois cria uma coluna nova "Email"
    cubo_sobrando_df2 = df3_new[df3_new["Observacoes"] == "cubo sobrando"]
    cubo_sobrando_df2 = cubo_sobrando_df2.reset_index(drop=True) # esse comando acerta a ordem dos index, pois sem ele o index fica igual ao do dataframe original "df2_new"
    df3_new["Email"] = ""

    # Faz um loop nas linhas do dataframe x, colocando a data e hora no campo Email
    lenght_cubo_sobrando_df2 = cubo_sobrando_df2.shape[0]
    for i in range(0,lenght_cubo_sobrando_df2):
        NP = cubo_sobrando_df2.loc[i,"NP Pai"]
        date = get_date()
        tempo = get_time()
        date_time = date + "  " + tempo
        cubo_sobrando_df2.loc[cubo_sobrando_df2["NP Pai"] == NP,"Email"] = date_time

    aux = cubo_sobrando_df2

    # Apagando as linhas do dataframe X que já estão no csv
    df_csv = pd.read_csv(r"C:\Users\RCAMINA\Desktop\teste_projeto2.csv")

    for i in range(0,lenght_cubo_sobrando_df2):
        NP = cubo_sobrando_df2.loc[i,"NP Pai"]
        df_cubo = df_csv.loc[df_csv["NP Pai"] == NP]
        if not df_cubo.empty:   # se tiver vazio é porque o NP sem cubo não está na tabela do CSV, logo deve ser adicionado na tabela
                cubo_sobrando_df2 = cubo_sobrando_df2.drop(i) # deleta a linha que possui o NP que já está no csv
            
    cubo_sobrando_df2= cubo_sobrando_df2 = cubo_sobrando_df2.reset_index(drop=True)

    # Aqui fazemos o append do dataframe no csv
    cubo_sobrando_df2.to_csv(r"C:\Users\RCAMINA\Desktop\teste_projeto2.csv", mode='a', index=False, header=False)

    # Apagando as linhas do csv que os cubos já foram tirados do buffer
    df_csv2 = pd.read_csv(r"C:\Users\RCAMINA\Desktop\teste_projeto2.csv")

    lenght_df_csv2 = df_csv2.shape[0]

    for i in range(0,lenght_df_csv2):
        NP = df_csv2.loc[i,"NP Pai"]
        df_cubo = aux.loc[aux["NP Pai"] == NP]
        if df_cubo.empty:   # se estiver vazio é porque o cubo ja foi tirado do buffer
            df_csv2 = df_csv2.drop(i) # deleta a linha do cubo que estava sobrando
                
    # Aqui fazemos o append do dataframe no csv
    df_csv2.to_csv(r"C:\Users\RCAMINA\Desktop\teste_projeto2.csv", mode='w', index=False, header=True)

    # Lendo o csv
    df_csv2 = pd.read_csv(r"C:\Users\RCAMINA\Desktop\teste_projeto2.csv")
    
    return(df_csv2)

def get_date():
    today = datetime.today()
    date_str = today.strftime("%d/%m/%Y")
    return date_str

# retorna o horario atual
def get_time():
    time = datetime.now().time()
    time_str = time.strftime("%H:%M")
    return time_str

def send_email_traseiro(df_csv):
    # enviar email com os dados do csv

    outlook = win32.Dispatch("outlook.application") # criando integração com o outlook

    email = outlook.CreateItem(0) # criando um email

    email.To = "raul_lopes.camina@daimler.com ; anderson.radael@daimlertruck.com ; andre.lopes@daimlertruck.com ; alex.almeida@daimler.com; franco.willian@daimlertruck.com"  # especificando o destino do email

    email.Subject = "Cubos Sobrando no Buffer" # especificando o assunto no email

    email.HTMLBody = '''<h3>Segue tabela dos cubos que estão sobrando no buffer:</h3>
                   {}'''.format(df_csv.to_html()) #com esse comando .format, é possível enviar um dataframe no modo tabela por email

    email.Send()

def send_email_dianteiro(df_csv2):
    # enviar email com os dados do csv

    outlook = win32.Dispatch("outlook.application") # criando integração com o outlook

    email = outlook.CreateItem(0) # criando um email

    email.To = "raul_lopes.camina@daimler.com ; valdemar.castro@daimler.com ; ed_carlos.cupertino@daimler.com ; anderson.radael@daimlertruck.com ; rodrigo.gandolpho@daimlertruck.com; franco.willian@daimlertruck.com"  # especificando o destino do email

    email.Subject = "Cubos Sobrando no Buffer" # especificando o assunto no email

    email.HTMLBody = '''<h3>Segue tabela dos cubos que estão sobrando no buffer:</h3>
                   {}'''.format(df_csv2.to_html()) #com esse comando .format, é possível enviar um dataframe no modo tabela por email

    email.Send()

# MAIN SCOPE

get_files() # baixando os arquivos

folder_path = r"C:\Users\RCAMINA\Downloads"
file_type = r'\*xlsx'

files = glob.glob(folder_path + file_type)
traseiro = max(files, key=os.path.getctime) # colocando na variável o arquivo do traseiro
dianteiro = sorted(files, key=os.path.getctime) # ordena os arquivos baixados de acordo com a data e hora que foi baixado
dianteiro = dianteiro[-2] # colocando na variável o arquivo do dianteiro; pega o 2º arquivo, pois o 1º foi o do traseiro

df = analyzes_traseiro(traseiro)
send_email_traseiro(df)
df2 = analyzes_dianteiro(dianteiro)
send_email_dianteiro(df2)
