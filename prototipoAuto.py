
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from time import sleep

'''
Instalado openpyxl
instalado pandas
instalado numpy
'''
def enviar_email():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    
    destinatario = "destinatario@gmail.com"
    assunto = "E-mail com anexo teste"
    mensagem = "resultado final."
    anexo = "C:\Caminho\Do\Anexo\Planilha_Teste_valor.xlsx"
    
    
    
    # 1 - Configurar informações do remetente e do servidor de e-mail
    remetente = "remetente@gmail.com"
    senha = "SenhaAppGoogle"
    servidor_sxmtp = "smtp.gmail.com"
    porta_smtp = 587

    # 2 - Criar objeto de mensagem multipart
    msg = MIMEMultipart()

    # 3 - Configurar os campos do cabeçalho do e-mail
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # 4 - Adicionar corpo do e-mail
    msg.attach(MIMEText(mensagem, 'plain'))

    # 5 - Anexar o arquivo 
    with open(anexo, 'rb') as arquivo:
        ArqAnexo = MIMEBase('application', 'octet-stream')
        ArqAnexo.set_payload(arquivo.read())
        encoders.encode_base64(ArqAnexo)
        ArqAnexo.add_header('Content-Disposition', f'attachment; filename = Planilha_Teste_valor.xlsx')
        arquivo.close()
        
        msg.attach(ArqAnexo)

    # 6 -Conectar ao servidor e enviar o e-mail
    with smtplib.SMTP(servidor_sxmtp, porta_smtp) as smtp:
        smtp.starttls()
        smtp.login(remetente, senha)
        smtp.send_message(msg)
        
        smtp.quit()

def Pegando_informacoes(CotLista):
    for c in range (0,3):
        driver.get('https://www.google.com.br/')
        sleep(1)
        Navegador = driver.find_element(By.XPATH, '//*[@id="APjFqb"]')
        Navegador.send_keys('Cotação ', cotacoes[c])
        sleep(1)
        Navegador = driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(Keys.ENTER)
        sleep(1)
        CotLista[c] = driver.find_element(By.XPATH,'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
        
        sleep(1)
        print(CotLista[c])
        
    sleep(1)
    driver.get('https://www.melhorcambio.com/ouro-hoje')

    sleep(1)
    CotLista[3] = driver.find_element(By.XPATH,'//*[@id="comercial"]').get_attribute("value")
    CotLista[3] = CotLista[3].replace(",",".")
    print(CotLista[3],' Por grama de ouro')

    sleep(1)
    driver.quit()

def Atualiza_Planilha():
    tabela = pd.read_excel("Planilha_Teste_valor.xlsx")
    print(tabela)
    sleep(1)

    for c in range (0,4):
        tabela.loc[0 ,cotacoes[c]] = float(CotLista[c])
    print(tabela)

    tabela.to_excel("Planilha_Teste_valor.xlsx", index=False)

    sleep(1)

# Definindo Caminho no Driver 

CaminhoDriver = "c:\Caminho\do\Driver\chromedriver.exe" # Inportando arquivo WebDriver
driver = webdriver.Chrome(CaminhoDriver)

# 1 - Definindo variáveis
Cotacao_Dolar = cotacao_Euro = cotacao_Iene = cotacao_Ouro = 0

CotLista = [Cotacao_Dolar, cotacao_Euro ,cotacao_Iene, cotacao_Ouro]

cotacoes = ['Dólar','Euro','Iene','Ouro']

# 2 - Pegando informações no google
Pegando_informacoes(CotLista)

# 3 - Lendo tabela 
Atualiza_Planilha()

# 4 - Puxando função e-mail
enviar_email()
