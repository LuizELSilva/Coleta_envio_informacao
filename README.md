# Protótipo de automação em python

Protótipo de automação: Esse projeto pega informações de cotação para a moeda real, as moedas consultadas são Dólar, Euro, Iene, Ouro, essas informações são coletadas a partir de um navegador usando a biblioteca Selenium, após coletar as informações o programa envia um E-mail com uma tabela em xlsx.

Biblioteca Selenium e sleep para navegação em sites, uso o sleep para dar mais tempo no carregamento do site (Tal recurso vai variar de máquina para máquina). 

Bibliotecas Pandas, openpyxl e numpy para leitura da planilha.

Bibliotecas para e-mail, MiMEText, MiMEBase, Encoders, MIMEMultipart e smtplib. Os quatro primeiros são para a construção do texto e arquivo em anexo. O último é para conectar com o servidor e enviar o E-mail.

Estou usando a estrutura do Gmail para esse projeto, ou seja, precisamos de uma senha de acesso que só pode ser pega configurando o Senhas de app do google.

Por fim, nesse projeto estou utilizando o ChromeDriver, com a ajuda desse arquivo o selenium controla a máquina navegando pelos sites.


Projeto é feito em no python 3.11




