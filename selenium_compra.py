import time, smtplib
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait	
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#servicedesk_email = "dft11.servicedesk@gmail.com"
#servicedesk_password = "?"

servicedesk_email = "darkneroh@gmail.com"
servicedesk_password = "?"

def comprar(url, site):

	browser.get("https://secure.{}.com.br/customer/account/login/".format(site))

	#login e senha para entrar no site
	#espera 10 segundos pelo elemento
	WebDriverWait(browser,10).until(EC.visibility_of_element_located((By.XPATH,"""//*[@id="LoginForm_email"]""")))

	browser.find_element_by_xpath("""//*[@id="LoginForm_email"]""").send_keys(servicedesk_email)
	time.sleep(0.5)
	browser.find_element_by_xpath("""//*[@id="LoginForm_password"]""").send_keys(servicedesk_password)
	time.sleep(1)
	browser.find_element_by_xpath("""//*[@id="customer-account-login"]""").click()
	time.sleep(4)

	#abre um item
	browser.get(url)
	time.sleep(4)

	#clicar em comprar
	WebDriverWait(browser,10).until(EC.visibility_of_element_located((By.XPATH,"""//*[@id="add-to-cart"]/button"""))).click()
	time.sleep(4)

	#ir para o carrinho
	browser.get("https://secure.{}.com.br/cart/".format(site))

	#clicar em finalizar compra
	browser.find_element_by_xpath("""//*[@id="button-finalize-order-1"]""").click()
	time.sleep(4)

	#selecionar o boleto
	browser.find_element_by_xpath("""//*[@id="boleto"]""").click()
	time.sleep(1)

	#clicar em finalizar compra
	browser.find_element_by_xpath("""//*[@id="btn_finalize_order"]""").click()


def enviar_email(erro, site):

	de = "matheuspsilva222@gmail.com"
	para = servicedesk_email
	senha = "?"

	msg = MIMEMultipart()
	msg ['From'] = de
	msg['To'] = para
	msg['Subject'] = "Erro na compra"

	body = "Ocorreu um erro na compra do site [{1}]:\n({0})".format(erro, site)

	# anexa imagem
	image = MIMEImage(open('error.png','rb').read())
	msg.attach(image)
	# anexa o texto
	msg.attach(MIMEText(body, 'plain'))

	s = smtplib.SMTP('smtp.gmail.com', 587)
	s.starttls()
	s.login(de, senha)
	s.sendmail(de, para, msg.as_string())
	print('mensagem enviada')
	s.quit()


# AQUI COMECA RODAR O CODIGO
if __name__ == "__main__":

	chrome_options = webdriver.ChromeOptions()

	prefs = {
	"profile.default_content_setting_values.notifications": 2,
	#"deviceName": "Galaxy S5" 
	}

	chrome_options.add_experimental_option("prefs", prefs)
	browser = webdriver.Chrome(options=chrome_options)
	browser.set_page_load_timeout(60)
	browser.maximize_window()

	#lista de skus para buscar
	skus = []

	url = [
	'https://www.tricae.com.br/Kit-3pcs-Pimpolho-Menina-Lisa-Branco-4758323.html', 
	'https://www.kanui.com.br/Meia-Zero-Full-Zero-Amarelo-3124630.html', 
	'https://www.dafiti.com.br/Meia-Zero-Skull-Line-Patt-Preto-3015915.html?hash_configuration=0c513651-0f44-45e2-a13c-d5ebeba55f6c'
	]

	sites = ['tricae', 'kanui', 'dafit']

	# tente
	try:
		for url,site in zip(url,sites):
			comprar(url, site)

	# exceto erro
	except Exception as erro:
		browser.save_screenshot('error.png')
		enviar_email(erro, site)
		# força o erro e interrupção do script
		raise
