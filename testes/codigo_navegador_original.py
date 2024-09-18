from selenium import webdriver

# As duas formas abaixo dá certo, desde que a versao do chromedriver seja a mesma versao que a do chrome - ainda que seja de teste;
# navegador = webdriver.Chrome(r'C:\Users\erics\Downloads\120.0.6099.109\chrome-win64 120.0.6099.109\chrome-win64\chrome.exe') - o site nao abriu
# Na forma abaixo o site abriu;
navegador = webdriver.Chrome(r'C:\Users\erics\Downloads\120.0.6099.109\chromedriver-win64 120.0.6099.109\chromedriver-win64\chromedriver')

navegador.get("https://vuetifyjs.com/en/components/date-pickers/#usage")