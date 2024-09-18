from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

def configurar_navegador():
    # Baixa e configura o ChromeDriver
    navegador = webdriver.Chrome(ChromeDriverManager().install())
    navegador.get("https://vuetifyjs.com/en/components/date-pickers/#usage")
    print(navegador.title)

    navegador.quit()

if __name__ == '__main__':
    configurar_navegador()
