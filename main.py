from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd


def extrair_produtos(driver, num_paginas=3):
    dados = []
    wait = WebDriverWait(driver, 15)
    produto_id = 1

    for pagina in range(1, num_paginas + 1):
        wait.until(EC.presence_of_all_elements_located(
            (By.CLASS_NAME, "ProductCard_ProductCard_Body__bnVUn")
        ))
        time.sleep(2)

        cards = driver.find_elements(By.CLASS_NAME, "ProductCard_ProductCard_Body__bnVUn")

        for card in cards:
            try:
                descricao = card.find_element(By.CLASS_NAME,
                    "ProductCard_ProductCard_Description__YNSOx")
                titulo = descricao.find_element(By.TAG_NAME, "h2").text.strip()
            except:
                titulo = "N/A"

            try:
                preco = card.find_element(By.CLASS_NAME,
                    "Text_MobileHeadingS__HEz7L").text.strip()
            except:
                preco = "N/A"

            dados.append({
                "ID": produto_id,
                "Produto": titulo,
                "Valor": preco
            })
            produto_id += 1

        if pagina < num_paginas:
            try:
                next_page = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, f'[data-testid="page-{pagina+1}"] a'))
                )
                driver.execute_script("arguments[0].click();", next_page)
                wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, f'[data-testid="page-{pagina+1}"].Paginator_selected__sOsqO'))
                )
                time.sleep(2)
            except Exception as e:
                print(f"Não consegui ir para a página {pagina+1}: {e}")
                break

    return dados


def aplicar_filtro(driver, wait, texto_opcao):
    """Seleciona a opção de ordenação no <select> pelo texto visível"""
    select_element = wait.until(
        EC.presence_of_element_located((By.ID, "orderBy"))
    )
    select = Select(select_element)
    select.select_by_visible_text(texto_opcao)
    time.sleep(3)


def salvar_excel(lista, nome_arquivo):
    df = pd.DataFrame(lista)
    df.to_excel(nome_arquivo, index=False)
    print(f"Arquivo salvo: {nome_arquivo}")


def main():
    # Configura o Selenium (Chrome)
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    driver.get("https://www.zoom.com.br")

    wait = WebDriverWait(driver, 15)

    search_box = wait.until(EC.presence_of_element_located((By.ID, "searchInput")))
    search_box.send_keys("Smart TV")

    search_button = driver.find_element(By.CLASS_NAME, "SearchInput_SearchButton__Cg9kC")
    search_button.click()

    produtos_relevantes = extrair_produtos(driver, num_paginas=3)
    salvar_excel(produtos_relevantes, "Produtos_Relevantes.xlsx")

    aplicar_filtro(driver, wait, "Menor preço")
    produtos_menor_preco = extrair_produtos(driver, num_paginas=3)
    salvar_excel(produtos_menor_preco, "Produtos_Menor_Preco.xlsx")

    aplicar_filtro(driver, wait, "Melhor avaliado")
    produtos_melhor_avaliado = extrair_produtos(driver, num_paginas=3)
    salvar_excel(produtos_melhor_avaliado, "Produtos_Melhor_Avaliado.xlsx")

    driver.quit()


if __name__ == "__main__":
    main()
