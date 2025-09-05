from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from collections import Counter


def extrair_produtos(driver, num_paginas=3):
    dados = []
    wait = WebDriverWait(driver, 15)
    produto_id = 1


    try:
        pagina1 = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="page-1"]'))
        )
        try:
            link = pagina1.find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", link)
        except:
            driver.execute_script("arguments[0].click();", pagina1)
        wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="page-1"].Paginator_selected__sOsqO'))
        )
        time.sleep(2)
    except:
        pass

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

            try:
                link_element = card.find_element(By.XPATH, ".//ancestor::a[@class='ProductCard_ProductCard_Inner__gapsh']")
                link = link_element.get_attribute("href")
            except:
                link = "N/A"

            dados.append({
                "ID": produto_id,
                "Produto": titulo,
                "Valor": preco,
                "Link": link
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

    todos_produtos = produtos_relevantes + produtos_menor_preco + produtos_melhor_avaliado
    contador = Counter([p["Produto"] for p in todos_produtos])
    ranking = contador.most_common()
    top5_produtos = ranking[:5]  # [(nome, freq), ...]
    
    top5_links = []
    for nome, freq in top5_produtos:
        for p in todos_produtos:
            if p["Produto"] == nome:
                top5_links.append((nome, p["Link"], freq))
                break


    driver = webdriver.Chrome(options=webdriver.ChromeOptions().add_argument("--start-maximized"))
    wait = WebDriverWait(driver, 15)

    detalhes_top5 = []

    for produto, link, freq in top5_links:  # freq = quantas vezes apareceu
        driver.get(link)
        time.sleep(2)

        try:
            # Captura a seção de descrição do produto
            descricao_div = driver.find_element(
                By.CSS_SELECTOR, '[data-testid="detailsSection-overflow"]'
            )
            paragrafos = descricao_div.find_elements(By.TAG_NAME, "p")
            descricao_texto = " ".join([p.text for p in paragrafos])
        except:
            descricao_texto = "N/A"

        detalhes_top5.append({
            "Produto": produto,
            "Frequencia": freq,
            "Descricao": descricao_texto
        })

    # Salva em Excel
    df_top5 = pd.DataFrame(detalhes_top5)
    df_top5.to_excel("Produtos_Top5.xlsx", index=False)
    print("Arquivo Produtos_Top5.xlsx salvo com sucesso!")


if __name__ == "__main__":
    main()
