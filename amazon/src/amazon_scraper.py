# Amazon - Web Scraping Automation

### Imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import pandas as pd
from pathlib import Path


### Navigating to the page
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

driver.get("https://www.amazon.com.br")
driver.maximize_window()

wait = WebDriverWait(driver, timeout=10)
menu = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@id='nav-hamburger-menu']")))
menu.click()

category_button = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@class='hmenu-item' and text()='Produtos em alta']")))
category_button.click()

ul = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[contains(@class, 'zg-browse-group')]")))

categories = ul.find_elements(By.XPATH, ".//li")

category_names = [name.text for name in categories]
print(category_names)


### Functions
def clicking_on_categories(name, initial_page='Qualquer departamento'):

    ul = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[contains(@class, 'zg-browse-group')]")))
    
    try:
        category = ul.find_element(By.XPATH, f".//a[contains(., '{name}')]")
        category.click()
        print(f"Clicked on category: {name}\n")
    except:
        ul = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[contains(@class, 'zg-browse-root')]")))
        qualquer_departamento = ul.find_element(By.XPATH, f".//a[contains(., '{initial_page}')]")
        qualquer_departamento.click()
        print(f"Returning to the initial page")
        ul = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[contains(@class, 'zg-browse-group')]")))
        category = ul.find_element(By.XPATH, f".//a[contains(., '{name}')]")
        category.click()
        print(f"Clicked on category: {name}\n")

def scrape_items(items, name):
    
    c = []
    r = []
    d = []
    p = []
    l = []

    while True:
        i = 0
        while i < len(items):
            action = ActionChains(driver)
            action.scroll_to_element(items[i])
            action.perform()
            sleep(1)

            c.append(name)
            
            ranking = int(items[i].find_element(By.XPATH, ".//span[@class='zg-bdg-text']").text.replace("#",''))
            r.append(ranking)
            
            description = items[i].find_element(By.XPATH, ".//a[@class='a-link-normal aok-block']//span").text
            d.append(description)

            try:
                price_raw = items[i].find_element(By.XPATH, ".//span[@class='a-size-base a-color-price']").text
            except:
                price_raw = "-"
            p.append(price_raw)

            link = items[i].find_element(By.XPATH, ".//a[@class='a-link-normal aok-block']").get_attribute("href")  
            l.append(link)

            print(f"{ranking} || {description} | {price_raw} | {link}")

            page = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'p13n-desktop-grid')]")))
            items = page.find_elements(By.XPATH, "//li[@class='zg-no-numbers']")
            i += 1
        
        try:
            button = wait.until(EC.visibility_of_element_located((By.XPATH, "//li[@class='a-last']")))
            button.click()
            sleep(2)
            page = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'p13n-desktop-grid')]")))
            items = page.find_elements(By.XPATH, "//li[@class='zg-no-numbers']")
            i = 0
        except:
            print("[INFO] No more pages found.\n")
            break

    return c, r, d, p, l

def Saving_DataFrames(dict_dfs, directory,  option=1):
    if option == 1:
        path = directory / "Consolidated.xlsx"
        df = pd.concat(dfs.values(), ignore_index=True)
        df.to_excel(path, index=False)

    elif option == 2:
        for category, df in dict_dfs.items():
            path = directory / f'DataFrame_{category}.xlsx'
            df.to_excel(path, index=False)

    elif option == 3:
        path = directory / "Consolidated_sheets.xlsx"
        with pd.ExcelWriter(path) as writer:
            for category, df in dict_dfs.items():
                df.to_excel(writer, sheet_name=category[:31], index=False)

    else:
        print("Please, select a valid options: 1 , 2 or 3")


### Web Scraping

if __name__ == "__main__":

    dfs = {}
    for name in category_names:

        clicking_on_categories(name)
        
        page = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'p13n-desktop-grid')]")))
        items = page.find_elements(By.XPATH, "//li[@class='zg-no-numbers']")
        
        category, ranking, description , price, link = scrape_items(items, name)

        dfs[name] = pd.DataFrame({
            'Category':category,
            'Rank':ranking,
            'Description':description,
            'Product Price':price,
            'Link':link
        })

    directory = Path(__file__).parent.parent / 'output'
    Saving_DataFrames(dfs, directory, ) # Saving Modes --> 1 (default): Consolidated | 2: One file per category | 3: One sheet per category
