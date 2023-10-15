from selenium_recaptcha_solver import RecaptchaSolver
from selenium import webdriver
from time import sleep
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.wait import WebDriverWait


headers = {
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'es-419,es;q=0.9,es-ES;q=0.8,en;q=0.7,en-GB;q=0.6,en-US;q=0.5',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Host': 'ge.ch',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36 Edg/117.0.2045.60',
    'sec-ch-ua': '"Microsoft Edge";v="117", "Not;A=Brand";v="8", "Chromium";v="117"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}

regions = [
    "Aire-la-Ville (1)",
    "Anières (2)",
    "Avully (3)",
    "Avusy (4)",
    "Bardonnex (5)",
    "Bellevue (6)",
    "Bernex (7)",
    "Carouge (8)",
    "Cartigny (9)",
    "Céligny (10)",
    "Chancy (11)",
    "Chêne-Bougeries (12)",
    "Chêne-Bourg (13)",
    "Choulex (14)",
    "Collex-Bossy (15)",
    "Collonge-Bellerive (16)",
    "Cologny (17)",
    "Confignon (18)",
    "Corsier (19)",
    "Dardagny (20)",
    "Genève-Cité (21)",
    "Genève-Eaux-Vives (22)",
    "Genève-Petit-Saconnex (23)",
    "Genève-Plainpalais (24)",
    "Genthod (25)",
    "Grand-Saconnex (26)",
    "Gy (27)",
    "Hermance (28)",
    "Jussy (29)",
    "Laconnex (30)",
    "Lancy (31)",
    "Meinier (32)",
    "Meyrin (33)",
    "Onex (34)",
    "Perly-Certoux (35)",
    "Plan-les-Ouates (36)",
    "Pregny-Chambésy (37)",
    "Presinge (38)",
    "Puplinge (39)",
    "Russin (40)",
    "Satigny (41)",
    "Soral (42)",
    "Thônex (43)",
    "Troinex (44)",
    "Vandoeuvres (45)",
    "Vernier (46)",
    "Versoix (47)",
    "Veyrier (48)"
]


def run():

    pass

def save_to_excel():

    pass

def scrape(heads, regions):

    service = Service(EdgeChromiumDriverManager().install())

    driver = webdriver.Edge(service=service)

    edge_options = Options()

    for key, value in heads.items():
            
            edge_options.add_argument(f'--{key}={value}')

    driver.get("https://ge.ch/terextraitfoncier/immeuble.aspx")

    wait = WebDriverWait(driver, 30).until(lambda x: x.find_element(By.ID, "ddlCommune"))

    region_input = driver.find_element(By.ID, "ddlCommune")

    region_input.click()

    sleep(1)

    select_region = driver.find_element(By.XPATH, f"//option[text()='{regions[0]}']")

    select_region.click()

    sleep(1)

    plot_input = driver.find_element(By.ID, "tbParcelle")

    sleep(0.5)

    plot_input.send_keys("150")

    sleep(1)

    iframe_element = driver.find_element(By.XPATH, "//iframe[@title='reCAPTCHA']")

    # driver.switch_to.frame(iframe_element)

    try:

        sleep(1)

        solver = RecaptchaSolver(driver=driver)

        solver.click_recaptcha_v2(iframe=iframe_element)

    except Exception as e:

        # captcha = driver.find_element(By.CLASS_NAME, "recaptcha-checkbox-border")

        # captcha.click()

        pass

    sleep(2)

    driver.switch_to.default_content()

    sleep(5)

    ex = driver.find_element(By.ID, "btnSubmit")

    ex.click()

    sleep(15)

    owners = driver.find_element(By.XPATH, "//select[@name='select_proprietaires']")

    owners = owners.find_elements(By.TAG_NAME, "option")

    for owner in owners:

        print(owner.text)



if __name__ == "__main__":

    scrape(headers, regions)