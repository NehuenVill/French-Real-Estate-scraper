from selenium_recaptcha_solver import RecaptchaSolver
from selenium import webdriver
from time import sleep
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.wait import WebDriverWait
import openpyxl
import json
from selenium_recaptcha_solver import DelayConfig

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

def save_to_excel(plot, region, owner):

    try:
        # Load an existing workbook or create a new one if it doesn't exist
        workbook = openpyxl.load_workbook("plots.xlsx")

        # Select the default active sheet (usually the first sheet)
        sheet = workbook.active

        # Add data to the sheet in the specified order
        sheet.append([plot, region, owner])

        # Save the workbook with the added data
        workbook.save("plots.xlsx")

        print(f"Data ({plot}, {region}, {owner}) added to the Excel file.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def save_progress(region, plot):

    try:
        with open("progress.json", "r") as file:
            progress = json.load(file)
    except FileNotFoundError:
        # If the file doesn't exist, initialize an empty progress dictionary
        progress = {}

    # Check if the region already exists in the progress data
    if region not in progress:
        # If the region doesn't exist, create a new list for it
        progress[region] = []

    # Append the plot number to the region's list
    progress[region].append(plot)

    # Save the updated progress data back to the JSON file
    with open("progress.json", "w") as file:
        json.dump(progress, file, indent=4)

def scrape(heads, regions):

    script = "Object.defineProperty(navigator, 'webdriver', {get: () => false})"

    edge_options = Options()

    # edge_options.add_argument("--headless")
    
    edge_options.add_argument("window-size=1280,1000")

     # Adding argument to disable the AutomationControlled flag 
    edge_options.add_argument("--disable-blink-features=AutomationControlled") 
    
    # Exclude the collection of enable-automation switches 
    edge_options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
    
    # Turn-off userAutomationExtension 
    edge_options.add_experimental_option("useAutomationExtension", False) 

    edge_options.add_experimental_option(
        'excludeSwitches', [
        'disable-extensions',
        'disable-default-apps',
        'disable-component-extensions-with-background-pages',
    ])

    for key, value in heads.items():
            
            edge_options.add_argument(f'--{key}={value}')

    service = Service(EdgeChromiumDriverManager().install())

    driver = webdriver.Edge(service=service, options=edge_options)

    driver.get("https://ge.ch/terextraitfoncier/immeuble.aspx")

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": script})

    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    driver.execute_script("""
    Object.defineProperty(navigator, 'plugins', {
        // This just needs to have `length > 0` for the current test,
        // but we could mock the plugins too if necessary.
        get: () => [1, 2, 3, 4, 5, 6, 7, 8],
    });

    """)

    try:
        with open("progress.json", "r") as file:
            progress = json.load(file)
    except FileNotFoundError:
        progress = {}

    regions_scraped, last_region = [], ''

    try:

        regions_scraped, last_region = list(progress.keys())[:-1], list(progress.keys())[-1] 

    except Exception:

        pass

    for region in regions:

        if region in regions_scraped:

            continue

        if region == last_region:

            for i in range(max(progress[last_region])+1, 10_001):

                wait = WebDriverWait(driver, 30).until(lambda x: x.find_element(By.ID, "ddlCommune"))

                region_input = driver.find_element(By.ID, "ddlCommune")

                region_input.click()

                sleep(1)

                select_region = driver.find_element(By.XPATH, f"//option[text()='{region}']")

                select_region.click()

                sleep(1)

                plot_input = driver.find_element(By.ID, "tbParcelle")

                sleep(0.5)

                plot_input.clear()

                sleep(0.5)

                plot_input.send_keys(f"{i}")

                sleep(1)

                try:

                    iframe_element = driver.find_element(By.XPATH, "//iframe[@title='reCAPTCHA']")

                    try:

                        sleep(1)

                        solver = RecaptchaSolver(driver=driver)

                        solver.click_recaptcha_v2(iframe=iframe_element)

                    except Exception as e:

                        # captcha = driver.find_element(By.CLASS_NAME, "recaptcha-checkbox-border")

                        # captcha.click()

                        pass

                except Exception:

                    pass

                # driver.switch_to.frame(iframe_element)

                sleep(1)

                driver.switch_to.default_content()

                ex = driver.find_element(By.ID, "btnSubmit")

                ex.click()

                save_progress(region, i)

                try:

                    wait = WebDriverWait(driver, 5).until(lambda x: x.find_element(By.XPATH, "//select[@name='select_proprietaires']"))

                except Exception as e:

                    continue

                try:

                    owners = driver.find_element(By.XPATH, "//select[@name='select_proprietaires']")

                    owners = owners.find_elements(By.TAG_NAME, "option")

                except Exception as e:

                    driver.back()

                    continue

                for owner in owners:

                    save_to_excel(i, region.split(" ")[0], owner)

                driver.back()

        else:

            for i in range(0, 10_001):

                wait = WebDriverWait(driver, 30).until(lambda x: x.find_element(By.ID, "ddlCommune"))

                region_input = driver.find_element(By.ID, "ddlCommune")

                region_input.click()

                sleep(1)

                select_region = driver.find_element(By.XPATH, f"//option[text()='{region}']")

                select_region.click()

                sleep(1)

                plot_input = driver.find_element(By.ID, "tbParcelle")

                sleep(0.5)

                plot_input.clear()

                sleep(0.5)

                plot_input.send_keys(f"{i}")

                sleep(1)

                try:

                    iframe_element = driver.find_element(By.XPATH, "//iframe[@title='reCAPTCHA']")

                    try:

                        sleep(1)

                        solver = RecaptchaSolver(driver=driver)

                        solver.click_recaptcha_v2(iframe=iframe_element)

                    except Exception as e:

                        # captcha = driver.find_element(By.CLASS_NAME, "recaptcha-checkbox-border")

                        # captcha.click()

                        pass

                except Exception:

                    pass

                # driver.switch_to.frame(iframe_element)

                sleep(1)

                driver.switch_to.default_content()

                ex = driver.find_element(By.ID, "btnSubmit")

                ex.click()

                save_progress(region, i)

                try:

                    wait = WebDriverWait(driver, 5).until(lambda x: x.find_element(By.XPATH, "//select[@name='select_proprietaires']"))

                except Exception as e:

                    continue

                try:

                    owners = driver.find_element(By.XPATH, "//select[@name='select_proprietaires']")

                    owners = owners.find_elements(By.TAG_NAME, "option")

                except Exception as e:

                    driver.back()

                    continue

                for owner in owners:

                    save_to_excel(i, region.split(" ")[0], owner.text)

                driver.back()


def run():

    while True:

        try:

            scrape(headers, regions)

            break

        except Exception as e:

            print(e)

            print(e.with_traceback(None))

            sleep(10_800)


if __name__ == "__main__":

    run()