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


service = Service(EdgeChromiumDriverManager().install())

driver = webdriver.Edge(service=service)

driver.get("https://www.google.com/recaptcha/api2/demo")

iframe_element = driver.find_element(By.XPATH, "//iframe[@title='reCAPTCHA']")

try:

    sleep(1)

    solver = RecaptchaSolver(driver=driver)

    solver.click_recaptcha_v2(iframe=iframe_element)

except Exception as e:

    # captcha = driver.find_element(By.CLASS_NAME, "recaptcha-checkbox-border")

    # captcha.click()

    print(e)

sleep(2)

driver.switch_to.default_content()

