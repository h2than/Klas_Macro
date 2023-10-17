from selenium import webdriver
import chromedriver_autoinstaller

url = 'https://www.google.com/'

# Chromedriver Auto Install


options = webdriver.ChromeOptions()

driver = webdriver.Chrome(options=options)

driver.get(url)
