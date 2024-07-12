from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import requests as rq
import xlwings as xw
import time
from time import sleep
from urllib.parse import urlparse, parse_qs
from pyotp import TOTP

# Path to your Edge WebDriver, adjust if necessary
edge_driver_path = 'D:\stocks web scraping\msedgedriver.exe'  # Example: 'C:/path/to/msedgedriver.exe'

# # You can add more automation steps here if needed

wb = xw.Book('TradeToolsUpstox.xlsx')
crd = wb.sheets("Cread")
api_key = crd['B1'].value
secret_key = crd['B2'].value
r_url = crd['B3'].value
totp_key = crd['B4'].value
mobile_no = crd['B5'].value
pin = crd['B6'].value
auth_url = f'https://api.upstox.com/v2/login/authorization/dialog?response_type=code&client_id={api_key}&redirect_uri={r_url}'


# # Set up the Edge options
edge_options = Options()
# # Set up the WebDriver service
service = Service(edge_driver_path)

# # Initialize the WebDriver with the options and service
driver = webdriver.Edge(service=service, options=edge_options)

# # Open YouTube
# driver.get(auth_url)

# Waiting for the Mobile Number Input Field and Entering the Mobile Number
# wait = WebDriverWait(driver, 3)
# wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(mobile_no)

driver.get(auth_url)
wait = WebDriverWait(driver,3)

wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(mobile_no)
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="getOtp"]'))).click()
totp = TOTP(totp_key).now()

sleep(2)
wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="otpNum"]'))).send_keys(totp)
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="continueBtn"]'))).click()
sleep(2)

wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pinCode"]'))).send_keys(pin)
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pinContinueBtn"]'))).click()
sleep(2)

token_url = driver.current_url
parsed = urlparse(token_url)
driver.close()
code = parse_qs(parsed.query)['code'][0]

url = 'https://api.upstox.com/v2/login/authorization/token'
headers = {
    'accept': 'application/json',
    'Api-Version': '2.0',
    'Content-Type': 'application/x-www-form-urlencoded'}

data = {
    'code': code,
    'client_id': api_key,
    'client_secret': secret_key,
    'redirect_uri': r_url,
    'grant_type': 'authorization_code'}

response = rq.post(url, headers=headers, data=data)
jsr = response.json()
# print(response.json())

with open('accessToken.txt','w') as file:
    file.write(jsr['access_token'])
print(f"Access Token : {jsr['access_token']}")

print(jsr['access_token'])

