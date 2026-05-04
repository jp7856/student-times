import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

MAIN_URL   = "https://www.netimes.co.kr/index.asp"
SAMPLE_URL = "https://www.netimes.co.kr/pages/Kids/reading.asp?seq=36588"

options = webdriver.ChromeOptions()
options.add_argument("--no-sandbox")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get(MAIN_URL)
print("👉 로그인해주세요.")
input("✅ 로그인 후 Enter...")

driver.get(SAMPLE_URL)
time.sleep(3)

print("\n=== 기사 내용 ===")
print(driver.find_element(By.TAG_NAME, "body").text[:3000])
driver.quit()