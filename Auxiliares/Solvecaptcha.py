from re import T
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
import time
import random
from selenium.webdriver.common.keys import Keys
import time
import speech_recognition as sr
from selenium import webdriver
import ffmpy
import requests
import urllib
import pydub
import os
import Functions_Bot as FB


from selenium.webdriver.support.ui import WebDriverWait


def get_clear_browsing_button(driver):
    """Find the "CLEAR BROWSING BUTTON" on the Chrome settings page."""
    return driver.find_element_by_id('clearBrowsingDataConfirm')


def clear_cache(driver, timeout=60):
    """Clear the cookies and cache for the ChromeDriver instance."""
    # navigate to the settings page
    driver.get('chrome://settings/clearBrowserData')

    # wait for the button to appear
    wait = WebDriverWait(driver, timeout)
    #wait.until(get_clear_browsing_button)
    time.sleep(5)
    # click the button to clear the cache
    driver.find_element_by_id("clearBrowsingDataConfirm")
    #get_clear_browsing_button(driver).click()

    # wait for the button to be gone before returning
    wait.until_not(get_clear_browsing_button)



ChromeDriver = "C:\drivers_selenium\chromedriver_102.exe"

chr_options = Options()
    #El explorador queda abierto al finalizar la preuba
chr_options.add_experimental_option("useAutomationExtension", False)
chr_options.add_experimental_option("excludeSwitches",["enable-automation"])
chr_options.add_experimental_option("detach", True)                             
D_CHR = webdriver.Chrome(options=chr_options, executable_path=ChromeDriver)
D_CHR.maximize_window()

D_CHR.find_element_by_id

D_CHR.get(FB.URL_Elemplo)  # Registrarse en cuenta de google
user_empleo = D_CHR.find_element_by_name("EmailField")
user_empleo.send_keys("joseph.rojas@exsis.com.co")
time.sleep(2)
pass_empleo = D_CHR.find_element_by_name("PasswordField")
pass_empleo.send_keys("Recru_2022_Auto")
pass_empleo.send_keys(Keys.ENTER)

time.sleep(15)

D_CHR.get("https://www.elempleo.com/co/empresas/buscar?&keywords=ingeniero%20de%20sistemas")
reciente = "/html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[1]/div[2]/div/span[1]/span[1]/span"
exp_min = "html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[3]/div[1]/div/form/div[1]/div[6]/div/div/div/div/span[1]"
exp_max = "/html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[3]/div[1]/div/form/div[1]/div[6]/div/div/div/div/span[2]"

time.sleep(5)
ased=D_CHR.find_element_by_xpath(reciente)
ased.click()
for i in range(2):
    time.sleep(5)
    ased.send_keys(Keys.ARROW_DOWN)
ased.send_keys(Keys.ENTER)

time.sleep(15)
el1 = D_CHR.find_element_by_xpath(exp_min)
el2 = D_CHR.find_element_by_xpath(exp_max)
time.sleep(10)

for i in range(10):
    t = int(input())
    y = int(input())
    ActionChains(D_CHR).drag_and_drop_by_offset(el1,t,0).perform()
    time.sleep(10)
    ActionChains(D_CHR).drag_and_drop_by_offset(el2,-y,0).perform()
    time.sleep(5)
time.sleep(10)
print("OK")

D_CHR.switch_to.default_content()
frames = D_CHR.find_element_by_xpath("html/body/div[13]/div[2]/iframe")
#print(frames)
D_CHR.switch_to.frame(frames)
delay()
D_CHR.find_element_by_id("recaptcha-audio-button").click()
#D_CHR.find_element_by_xpath("/html/body/div/div/div[3]/div[2]/div/div/div[2]/button").click()
print("OK 2")
time.sleep(3)
try:
    D_CHR.find_element_by_xpath("/html/body/div/div/div[3]/div/button").click()
except:
    D_CHR.find_element_by_xpath("/html/body/div/div/div[3]/div/div/div//div[2]/button").click()

src = D_CHR.find_element_by_id("audio-source").get_attribute("src")
print("[INFO] Audio src: %s"%src)

urllib.request.urlretrieve(src,os.getcwd()+"\\sample.mp3")
time.sleep(2)
sound = pydub.AudioSegment.from_mp3(os.getcwd()+"\\sample.mp3")
time.sleep(2)
sound.export(os.getcwd()+"/sample.wav",format="wav")
time.sleep(6)
sample_audio = sr.AudioFile(os.getcwd()+"\\sample.wav")
time.sleep(2)
r = sr.Recognizer()
time.sleep(2)
with sample_audio as source:
    time.sleep(2)
    audio = r.record(source)
time.sleep(2)
karte = r.recognize_google(audio)
time.sleep(2)
print("[INFO] Recaptcha code:%s"%karte)
time.sleep(10)
tr = D_CHR.find_element_by_id("audio-response")
tr.send_keys(karte)
time.sleep(3)
tr.send_keys(Keys.ENTER)