import os
import time
from selenium import webdriver
from pptx import Presentation
from pptx.util import Cm
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options as ChromeOptions
import win32com.client
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


options = ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
driver.get("https://freemeteo.gr/kairos/plati/7-imeres/pinakas/?gid=734573&language=greek&country=greece")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,"//*[contains(text(), 'Συναίνεση')]"))).click()

weather = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "today.table")))
tooltip_element = driver.find_element(By.CLASS_NAME, "prev.sevendays")  
actions = ActionChains(driver)
actions.move_to_element(tooltip_element).perform()

weather.screenshot("weather_screenshot.png")
driver.quit()

prs = Presentation("template.pptx")

slide = prs.slides[0]

left = Cm(4.18)  
top = Cm(1.29)   
width = Cm(17.03) 
height = Cm(11.69) 

slide.shapes.add_picture("weather_screenshot.png", left, top, width, height)

prs.save("updated_presentation.pptx")

powerpoint = win32com.client.Dispatch("Powerpoint.Application")
try:
    presentation = powerpoint.Presentations.Open(FileName=r'C:\Users\CAMELS\Documents\GitHub\powerpoint_automator\updated_presentation.pptx')
except Exception as e:
    print('File cannot be found')
    exit

presentation.CreateVideo(r'C:\Users\CAMELS\Documents\GitHub\powerpoint_automator\out.mp4',VertResolution=1080,Quality=100)
while presentation.CreateVideoStatus == 1:
    time.sleep(1)
presentation.Close()

powerpoint.Quit()