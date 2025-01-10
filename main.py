import time
import logging
from pathlib import Path
from contextlib import contextmanager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from pptx import Presentation
from pptx.util import Cm
import win32com.client

# Configuration
CONFIG = {
    'url': "https://freemeteo.gr/kairos/plati/7-imeres/pinakas/?gid=734573&language=greek&country=greece",
    'screenshot_path': "weather_screenshot.png",
    'template_path': "template.pptx",
    'output_pptx': "updated_presentation.pptx",
    'output_video': "out.mp4",
    'slide_dimensions': {'left': 4.18, 'top': 1.29, 'width': 17.03, 'height': 11.69}
}

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

@contextmanager
def chrome_driver():
    """Context manager for Chrome driver"""
    options = ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    try:
        yield driver
    finally:
        driver.quit()

def capture_weather_screenshot():
    """Capture weather data screenshot using Selenium"""
    with chrome_driver() as driver:
        logging.info("Accessing weather website")
        driver.get(CONFIG['url'])
        
        # Accept cookies
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Συναίνεση')]"))
        ).click()
        
        # Wait for and find weather element
        weather = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "today.table"))
        )
        
        # Move to tooltip element
        tooltip_element = driver.find_element(By.CLASS_NAME, "prev.sevendays")
        ActionChains(driver).move_to_element(tooltip_element).perform()
        
        # Take screenshot
        weather.screenshot(CONFIG['screenshot_path'])
        logging.info("Weather screenshot captured successfully")

def update_powerpoint():
    """Update PowerPoint presentation with screenshot"""
    try:
        prs = Presentation(CONFIG['template_path'])
        slide = prs.slides[0]
        dims = CONFIG['slide_dimensions']
        
        slide.shapes.add_picture(
            CONFIG['screenshot_path'],
            Cm(dims['left']),
            Cm(dims['top']),
            Cm(dims['width']),
            Cm(dims['height'])
        )
        
        prs.save(CONFIG['output_pptx'])
        logging.info("PowerPoint presentation updated successfully")
    except Exception as e:
        logging.error(f"Error updating PowerPoint: {str(e)}")
        raise

def create_video():
    """Convert PowerPoint to video"""
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    try:
        presentation = powerpoint.Presentations.Open(
            FileName=str(Path.cwd() / CONFIG['output_pptx'])
        )
        
        presentation.CreateVideo(
            str(Path.cwd() / CONFIG['output_video']),
            VertResolution=1080,
            Quality=100
        )
        
        while presentation.CreateVideoStatus == 1:
            time.sleep(1)
            
        presentation.Close()
        logging.info("Video created successfully")
    except Exception as e:
        logging.error(f"Error creating video: {str(e)}")
        raise
    finally:
        powerpoint.Quit()

def main():
    """Main execution function"""
    try:
        capture_weather_screenshot()
        update_powerpoint()
        create_video()
        logging.info("Process completed successfully")
    except Exception as e:
        logging.error(f"Process failed: {str(e)}")
        return 1
    return 0

if __name__ == "__main__":
    exit(main())