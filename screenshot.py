from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import requests
from time import sleep
import loginFormDR_up
from selenium.webdriver.common.by import By
from PIL import Image
from io import BytesIO
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

screenshot_path = ".\screenshot.png"

def login():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
   
    driver.get("https://app.formdr.com/login")
    # Capture the full page height by scrolling
    # total_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")

    # Set the window size to the full page height
    driver.set_window_size(1920, 1080)
    sleep(6)
    # Find the username and password input fields and enter your credentials
    username_input = driver.find_element("name", "email")
    password_input = driver.find_element("name", "password")

    username_input.send_keys("admin@oregonfirstresponderevals.com")
    password_input.send_keys("Packets66!")

    # Submit the login form
    password_input.send_keys(Keys.RETURN)

    # Wait for the page to load after login (adjust the sleep time or use WebDriverWait)
    sleep(7)

    return driver

def find_candidate_item(driver, candidate_name):
    driver.get("https://app.formdr.com/submissions")
    sleep(5)
    elements = driver.find_elements(By.CSS_SELECTOR, ".MuiInputBase-input.MuiInput-input.MuiInputBase-inputAdornedEnd")
    if len(elements) == 2:
        search_input = elements[1]
        search_input.send_keys(candidate_name)
    else:
        search_input = None
        return None
    

    elements = driver.find_elements(By.CSS_SELECTOR, ".MuiInputAdornment-root.MuiInputAdornment-positionEnd")
    if len(elements) == 2:
        start_search_icon = elements[1]
        start_search_icon.click()
    else:
        print("There is no search icon")
        return None
    sleep(5)

    elements = driver.find_elements(By.CSS_SELECTOR, ".name-container.formdr-flex-grow")
    if len(elements) > 0:
        candidate_item = elements[0]
    else:
        print("there is no candidate with such name")
        return None

    candidate_item.click()
    sleep(4)

    tabs = driver.find_elements(By.CSS_SELECTOR, ".MuiButtonBase-root.MuiTab-root.MuiTab-textColorInherit.sc-iUuytg.esdhrP")
    if len(tabs) < 6:
        print("There are no 6 tabs")
        return None
    background_tab = None
    for tab in tabs:
        if "THE PERSONAL BACKGROUND HISTORY FORM" in tab.accessible_name:
            background_tab = tab
    if background_tab != None:
        background_tab.click()
    else:
        return None
    
    sleep(1)

    return driver

def screenshot_signature(driver, candidate_name):

    driver.get("https://app.formdr.com/submissions")
    sleep(5)
    elements = driver.find_elements(By.CSS_SELECTOR, ".MuiInputBase-input.MuiInput-input.MuiInputBase-inputAdornedEnd")
    if len(elements) == 2:
        search_input = elements[1]
    else:
        search_input = None
    search_input.send_keys(candidate_name)

    elements = driver.find_elements(By.CSS_SELECTOR, ".MuiInputAdornment-root.MuiInputAdornment-positionEnd")
    if len(elements) == 2:
        start_search_icon = elements[1]
        start_search_icon.click()
    else:
        print("There is no search icon")
        return None
    sleep(5)

    elements = driver.find_elements(By.CSS_SELECTOR, ".name-container.formdr-flex-grow")
    if len(elements) > 0:
        candidate_item = elements[0]
    else:
        print("there is no candidate with such name")
        return None
    
    candidate_item.click()
    sleep(4)

    tabs = driver.find_elements(By.CSS_SELECTOR, ".MuiButtonBase-root.MuiTab-root.MuiTab-textColorInherit.sc-iUuytg.esdhrP")
    if len(tabs) < 6:
        print("There are no 6 tabs")
        return None
    background_tab = None
    for tab in tabs:
        if "THE PERSONAL BACKGROUND HISTORY FORM" in tab.accessible_name:
            background_tab = tab
    if background_tab != None:
        background_tab.click()
    else:
        return None
    sleep(1)
    elements = driver.find_elements(By.CSS_SELECTOR, ".fd-field-item.field-type-group")

    ROI_form = None
    element_title = None
    # element_title = driver.find_element(By.XPATH("//h2[contains(text(), 'AUTHORIZATION FOR RELEASE OF HEALTH INFORMATION')]"))
    for container in elements:
        sleep(1)
        try:
            container_title_element = container.find_element(By.TAG_NAME, "h2")
            container_title = container_title_element.text
            if container_title == "AUTHORIZATION FOR RELEASE OF HEALTH INFORMATION":
                ROI_form = container
                element_title = container_title_element
                break
            else:
                continue
        except:
            continue
    if ROI_form == None:
        child_elements = driver.find_elements(By.CSS_SELECTOR, 'h2.fd-field-item-title[style="font-size: 30px; text-align: left;"]')
        for element in child_elements:
            if element.text == "AUTHORIZATION FOR RELEASE OF HEALTH INFORMATION":
                element_title = element
    if element_title:
        location = element_title.location
        if ROI_form:
            top = location['y']

            driver.execute_script("window.scrollTo(0, arguments[0]);", top-65)  # Replace 345 with the Y-coordinate of the element
            sleep(2)
            screenshot = driver.get_screenshot_as_png()
            screenshot = Image.open(BytesIO(screenshot))
            image_byte_array = BytesIO()
            screenshot.save(image_byte_array, format='PNG')

            # Upload the file to Google Drive
            media_body = BytesIO(image_byte_array.getvalue())

            return media_body
        return None
    else:
        driver.quit()
        return None

# Close the browser

def get_va_download_url(driver, candidate_name):
    driver = find_candidate_item(driver, candidate_name)

    sleep(5)

    a_elements = driver.find_elements(By.TAG_NAME, "a")

    download_element = None

    for element in a_elements:
        if element.text == "Download File":
            download_element = element
    if download_element:
        return download_element.get_attribute('href')
    
    return download_element