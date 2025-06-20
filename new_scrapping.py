import os
import time
import requests
import pandas as pd
from flask import Flask, render_template, request
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoAlertPresentException, InvalidSessionIdException
from webdriver_manager.chrome import ChromeDriverManager

def writeCSV(enroll, name, *args, sgpa, cgpa, remark, filename):
    gradesString = [str(a) + "," for a in args]
    information = [enroll, ",", name, ","] + gradesString + [sgpa, ",", cgpa, ",", remark, "\n"]
    with open(filename, 'a') as f:
        f.writelines(information)
        f.close()

def makeXslx(filename):
    csvFile = f'{filename}.csv'
    df = pd.read_csv(csvFile)
    excelFile = f'{filename}.xlsx'
    df.index += 1
    df.to_excel(excelFile)

def readFromImage(url: str) -> str:
    api_key = os.environ.get('OCR_API_KEY', 'K86969399988957')  # Use environment variable
    
    image_response = requests.get(url)
    if image_response.status_code != 200:
        print("Failed to fetch image from URL.")
        return ""

    ocr_response = requests.post(
        'https://api.ocr.space/parse/image',
        files={'filename': ('captcha.jpg', image_response.content)},
        data={'apikey': api_key, 'OCREngine': 2}
    )

    try:
        result = ocr_response.json()
        if result.get('IsErroredOnProcessing'):
            return ""
        parsed = result.get('ParsedResults')
        if not parsed:
            print("No ParsedResults found in response.")
            return ""
        text = parsed[0].get('ParsedText', "")
        return text.upper().replace(" ", "").strip()
    except Exception as e:
        return ""

def get_chrome_driver():
    """Setup Chrome driver for Railway deployment"""
    chrome_options = Options()
    
    # Railway-specific Chrome options
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--mute-audio")
    chrome_options.add_argument("--disable-background-timer-throttling")
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")
    chrome_options.add_argument("--disable-renderer-backgrounding")
    chrome_options.add_argument("--disable-features=TranslateUI")
    chrome_options.add_argument("--disable-ipc-flooding-protection")
    
    # Try to use system Chrome on Railway
    if os.path.exists("/usr/bin/chromium"):
        chrome_options.binary_location = "/usr/bin/chromium"
        service = Service("/usr/bin/chromedriver")
    else:
        # Fallback to ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"Error creating Chrome driver: {e}")
        # Try with ChromeDriverManager as fallback
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver

def resultFound(start: int, end: int, branch: str, year: str, sem: int):
    if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
        print("Wrong Branch Entered")
        return
    
    noResult = []
    driver = get_chrome_driver()
    firstRow = True
    filename = f'{branch}_sem{sem}_result.csv'
    driver.implicitly_wait(0.5)
    
    try:
        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
        driver.find_element(By.ID, "radlstProgram_1").click()

        while start <= end:
            if start < 10:
                num = "00" + str(start)
            elif start < 100:
                num = "0" + str(start)
            else:
                num = str(start)
            
            enroll = f"0105{branch}{year}1{num}"
            print(f"Currently compiling ==> {enroll}")

            img_element = driver.find_element(By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')
            img_src = img_element.get_attribute("src")
            url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'
            print(url)

            captcha = readFromImage(url)
            captcha = captcha.replace(" ", "")

            Select(driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]')).select_by_value(str(sem))
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').send_keys(captcha)
            time.sleep(1)
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').send_keys(enroll)
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]').send_keys(Keys.ENTER)

            time.sleep(2)
            alert = Alert(driver)
            try:
                alerttext = alert.text
                alert.accept()
            except NoAlertPresentException:
                pass
            except InvalidSessionIdException:
                pass

            if "Total Credit" in driver.page_source:
                if firstRow == True:
                    details = []
                    rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                    for row in rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 4 and '[T]' in cells[0].text:
                            details.append(cells[0].text.strip('- [T]'))
                    firstRow = False
                    writeCSV("Enrollment No.", "Name", *details, sgpa="SGPA", cgpa="CGPA", remark="REMARK", filename=filename)
                
                roll_nu = driver.find_element("id", 'ctl00_ContentPlaceHolder1_lblRollNoGrading').text
                name = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblNameGrading").text
                grades = []
                rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4 and '[T]' in cells[0].text:
                        grades.append(cells[3].text.strip())
                sgpa = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblSGPA").text
                cgpa = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblcgpa").text
                result = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblResultNewGrading").text

                result = result.replace(",", " ")
                name = name.replace("\n", " ")
                writeCSV(enroll, name, *grades, sgpa=sgpa, cgpa=cgpa, remark=result, filename=filename)
                print("Compilation Successful")

                driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
                start = start + 1
            else:
                if "Result" in alerttext:
                    driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
                    start = start + 1
                    noResult.append(enroll)
                    print(f"Enrollment NO: {enroll} not found.")
                else:
                    driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').clear()
                    driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceFolder1_txtrollno"]').clear()
                    print("Wrong Captcha Entered")
                    continue

        print(f'Enrollment Numbers not found ====> {noResult}')
    
    except Exception as e:
        print(f"Error during scraping: {e}")
    finally:
        driver.quit()

# Flask App
app = Flask(__name__)

@app.route('/')
def form():
    return render_template('1st.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        branch = request.form['branch']
        year = request.form['year']
        sem = int(request.form['sem'])
        start = int(request.form['start'])
        end = int(request.form['end'])

        resultFound(start, end, branch.upper(), year, sem)

        return f"""
        <h2>✅ Result fetching completed for {branch.upper()} Sem {sem}</h2>
        <p>Check the server logs for details.</p>
        <a href="/">Go back</a>
        """
    except Exception as e:
        return f"""
        <h2>❌ Error occurred</h2>
        <p>Error: {str(e)}</p>
        <a href="/">Go back</a>
        """

@app.route('/health')
def health():
    return {'status': 'healthy'}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
