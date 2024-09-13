from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
from random import randint
from datetime import datetime
from time import perf_counter
from tqdm import tqdm
from openpyxl import Workbook

t0 = perf_counter()
current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
processed_data = []
date = "mmddyyyy"   #mm/dd/yyyy
process_header = ['TOWNSHIP', 'SECTION', 'Others', 'TYPE', 'FILED_DATE', 'GRANTOR', 'GRANTEE', 'Instrument', 'Doc#']
base_url = "https:"

grantor = ''
grantee = ''
sec = ''
tship = ''
rng = ''
other = ''
rang = ''

urls = []
err_urls_header = ['Error']
err_urls = []
total_url = ['Total Url #']
curr_url_header = ['Current URL']
curr_url = []

# openpyxl implementation code
file = f'raw_data_{str(current_datetime)}.xlsx'
# Load existing workbook
wb = Workbook()
# Select the active sheet
ws = wb.active
ws.append(process_header)

def convert_to_excel(data, column_header, filename, mode):
    df = pd.DataFrame(data, columns =column_header)
    if mode == 'Y':
        df.drop_duplicates(inplace=True)
    df.to_excel(f"{filename}_{str(current_datetime)}.xlsx", index=False)
# Define a custom user agent
my_user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
options = webdriver.EdgeOptions()
options.add_experimental_option("excludeSwitches", ['enable-automation'])
options.add_argument(f"--user-agent={my_user_agent}")
options.add_argument('--inprivate')
options.add_argument("disable-infobars")
options.add_argument('--allow-running-insecure-content')
options.add_argument('--ignore-certificate-errors')

driver = webdriver.Edge(options=options)

driver.get("http://")

# click on the Clerk
for i in range(10):
    try:
        clerk = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Clerk')]"))
            )
        driver.execute_script("arguments[0].click();", clerk)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

# click on the Grantee
for i in range(10):
    try:
        grantor = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Grantor')]"))
            )
        driver.execute_script("arguments[0].click();", grantor)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

sleep(2)
# Enter the data and press enter to search dynamic content
for i in range(10):
    try:
        date_field = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.NAME, "filedte"))
            ).send_keys(date, Keys.ENTER)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

while True:
    try:
        curr_url.append(driver.current_url)
        page_content = BeautifulSoup(driver.page_source, 'html.parser')

        table = page_content.find('table', {"id": "tableResults"})
        body = table.find('tbody')
        trb = body.find_all('tr')
        for tr in trb:
            for a in tr.find_all('a', href=True):
                link = f"{base_url}{a['href']}"
                if link not in urls:
                    urls.append(link)
        # click on the Next
        for i in range(10):
            try:
                next_button = WebDriverWait(driver, 300).until(
                        EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Next')]"))
                    )
                driver.execute_script("arguments[0].click();", next_button)
                sleep(randint(1, 3))
                break
            except NoSuchElementException as e:
                sleep(randint(1, 3))
                driver.refresh()
        else:
            raise e

    except:
        break

# convert_to_excel(curr_url, curr_url_header, 'current_url_', "Y")
# print(f"`current_url_{str(current_datetime)}` excel file created successfully.")

# convert_to_excel(urls, total_url, 'total_url_', "Y")
# print(f"`total_url_{str(current_datetime)}` excel file created successfully.")  
"""
with open("done.txt", "w") as done:
"""
for url in tqdm(urls):
    # done.write(url+'\n')
    sleep(randint(3, 7))
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    try:
        reception_no = soup.find("b", string="Reception #:").next_sibling.next_sibling.get_text().strip()
        description = soup.find("b", string="Kind of Instrument:").next_sibling.next_sibling.get_text().strip()
        f_date= soup.find("b", string="Date Filed:").next_sibling.next_sibling.get_text().strip()
        filed_date = datetime.strptime(f_date, '%Y%m%d').strftime('%m/%d/%Y')
        i_date = soup.find("b", string="Intrument Date:").next_sibling.next_sibling.get_text().strip()
        instrument = datetime.strptime(i_date, '%Y%m%d').strftime('%m/%d/%Y')

        grtee= soup.find_all("fieldset")[1].find_all("b", string="Grantee Name:")
        if len(grtee) > 1:
            grantee = f"{grtee[0].next_sibling.get_text().strip()} ET AL"
        else:
            grantee = grtee[0].next_sibling.get_text().strip()

        grtor = soup.find_all("fieldset")[2].find_all("b", string="Grantor Name:")
        if len(grtor) > 1:
            grantor = f"{grtor[0].next_sibling.get_text().strip()} ET AL"
        else:
            grantor = grtor[0].next_sibling.get_text().strip()

        try:
            sections = soup.find_all("b", string="Section:")
            townships = soup.find_all("b", string="Township:")
            ranges = soup.find_all("b", string="Range:")
            other_descs = soup.find_all("b", string="Description:")
            for (section, township, rnge, other_desc) in zip(sections, townships, ranges, other_descs):
                sec = section.next_sibling.next_sibling.get_text().strip()
                tship = township.next_sibling.next_sibling.get_text().strip()
                rng = rnge.next_sibling.next_sibling.get_text().strip()
                other = other_desc.next_sibling.next_sibling.get_text().strip()
                rang = f"{tship};{rng}"
                processed_data.append([rang, sec, other, description, filed_date, grantor, grantee, instrument, reception_no])
                ws.append([rang, sec, other, description, filed_date, grantor, grantee, instrument, reception_no])
            wb.save(file)
        except Exception as e:
            # print(e)
            pass
    except:
        err_urls.append(url)
        # print("Error: ", url)
        pass


convert_to_excel(processed_data, process_header, 'processed_leacounty', "Y")
print(f"Raw data based `processed_leacounty_{str(current_datetime)}` excel file created successfully.")

# convert_to_excel(err_urls, err_urls_header, 'error_urls_leacounty', "Y")
# print(f"Error `error_urls_leacounty_{str(current_datetime)}` excel file created successfully.")

sleep(10) 
driver.close()
