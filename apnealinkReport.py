from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from docx import Document
import time
import sys

sign_in_account=sys.argv[1]
sign_in_password=sys.argv[2]
which_page=sys.argv[3]
which_patient=sys.argv[4]

def login(chrome):
    account = chrome.find_element(By.ID, "username")
    account_value = account.get_attribute("value")
    print(account_value)
    password = chrome.find_element(By.ID, "password")
    password_value = password.get_attribute("value")
    print(password_value)
    if not account_value:
        account.send_keys(sign_in_account)
    if not password_value:
        password.send_keys(sign_in_password)
    chrome.find_element(By.ID, "userSubmit").click()

service = Service(executable_path='./chromedriver.exe')
options = webdriver.ChromeOptions()
chrome = webdriver.Chrome(service=service, options=options)
url = "https://ap-airview.resmed.com"
wait = WebDriverWait(chrome, 10)
chrome.get(url)

def is_element_present(by, value):
    try:
        element = chrome.find_element(by, value)
    except NoSuchElementException as e:
        return False
    return True

if(is_element_present(By.ID,"onetrust-accept-btn-handler")==True):
    chrome.find_element(By.ID, "onetrust-accept-btn-handler").click()
    
login(chrome)
time.sleep(10)

diagnostic_href=chrome.find_element(By.XPATH, "//*[@id='diagnostic-patients-link']").get_attribute("href")
chrome.get(diagnostic_href)
time.sleep(1)

select = Select(chrome.find_element(By.ID,"selectPageNum"))
select.select_by_value(str(which_page))
time.sleep(1)

patientName=chrome.find_element(By.XPATH, "//*[@id='hstPatientsTable']/tbody/tr["+str(which_patient)+"]/td[1]/a").text
fileNames = patientName.split(', ')
fileName=fileNames[0]+fileNames[1]
# print(fileName)
patient_href=chrome.find_element(By.XPATH, "//*[@id='hstPatientsTable']/tbody/tr["+str(which_patient)+"]/td[1]/a").get_attribute("href")
chrome.get(patient_href)
report_values=chrome.find_elements(By.CLASS_NAME, "column.report-value")

def replace_text_in_table(table, target_dict):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for target_text, new_text in target_dict.items():
                    if target_text in paragraph.text:
                        inline = paragraph.runs
                        inlineText=""
                        for i in range(len(inline)):
                            inlineText=inlineText+inline[i].text
                            inline[i].text=""
                        if target_text in inlineText:
                            text = inlineText.replace(target_text, new_text)
                            inline[0].text = text

def main():
    doc = Document('成大Home ApneaLink screen報告v1.docx')
    replacement_dict = {
        '{a1}': report_values[13].text,
        '{a2}': report_values[12].text,
        '{a3}': report_values[11].text,
        '{a4}': report_values[30].text,
        '{a5}': report_values[29].text,
        '{a6}': report_values[34].text,
        '{a7}': report_values[33].text,
        '{a8}': report_values[32].text,
        '{a9}': report_values[31].text,
        '{a10}': report_values[16].text,
        '{a11}': report_values[15].text,
        '{a12}': report_values[14].text,
        '{a13}': report_values[18].text,
        '{a14}': report_values[17].text[1:-2],
        '{a15}': report_values[21].text,
        '{a16}': report_values[20].text,
        '{a17}': report_values[19].text,
        '{a18}': report_values[23].text,
        '{a19}': report_values[22].text[1:-2],
        '{a20}': report_values[26].text,
        '{a21}': report_values[25].text,
        '{a22}': report_values[24].text,
        '{a23}': report_values[28].text,
        '{a24}': report_values[27].text[1:-2],
        '{a25}': report_values[36].text,
        '{a26}': report_values[35].text,

        '{b1}': report_values[3].text,
        '{b2}': report_values[2].text,
        '{b3}': report_values[1].text,
        '{b4}': report_values[7].text,
        '{b5}': report_values[6].text,
        '{b6}': report_values[5].text,
        '{b7}': report_values[10].text,
        '{b8}': report_values[9].text,
        '{b9}': report_values[8].text,

        '{c1}': report_values[41].text,
        '{c2}': report_values[40].text,
        '{c3}': report_values[39].text,
        '{c4}': report_values[44].text,
        '{c5}': report_values[43].text,
        '{c6}': report_values[42].text,
        '{c7}': report_values[46].text,
        '{c8}': report_values[45].text,
        '{c9}': report_values[37].text,
        '{c10}': report_values[38].text,

        '{d1}': report_values[50].text,
        '{d2}': report_values[49].text,
        '{d3}': report_values[48].text,

        '{e1}': report_values[53].text,
        '{e2}': report_values[52].text,
        '{e3}': report_values[51].text,
    }
    replace_text_in_table(doc.tables[0], replacement_dict)
    replace_text_in_table(doc.tables[2], replacement_dict)
    doc.save('成大Home ApneaLink screen報告v1 - '+fileName+'.docx')

if __name__ == "__main__":
    main()