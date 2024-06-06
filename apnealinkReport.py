from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from docx import Document
import time
import sys
import os

def login(chrome):
    account=chrome.find_element(By.ID, "username")
    accountValue=account.get_attribute("value")
    password=chrome.find_element(By.ID, "password")
    passwordValue=password.get_attribute("value")
    if not accountValue:
        account.send_keys(sys.argv[1])
    if not passwordValue:
        password.send_keys(sys.argv[2])
    chrome.find_element(By.ID, "userSubmit").click()

if "crawler" in os.getcwd():
    pythonPath='.\\'
else:
    pythonPath='.\\crawler\\'
service=Service(executable_path=os.path.join(pythonPath, 'chromedriver.exe'))
options=webdriver.ChromeOptions()
chrome=webdriver.Chrome(service=service, options=options)
url="https://ap-airview.resmed.com"
wait=WebDriverWait(chrome, 10)
chrome.get(url)

def is_element_present(by, value):
    try:
        element=chrome.find_element(by, value)
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

select=Select(chrome.find_element(By.ID,"selectPageNum"))
reportValues=[]
dir=sys.argv[3].split('\\')[-1]
for option in select.options:
    if len(reportValues)>1:
        break
    select.select_by_value(str(option.text))
    patients=len(chrome.find_elements(By.XPATH, "//*[@id='hstPatientsTable']/tbody/tr"))
    time.sleep(1)
    for i in range(1, int(patients)+1):
        patientName=chrome.find_element(By.XPATH, "//*[@id='hstPatientsTable']/tbody/tr["+str(i)+"]/td[1]/a").text
        fileNames=patientName.split(', ')
        if len(fileNames[1].split('_'))<2:
            continue
        date=fileNames[1].split('_')[0]
        patientID=fileNames[1].split('_')[1]
        if date==dir.split('_')[0] and patientID==dir.split('_')[1] and fileNames[0]==dir.split('_')[2]:
            patient_href=chrome.find_element(By.XPATH, "//*[@id='hstPatientsTable']/tbody/tr["+str(i)+"]/td[1]/a").get_attribute("href")
            chrome.get(patient_href)
            reportValues=chrome.find_elements(By.CLASS_NAME, "column.report-value")
            break

def replace_text_in_table(table, target_dict):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for targetText, newText in target_dict.items():
                    if targetText in paragraph.text:
                        inline=paragraph.runs
                        inlineText=""
                        for i in range(len(inline)):
                            inlineText=inlineText+inline[i].text
                            inline[i].text=""
                        if targetText in inlineText:
                            text=inlineText.replace(targetText, newText)
                            inline[0].text=text

def main(content):
    if os.path.exists(sys.argv[3]+'\\'+dir+'.docx'):
        try:
            os.remove(sys.argv[3]+'\\'+dir+'.docx')
        except PermissionError:
            pass
        except Exception as e:
            pass
    doc=Document(os.path.join(pythonPath, '報告模板.docx'))
    # f=open(sys.argv[3]+'\\reportTemp.txt', 'r')
    # lines=f.readlines()
    
    replacementDict={

    }
    replace_text_in_table(doc.tables[0], replacementDict)
    replace_text_in_table(doc.tables[2], replacementDict)
    doc.save(sys.argv[3]+'\\'+dir+'.docx')
    # f.close()
    # os.remove(sys.argv[3]+'\\reportTemp.txt')

# if __name__ == "__main__":
#     main()