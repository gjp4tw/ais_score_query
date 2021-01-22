from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from urllib.parse import quote, urlencode
import requests, yaml, os
from bs4 import BeautifulSoup
def login(driver,config):
    driver.get('https://ais.ntou.edu.tw')
    print("[登入中]")
    WebDriverWait(driver,30).until(EC.presence_of_element_located((By.ID,"M_PORTAL_LOGIN_ACNT")))
    username=driver.find_element_by_id("M_PORTAL_LOGIN_ACNT")
    username.clear()
    username.send_keys(config['account'])
    pwd=driver.find_element_by_id("M_PW")
    pwd.clear()
    pwd.send_keys(config['password'])
    driver.find_element_by_id("LGOIN_BTN").click()
    driver.get('https://ais.ntou.edu.tw/title.aspx')
    WebDriverWait(driver,30).until(EC.presence_of_element_located((By.ID,"USERNAME")))
    global name
    name = driver.find_element_by_id("USERNAME").text
    print('[登入成功] username: ',name)
    return  driver.get_cookies()

if __name__ == '__main__':
    studentid = str(input("要查詢的學號:"))
    op=webdriver.ChromeOptions()
    # op.add_argument('headless')
    op.add_experimental_option("excludeSwitches", ["enable-logging"])
    data = urlencode({"Mode":"MOD", "STNO":studentid, "QUERY_TYPE": "3"})#歷年成績
    data1 = urlencode({"Mode":"MOD", "STNO":studentid, "QUERY_TYPE": "1", "QUERY_VALUE":"1091"})
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36 Edg/87.0.664.75',
        "Connection":"close"
    }
    try:
        with open('config.yaml','r',encoding='UTF-8') as f:
            config = yaml.load(f, Loader = yaml.FullLoader)
        driver=webdriver.Chrome(options=op)
        cookies = login(driver,config)
        c = requests.cookies.RequestsCookieJar()
        for i in cookies:
            c.set(i["name"], i['value'])
        session = requests.Session()
        session.cookies.update(c)
        req = session.post(url="https://ais.ntou.edu.tw/Application/GRD/GRD50/GRD5010_02.aspx", data = data, headers=headers)
        soup = BeautifulSoup(req.text, 'lxml')
        __viewstate = quote(soup.find(id = '__VIEWSTATE')['value'])
        __EVENTVALIDATION = quote(soup.find(id = '__EVENTVALIDATION')['value'])
        __VIEWSTATEGENERATOR = quote(soup.find(id = '__VIEWSTATEGENERATOR')['value'])
        __CRYSTALSTATECrystalReportViewer = quote(soup.find(id = '__CRYSTALSTATECrystalReportViewer')['value'])
        script  = f"ScriptManager1=AjaxPanel%7CReQuery&__EVENTTARGET=ReQuery&__EVENTARGUMENT=&__CRYSTALSTATECrystalReportViewer=&__VIEWSTATE={__viewstate}&__VIEWSTATEGENERATOR=&__VIEWSTATEENCRYPTED=&__EVENTVALIDATION={__EVENTVALIDATION}&ActivePageControl=&ColumnFilter=&QUERY_TYPE=3&STNO=&STUDY_STATUS=01&ASYS_CODE=0&DEGREE_CODE=0&COLLEGE_CODE=0160&FACULTY_CODE=0162&TEACH_GRP=0507&__ASYNCPOST=true&"
        script1 = f"ScriptManager1=AjaxPanel%7CReQuery&__EVENTTARGET=ReQuery&__EVENTARGUMENT=&__CRYSTALSTATECrystalReportViewer=&__VIEWSTATE={__viewstate}&__VIEWSTATEGENERATOR=&__VIEWSTATEENCRYPTED=&__EVENTVALIDATION={__EVENTVALIDATION}&ActivePageControl=&ColumnFilter=&QUERY_TYPE=1&QUERY_VALUE=1091&STNO=&STUDY_STATUS=01&ASYS_CODE=0&DEGREE_CODE=0&COLLEGE_CODE=0160&FACULTY_CODE=0162&TEACH_GRP=0507&IS_MOMENT=True&__ASYNCPOST=true&"
        req.close()
        req = session.post("https://ais.ntou.edu.tw/Application/GRD/GRD50/GRD5010_02.aspx",data = script, headers = headers)
        soup = BeautifulSoup(req.text, 'lxml')
        f = open("result.html","w",encoding="utf-8")
        f.write(soup.prettify())
        f.close()
        req.close()
        driver.get(os.getcwd()+'/result.html')
        os.remove("result.html")
    except Exception as e:
        print(e)
        driver.quit()
    # finally:
        # driver.quit()