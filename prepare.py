import re
import os

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import openpyxl

def openxlsx(xlsx_file_path):
        # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(xlsx_file_path)

    # 첫 번째 시트 선택
    worksheet = workbook.active

    # A열의 모든 값을 리스트로 저장
    column_a = []
    for cell in worksheet['A']:
        column_a.append(cell.value)
        
    return column_a
    

def loop1(book_title_box,xlsx_file_path , val):

    title = str(book_title_box.get_attribute('value'))
    column = openxlsx(xlsx_file_path)
    title = column[val] + title
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
        
def loop2(book_title_box,xlsx_file_path, val):

    title = str(book_title_box.get_attribute('value'))
    column = openxlsx(xlsx_file_path)
    title = title + column[val]
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
        
def loop3(book_title_box,input_text):
    title = str(book_title_box.get_attribute('value'))
    title = input_text + title
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
    
def loop4(book_title_box,input_text):
    title = str(book_title_box.get_attribute('value'))
    title = title + input_text
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
    
def loop5(book_title_box):
    title = str(book_title_box.get_attribute('value'))
    pattern = re.compile(r'\[(\d+)\]')
    matches = pattern.findall(title)
    title = pattern.sub('',title)
    title = title.strip()

    if matches:
        book_title_box.clear()
        paste = "[{}] {}".format(matches[0],title)
        book_title_box.send_keys(paste)
        book_title_box.send_keys(Keys.ENTER)

def loop6(book_title_box):
    title = str(book_title_box.get_attribute('value'))
    pattern = re.compile(r"\[[^\d]+\]")
    matches = pattern.findall(title)
    
    if matches:
        book_title_box.clear()
        title = re.sub(r"\[[^\d]+\]", "", title)
        title = title.strip()
        book_title_box.send_keys(title)
        book_title_box.send_keys(Keys.ENTER)
        
def loop7(book_title_box):
    title = str(book_title_box.get_attribute('value'))
    pattern = r'\[[^\]]*\]'
    title = re.sub(pattern, '', title)
    title = title.strip()
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)


def common(ins):
    
    try :
        # 저장 버튼
        ins.driver.execute_script("arguments[0].click()", ins.save_btn)
    except :
        ins.error.emit("저장 버튼 에러")

        # 저장 확인 알림
    try:
        WebDriverWait(ins.driver, 3).until(EC.alert_is_present())
        alert = ins.driver.switch_to.alert
        alert.accept()
    except:
        ins.error.emit("알림 확인 버튼 에러")

        # 중복된 등록번호 알림
    try:
        wait = WebDriverWait(ins.driver, 3)
        dup_msg = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]')))
        if dup_msg.is_displayed() == True:
            dup_btn = ins.driver.find_element(By.XPATH, '//*[@id="confirm_ok"]')
            ins.driver.execute_script("arguments[0].click()", dup_btn)
    except:
        pass
        
        # 신텍스 오류
    try:
        WebDriverWait(ins.driver, 3).until(EC.alert_is_present())
        syn_alert = ins.driver.switch_to.alert.get_attribute("style")
        if syn_alert == 'block' :
            ins.next_btn.send_keys(Keys.ENTER)
            if ins.val == ins.num :
                ins.flag == False
            else:
                ins.val -= 1
            # 오류난 파일, 바탕화면의 "서지정보 오류 목록.txt" 파일에 따로 저장할것
                # desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
                # file_path = os.path.join(desktop, '신텍스 오류 목록.txt')
                # with open(file_path, mode='w', encoding='utf-8') as f:
                #     f.write('\n')
    except:
        ins.error.emit("신텍스 오류 에러")

    try :
        admin_error = ins.driver.find_element(By.XPATH, '/html/body/div[6]')
        if admin_error.get_attribute("style") == 'block':
            ins.next_btn.send_keys(Keys.ENTER)
            if ins.val == ins.num :
                ins.flag == False
            else:
                ins.val -= 1
    except :
        ins.error.emit("관리자 문의 에러")
    
    
def klas_upload(ins):
    ins.driver.get("https://klas.jeonju.go.kr/klas3/Admin/")
    
    id_box = ins.driver.find_element(By.ID, 'manager_id')
    id_box.send_keys(ins.id)
    
    pw_box = ins.driver.find_element(By.ID, 'password')
    pw_box.send_keys(ins.pw)
    
    ins.driver.find_element(By.CLASS_NAME, 'btn_login').click()

    # Books Page
    try:
        ins.driver.get("https://klas.jeonju.go.kr/klas3/Books/booksPage/")
        ins.driver.execute_script("TitleViewChange('file');")
        # 파일 업로드
        ins.driver.find_element(By.CSS_SELECTOR, "input[type='file']").send_keys(ins.txt_file_path)
        ins.driver.execute_script("speciesFileSearch(1, true);")
        
        # 진행 사이즈 선택
        
        
        # 자료량 기입
        data_num = ins.driver.find_element(By.XPATH, '//*[@id="content"]/div[5]/span[2]')
        ins.num = re.sub(r'[^0-9]', '', data_num.text)
        ins.num = int(ins.num)
        
        # 체크박스 선택
        box = ins.driver.find_element(By.XPATH, '//*[@id="all_solr_check"]')
        ins.driver.execute_script("arguments[0].click()", box)
    except :
        ins.driver.quit()

    try:
        marc_editor = ins.driver.find_element(By.XPATH, '//*[@id="content"]/div[8]/div/input[3]')
        ins.driver.execute_script("arguments[0].click()", marc_editor)
        ins.driver.switch_to.window(ins.driver.window_handles[1])
    except :
        ins.driver.quit()

    try :
        wait = WebDriverWait(ins.driver, 10)
        ins.book_title_box = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="book_title"]')))    
        ins.save_btn = ins.driver.find_element(By.XPATH, '//*[@id="marcForm"]/div[2]/div[1]/input[18]')
        ins.next_btn = ins.driver.find_element(By.XPATH, '//*[@id="nextBtn"]')
        ins.syn_tex_box = ins.driver.find_element(By.XPATH, '//*[@id="syntexMsg_div"]')
    
    except:
        pass