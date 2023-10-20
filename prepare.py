import re
import os

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import openpyxl

def alert_ok(self):
    try:
        WebDriverWait(self.driver, 3).until(EC.alert_is_present())
        alert = self.driver.switch_to.alert
        alert.accept()
    except:
        pass

def column_A(xlsx_file_path):
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
    pattern = r'\[.*\]'

    if re.search(pattern, title):
        return

    column = column_A(xlsx_file_path)
    title = column[val] + title
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
        
def loop2(book_title_box,xlsx_file_path, val):

    title = str(book_title_box.get_attribute('value'))
    pattern = r'\[.*\]'

    if re.search(pattern, title):
        return
    column = column_A(xlsx_file_path)
    title = title + column[val]
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
        
def loop3(book_title_box,tab2_input):
    title = str(book_title_box.get_attribute('value'))
    pattern = r'\[.*\]'

    if re.search(pattern, title):
        return

    title = tab2_input + title
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
    
def loop4(book_title_box,tab2_input):
    title = str(book_title_box.get_attribute('value'))
    pattern = r'\[.*\]'

    if re.search(pattern, title):
        return
    title = title + tab2_input
    book_title_box.clear()
    book_title_box.send_keys(title)
    book_title_box.send_keys(Keys.ENTER)
    
def loop5(book_title_box):
    title = str(book_title_box.get_attribute('value'))
    pattern = re.compile(r'\[\d+\]')
    matches = pattern.findall(title)

    if matches:
        book_title_box.clear()
        title = re.sub(pattern, '',title)
        title = title.strip()
        book_title_box.send_keys(title)
        book_title_box.send_keys(Keys.ENTER)

def loop6(book_title_box):
    title = str(book_title_box.get_attribute('value'))
    pattern = re.compile(r"\[^\d]+\]")
    matches = pattern.findall(title)

    if matches:
        book_title_box.clear()
        title = re.sub(pattern, "", title)
        title = title.strip()
        book_title_box.send_keys(title)
        book_title_box.send_keys(Keys.ENTER)

        
def loop7(book_title_box):
    title = str(book_title_box.get_attribute('value'))
    pattern = re.compile(r'\[[^\]]*\]')
    matches = pattern.findall(title)

    if matches:
        book_title_box.clear()
        title = re.sub(pattern, "", title)
        title = title.strip()
        book_title_box.send_keys(title)
        book_title_box.send_keys(Keys.ENTER)


def common(self):

    WebDriverWait(self.driver, 10)

    ddc = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ddc_class"]')))
    ddc_text = str(ddc.get_attribute('value'))
    if not ddc_text.isnumeric() and len(ddc_text) > 0:
        ddc.clear()
        ddc.send_keys(Keys.ENTER)

    self.driver.execute_script("arguments[0].click()", self.save_btn)
    alert_ok(self)

    dup_msg = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]')))
        # 중복된 등록번호 알림
    if dup_msg.is_displayed() == True:
        dup_btn = self.driver.find_element(By.XPATH, '//*[@id="confirm_ok"]')
        self.driver.execute_script("arguments[0].click()", dup_btn)

    try:
        # markup Systex error
        is_error = self.syn_tex_box.find_element(By.TAG_NAME, 'div').text
        if len(is_error) > 0:
                    
            divs = self.driver.find_elements(By.XPATH, '//*[@id="marcEditor"]/div')
            div = divs[-2]
            fonts = div.find_elements(By.XPATH, './font')
            Reg_error = fonts[2].find_element(By.XPATH, './pre').text

            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            file_path = os.path.join(desktop, '신텍스 오류 목록.txt')
            with open(file_path, mode='a', encoding='utf-8') as f:
                f.write(f'{Reg_error}\n')
            self.driver.execute_script("arguments[0].click()", self.next_btn)
            return
    except:
        pass

    # try :
    #     # Other Failure
    #     failure_btn = self.driver.find_element(By.XPATH, '//*[@id="failAlertbtn"]')
    #     self.driver.execute_script("arguments[0].click()", failure_btn)
    #     self.mark_editor.send_keys(Keys.ENTER)
    #     self.driver.execute_script("arguments[0].click()", self.save_btn)
    #     alert_ok(self)
    # except :
    #     pass

def klas_upload(self):
    try :
        self.driver.get("https://klas.jeonju.go.kr/klas3/Admin/")
        
        id_box = self.driver.find_element(By.ID, 'manager_id')
        id_box.send_keys(self.id)
        
        pw_box = self.driver.find_element(By.ID, 'password')
        pw_box.send_keys(self.pw)
        
        self.driver.find_element(By.CLASS_NAME, 'btn_login').click()
    except:
        self.error.emit("KLAS Login Fail")
        return False
        
    # Books Page
    try:
        self.driver.get("https://klas.jeonju.go.kr/klas3/Books/booksPage/")
        self.driver.execute_script("TitleViewChange('file');")
        # 파일 업로드
        self.driver.find_element(By.CSS_SELECTOR, "input[type='file']").send_keys(self.txt_file_path)
        self.driver.execute_script("speciesFileSearch(1, true);")
        
        # 자료량 기입
        data_num = self.driver.find_element(By.XPATH, '//*[@id="content"]/div[5]/span[2]')
        self.num = re.sub(r'[^0-9]', '', data_num.text)
        self.num = int(self.num)
        
        # 체크박스 선택
        box = self.driver.find_element(By.XPATH, '//*[@id="all_solr_check"]')
        self.driver.execute_script("arguments[0].click()", box)
    except :
        self.error.emit("KLAS file upload Fail")
        return False

    try:
        marc_editor = self.driver.find_element(By.XPATH, '//*[@id="content"]/div[8]/div/input[3]')
        self.driver.execute_script("arguments[0].click()", marc_editor)
        self.driver.switch_to.window(self.driver.window_handles[1])
    except :
        self.error.emit("MarkEditor Access Fail")
        return False
