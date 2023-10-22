from PyQt5.QtCore import QThread,pyqtSignal
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import time
import openpyxl
import re
import os


class Thread(QThread):
    # signal
    progress = pyqtSignal(int)
    context = pyqtSignal(str)
    error = pyqtSignal(str)
    
    # stop
    flag = True
    
    # data val
    num = 0 # 총 자료 개수
    val = 0 # 현재 자료 위치
    
    # Selenium
    driver = None
    options = None
    
    # elements
    book_title_box = None
    save_btn = None
    next_btn = None
    syn_tex_box = None
    mark_editor = None

    # loop option
    select = 0
    col = []
    pattern = r'\[.*\]'
    title = ""
    
    def __init__(self , id, pw, txt_path , xlsx_path, tab2_input, select):
        super().__init__()
        try:
            chromedriver_autoinstaller.install()
        except Exception:
            self.error.emit("Chrome Driver Error")
        
        # selenium option
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("detach", True)
        
        # Loop arguments
        self.id = id
        self.pw = pw
        self.txt_file_path = txt_path
        self.xlsx_file_path = xlsx_path
        self.tab2_input = tab2_input
        self.select = select

    def alert_ok(self):
        try:
            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            alert.accept()
        except:
            pass

    def common(self):

        WebDriverWait(self.driver, 10)

        ddc = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ddc_class"]')))
        ddc_text = str(ddc.get_attribute('value'))
        if not ddc_text.isnumeric() and len(ddc_text) > 0:
            ddc.clear()
            ddc.send_keys(Keys.ENTER)

        self.driver.execute_script("arguments[0].click()", self.save_btn)
        self.alert_ok(self)

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

    def column_A(self):
        # 엑셀 파일 열기
        workbook = openpyxl.load_workbook(self.xlsx_file_path)

        # 첫 번째 시트 선택
        worksheet = workbook.active

        # A열의 모든 값을 리스트로 저장
        self.col = []
        for cell in worksheet['A']:
            self.col.append(cell.value)
        
    def finish(self):
        self.book_title_box.send_keys(self.title)
        self.book_title_box.send_keys(Keys.ENTER)
        self.progress.emit(int(self.val*100/self.num))
        
    def loop(self):
        self.title = str(self.book_title_box.get_attribute('value'))
        if self.select == 1:
            if re.search(self.pattern, self.title):
                return
            self.title = self.col[self.val] + self.title
            self.book_title_box.clear()
            self.finish()
            self.common(self)

        elif self.select == 2:
            if re.search(self.pattern, self.title):
                return
            self.title = self.title + self.col[self.val]
            self.book_title_box.clear()
            self.finish()
            self.common(self)

        elif self.select == 3:
            if re.search(self.pattern, self.title):
                return
            self.title = self.tab2_input + self.title
            self.book_title_box.clear()
            self.finish()
            self.common(self)

        elif self.select == 4:
            self.title = self.title + self.tab2_input
            self.book_title_box.clear()
            self.finish()
            self.common(self)
        
        elif self.select <= 5:
            matches = self.pattern.findall(self.title)
            if matches:
                self.book_title_box.clear()
                self.title = re.sub(self.pattern, '',self.title)
                self.title = self.title.strip()
                self.finish()
            self.common(self)
        
    def run(self):
        self.context.emit("진행중")

        if self.xlsx_file_path is not "" :
            self.column_A()
    
        if self.select == 7 :
            self.pattern = re.compile(r'\[[^\]]*\]')
        elif self.select == 6 :
            self.pattern = re.compile(r"\[^\d]+\]")
        elif self.select == 5 :
            self.pattern = re.compile(r'\[\d+\]')

        self.driver = webdriver.Chrome(options=self.options)
        self.flag = self.klas_upload()
        time.sleep(2)
        # enter mark editor tab & find essential elements
        self.book_title_box = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="book_title"]')))    
        self.save_btn = self.driver.find_element(By.XPATH, '//*[@id="marcForm"]/div[2]/div[1]/input[18]')
        self.next_btn = self.driver.find_element(By.XPATH, '//*[@id="nextBtn"]')
        self.syn_tex_box = self.driver.find_element(By.XPATH, '//*[@id="syntexMsg_div"]')
        self.mark_editor = self.driver.find_element(By.XPATH, '//*[@id="marcEditor"]')
        
        while True:
            if(self.flag == False):
                time.sleep(1) # 일시정지
            else :
                if self.val <= self.num:
                    self.driver.implicitly_wait(2)
                    self.loop()
                    self.val += 1
                else:
                    self.context.emit("작업 완료")
                    break

