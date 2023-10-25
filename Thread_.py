import re
import time
import openpyxl
import os
import chromedriver_autoinstaller
from PyQt5.QtCore import QThread, pyqtSignal
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

class Thread(QThread):
    # signal
    progress = pyqtSignal(int)
    context = pyqtSignal(str)
    error = pyqtSignal(str)
    
    # stop
    flag = True
    
    # data val
    num = 0  # 총 자료 개수
    val = 0  # 현재 자료 위치
    
    # Selenium
    driver = None
    options = None
    
    # elements
    book_title_box = None
    save_btn = None
    next_btn = None
    syn_tex_box = None
    mark_editor = None

    ddc = None
    dup_msg = None
    dup_btn = None

    # loop option
    opt = 0
    col = []
    pattern = r'\[.*\]'
    title = ""
    
    def __init__(self, id, pw, txt_path, xlsx_path, tab2_input, opt):
        super().__init__()

        try:
            chromedriver_autoinstaller.install()
        except Exception:
            self.error.emit("Chrome Driver Error")
        
        # Selenium option
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("detach", True)
        
        # Loop arguments
        self.id = id
        self.pw = pw
        self.txt_file_path = txt_path
        self.xlsx_file_path = xlsx_path
        self.tab2_input = tab2_input
        self.opt = opt
        
        self.method_mapping = {
            1: self.loop1,
            2: self.loop2,
            3: self.loop3,
            4: self.loop4,
        }
        
        self.re_mapping = {
            5: re.compile(r'\[\d+\]'),
            6: re.compile(r"\[^\d]+\]"),
            7: re.compile(r'\[[^\]]*\]'),
        }
        
    def alert_ok(self):
        try:
            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            alert.accept()
        except:
            pass

    def common(self):
        WebDriverWait(self.driver, 10)

        ddc_text = str(self.ddc.get_attribute('value'))
        if not ddc_text.isnumeric() and len(ddc_text) > 0:
            self.ddc.clear()
            self.ddc.send_keys(Keys.ENTER)

        self.driver.execute_script("arguments[0].click()", self.save_btn)
        self.alert_ok()

        dup_msg = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]')))
        
        if dup_msg.is_displayed() == True:
            dup_btn = self.driver.find_element(By.XPATH, '//*[@id="confirm_ok"]')
            self.driver.execute_script("arguments[0].click()", dup_btn)

        try:
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
        try:
            self.driver.get("https://klas.jeonju.go.kr/klas3/Admin/")
            
            id_box = self.driver.find_element(By.XPATH, '//*[@id="manager_id"]')
            id_box.send_keys(self.id)
            
            pw_box = self.driver.find_element(By.XPATH, '//*[@id="password"]')
            pw_box.send_keys(self.pw)
            
            self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/article[1]/input[4]').click()
        except:
            self.error.emit("KLAS Login Fail")
            return False
        
        # Books Page
        try:
            self.driver.get("https://klas.jeonju.go.kr/klas3/Books/booksPage/")
            self.driver.execute_script("TitleViewChange('file');")

            upload_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="reg_no_file"]'))
            )
            upload_box.send_keys(self.txt_file_path)
            self.driver.execute_script("speciesFileSearch(1, true);")


            label_data_num = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[5]/span[2]'))
            )

            self.num = re.sub(r'[^0-9]', '', label_data_num.text)
            self.num = int(self.num) - 1

            select_all_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="all_solr_check"]'))
            )
            self.driver.execute_script("arguments[0].click()", select_all_box)
        except:
            self.error.emit("KLAS file upload Fail")
            return False
        
        try:
            enter_marc_editor_btn = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[8]/div/input[3]'))
            )
            self.driver.execute_script("arguments[0].click()", enter_marc_editor_btn)
            self.driver.switch_to.window(self.driver.window_handles[1])
        except:
            self.error.emit("MarkEditor Access Fail")
            return False

    def column_A(self):
        # 엑셀 파일 열기
        workbook = openpyxl.load_workbook(self.xlsx_file_path)

        # 첫 번째 시트 선택
        worksheet = workbook.active

        # A열의 모든 값을 리스트로 저장
        for cell in worksheet['A']:
            self.col.append(cell.value)

    def finish(self):
        self.book_title_box.send_keys(self.title)
        self.book_title_box.send_keys(Keys.ENTER)
        self.progress.emit(int(self.val * 100 / self.num))

    def loop1(self):
        self.title = self.col[self.val] + self.title

    def loop2(self):
        self.title = self.title + self.col[self.val]

    def loop3(self):
        self.title = self.tab2_input + self.title

    def loop4(self):
        self.title = self.title + self.tab2_input

    def remove(self):
        self.title = str(self.book_title_box.get_attribute('value'))
        matches = self.pattern.findall(self.title)
        if matches:
            self.book_title_box.clear()
            self.title = re.sub(self.pattern, '', self.title)
            self.title = self.title.strip()
            self.finish()
            self.common()

    def loop(self):
        self.title = str(self.book_title_box.get_attribute('value'))
        if not re.search(self.pattern, self.title):
            method_to_call = self.method_mapping.get(self.opt)
            method_to_call()
            self.book_title_box.clear()
            self.finish()
        self.common()

    def run(self):
        self.context.emit("진행중")

        if self.xlsx_file_path != "":
            self.column_A()

        self.driver = webdriver.Chrome(options=self.options)
        self.flag = self.klas_upload()

        # enter mark editor tab & find essential elements
        self.book_title_box = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="book_title"]')))
        self.save_btn = self.driver.find_element(By.XPATH, '//*[@id="marcForm"]/div[2]/div[1]/input[18]')
        self.next_btn = self.driver.find_element(By.XPATH, '//*[@id="nextBtn"]')
        self.syn_tex_box = self.driver.find_element(By.XPATH, '//*[@id="syntexMsg_div"]')
        self.mark_editor = self.driver.find_element(By.XPATH, '//*[@id="marcEditor"]')
        self.ddc = self.driver.find_element(By.XPATH, '//*[@id="ddc_class"]')

        if self.opt > 4:
            re_to_call = self.re_mapping.get(self.opt)
            self.pattern = re_to_call()
            while True:
                if self.flag == False:
                    time.sleep(1)  # 일시정지
                else:
                    if self.val <= self.num:
                        self.driver.implicitly_wait(2)
                        self.remove()
                        self.val += 1
                    else:
                        self.context.emit("작업 완료")
                        return
        else:
            while True:
                if self.flag == False:
                    time.sleep(1)  # 일시정지
                else:
                    if self.val <= self.num:
                        self.driver.implicitly_wait(2)
                        self.loop()
                        self.val += 1
                    else:
                        self.context.emit("작업 완료")
                        return
