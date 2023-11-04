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
    reinit = pyqtSignal()
    
    def __init__(self, id, pw, txt_path, xlsx_path, tab2_input, opt):
        super().__init__()

        try:
            chromedriver_autoinstaller.install()
        except Exception:
            self.error.emit("Chrome Driver Error")
        
        # Selenium option
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("detach", True)
        self.driver = None
        
        # KLAS mark editor elements
        self.book_title_box = None
        self.save_btn = None
        self.next_btn = None
        self.syn_tex_box = None
        self.mark_editor = None
        self.ddc = None
        self.is_loding = None
        
        # Loop arguments
        self.id = id
        self.pw = pw
        self.txt_file_path = txt_path
        self.xlsx_file_path = xlsx_path
        self.tab2_input = tab2_input
        self.opt = opt

        self.title = ""
        self.num = 0
        self.val = 0
        
        self.pattern = self.get_pattern()
        self.work_function = self.get_work_function()
        self.loop_function = self.get_loop_function()

    def get_loop_function(self) :
        self.loop_functions = {
            1: self.loop1,
            2: self.loop2,
            3: self.loop3,
            4: self.loop4
        }
        return self.loop_functions.get(self.opt)

    def get_pattern(self):
        re_mapping = {
            5: r'\[\d+\]',
            6: r'\[(.*?)\]',
            7: r'\[[^\]]*\]',
        }
        return re.compile(re_mapping.get(self.opt, r'\[.*\]'))
    
    def get_work_function(self):
        work_mapping = {
            5: self.remove,
            6: self.remove,
            7: self.remove
        }
        return work_mapping.get(self.opt, self.insert)
        
    def common(self):
        ddc_text = str(self.ddc.get_attribute('value'))
        if not ddc_text.isnumeric() and len(ddc_text) > 0:
            self.ddc.clear()
            self.ddc.send_keys(Keys.ENTER)

        self.driver.execute_script("arguments[0].click()", self.save_btn)
        try:
            WebDriverWait(self.driver, 3).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            alert.accept()
        except:
            pass

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
        self.driver.get("https://klas.jeonju.go.kr/klas3/Admin/")
        id_box = self.driver.find_element(By.XPATH, '//*[@id="manager_id"]')
        id_box.send_keys(self.id)
        pw_box = self.driver.find_element(By.XPATH, '//*[@id="password"]')
        pw_box.send_keys(self.pw)
        self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/article[1]/input[4]').click()

        # Books Page
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
        self.num = int(self.num)

        select_all_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="all_solr_check"]'))
        )
        self.driver.execute_script("arguments[0].click()", select_all_box)
        
        enter_marc_editor_btn = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[8]/div/input[3]'))
        )
        self.driver.execute_script("arguments[0].click()", enter_marc_editor_btn)
        self.driver.switch_to.window(self.driver.window_handles[1])


    def column_A(self):
        workbook = openpyxl.load_workbook(self.xlsx_file_path)
        worksheet = workbook.active
        for cell in worksheet['A']:
            self.col.append(cell.value)

    def loop1(self):
        self.title = self.col[self.val] + self.title

    def loop2(self):
        self.title = self.title + self.col[self.val]

    def loop3(self):
        self.title = self.tab2_input + self.title

    def loop4(self):
        self.title = self.title + self.tab2_input

    def insert(self):
        if not re.search(self.pattern, self.title):
            self.loop_function()

    def remove(self):        
        matches = self.pattern.findall(self.title)
        if matches:
            self.title = re.sub(self.pattern, '', self.title)
            self.title = self.title.strip()

    def run(self):
        self.context.emit("진행중")

        if self.xlsx_file_path != "":
            self.col = []
            self.column_A()

        try :
            self.driver = webdriver.Chrome(options=self.options)
            self.klas_upload()
        except:
            self.error.emit("Klas 접속 실패")
            self.driver.quit()
            self.context.emit("재시작")
            self.reinit.emit()
            return

        # enter mark editor tab & find essential elements
        self.book_title_box = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="book_title"]')))
        self.save_btn = self.driver.find_element(By.XPATH, '//*[@id="marcForm"]/div[2]/div[1]/input[18]')
        self.next_btn = self.driver.find_element(By.XPATH, '//*[@id="nextBtn"]')
        self.syn_tex_box = self.driver.find_element(By.XPATH, '//*[@id="syntexMsg_div"]')
        self.mark_editor = self.driver.find_element(By.XPATH, '//*[@id="marcEditor"]')
        self.ddc = self.driver.find_element(By.XPATH, '//*[@id="ddc_class"]')
        self.is_loding = self.driver.find_element(By.XPATH, '/html/body/div[5]')


        while True:
            if self.flag == False:
                time.sleep(1)
            else:
                if self.val < self.num:
                    self.title = str(self.book_title_box.get_attribute('value'))
                    self.book_title_box.clear()
                    self.work_function()
                    self.book_title_box.send_keys(self.title)
                    self.book_title_box.send_keys(Keys.ENTER)
                    self.val += 1
                    self.common()
                    WebDriverWait(self.driver, 10).until(EC.invisibility_of_element_located((By.ID, "divLoadingBar")))
                    self.progress.emit(int(self.val * 100 / self.num))
                else:
                    self.context.emit("작업 완료")
                    self.reinit.emit()
                    return
