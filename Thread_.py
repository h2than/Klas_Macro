from PyQt5.QtCore import QThread,pyqtSignal
import chromedriver_autoinstaller
from selenium import webdriver
from prepare import *
import time
import openpyxl

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

    def column_A(self):
        # 엑셀 파일 열기
        workbook = openpyxl.load_workbook(self.xlsx_file_path)

        # 첫 번째 시트 선택
        worksheet = workbook.active

        # A열의 모든 값을 리스트로 저장
        column_a = []
        for cell in worksheet['A']:
            column_a.append(cell.value)
            
        return column_a

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
            common(self)

        elif self.select == 2:
            if re.search(self.pattern, self.title):
                return
            self.title = self.title + self.col[self.val]
            self.book_title_box.clear()
            self.finish()
            common(self)

        elif self.select == 3:
            if re.search(self.pattern, self.title):
                return
            self.title = self.tab2_input + self.title
            self.book_title_box.clear()
            self.finish()
            common(self)

        elif self.select == 4:
            self.title = self.title + self.tab2_input
            self.book_title_box.clear()
            self.finish()
            common(self)
        
        elif self.select <= 5:
            matches = self.pattern.findall(self.title)
            if matches:
                self.book_title_box.clear()
                self.title = re.sub(self.pattern, '',self.title)
                self.title = self.title.strip()
                self.finish()
            common(self)
        
    def run(self):
        self.context.emit("진행중")

        if self.xlsx_file_path is not "" :
            self.col = self.column_A()
    
        if self.select == 7 :
            self.pattern = re.compile(r'\[[^\]]*\]')
        elif self.select == 6 :
            self.pattern = re.compile(r"\[^\d]+\]")
        elif self.select == 5 :
            self.pattern = re.compile(r'\[\d+\]')

        self.driver = webdriver.Chrome(options=self.options)
        self.flag = klas_upload(self)
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

