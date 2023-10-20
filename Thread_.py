from PyQt5.QtCore import QThread,pyqtSignal
import chromedriver_autoinstaller
from selenium import webdriver
from prepare import *
import time

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
    end_count = 0 # common() 루프방지
    
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
        
    def loop(self):
        if self.select == 1:
            loop1(self.book_title_box,self.xlsx_file_path,self.val)
            self.progress.emit(int(self.val*100/self.num))
            common(self)

        elif self.select == 2:
            loop2(self.book_title_box,self.xlsx_file_path,self.val)
            self.progress.emit(int(self.val*100/self.num))
            common(self)

        elif self.select == 3:
            loop3(self.book_title_box,self.tab2_input)
            self.progress.emit(int(self.val*100/self.num))
            common(self)

        elif self.select == 4:
            loop4(self.book_title_box,self.tab2_input)
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        
        elif self.select == 5:
            loop5(self.book_title_box)
            self.progress.emit(int(self.val*100/self.num))
            common(self)

        elif self.select == 6:
            loop6(self.book_title_box)
            self.progress.emit(int(self.val*100/self.num))
            common(self)

        elif self.select == 7:
            loop7(self.book_title_box)
            self.progress.emit(int(self.val*100/self.num))
            common(self)

        
    def run(self):
        self.context.emit("진행중")
        self.driver = webdriver.Chrome(options=self.options)
        self.flag = klas_upload(self)
        # enter mark editor tab & find essential elements
        self.book_title_box = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="book_title"]')))    
        self.save_btn = self.driver.find_element(By.XPATH, '//*[@id="marcForm"]/div[2]/div[1]/input[18]')
        self.next_btn = self.driver.find_element(By.XPATH, '//*[@id="nextBtn"]')
        self.syn_tex_box = self.driver.find_element(By.XPATH, '//*[@id="syntexMsg_div"]')
        self.mark_editor = self.driver.find_element(By.XPATH, '//*[@id="marcEditor"]')
        
        time.sleep(2)
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

