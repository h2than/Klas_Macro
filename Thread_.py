from PyQt5.QtCore import QThread,pyqtSignal
import chromedriver_autoinstaller
from selenium import webdriver
from prepare import klas_upload,common,loop1,loop2,loop3,loop4,loop5,loop6,loop7

class Thread(QThread):
    # signal
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    context = pyqtSignal(str)
    
    driver = None
    options = None
    
    num = 0
    val = 0
    
    id = ""
    pw = ""
    txt_file_path = ""
    xlsx_file_path = ""
    input_text = ""
    
    book_title_box = None
    save_btn = None
    next_btn = None
    syn_tex_box = None
    final = None

    
    select = 0
    
    def __init__(self , id, pw, txt_path , xlsx_path, input_text, select):
        super().__init__()
        self.flag = True
        try:
            chromedriver_autoinstaller.install()
        except:
            pass
        
        # selenium option
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("detach", True)
        
        # Loop arguments
        self.id = id
        self.pw = pw
        self.txt_file_path = txt_path
        self.xlsx_file_path = xlsx_path
        self.input_text = input_text
        self.select = select
        
    def loop(self):
        if self.select == 1:
            while self.flag and self.val <= self.num:
                loop1(self.book_title_box,self.xlsx_file_path,self.val)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        elif self.select == 2:
            while self.flag and self.val <= self.num:
                loop2(self.book_title_box,self.xlsx_file_path,self.val)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        elif self.select == 3:
            while self.flag and self.val <= self.num:
                loop3(self.book_title_box,self.input_text)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        elif self.select == 4:
            while self.flag and self.val <= self.num:
                loop4(self.book_title_box,self.input_text)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        elif self.select == 5:
            while self.flag and self.val <= self.num:
                loop5(self.book_title_box)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        elif self.select == 6:
            while self.flag and self.val <= self.num:
                loop6(self.book_title_box)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        elif self.select == 7:
            while self.flag and self.val <= self.num:
                loop7(self.book_title_box)
                self.val += 1
                self.progress.emit(int(self.val*100/self.num))
                common(self)
        
    def run(self):
        self.context.emit("진행중")
        self.driver = webdriver.Chrome(options=self.options)
        klas_upload(self)
        self.loop()
        self.context.emit("종료")
        self.finished.emit()
        self.driver.quit()
    def stop(self):
        self.driver.quit()
        self.flag = False