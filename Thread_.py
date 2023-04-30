from PyQt5.QtCore import QThread,pyqtSignal
import chromedriver_autoinstaller
from selenium import webdriver
from prepare import klas_upload,common,loop1,loop2,loop3,loop4,loop5,loop6,loop7
<<<<<<< HEAD

class Thread(QThread):
    # signal
    finished = pyqtSignal()
=======
import time

class Thread(QThread):
    # signal
>>>>>>> origin/V1.2
    progress = pyqtSignal(int)
    context = pyqtSignal(str)
    error = pyqtSignal(str)
    
    flag = True
    
    driver = None # Selenium driver
    options = None # Selenium options
    
    num = 0 # 총 자료 개수
    val = 0 # 현재 자료 위치
    
    id = ""
    pw = ""
    txt_file_path = ""
    xlsx_file_path = ""
    input_text = ""
    
    book_title_box = None
    save_btn = None
    next_btn = None
    syn_tex_box = None
<<<<<<< HEAD
    # //*[@id="books_print_count"] -> 1000개
=======
>>>>>>> origin/V1.2

    select = 0
    
    def __init__(self , id, pw, txt_path , xlsx_path, input_text, select):
        super().__init__()
        try:
            chromedriver_autoinstaller.install()
<<<<<<< HEAD
        except:
            raise Exception("Chrome Driver Not Installed")
=======
        except Exception:
            self.error.emit("Chrome Driver Error")
            self.terminate()
>>>>>>> origin/V1.2
        
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
<<<<<<< HEAD
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
=======
            loop1(self.book_title_box,self.xlsx_file_path,self.val)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        elif self.select == 2:
            loop2(self.book_title_box,self.xlsx_file_path,self.val)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        elif self.select == 3:
            loop3(self.book_title_box,self.input_text)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        elif self.select == 4:
            loop4(self.book_title_box,self.input_text)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        elif self.select == 5:
            loop5(self.book_title_box)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        elif self.select == 6:
            loop6(self.book_title_box)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
        elif self.select == 7:
            loop7(self.book_title_box)
            self.val += 1
            self.progress.emit(int(self.val*100/self.num))
            common(self)
>>>>>>> origin/V1.2
        
    def run(self):
        self.context.emit("진행중")
        self.driver = webdriver.Chrome(options=self.options)
        try:
            klas_upload(self)
<<<<<<< HEAD
        except Exception as e:
            self.error.emit(str(e))
            self.stop()
        try:
            self.loop()
        except Exception as e:
            self.error.emit(str(e))
            self.stop()
            
        self.context.emit("종료")
        self.finished.emit()
        self.stop()
        
    def stop(self):
        self.quit()
        self.wait(5000)
=======
        except Exception :
            self.error.emit("Login Error")
            self.terminate()
        while True:
            if self.flag == False:
                time.sleep(3)
            elif self.flag and self.val <= self.num:
                self.loop()
            else:
                self.context.emit("작업 완료")
                self.quit()
>>>>>>> origin/V1.2
