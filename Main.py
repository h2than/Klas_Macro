import sys
import os

from PyQt5.QtWidgets import QMainWindow,QApplication,QFileDialog,QMessageBox,QDialog
from PyQt5.QtCore import Qt,QDir
from PyQt5 import uic

from Thread_ import *

# 에러 메세지
def login_check(id,pw):
    if id == "" or pw == "":
        errormsg("id, pw 를 입력하세요")
    else : return True

def errormsg(text):
    error_box = QMessageBox()
    error_box.setIcon(QMessageBox.Critical)
    error_box.setText(text)
    error_box.setWindowTitle("Error")
    error_box.exec_()

def infomsg(text):
    info_box = QMessageBox()
    info_box.setIcon(QMessageBox.information)
    info_box.setText(text)
    info_box.setWindowTitle("알림")
    info_box.exec_()

# UI 파일 경로

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = os.path.join(sys._MEIPASS, 'resources')
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

Ui_MainWindow, _ = uic.loadUiType(resource_path('main.ui'))

# PyQt5 메인 윈도우
class WindowClass(QMainWindow, Ui_MainWindow):
    
    # 변수
    id = ""
    pw = ""
    txt_file_path = ""
    xlsx_file_path = ""
    opt = 0 
    input_text = ""
    
    
    def __init__(self):
        super( ).__init__()
        self.setupUi(self)
        
        # Tab 
        self.TabWidget.currentChanged.connect(self.tab_changed)

        self.t1 = None

        # 상단 Frame
        self.btn_txt_add.clicked.connect(self.txt_add)
        self.btn_xlsx_add.clicked.connect(self.xlsx_add)
        self.btn_id.clicked.connect(self.adp_id_pw) # label_id , label_pw
        # 종료 버튼
        self.btn_Quit.clicked.connect(self.exit)

        # 추가
        self.CB_add_front_file.stateChanged.connect(self.chk_file)
        self.CB_add_back_file.stateChanged.connect(self.chk_file)
        self.btn_start.clicked.connect(self.start_btn)
        self.btn_pause_file.clicked.connect(self.pause)
        
        # 입력 
        self.btn_input.clicked.connect(self.adp_input_text) # label_input
        self.btn_start_2.clicked.connect(self.start_btn_2)
        self.CB_back_add_file.stateChanged.connect(self.input_title_chk)
        self.CB_front_add_text.stateChanged.connect(self.input_title_chk)
        self.btn_pause_txt.clicked.connect(self.pause)
        
        # 자동
        self.CB_auto_int.stateChanged.connect(self.chk_auto)
        self.CB_auto_txt.stateChanged.connect(self.chk_auto)
        self.btn_start_3.clicked.connect(self.start_btn_3)
        self.btn_pause_auto.clicked.connect(self.pause)
        
    # 탭 전환시
    def tab_changed(self):
        self.opt = 0
        
    # 일시정지
    def pause(self):
        if self.t1.flag == True:
            self.t1.flag == False
            self.label_ing.setText("일시정지")
        else : 
            self.t1.flag == True
            self.label_ing.setText("진행중")
        
    # 강제 종료
    def exit(self):
        if self.t1 == None:
            self.t1.terminate()
            self.wait(5000)
            self.t1 = None
        app.exec_( )
        
        
    # 상단 Frame
    def txt_add(self):
        
        default_path = QDir.homePath() +'/Desktop'
        
        file_dialog = QFileDialog(self,'Open File', default_path, 'txt Files (*.txt)')
        file_dialog.setDirectory(default_path)
        
        if file_dialog.exec_() == QDialog.Accepted:
            self.txt_file_path = file_dialog.selectedFiles()[0]
            fileName = file_dialog.selectedFiles()[0].split('/')[-1]
            self.label_txt.setText(fileName)
    
    # 상단 Frame id, pw 적용
    def adp_id_pw(self):

        if self.label_id.isEnabled():
            self.id = self.label_id.text()
            self.pw = self.label_pw.text()
            self.label_id.setDisabled(True)
            self.label_pw.setDisabled(True)
        else :
            self.label_id.clear()
            self.label_pw.clear()
            self.id = ""
            self.pw = ""
            self.label_id.setDisabled(False)
            self.label_pw.setDisabled(False)
    
# 탭 위젯
    
    # 탭 1 파일 추가
    
    def xlsx_add(self):
        
        default_path = QDir.homePath() +'/Desktop'
        
        file_dialog = QFileDialog(self,'Open File', default_path, 'Xlsx Files (*.xlsx)')
        file_dialog.setDirectory(default_path)
        
        if file_dialog.exec_() == QDialog.Accepted:
            self.xlsx_file_path = file_dialog.selectedFiles()[0]
            fileName = file_dialog.selectedFiles()[0].split('/')[-1]
            self.label_xlsx.setText(fileName)
        else:
            self.xlsx_file_path = ""
    
    # 탭1 옵션 체크   
    def chk_file(self,state):
        
        self.opt = 0
        if state == Qt.Checked:
            if self.sender() == self.CB_add_front_file:
                self.CB_add_back_file.setChecked(False)
                self.opt = 1
            elif self.sender() == self.CB_add_back_file:
                self.CB_add_front_file.setChecked(False)
                self.opt = 2
                
    # 탭 1 시작 버튼
    def start_btn(self):
        if self.opt == 0 and self.xlsx_file_path == ""  :
            errormsg(text="옵션을 다시 확인 해주세요")
            if self.label_txt.text() == "등록번호":
                errormsg(text="등록번호 파일을 추가해 주세요")
        else :
            if login_check(self.id,self.pw) :
                try:
                    if self.opt == 1:
                        self.btn_start_2.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path=self.xlsx_file_path,input_text="", select=1)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                    else:
                        self.btn_start_2.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path=self.xlsx_file_path,input_text="", select=2)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                except:
                    pass

    # 탭 2 텍스트 적용

    def adp_input_text(self):
                
        if self.label_input.isEnabled():
            self.input_text = self.label_input.text()
            self.label_input.setDisabled(True)
        else :
            self.label_input.clear()
            self.input_text = ""
            self.label_input.setDisabled(False)
        
    # 탭 2 텍스트 옵션 체크
    def input_title_chk(self,state):
        
        self.opt = 0
        if state == Qt.Checked:
            if self.sender() == self.CB_front_add_text:
                self.CB_back_add_file.setChecked(False)
                self.opt = 1
            elif self.sender() == self.CB_back_add_file:
                self.CB_front_add_text.setChecked(False)
                self.opt = 2

    # 탭 2 시작 버튼
    def start_btn_2(self):
        if self.opt == 0 or self.input_text == "" :
            errormsg(text="옵션을 다시 확인 해주세요")
            if self.label_txt.text() == "등록번호":
                errormsg(text="등록번호 파일을 추가해 주세요")
        else :
            if login_check(self.id,self.pw) :
                try:
                    if self.opt == 1:
                        self.btn_start_2.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path="",input_text=self.input_text, select=3)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                    else:
                        self.btn_start_2.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path="",input_text=self.input_text, select=4)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                except:
                    pass
        
    # 탭 3 옵션 
    def chk_auto(self):
        self.opt = 0
        if self.CB_auto_int.isChecked():
            self.opt += 1
        if self.CB_auto_txt.isChecked():
            self.opt += 2
            
    # 탭 3 시작 버튼
    def start_btn_3(self):
        if not self.CB_auto_int.isChecked() and not self.CB_auto_txt.isChecked() :
            errormsg(text="체크 박스를 선택해 주세요!")
            if self.label_txt.text() == "등록번호":
                errormsg(text="등록번호 파일을 추가해 주세요")
        else:
            if login_check(self.id,self.pw) :
                try:
                    if self.opt == 1:
                        self.btn_start_3.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path="",input_text="", select=5)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                    elif self.opt == 2:
                        self.btn_start_3.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path="",input_text="", select=6)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                    else:
                        self.btn_start_3.setEnabled(False)
                        self.t1 = Thread(id=self.id, pw=self.pw, txt_path=self.txt_file_path,xlsx_path="",input_text="", select=7)
                        self.t1.context.connect(self.state)
                        self.t1.finished.connect(self.onFinished)
                        self.t1.progress.connect(self.onProgress)
                        self.t1.error.connect(self.handel_error)
                        self.t1.start()
                except:
                    pass
                
    # 스레드 slot
    def handel_error(self, err_str):
        errormsg(text=err_str)
        self.exit()
    
    def state(self, context):
        self.label_ing.setText(context)
            
    def onProgress(self, value):
        self.progressBar.setValue(value)
            
    def onFinished(self):
        self.label_ing.setText("완료")
        infomsg(text="작업 완료")
        self.exit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = WindowClass( )
    myWindow.show( )
    app.exec_( )