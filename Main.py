import sys

from PyQt5.QtWidgets import QMainWindow,QApplication,QFileDialog,QMessageBox,QDialog
from PyQt5.QtCore import QDir
from gui import Ui_MainWindow

from Thread_ import Thread

class WindowClass(QMainWindow, Ui_MainWindow):
    
    # 변수
    tab2_input = ""
    id = ""
    pw = ""
    txt_path = ""
    xlsx_path = ""
    opt = 0

    def __init__(self):
        super( ).__init__()
        self.setupUi(self)
        
        # Tab 
        self.TabWidget.currentChanged.connect(self.tab_changed)

        self.mainthread = None

        # 상단 Frame
        self.btn_txt_add.clicked.connect(self.txt_add)
        self.btn_login.clicked.connect(self.login) # label_id , label_pw
        # 종료 버튼
        self.btn_Quit.clicked.connect(self.exit)

        # tab1 xlsx
        self.btn_tab1.clicked.connect(self.tab1_apt)
        self.tab1_cb0.stateChanged.connect(self.tab1_chk)
        self.tab1_cb1.stateChanged.connect(self.tab1_chk)
        self.btn_start0.clicked.connect(self.tab1_start_btn)
        self.btn_pause0.clicked.connect(self.pause)
    
        # tab2 [text]
        self.btn_tab2.clicked.connect(self.tab2_apt) # label_tab2
        self.tab2_cb0.stateChanged.connect(self.tab2_chk)
        self.tab2_cb1.stateChanged.connect(self.tab2_chk)
        self.btn_start1.clicked.connect(self.tab2_start_btn)
        self.btn_pause1.clicked.connect(self.pause)
        
        # tab3 remove
        self.tab3_cb0.stateChanged.connect(self.tab3_chk)
        self.tab3_cb1.stateChanged.connect(self.tab3_chk)
        self.btn_start2.clicked.connect(self.tab3_start_btn)
        self.btn_pause2.clicked.connect(self.pause)

    # 탭 전환시
    def tab_changed(self):
        self.opt = 0
        self.tab1_cb0.setChecked(False)
        self.tab1_cb1.setChecked(False)
        self.tab2_cb0.setChecked(False)
        self.tab2_cb1.setChecked(False)
        self.tab3_cb0.setChecked(False)
        self.tab3_cb1.setChecked(False)
        
    def essential_val_check(self):
        if self.label_top.text() == "등록번호":
            self.errormsg(text="등록번호 파일을 추가해 주세요")
            return False
        
        if self.id == "" or self.pw == "" :
            self.errormsg("id, pw 를 입력하세요")
            return False

        if self.opt <= 2 :
            if self.opt == 0 or self.xlsx_path == "" :
                self.errormsg(text="옵션을 다시 확인 해주세요")
                return False  
        elif self.opt <= 4 :
            if self.opt == 0 or self.tab2_input == "" :
                self.errormsg(text="옵션을 다시 확인 해주세요")
                return False

        return True

    def macro_run(self):
        is_ready = self.essential_val_check()
        if is_ready:
                self.btn_start0.setDisabled(True)
                self.btn_start1.setDisabled(True)
                self.btn_start2.setDisabled(True)
                self.mainthread = Thread(id=self.id, pw=self.pw, txt_path=self.txt_path,xlsx_path=self.xlsx_path,tab2_input=self.tab2_input, opt=self.opt)
                self.mainthread.error.connect(self.errormsg)
                self.mainthread.reinit.connect(self.re_init)
                self.mainthread.context.connect(self.state)
                self.mainthread.progress.connect(self.onProgress)
                self.mainthread.start()
            
    # 등록번호 upload
    def txt_add(self):
        default_path = QDir.homePath() +'/Desktop'
        
        file_dialog = QFileDialog(self,'Open File', default_path, 'txt Files (*.txt)')
        file_dialog.setDirectory(default_path)
        
        if file_dialog.exec_() == QDialog.Accepted:
            self.txt_path = file_dialog.selectedFiles()[0]
            fileName = file_dialog.selectedFiles()[0].split('/')[-1]
            self.label_top.setText(fileName)
    
    # id,pw 
    def login(self):
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
    
    # 엑셀 파일 추가
    
    def tab1_apt(self):
        default_path = QDir.homePath() +'/Desktop'
        file_dialog = QFileDialog(self,'Open File', default_path, 'Xlsx Files (*.xlsx)')
        file_dialog.setDirectory(default_path)
        
        if file_dialog.exec_() == QDialog.Accepted:
            self.xlsx_path = file_dialog.selectedFiles()[0]
            fileName = file_dialog.selectedFiles()[0].split('/')[-1]
            self.label_tab1.setText(fileName)

    # 탭1 옵션 체크   
    def tab1_chk(self):
        self.opt = 0
        if self.tab1_cb0.isChecked() :
            self.tab1_cb1.setChecked(False)
            self.opt = 1
        if self.tab1_cb1.isChecked() :
            self.tab1_cb0.setChecked(False)
            self.opt = 2
                
    # 탭 1 시작 버튼
    def tab1_start_btn(self):
        self.macro_run()
    # 탭 2 텍스트 적용

    def tab2_apt(self):
        if self.label_tab2.isEnabled():
            self.tab2_input = self.label_tab2.text()
            self.label_tab2.setDisabled(True)
        else :
            self.label_tab2.clear()
            self.tab2_input = ""
            self.label_tab2.setDisabled(False)
        
    # 탭 2 텍스트 옵션 체크
    def tab2_chk(self):
        self.opt = 0
        if self.tab2_cb0.isChecked() :
            self.tab2_cb1.setChecked(False)
            self.opt = 3
        if self.tab2_cb1.isChecked() :
            self.tab2_cb0.setChecked(False)
            self.opt = 4

    # 탭 2 시작 버튼
    def tab2_start_btn(self):
        self.macro_run()
        
    # 탭 3 옵션 
    def tab3_chk(self):
        self.opt = 0
        if self.tab3_cb0.isChecked() and self.tab3_cb1.isChecked():
            self.opt = 7
        else:
            if self.tab3_cb0.isChecked():
                self.opt = 5
            if self.tab3_cb1.isChecked():
                self.opt = 6
                
    # 탭 3 시작 버튼
    def tab3_start_btn(self):
        self.macro_run()

    def errormsg(self, text):
        error_box = QMessageBox(self)
        error_box.setIcon(QMessageBox.Critical)
        error_box.setText(text)
        error_box.setWindowTitle("Error")
        error_box.exec_()
    
    def pause(self):
        try:
            if self.mainthread.flag :
                self.mainthread.flag = False
                self.label_status.setText("멈춤")
                self.btn_pause0.setText("재시작")
                self.btn_pause1.setText("재시작")
                self.btn_pause2.setText("재시작")
            else:
                self.mainthread.flag = True
                self.label_status.setText("진행중")
                self.btn_pause0.setText("일시정지")
                self.btn_pause1.setText("일시정지")
                self.btn_pause2.setText("일시정지")
        except:
            pass
        
        
    def stop(self):
        self.mainthread.driver.quit()
        self.mainthread.flag = False
        self.mainthread.terminate()
        
    # 강제 종료
    def exit(self):
        self.stop()
        QApplication.instance().exit()
        
    def re_init(self):
        self.stop()
            
        self.btn_start0.setDisabled(False)
        self.btn_start1.setDisabled(False)
        self.btn_start2.setDisabled(False)
                
        self.tab2_input = ""
        self.txt_path = ""
        self.xlsx_path = ""
        self.opt = 0
        self.label_top.setText("등록번호")
        self.label_tab1.setText("Excel file")
        self.label_tab2.setText("[출판사]")
        self.label_tab2.setDisabled(False)
        self.progressBar.setValue(0)

    # 스레드 slot
    def state(self, context):
        self.label_status.setText(context)
            
    def onProgress(self, value):
        self.progressBar.setValue(value)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = WindowClass( )
    myWindow.show( )
    sys.exit(app.exec_())

# https://github.com/h2than/Klas_Macro.git
