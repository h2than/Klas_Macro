import re
import os

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

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
