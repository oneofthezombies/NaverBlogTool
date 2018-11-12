"""
네이버 블로그 도구
"""

# 시스템
import sys
from multiprocessing import Process, Queue as MPQueue

# 크롬 브라우저 제어 모듈
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# 캡챠 모듈
from python3_anticaptcha import ImageToTextTask

# 엑셀 모듈
from openpyxl import load_workbook

# 데스크탑 앱 모듈
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, QThread
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, 
    QDesktopWidget, QPushButton, QMessageBox)


JOBS_DONE = 'jobs_done' # 멀티프로세스 메시지 큐 종료 플래그


"""
데스크탑 앱
"""
class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle('네이버 블로그 도구') # 앱 이름        
        self.setGeometry(0, 0, 170, 190) # 앱 크기

        # 앱을 화면 가운데로 옮기기
        rect = self.frameGeometry()
        center = QDesktopWidget().availableGeometry().center()
        rect.moveCenter(center)
        self.move(rect.topLeft())

        self.createButton('새글 쓰기', 150, 50, 10, 10, self.writeNewPost) # 새글 쓰기 버튼 생성
        self.createButton('마지막 글 지우기', 150, 50, 10, 70, self.deleteLastPost) # 마지막 글 지우기 버튼 생성
        self.createButton('다른 글 바꾸기', 150, 50, 10, 130, self.modifyOtherPost) # 다른 글 바꾸기 버튼 생성

        self.q = MPQueue() # 멀티프로세스 큐

        # 팝업 메시지박스 스레드
        self.ui_thread = UIThread(self.q)
        self.ui_thread.popup_message.connect(self.createMessageBox)
        self.ui_thread.start()


    """
    팝업 메시지
    """
    @pyqtSlot(str)
    def createMessageBox(self, message):
        QMessageBox.about(self, '', message)


    """
    버튼 만들기
    """
    def createButton(self, text, width, height, x, y, callback):
        btn = QPushButton(text, self) # 버튼 만들기
        btn.resize(width, height) # 크기
        btn.move(x, y) # 위치 (좌상단 기준)
        btn.clicked.connect(callback) # 클릭 콜백 연결


    """
    엑셀 파일 열기
    """
    def openExcelFile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog

        filename, _ = QFileDialog.getOpenFileName(self, 
            '엑셀 파일 열기', './', '엑셀 파일 (*.xlsx)', options=options) # 엑셀 파일만 열기 허용

        success = True if filename else False
        return (success, filename)

    
    """
    새글 쓰기
    """
    def writeNewPost(self):
        success, filename = self.openExcelFile()

        # 엑셀 파일 열기 실패시 종료
        if not success:
            return

        # 새글 쓰기 프로세스 시작
        writer_proc = Process(target=writerProcess, args=(filename, self.q))
        writer_proc.start()
        

    """
    마지막 글 지우기
    """
    def deleteLastPost(self):
        filename, success = self.openExcelFile()

        # 엑셀 파일 열기 실패시 종료
        if not success:
            return

        # 마지막 글 지우기 프로세스 시작
        deleter_proc = Process(target=deleterProcess, args=(filename, self.q))
        deleter_proc.start()


    """
    다른 글 바꾸기
    """
    def modifyOtherPost(self):
        filename, success = self.openExcelFile()

        # 엑셀 파일 열기 실패시 종료
        if not success:
            return

        # 다른 글 바꾸기 프로세스 시작
        modifier_proc = Process(target=modifierProcess, args=(filename, self.q))
        modifier_proc.start()
        

"""
다른 프로세스로부터 받은 메시지 처리 스레드
"""
class UIThread(QThread):
    popup_message = pyqtSignal(str) # 팝업 메시지 시그널 생성

    def __init__(self, q):
        super().__init__()
        self.q = q


    """
    메시지 처리 루프
    """
    def run(self):
        while True:
            msg = self.q.get()

            if msg == JOBS_DONE:
                break

            self.popup_message.emit(msg) # 메시지 팝업


"""
새글 쓰기 프로세스
"""
def writerProcess(excel_filename, q):
    book = load_workbook(excel_filename, read_only=True) # 엑셀 파일
    sheet = book.worksheets[0] # 첫번째 시트

    """
    시트 포맷
        A   B   C        D        E         F         G     H     I     J       K
    1 [id][pw][subject][content][content2][content3][imgs][tags][hour][minute][is_open (yes or no)]
    2 [my_id][my_pw][my_subject]...
    3 [other_id][other_pw][other_subject]...
    4 ...
    """
    rows = sheet.iter_rows(min_row=2) # 첫번째 행을 제외한 나머지 행

    # 크롬 시크릿 창으로 열기 설정
    chrome_opts = webdriver.ChromeOptions()
    chrome_opts.add_argument('--incognito')

    driver = webdriver.Chrome('./program/chromedriver.exe', chrome_options=chrome_opts) # 크롬 열기

    # 크롬 열기 실패시 종료
    if not driver:
        q.put('크롬을 열 수 없습니다.')
        return

    prev_my_id = '' # 이전 행 아이디
    for row in rows:
        my_id = row[0].value
        my_pw = row[1].value
        my_subject = row[2].value
        my_content = row[3].value
        my_content2 = row[4].value
        my_content3 = row[5].value
        my_imgs = row[6].value
        my_tags = row[7].value
        my_hour = row[8].value
        my_minute = row[9].value
        my_is_open = row[10].value

        # 아이디가 다르면 새로 로그인하기
        if prev_my_id != my_id:
            prev_my_id = my_id # 다음 루프 확인용으로 아이디 저장하기

            succ = naverLogout(driver, q) # 로그아웃 한번 하기
            if not succ:
                return

            succ = naverLogin(driver, q, my_id, my_pw) # 로그인 하기
            if not succ:
                return

        writeNewPost(driver, q, my_subject, my_content, my_content2, my_content3, my_tags, my_imgs, my_hour, my_minute, my_is_open)
            

"""
네이버 로그아웃하기
"""
def naverLogout(driver, q):
    driver.get('https://nid.naver.com/nidlogin.logout') # 네이버 로그아웃

    try:
        elem_id_fld = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="id"]'))) # 아이디 입력 필드

    except TimeoutException:
        q.put('타임아웃: 네이버 로그아웃')
        return False

    return True


"""
네이버 로그인하기
"""
def naverLogin(driver, q, my_id, my_pw):
    driver.get('https://nid.naver.com/nidlogin.login') # 네이버 로그인 창 열기

    try:
        elem_id_fld = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="id"]'))) # 아이디 입력 필드

    except TimeoutException:
        q.put('타임아웃: 네이버 로그인')
        return False

    elem_id_fld.send_keys(my_id) # 아이디 쓰기    
    driver.find_element_by_xpath('//*[@id="pw"]').send_keys(my_pw) # 비밀번호 쓰기    
    driver.find_element_by_xpath('//*[@id="frmNIDLogin"]/fieldset/input').click() # 로그인 버튼 누르기

    try:
        elem_logout_btn = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="btn_logout"]/span'))) # 로그아웃 버튼 있는지 확인

    # 로그아웃 버튼이 없다면
    except TimeoutException:
        try:
            elem_captcha_fld = WebDriverWait(driver, timeout=3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="chptcha"]'))) # 캡챠 입력 필드가 있는지 확인

        # TODO: 캡챠 익셉션 핸들링            
        except TimeoutException:
            return True

        elem_captcha_img = driver.find_element_by_xpath('//*[@id="captchaimg"]') # 캡챠 이미지 원소

        captcha_img_url = elem_captcha_img.get_attribute('src') # 캡챠 이미지 주소

        ANTICAPTCHA_KEY = '5330b0f08fe52776ce6caaf56321539d' # 안티캡챠 키

        # 캡챠 정답
        captcha_answer = ImageToTextTask.ImageToTextTask(
            anticaptcha_key=ANTICAPTCHA_KEY, save_format='const').captcha_handler(captcha_link=captcha_img_url)

        captcha_text = captcha_answer['solution']['text'] # 캡챠 정답 문자열

        driver.find_element_by_xpath('//*[@id="pw"]').send_keys(my_pw) # 비밀번호 쓰기
        elem_captcha_fld.send_keys(captcha_text) # 캡챠 정답 문자열 쓰기
        driver.find_element_by_xpath('//*[@id="login_submit"]').click() # 로그인 버튼 누르기

        return True


"""
새글 쓰기
"""
def writeNewPost(driver, q, subject, content, content2, content3, tags, imgs, hour, minute, is_open):
    driver.get('https://blog.editor.naver.com/editor') # 에디터창 열기

    try:
        elem_popup = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[6]/div/div/div[2]/a'))) # 새글쓰기 팝업창 있는지 확인

    except TimeoutException:
        pass

    finally:
        try:
            driver.execute_script('arguments[0].click();', elem_popup) # 새글쓰기 팝업창이 있다면 닫기

        except:
            pass

    while True:
        pass

    # 제목 쓰기 //*[@id="documentTitle_8761958271541983606916"]/div[2]/div/div[4]/div/div/textarea
    #driver.find_element_by_xpath('//*[@id="documentTitle_1715842561541981314717"]/div[2]/div/div[4]/div/div/textarea').send_keys(subject)

    # 본문 쓰기
    #driver.find_element_by_xpath('//*[@id="paragraph_1436506121541981314727"]/div[1]/div/div[4]/div/div/div/div').send_keys(content)

    # 태그 쓰기
    #driver.find_element_by_xpath('//*[@id="se_canvas_body"]/div[3]/div/div/div/div/span/ul/li/input').send_keys(tags)

    # 그림 아이콘 누르기
    #driver.find_element_by_xpath('//*[@id="se_side_comp_list"]/li[2]/button').click()



def deleterProcess(excel_filename, q):
    q.put('실패')


def modifierProcess(excel_filename, q):
    q.put(JOBS_DONE)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())
