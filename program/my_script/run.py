"""
네이버 블로그 도구
"""

# 시스템
import sys, os, time
from multiprocessing import Process, Queue as MPQueue

# 크롬 브라우저 제어
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

from python3_anticaptcha import ImageToTextTask # 안티캡챠

from openpyxl import load_workbook # 엑셀 

# 데스크탑 앱 모듈
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, QThread
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, 
    QDesktopWidget, QPushButton, QMessageBox)

import pyautogui # 마우스/키보드 컨트롤

import pyperclip # 클립보드


JOBS_DONE = 'jobs_done' # 멀티프로세스 메시지 큐 종료 플래그


"""
데스크탑 앱
"""
class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle('네이버 블로그 도구') # 앱 이름        
        self.setGeometry(0, 0, 350, 190) # 앱 크기

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
    1 [id][pw][subject][content1][content2][content3][tags][imgs][year][month][day][hour][minute][is_open]
    2 [my_id][my_pw][my_subject]...
    3 [other_id][other_pw][other_subject]...
    4 ...
    """
    rows = sheet.iter_rows(min_row=2) # 첫번째 행을 제외한 나머지 행

    # 크롬 시크릿 창으로 열기 설정
    chrome_opts = webdriver.ChromeOptions()
    chrome_opts.add_argument('--incognito')

    driver = webdriver.Chrome(executable_path='./program/chromedriver.exe', chrome_options=chrome_opts) # 크롬 열기

    # 크롬 열기 실패시 종료
    if not driver:
        q.put('크롬을 열 수 없습니다.')
        return

    prev_id = '!' # 이전 행 아이디
    for row in rows:
        id_ = row[0].value
        if id_ == None: # 아이디가 비어있으면 넘어간다
            continue

        pw = row[1].value
        subject = row[2].value
        content = row[3].value
        content2 = row[4].value
        content3 = row[5].value
        tags = row[6].value
        imgs = row[7].value
        year = row[8].value
        month = row[9].value
        day = row[10].value
        hour = row[11].value
        minute = row[12].value
        is_open = row[13].value

        # 아이디가 다르면 새로 로그인하기
        if prev_id != id_:
            prev_id = id_ # 다음 루프 확인용으로 아이디 저장하기

            succ = naverLogout(driver, q) # 로그아웃 한번 하기
            if not succ:
                return

            print('로그인 시도')
            succ = naverLogin(driver, q, id_, pw) # 로그인 하기
            if not succ:
                return
            
            print('로그인 성공')

        writeNewPost(driver, q, id_, subject, content, content2, content3, tags, imgs, year, month, day, hour, minute, is_open)

    print('새글쓰기 프로세스 종료')


"""
네이버 로그아웃하기
"""
def naverLogout(driver, q):
    driver.get('https://nid.naver.com/nidlogin.logout') # 네이버 로그아웃

    try:
        elem_logout_msg = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[1]/p'))) # 로그아웃 메시지

    except TimeoutException:
        q.put('타임아웃: 네이버 로그아웃')
        return False

    return True


"""
네이버 로그인하기
"""
def naverLogin(driver, q, id_, pw):
    driver.get('https://nid.naver.com/nidlogin.login') # 네이버 로그인 창 열기

    try:
        elem_id = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="id"]'))) # 아이디 입력 필드

    except TimeoutException:
        q.put('타임아웃: 네이버 로그인')
        return False

    elem_id.send_keys(id_) # 아이디 쓰기
    elem_pw = driver.find_element_by_xpath('//*[@id="pw"]') # 비밀번호 쓰기   
    elem_pw.send_keys(pw)
    elem_pw.submit()

    try:
        elem_main = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="PM_ID_ct"]'))) # 네이버 메인

    # 네이버 메인이 아니라면
    except TimeoutException:
        print('로그인 실패')
        try:
            elem_captcha = WebDriverWait(driver, timeout=3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="chptcha"]'))) # 캡챠 입력 필드가 있는지 확인

        except TimeoutException:
            print('로그인 재시도')
            return naverLogin(driver, q, id_, pw)

        return antiCaptcha(driver, q, id_, pw)

    return True


"""
안티캡챠
"""
def antiCaptcha(driver, q, id_, pw):
    print('안티캡챠 실행')

    elem_captcha_img = driver.find_element_by_xpath('//*[@id="captchaimg"]') # 캡챠 이미지 원소

    captcha_img_url = elem_captcha_img.get_attribute('src') # 캡챠 이미지 주소

    ANTICAPTCHA_KEY = '5330b0f08fe52776ce6caaf56321539d' # 안티캡챠 키

    # 캡챠 정답
    captcha_answer = ImageToTextTask.ImageToTextTask(
        anticaptcha_key=ANTICAPTCHA_KEY, save_format='const').captcha_handler(captcha_link=captcha_img_url)

    if not 'solution' in captcha_answer:
        print('로그인 재시도')
        return naverLogin(driver, q, id_, pw)
    
    captcha_text = captcha_answer['solution']['text'] # 캡챠 정답 문자열

    elem_pw = driver.find_element_by_xpath('//*[@id="pw"]') # 비밀번호 쓰기
    elem_pw.send_keys(pw)
    driver.find_element_by_xpath('//*[@id="chptcha"]').send_keys(captcha_text) # 캡챠 정답 문자열 쓰기
    elem_pw.submit()

    try:
        elem_main = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="PM_ID_ct"]'))) # 네이버 메인

    except TimeoutException:
        print('로그인 재시도')
        return naverLogin(driver, q, id_, pw)

    return True


"""
새글 쓰기
"""
def writeNewPost(driver, q, id_, subject, content, content2, content3, tags, imgs, year, month, day, hour, minute, is_open):
    driver.get('https://blog.editor.naver.com/editor') # 에디터창 열기
    
    try:
        elem_popup = WebDriverWait(driver, timeout=6).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[6]/div/div/div[2]/a'))) # 새글쓰기 팝업창 있는지 확인
    except:
        pass

    try:
        driver.execute_script('arguments[0].click();', elem_popup) # 새글쓰기 팝업창이 있다면 닫기
    except:
        pass

    try:
        elem_canvas_frm = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="se_canvas_frame"]')))

    except TimeoutException:
        q.put('타임아웃: 캔버스 프레임')
        return False

    try:
        elem_img = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="se_side_comp_list"]/li[2]/button'))) # 이미지 올리기 버튼

    except TimeoutException:
        q.put('타임아웃: 이미지 올리기 버튼')
        return False

    cwd = os.getcwd() # 현재 디렉터리
    cwd += '/'

    imgs = [img for img in imgs.splitlines() if img != ''] # 이미지별 분리하기

    # 경로와 파일이름 분리
    path_filenames = [[]]
    path_filenames.pop()
    for img in imgs:
        i = img.rfind('/')
        if i == -1:
            path_filenames.append([None, img])
        else:
            path_filenames.append([img[:i+1], img[i+1:]])

    # 이미지 업로드
    for pf in path_filenames:
        elem_img.click() # 이미지 업로드 버튼 클릭
        time.sleep(2) # 파일 열기창 기다리기

        # 경로 입력
        for i in range(5):
            pyautogui.press('tab')
        pyautogui.press('enter')
        if pf[0] != None:
            pyautogui.typewrite(cwd + pf[0])
        else:
            pyautogui.typewrite(cwd)
        pyautogui.press('enter')

        # 파일 입력
        for i in range(5):
            pyautogui.press('tab')
        pyautogui.typewrite(pf[1])
        pyautogui.press('enter')
        time.sleep(4) # 파일 업로드 기다리기

    driver.switch_to.frame(elem_canvas_frm) # 본문 프레임
    
    pyperclip.copy(subject)
    driver.find_element_by_xpath('//textarea').send_keys(Keys.CONTROL, 'v') # 제목
    time.sleep(2)

    pyperclip.copy('\n'.join([content, content2, content3]))
    driver.find_element_by_xpath('//div[@contenteditable="true"]').send_keys(Keys.CONTROL, 'v') # 본문
    time.sleep(2)

    pyperclip.copy(tags)    
    driver.find_element_by_xpath('//li/input[@type="text"]').send_keys(Keys.CONTROL, 'v') # 태그 #,(컴마)로 구분되어야 함
    time.sleep(2)

    driver.switch_to.default_content()

    elem_align_center = driver.find_element_by_xpath('//a[@class="btn_alignCenter __se_align_btn"]')
    driver.execute_script('arguments[0].click();', elem_align_center)

    elem_publish = driver.find_element_by_xpath('//a[@id="se_top_publish_btn"]')
    driver.execute_script('arguments[0].click();', elem_publish)

    try:
        elem_appointment_btn = WebDriverWait(driver, timeout=3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="se_top_publish_setting_layer"]/div[1]/div[2]/div[3]/button[2]')))

    except TimeoutException:
        q.put('타임아웃: 예약발행 버튼 유무')
        return False

    ymd = '예약 안함'
    if year != '예약 안함':
        driver.execute_script('arguments[0].click();', elem_appointment_btn) # 예약발행 버튼 클릭

        try:
            elem_hour = WebDriverWait(driver, timeout=3).until(
                EC.presence_of_element_located((By.XPATH, '//select[@class="hour"]')))
        except TimeoutException:
            q.put('타임아웃: 시간 예약 버튼')
            return False

        elem_ymd = driver.find_element_by_xpath('//input[@type="text" and @class="date hasDatepicker" and @readonly="true"]')
        driver.execute_script('arguments[0].removeAttribute("readonly");', elem_ymd) # 읽기 전용 속성 제거
        ymd = '{0}.{1:02d}.{2:02d}'.format(year, int(month), int(day))
        driver.execute_script("arguments[0].setAttribute('value', arguments[1])", elem_ymd, ymd) # 년월일

        driver.find_element_by_xpath('//select[@class="hour"]/option[text()="{:02d}"]'.format(int(hour))).click() # 시간 입력
        driver.find_element_by_xpath('//select[@class="minutes"]/option[text()="{:02d}"]'.format(minute)).click() # 분 입력

    if is_open == '공개':
        driver.find_element_by_xpath('//label[@for="lv_public_1"]').click() # 공개로 설정
    else:
        driver.find_element_by_xpath('//label[@for="lv_public_4"]').click() # 비공개로 설정

    elem_publish_confirm = driver.find_element_by_xpath('//button[@class="btn_publish"]')
    driver.execute_script('arguments[0].click();', elem_publish_confirm) # 발행하기
    
    if year == '예약 안함':
        print('{0}, {1}, {2}, {3} 게시 완료'.format(id_, subject, year, is_open))
    else:
        print('{0}, {1}, {2}, {3}:{4}, {5} 게시 완료'.format(id_, subject, ymd, hour, minute, is_open))

    time.sleep(5) # 업로드 완료 기다리기


def deleterProcess(excel_filename, q):
    q.put('실패')


def modifierProcess(excel_filename, q):
    q.put(JOBS_DONE)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())
