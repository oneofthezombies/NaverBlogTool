"""
내 함수
"""


# 시스템
import time, os, platform
from random import randrange

# 엑셀
from pandas import read_excel

# 파일 선택 창
import tkinter
from tkinter.filedialog import askopenfilename

# 크롬 브라우저 제어
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# 안티캡챠
from python3_anticaptcha import ImageToTextTask
from python3_anticaptcha.errors import IdGetError

# 클립보드
import pyperclip 

# 마우스/키보드 컨트롤
import pyautogui 


def has_keyword(text, keywords):
    return any(keyword in text for keyword in keywords)


def load_keywords():
    df = read_excel('키워드.xlsx')
    res = []
    for i in range(len(df)):
        res.append(str(df.loc[i]['키워드 입력(↓)']))
    return res


def load_my_file():
    root = tkinter.Tk()
    root.withdraw()
    root.update()
    filename = askopenfilename(title='아이디,비밀번호,제목,본문1,본문2,본문3,태그 (컴마로 분리, 예. tag0, tag1, tag2),이미지들 (run.bat이 있는 폴더 내에),예약 유무,년,월,일,시,분,공개 유무',
                               filetypes=[('Excel file', '*.xl*')])
    root.destroy()
    return read_excel(filename)


def get_count_tab():

    # os 별로 파일 업로드 창이 다르다
    os_version = platform.release()
    print('OS 버전: {}'.format(os_version))

    count_tab = 5
    if '7' in os_version:
        count_tab = 4
    elif '8' in os_version:
        count_tab = 4

    return count_tab


def open_chrome():

    # 크롬 시크릿 창으로 열기 설정
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--incognito')

    browser = webdriver.Chrome(executable_path='./chromedriver.exe', chrome_options=chrome_options)
    return browser


class Browser:
    def __init__(self, delay, anticaptcha_key, is_manual):
        self.browser = open_chrome()
        self.count_tab = get_count_tab()

        self.delay = delay
        self.is_manual = is_manual
        self.anticaptcha_key = anticaptcha_key


    def find_element(self, xpath):
        try:
            elem = WebDriverWait(self.browser, timeout=self.delay) \
                .until(EC.presence_of_element_located((By.XPATH, xpath)))

        except TimeoutException:
            return None

        return elem


    def find_elements(self, xpath):
        try:
            WebDriverWait(self.browser, timeout=self.delay) \
                .until(EC.presence_of_element_located((By.XPATH, xpath)))

            elems = self.browser.find_elements_by_xpath(xpath)

        except TimeoutException:
            return None

        return elems 


    def naver_logout(self):

        # 네이버 로그아웃 
        self.browser.get('https://nid.naver.com/nidlogin.logout') 

        # 로그아웃 메시지
        elem_logout_msg = self.find_element('//*[@id="content"]/div[1]/p')

        if elem_logout_msg is None:
            print('타임아웃: 로그아웃 메시지')
            exit()
    

    def naver_login(self, my_id, my_pw):

        # 네이버 메인 창 경유
        self.browser.get('https://www.naver.com')

        # 로그인 버튼
        elem_login = self.find_element('//*[@id="account"]/div/a/i')
        elem_login.click()

        # 아이디 입력 필드
        elem_id = self.find_element('//*[@id="id"]')

        if elem_id is None:
            print('타임아웃: 아이디 입력 필드')
            exit()

        # 아이디 쓰기
        elem_id.send_keys(my_id) 

        # 비밀번호 필드
        elem_pw = self.find_element('//*[@id="pw"]')    

        # 비밀번호 쓰기
        elem_pw.send_keys(my_pw)
        
        if self.is_manual:
            # 비밀번호 클립보드에 복사
            pyperclip.copy(my_pw)

            root = tkinter.Tk()
            tkinter.Button(root, text="로그인 완료", command=root.destroy).pack()
            root.mainloop()

            # 네이버 메인
            elem_main = self.find_element('//*[@id="PM_ID_ct"]')

            if elem_main:
                print('로그인 성공')
            else:
                self.naver_login(my_id, my_pw)

        else:
            elem_pw.submit()

            # 네이버 메인
            elem_main = self.find_element('//*[@id="PM_ID_ct"]')

            if elem_main:
                print('로그인 성공')

            else:
                # 캡챠 창
                elem_capt = self.find_element('//*[@id="chptcha"]')        

                if elem_capt:

                    # 안티캡챠 키 재발급 받아야 함
                    self.anti_captcha(my_id, my_pw)

                else:
                    print('로그인 재시도')
                    self.naver_login(my_id, my_pw)


    def anti_captcha(self, my_id, my_pw):
        print('안티캡챠 실행')

        # 캡챠 이미지 원소
        elem_captcha_img = self.find_element('//*[@id="captchaimg"]')

        # 캡챠 이미지 주소
        captcha_img_url = elem_captcha_img.get_attribute('src') 

        try:

            # 캡챠 정답
            captcha_answer = ImageToTextTask.ImageToTextTask(anticaptcha_key=self.anticaptcha_key, save_format='const') \
                .captcha_handler(captcha_link=captcha_img_url)

        except IdGetError:
            print('안티캡챠 밸런스 모두 사용함! 결제해야 함!')
            exit()
        
        if not 'solution' in captcha_answer:
            print('로그인 재시도')
            self.naver_login(my_id, my_pw)
            return
        
        # 캡챠 정답
        captcha_text = captcha_answer['solution']['text'] 

        # 비밀번호 쓰기
        elem_pw = self.find_element('//*[@id="pw"]') 
        elem_pw.send_keys(my_pw)

        # 캡챠 정답 쓰기
        elem_captcha_fld = self.find_element('//*[@id="chptcha"]')
        elem_captcha_fld.send_keys(captcha_text)

        elem_pw.submit()

        # 네이버 메인
        elem_main = self.find_element('//*[@id="PM_ID_ct"]')

        if elem_main is None:
            print('로그인 재시도')
            self.naver_login(my_id, my_pw)


    def click_noexcept(self, element):
        try:
            self.click(element)
        except:
            pass

    
    def click(self, element):
        self.browser.execute_script('arguments[0].click();', element)


    def open_new_tab(self, element):
        ActionChains(self.browser) \
            .key_down(Keys.CONTROL) \
            .click(element) \
            .key_up(Keys.CONTROL) \
            .perform()


    def work_write(self, subject, content1, content2, content3, tags, images, is_reserved, year, month, day, hour, minute, is_open):

        # 본문 프레임
        elem_frame = self.find_element('//*[@id="se_canvas_frame"]')

        if elem_frame is None:
            print('타임아웃: 본문 프레임')
            exit()

        # 이미지 올리기 버튼
        elem_upload_img = self.find_element('//*[@id="se_side_comp_list"]/li[2]/button')

        if elem_upload_img is None:
            print('타임아웃: 이미지 올리기 버튼')
            exit()

        # 본문 프레임으로 전환
        self.browser.switch_to.frame(elem_frame)
        time.sleep(1)

        # 제목 쓰기
        elem_subject = self.find_element('//textarea[@class="se_editable se_textarea"]')
        pyperclip.copy(subject)
        elem_subject.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        # 본문 1, 2 쓰기
        elem_content = self.find_element('//*[@contenteditable="true"]')
        pyperclip.copy('\n'.join([content1, content2]))
        elem_content.send_keys(Keys.CONTROL, 'v')
        time.sleep(self.delay)
        
        # 에디터 프레임으로 전환
        self.browser.switch_to.default_content()
        time.sleep(1)

        # 이미지 업로드
        self.upload_image(elem_upload_img, images)

        # 본문 프레임으로 전환
        self.browser.switch_to.frame(elem_frame)     
        time.sleep(1)

        # 본문 3 쓰기
        elem_content = self.find_element('//div[@contenteditable="true"]')
        pyperclip.copy(content3)
        elem_content.send_keys(Keys.CONTROL, 'v')
        time.sleep(self.delay)

        # 태그 쓰기
        if tags != '':  
            elem_tag = self.find_element('//li/input[@type="text"]')
            elem_tag.send_keys(tags)
            elem_tag.send_keys(',')
            time.sleep(1)

        # 에디터 프레임으로 전환
        self.browser.switch_to.default_content()    
        time.sleep(1)   

        # 중앙 정렬 버튼 누르기
        elem_align_center = self.find_element('//a[@class="btn_alignCenter __se_align_btn"]')
        self.click(elem_align_center)
        time.sleep(1)

        # 상단 발행 버튼 누르기
        elem_publish = self.find_element('//a[@id="se_top_publish_btn"]')
        self.click(elem_publish)

        # 발행 버튼
        elem_final_publish = self.find_element('//*[@id="se_top_publish_setting_layer"]/div[1]/div[2]/div[3]/button[@class="btn_publish"]')

        if is_reserved == '예약 발행':
            print('년월일 예약 발행은 미구현 상태입니다. 시,분만 예약됩니다.')

            # 예약발행 버튼 클릭
            elem_appoint = self.find_element('//*[@id="se_top_publish_setting_layer"]//*[@class="btn_appointment"]')
            self.click(elem_appoint)
            time.sleep(self.delay)

            # # 시간 선택 박스
            # succ, _ = find_element_wait_for(browser, )

            # if not succ:
            #     message_queue.put('타임아웃: 시간 선택 박스')
            #     return False

            # # 년월일 선택 박스
            # elem_year_month_day = browser.find_element_by_xpath('//input[@type="text" and @class="date hasDatepicker"]')
            # time.sleep(2)

            # # 읽기 전용 속성 제거
            # browser.execute_script('arguments[0].removeAttribute("readonly")', elem_year_month_day) 
            # time.sleep(2)
            
            # # 년월일
            # year_month_day = '{0}.{1:02d}.{2:02d}'.format(year, int(month), int(day))
            # print(year_month_day)

            # # 년월일 입력
            # browser.execute_script("arguments[0].setAttribute('value', arguments[1])", elem_year_month_day, year_month_day)
            # time.sleep(2)

            # # # 읽기 전용 속성 추가
            # # browser.execute_script('arguments[0].createAttribute("readonly")', elem_year_month_day) 
            # # time.sleep(2)

            # # # 읽기 전용 속성 true
            # # browser.execute_script('arguments[0].setAttribute("readonly", arguments[1])', elem_year_month_day, 'true') 
            # # time.sleep(2)

            # 시간 선택 버튼
            elem_hour = self.find_element('//select[@class="hour"]')

            # 시간 입력
            elem_hour_opt = self.find_element('//select[@class="hour"]/option[text()="{:02d}"]'.format(int(hour)))
            elem_hour_opt.click()
            time.sleep(1)

            # 분 입력
            elem_minute_opt = self.find_element('//select[@class="minutes"]/option[text()="{:02d}"]'.format(int(minute)))
            elem_minute_opt.click()
            time.sleep(1)

        if is_open == '공개':

            # 공개로 설정
            elem_public = self.find_element('//label[@for="lv_public_1"]')
            elem_public.click()
            time.sleep(1)

        else:

            # 비공개로 설정
            elem_private = self.find_element('//label[@for="lv_public_4"]')
            elem_private.click()
            time.sleep(1)

        # 발행하기
        self.click(elem_final_publish)
        time.sleep(self.delay)
        

    def write_new_post(self, subject, content1, content2, content3, tags, images, is_reserved, year, month, day, hour, minute, is_open):

        # 에디터창 열기
        self.browser.get('https://blog.editor.naver.com/editor') 

        # 발행 버튼
        elem_publish = self.find_element('//*[@id="se_top_publish_btn"]')

        # 팝업창 버튼
        elem_popup = self.find_element('/html/body/div[6]/div/div/div[2]/a')

        # 팝업창 닫기
        self.click_noexcept(elem_popup)

        # 작성
        self.work_write(subject, content1, content2, content3, tags, images, is_reserved, year, month, day, hour, minute, is_open)

        if is_reserved == '예약 발행':
            print('{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}'.format(subject, is_reserved, year, month, day, hour, minute, is_open))
        else:
            print('{0}, {1}, {2}'.format(subject, is_reserved, is_open))


    def upload_image(self, element_image, image_paths):

        # 현재 디렉터리
        cwd = os.getcwd() + '/'

        # 이미지가 여러장일 경우 분리
        image_paths = [image_path for image_path in image_paths.splitlines() if image_path is not ''] 

        # 이미지의 경로와 파일이름 분리
        path_filenames = [[]]
        path_filenames.pop()

        for image_path in image_paths:
            index = image_path.rfind('/')

            if index == -1:
                path_filenames.append([None, image_path])

            else:
                path_filenames.append([image_path[:index+1], image_path[index+1:]])

        # 이미지 업로드
        for pf in path_filenames:

            # 이미지 업로드 버튼 클릭
            element_image.click() 
            time.sleep(self.delay)

            for _ in range(self.count_tab):
                pyautogui.press('tab')

            pyautogui.press('enter')

            # 경로 입력
            if pf[0] != None:
                pyautogui.typewrite(cwd + pf[0])

            else:
                pyautogui.typewrite(cwd)

            pyautogui.press('enter')
            time.sleep(1)

            for _ in range(5):
                pyautogui.press('tab')

            # 파일 입력
            pyautogui.typewrite(pf[1])

            # 엔터 누르기
            pyautogui.press('enter')
            time.sleep(self.delay)    


    def work_modify(self, elem_post, subject, content1, content2, content3, tags, images):

        # 새 창으로 열기
        self.open_new_tab(elem_post)
        time.sleep(self.delay)

        # 수정 탭으로 이동
        self.browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.TAB)
        self.browser.switch_to.window(self.browser.window_handles[1])
        time.sleep(1)

        # 수정 버튼 클릭
        elem_modify = self.find_element('//*[@id="printPost1"]/tbody/tr/td[2]/div[3]/div[2]/div[2]/a[1]')
        elem_modify.click()

        # 발행 버튼
        elem_publish = self.find_element('//*[@id="se_top_publish_btn"]')

        # 팝업창 버튼
        elem_popup = self.find_element('/html/body/div[6]/div/div/div[2]/a')

        # 팝업창 닫기
        self.click_noexcept(elem_popup)

        # 본문 프레임으로 전환
        elem_frame = self.find_element('//*[@id="se_canvas_frame"]')
        self.browser.switch_to.frame(elem_frame)
        time.sleep(1)

        # 본문 지우기
        pyautogui.keyDown('ctrl')
        for _ in range(4):
            pyautogui.press('a')
        pyautogui.keyUp('ctrl')
        pyautogui.press('backspace')
        time.sleep(1)

        # 태그 지우기
        elem_tags = self.find_elements('//*[@id="se_canvas_body"]/div[3]/div/div/div/div/span/ul/li')

        try:
            for _ in elem_tags:
                elem_tag_fld = self.find_element('//li/input[@type="text"]')
                elem_tag_fld.send_keys(Keys.BACK_SPACE)
        except:
            pass

        # 제목 지우기
        elem_subject = self.find_element('//textarea[@class="se_editable se_textarea"]')
        elem_subject.click()
        elem_subject.send_keys(Keys.CONTROL + 'a')
        elem_subject.send_keys(Keys.BACK_SPACE)
        time.sleep(1)

        # 본문 에디터창 찾기
        while True:
            pyautogui.press('down')
            time.sleep(1)
            elem_content = self.find_element('//div[@contenteditable="true"]')
            if elem_content:
                break
            else:
                print('본문 에디터창 찾는 중')

        # 에디터 프레임으로 전환
        self.browser.switch_to.default_content()
        time.sleep(1)

        # 글 쓰기
        self.work_write(subject, content1, content2, content3, tags, images, '예약 안함', '', '', '', '', '', '공개')

        # 수정 탭 닫기
        pyautogui.keyDown('ctrl')
        pyautogui.press('w')
        pyautogui.keyUp('ctrl')

        # 블로그 탭으로
        self.browser.switch_to.window(self.browser.window_handles[0])
        time.sleep(1)


    def modify_other_post(self, keywords, my_id, subject, content1, content2, content3, tags, images, is_open):

        # 블로그 메인
        self.browser.get('https://blog.naver.com/PostList.nhn?blogId={}'.format(my_id)) 

        # 전체보기 버튼 클릭
        elem_total = self.find_element('//*[@id="category0"]')
        elem_total.click()

        # 목록 열기 버튼
        elem_list = self.find_element('//*[@id="toplistSpanBlind"]')
        elem_list.click()
        
        # 맨 뒤 페이지로 가기 (제일 과거로)
        while True:

            # 페이지 링크들
            elem_page_links = self.find_elements('//div[@id="toplistWrapper"]//div[@class="blog2_paginate"]//a[@href="#"]')

            # '다음' 링크가 있다면 클릭
            elem_page_links = [epl for epl in elem_page_links if epl.text == '다음']
            if elem_page_links:
                elem_page_links[0].click()
                time.sleep(1)

            # 없다면 마지막 페이지 클릭
            else:
                # 페이지 링크들
                elem_page_links = self.find_elements('//div[@id="toplistWrapper"]//div[@class="blog2_paginate"]//a[@href="#"]')
                elem_page_texts = [epl.text for epl in elem_page_links if epl.text != '다음']
                if elem_page_texts:
                    elem_page_texts.reverse()
                    elem_page = self.find_element('//div[@id="toplistWrapper"]//div[@class="blog2_paginate"]//a[@href="#" and contains(text(), "{}")]'.format(elem_page_texts[0]))
                    elem_page.click()
                    time.sleep(1)
                    break

        # 맨 뒤부터 작업하면서 이전 페이지 그룹으로 오기
        while True:

            # 페이지 그룹에서의 마지막 페이지
            # 현재 페이지 포스트 목록 
            elem_posts = self.find_elements('//*[@id="listTopForm"]/table/tbody/tr/td/div/span/a') 

            # 현재 페이지 포스트 목록 뒤집기
            elem_posts.reverse()

            for elem_post in elem_posts:

                # 키워드가 하나라도 없다면
                if not has_keyword(elem_post.text, keywords):

                    # 수정 작업
                    self.work_modify(elem_post, subject, content1, content2, content3, tags, images)

                    # 수정 완료
                    print('포스트 수정 완료')
                    return

            # 페이지 링크들
            elem_page_links = self.find_elements('//div[@id="toplistWrapper"]//div[@class="blog2_paginate"]//a[@href="#"]')

            # '이전'이 아닌 그리고 다음'이 아닌 모든 페이지 링크의 텍스트들
            elem_page_texts = [elem_page_link.text for elem_page_link in elem_page_links if elem_page_link.text != '이전' and elem_page_link.text != '다음']

            # 페이지 링크 텍스트 목록 뒤집기
            elem_page_texts.reverse()

            for text in elem_page_texts:

                # 페이지 클릭
                elem_page = self.find_element('//div[@id="toplistWrapper"]//div[@class="blog2_paginate"]//a[@href="#" and contains(text(), "{}")]'.format(text))
                elem_page.click()
                time.sleep(1)

                # 현재 페이지 포스트 목록 
                elem_posts = self.find_elements('//*[@id="listTopForm"]/table/tbody/tr/td/div/span/a') 

                # 현재 페이지 포스트 목록 뒤집기
                elem_posts.reverse()

                for elem_post in elem_posts:

                    # 키워드가 하나라도 없다면
                    if not has_keyword(elem_post.text, keywords):

                        self.work_modify(elem_post, subject, content1, content2, content3, tags, images)

                        # 수정 완료
                        print('포스트 수정 완료')
                        return

            # '이전' 링크가 있다면 클릭
            elem_page_links = self.find_elements('//div[@id="toplistWrapper"]//div[@class="blog2_paginate"]//a[@href="#"]')
            elem_page_links = [epl for epl in elem_page_links if epl.text == '이전']
            if elem_page_links:
                elem_page_links[0].click()
                time.sleep(1)
            
            else:
                print('바꿀 포스트가 없습니다.')


    def delete_last_post(self):
        pass

