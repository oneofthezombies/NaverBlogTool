"""
네이버 블로그 도구
쓰기_지우기_수정하기
"""


# 내 함수
from my_function import load_keywords, load_my_file, Browser




# 딜레이
MIN_DELAY = 5

# 수동 로그인
IS_MANUAL = True # True or False

# 안티캡챠 키
ANTICAPTCHA_KEY = '3a8795821736b65d999a30ee1f210d2f' 




# 프로그램 실행
if __name__ == '__main__':

    # 키워드 엑셀 파일 로드
    keywords = load_keywords()

    # 작성 정보 엑셀 파일 로드
    my_file = load_my_file()

    # 크롬 브라우저
    browser = Browser(MIN_DELAY, ANTICAPTCHA_KEY, IS_MANUAL)

    for i in range(0, len(my_file), 3):

        # 네이버 로그아웃
        browser.naver_logout()

        # 네이버 로그인
        my_id = my_file.loc[i]['아이디']
        my_pw = my_file.loc[i]['비밀번호']

        browser.naver_login(my_id, my_pw)

        # 새 글 작성
        subject = my_file.loc[i]['제목']
        content1 = my_file.loc[i]['본문1']
        content2 = my_file.loc[i]['본문2']
        content3 = my_file.loc[i]['본문3']
        tags = my_file.loc[i]['태그 (컴마로 분리, 예. tag0, tag1, tag2)']
        images = my_file.loc[i]['이미지들 (run.bat이 있는 폴더 내에)']
        is_reserved = my_file.loc[i]['예약 유무']
        year = my_file.loc[i]['년']
        month = my_file.loc[i]['월']
        day = my_file.loc[i]['일']
        hour = my_file.loc[i]['시']
        minute = my_file.loc[i]['분']
        is_open = my_file.loc[i]['공개 유무']

        while True:
            browser.write_new_post(subject, content1, content2, content3, tags, images, is_reserved, year, month, day, hour, minute, is_open)
        
        # 다른 글 수정
        subject = my_file.loc[i + 1]['제목']
        content1 = my_file.loc[i + 1]['본문1']
        content2 = my_file.loc[i + 1]['본문2']
        content3 = my_file.loc[i + 1]['본문3']
        tags = my_file.loc[i + 1]['태그 (컴마로 분리, 예. tag0, tag1, tag2)']
        images = my_file.loc[i + 1]['이미지들 (run.bat이 있는 폴더 내에)']

        browser.modify_other_post(keywords, my_id, subject, content1, content2, content3, tags, images, is_open)

        # 마지막글 삭제
        browser.delete_last_post()
        print('test: 한 사이클')

    print('test: 프로그램 완료')