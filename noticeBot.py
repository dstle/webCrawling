import time
import datetime
import win32con
import win32api
import win32gui
import requests
import json
import logging
import requests 

from bs4 import BeautifulSoup
from operator import eq
from apscheduler.schedulers.background import BackgroundScheduler
from logging.handlers import TimedRotatingFileHandler

# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = 'noticebot'
idx = 89975

# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, noticeLists):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RICHEDIT50W", None)

    check = len(noticeLists)

  
    for noticeList in noticeLists:
        win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, noticeList)
        SendReturn(hwndEdit)
        botLogger = logging.getLogger()
        botLogger.debug(noticeList)
        time.sleep(3)


# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx(
        hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx(
        hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx(
        hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx(
        hwndkakao_edit2_2, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(
        hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)

# 공지사항 크롤링하기
def get_dwu_notice():
    global idx
    url = 'https://www.dongduk.ac.kr/www/contents/kor-noti.do?gotoMenuNo=kor-noti'  
    response = requests.get(url)    
    dongduk_url = 'https://www.dongduk.ac.kr/www/contents/kor-noti.do?schM=view&page=1&viewCount=10&id='

    if response.status_code != 200:
        botLogger.error(f"[get_dwu_notice] Failed to fetch notices. HTTP {response.status_code}")
        return []

    html = response.text
    soup = BeautifulSoup(html, 'html.parser')
    notices = soup.select_one('ul.board-basic')
    elements = notices.select('li > dl')

    notice_set = []
    existing_ids = set()

    for element in elements:
        id = int(element.a.get('onclick').split("'")[1])
        if id in existing_ids: # 중복 방지
            continue  
        existing_ids.add(id)
    
        title = element.a.text.strip()
        date = element.find_all('span', 'p_hide')[1].text
        notice_set.append({"id": id, "title": title, "date": date, "link": f"{dongduk_url}{id}"})

    new_notices = [el for el in notice_set if el["id"] > idx]
    if new_notices:
        new_notices.sort(key=lambda x: x["id"])
        idx = new_notices[-1]["id"]
        botLogger.info(f"[get_dwu_notice] {len(new_notices)} new notices fetched.")
        return new_notices
    
    botLogger.info("[get_dwu_notice] No new notices found.")
    return []


# # 스케줄러 job : 매 시간마다 공지사항 크롤링해서 가져오기
def job():
    p_time_ymd_hms = \
        f"{time.localtime().tm_year}-{time.localtime().tm_mon}-{time.localtime().tm_mday} / " \
        f"{time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"

    # 채팅방 열기
    open_chatroom(kakao_opentalk_name)
    noticeList = get_dwu_notice()

    # 메시지 전송, time/실검
    kakao_sendtext(kakao_opentalk_name, noticeList)


# # log 환경설정
def set_logger():
    global botLogger 
    botLogger = logging.getLogger()
    botLogger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                  "%Y-%m-%d %H:%M:%S")

    # 일주일에 한번 월요일 자정에 로그 파일 새로 생성. 최대 5개까지 파일 관리.
    rotatingHandler = TimedRotatingFileHandler(
        filename='./noticebot_log/webCrawling.log', when='W0', encoding='utf-8', backupCount=5, atTime=datetime.time(0, 0, 0))
    rotatingHandler.setLevel(logging.DEBUG)
    rotatingHandler.setFormatter(formatter)

    # 파일 이름 suffix 설정 (webCrawling.log.yyyy-mm-dd-hh-mm 형식)
    rotatingHandler.suffix = datetime.datetime.today().strftime("%Y-%m-%d-%H-%M")
    botLogger.addHandler(rotatingHandler)


def main():
    sched = BackgroundScheduler()
    sched.start()
    set_logger()

    # 15분마다 실행
    sched.add_job(job, 'interval', minutes=15)

    while True:
        botLogger = logging.getLogger()
        botLogger.debug("-------------실행 중-------------")
        time.sleep(900)


if __name__ == '__main__':
    main()
