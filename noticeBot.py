import time
import datetime
import win32con
import win32api
import win32gui
import requests
import json
import logging
from bs4 import BeautifulSoup
from operator import eq
from apscheduler.schedulers.background import BackgroundScheduler
from logging.handlers import TimedRotatingFileHandler

# 카톡창 이름 리스트
kakao_opentalk_name = ['noticebot', 'noticebot2']
idx = 0

# 로거 설정
def set_logger():
    global botLogger
    botLogger = logging.getLogger("KakaoBot")
    botLogger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    
    rotatingHandler = TimedRotatingFileHandler(
        filename='./noticebot_log/webCrawling.log', when='midnight', encoding='utf-8', backupCount=7)
    rotatingHandler.setLevel(logging.DEBUG)
    rotatingHandler.setFormatter(formatter)
    rotatingHandler.suffix = "%Y-%m-%d"
    
    botLogger.addHandler(rotatingHandler)
    botLogger.info("Logger initialized.")

# 채팅방 열기
def open_chatroom(chatroom_name):
    botLogger.info(f"[open_chatroom] Trying to open chatroom: {chatroom_name}")
    hwnd_kakao = win32gui.FindWindow(None, "카카오톡")
    hwnd_edit1 = win32gui.FindWindowEx(hwnd_kakao, None, "EVA_ChildWindow", None)
    hwnd_edit2_1 = win32gui.FindWindowEx(hwnd_edit1, None, "EVA_Window", None)
    hwnd_edit2_2 = win32gui.FindWindowEx(hwnd_edit1, hwnd_edit2_1, "EVA_Window", None)
    hwnd_edit3 = win32gui.FindWindowEx(hwnd_edit2_2, None, "Edit", None)

    if hwnd_edit3 == 0:
        botLogger.error(f"[open_chatroom] Failed to find chatroom search box.")
        return False
    
    win32api.SendMessage(hwnd_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)
    SendReturn(hwnd_edit3)
    time.sleep(1)
    botLogger.info(f"[open_chatroom] Chatroom '{chatroom_name}' opened.")
    return True

# 엔터 입력
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, noticeLists):
    botLogger.info(f"[kakao_sendtext] Sending {len(noticeLists)} messages to '{chatroom_name}'")
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RICHEDIT50W", None)
    
    if hwndEdit == 0:
        botLogger.error(f"[kakao_sendtext] Failed to find chat input box for '{chatroom_name}'")
        return
    
    for notice in noticeLists:
        message = f"📢 [공지사항] {notice['date']}\n🔹 제목: {notice['title']}\n🔗 링크: {notice['link']}"
        win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, message)
        SendReturn(hwndEdit)
        botLogger.info(f"[kakao_sendtext] Message sent: {message}")
        time.sleep(3)
    
    botLogger.info(f"[kakao_sendtext] Completed sending messages to '{chatroom_name}'")

# 공지사항 크롤링
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

    notice_set = set()
    for element in elements:
        id = int(element.a.get('onclick').split("'")[1])
        title = element.a.text.strip()
        date = element.find_all('span', 'p_hide')[1].text
        notice_set.add({"id": id, "title": title, "date": date, "link": f"{dongduk_url}{id}"})

    new_notices = [el for el in notice_set if el["id"] > idx]
    if new_notices:
        new_notices.sort(key=lambda x: x["id"])
        idx = new_notices[-1]["id"]
        botLogger.info(f"[get_dwu_notice] {len(new_notices)} new notices fetched.")
        return new_notices
    
    botLogger.info("[get_dwu_notice] No new notices found.")
    return []

# 스케줄러 job
def job():
    botLogger.info("[job] Running scheduled job...")
    noticeList = get_dwu_notice()
    
    for chatroom in kakao_opentalk_name: 
        if open_chatroom(chatroom):
            kakao_sendtext(chatroom, noticeList)
    
    botLogger.info("[job] Job completed.")

# 메인 함수
def main():
    set_logger()
    botLogger.info("Bot is starting...")
    sched = BackgroundScheduler()
    sched.start()
    sched.add_job(job, 'interval', minutes=15)
    
    while True:
        botLogger.debug("[main] Bot is running...")
        time.sleep(900)

if __name__ == '__main__':
    main()
