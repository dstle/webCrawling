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

idx = 0

# setting logger
def set_logger():
    global botLogger
    botLogger = logging.getLogger("KakaoBot")
    botLogger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s | PID %(process)d | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S")
    
    rotatingHandler = TimedRotatingFileHandler(
        filename='./noticebot_log/webCrawling.log', when='midnight', encoding='utf-8', backupCount=7)
    rotatingHandler.setLevel(logging.DEBUG)
    rotatingHandler.setFormatter(formatter)
    rotatingHandler.suffix = "%Y-%m-%d"
    
    botLogger.addHandler(rotatingHandler)
    botLogger.info("Logger initialized.")

# Search the chatroom and Open
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

# Click the "Enter"
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    
# Click the "ESC"
def SendEsc(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)

# clean chatroom search box
def cleanChatroom():
    hwnd_kakao = win32gui.FindWindow(None, "카카오톡")
    hwnd_edit1 = win32gui.FindWindowEx(hwnd_kakao, None, "EVA_ChildWindow", None)
    hwnd_edit2_1 = win32gui.FindWindowEx(hwnd_edit1, None, "EVA_Window", None)
    hwnd_edit2_2 = win32gui.FindWindowEx(hwnd_edit1, hwnd_edit2_1, "EVA_Window", None)
    hwnd_edit3 = win32gui.FindWindowEx(hwnd_edit2_2, None, "Edit", None)
    win32api.SendMessage(hwnd_edit3, win32con.WM_SETTEXT, 0, '')

    return hwnd_edit3

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
    SendEsc(hwndEdit)
    botLogger.info(f"[kakao_close_chatroom] Close the '{chatroom_name}' !!")

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

# scheduler job
def job(chatroom_name):
    botLogger.info("[job] Running scheduled job...")
    noticeList = get_dwu_notice()
    
    botLogger.info(f"[job] chatroom name is {chatroom_name}")
    if open_chatroom(chatroom_name):
        kakao_sendtext(chatroom_name, noticeList)
        if cleanChatroom() == 0:
            botLogger.error(f"[job] Failed to clean chatroom.")
            return
        
    botLogger.info("[job] Job completed.")

# Main
def main():
    # add parser
    parser = argparse.ArgumentParser(description='Notice Bot for Dongduk Women\'s University')

    # add parameters
    parser.add_argument('--chatroom', type=str, help='chatroom name')
    parser.add_argument('--verbose', action='store_true', help='verbose output')

    # parse parameters
    args = parser.parse_args()
    chatroom_name = args.chatroom

    # use verbose option to print chatroom name
    if args.verbose:
        print(f"chatroom name: {args.chatroom}")


    set_logger()
    botLogger.info("Bot is starting...")

    sched = BackgroundScheduler()
    sched.start()
    sched.add_job(job, 'interval', minutes=15, args=[chatroom_name])
    
    while True:
        botLogger.debug("[main] Bot is running...")
        time.sleep(900)

if __name__ == '__main__':
    main()