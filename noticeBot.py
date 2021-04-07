import time
import win32con
import win32api
import win32gui
import requests
import json

from operator import eq
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler


# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = 'noticebot'
idx = 0


# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, notice):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RICHEDIT50W", None)

    check = len(notice)
    global idx
    if(idx < check):
        for i in range(idx, check):
            win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, notice[i])
            SendReturn(hwndEdit)
            print(notice[i])
            time.sleep(3)
    idx = check


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

    today = datetime.today().strftime("%Y%m%d")
    url = 'https://www.dongduk.ac.kr/ajax/board/kor/kor_notice/list.json'
    req = requests.get(url)

    html = req.text

    data = json.loads(html)

    temp = data.get('data')
    rslt = temp.get('list')

    noticelist = []
    for i in rslt:
        date = i.get('REG_DT')[:8]
        index = i.get('B_IDX')
        title = i.get('B_TITLE')
        href = 'https://www.dongduk.ac.kr/board/kor/kor_notice/detail.do?curPageNo=1&pageStatus=N&rowSize=15&B_IDX=' + \
            str(index)

        #공지사항 리스트
        if eq(today, date):
            rslt = "[" + date + "]\n" + title + "\n" + href
            noticelist.append(rslt)

    # 리스트 역순 정렬
    noticelist.reverse()

    return noticelist


# # 스케줄러 job : 매 시간마다 공지사항 크롤링해서 가져오기
def job():
    p_time_ymd_hms = \
        f"{time.localtime().tm_year}-{time.localtime().tm_mon}-{time.localtime().tm_mday} / " \
        f"{time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"

    open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    notice = get_dwu_notice()
    kakao_sendtext(kakao_opentalk_name, notice)  # 메시지 전송, time/실검


def main():
    sched = BackgroundScheduler()
    sched.start()

    # 15분마다 실행
    sched.add_job(job, 'interval', minutes=15)

    while True:
        print("실행중.................")
        time.sleep(60)


if __name__ == '__main__':
    main()
