import sys, requests
from bs4 import BeautifulSoup
import re
import datetime
import threading
import time
import requests
import win32api
import win32con
import win32gui

normalorigindata = set()
normaloriginlist = []
normalorigin = []

sworigindata = set()
sworiginlist = []
sworigin = []

bugcounter = 0

def findPost(flag, newText, lateText, chatNumber):  # 놓친거 있을때 놓친거 제일 최신거와 제일 늦은거 입력 , 공백은 모두 지워서
    global normalorigin, sworigin, normaloriginlist, sworiginlist
    normaltemp = normalorigin
    swtemp = sworigin

    swhref = "https://www.gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6695&bbsId=2352&nttSn="
    normalhref = "https://www.gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6694&bbsId=2351&nttSn="
    startAt = int(newText)
    endAt = int(lateText)
    point = int(flag)
    chatChoice = int(chatNumber)

    for i in range(len(normaltemp)):  # 공백제거
        normaltemp[i] = normaltemp[i].replace(' ', '')
    for i in range(len(swtemp)):
        swtemp[i] = swtemp[i].replace(' ', '')

    if (point == 1):
        for i in range(startAt, endAt + 1):
            title = str(sworiginlist[i].get('title')) + swhref + str(sworiginlist[i].get('href'))
            print(title)
            kakao_send_text(title)

    elif (point == 0):
        for i in range(startAt, endAt + 1):
            title = str(normaloriginlist[i].get('title')) + normalhref + str(normaloriginlist[i].get('href'))
            print(title)
            kakao_send_text(title, chatChoice)



def init_notice():
    global normalorigindata, normaloriginlist, normalorigin, bugcounter
    bugcounter = 0
    normalorigindata = set()
    normaloriginlist = []
    normalorigin = []
    print("공지 초기화")
    normalurl = "https://www.gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6694&bbsId=2351"
    normal = requests.get(normalurl)
    normalcroll = BeautifulSoup(normal.content, "lxml")
    normalnotice = normalcroll.select("a.nttInfoBtn")

    for i in range(len(normalnotice)):
        temptitle = re.sub("\t|\r|\n", "", normalnotice[i].text)
        templink = normalnotice[i].get("data-id")
        normaloriginlist.append({"href": templink, "title": temptitle})
        normalorigin.append(temptitle)
    for i in range(len(normalorigin)):
        print(str(i)+ " " + normalorigin[i])

def init_swnotice():
    global sworigin, sworiginlist, sworigindata, bugcounter
    bugcounter = 0
    sworigin = []
    sworiginlist = []
    sworigindata = set()
    print("SW공지 초기화")
    swurl = "https://www.gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6695&bbsId=2352"
    sw = requests.get(swurl)
    swcroll = BeautifulSoup(sw.content, "lxml")
    swnotice = swcroll.select("a.nttInfoBtn")

    for i in range(len(swnotice)):
        temptitle = re.sub("\t|\r|\n", "", swnotice[i].text)
        templink = swnotice[i].get("data-id")
        sworiginlist.append({"href": templink, "title": temptitle})
        sworigin.append(temptitle)
    for i in range(len(sworigin)):
        print(str(i)+ " " + sworigin[i])

# 키보드 엔터 입력
def push_enter(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.1)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


# 카카오 메세지 보내기
def kakao_send_text(text, chatNumber):
    # tester = ['관리자1', '관리자2', '관리자3', '관리자4']
    captions = ['cs 1학년', 'cs 2학년', 'cs 3학년', 'cs 4학년', 'cs 외국인']
    #captions = ['구석방', '이가현조교님','이자룡조교님','박유리조교님']
    #captions = ['황혁주', '구석방']
    if chatNumber == 5:
        for caption in captions:
            hwnd_main = win32gui.FindWindow(None, caption)
            hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RichEdit50W", None)
            win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, text)
            push_enter(hwnd_edit)
            time.sleep(0.1)
    else :
        hwnd_main = win32gui.FindWindow(None, captions[chatNumber])
        hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RichEdit50W", None)
        win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, text)
        push_enter(hwnd_edit)
        time.sleep(0.1)


if __name__ == "__main__":
    print(datetime.datetime.now().strftime('%c'), ': 구석봇v2.0 가동 시작')
    init_notice()
    print("-------------------------------------------------------------------------------------------------")
    init_swnotice()
    d = input("원하는 공지방을 선택하세요 (예시 0번 : 'cs 1학년', 1번 : 'cs 2학년', 2번 : 'cs 3학년', 3번 : 'cs 4학년', 4번 : 'cs 외국인', 5번 : 전체) : ")
    c = input("공지를 선택하세요(예시 1번 sw공지, 0번 일반 공지) : ")
    a = input("시작 인덱스를 선택하세요 : ")
    b = input("끝 인덱스를 선택하세요 : ")

    findPost(c, a, b, d)