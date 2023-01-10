import sys, requests
from tkinter import EXCEPTION
from bs4 import BeautifulSoup
import re
import datetime
import threading
import time
import requests
import win32api
import win32con
import win32gui
from socket import *

normalorigindata = set()
normaloriginlist = []
normalorigin = []

sworigindata = set()
sworiginlist = []
sworigin = []

joborigindata = set()
joboriginlist = []
joborigin = []


def init_notice():
    global normalorigindata, normaloriginlist, normalorigin
    normalorigindata = set()
    normalorigin = []
    normaloriginlist = []
    print("공지 초기화")
    normalurl = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6694&bbsId=2351"
    normal = requests.get(normalurl, verify=True)
    normalcroll = BeautifulSoup(normal.content, "lxml")
    normalnotice = normalcroll.select("a.nttInfoBtn")

    for i in range(len(normalnotice)):
        temptitle = re.sub("\t|\r|\n", "", normalnotice[i].text)
        templink = normalnotice[i].get("data-id")
        normaloriginlist.append({"href": templink, "title": temptitle})
        normalorigin.append(templink)
    normalorigindata = set(normalorigin)


def init_swnotice():
    global sworigin, sworiginlist, sworigindata
    sworigin = []
    sworiginlist = []
    sworigindata = set()
    print("SW공지 초기화")
    swurl = "https://gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6695&bbsId=2352"
    sw = requests.get(swurl, verify=True)
    swcroll = BeautifulSoup(sw.content, "lxml")
    swnotice = swcroll.select("a.nttInfoBtn")

    for i in range(len(swnotice)):
        temptitle = re.sub("\t|\r|\n", "", swnotice[i].text)
        templink = swnotice[i].get("data-id")
        sworiginlist.append({"href": templink, "title": temptitle})
        sworigin.append(templink)

    sworigindata = set(sworigin)

def init_jobinfo():
    global joborigin, joborigindata, joboriginlist
    joborigin = []
    joboriginlist = []
    joborigindata = set()
    print("취업정보 초기화")
    joburl = "https://gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6696&bbsId=2353"
    job = requests.get(joburl, verify=True)
    jobcroll = BeautifulSoup(job.content, "lxml")
    jobinfo = jobcroll.select("a.nttInfoBtn")

    for i in range(len(jobinfo)):
        temptitle = re.sub("\t|\r|\n", "", jobinfo[i].text)
        templink = jobinfo[i].get("data-id")
        joboriginlist.append({"href": templink, "title": temptitle})
        joborigin.append(templink)

    joborigindata = set(joborigin)

def init_workShop():
    return


def normalcheck_notice():
    global normalorigindata
    normalhref = "https://gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6694&bbsId=2351&nttSn="
    normalurl = "https://gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6694&bbsId=2351"
    try:
        normal = requests.get(normalurl, verify=True)
        normalcroll = BeautifulSoup(normal.content, "lxml")
    except:
        errorstring = "에러발생 확인 하셈 빨리"
        kakao_send_error(errorstring)
        init_notice()
        init_swnotice()
        starting_crolling()

    tempnotice = normalcroll.select("a.nttInfoBtn")
    temp = []
    tempdata = set()
    templist = []

    for i in range(len(tempnotice)):
        temptitle = re.sub("\t|\r|\n", "", tempnotice[i].text)
        templink = tempnotice[i].get("data-id")
        temp.append(templink)
        templist.append({"href": templink, "title": temptitle})
    nmtemp = normalorigindata
    tempdata = set(temp)
    temp = tempdata - nmtemp

    if (len(temp) > 0):
        for i in temp:
            for index in range(len(templist)):
                if (i == templist[index].get("href")):
                    title = templist[index].get("title") + "\n" + normalhref + templist[index].get("href")
                    kakao_send_text(title)
        init_notice()


def swcheck_notice():
    global sworigindata, sworigin

    swurl = "https://gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6695&bbsId=2352"
    swhref = "https://gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6695&bbsId=2352&nttSn="
    try:
        sw = requests.get(swurl, verify=True)
        swcroll = BeautifulSoup(sw.content, "lxml")
    except:
        errorstring = "에러발생 확인 하셈 빨리"
        kakao_send_error(errorstring)
        init_notice()
        init_swnotice()
        starting_crolling()

    tempnotice = swcroll.select("a.nttInfoBtn")
    temp = []
    tempdata = set()
    templist = []

    for i in range(len(tempnotice)):
        temptitle = re.sub("\t|\r|\n", "", tempnotice[i].text)
        templink = tempnotice[i].get("data-id")
        temp.append(templink)
        templist.append({"href": templink, "title": temptitle})
    swtemp = sworigindata
    tempdata = set(temp)
    temp = tempdata - swtemp

    if (len(temp) > 0):
        for i in temp:
            for index in range(len(templist)):
                if (i == templist[index].get("href")):
                    title = templist[index].get("title") + "\n" + swhref + templist[index].get("href")
                    kakao_send_text(title)
        init_swnotice()

def job_info():
    global joborigindata, joborigin

    joburl = "https://gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6696&bbsId=2353"
    jobhref = "https://gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6696&bbsId=2353&nttSn="
    try:
        job = requests.get(joburl, verify=True)
        jobcroll = BeautifulSoup(job.content, "lxml")
    except:
        errorstring = "에러발생 확인 하셈 빨리"
        kakao_send_error(errorstring)
        init_notice()
        init_swnotice()
        starting_crolling()

    tempnotice = jobcroll.select("a.nttInfoBtn")
    temp = []
    tempdata = set()
    templist = []

    for i in range(len(tempnotice)):
        temptitle = re.sub("\t|\r|\n", "", tempnotice[i].text)
        templink = tempnotice[i].get("data-id")
        temp.append(templink)
        templist.append({"href": templink, "title": temptitle})
    jobtemp = joborigindata
    tempdata = set(temp)
    temp = tempdata - jobtemp

    if (len(temp) > 0):
        for i in temp:
            for index in range(len(templist)):
                if (i == templist[index].get("href")):
                    title = templist[index].get("title") + "\n" + jobhref + templist[index].get("href")
                    kakao_send_text_jobInfo(title)
        init_jobinfo()



# 키보드 엔터 입력
def push_enter(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.1)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


# 카카오 메세지 보내기
def kakao_send_text(text):
    # tester = ['관리자1', '관리자2', '관리자3', '관리자4']
    captions = ['cs 1학년', 'cs 2학년', 'cs 3학년', 'cs 4학년', 'cs 외국인']
    # captions = ['구석방', '이가현조교님','이자룡조교님','박유리조교님']
    # captions = ['구석방']
    for caption in captions:
        hwnd_main = win32gui.FindWindow(None, caption)
        hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RichEdit50W", None)
        win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, text)
        push_enter(hwnd_edit)
        time.sleep(0.1)

def kakao_send_text_jobInfo(text):
    # tester = ['관리자1', '관리자2', '관리자3', '관리자4']
    captions = ['cs 4학년', 'cs 졸업']
    # captions = ['구석방', '이가현조교님','이자룡조교님','박유리조교님']
    # captions = ['구석방']
    for caption in captions:
        hwnd_main = win32gui.FindWindow(None, caption)
        hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RichEdit50W", None)
        win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, text)
        push_enter(hwnd_edit)
        time.sleep(0.1)


# 카카오 메세지 에러용
def kakao_send_error(error):
    captions = ['구석방']
    # captions = ['관리자1','관리자2','관리자3']
    for caption in captions:
        hwnd_main = win32gui.FindWindow(None, caption)
        hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RichEdit50W", None)
        win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, error)
        push_enter(hwnd_edit)


def starting_crolling():
    normalcheck_notice()
    swcheck_notice()
    job_info()
    threading.Timer(600, starting_crolling).start()

def still_Alive():
    kakao_send_error("쿠쿸")
    threading.Timer(28800, still_Alive).start()

if __name__ == "__main__":
    print(datetime.datetime.now().strftime('%c'), ': 구석봇v4 가동 시작')
    init_notice()
    init_swnotice()
    init_jobinfo()
    starting_crolling()
    still_Alive()