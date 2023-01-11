import sys, requests
from tkinter import EXCEPTION
from bs4 import BeautifulSoup # 크롤러 라이브러리
import re
import datetime
import threading
import time
import requests
import win32api
import win32con
import win32gui
from socket import *


# 일단 작동방식:
# 최초에 init_notice()와 init_swnotice()를 통해서 학과 공지 랑 sw공지를 크롤링해서 normalorigindata, sworigindata이라는 집합 변수에 집어 넣음
# 이후에 main함수에서 10분에 한번씩 크롤링을 해서 최초 집합 normalorigindata, sworigindata랑 비교해서 변동사항이 생기면 kakao_send_text()함수를 이용해서
# 프로세스에 있는 카카오 채팅방의 이름을 확인함 비교하는 데이터는 kakao_send_text에 있는 caption이라는 문자열과 비교해서 같은 문자열의 프로세스가 있으면
# 다른 자료를 보내는데, 채팅방을 먼저 active해서 띄운 다음 문자열을 입력, 엔터를 눌러 보내는 형식으로 구석봇이 작동함
class KakaoUtil:
# 키보드 엔터 입력
    def push_enter(self,hwnd):
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        time.sleep(0.1)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

    # 카카오 메세지 보내기
    def kakao_send_text(text, captions):
        for caption in captions:
            hwnd_main = win32gui.FindWindow(None, caption)
            hwnd_edit = win32gui.FindWindowEx(hwnd_main, None, "RichEdit50W", None)
            win32api.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, text)
            self.push_enter(hwnd_edit)
            time.sleep(0.1)

class CornerBot:
    utility = KakaoUtil()

    def __init__(self):
        self.normalorigindata = set()
        self.sworigindata = set()
        self.joborigindata = set()
        
        self.normalurl = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6694&bbsId=2351"
        self.swurl = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6695&bbsId=2352"
        self.joburl = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttList.do?mi=6696&bbsId=2353"

        self.normalhref = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6694&bbsId=2351&nttSn="
        self.swhref = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6695&bbsId=2352&nttSn="
        self.jobhref = "https://newgh.gnu.ac.kr/cs/na/ntt/selectNttInfo.do?mi=6696&bbsId=2353&nttSn="

        self.common_captions = ['cs 1학년', 'cs 2학년', 'cs 3학년', 'cs 4학년', 'cs 외국인']
        self.senior_captions = ['cs 4학년', 'cs 졸업', '졸업']
        self.admin_captions = ['구석방']

    def get_crawl_data(self,url):
        get_data = requests.get(url, verify=False) # requests를 이용하여 해당 url 주소로 get 요청을 보내고 응답을 받는다. 상태 코드와 HTML 내용을 응답 받을 수 있음
        return BeautifulSoup(get_data.content, "lxml") # 응답받은 HTML 내용을 BeautifulSoup 클래스의 객체 혀애로 생성 후 반환

    def formatting_data(self, data):
        origin=[]
        for i in data:
            templink = i.get("data-id")
            origin.append(templink)
        return origin

    def init_notice(self):
        self.normalorigindata = set()
        self.temp_list = []
        print("공지 초기화")    

        normalcroll = self.get_crawl_data(self.normalurl)
        normalnotice = normalcroll.select("a.nttInfoBtn")
        self.normalorigindata = set(self.formatting_data(normalnotice))

    def init_swnotice(self):
        self.sworigindata = set()
        self.temp_list = []
        print("SW공지 초기화")

        swcroll = self.get_crawl_data(self.swurl)
        swnotice = swcroll.select("a.nttInfoBtn")

        self.sworigindata = set(self.formatting_data(swnotice))

    def init_jobinfo(self):
        self.joborigindata = set()
        self.temp_list = []
        print("취업정보 초기화")

        jobcroll = self.get_crawl_data(self.joburl)
        jobinfo = jobcroll.select("a.nttInfoBtn")

        self.joborigindata = set(self.formatting_data(jobinfo))
    
    def checking_data(self, url):
        try:
            croll_data = self.get_crawl_data(url)
        except:
            errorstring = "에러발생 확인 하셈 빨리"
            self.utility.kakao_send_error(errorstring, self.admin_captions)
            self.init_notice()
            self.init_swnotice()
            self.init_jobinfo()
            starting_crolling()

        tempnotice = croll_data.select("a.nttInfoBtn")
        temp = []
        templist = []
        for i in tempnotice:
            temptitle = re.sub("\t|\r|\n", "", i.text)
            templink = i.get("data-id")
            temp.append(templink)
            templist.append({"href": templink, "title": temptitle})
        return templist, temp
    
    def send_diff(self, origin, temps ,templist, href, caption):
        temp = set(temps) - origin
        if (len(temp) > 0):
            for i in temp:
                for index in range(len(templist)):
                    if (i == templist[index].get("href")):
                        title = templist[index].get("title") + "\n" + href + templist[index].get("href")
                        print(title)
                        self.utility.kakao_send_text(title, caption)
            origin.update(temp)

    def normalcheck_notice(self):
        print("normalcheck")
        templist, temp= self.checking_data(self.normalurl)
        self.send_diff(self.normalorigindata, temp, templist, self.normalhref, self.common_captions)

    def swcheck_notice(self):
        print("swcheck")
        templist, temp= self.checking_data(self.swurl)
        self.send_diff(self.sworigindata, temp, templist, self.swhref, self.common_captions)

    def job_info(self):
        print("jobinfocheck")
        templist, temp= self.checking_data(self.joburl)
        self.send_diff(self.joborigindata, temp, templist, self.jobhref, self.senior_captions)



def starting_crolling(instance):
    instance.normalcheck_notice()
    instance.swcheck_notice()
    instance.job_info()
    threading.Timer(600, starting_crolling).start()

if __name__ == "__main__":
    print(datetime.datetime.now().strftime('%c'), ': 구석봇v6 가동 시작')
    bot1 = CornerBot()
    bot1.init_notice()
    bot1.init_swnotice()
    bot1.init_jobinfo()
    starting_crolling(bot1)
