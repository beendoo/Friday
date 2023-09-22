# FRIDAY 껍데기 # 알림 # 스케줄링
# 폴더에 포함되어야하는 것 - 프라이데이.jpg, win11toast_b.py
# 윈도우 시작 시 자동실행 켜기 설정 안내

import schedule
import time
from pprint import pprint
import os
from win11toast_b import notify
import subprocess
import csv
import re
import requests
import tableread
import pandas as pd
import urllib3
from bs4 import BeautifulSoup


# pyinstaller로 실행파일1개만 생성, 터미널 창 뜨지않는 옵션, 실행파일 아이콘 변경

def job():
 
    

    # 1. 설치된 프로그램 목록 가져오기
    subprocess.call(["powershell.exe", os.path.dirname(os.path.abspath(__file__))+"\\start.ps1"])
    urllib3.disable_warnings()
    
    # 2. Kisa홈페이지에서 테이블 크롤링 해오기
    # 2-1. 파일 초기설정 - 열 제목

    try:
        with open("final.tsv",'r',encoding='cp949') as f:
            old_num=int(f.readlines()[-1].split("\t")[-1].strip())
            if old_num =="페이지번호":
                old_num=5816
    except:
        with open("final.tsv",'w',encoding='cp949') as f:
            f.write("제품명\t해결버전\t페이지번호" + '\n')
            old_num=5816 #2023년 첫번쨰 글이 5817

    # 2-2. 원하는 페이지 찾아서 table가져와서 저장하기.
    wrong_page = 0

    for page_index in range(old_num + 1, old_num + 100):
        url=f"https://knvd.krcert.or.kr/detailSecNo.do?IDX={page_index}"
        response=requests.get(url,verify=False)
        notice=response.text
        soup = BeautifulSoup(notice, features="html.parser")
        new_soup = soup.select_one('body').get_text()
    
        if 'WRONG ACCESS' in new_soup : 
            wrong_page += 1
            if wrong_page > 10 :
                break
            else :
                continue
        
        wrong_page = 0

        try:
            rows = soup.select_one('#tab_1 > div > .notice_contents').select_one('.basicView > tbody > tr:nth-child(2) > td > table.se_tbl_ext') # table의 클래스 이름이 se_tbl_ext일 때 case3
            if rows==None:
                rows = soup.select_one('#tab_1 > div > .notice_contents').select_one('.basicView > tbody > tr:nth-child(2) > td > table > tbody > tr > td ').select_one('table > tbody') # table의 클래스이름이 table일 때 case4
            try:
                totalColCnt=(len(rows.select_one('tr:nth-child(1)').find_all('td')))
                df=pd.DataFrame(tableread.table_to_2d(BeautifulSoup(str(rows).replace("\n","").replace("  ",""), features='html.parser')))
                df.drop(df.index[0], inplace = True)
                add_page=pd.DataFrame(list(page_index for _ in range(len(df.index))), columns = ['페이지번호'])
                add_page.index+=1
                result=pd.concat([df,add_page],axis=1)
                save_file=result.dropna()[[0,totalColCnt-1, '페이지번호']]
                save_file.to_csv("final.tsv",sep="\t",header=False,index=False,mode='a',encoding='cp949')
            finally:
                continue
        finally:
            continue



    # 3.가져온 데이터 정제하기 - 엔터,한글,영어로만 된 버전 제거
    korean = re.compile('[\u3131-\u3163\uac00-\ud7a3]+') #한글 없애는 코드
    english = re.compile('^[a-zA-Z]*$') # 영어만 이루어진 버전 없애는 코드

    with open('final.tsv', "r", encoding='cp949') as update_list:
        update_list = csv.reader(update_list, delimiter="\t")
        secure_dic = {}

        for search in update_list:
            name = search[0]
            version = re.sub(korean,'',search[1]).strip()
            page = search[2].strip()
            if version == "": continue        
            if secure_dic.get(name):
                dic_version = secure_dic[name][0]
                dic_page = secure_dic[name][-1]
                if dic_page < page:
                    secure_dic[name] = [[version],page]                
                elif dic_page == page:
                    dic_version.append(version)
                    secure_dic[name] = [dic_version,page]
            else:
                secure_dic[name] = [[version],page]


    # 4. 기존 설치된 파일들과 업데이트 리스트 비교하여 알림.

    img_path = os.path.dirname(os.path.abspath(__file__))

    with open(os.path.dirname(os.path.abspath(__file__))+'\\installed_programs.csv', 'r', encoding='utf-8') as installed:
        for search in installed.readlines():
            name = search.split(",")[0].strip('"')
            version = search.split(",")[1]
            for key_list in list(secure_dic.keys()):
                if name in key_list:
                    if secure_dic[name][0][-1] == version:
                        notify(f'{name} : 최신 버전입니다.\n굿b',
                        buttons=[
                                {'activationType':'protocol', 'arguments':f'https://knvd.krcert.or.kr/detailSecNo.do?IDX={secure_dic[name][-1]}', 'content':'관련링크열기'},
                                {'activationType':'protocol', 'arguments':'dismiss', 'content':'닫기'}
                                ],
                                icon=(img_path+'\\'+'F.R.I.D.A.Y._(Marvel_Comics).jpg'))
                        
                    else:
                        notify(f'{name} : 이전 버전이니\n{secure_dic[name][0][-1]}으로 업데이트 하세요',
                        buttons=[
                                {'activationType':'protocol', 'arguments':f'https://knvd.krcert.or.kr/detailSecNo.do?IDX={secure_dic[name][-1]}', 'content':'관련링크열기'},
                                {'activationType':'protocol', 'arguments':'dismiss', 'content':'닫기'}
                                ],
                                icon=(img_path+'\\'+'F.R.I.D.A.Y._(Marvel_Comics).jpg'))
                        
                elif key_list in name:
                    if secure_dic[key_list][0][-1] == version:
                        notify(f'{key_list} : 최신 버전입니다.\n굿b',
                        buttons=[
                                {'activationType':'protocol', 'arguments':f'https://knvd.krcert.or.kr/detailSecNo.do?IDX={secure_dic[key_list][-1]}', 'content':'관련링크열기'},
                                {'activationType':'protocol', 'arguments':'dismiss', 'content':'닫기'}
                                ],
                                icon=(img_path+'\\'+'F.R.I.D.A.Y._(Marvel_Comics).jpg'))
                        
                    else:
                        notify(f'{key_list} : 이전 버전이니\n{secure_dic[key_list][0][-1]}으로 업데이트 하세요',
                        buttons=[
                                {'activationType':'protocol', 'arguments':f'https://knvd.krcert.or.kr/detailSecNo.do?IDX={secure_dic[key_list][-1]}', 'content':'관련링크열기'},
                                {'activationType':'protocol', 'arguments':'dismiss', 'content':'닫기'}
                                ],
                                icon=(img_path+'\\'+'F.R.I.D.A.Y._(Marvel_Comics).jpg'))
                        


#실행주기

schedule.clear()    #실행 전 스케줄 삭제

#schedule.every(10).seconds.do(job)  #10초마다 실행
schedule.every().day.at("22:14:10").do(job)    #매일 정해진 시간에 실행

pprint(schedule.jobs)   #스케줄 목록 확인

while True:
   
    #while로 무한히 스케줄 체크

    schedule.run_pending()

    #위의 함수를 1초 주기로 호출, 등록 스케줄 job 실행

    time.sleep(1)
    
