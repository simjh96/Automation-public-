import os
from win32com.client import Dispatch
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import random
import re
import os
import datetime
import pandas as pd
from soynlp.hangle import jamo_levenshtein
import fnmatch
import shutil
import os
import re



class PathFinder():
    """
    lots of cleaning up to do....
    """

    #0계층
    def __init__(self,out_dir,target_dir):
        print('시스템>>> init')
        #variables
        self.out_dir = out_dir
        self.url_dic = dict()
        self.url_dic['root_dir'] = target_dir

        self.matched_folder = None
        self.second_matched_folder = None
        self.search_dirs = None
        self.upchae_name = None

        #action
        self.get_search_dirs()

    #1계층
    def get_search_dirs(self):

        #최상위폴더는 날짜 기준
        date_folders = []
        for file in os.listdir(self.url_dic['root_dir']):
            #임시방편... num_to_date 만들어야함

            if len(re.findall(r'^20[0-9]{6}',file))>0:
                if '20210630' not in file:
                    date_folders.append(file)

        #진행중인 20210430, 20210331 제외한 20210228 부터 확인 시작 하려면 뒤에 인덱스 달아줘 ~ 9개 폴더까지만...[2:]
        date_folders = pd.Series(date_folders).sort_values(ascending=False)

        search_dirs = [(self.url_dic['root_dir'] + '\\' + date_folder) for date_folder in date_folders][:9]
        self.search_dirs = search_dirs
        
        return search_dirs

    def match_and_save(self, upchae_name, copy_tree = False):
        """
        match + save
        at re.findall(r'[^\\]+$',result.iloc[0,0])[0]
        원본 폴더명 그대로 가져옴
        """

        result = self.match_folder(upchae_name)
        self.save_file(result, copy_tree)
        return result

    #2계층
    def match_folder(self, upchae_name):
        """
        return exact matched folder dir, and print next nearest dir
        if copy_tree is True, folder is copied to 'fresh_baked'
        upchae_name 기준으로 sc_type_strip 한 폴더명과 매칭
        
        """
        self.upchae_name = upchae_name
        dist_of_dirs = pd.DataFrame()
        exact = 0
        
        #월별 폴더 주소 loop
        for search_dir in self.search_dirs:
            print('시스템>>> {} 시점 폴더를 확인 합니다.'.format(re.findall(r'\\[0-9]{8}$',search_dir)[0]))
            for root, dirs, files in os.walk(search_dir):
                for dir0 in dirs:
                    dist = jamo_levenshtein(upchae_name, self.sc_type_strip(dir0))
                    full_dir = root + '\\' + dir0
                    result = pd.DataFrame([full_dir, dist]).T
                    if dist == 0:
                        exact = 1
                        print('시스템>>> exact match 된 "{}" 폴더를 찾았습니다.'.format(dir0))
                        print('시스템>>> 두번째로 유사한 위치는 "{}" 였습니다.'.format(dist_of_dirs.sort_values(by=1).iloc[0,0]))

                        #log
                        self.matched_folder = full_dir
                        self.second_matched_folder = dist_of_dirs.sort_values(by=1).iloc[0,0]


                        return result
                    else:
                        dist_of_dirs = dist_of_dirs.append(result)
        
        #exact 가 없으면 가장 유사한 주소 반환                    
        print('시스템>>> exact match 를 찾을 수 없어 가장 유사한 "{}" 폴더를 반환합니다.'.format(dist_of_dirs.sort_values(by=1).iloc[0,0]))
        result = dist_of_dirs.sort_values(by=1).iloc[[0],:]

        #log
        self.matched_folder = 'Not Matched'
        self.second_matched_folder = dist_of_dirs.sort_values(by=1).iloc[0,0]

        return result


    def save_file(self,result, copy_tree = False):
        if copy_tree:
            try:
                shutil.copytree(result.iloc[0,0],self.out_dir +'\\'+ '{}'.format(re.findall(r'[^\\]+$',result.loc[0,0])[0]) + '_전기')
                print('시스템>>> "{}" 복사 완료'.format(self.upchae_name))
            except:
                print('시스템>>> "{}" 관련 폴더가 이미 존재합니다!'.format(self.upchae_name))
        else:
            try:
                path = self.out_dir + '\\{}_바로가기.lnk'.format(re.findall(r'[^\\]+$',result.iloc[0,0])[0])
                target = result.iloc[0,0]
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(path)
                shortcut.Targetpath = target
                shortcut.save()
                time.sleep(0.1)
                print('시스템>>> "{}" 바로가기 생성 완료'.format(self.upchae_name))
            except:
                print('시스템>>> "{}" 관련 바로가기가 이미 존재합니다!'.format(self.upchae_name))


    #3계층
    def sc_type_strip(self, name):
        """
        strip ['RCPS','BW','CB','CPS','보통주','우선주','평가제외','제외']
        + re.sub(r'\(.+?\)','',name.replace('㈜','')).strip()
        """
        
        for i in ['RCPS','BW','CB','CPS','보통주','우선주','평가제외','제외']:
            name = name.replace(i,'')
            name = name.strip()

        name = re.sub(r'\(.+?\)','',name.replace('㈜','')).strip()
        return name


# upchae_name = '크레아아이엔'
# final_dir = match_folder(upchae_name, search_dirs)
# print(final_dir)


