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


class KisChrome():
    """
    - init에 action 없이 함수만 빼서 사용하는게 좋을듯
        driver는 불러오고 써야하니까 instance 만들기는 해야함

    - dir 변수로 받을 수 있도록 해줘야함
        - get upmoo series dir param으로

    - 시스템 안내문 달아    

    - url dic 이랑 ss path 정리해 self.variables로
    """

    # 0계층
    def __init__(self, upmoo_dir):
        print('시스템>>> init')
        # >>> 간이로 만든 내부변수.. 업무배분 파일.. 원래 전체로 넣어서 돌리면 좋은데
        # 지금 구현은 귀찮으니까... 하나의 고객사 하나의 업무담당자 파일 받아서 일단 for
        self.upmoo_dir = upmoo_dir

        # variables
        self.url_dic = {'root_chrome': ''}
        self.upmoo_df = None
        self.loged_in = False
        self.rep = 'None'

        # 업무 리스트가 있어야 루프를 돌린다... 나중에 따로 빼던지

        # 업무 배정된 사람 걸로 바꿔야함
        self._id = ''
        self._pw = ''
        # actions
        self.start_driver()
        self.log_in_kischrome()
        self.get_upmoo_df()

    # 1계층

    def get_upmoo_df(self):
        """
        get_upmoo_series -> upmoo_df
        """
        print('시스템>>> get_upmoo_df')
        # 종목 정보(크롬) 가져오기
        upmoo_series = self.get_upmoo_series()
        final = []
        for upmoo_name in upmoo_series:
            final.append(self.get_upchae_info(upmoo_name))

        final = pd.concat([upmoo_series, pd.DataFrame(final)], axis=1)
        final.columns = ['종목명(업무배분_clean)', '종목명(크롬)', '업체코드', '이질도']
        self.upmoo_df = final
        return final

    def start_driver(self):
        # chrome driver dir check
        # go to dart
        # print("시스템>>> 현재 디렉토리: {}".format(os.getcwd()))
        # print("시스템>>> 디렉토리가 맞다면(Y) 아니면(N): ")
        # inp = input()
        # if inp == 'N':
        #     print("시스템>>> 크롬드라이버가 있는 위치 입력...")
        #     os.chdir(input())
        self.driver = webdriver.Chrome("./chromedriver.exe")
        print("시스템>>> 드라이버를 실행 합니다...")

    def log_in_kischrome(self):
        print('시스템>>> log_in_kischrome')
        self.driver.get(self.url_dic['root_chrome'])
        self.driver.find_elements_by_class_name(
            'form-control')[0].send_keys(self._id)
        self.driver.find_elements_by_class_name(
            'form-control')[1].send_keys(self._pw)
        self.driver.find_elements_by_class_name(
            'form-control')[1].send_keys(Keys.ENTER)

    def kischrome_ss(self, series_with_code, ss_path):
        """
        input upchae_code -> takes ss 
        """
        print('시스템>>> kischrome_ss')
        os.makedirs(ss_path, exist_ok=True)

        self.driver.get('?upchecd={}'.format(series_with_code['업체코드']))
        self.driver.maximize_window()
        try:
            WebDriverWait(self.driver, 5).until(EC.text_to_be_present_in_element(
                (By.XPATH, '//*[@id="gpinfo"]/div[1]/div/div/div/table/tbody/tr[1]/td[1]'), r"기업명"))
        except:
            print('시스템>>> 크롬 업체정보에 등록이 안되어 있음')
        time.sleep(1)
        self.driver.get_screenshot_as_file(
            ss_path + '\\'+'{}_chrome_upchae.png'.format(series_with_code['종목명(업무배분_clean)']))
        self.driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        self.driver.get_screenshot_as_file(
            ss_path + '\\'+'{}_chrome_history.png'.format(series_with_code['종목명(업무배분_clean)']))
        time.sleep(1)

    # 업체 코드로 kisline 스크린샷해서 엑셀 저장

    def kisline_ss(self, series_with_code, ss_path, with_ratings=True):
        """
        series_with_code = get_upmoo_df(self) 중 1줄 준용
        1 series -> 1 ss 준용

        series['업체코드']

        _kisline_main.png
        _kisline_rating.png
        _kisline_credit.png

        return rating

        rating 분리 필요...
        """
        print('시스템>>> kisline_ss')
        os.makedirs(ss_path, exist_ok=True)

        self.driver.get('https://www.kisline.com/')
        self.driver.maximize_window()

        # 팝업 처리
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.ID, 'userGuideCookie')))
            self.driver.find_element('id', 'userGuideCookie').click()
            self.driver.find_element_by_xpath(
                '//*[@id="layout_guide"]/div[2]/a').click()
        except:
            time.sleep(0.1)
        time.sleep(1)

        # login
        if self.loged_in == False:
            self.driver.find_element('id', "btn_login").click()
            time.sleep(4)

            _id = 'simjh96'
            _pw = '104ehd902gh@'
            input_js = ' \
            document.getElementById("lgnuid").value = "{id}"; \
            '.format(id=_id)
            self.driver.execute_script(input_js)
            time.sleep(1)

            try:
                self.driver.find_element('id', 'tmp_lgnupassword').click()
            except:
                self.driver.find_element('id', 'lgnupassword').click()
            time.sleep(1)

            try:
                input_js = ' \
                    document.getElementById("lgnupassword").value = "{pw}"; \
                        '.format(pw=_pw)
                self.driver.execute_script(input_js)
            except:
                input_js = ' \
                    document.getElementById("tmp_lgnupassword").value = "{pw}"; \
                        '.format(pw=_pw)
                self.driver.execute_script(input_js)

            try:
                self.driver.find_element(
                    'id', 'tmp_lgnupassword').send_keys(Keys.ENTER)
            except:
                self.driver.find_element(
                    'id', 'lgnupassword').send_keys(Keys.ENTER)
            time.sleep(1)
            self.loged_in = True

        ratings = 'None'

        # single entry

        if (series_with_code['업체코드'] is not None):
            if (series_with_code['업체코드'][0] == 'K'):

                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.ID, 'q')))
                    self.driver.find_element_by_xpath(
                        '//*[@id="q"]').send_keys(series_with_code['업체코드'][1:])
                except:
                    self.driver.refresh()
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.ID, 'q')))
                    self.driver.find_element_by_xpath(
                        '//*[@id="q"]').send_keys(series_with_code['업체코드'][1:])

                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, 'searchView')))
                self.driver.find_element('id', 'searchView').click()

                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="eprTable"]/tbody/tr/td[2]/a')))
                self.rep = self.driver.find_element_by_xpath(
                    '//*[@id="eprTable"]/tbody/tr/td[3]').text

                self.driver.find_element_by_xpath(
                    '//*[@id="eprTable"]/tbody/tr/td[2]/a').click()
                time.sleep(1.5)
                self.driver.get_screenshot_as_file(
                    ss_path + '\\' + '{}_kisline_main.png'.format(series_with_code['종목명(업무배분_clean)']))
                self.driver.execute_script("window.scrollTo(0, 600);")
                self.driver.get_screenshot_as_file(
                    ss_path + '\\' + '{}_kisline_main_kpi.png'.format(series_with_code['종목명(업무배분_clean)']))

                self.driver.find_element_by_xpath(
                    '//*[@id="content"]/div[1]/ul[6]/li/a').click()
                time.sleep(0.1)
                self.driver.find_element_by_xpath(
                    '//*[@id="content"]/div[1]/ul[6]/li/ul/li[1]/a').click()

                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="content"]/div/ul[6]/li/ul/li[3]/a')))
                self.driver.get_screenshot_as_file(
                    ss_path + '\\' + '{}_kisline_rating.png'.format(series_with_code['종목명(업무배분_clean)']))

                src = BeautifulSoup(self.driver.page_source, 'lxml')
                rating = src.find('div', 'right_area').strong.text
                ratings = rating

                self.driver.find_element_by_xpath(
                    '//*[@id="content"]/div/ul[6]/li/ul/li[3]/a').click()

                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="content"]/div/ul[6]/li/ul/li[3]/a')))
                self.driver.get_screenshot_as_file(
                    ss_path + '\\' + '{}_kisline_credit.png'.format(series_with_code['종목명(업무배분_clean)']))

                print('시스템>>> {}의 KisLine 수집을 완료했습니다 ({})'.format(
                    series_with_code['종목명(업무배분_clean)'], rating))
                time.sleep(1)

        else:
            print('시스템>>> {}은 KisLine 수집이 불가능합니다 ({})'.format(
                series_with_code['종목명(업무배분_clean)'], series_with_code['업체코드']))
            ratings.append(0)

        if with_ratings:
            self.ratings = ratings
            return ratings
        else:
            pass

    # 2계층

    def get_upchae_info(self, upmoo_name, dist=True):
        """
        upmoo 1개 -> [upchae, code, 거리] 1개
        """
        print('시스템>>> {}의 종목정보를 통해 기업 코드를 확인합니다'.format(upmoo_name))
        # get src
        self.driver.get('?query={}'.format(self.sc_type_strip(upmoo_name)))
        src = BeautifulSoup(self.driver.page_source, 'lxml')

        # parse all ""
        result = [self.utf_decoder(i.replace('"', '')) for i in re.findall(
            r'"[\\a-zA-Z0-9\(\)_\s\.]+"', src.text)]
        # print('시스템>>> 종목검색 result는 {} 입니다'.format(result))
        if len(result) != 2:

            upchae_names = result[1:(len(result)//2)]
            subj_codes = result[(len(result)//2)+1:]

            print('시스템>>> 종목검색 upchae_names 는 {} 입니다'.format(upchae_names))
        else:
            # 종목 검색으로 내용이 없다면->업체검색으로
            for ad in ['', '㈜', '(주)']:
                self.driver.get(
                    '/?query={}'.format(self.sc_type_strip(upmoo_name)))
                src = BeautifulSoup(self.driver.page_source, 'lxml')

                # 이것도 안나오면 앞에 첨자 붙여서 검색
                if len(re.findall(r'K[a-zA-Z0-9]+', src.text)) != 0:

                    upchae_names = [self.utf_decoder(i[2:].replace('"', '')) for i in re.findall(
                        r'\:\s[?\s\\a-zA-Z0-9\(\)]+"', src.text)]
                    subj_codes = re.findall(r'K[a-zA-Z0-9]+', src.text)
                    print('시스템>>> 업체검색 upchae_names 는 {} 입니다'.format(upchae_names))
                    break

                # 정보 없다면
                else:
                    if ad == '(주)':
                        print(
                            '시스템>>> {} 은 신규 입력이 필요하기에 None을 반환합니다'.format(upmoo_name))
                        return [None, None, None]

                    continue

        # 구한 후보군들 거리 쟤서 결과 반환
        levi_dist = pd.Series([jamo_levenshtein(self.sc_type_strip(
            upmoo_name), self.sc_type_strip(i)) for i in upchae_names])

        # find nearest
        _idx = levi_dist[levi_dist == levi_dist.min()].index[0]

        print('시스템>>> 종목명 : {}, 업체코드 : {}, 거리 : {}'.format(
            upchae_names[_idx], subj_codes[_idx], levi_dist[_idx]))

        if dist:
            return [upchae_names[_idx], subj_codes[_idx], levi_dist[_idx]]

        else:
            return [upchae_names[_idx], subj_codes[_idx]]

    def get_upmoo_series(self):
        """
        #dir을 parameter 로 받을 수 있도록 추가해야함...

        self.upmoo_names 에 저장
        """
        # 엑셀에서 가져오기
        # my_xl = input('단일 고객사의 xlsx 업무배분 파일 주소를 알려주세요')
        # 간이로 수정함... 단일 고객사 단일 담당자 dir 가져오는걸로
        my_xl = self.upmoo_dir

        df0 = pd.read_excel(r'{}'.format(my_xl))

        upmoo_names = df0.loc[:, '종목명'].drop_duplicates().apply(
            self.sc_type_strip)
        upmoo_names.index = range(len(upmoo_names))

        return upmoo_names

    # 3계층

    def utf_decoder(self, captured):
        return captured.encode('utf-8').decode('unicode_escape')

    # def basic_strip(self, with_zoo, space=True):
    #     result = re.sub(r'\(.+?\)','',with_zoo.replace('㈜','')).strip()
    #     if space:
    #         return result
    #     else:
    #         return result.replace('\s','')

    def sc_type_strip(self, name):
        """
        strip ['RCPS','BW','CB','CPS','보통주','우선주','평가제외','제외']
        + re.sub(r'\(.+?\)','',name.replace('㈜','')).strip()
        """

        for i in ['RCPS', 'BW', 'CB', 'CPS', '보통주', '우선주', '상환', '주식회사', '전환', '교환권', '교환', '교부', '채권', '평가제외', '제외']:
            name = name.replace(i, '')
            name = name.strip()

        name = re.sub(r'\(.+?\)', '', name.replace('㈜', '')).strip()
        return name
