<h1><center> <font> 자료 수집 자동화 </font> </center></h1>

#### - 해당 repository는 금융자산 평가 회사 재직 당시 제작한 private repository를 토대로 민감한 정보를 모두 제거하여 작성하였습니다.

#### - 또한 개인적으로 진행한 첫 프로젝트인만큼 많은 비효율과 에러가 내재되어 있지만, (ex: promise 사용하지 않고 충분한 시간을 sleep으로 잡아줌)

#### 실제로 재직 당시 많은 시간을 줄여주었던 패키지입니다.

<br>

## 1. 파일 별 설명

- **main_run** : 아래 class 들을 이용하여 전체 작업을 진행
- **path_finder** : 전기 자료 유무 확인 및 유사 자료와 동명 회사 구분
  - `(soynlp, os, shutil, re)`
- **dart** : 회사 평가 외부 자료인 dart 에서 공시 자료 parsing 및 전처리
  - `(selenium, bs4, pandas)`
- **kis_chrome** : 내부 db 자료 parsing 및 전처리
  - `(selenium, bs4, pandas)`
- **to_excel** : 취합 자료들 하나의 엑셀 파일에 기입 및 정리
  - `(openpyxl)`
- **post_office** : 필요한 directory에 파일 저장 - `(soynlp, os, shutil, re)`
  <br>

## 2. 작동 구상

#### action

```
    담당 종목 조회
        # 1. 업무배분 -> 담당자 입력 -> 검색 -> 펀드 아이디 복사

	패스 다루기
		# 1. 전기자료
			# 없으면 없다고 나와야함

		# 2. 현재자료
			# 이미 만들어진게 있을듯

	Dart 정보
		# 1. 종목 이름 가져오기
			업체이름 검색 되는것(전기 엑셀 참조)

		# 2-1. 거래소 공시 제목/link
			전환청구권 행사 parsing

		# 2-2. 지분공시 제목/link
			# filter 목록 만들어야함
			세부변동내역 parsing

		# 2-3. 주요사항 보고/link
			# 발행정보

    DB 자료 조회
        # 1. 펀드 아이디 기입 -> 펀드 이력, 종목 history 조회 -> 전기 평가값 조회

	엑셀 기입
		# 1. 자료 읽기
			# 각 action별로 필요한 정보 부분 찾기

		# 2. 자료 쓰기
			# 각 action별로 필요한 정보 부분 찾기

		# 유효성 검사
			# 필요할때마다

```

## 3. 세부 작업 내용

### **Dart**

1. 기간 : 발행일 10일 전 이후부터
2. 필요 자료 uri : 주요사항보고(주요사항보고서)/발행공시(전체)/지분공시(주식등의대량보유상황보고서)/거래소공시(전체)
3. **parsing keyword** :

- '발행결정'

  - '회차','종류','권면총액','표면이자율','만기이자율','사채만기일','전환비율','전환가액','전환청구기간','납입일'
  - '전환에 관한 사항','옵션에 관한 사항','기타 투자판단에 참고할 사항'
  - '조정'
    - '회차','조정전 전환가액','조정후 전환가액','조정사유','조정가액 적용일'
  - '행사'
    - '일별 전환청구내역'
      - '청구일자','회차','전환가액','발행한 주식수'
    - '전환사채 잔액'
      - '회차','발행당시 사채의 권면','신고일 현재 미전환','전환가액','전환가능 주식수'
  - '대량보유상황'(제출인 확인)
    - '보고사유'
    - '세부변동내역'
      - '변동일','취득/처분방법','주식등의종류','변동 내역','취득/처분단가'
  - '증자'

  - '분할'
    - '주식분할 결정'
      1주당 가액/ 분할 전 후

### **전기 자료 및 동명기업 구분**

1. 유사도 distance(단어1,단어2,알림)

- 단어가 비는건 괜찮은데 틀리는건 안됨
- 주식회사,형,인프라,특별,자산,창업,벤처,PEF,pef,괄호 안에 말들,사모,투자, 합자,회사,조합,전문 - 띄어쓰기로 바꿈
  업체 검색시는 BW, CB, TRS, RCPS, 등등도 빼야함
- 공백은 매치 안함
- 숫자 다르면 무조건 다름(1호 2호)
- 100% 나오면 묻지 말고 ~70%는 물어봐
- 검색할때는 형식어와 공백 빼고 첫 단어 검색 후, 목록 다 받아온 다음 유사도 검사
- return(동일 여부, score)

- 고객사 폴더 찾기
  - 명단 관리
    - 평가일 찾기 -> 21-04-30, [만약 분기가 되었다면 3,6,9,12 :21 1분기, 21 1Q]
    - 평가내역 폴더 찾기(평가일 있었으면 무조건 있음)
    - 내 이름 찾기(내이름 없으면 무조건 만듬)
    - 종목명 찾기(없으면 무조건 만들기)
    - 자료수령, 평가내역 폴더 만들기
    - 전기자료 찾아서 바로가기 만들어 놓기(종목명 매칭 못하면 로그로 남겨놓기)
  - 파일 생성은 결국 마지막에 하기때문에 원자성 유지

## 4. 고민거리

- assigning instance variable while get...

  - should get only just 'get'
    1. self.variable = result
    2. results['query'] = result
       to obtain latest result... would it be unconventional?

- object oriented...

  - task 하나당 객체를 하나씩 만들수도 있고...
    - 생각없이 쓰기 편함
    - list로 기록 남길 생각을 안하고 모듈을 짤 수 있음
  - 객체의 method 하나를 task에 맞게 만들어서 for loop 안에 쓸수도 있고...
    - 기록 남기기 편함

- 버전관리 conda 때문에 안되는거면

  - executer를 python 으로 바꾸던지
  - conda dos 에서 조작하던지

- to_excel.py

  - 에서 객체로 쓸거는 class(객체 고유 variable)
  - 에서 함수로 쓸거는 따로 빼서(필요한 argument)
    - 이렇게 만들어야 다른 .py 파일에서도 import 해서 함수 쓰기 편함....
    - 함수 따로 안빼놓으면 self 걸려있어서 못씀 ㅠㅠ
  - 이것보단 class 만드는건 무조건인데... init이나 self 안들어 있게 짜는게 나을듯
  - class 안만들고 만드니까 instance 생성 안해도 되고 너무 좋은데?

- from mtp_xl_to_ss import google_ss
  - 함수만으로 만든 .py 파일 들고올때 함수 아닌것들은 미리 불러와짐
  - 이경우 chrome driver가 import 할때 미리 켜짐

## 5. 미구현 아이디어

### 멀티플 가치 평가 방식 Range(seudo code)

```

멀티플 경우의 수 2^n*2^4*5

	컴셋 n
		산업코드 기준
		주력상품 기준

	지표 4 (all given)
		PBR 시총/장부가
		EV/EBITDA (시총+순차입금-비영업용자산 := 영업용자산의 평가가치)/영업이익
			EV/EBIT 우린 감가상각까지 반영하겠다
		EV/SALES 영업용자산의 평가가치/매출

	적용배수 5
		평균 mean
		가중평균 vector inner product/n
		median median
		min min
		max max


lst = list(map(list, itertools.product([0, 1], repeat=n)))
# OR
lst = [list(i) for i in itertools.product([0, 1], repeat=n)]

<노트에 정리됨>
([pbr1,pbr2,..pbrn],
[EV_EBITDA1,EV_EBITDA2,...EV_EBITDAn,],
[EV_SALES1,EV_SALES2,...EV_SALESn])

*

lst_(2^n)

*

[mean,promean,median,min,max]

*

lst_(n=3)


All possible MTP

lst = [list(i) for i in itertools.product([0, 1], repeat=n)][1:]

[0,1,1,1,0,0] 의 경우
->	[[0,1,0,0,0,0],
	[0,0,1,0,0,0],
	[0,0,0,1,0,0]]

x1 = [0,1,1,1,0,0]
I1 = []

for i in range(sum(x)):
	result = [0,0,0,0,0,0]	#compset 갯수와 같음
	for j in range(len(x1)):
		if x1[j] == 1:
			x1[j] = 0
			result[j] = 1
			I1.append(result)

I2 = []
x2 = [0,1,1,1,0]
for i in range(sum(x)):
	result = [0,0,0,0,0]		#무조건 5개
	for j in range(len(x2)):
		if x2[j] == 1:
			x2[j] = 0
			result[j] = 1
			I1.append(result)

#openpyxl 에서 가져와
Comp

result1 = [I1@Comp@I2.apply(mean),
	I1@Comp@I2.apply(weightedmean),
	I1@Comp@I2.apply(median),
	I1@Comp@I2.apply(min),
	I1@Comp@I2.apply(max),]

Blank = [[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],]

for i in range(len(My)):
	Blank[i,i] = My[i]

My = Blank

EV_amend = [[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],
	[0,0,0,0,0],]
	#무조건 (6,5)

x3 = -현금성자산+이자부부채
EV_amend[3,:] = [x3,x3,x3]

result2 = (I1@My@I1.T)@result1 + I1@EV_amend

J = np.ones((k,), dtype=int)
result3 = (1/k)*J.T@result2

```

### 회계정보 무결성 체킹

```

재무제표 장부가 긁어오기(https://pbpython.com/pandas-html-table.html)
	pef case
		1. classifier 학습?
			strip으로 one hot encoding


		1-1. non dart parser

		1-2. 재무상태표 업데이트(재무기준일)
			투자자산 현황에서
				신규 취득 (주식수*취득단가[투자금액])
					당좌자산 -
					투자자산 +

				회수(주식수*처분단가)
					당좌자산 +
					투자자산 -
					이익잉여금 +

			이후인 셀(2021-03-31) 잡아서 아래 합계값
			출자금에서
				원본분배
					자본금 -
					당좌자산 -
				이익분배
					이익잉여금 -
					당좌자산 -

				신규출자
					자본금 +
					당좌자산 +

		1-3. 투자주식수 업데이트 숫자 맞는지 확인(투자자산 현황)
			(최초)최초 투자금액/취득단가 = 투자주식수
				if BS상장부가액 != 최초투자금액 ( 알림 ' 변동 존재 ' )

				else:
					if (기준일)BS상장부가액/기준일 잔여주식수 != 취득단가:
						알림 '액면분할 가능성!'




		2. rule based
			- 우선적으로 세부 계정의 합이 총계와 맞는지
				세부 계정들 분류 다 tab으로 맞춰줘야 parse 가능
					세부항목과 y축선에 있는 자료들 list로 객체 취급

			앞 부분 tab 숫자중 오름차순 아닌것만(Dart parser)
				보통예금,기타단기금융상품,미수수익,미수금,선납세금
					'예금'과 '단기금융' 은 당좌
					'투자자산' 하위, '증권'은 다 투자자산
				재고자산
				매도가능증권
				기타보증금

			부채 기준으로 잘라서 기타 들어가는거 다 기타자산에 기입
				부채는 부채총계가 부채지


			손익계산서와 연결 가능성
				미처분이익잉여금

			투자현황과 연결
				투자자산 = sum(bs상 장부가액) if prodsum(투자주식수*취득단가)

			출자금 현황과 연결
				원본/출자금
				재무상 원본 == 원본 + 추가출자 - 원본 분배
				차액이 이익잉여금

				출자,합계 셀의 값이 자본금 값과 같은지
					나머지는 자본 잉여금으로 들어감

```
