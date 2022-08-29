from selenium import webdriver as wd
import time

# 3개 모듈세트로 자주 사용
from selenium.webdriver.common.by import By
# 브라우저 로딩대기
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# chrome 옵션 
from selenium.webdriver.chrome.options import Options

# BeautifulSoup
from bs4 import BeautifulSoup

#pip install xlsxwriter
import xlsxwriter as xw
# CSV
import csv
# pandas
import pandas as pd

#이미지 바이트 처리하기 위한 패키지
from io import BytesIO
import urllib.request as req

chrome_options = Options()

# headless 모드(브라우저를 실행하지 않는 모드)
chrome_options.add_argument('--headless')

#엑셀파일설정
workbook = xw.Workbook('./crawl_result.xlsx')


#워크시트 추가
worksheet = workbook.add_worksheet()

#이미지리스트
img_list = [] 

# webDriver 설정
browser = wd.Chrome('./chromeDriver/chromedriver.exe',options=chrome_options)

# 일반모드
browser = wd.Chrome('./chromeDriver/chromedriver.exe')
browser.implicitly_wait(3)
browser.set_window_size(1024, 768)
browser.get('http://prod.danawa.com/list/?cate=112758&15main_11_02')

# 페이지소스 확인하기
# print(f'first page contents : {browser.page_source}')
# css 선택자
# #dlMaker_simple button.btn_spec_view.btn_view_more
# WebDriverWait(browser, 2) : browser가 2초간 기다림
# EC.presence_of_element_located : 해당 요소가 자리잡을 때까지
# By.CSS_SELECTOR : css 선택자에 의해 선택된 요소
### 제조사 더보기
# 2초안에 로딩되지 않으면 종료
WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#dlMaker_simple button.btn_spec_view.btn_view_more'))).click()


### 원하는 노트북 제조사 선택
WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#selectMaker_simple_priceCompare_A > li:nth-child(17) > label'))).click()
select_company = input("제조사를 입력하세요:")


# 애플 제조사 페이지 확인하기
# print(f'second page contents : {browser.page_source}')

# 로딩 타임 주기
time.sleep(2)


#시작페이지
cur_page=1

#전체페이지수
page_all=6

#액셀의 첫번째 행 지정
ins_row =1

while cur_page<=page_all:
  # cur_page+1


  # BeautifulSoup 생성
  bs = BeautifulSoup(browser.page_source, 'lxml')

  # prod_list = bs.select('.prod_item.prod_layer')
  prod_list = bs.select('.prod_item.prod_layer:not(.product-pot)')

  # DataFrame으로 2D 컬럼 생성
  df = pd.DataFrame(
      columns = ['PNAME','PRICE','IMG','PIMAGE','PCATEGORY_FK','PCOMPANY','PCONTENT'])
  

  

  # print('맥북 리스트 : ', len(prod_list))

  print(f'============================현재 페이지 :{cur_page}============================')
  print()

  # 데이터 추출하기
  for prod in prod_list:
      # 상품명, 이미지 링크주소, 가격
      #dataframe = ['PNAME','PRICE','IMG','PIMAGE','PCONTENT','PCOMPANY']
      name = prod.select('p.prod_name a')[0].text.strip()
      price = prod.select('p.price_sect > a')[0].text.strip()
      content = prod.select('dd > div.spec_list ')[0].text.strip()


      # img_url ="http:"+prod.select('a.thumb_link > img')[0]['src']

      img_attr = prod.select('a.thumb_link > img')[0].get('data-original')
      img_src = prod.select('a.thumb_link > img')[0]['src']
      img_url = req.Request("http:" + (img_attr if img_attr else img_src))

      url = (str("http:" + (img_attr if img_attr else img_src)))

      
      img_list.append(url)
      
      print(img_list)

      



      # print(img_url)

      #수신후 바이트로 변환
      img_data = BytesIO(req.urlopen(img_url).read())

      print("url : ",url)


      # 엑셀 저장
      worksheet.write_row(0,0,df)
      worksheet.write(f'A{ins_row}',name)
      worksheet.write(f'B{ins_row}',price)
      # 이미지 저장               위치    이미지이름, 이미지데이터(딕셔너리)
      worksheet.insert_image(f"C{ins_row}",name,{'image_data':img_data})
      worksheet.write(f'D{ins_row}',str(ins_row-2)+'.jpg')  
      worksheet.write(f'E{ins_row}',str(100))  
      worksheet.write(f'F{ins_row}',select_company)  
      worksheet.write(f'G{ins_row}',content)  
      ins_row +=1
      

  print()

  # 페이지 캡처
  browser.save_screenshot(f'./capture_page{cur_page}.png')

  #페이지 증가
  cur_page += 1

  if cur_page>page_all:
    print("성공")
    break

  #
  WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'.number_wrap a:nth-child({cur_page})'))).click()

  # BeautifulSoup 겍체삭제
  del bs
  #3초간 대기
  time.sleep(3)

file_no = 0
for j in range(0,len(img_list)) :
    try :
      # path = "C:\\CrawlForShop\\crawl_img\\"
      path = './crawl_img/'
      # req.urlretrieve(img_list[j],str(file_no)+'.jpg')
      req.urlretrieve(img_list[j],path+(str(file_no)+'.jpg'))
      file_no += 1
      time.sleep(0.5)  
    except :
      continue
  
browser.close()
workbook.close()

