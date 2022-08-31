#ID저장 클래스#
class IDClass:
    def __init__(self):
        self.Id='howjung'
        self.Pw='COCO4241++'
        self.Em='5466hee@gmail.com'
        self.EmPw='howjung1++'


#이메일 클래스#
import smtplib
from email.mime.text import MIMEText

class Email:
    def __init__(self):
        self.from_email = ''
        self.to_email = ''
        self.subject = ''
        self.contents = ''
        self.m_Ids = IDClass()

    def send_mail(self):
        print('이메일 정보 설정')
        print(self.m_Ids.Id + ',' + self.m_Ids.Pw + ',' + self.m_Ids.Em + ',' + self.m_Ids.EmPw)
        msg = MIMEText(self.contents, _charset='euc-kr')
        msg['Subject'] = self.subject
        msg['From'] = self.from_email
        msg['To'] = self.to_email

        # login에 들어갈 아이디 비번은 판매자 이메일아이디 비번
        print('이메일 전송 시작')
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.m_Ids.Em, self.m_Ids.EmPw)
        server.sendmail(self.from_email, self.to_email, msg.as_string())
        server.quit()


#자동화 시스템#
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.common.alert import Alert
import xlwings as xw
from selenium.common.exceptions import NoAlertPresentException
from keyboard import press

# SalesID 에 자신의 스마트스토어 아이디, 비번(네이버 기준) 입력하고 이메일에는 판매자에게 보낼 발송 메일 발신인으로 사용할 이메일 주소를 입력한다
def SmartStoreAuto(Id='', PW='', EmID='', EmPW=''):
    driver = webdriver.Chrome('./chromedriver')
    wb = xw.Book('스토어 재고 관리.xlsm')
    #wb = load_workbook('스토어 재고 관리.xlsm', read_only=False, keep_vba=True)
    ws = wb.sheets['Items']
    # 재고 없으면 0, 있으면 1
    StockState = 0
    title = ''
    SendMessage = ''
    m_email = Email()
    m_Ids = IDClass()
    # 몇번째 주문인지
    PayNum = 0
    # 행 위치 체크
    rownum = 3
    # 이미 재고 없음을 확인했는지 여부 / 0이면 확인 안함
    StockEmpty = 0
    # 재고 체크용 변수
    LastItemRow = 0
    # 모든 주문이 결제완료 상태가 아닌지 확인
    CheckOrder = 0

    #외부 변수 적용
    if Id != '':
        m_Ids.Id = Id
    if PW != '':
        m_Ids.Pw = PW
    if EmID != '':
        m_Ids.Em = EmID
    if EmPW != '':
        m_Ids.EmPw = EmPW

    try:
        # 스토어 사이트 접속 후 로그인
        driver.get('https://sell.smartstore.naver.com/#/home/dashboard')

        time.sleep(1)

        elem = driver.find_element_by_xpath("//em[text()='로그인하기']/../..")
        elem.click()

        time.sleep(2)

        elem = driver.find_element_by_link_text("네이버 아이디 로그인")
        elem.click()

        time.sleep(1)

        elem = driver.find_element_by_id("id")
        elem.send_keys(m_Ids.Id)
        elem = driver.find_element_by_id("pw")
        elem.send_keys(m_Ids.Pw)
        elem.send_keys(Keys.RETURN)

        time.sleep(3)

        while True:
            # 주문 조회 접속
            print('시작')
            elem = driver.find_element_by_xpath("//ul[@class='metisMenu']/li[2]")
            a = elem.find_element_by_xpath("./a")
            if(a.get_attribute("aria-expanded") == "false"):
                elem.click()
            time.sleep(1)

            print('주문 조회 클릭')
            elem = elem.find_element_by_xpath("./ul/li")
            elem.click()
            time.sleep(1)

            # 주문 조회 페이지 읽기
            print('프레임으로')
            iframe = driver.find_element_by_id('__naverpay')
            driver.switch_to_frame(iframe)

            frame = driver.find_element_by_xpath("//div[@class='npay_grid_area']")
            tbody = frame.find_element_by_xpath("./div/div/div[2]/div[2]/table/tbody")
            trs = tbody.find_elements_by_xpath("./tr")

            # 주문이 없다면 5분 대기 후 처음부터
            if len(trs) == 0:
                print('주문 없음')
                driver.switch_to_default_content()
                time.sleep(300)
                continue

            # 주문이 있다면 현재 주문 상태 확인
            for tr in trs:
                PayNum = PayNum + 1
                print(str(PayNum) + '번째 주문')
                td = tr.find_element_by_xpath("./td[2]")
                stat = td.get_attribute('title')
                print('주문자 정보 확인')

                # 결제완료 상태인 주문 상품명과 주문자의 이메일 주소, 주문자 이름을 가져옴
                if stat == "결제완료":
                    if CheckOrder == 0:
                        CheckOrder = 1
                    titletd = tr.find_element_by_xpath("./td[4]")
                    title = titletd.get_attribute('title')
                    emailtd = tr.find_element_by_xpath("./td[5]")
                    emailfull = emailtd.get_attribute('title')
                    email = emailfull[22:]
                    Nametd = tr.find_element_by_xpath("./td[7]")
                    Name = Nametd.get_attribute('title')
                    print('주문자 정보 가져옴')
                else:
                    print('결제완료 상태가 아님')
                    continue

                # 상품 재고가 있는지 확인
                if ws.range('C3').value == '' or ws.range('C3').value == None:
                    print('재고 없음')
                    break

                # 상품 재고가 있다면 상품명이 같고 사용하지 않은 상품의 시리얼 코드를 가져온다
                while True:
                    print('액셀체크 시작')
                    ItemName = ws.range('C' + str(rownum)).value
                    ItemStock = ws.range('E' + str(rownum)).value
                    Key = ws.range('D' + str(rownum)).value
                    print(ItemName)

                    if ItemName == '' or ItemName == None:
                        print('재고 모두 체크함')
                        if ws.range('G' + str(LastItemRow)).value == '' or ws.range('G' + str(LastItemRow)).value == None:
                            if LastItemRow == 0:
                                print('해당 아이템이 없음')
                                ws.range('G' + str(rownum)).value = '본 제품 재고 없음'
                                ws.range('H' + str(rownum)).value = Name
                                wb.save()
                            else:
                                print('해당 아이템은 있지만 재고가 없음')
                                ws.range('G' + str(LastItemRow)).value = '본 제품 재고 없음'
                                ws.range('H' + str(LastItemRow)).value = Name
                                wb.save()
                        else:
                            if LastItemRow == 0:
                                print('해당 아이템 없으며 이미 한번 이상 체크한 재고')
                                CheckString = ws.range('H' + str(rownum)).value
                                CheckStrings = str(CheckString).split(',')
                                for string in CheckStrings:
                                    if Name == string:
                                        print('이미 체크한 사용자')
                                        StockEmpty = 1
                                        break
                                if StockEmpty == 0:
                                    print('이름 추가')
                                    ws.range('H' + str(rownum)).value = str(CheckString) + ',' + Name
                                    wb.save()
                            else:
                                CheckString = ws.range('H' + str(LastItemRow)).value
                                CheckStrings = str(CheckString).split(',')
                                for string in CheckStrings:
                                    if Name == string:
                                        print('이미 체크한 사용자')
                                        StockEmpty = 1
                                        break
                                if StockEmpty == 0:
                                    print('이름 추가')
                                    ws.range('H' + str(LastItemRow)).value = str(CheckString) + ',' + Name
                                    wb.save()

                        rownum = 3
                        break             

                    if ItemName == title:
                        print('재고 확인')
                        LastItemRow = rownum
                        # 재고가 있다면 재고 상태를 사용으로 변경
                        if ItemStock != '사용':
                            print('재고 정보 가져오기')
                            if ItemStock == 1 or ItemStock == '1':
                                print('재고가 한개')
                                now = time.localtime()
                                CurrentTime = '%04d-%02d-%02d %02d:%02d:%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
                                ws.range('E' + str(rownum)).value = '사용'
                                ws.range('F' + str(rownum)).value = Name + ' / ' + email + ' / ' + CurrentTime
                            else:
                                print('재고가 여러개')
                                ws.range('E' + str(rownum)).value = str(int(ItemStock) - 1)
                            wb.save()
                            StockState = 1
                            SendMessage = '제품명 : ' + title + '/n' + '시리얼 키 : ' + str(Key)
                            print(SendMessage)
                            rownum = 3
                            break
                        else:
                            print('해당 상품 재고 없음')
                            rownum = rownum + 1
                            continue
                    else:
                        print('해당 상품이 아님')
                        rownum = rownum + 1
                        continue
                
                # 재고가 있다면 재고를 이메일로 전송하고 없다면 판매자에게 재고 추가할 것을 알림
                if StockState == 0:
                    if StockEmpty == 1:
                        print('이미 지연 메일 보냄')
                        #변수 초기화
                        StockState = 0
                        PayNum = 0
                        break

                    print('구매자에게 발송')
                    m_email.from_email = m_Ids.Em
                    m_email.to_email = email
                    m_email.contents = '24시간 안에 해당 제품의 재고를 확보하여 발송해 드리겠습니다'
                    m_email.subject = '발송 지연 안내'
                    m_email.send_mail()

                    print('판매자에게 발송')
                    m_email.from_email = m_Ids.Em
                    m_email.to_email = m_Ids.Em
                    m_email.contents = '재고가 없으니 확보하여 Excel에 업데이트해 주세요'
                    m_email.subject = '재고 부족 안내'
                    m_email.send_mail()

                    #변수 초기화
                    StockState = 0
                    PayNum = 0
                    break

                else:
                    print('제품을 배송')
                    m_email.from_email = m_Ids.Em
                    m_email.to_email = email
                    m_email.contents = SendMessage
                    m_email.subject = '상품 배송'
                    m_email.send_mail()

                    # 배송 상태 스토어에 반영
                    print('발송 상태 반영')
                    frame = driver.find_element_by_xpath("//div[@class='npay_grid_area']")
                    tbody = frame.find_element_by_xpath("./div/div[2]/div[2]/div[2]/table/tbody")
                    tr = tbody.find_element_by_xpath("./tr[" + str(PayNum) + "]")

                    elem = tr.find_element_by_xpath('./td/input')
                    elem.click()

                    time.sleep(1)

                    elem = driver.find_element_by_id('_link_dispatch')
                    elem.click()

                    time.sleep(1)

                    #발송 처리
                    print('발송 처리')
                    btn = driver.find_element_by_xpath("//div[@class='npay_grid_area htmlx_grid_container']/div/div/div[2]/table/tbody/tr[2]/td")
                    btn.click()

                    time.sleep(1)

                    #정보 입력
                    print('배송방법 입력')
                    elem = driver.find_element_by_xpath("//div[@class='npay_grid_area htmlx_grid_container']/div/div/div[2]/table/tbody/tr[2]/td[4]/select/option[5]")
                    elem.click()

                    #발송처리 버튼 클릭
                    print('발송처리 버튼 클릭')
                    elem = driver.find_element_by_xpath("//div[@class='npay_button_major']/div/button[2]")
                    elem.click()

                    time.sleep(1)

                    #알람 창 처리
                    print('알람 창 처리')
                    driver.switch_to.alert.accept()
                    time.sleep(3)

                    #알람 창 처리 시도
                    print('알람 창 끄기')
                    press('enter')

                    time.sleep(2)

                    #바깥 프레임으로
                    print('발송완료')
                    driver.switch_to.default_content()

                    #변수 초기화
                    StockState = 0
                    PayNum = 0
                    break

            StockState = 0
            PayNum = 0
            #모든 주문이 결제완료가 아니라면 5분 대기
            if CheckOrder == 0:
                print("모든 주문이 결제완료가 아님")
                time.sleep(300)
            else:
                CheckOrder = 0
            driver.switch_to.default_content()
            continue

    except Exception as e:
        print(e)
    finally:
        wb.app.quit()
        driver.quit()


#GUI 생성#
import tkinter

class GUIT():
    def __init__(self):
        self.tkhandler = tkinter.Tk()
        self.tkhandler.geometry('500x500')
        self.tkhandler.title('스마트스토어 자동화프로그램')

        self.label_title = tkinter.Label(self.tkhandler, text='자동화 프로그램')
        self.label_title.pack(pady=10)

        #ID프레임
        self.ID_frame = tkinter.Frame(self.tkhandler)
        self.ID_frame.pack(fill="x")

        self.label_ID = tkinter.Label(self.ID_frame, text='스토어 네이버 ID', width=15)
        self.label_ID.pack(side='left', padx=10, pady=5)

        self.text_ID = tkinter.Text(self.ID_frame, height=1)
        self.text_ID.pack(fill='x', padx=10, expand=True)
        #self.text_ID.insert('current', 'howjung')

        #IDPassword프레임
        self.IDPass_frame = tkinter.Frame(self.tkhandler)
        self.IDPass_frame.pack(fill="x")

        self.label_IDPass = tkinter.Label(self.IDPass_frame, text='스토어 네이버 PW', width=15)
        self.label_IDPass.pack(side='left', padx=10, pady=5)

        self.text_IDPass = tkinter.Text(self.IDPass_frame, height=1)
        self.text_IDPass.pack(fill='x', padx=10, expand=True)
        #self.text_IDPass.insert('current', 'COCO4241++')

        #Email프레임
        self.Em_frame = tkinter.Frame(self.tkhandler)
        self.Em_frame.pack(fill="x")

        self.label_Em = tkinter.Label(self.Em_frame, text='판매자 gmail', width=15)
        self.label_Em.pack(side='left', padx=10, pady=5)

        self.text_Em = tkinter.Text(self.Em_frame, height=1)
        self.text_Em.pack(fill='x', padx=10, expand=True)
        #self.text_Em.insert('current', '5466hee@gmail.com')

        #Email Password 프레임
        self.EmPW_frame = tkinter.Frame(self.tkhandler)
        self.EmPW_frame.pack(fill="x")

        self.label_EmPW = tkinter.Label(self.EmPW_frame, text='판매자 gmail PW', width=15)
        self.label_EmPW.pack(side='left', padx=10, pady=5)

        self.text_EmPW = tkinter.Text(self.EmPW_frame, height=1)
        self.text_EmPW.pack(fill='x', padx=10, expand=True)
        #self.text_EmPW.insert('current', 'howjung1++')

        #시작 버튼
        self.btn = tkinter.Button(self.tkhandler, text='스마트스토어 자동화 프로그램 시작', width=30, command=self.runAutoSystem)
        self.btn.pack(pady=10)

        #테스트용 라벨
        self.label_test = tkinter.Label(self.tkhandler, text='버튼 클릭 대기', width=30)
        self.label_test.pack(side='bottom')

    def runAutoSystem(self):
        self.label_test.config(text='자동화 시스템 시작')
        time.sleep(1)
        Id = self.text_ID.get('1.0', 'end').strip()
        PW = self.text_IDPass.get('1.0', 'end').strip()
        EmID = self.text_Em.get('1.0', 'end').strip()
        EmPW = self.text_EmPW.get('1.0', 'end').strip()
        SmartStoreAuto(Id, PW, EmID, EmPW)
        self.label_test.config(text="버튼 클릭 대기")

    def run(self):
        self.tkhandler.mainloop()

g = GUIT()
g.run()