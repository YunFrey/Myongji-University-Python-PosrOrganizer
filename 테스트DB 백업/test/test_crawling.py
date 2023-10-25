from selenium import webdriver

driver = webdriver.Chrome('chromedriver')
driver.get("http://www.koreapost.go.kr/kpost/subIndex/131.do?pSiteIdx=125")
naver_login = driver.find_element_by_id("id")
naver_login.clear()
#키보드 입력
naver_login.send_keys("naver_id")

naver_login = driver.find_element_by_id("pw")
naver_login.clear()
#키보드 입력
naver_login.send_keys("naver_pw")


pd.read_csv = '/'