import time
from selenium import webdriver
from readexcel import ReadExcel
# broswer = webdriver.Chrome()
from selenium.webdriver.common.by import By

driver = webdriver.Edge()
driver.implicitly_wait(10)
driver.get("https://opensea.io/collection/perdidos-no-tempo")
driver.maximize_window()
# 获取当前页面的 href链接列表
# //a[contains(@href,"/assets/matic/0xecc82095b2e23605cd95552d90216faa87606c40/")]
items = driver.find_elements(by=By.XPATH,
                             value='//a[contains(@href,"/assets/matic/0xecc82095b2e23605cd95552d90216faa87606c40/")]')
i = 2
links = []
for l in items:
    link = l.get_attribute("href")
    # 将获取到的link写入excel
    ReadExcel(r"t.xlsx", "Sheet1").write_date(i, 1, i - 1)
    ReadExcel(r"t.xlsx", "Sheet1").write_date(i, 2, link)
    links.append(link)
    i += 1
# 直接加载页面 ，一个一个点击
i = 2
for link in links:
    driver.get(link)
    a = driver.find_element(by=By.XPATH, value='//i[@value="refresh"]/../..')
    a.click()
    print("点击刷新")
    ReadExcel(r"t.xlsx", "Sheet1").write_date(i, 3, "Click")
    time.sleep(1)  # 用显性等待会好很多
    text = driver.find_element(by=By.XPATH, value='//div[@class="sc-1xf18x6-0 cKmASc"]/../../..').text
    if "queued" in text:
        ReadExcel(r"t.xlsx", "Sheet1").write_date(i, 3, "Queued")
    else:
        ReadExcel(r"t.xlsx", "Sheet1").write_date(i, 3, "Error")
    i += 1
driver.quit()
# =============================================
# 自动下拉滚动
# js = 'return document.body.scrollHeight;'
# height = 0
# while True:
#     new_height = driver.execute_script(js)
#     if new_height > height:
#         driver.execute_script('window.scrollTo(0, document.body.scrollHeight)')
#         height = new_height
#         time.sleep(5)
#         break
#     else:
#         print("滚动条已经处于页面最下方!")
#         driver.execute_script('window.scrollTo(0, 0)')  # 页面滚动到顶部
#         break
