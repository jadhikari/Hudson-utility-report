from selenium import webdriver
import chromedriver_autoinstaller
import time
import xlsxwriter
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
chromedriver_autoinstaller.install()
# Check if the current version of chromedriver exists
# and if it doesn't exist, download it automatically,
# then add chromedriver to path

driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://www.k-nw.kyuden.co.jp/l-demand/pages/")
time.sleep(2)

#there are 4 user, you can chenge it according to your requerement.
user_id = "qden-jenson"
password = "qden-jenson"

id_name = driver.find_element(By.XPATH, '//*[@id="k_id"]')
id_name.send_keys(user_id)
id_pass = driver.find_element(By.XPATH,'//*[@id="k_pw"]')
id_pass.send_keys(password)

log_in = driver.find_element(By.XPATH, value='//*[@id="loginFormDtl"]/button')
log_in.click()
time.sleep(3)

driver.find_element(By.XPATH, value='//*[@id="globalNavPC"]/div/ul/li[2]/a').click()

driver.find_element(By.XPATH, value='//*[@id="globalNavPC"]/div/table[2]/tbody/tr[2]/td[2]/a').click()
time.sleep(3)

driver.find_element(By.XPATH, value='//*[@id="contListBtn"]').click()
time.sleep(3)

system_ids = []
date = []
sdate = []
edate = []
powers = []
purchases = []
payments = []
amounts = []
days = []
averages = []


for x in range(60,0,-1):
    print(x)
    path = '//*[@id="contractInfo'
    path += str(x)
    path += ' "]/table/tbody'
    element = driver.find_element(By.XPATH, value= path)
    driver.execute_script("arguments[0].click();", element)
    time.sleep(4)
    system_id = driver.find_element(By.XPATH, value='//*[@id="gaiyouKeiyakuInfo"]/table/tbody/tr[3]/td[2]').text
    system_id = system_id[1:]
    print(system_id)
    system_ids.append(system_id)
    driver.find_element(By.XPATH, value='//*[@id="menu_month"]').click()
    time.sleep(2)
    for i in range(2, 4):
        row ='//*[@id="wrapper"]/div/div/table[1]/tbody/tr['
        row += str(i)
        row += ']'
        d = row + '/td[1]'
        month = driver.find_element(By.XPATH, value=d).text
        if month == '2023/11':
            p = row + '/td[2]'
            power = driver.find_element(By.XPATH, value=p).text
            pur = row + '/td[3]'
            purchase = driver.find_element(By.XPATH, value=pur).text
            pa = row + '/td[4]'
            payment = driver.find_element(By.XPATH, value=pa).text
            am = row + '/td[5]'
            amount = driver.find_element(By.XPATH, value=am).text
            da = row + '/td[6]'
            day = driver.find_element(By.XPATH, value=da).text
            av = row + '/td[7]'
            average = driver.find_element(By.XPATH, value=av).text

            date.append(month)
            powers.append(power)
            purchases.append(purchase)
            payments.append(payment)
            amounts.append(amount)
            days.append(day)
            averages.append(average)

    driver.find_element(By.XPATH, value='//*[@id="menu_kenshin"]').click()
    time.sleep(3)
    #driver.find_element(By.XPATH, value='/html/body/div/div[2]/div[2]/div[5]/div[1]/div/select').click()
    dropdown_xpath = '//*[@id="kenshinKounyuSelect"]/select'
    dropdown_element = driver.find_element(By.XPATH, value=dropdown_xpath)
    dropdown = Select(dropdown_element)
    print(dropdown)

    # change the date according to your need
    desired_option = '2023年11月分'

    # Check if the desired option is already selected
    dropdown.select_by_visible_text(desired_option)
    time.sleep(3)
    se_date = driver.find_element(By.XPATH,
                                      value='//*[@id="kounyu_container_inner"]/div[2]/table[2]/tbody/tr[1]/td[1]').text
    print(se_date)
    b, c = se_date.split('～')
    sdate.append(b)
    edate.append(c)

    driver.find_element(By.XPATH, value='//*[@id="contListBtn"]').click()
    time.sleep(2)

result = []
result.append(system_ids)
result.append(date)
result.append(sdate)
result.append(edate)
result.append(powers)
result.append(purchases)
result.append(payments)
result.append(amounts)
result.append(days)
result.append(averages)

print(result)

workbook = xlsxwriter.Workbook('result.xlsx')
worksheet1 = workbook.add_worksheet('Sheet1')

row = 1
for col, data in enumerate(result):
    worksheet1.write_column(row, col, data)

workbook.close()


