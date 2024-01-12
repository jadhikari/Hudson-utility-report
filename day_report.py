from selenium import webdriver
import chromedriver_autoinstaller
import time
import re
import pandas as pd
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By

chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
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

s_number = []
f_data = []
s_data = []
dates = []
class MyClass:
    # Define a method inside the class
    def my_method(self):
        time.sleep(1)
        iframe = driver.find_element(By.XPATH, '//*[@id="mieruframe"]')
        driver.switch_to.frame(iframe)
        e = driver.find_element(By.XPATH, value='//*[@id="mi_title"]').text
        print(e)
        table = driver.find_element(By.XPATH, value='//*[@id="mi_month_list_table"]/tbody')
        # Extract data from the first and second columns

        first_column_data = []
        second_column_data = []
        # Iterate through rows and extract data from the first and second columns
        rows = table.find_elements(By.TAG_NAME, 'tr')
        for row in rows:
            columns = row.find_elements(By.TAG_NAME, 'td')
            if len(columns) >= 2:
                # Extract data from the first and second columns
                input_string = columns[0].text
                formatted_date = re.sub(r"(\d{2})月(\d{2})日", r"\1/\2", input_string)
                cleaned_date = re.sub(r'\([^)]*\)', '', formatted_date)
                first_column_data.append(cleaned_date)
                second_column_data.append(columns[1].text)
        f_data.append(first_column_data)
        s_data.append(second_column_data)

        # Print the extracted data


my_object = MyClass()

#change the first number in () according to the User ID. You can find the num in read_me
for x in range(60,0,-1):
    path = '//*[@id="contractInfo'
    path += str(x)
    path += ' "]/table/tbody/tr[4]/td[1]'
    print(x)
    td = driver.find_element(By.XPATH, value= path)
    number = td.text
    if "/" in number:
        button = '//*[@id="contractInfo'
        button += str(x)
        button += ' "]/table/tbody'
        time.sleep(1)
        element = driver.find_element(By.XPATH, value=button)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(4)
    else:
        s_number.append(number)
        button = '//*[@id="contractInfo'
        button += str(x)
        button += ' "]/table/tbody//button'
        element = driver.find_element(By.XPATH, value=button)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(3)
        driver.find_element(By.XPATH, value='//*[@id="menu_day"]').click()
        time.sleep(4)
        iframe = driver.find_element(By.XPATH, '//*[@id="mieruframe"]')
        driver.switch_to.frame(iframe)
        d = driver.find_element(By.XPATH, value='//*[@id="mi_title"]').text
        # change the date according to your need
        if '2023年11月分' in d:
            time.sleep(2)

            driver.switch_to.default_content()
            my_object.my_method()

        else:
            driver.find_element(By.XPATH, value='//*[@id="before_btn"]/img').click()
            driver.switch_to.default_content()
            time.sleep(3)
            my_object.my_method()

    driver.switch_to.default_content()
    driver.find_element(By.XPATH, value='//*[@id="contListBtn"]').click()
    time.sleep(2)


print("system id:", s_number)
print("First Column Data:", f_data)
print("Second Column Data:", s_data)
max_date = max(max(f_data, key=lambda x: datetime.strptime(x[-1], '%m/%d')))
date_range = pd.date_range(start=datetime.strptime(f_data[0][0], '%m/%d'), end=datetime.strptime(max_date, '%m/%d'))
df = pd.DataFrame(index=date_range.strftime('%m/%d'), columns=s_number)

# Fill the DataFrame with s_data
for i, s_data_item in enumerate(s_data):
    start_index = datetime.strptime(f_data[i][0], '%m/%d')  # Start index based on f_data
    for j, value in enumerate(s_data_item):
        df.at[(start_index + timedelta(days=j)).strftime('%m/%d'), s_number[i]] = value

# Sort the DataFrame by index (date)
df.sort_index(inplace=True)

# Reverse the order of columns
df = df.iloc[:, ::-1]

# Save the DataFrame to an Excel file
df.to_excel('output.xlsx')

