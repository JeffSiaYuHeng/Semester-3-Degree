import time
import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

# 读取 Excel 文件中的账号信息
df = pd.read_excel("credentials.xlsx")  # 确保 Excel 文件包含 `username` 和 `password` 列

# 设置 WebDriver（Chrome）
driver = webdriver.Chrome()

for index, row in df.iterrows():
    username_value = str(row['username']).zfill(8)  # 确保用户名是字符串
    password_value = str(row['password']).zfill(8)  # 确保密码长度为8位，前面补0

    print(username_value)
    print("p" , password_value)

    try:
        # **1. 打开登录页面**
        driver.get("http://cms.nilai.edu.my/psp/HRCSX/?cmd=login&languageCd=UKE&")  # 替换成你的 URL
        wait = WebDriverWait(driver, 10)

        time.sleep(3)

        try:
            # Wait until "Sign in to PeopleSoft" button is clickable and click it
            wait = WebDriverWait(driver, 10)
            sign_in_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//a[contains(@href, "cmd=login") and contains(text(), "Sign in to PeopleSoft")]')))
            sign_in_button.click()

            # Wait 10 seconds before clicking the English button
            time.sleep(5)

            # Wait until the "English" button is clickable and click it
            english_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@title="English"]')))
            english_button.click()

        except Exception as e:
            print(f"Initialization Error: {e}")

        try:
            # **4. Locate Username & Password Fields**
            username = wait.until(EC.presence_of_element_located((By.NAME, "userid")))
            password = driver.find_element(By.NAME, "pwd")

            # **5. Enter Credentials**
            username.send_keys(username_value)
            password.send_keys(password_value)

            # **6. Click Login Button**
            login_button = driver.find_element(By.XPATH, '//input[@type="submit"]')  # Modify if needed
            login_button.click()

        except Exception as e:
            print(f"⚠️ Input Box Error: {e}")



        time.sleep(5)

        # **4. 进入 Instructor Evaluation**
        wait.until(EC.element_to_be_clickable((By.ID, "fldra_L_EXTENSIONS"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "fldra_EVALUATIONS"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "fldra_L_STUDENT_EVAL"))).click()
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Online Faculty Evaluation"))).click()

        time.sleep(10)

        try:
            # ✅ Switch to the iframe
            iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ptifrmtgtframe"))  # Ensure correct ID
            )
            driver.switch_to.frame(iframe)
            print("✅ Switched to iframe")

            while True:
                # ✅ Locate the table
                table = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'PSLEVEL1GRIDWBO')]")
                                                   ))
                print("✅ Table found!")

                # ✅ Find all rows in the table
                rows = table.find_elements(By.TAG_NAME, "tr")

                if not rows:
                    print("⚠️ No rows found in the table.")
                    break  # Exit loop if no rows are found

                row_found = False
                index = 0

                while index < len(rows):
                    row = rows[index]
                    cells = row.find_elements(By.TAG_NAME, "td")

                    if len(cells) >= 2:
                        second_td = cells[1]
                        try:
                            # ✅ Try to click the 'Select Class' button
                            select_link = WebDriverWait(second_td, 5).until(
                                EC.element_to_be_clickable((By.TAG_NAME, "a"))
                            )
                            print(f"✅ Found 'Select Class' button in row {index + 1}: {select_link.text}")

                            # 🎯 Click the button
                            select_link.click()
                            print("🎯 Clicked 'Select Class'")

                            # ✅ Wait for the next page to load
                            time.sleep(20)  # Adjust this if needed

                            import random

                            # ✅ Fill Dropdowns (Ratings)
                            try:
                                for i in range(12):  # Loop from 0 to 11
                                    select_id = f"L_INST_EVAL_TBL_L_QUESTION_RATING${i}"
                                    try:
                                        dropdown = WebDriverWait(driver, 5).until(
                                            EC.presence_of_element_located((By.ID, select_id))
                                        )

                                        random_choice = random.choice(["ST", "GD", "EX"])
                                        Select(dropdown).select_by_value(random_choice)
                                        print(f"✅ Selected '{random_choice}' for: {select_id}")
                                    except Exception as e:
                                        print(f"⚠️ Could not select option for {select_id}: {e}")
                            except Exception as e:
                                print("❌ Error processing dropdowns:", e)

                            # ✅ Fill Textareas (Comments)
                            try:
                                for i in range(12, 15):  # Fill textareas from ID 12 to 14
                                    textarea_id = f"L_INST_EVAL_TBL_COMMENTS${i}"
                                    try:
                                        textarea = WebDriverWait(driver, 5).until(
                                            EC.presence_of_element_located((By.ID, textarea_id))
                                        )
                                        textarea.clear()
                                        textarea.send_keys(f"This is my feedback for question {i}.")
                                        print(f"✅ Filled textarea {textarea_id}")
                                    except Exception as e:
                                        print(f"⚠️ Could not fill {textarea_id}: {e}")
                            except Exception as e:
                                print("❌ Error processing textareas:", e)

                            time.sleep(20)
                            row_found = True
                            break  # Stop after clicking the first available row
                        except Exception as e:
                            print(f"⚠️ Unable to click the link in row {index + 1}: {e}")
                            index += 1  # Move to the next row

                if row_found:
                    break  # Exit the while loop if a row was processed

        except Exception as e:
            print("❌ Error locating the table or rows:", e)

        # Close the driver after execution
        driver.quit()


    #             # **8. 提交表单**
    #             submit_button = driver.find_element(By.XPATH, "//button[contains(text(),'Submit')]")
    #             submit_button.click()
    #             time.sleep(2)  # 等待提交完成
    #

        print(f"✅ 账户 {username_value} 处理完成")
    except Exception as e:
        print(f"❌ 账户 {username_value} 发生错误: {e}")

    # **9. 退出当前账户**
    driver.get("https://example.com/logout")  # 替换成登出的 URL
    time.sleep(3)

# **10. 关闭浏览器**
driver.quit()
print("🎉 所有账户处理完成！")
