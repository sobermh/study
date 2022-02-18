"""
@author:maohui
@time:2022/2/18 11:27
"""
import time

from selenium import webdriver
from selenium.webdriver.common.by import By  # 定位符
from selenium.webdriver.support.wait import WebDriverWait  # 等待类
from selenium.webdriver.support import expected_conditions  # 等待条件


class AboutPage():
    # 定义类的熟悉并完成初始化
    def __init__(self, driver):
        self.menu_locator = (By.CSS_SELECTOR, "/html/body/div[1]/div[1]/div/div/div[2]")
        self.driver = driver
        self.driver=webdriver.Edge()#调试使用
        self.wait = WebDriverWait(driver, 30)

    # 定义页面操作方法
    # 点击关于汇健
    def click_menu(self):
        try:
            # self.wait.until(expected_conditions.presence_of_element_located((self.menu_locator)))
            time.sleep(5)
            self.driver.find_element(By.XPATH,"/html/body/div[6]/div/ul/li[1]/a/div[1]/img").click()
        except Exception as e:
            raise e


# 调试
if __name__ == "__main__":
    driver = webdriver.Edge()
    driver.get("http://hj.demoweb.68hanchen.com/")
    AboutPage = AboutPage(driver)
    AboutPage.click_menu()
    time.sleep(1)
    driver.quit()
