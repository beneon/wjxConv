from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import yaml
import os

with open(os.path.join('..', 'env.yaml')) as _env_conf:
    _env_conf_yaml = yaml.safe_load(_env_conf)
    _chrome_path = _env_conf_yaml['chrome_driver_path']
driver = webdriver.Chrome(_chrome_path)
driver.get("https://www.wjx.cn/")
# click wechat login
elem_wechat = driver.find_element_by_xpath("//li[@id='ctl00_liWeiXin']/a")
elem_wechat.click()
# 点击了微信扫描登录, 接下来是需要互动扫描登录
# 当顺利登录以后，检查是否进入了登录后页面
try:
    elem_wait_for_qr = WebDriverWait(driver,30).until(
        EC.presence_of_element_located((By.ID,'qrcode'))
    )
    print('loaded')
finally:
    pass


# assert "Python" in driver.title
# elem = driver.find_element_by_name("q")
# elem.clear()
# elem.send_keys("pycon")
# elem.send_keys(Keys.RETURN)
# assert "No results found." not in driver.page_source
# driver.close()