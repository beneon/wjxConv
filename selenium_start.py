from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import yaml
import os

with open(os.path.join('..', 'env.yaml')) as _env_conf:
    _env_conf_yaml = yaml.safe_load(_env_conf)['chrome_driver']
    _chrome_path = _env_conf_yaml['chrome_driver_path']
    _chrome_app_path = _env_conf_yaml['binary']
options = ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('start-maximized')
options.binary_location = _chrome_app_path

driver = webdriver.Chrome(_chrome_path,options=options)
driver.get("https://www.wjx.cn/")
# click wechat login
elem_wechat = driver.find_element_by_xpath("//li[@id='ctl00_liWeiXin']/a")
elem_wechat.click()
# 点击了微信扫描登录, 接下来是需要互动扫描登录
# 当顺利登录以后，检查是否进入了登录后页面

elem_wait_for_qr = WebDriverWait(driver,30).until(
    EC.title_contains('我的问卷')
)

# 接下来由操作者选择对应的问卷页面，点击进入问卷数据下载页面
if elem_wait_for_qr:
    print('进入我的问卷界面')
    elem_wait_for_download = WebDriverWait(driver,120).until(
        EC.title_contains('查看下载答卷')
    )
# 进入下载页面以后，自动下载按序号编码的数据
if elem_wait_for_download:
    print('进入下载页面')
    elem_download_pop = driver.find_element_by_xpath('//a[@id="ctl02_ContentPlaceHolder1_ViewStatSummary1_ViewQuery1_hlAll"]')
    elem_download_btn = driver.find_element_by_xpath('//ul[@id="divDown"]/li[1]/a')
    ActionChains(driver).move_to_element(elem_download_pop).click(elem_download_btn).perform()
# 接下来开始下载附件
    elem_attachment = driver.find_element_by_xpath('//a[@id="ctl02_ContentPlaceHolder1_ViewStatSummary1_ViewQuery1_hrefDownAttach"]')
    elem_attachment.click()
    # 切换到新生成的tab
    driver.switch_to.window(driver.window_handles[1])
    elem_wait_for_attach_download = WebDriverWait(driver,30).until(
        EC.text_to_be_present_in_element((By.XPATH,'//div[@id="layui-layer1"]/div[1]'),'提示')
    )

# 提交下载任务的说明已经生成
if elem_wait_for_attach_download:
    print('attachment download submitted')
    elem_attach_submit_ok = driver.find_element_by_xpath('//div[@id="layui-layer1"]/div[3]/a')
    elem_attach_submit_ok.click()
    # TODO：正常来说会先来到一个“下载任务列表”这样的页面，这里应该隔10秒点击一次刷新页面比较好
    # elem_refresh_page = driver.find_element_by_xpath('//div[@id="ctl02_ContentPlaceHolder1_divDownTime"]/abc')
    # elem_refresh_page.click()
    # time.sleep(10)
    # 等待附件打包完成
    elem_wait_for_attach_pack = WebDriverWait(driver,120).until(
        EC.text_to_be_present_in_element((By.XPATH,'//table[@id="ctl02_ContentPlaceHolder1_tbSummary"]/tbody/tr[1]/td[2]/a'),'下载')
    )

# 此时打包也已经完成
if elem_wait_for_attach_pack:
    driver.find_element_by_xpath('//table[@id="ctl02_ContentPlaceHolder1_tbSummary"]/tbody/tr[1]/td[2]/a').click()

# future dev：1. 能否知道文件是否已经下载完成？
# 对于下载的文件，找到最新的几个文件的方式是：
# list_of_files=glob.glob(os.path.expanduser(r'~\Downloads\*'))
# sorted(list_of_files,key=os.path.getctime)


# assert "Python" in driver.title
# elem = driver.find_element_by_name("q")
# elem.clear()
# elem.send_keys("pycon")
# elem.send_keys(Keys.RETURN)
# assert "No results found." not in driver.page_source
# driver.close()