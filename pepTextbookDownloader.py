import requests
import time
import re
from PIL import Image
import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

print("正在初始化...")
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # 禁用自动化标签

print("初始化已完成.")
input("请在即将打开的浏览器窗口中直接使用[在线阅读]按钮选择要下载的教材.按Enter键继续.")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    driver.get("https://jc.pep.com.cn/") 
    original_window = driver.current_window_handle
    WebDriverWait(driver, 60).until(EC.new_window_is_opened([original_window]))
    new_window = [window for window in driver.window_handles if window != original_window][0]
    driver.switch_to.window(new_window)
    new_url = driver.current_url
    
finally:
    driver.quit()

if ("https://book.pep.com.cn/" in new_url) == False:
    print("打开了错误的教材, 程序退出.")
    exit()
else:
    BookId = ''
    for i in range(0,len(new_url)):
        if new_url[i-11:i] == "pep.com.cn/":
            j = i
            while new_url[j] != "/":
                BookId = BookId + str(new_url[j])
                j += 1

BookId = int(BookId)

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Referer": f"https://book.pep.com.cn/{BookId}/mobile/index.html",
}
cookies = {}

def getCookie():
    global cookies
    driver.get(f"https://book.pep.com.cn/{BookId}/mobile/index.html")
    initial_title = driver.title
    for _ in range(90):
        current_title = driver.title
        if current_title != initial_title:
            cookie = driver.get_cookies()
            cookies = {cookie['name']: cookie['value'] for cookie in cookie} #直接以captcha页的cookie覆盖现有cookie
            return True
        time.sleep(1)
        
    print("超时")
    return False

def wget(url,filePath = None):
    response = requests.get(url, headers=headers,cookies=cookies)
    if ("text/html" in response.headers.get('Content-Type', '').lower()): # 检查内容类型是否包含 "text/html" 字符串
        print("请通过人机验证")
        if(getCookie()):
            print("人机验证成功。")
            response = requests.get(url, headers=headers,cookies=cookies) #再次尝试
        else:
            print("人机验证失败，程序自动结束。")
            quit()
        
    if response.status_code == 200:
        print(f"抓取{url}成功")
        if(filePath != None):
            with open(filePath, "wb") as file:
                file.write(response.content)
        else:
            return response
    else:
        print("下载失败，程序自动结束。")
        quit()

config = wget(f"https://book.pep.com.cn/{BookId}/mobile/javascript/config.js").text

createTimeMatch = re.search(r'bookConfig\.CreatedTime\s*=\s*"(\d+)"', config)
BookTitleMatch = re.search(r'bookConfig\.bookTitle\s*=\s*"(.*?)"', config)
totalPageCntMatch = re.search(r'bookConfig\.totalPageCount\s*=\s*(\d+)', config)

tPage = totalPageCntMatch.group(1) if totalPageCntMatch else None
cTime = createTimeMatch.group(1) if createTimeMatch else None
bookTitle = BookTitleMatch.group(1) if BookTitleMatch else None

if input("是否确认下载(Y/N) : 当前下载书本为[ " + bookTitle + " ], 总页数为[ " + str(tPage) + " ], 书籍编号为[ " + str(BookId) + " ]") != "Y":
    exit()

curDir = os.path.dirname(os.path.abspath(__file__))
saveDir = os.path.join(curDir, bookTitle)

os.makedirs(saveDir, exist_ok=True)
for i in range (1, int(tPage)+1):
    SingleImgPath = os.path.join(saveDir, f"{i}.jpg")
    wget(f"https://book.pep.com.cn/{BookId}/files/mobile/{i}.jpg?{cTime}",SingleImgPath)
    print("进度:{}%. ".format(str(round(i/int(tPage)*1000)/10)))

imgFolder = "./{}".format(bookTitle)
pdfOutput = "{}.pdf".format(bookTitle)
image_files = sorted([f for f in os.listdir(imgFolder) if f.endswith(".jpg")], key=lambda x: int(os.path.splitext(x)[0]))
images = [Image.open(os.path.join(imgFolder, img)).convert("RGB") for img in image_files]
if images:
    images[0].save(pdfOutput, save_all=True, append_images=images[1:])
    print(f"已保存为 {pdfOutput}")
else:
    print("未找到 JPG 图片文件。")
