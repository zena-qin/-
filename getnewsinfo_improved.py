import os
import time
import re
import json
import random
import pandas as pd
import requests
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class WeChatSuperCrawler:
    def __init__(self, start_date, end_date):
        self.base_url = "https://mp.weixin.qq.com"
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Connection": "keep-alive",
            "Referer": "https://mp.weixin.qq.com/",
            "Origin": "https://mp.weixin.qq.com",
            "X-Requested-With": "XMLHttpRequest"
        }
        self.session = requests.Session()
        self.token = ""
        self.cookies = {}

        self.driver_path = os.path.join(os.getcwd(), "chromedriver.exe")
        self.start_ts = int(time.mktime(time.strptime(start_date, "%Y%m%d")))
        self.end_ts = int(time.mktime(time.strptime(end_date, "%Y%m%d"))) + 86399
        self.output_excel = f"文章信息_{start_date}至{end_date}.xlsx"
        self.all_data = []
        self.debug = True   # 开启调试打印

    def login(self):
        print("🚀 正在启动登录浏览器，请扫码...")
        try:
            service = Service(executable_path=self.driver_path)
            driver = webdriver.Chrome(service=service)
        except Exception as e:
            print(f"❌ 启动失败: {e}")
            print(f"💡 请确认 {self.driver_path} 是否在当前目录下！")
            return

        driver.get(self.base_url)
        while "token=" not in driver.current_url:
            time.sleep(1)
        print("✅ 登录成功，正在同步 Cookies...")
        self.token = re.findall(r'token=(\d+)', driver.current_url)[0]
        selenium_cookies = driver.get_cookies()
        for cookie in selenium_cookies:
            self.cookies[cookie['name']] = cookie['value']
        self.session.headers.update(self.headers)
        self.session.cookies.update(self.cookies)

        # 额外访问一次首页，刷新会话状态
        driver.get(self.base_url)
        time.sleep(2)
        driver.quit()
        print("✅ 登录态已同步")

    def init_robust_driver(self):
        options = Options()
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1280,800')
        service = Service(executable_path=self.driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(30)
        return driver

    def search_account(self, query):
        url = f"{self.base_url}/cgi-bin/searchbiz"
        params = {
            "action": "search_biz",
            "begin": "0",
            "count": "5",
            "query": query,
            "token": self.token,
            "lang": "zh_CN",
            "f": "json",
            "ajax": "1"
        }
        try:
            res = self.session.get(url, params=params, timeout=10)
            if self.debug:
                print(f"   [DEBUG] search_account 返回: {res.text[:200]}...")
            data = res.json()
            if data.get('list'):
                return data['list'][0]['fakeid'], data['list'][0]['nickname']
        except Exception as e:
            print(f"   ⚠️ 搜索 '{query}' 出错: {e}")
        return None, None

    def get_article_list(self, fakeid):
        url = f"{self.base_url}/cgi-bin/appmsg"
        begin = 0
        count = 10  # 一次取10条
        while True:
            params = {
                "token": self.token,
                "lang": "zh_CN",
                "f": "json",
                "ajax": "1",
                "action": "list_ex",
                "begin": str(begin),
                "count": str(count),
                "query": "",
                "fakeid": fakeid,
                "type": "9"
            }
            try:
                res = self.session.get(url, params=params, timeout=10)
                if self.debug:
                    print(f"   [DEBUG] appmsg 返回: {res.text[:300]}...")
                data = res.json()
                # 检查是否有错误信息
                if data.get('ret') != 0:
                    print(f"   ⚠️ 接口返回错误: {data.get('errmsg', '未知错误')}")
                    break
                if not data.get('app_msg_list'):
                    break
                for item in data['app_msg_list']:
                    yield item
                # 如果返回的数量少于 count，说明已到末尾
                if len(data['app_msg_list']) < count:
                    break
                begin += count
                time.sleep(random.uniform(2, 4))
            except Exception as e:
                print(f"   ⚠️ 获取文章列表出错: {e}")
                break

    def get_real_author(self, driver, url):
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "js_name")))
            author_res = driver.execute_script("""
                let el = document.querySelector('#js_author_name');
                return el ? el.innerText.trim() : (window.author || '');
            """)
            return author_res
        except Exception as e:
            print(f"      [DEBUG] 获取作者失败: {e}")
            return ""

    def run(self, excel_file):
        if not os.path.exists(excel_file):
            print(f"❌ 找不到: {excel_file}")
            return

        df = pd.read_excel(excel_file)
        accounts = df.iloc[:, 0].dropna().tolist()

        self.login()

        print("🔧 初始化后台浏览器获取作者信息...")
        detail_driver = self.init_robust_driver()

        try:
            print(f"\n📋 任务开始：{len(accounts)} 个公众号")
            for idx, ac_name in enumerate(accounts, 1):
                print(f"\n📍 [{idx}/{len(accounts)}] 搜索: {ac_name}")
                fakeid, nickname = self.search_account(ac_name)
                if not fakeid:
                    print(f"   ⚠️ 未找到: {ac_name}")
                    continue

                print(f"   ✅ {nickname} | 获取列表中...")
                article_count = 0
                for item in self.get_article_list(fakeid):
                    update_time = item['update_time']
                    if update_time > self.end_ts:
                        continue
                    if update_time < self.start_ts:
                        print("      ⏰ 达到起始日期，跳过该号余下文章")
                        break

                    dt = datetime.fromtimestamp(update_time)
                    date_str = dt.strftime('%Y-%m-%d')
                    title = item['title']
                    link = item['link']

                    print(f"      📝 [{date_str}] {title[:20]}...")
                    real_author = self.get_real_author(detail_driver, link)

                    self.all_data.append({
                        "发表日期": date_str,
                        "作者": real_author if real_author else item.get('author', '未知'),
                        "来源公众号": nickname,
                        "文章标题": title,
                        "原文链接": link,
                        "发表年": dt.strftime('%Y'),
                        "发表月日": dt.strftime('%m月%d日')
                    })
                    article_count += 1
                    time.sleep(25)  # 避免风控

                print(f"   📊 共抓取 {article_count} 篇文章")
                time.sleep(40)  # 公众号间隔

        finally:
            detail_driver.quit()
            if self.all_data:
                df_out = pd.DataFrame(self.all_data)
                cols = ["发表日期", "作者", "来源公众号", "文章标题", "原文链接", "发表年", "发表月日"]
                df_out[cols].to_excel(self.output_excel, index=False)
                print(f"\n🎉 抓取完成！文件已保存: {self.output_excel}")
            else:
                print("\n⚠️ 未抓取到有效数据")

if __name__ == "__main__":
    try:
        s_date = input("请输入起始日期 (YYYYMMDD): ").strip()
        e_date = input("请输入结束日期 (YYYYMMDD): ").strip()
        EXCEL_List = "公众号列表.xlsx"
        crawler = WeChatSuperCrawler(s_date, e_date)
        crawler.run(EXCEL_List)
    except Exception as e:
        print(f"❌ 运行错误: {e}")
    finally:
        input("\n按回车退出...")
