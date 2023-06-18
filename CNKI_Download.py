import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import PySimpleGUI as sg

def open_page(driver, theme):
    driver.get("https://www.cnki.net")
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txt_SearchText"]'))).send_keys(theme)
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/input[2]"))).click()
    time.sleep(3)

    # 获取搜索结果的文献数
    res_unm_element = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[2]/div[2]/div[2]/form/div/div[1]/div[1]/span[1]/em")))
    res_unm = int(res_unm_element.text.replace(",", ''))
    page_unm = int(res_unm / 20) + 1
    print(f"共找到 {res_unm} 条结果, {page_unm} 页。")
    print(f"\n")
    return res_unm

def crawl(driver, papers_need, theme):
    # 创建DataFrame保存文献信息
    df = pd.DataFrame(columns=['序号', '标题', '作者', '机构', '日期', '来源', '数据库', '关键词', '摘要', '链接'])
    # 赋值序号, 控制爬取的文章数量
    count = 1
    # 当爬取数量小于需求时，循环网页页码
    while count <= papers_need:
        # 等待加载完全，休眠3S
        time.sleep(3)
        title_list = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "fz14")))
        # 循环网页一页中的条目
        for i in range(len(title_list)):
            try:
                if count % 20 != 0:
                    term = count % 20  # 本页的第几个条目
                else:
                    term = 20 # 本页的第20个条目
                title_xpath = f"/html/body/div[3]/div[2]/div[2]/div[2]/form/div/table/tbody/tr[{term}]/td[2]"
                author_xpath = f"/html/body/div[3]/div[2]/div[2]/div[2]/form/div/table/tbody/tr[{term}]/td[3]"
                source_xpath = f"/html/body/div[3]/div[2]/div[2]/div[2]/form/div/table/tbody/tr[{term}]/td[4]"
                date_xpath = f"/html/body/div[3]/div[2]/div[2]/div[2]/form/div/table/tbody/tr[{term}]/td[5]"
                database_xpath = f"/html/body/div[3]/div[2]/div[2]/div[2]/form/div/table/tbody/tr[{term}]/td[6]"
                title = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, title_xpath))).text
                authors = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, author_xpath))).text
                source = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, source_xpath))).text
                date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, date_xpath))).text
                database = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, database_xpath))).text
                # 点击条目
                title_list[i].click()
                # 获取driver的句柄
                n = driver.window_handles
                # driver切换至最新生产的页面
                driver.switch_to.window(n[-1])
                time.sleep(3)
                # 开始获取页面信息
                title = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h1"))).text
                authors = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h3[1]"))).text
                institute = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h3[2]"))).text
                abstract = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "abstract-text"))).text
                try:
                    keywords = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "keywords"))).text[:-1]
                except:
                    keywords = '无'
                url = driver.current_url
                # 获取下载链接
                # link = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "btn-dlcaj")))[0].get_attribute('href')
                # link = urljoin(driver.current_url, link)
                # 写入DataFrame
                df.loc[count] = [count, title, authors, institute, date, source, database, keywords, abstract, url]
                # 结果显示在用户界面上
                print(f"【检索编号】{count}\n【文献标题】{title}\n【文献作者】{authors}\n【作者机构】{institute}\n【发表日期】{date}\n【文献来源】{source}\n【数据库】{database}\n【关键词】{keywords}\n\n【文献摘要】{abstract}\n\n【来源地址】{url}\n\n")

            except Exception as e:
                print(f"第{count-1}条爬取失败: {str(e)}\n")
                # 跳过本条，继续下一个
                continue
            finally:
                # 如果有多个窗口，关闭第二个窗口，切换回主页
                n2 = driver.window_handles
                if len(n2) > 1:
                    driver.close()
                    driver.switch_to.window(n2[0])
                # 计数,判断需求是否足够
                count += 1
                if count > papers_need:
                    break
        # 切换到下一页
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@id='PageNext']"))).click()
    return df

# 图形化界面
def gui_layout():
    sg.theme('LightGrey')
    layout = [
        [sg.Text('中国知网(CNKI)文献检索与信息下载程序', size=(33, 1), justification='center', font=('微软雅黑', 20), relief=sg.RELIEF_RIDGE)],
        [sg.Text('搜索关键词:', size=(15, 1), font=('微软雅黑', 12)), sg.Input(key='-KEYWORD-', size=(40, 1), default_text=('例如：在线学习'), font=('微软雅黑', 12))],
        [sg.Text('搜索文献数:', size=(15, 1), font=('微软雅黑', 12)), sg.Input(key='-PAPERS_NEED-', size=(40, 1), default_text=('例如：10'), font=('微软雅黑', 12))],
        [sg.Button('搜索', key='-SEARCH-', size=(10, 1), font=('微软雅黑', 12))],
        [sg.Output(size=(57, 20), font=('微软雅黑', 12), key='-OUTPUT-')]
    ]
    window = sg.Window('CNKI文献搜索', layout, icon='icon.ico')
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == '-SEARCH-':
            keyword = values['-KEYWORD-']
            papers_need = int(values['-PAPERS_NEED-'])
            # get直接返回，不再等待界面加载完成
            desired_capabilities = DesiredCapabilities.CHROME
            desired_capabilities["pageLoadStrategy"] = "none"
            # 设置Edge驱动器的环境
            options = webdriver.EdgeOptions()
            # 设置Edge不加载图片，提高速度
            options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
            # 创建一个Edge驱动器
            driver = webdriver.Edge(options=options)
            driver.maximize_window()
            res_unm = open_page(driver, keyword)
            # 判断所需篇数是否大于总篇数
            papers_need = papers_need if papers_need <= res_unm else res_unm
            df = crawl(driver, papers_need, keyword)
            # 保存到Excel文件
            df.to_excel(f'CNKI_{keyword}.xlsx', index=False)
            # 关闭浏览器
            driver.close()
            print(f"\n爬取完成！结果已保存到文件 <CNKI_{keyword}.xlsx> 中。\n")
    window.close()

if __name__ == "__main__":
    gui_layout()
