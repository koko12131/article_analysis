#read me before using
#该程式基于pycharm开发环境使用python3.6.1，抓取对象为DII数据库
#需要安装的库包括：selenium、pyautogui、chromedriver
#本程式对于DII数据库指定年份时间间隔进行逐年专利数据进行提取，并保存在指定的路径中
#记录下载顺序为按照被引频次 如果想要按照更新时间请将第84行修改为#sort_type.click()【注意保留原代码前的缩进】
#单次记录下载不能超过5万条，同时德文特不支持输出检索记录10万条以后的记录




from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
import time,random
from pathlib import Path
import os
import getpass
import pyautogui
import shutil
import glob

def search(begin_year,end_year,browser,wait,search_keywords,output_path):
    browser.get('http://apps.webofknowledge.com/')
    browser.maximize_window()
    databases_select_button = wait.until(EC.element_to_be_clickable(
        (By.CLASS_NAME, 'select2-selection__arrow')))
    actions = ActionChains(browser)
    actions.move_to_element(databases_select_button)
    actions.click()
    actions.perform()
    time.sleep(1)
    pyautogui.press('down')
    pyautogui.press('enter')  # 选择DII数据库成功
    time.sleep(1)
    search_type = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'body > div.EPAMdiv.main-container > div.block-search.block-search-header > div > ul > li:nth-child(4) > a')
    ))
    search_type.click()  # 选择高级检索成功
    time.sleep(1)
    TS_input = wait.until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="value(input1)"]')))
    TS_input.clear()
    TS_input.send_keys(search_keywords)
    language_type = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '#value\(input2\) > option:nth-child(2)')
    ))
    language_type.click()
    doucumet_type = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '#value\(input3\) > option:nth-child(2)')
    ))
    doucumet_type.click()
    more_setting = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR,
         "#timespan > div:nth-child(2) > div > span > span.selection > span > span.select2-selection__arrow")))
    # more_setting.click()
    # pyautogui.press('down')
    # pyautogui.press('down')
    # pyautogui.press('down')
    # pyautogui.press('down')
    # pyautogui.press('down')
    # pyautogui.press('enter')
    # begin_year_button = wait.until(EC.presence_of_element_located(
    #     (By.XPATH, '//*[@id="timespan"]/div[3]/div/span[2]/span[1]/span/span[2]')))
    # begin_year_button.click()
    # pyautogui.typewrite(begin_year)
    # pyautogui.press('enter')
    # end_year_button = wait.until(EC.presence_of_element_located(
    #     (By.XPATH, '//*[@id="timespan"]/div[3]/div/span[4]/span[1]/span')))
    # end_year_button.click()
    # pyautogui.typewrite(end_year)
    # pyautogui.press('enter')
    time.sleep(1)
    search_button = wait.until(EC.presence_of_element_located(
        (By.CLASS_NAME, "searchButtons")))
    search_button.click()
    result = wait.until(EC.presence_of_element_located(
        (By.CLASS_NAME, "historyResults")))
    result.click()
    time.sleep(7)
    refresh_i = 0
    while refresh_i == 0:
        try:
            sort_type = wait.until(EC.presence_of_element_located(
                (By.ID, "LC.D;PY.D;AU.A.en;SO.A.en;VL.D;PG.A")))
            refresh_i = 1
            break
        except Exception:
            browser.refresh()
            time.sleep(10)
            refresh_i = 0
            continue
    hit_count = browser.find_element_by_id(r'hitCount.top').text.replace(',', '')
    print('Start', hit_count + '条记录')
    if int(hit_count) > 500:
        recoding_range_list = int(int(hit_count) - int(hit_count) % 500)
    else:
        recoding_range_list = None
    # sort_type.click()
    '''s1 = Select(browser.find_element_by_id('saveToMenu'))
    s1.select_by_visible_text('其他文件格式')'''
    save_type = wait.until(EC.presence_of_element_located(
        (By.ID, "exportTypeName")))
    save_type.click()
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    file_save_f(wait, recoding_range_list, hit_count,output_path,browser)


def file_save_f(wait,recoding_range_list,hit_count,output_path,browser):
    mark = False
    for i in range(0, recoding_range_list, 500):
        while i == 0:
            i = i + 1
            result_save1(wait, str(i), str(i + 499))
        else:
            type_click = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR,
                 "#page > div.EPAMdiv.main-container > div.NEWsummaryPage > div.NEWsummaryDataContainer > div > div > div > div.l-column-content > div.l-content > div:nth-child(6) > div.export_options > div.selectedExportOption > ul > li > button")))
            i = i + 1
            if i == 2:
                continue
            else:
                type_click.click()
                time.sleep(2)
                result_save2(wait,str(i), str(i + 499),browser)
            if i+499 == 50500:
                mark=True
                time.sleep(60)
                file_move(output_path,mark = mark)
    type_click = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR,
         "#page > div.EPAMdiv.main-container > div.NEWsummaryPage > div.NEWsummaryDataContainer > div > div > div > div.l-column-content > div.l-content > div:nth-child(6) > div.export_options > div.selectedExportOption > ul > li > button")))
    type_click.click()
    result_save2(wait, str(recoding_range_list + 1), str(int(hit_count)),browser)
    #最后不能被500除尽的余数 进行下载
    time.sleep(60)
    file_move(output_path,mark= mark)
    print('finish', hit_count + '条记录')
    browser.close()

def file_move(output_path,mark=False):
    path_temp = r'C:\Users' + os.sep + getpass.getuser() + os.sep + 'Dii_file_download_temp'
    txt_filenames = glob.glob(path_temp + os.sep +'*.txt')
    if mark and len(txt_filenames) == 101:
        for filename,i in zip(txt_filenames,list(range(1,len(txt_filenames)+1))):
            shutil.move(filename,os.path.join(output_path,'download_'+str(i)+'.txt'))
    elif mark and len(txt_filenames) < 101:
        for filename,i in zip(txt_filenames,list(range(102,102+len(txt_filenames)))):
            shutil.move(filename,os.path.join(output_path,'download_'+str(i)+'.txt'))
    elif len(txt_filenames)<100 and mark==False:
        for filename, i in zip(txt_filenames, list(range(1,len(txt_filenames)+1))):
            shutil.move(filename, os.path.join(output_path,'download_'+str(i)+'.txt'))

def result_save1(wait,begin_numbers,end_numbers):
    save_window = wait.until(EC.presence_of_element_located(
            (By.ID,"numberOfRecordsRange")))
    save_window.click()
    begin_number =  wait.until(EC.presence_of_element_located(
            (By.XPATH,'//*[@id="markFrom"]')))
    begin_number.clear()
    begin_number.send_keys(begin_numbers)
    end_number = wait.until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="markTo"]')))
    end_number.clear()
    end_number.send_keys(end_numbers)
    recording_text = wait.until(EC.presence_of_element_located(
        (By.ID, 'select2-bib_fields-container')))
    recording_text.click()
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    file_type = wait.until(EC.presence_of_element_located(
        (By.ID, 'select2-saveOptions-container')))
    file_type.click()
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(1)
    save_to_PC = wait.until(EC.presence_of_element_located(
            (By.CLASS_NAME,"quickoutput-action" )))
    save_to_PC.click()
    time.sleep(7)

def refresh_browser(wait,begin_numbers,end_numbers,url):
    options = webdriver.ChromeOptions()
    out_path_temp = r'C:\Users' + os.sep + getpass.getuser() + os.sep + 'Dii_file_download_temp'
    if not Path(out_path_temp).exists():
        os.mkdir(out_path_temp)
    # 指定文件保存路径
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': out_path_temp}
    # 设置偏好
    options.add_experimental_option('prefs', prefs)
    # 添加偏好
    browser = webdriver.Chrome(options=options)
    browser.get(url)
    save_type = wait.until(EC.presence_of_element_located(
        (By.ID, "exportTypeName")))
    save_type.click()
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    result_save1(wait, begin_numbers, end_numbers)
    return result_save2(wait, str(int(end_numbers) + 1), str(int(end_numbers) + 500), browser)

def result_save2(wait,begin_numbers,end_numbers,browser):
    try:
        save_window = wait.until(EC.presence_of_element_located(
                (By.ID,"numberOfRecordsRange")))
        save_window.click()
        begin_number =  wait.until(EC.presence_of_element_located(
                (By.XPATH,'//*[@id="markFrom"]')))
        begin_number.clear()
        begin_number.send_keys(begin_numbers)
        end_number = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="markTo"]')))
        end_number.clear()
        end_number.send_keys(end_numbers)
        recording_text = wait.until(EC.presence_of_element_located(
            (By.ID, 'select2-bib_fields-container')))
        recording_text.click()
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(1)
        save_to_PC = wait.until(EC.presence_of_element_located(
                (By.CLASS_NAME,"quickoutput-action" )))
        save_to_PC.click()
        time.sleep(random.randint(8,15))
    except Exception:
        url = browser.current_url
        browser.close()
        refresh_browser(wait,begin_numbers,end_numbers,url)

    '''close_alter = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR,"#page > div.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-front.ui-dialog-quickoutput.qoOther > div.ui-dialog-titlebar.ui-widget-header.ui-corner-all.ui-helper-clearfix > button" )))
    close_alter.click()
    time.sleep(5)'''

def Dii_infor_spider_f(output_path,year_begin,year_end,search_keywords):
    options = webdriver.ChromeOptions()
    # 实例化浏览器option
    out_path_temp = r'C:\Users' + os.sep +getpass.getuser()+os.sep+'Dii_file_download_temp'
    if not Path(out_path_temp).exists():
        os.mkdir(out_path_temp)
    # 指定文件保存路径
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': out_path_temp}
    # 设置偏好
    options.add_experimental_option('prefs', prefs)
    # 添加偏好
    browser = webdriver.Chrome(options=options)
    # 实例化浏览器对象
    wait = WebDriverWait(browser, 7)
    # 设置响应时间
    if year_begin > year_end:
        print('开始年份:',year_begin,'大于','结束年份：',year_end)
    else:
        search(begin_year=year_begin,end_year=year_end, browser=browser, wait=wait, search_keywords=search_keywords,output_path = output_path)
    time.sleep(60)
    shutil.rmtree(out_path_temp)

def main():
    print('在德温特预检确认后再进行抓取')
    output_path = input('请输入文件保存路径：')
    year_begin = input('请输入检索开始年份：')
    year_end = input('请输入检索结束年份：')
    search_keywords = input('请输入检索式：')
    search_func_type = input('请输入需要抓取的方式（1：一次性抓取，2：逐年抓取）：')
    if search_func_type == '1':
        Dii_infor_spider_f(output_path=output_path,year_begin=year_begin,year_end=year_end,search_keywords=search_keywords)

    # else:
    #     D2.Dii_infor_spider_f(output_path=output_path,year_begin=year_begin,year_end=year_end,search_keywords=search_keywords)

if __name__ =='__main__':
    main()


