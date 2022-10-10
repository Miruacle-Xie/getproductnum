# -*- coding: utf-8 -*-
"""
Created on Tue Sep 20 14:23:56 2022

@author: Administrator
"""
import sys
import html
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import re
import os
import time
import random
DEBUG = False


def amazonDeliverInit(driver):
    url = "https://www.amazon.com"
    try:
        logtext = []
        driver.get(url)
        logtext.append("driver.get(https://www.amazon.com)")
        content = driver.page_source
        logtext.append("driver.page_source")
        if "New York 10041" not in content:
            logtext.append("开始设置配送地址...")
            tmplog = driver.find_element_by_xpath('//*[@id="nav-packard-glow-loc-icon"]').click()
            logtext.append(tmplog)
            tmplog = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="GLUXZipUpdateInput"]')))  # 元素是否可见
            logtext.append(tmplog)
            tmplog = driver.find_element_by_xpath('//*[@id="GLUXZipUpdateInput"]').send_keys("10041")
            logtext.append(tmplog)
            time.sleep(1)
            tmplog = driver.find_element_by_xpath('//*[@id="GLUXZipUpdate"]/span/input').click()
            logtext.append(tmplog)
            time.sleep(3)
            driver.refresh()
            content1 = driver.page_source
            logtext.append(content1)
            # fd = open("1.txt", mode='w', encoding="utf8")
            # fd.write(content1)
            # fd.close()
            if "New York 10041" not in content1:
                print("设置地址失败")
                return False
            else:
                print("地址设为10041成功")
                return True
        else:
            print("初始化成功")
            return True
    except Exception as e:
        print("设置地址异常")
        repr(e)
        logtext.append(repr(e))
        exepath = sys.executable
        print(exepath)
        exepath = os.path.dirname(exepath)
        print(exepath)
        print(logtext)

        logContent = exepath + '\\' + "logContent.txt"
        fd = open(logContent, mode='w', encoding="utf8")
        fd.write(content)
        fd.close()
        logFile = exepath + '\\' + "log.txt"
        fd = open(logFile, mode='w', encoding="utf8")
        fd.write(str(logtext))
        fd.close()
        return False


def getSubjectName(fileName):
    df = pd.read_excel(fileName, sheet_name=0)
    # print(df)
    return df


def getAmazonResult(driver, subjectName):
    try:
        # url = "https://www.amazon.com/s?k=" + subjectName.replace(" ", "+") + "+shirt"
        # print(url)

        driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').click()

        eleValue = driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]')
        if eleValue.get_attribute('value') is not None:
            driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').clear()
        driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').send_keys(subjectName)
        print('{:30s}{}'.format('getAmazonResult-<ENTER>-start', time.strftime('%Y-%m-%d %H:%M:%S')))
        # driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]').send_keys(Keys.ENTER)
        driver.find_element_by_xpath('//*[@id="nav-search-submit-button"]').click()
        print('{:30s}{}'.format('getAmazonResult-<ENTER>-end', time.strftime('%Y-%m-%d %H:%M:%S')))
        pageSource = driver.page_source

# =============================================================================
#         fd = open("2.txt", mode='w', encoding="utf8")
#         fd.write(pageSource)
#         fd.close()
# =============================================================================
        return True, pageSource
    except Exception as e:
        print("getAmazonResult异常")
        repr(e)
        return False, None


def getresult(driver, df, filepath):
    time_start = time.time()
    colName = df.columns.values
    try:
        for cnt in range(1, len(colName)):
            productNumReslut = []
            if cnt == 1:
                productName = ""
            else:
                productName = " " + colName[cnt]

            for i in range(len(df)):
            # for i in range(2):
                print("第{}次验证".format((i + 1)))
                print('{:30s}{}       {}'.format('getAmazonResult-start', time.strftime('%Y-%m-%d %H:%M:%S'), (df[colName[0]][i] + productName).capitalize()))
                accessFlag, pageSource = getAmazonResult(driver, df[colName[0]][i] + productName)
                print('{:30s}{}'.format('getAmazonResult-end', time.strftime('%Y-%m-%d %H:%M:%S')))

                try:
                    if accessFlag:
                        searchResult = re.search(r"<span>.* of over (.*) results for</span>|<span>.* of (.*) results for</span>|<span>(.*) results for</span>", pageSource)

                        if searchResult is not None:
                            # for j in range(1, len(searchResult)):
                            for tmpsearchResult in searchResult.groups():
                                if tmpsearchResult is not None:
                                    productNumReslut.append(tmpsearchResult)
                                    print(tmpsearchResult)
                        else:
                            productNumReslut.append("亚马逊产品数量为0")
                except Exception as e:
                    print(repr(e))
                    print("第{}次验证异常-{}".format((i + 1), (df[colName[0]][i] + productName).capitalize()))
                    exepath = sys.executable
                    exepath = os.path.dirname(exepath)
                    pageSourceContent = exepath + '\\' + "pageSourceContent.txt"
                    fd = open(pageSourceContent, mode='w', encoding="utf8")
                    fd.write(pageSource)
                    fd.close()
            df[colName[cnt]] = pd.Series(productNumReslut)
    finally:
        reportpath = os.path.splitext(filepath)[0] + "-亚马逊检测报告" + ".xlsx"
        writer = pd.ExcelWriter(reportpath)
        # header = None：数据不含列名，index=False：不显示行索引（名字）
        df.to_excel(writer, index=False)
        writer.save()
        time_end = time.time()
        driver.quit()
        input("已生成报告, 耗时时间:{}, 平均耗时:{}, 按回车键结束".format(time_end - time_start, (time_end - time_start) / len(df)))

def main():
    global DEBUG
    filepath = input("\n文件路径：\n")
    # filepath = "F:\\JetBrains\\affirm_amazon\\selenium\\0905-快主题-拆分结果.xlsx"
    if " --DEBUG" in filepath:
        DEBUG = True
        filepath = filepath.replace(" --DEBUG", "")
        print("filepath:{}".format(filepath))
    filepath = filepath.replace("\"", "").replace("\'", "")
    df = getSubjectName(filepath)
    print(df)
    chrome_options = Options()
    if not DEBUG:
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"')
    chrome_options.add_argument('--incognito')
    chrome_options.page_load_strategy = 'eager'
    driver = webdriver.Chrome(chrome_options=chrome_options)
    # driver = webdriver.Chrome(options=chrome_options)

    # driver = webdriver.Chrome()
    print("初始化地址中...")
    if amazonDeliverInit(driver):
        print("即将开始...")
        getresult(driver, df, filepath)
    else:
        input("访问亚马逊异常, 回车结束")


if __name__ == "__main__":
    # df = getSubjectName("0905-快主题-拆分结果.xlsx")
    main()
