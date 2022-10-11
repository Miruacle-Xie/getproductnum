# -*- coding: utf-8 -*-
"""
Created on Tue Sep 20 14:23:56 2022

@author: Administrator
"""
import sys
import html

import numpy as np
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

DEBUG = True


def amazonDeliverInit(driver):
    url = "https://www.amazon.com"
    try:
        logtext = []
        driver.get(url)
        logtext.append("driver.get(https://www.amazon.com)")
        content = driver.page_source
        logtext.append("driver.page_source")
        # input("111")
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


def productarrange(dftmp):
    # print("productarrange-df")
    # print(dftmp)
    arrange = []
    arrangeresult = []
    for i in range(len(dftmp)):
        # print(dftmp.iloc[i, :])
        num = np.array(dftmp.iloc[i, :].tolist())
        sorted_index = sorted(range(len(num)), key=lambda k: num[k])
        # print(sorted_index)
        arrange.append(sorted_index)

    # tmplist = np.argsort(df, axis=1)
    # print("tmplist:")
    # print(tmplist)
    # for i in range(len(df)):
    #     print(i)
    #     print(tmplist.iloc[i, :].tolist())
    #     arrange.append(tmplist.iloc[i, :].tolist())
    # arrange = [np.argsort(row) for row in tmplist]
    # print("arrange:")
    # print(arrange)
    for tmp in arrange:
        tmpvalue = ""
        for m in tmp:
            # print(m)
            tmpvalue = tmpvalue + dftmp.columns.values[m] + " "
        # print(tmpvalue.strip())
        arrangeresult.append(tmpvalue.strip())
    # print(arrangeresult)

    return arrangeresult


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
                print('{:30s}{}       {}'.format('getAmazonResult-start', time.strftime('%Y-%m-%d %H:%M:%S'),
                                                 (df[colName[0]][i] + productName).capitalize()))
                accessFlag, pageSource = getAmazonResult(driver, df[colName[0]][i] + productName)
                print('{:30s}{}'.format('getAmazonResult-end', time.strftime('%Y-%m-%d %H:%M:%S')))

                try:
                    if accessFlag:
                        searchResult = re.search(
                            r"<span>.* of over (.*) results for</span>|<span>.* of (.*) results for</span>|<span>(.*) results for</span>",
                            pageSource)

                        if searchResult is not None:
                            # for j in range(1, len(searchResult)):
                            for tmpsearchResult in searchResult.groups():
                                if tmpsearchResult is not None:
                                    productNumReslut.append(tmpsearchResult)
                                    # print(tmpsearchResult)
                        else:
                            productNumReslut.append("0")
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
        writer.close()
        time_end = time.time()
        driver.quit()
        # writer = pd.ExcelWriter(reportpath)
        # dfnew = getSubjectName(reportpath)
        # # print(dfnew)
        # dftmp = dfnew.iloc[:, 2:]
        # # print(dftmp)
        # dfnew["产品数量排序"] = pd.Series(productarrange(dftmp)).str.replace(" ", "<")
        # dfnew.to_excel(writer, index=False)
        # writer.save()
        input("已生成报告, 耗时时间:{}, 平均耗时:{}, 按回车键结束".format(time_end - time_start, (time_end - time_start) / len(df)))


def main(filepath):
    df = getSubjectName(filepath)
    print(df)
    chrome_options = Options()
    if not DEBUG:
        chrome_options.add_argument('--headless')
        chrome_options.add_argument(
            'user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"')
    chrome_options.add_argument('--incognito')
    chrome_options.page_load_strategy = 'eager'
    driver = webdriver.Chrome(chrome_options=chrome_options)
    # driver = webdriver.Chrome(options=chrome_options)

    # driver = webdriver.Chrome()
    print("初始化地址中...")
    if amazonDeliverInit(driver):
        print("即将开始...")
        getresult(driver, df, filepath)
        # df["产品数量排序"] = pd.Series(productarrange(dftmp)).str.replace(" ", "<")
    else:
        input("访问亚马逊异常, 回车结束")
        driver.quit()


def main1(filepath):
    try:
        df = getSubjectName(filepath)
        print(df)
        # reportpath = os.path.splitext(filepath)[0] + "-亚马逊检测报告" + ".xlsx"
        reportpath = os.path.splitext(filepath)[0] + "-tmp" + ".xlsx"
        writer = pd.ExcelWriter(reportpath)
        dftmp = df.iloc[:, 2:]
        # print(dftmp)
        df["产品数量排序"] = pd.Series(productarrange(dftmp)).str.replace(" ", "<")
        df.to_excel(writer, index=False)
        writer.save()
        writer.close()
        os.rename(filepath, filepath + "-backup")
        os.rename(reportpath, filepath)
        os.remove(filepath + "-backup")
        input("报告处理完成, 回车结束")
    except Exception as e:
        print(repr(e))
        input("出现异常, 请确保表格格式正确, 按回车结束")


if __name__ == "__main__":
    filePath = input("\n请输入文件路径：\n（空格+y为统计排序模式）\n")
    # filepath = "F:\\JetBrains\\affirm_amazon\\selenium\\0905-快主题-拆分结果.xlsx"
    if " --DEBUG" in filePath:
        DEBUG = True
        filePath = filePath.replace(" --DEBUG", "")
        print("filepath:{}".format(filePath))
    filePath = filePath.replace("\"", "").replace("\'", "")
    if filePath[-1].lower() == "y":
        try:
            filePath = filePath.replace(" y", "")
            main1(filePath)
        except:
            input("请输入小写y,再重新使用")
    else:
        main(filePath)
    # lis = np.array([413,172,20])
    # sorted_index = sorted(range(len(lis)), key=lambda k: lis[k])
    # print(sorted_index)
    # lis = lis[sorted_index]
    # print(lis)
    # df = getSubjectName("916-快主题第-修改结果3-亚马逊检测报告.xlsx")
    # print(df)
    # dftmp = df.iloc[:, 2:]
    # print(dftmp)
    # for i in range(len(dftmp)):
    #     print(dftmp.iloc[i, :])
    #     num = np.array(dftmp.iloc[i, :])
    #     a = num.argsort()
    #     print(num)
    #     print(a)
    # tmplist = np.argsort(dftmp, axis=1)
    # print(tmplist)
    # arrange = []
    # for i in range(len(dftmp)):
    #     print(i)
    #     print(tmplist.iloc[i, :].tolist())
    #     arrange.append(tmplist.iloc[i, :].tolist())
    # print(arrange)
    # arrangeresult = []
    # for tmp in arrange:
    #     tmpvalue = ""
    #     for m in tmp:
    #         print(m)
    #         tmpvalue = tmpvalue + df.columns.values[m + 2] + " "
    #     print(tmpvalue.strip())
    #     arrangeresult.append(tmpvalue.strip())
    # print(arrangeresult)
    #
    # df["aa"] = pd.Series(arrangeresult).str.replace(" ", "<")
    # df.to_excel("1.xlsx", index=False)
