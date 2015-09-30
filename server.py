from flask import Flask, request, send_from_directory

from selenium import webdriver
import urllib
import csv
import xlsxwriter


from urlparse import urljoin
import selenium.webdriver as webdriver
#import urllib.request
from bs4 import BeautifulSoup
import contextlib
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.by import By



# LINK: https://www.wolframcloud.com/objects/f51a1d4b-7e02-47cf-8e9f-67b8516bf802?x={.....}

def find_by_xpath(locator):
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, locator))
    )

    return element

def makelist(table):
    result = []
    allrows = table.findAll('tr')
    allrows = filter(None, allrows)
    result = filter(None, result)
    for row in allrows:
        result.append([])
        allcols = row.findAll('td')
        for col in allcols:
            thestrings = [unicode(s) for s in col.findAll(text=True)]
            thetext = ''.join(thestrings)
            result[-1].append(thetext)
            result = filter(None, result)

    data = []

    for i in range(len(result)):
        data.append([])

        for j in range (len(result[1])):
            if result[i][j] != "":
                data[i].append(result[i][j])

    data1 =[]

    for i in range(len(data)):
        data1.append([])

        for j in range (len(data[1])):
            if data[i][j] != ">":
                data1[i].append(data[i][j])

    return data1

def getFormat(workbook, style, wide):
    if style == "header1":
        format = workbook.add_format({'bold': 1})
        format.set_bg_color('#000080')
        format.set_font_color('white')
        format.set_bottom()
        format.set_top()
        format.set_font_name("Times New Roman")




    elif style == "header2":
        format = workbook.add_format({'bold': 1})
        format.set_bg_color('blue')
        format.set_font_color('white')
        format.set_bottom()
        format.set_top()
        format.set_font_name("Times New Roman")


    elif style == "header3":
        format = workbook.add_format({'bold': 1})
        format.set_bg_color('blue')
        format.set_font_color('white')
        format.set_bottom()
        format.set_top()
        format.set_font_name("Times New Roman")

    elif style == "dollars_top":
        format = workbook.add_format({'num_format': '#,##0'})
        format.set_bg_color('white')
        format.set_font_color('black')
        format.set_font_name("Times New Roman")

    elif style == "dollars":
        format = workbook.add_format({'num_format': '#,##0'})
        format.set_bg_color('white')
        format.set_font_color('black')
        format.set_font_name("Times New Roman")

    elif style == "percentage":
        format = workbook.add_format({'num_format': '0.0%'})
        format.set_bg_color('white')
        format.set_font_color('black')
        format.set_font_name("Times New Roman")

    else:
        format = workbook.add_format()
        format.set_bg_color('white')
        format.set_font_name("Times New Roman")

    return format

def getCell(row, col):
    letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

    colAddress = letters[col]
    rowAddress = row + 1

    return colAddress + str(rowAddress)


def get_statements(ticker):

    if len(ticker) == 1:
        ticker += "  "
    elif len(ticker) == 2:
        ticker += " "

    driver = webdriver.PhantomJS()
    driver.set_window_size(1120, 550)

    driver.delete_all_cookies()
    driver.get("https://www.lib.uwo.ca/cgi-bin/ezpauthenticate.cgi?url=http://www.MergentOnline.com")



    ## LOGIN
    element_user = driver.find_element_by_name("user")
    element_user.send_keys("akhan552")

    with file("pass.txt") as f:
        password = f.read()

    element_password = driver.find_element_by_name("pass")
    element_password.send_keys(password)

    def find_by_xpath(locator):
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, locator))
        )

        return element


    find_by_xpath('//input[@value = "Login"]').click()

    element_ticker = driver.find_element_by_name("searchtext")

    # TODO - GET FROM REQUEST
    element_ticker.send_keys(ticker)
    driver.find_element_by_id("basicsearchbtn").click()

    ret = driver.find_element_by_name("chkopt[]")
    parent = ret.find_element_by_xpath('..')

    # Navigation
    ret1 = parent.find_element_by_xpath('following-sibling::td/a')
    #driver.find_element_by_xpath("//input[@name, 'chkopt[]']/following-sibling")
    link1 = driver.find_element_by_link_text(ret1.text)
    link1.click()

    link2 = driver.find_element_by_link_text("Company Financials")
    link2.click()

    element5 = driver.find_element_by_name("range")
    element5.send_keys("7")
    element5.send_keys("enter")
    element20 = driver.find_element_by_name("scale")
    element20.send_keys("n")
    element20.send_keys("enter")
    element6 = driver.find_element_by_name("Submit").click()

    html_doc = driver.page_source

    soup = BeautifulSoup(html_doc)
    soup.prettify()
    table = soup.find("table",{"id": "table01"})

    BS = makelist(table)

    target = open(ticker + "_BS.csv", 'w')

    for i in range (len(BS)):
        for j in range(len(BS[i])):
            target.write('"' + BS[i][j] + '"')
            target.write(",")

        target.write("\n")

    element11 = driver.find_element_by_name("dataarea")
    element11.send_keys("i")
    element11.send_keys("enter")
    element12 = driver.find_element_by_name("Submit").click()


    html_doc2 = driver.page_source

    soup2 = BeautifulSoup(html_doc2)
    soup2.prettify()

    table2 = soup2.find("table",{"id": "table02"})

    IS = makelist(table2)

    target = open(ticker + "_IS.csv", 'w')

    for i in range (len(IS)):
        for j in range(len(IS[i])):
            target.write('"' + IS[i][j] + '"');
            target.write(",")

        target.write("\n")

    link3 = driver.find_element_by_link_text("Analysis")
    link3.click()

    html_doc3 = driver.page_source

    soup3 = BeautifulSoup(html_doc3)
    soup3.prettify()

    table3 = soup3.find("table",{"class": "tablesorter"})

    analysis = makelist(table3)

    target = open(ticker + "_Analysis.csv", 'w')

    for i in range (len(analysis)):
        for j in range(len(analysis[i])):
            target.write('"' + analysis[i][j] + '"');
            target.write(",")

        target.write("\n")


    link4 = driver.find_element_by_link_text("Ratios")
    link4.click()
    element21 = driver.find_element_by_name("range")
    element21.send_keys("7")
    element21.send_keys("enter")

    html_doc4 = driver.page_source
    soup4 = BeautifulSoup(html_doc4)
    soup4.prettify()

    table4 = soup4.find("table",{"id": "table25"})

    ratios = makelist(table4)
    print ratios

    target = open(ticker + "_Ratios.csv", 'w')

    for i in range (len(ratios)):
        for j in range(len(ratios[i])):
            target.write('"' + ratios[i][j] + '"');
            target.write(",")

        target.write("\n")
    target.close()




    link = driver.find_element_by_link_text("Log Out")
    link.click()

    driver.quit()


    path_to_file = analyze(ticker, ticker + "_IS.csv", ticker + "_BS.csv", ticker +  "_Ratios.csv", ticker + "_Analysis.csv")

    # print path_to_file

    return path_to_file

    #############################

def analyze(ticker, IS_PATH, BS_PATH, RATIO_PATH, ANALYSIS_PATH):


    IS_reader = csv.reader(open(IS_PATH, 'r'))

    IS_rows_reversed = []
    IS_rows = []

    for r in IS_reader:
        IS_rows_reversed.append(r)

    for e in IS_rows_reversed:
        #print e
        pass




    for i in range (len(IS_rows_reversed)):

        IS_rows.append([]);
        if len(IS_rows_reversed[i]) > 8:
             IS_rows_reversed[i] = IS_rows_reversed[i][0:8]



        IS_rows[i].append(IS_rows_reversed[i][0])
        IS_rows[i].append(IS_rows_reversed[i][7])
        IS_rows[i].append(IS_rows_reversed[i][6])
        IS_rows[i].append(IS_rows_reversed[i][5])
        IS_rows[i].append(IS_rows_reversed[i][4])
        IS_rows[i].append(IS_rows_reversed[i][3])
        IS_rows[i].append(IS_rows_reversed[i][2])
        IS_rows[i].append(IS_rows_reversed[i][1])


    for i in IS_rows:
        for j in range (len(i)):
            if j != 0:
                if i[j] == "-":
                    i[j] = "0";


    ############## BS

    BS_reader = csv.reader(open(BS_PATH, 'r'))

    BS_rows_reversed = []
    BS_rows = []

    for r in BS_reader:
        BS_rows_reversed.append(r)

    for e in BS_rows_reversed:
        #print e
        pass




    for i in range (len(BS_rows_reversed)):

        BS_rows.append([]);
        if len(BS_rows_reversed[i]) > 8:
             BS_rows_reversed[i] = BS_rows_reversed[i][0:8]



        BS_rows[i].append(BS_rows_reversed[i][0])
        BS_rows[i].append(BS_rows_reversed[i][7])
        BS_rows[i].append(BS_rows_reversed[i][6])
        BS_rows[i].append(BS_rows_reversed[i][5])
        BS_rows[i].append(BS_rows_reversed[i][4])
        BS_rows[i].append(BS_rows_reversed[i][3])
        BS_rows[i].append(BS_rows_reversed[i][2])
        BS_rows[i].append(BS_rows_reversed[i][1])


    for i in BS_rows:
        for j in range (len(i)):
            if j != 0:
                if i[j] == "-":
                    i[j] = "0";




    ratio_reader = csv.reader(open(RATIO_PATH, 'r'))
    print "reader", ratio_reader


    ratio_rows_reversed = []
    ratio_rows = []

    print "READING"
    for r in ratio_reader:
        print r
        ratio_rows_reversed.append(r)

    print "REVERSED"
    for e in ratio_rows_reversed:
        print e
        pass


    for i in range (len(ratio_rows_reversed)):

        ratio_rows.append([]);
        if len(ratio_rows_reversed[i]) > 8:
             ratio_rows_reversed[i] = ratio_rows_reversed[i][0:8]



        ratio_rows[i].append(ratio_rows_reversed[i][0])
        ratio_rows[i].append(ratio_rows_reversed[i][7])
        ratio_rows[i].append(ratio_rows_reversed[i][6])
        ratio_rows[i].append(ratio_rows_reversed[i][5])
        ratio_rows[i].append(ratio_rows_reversed[i][4])
        ratio_rows[i].append(ratio_rows_reversed[i][3])
        ratio_rows[i].append(ratio_rows_reversed[i][2])
        ratio_rows[i].append(ratio_rows_reversed[i][1])

    print "test1: ", ratio_rows
    for i in ratio_rows:
        for j in range (len(i)):
            if j != 0:
                if i[j] == "-":
                    i[j] = "0";

    print "test"
    for i in ratio_rows:
        print i


    ###################### Analysis SUBTOTALS

    analysis_reader = csv.reader(open(ANALYSIS_PATH, 'r'))

    analysis_rows = []

    for r in analysis_reader:
        analysis_rows.append(r)

    ## TEST PRINTING ##

    # IS_rows, BS_rows, ratio_rows, analysis

    # for i in IS_rows:
        # print i

    # print ""
    # print ""
    # print ""
    # print ""

    # for i in BS_rows:
        # print i

    # print ""
    # print ""
    # print ""
    # print ""

    # for i in ratio_rows:
    #     print i

    ratio_rows1 = ratio_rows[:]
    print "ratiorows1"
    print ratio_rows1

    # print ""
    # print ""
    # print ""
    # print ""


    # for i in analysis_rows:
    #     print i



    # print ""
    # print ""
    # print ""
    # print ""


    ####### BS/IS #######

    workbook = xlsxwriter.Workbook(ticker + '.xlsx')
    DCF_worksheet = workbook.add_worksheet("DCF")

    FS_worksheet = workbook.add_worksheet("Financial Statements")

    format = workbook.add_format()

    format.set_pattern(1)  # This is optional when using a solid fill.
    format.set_bg_color('white')

    format1 = workbook.add_format()
    format1.set_font_size(30)

    header1 = getFormat(workbook, "header1", False)
    header1Wide = getFormat(workbook, "header1", True)

    header2 = getFormat(workbook, "header2", False)
    header2Wide = getFormat(workbook, "header2", True)

    header3 = getFormat(workbook, "header3", False)
    header3Wide = getFormat(workbook, "header3", True)

    dollar_top = getFormat(workbook, "dollars_top", False)
    dollar = getFormat(workbook, "dollars", False)

    percentage_top = getFormat(workbook, "percentage_top", False)
    percentage = getFormat(workbook, "percentage", False)

    regularWide = getFormat(workbook, "header1", True)


    IS = 6
    IS_length = 0

    DCF = IS + IS_length + 3




    for col in range(26):
        for row in range (1000):
            FS_worksheet.write(row,col, '', format)


    bold = workbook.add_format({'bold': 1})
    bold.set_bg_color('white')
    bold.set_font_name("Times New Roman")


    FS_worksheet.set_column(0, 0, 1)
    FS_worksheet.set_column(1, 1, 45)
    FS_worksheet.set_column(2, 2, 16)
    FS_worksheet.set_column(3, 3, 16)
    FS_worksheet.set_column(4, 4, 16)
    FS_worksheet.set_column(5, 5, 16)
    FS_worksheet.set_column(6, 6, 16)
    FS_worksheet.set_column(7, 7, 16)
    FS_worksheet.set_column(8, 8, 16)
    FS_worksheet.set_column(9, 9, 16)
    FS_worksheet.set_column(10, 10, 16)
    FS_worksheet.set_column(11, 11, 16)
    FS_worksheet.set_column(12, 12, 16)
    FS_worksheet.set_column(13, 13, 16)

    driver2 = webdriver.PhantomJS()
    driver2.set_window_size(1120, 550)

    driver2.delete_all_cookies()
    driver2.get("https://www.google.com/finance")
    element150 = driver2.find_element_by_id("gbqfq")
    element150.send_keys(ticker)
    driver2.find_element_by_class_name("gbqfb").click()
    driver2.get_screenshot_as_file('mhacks200.png')

    element200 = driver2.find_element_by_xpath('.//span[@class = "pr"]').text
    element201 = driver2.find_element_by_xpath('.//span[@class = "nwp"]').text
    #element200 = driver2.find_element_by_id("price-panel")


    driver2.quit()



    FS_worksheet.write(1,1, 'Microsoft Corporation', bold)
    FS_worksheet.write(2,1, 'Financial Statements', bold)
    FS_worksheet.write(2,3, 'Stock Price:', bold)
    FS_worksheet.write(2,4, element200, format1)
    FS_worksheet.write(3,3, 'Date Compiled :', bold)
    FS_worksheet.write(3,4, element201, bold)

    FS_worksheet.write(IS, 1, 'Income Statement', header1)

    historical_years = 7
    projected_years = 5

    total_years = historical_years + projected_years

    FS_worksheet.write(7,1, 'Account Name', bold)
    FS_worksheet.write(7,2, '2009', bold)
    FS_worksheet.write(7,3, '2010', bold)
    FS_worksheet.write(7,4, '2011', bold)
    FS_worksheet.write(7,5, '2012', bold)
    FS_worksheet.write(7,6, '2013', bold)
    FS_worksheet.write(7,7, '2014', bold)
    FS_worksheet.write(7,8, '2015', bold)
    IS_length += 1



    for i in range (2, 9):
        FS_worksheet.write(IS, i, '', header1)
        row = 1
        col = 0
        # IS_length += 1

        for j in IS_rows:
            FS_worksheet.write(row, col + 1, IS_rows[row][col], bold)


    for j in range(0,len(IS_rows)):
        IS_length += 1

        for x in range(0,8):
            if j == 0:
                FS_worksheet.write(IS + 3 + j, x + 1, IS_rows[j][x], dollar_top)
            else:
                FS_worksheet.write(IS + 3 + j, x + 1, IS_rows[j][x], dollar)

    BS_position = IS + IS_length + 5

    FS_worksheet.write(BS_position, 1, 'Balance Sheet', header1)

    for i in range (2, 9):
        FS_worksheet.write(BS_position, i, '', header1)
        row = 1
        col = 0
        # IS_length += 1
        for j in BS_rows:
            FS_worksheet.write(row, col + 1, BS_rows[row][col], dollar)


    for j in range(0,len(BS_rows)):

        for x in range(0,8):
            if j == 0:
                FS_worksheet.write(BS_position + 3 + j, x + 1, BS_rows[j][x], dollar_top)
            else:
                FS_worksheet.write(BS_position + 3 + j, x + 1, BS_rows[j][x], dollar)


    FS_worksheet.write(BS_position + 1,1, 'Account Name', bold)
    FS_worksheet.write(BS_position + 1,2, '2009', bold)
    FS_worksheet.write(BS_position + 1,3, '2010', bold)
    FS_worksheet.write(BS_position + 1,4, '2011', bold)
    FS_worksheet.write(BS_position + 1,5, '2012', bold)
    FS_worksheet.write(BS_position + 1,6, '2013', bold)
    FS_worksheet.write(BS_position + 1,7, '2014', bold)
    FS_worksheet.write(BS_position + 1,8, '2015', bold)

    FS_worksheet.write(1,1, '', bold)

    ####### RATIO SUMMARY ########



    ratio_position = 6

    ratio_worksheet = workbook.add_worksheet("Ratio Summary")

    ratio_worksheet.set_column(0, 0, 1)
    ratio_worksheet.set_column(1, 1, 45)
    ratio_worksheet.set_column(2, 2, 16)
    ratio_worksheet.set_column(3, 3, 16)
    ratio_worksheet.set_column(4, 4, 16)
    ratio_worksheet.set_column(5, 5, 16)
    ratio_worksheet.set_column(6, 6, 16)
    ratio_worksheet.set_column(7, 7, 16)
    ratio_worksheet.set_column(8, 8, 16)
    ratio_worksheet.set_column(9, 9, 16)
    ratio_worksheet.set_column(10, 10, 16)
    ratio_worksheet.set_column(11, 11, 16)
    ratio_worksheet.set_column(12, 12, 16)
    ratio_worksheet.set_column(13, 13, 16)

    for col in range(26):
        for row in range (1000):
            ratio_worksheet.write(row,col, '', format)

    ratio_position +=1


    ratio_position +=1

    ratio_worksheet.write(1,1, 'Microsoft Corporation', bold)
    ratio_worksheet.write(2,1, 'Key Ratio Summary', bold)

    ratio_worksheet.write(2,3, 'Stock Price:', bold)
    ratio_worksheet.write(2,4, element200, format1)
    ratio_worksheet.write(3,3, 'Date Compiled: ', bold)
    ratio_worksheet.write(3,4, element201, bold)

    ratio_worksheet.write(ratio_position, 1, 'Ratio Analysis', header1)


    ratio_worksheet.write(ratio_position + 1, 1, 'Description', bold)
    ratio_worksheet.write(ratio_position + 1, 2, '2009', bold)
    ratio_worksheet.write(ratio_position + 1, 3, '2010', bold)
    ratio_worksheet.write(ratio_position + 1, 4, '2011', bold)
    ratio_worksheet.write(ratio_position + 1, 5, '2012', bold)
    ratio_worksheet.write(ratio_position + 1, 6, '2013', bold)
    ratio_worksheet.write(ratio_position + 1, 7, '2014', bold)
    ratio_worksheet.write(ratio_position + 1, 8, '2015', bold)

    for i in range (2, 9):
        ratio_worksheet.write(ratio_position, i, '', header1)
        row = 1
        col = 0
        # IS_length += 1

        print ratio_rows

        for j in ratio_rows:
            ratio_worksheet.write(row, col + 1, ratio_rows[row][col], bold)


    for j in range(0,len(ratio_rows)):
        ratio_position += 1

        for x in range(0,8):
            if j == 0:
                ratio_worksheet.write(ratio_position + j + 2, x + 1, ratio_rows[j][x], dollar_top)
            else:
                ratio_worksheet.write(ratio_position + j + 2, x + 1, ratio_rows[j][x], dollar)

    ratio_worksheet.write(1,1, '', bold)


    ######### DCF #########

    DCF_position = 6


    DCF_worksheet.set_column(0, 0, 1)
    DCF_worksheet.set_column(1, 1, 45)
    DCF_worksheet.set_column(2, 2, 16)
    DCF_worksheet.set_column(3, 3, 16)
    DCF_worksheet.set_column(4, 4, 16)
    DCF_worksheet.set_column(5, 5, 16)
    DCF_worksheet.set_column(6, 6, 16)
    DCF_worksheet.set_column(7, 7, 16)
    DCF_worksheet.set_column(8, 8, 16)
    DCF_worksheet.set_column(9, 9, 16)
    DCF_worksheet.set_column(10, 10, 16)
    DCF_worksheet.set_column(11, 11, 16)
    DCF_worksheet.set_column(12, 12, 16)
    DCF_worksheet.set_column(13, 13, 16)

    for col in range(26):
        for row in range (1000):
            DCF_worksheet.write(row,col, '', format)

    DCF_position +=1

    DCF_worksheet.write(DCF_position + 1,1, 'Account Name', bold)
    DCF_worksheet.write(DCF_position + 1,2, '2009', bold)
    DCF_worksheet.write(DCF_position + 1,3, '2010', bold)
    DCF_worksheet.write(DCF_position + 1,4, '2011', bold)
    DCF_worksheet.write(DCF_position + 1,5, '2012', bold)
    DCF_worksheet.write(DCF_position + 1,6, '2013', bold)
    DCF_worksheet.write(DCF_position + 1,7, '2014', bold)
    DCF_worksheet.write(DCF_position + 1,8, '2015', bold)



    DCF_worksheet.write(1,1, 'Microsoft Corporation', bold)
    DCF_worksheet.write(2,1, 'Discounted Cash Flow Analysis', bold)
    DCF_worksheet.write(2,3, 'Stock Price:', bold)
    DCF_worksheet.write(2,4, element200, format1)
    DCF_worksheet.write(3,3, 'Date Compiled:', bold)
    DCF_worksheet.write(3,4, element201, bold)

    DCF_worksheet.write(DCF_position, 1, 'Discounted Cash Flow Analysis', header1)

    for i in range (2, 1 + 11):
        DCF_worksheet.write(DCF_position, i, '', header1)
        row = 1
        col = 0

    DCF_worksheet.write(DCF_position + 1, 1, 'Description', bold)
    DCF_worksheet.write(DCF_position + 1, 2, '2013', bold)
    DCF_worksheet.write(DCF_position + 1, 3, '2014', bold)
    DCF_worksheet.write(DCF_position + 1, 4, '2015E', bold)
    DCF_worksheet.write(DCF_position + 1, 5, '2016E', bold)
    DCF_worksheet.write(DCF_position + 1, 6, '2017E', bold)
    DCF_worksheet.write(DCF_position + 1, 7, '2018E', bold)
    DCF_worksheet.write(DCF_position + 1, 8, '2019E', bold)
    DCF_worksheet.write(DCF_position + 1, 9, '2020E', bold)
    DCF_worksheet.write(DCF_position + 1, 10, '2021E', bold)
    DCF_worksheet.write(DCF_position + 1, 11, '2022E', bold)



    DCF_position += 1

    # REVENUES

    DCF_worksheet.write(DCF_position + 1, 1, analysis_rows[0][0], dollar) # REVENUES
    DCF_worksheet.write(DCF_position + 1, 2, analysis_rows[0][2], dollar) # 2 YEARS AGO
    DCF_worksheet.write(DCF_position + 1, 3, analysis_rows[0][1], dollar) # LAST YEAR


    DCF_worksheet.write(DCF_position + 2, 1, "Growth", dollar)
    DCF_worksheet.write(DCF_position + 2, 2, "", dollar)
    DCF_worksheet.write(DCF_position + 2, 3, "=" + getCell(DCF_position + 1, 3) + "/" + getCell(DCF_position + 1, 2) + "-1", percentage)

    # projected growth
    DCF_worksheet.write(DCF_position + 2, 4, '=' + getCell(DCF_position + 2, 3) + "*0.6", percentage)
    DCF_worksheet.write(DCF_position + 2, 5, '=' + getCell(DCF_position + 2, 4), percentage)
    DCF_worksheet.write(DCF_position + 2, 6, '=' + getCell(DCF_position + 2, 5), percentage)
    DCF_worksheet.write(DCF_position + 2, 7, '=' + getCell(DCF_position + 2, 6), percentage)
    DCF_worksheet.write(DCF_position + 2, 8, '=' + getCell(DCF_position + 2, 7), percentage)
    DCF_worksheet.write(DCF_position + 2, 9, '=' + getCell(DCF_position + 2, 8), percentage)
    DCF_worksheet.write(DCF_position + 2, 10, '=' + getCell(DCF_position + 2, 9), percentage)
    DCF_worksheet.write(DCF_position + 2, 11, '=' + getCell(DCF_position + 2, 10), percentage)

    # resulting revenues
    DCF_worksheet.write(DCF_position + 1, 4, '=(1+' + getCell(DCF_position + 2, 4) + ")*" + getCell(DCF_position + 1, 3), dollar)
    DCF_worksheet.write(DCF_position + 1, 5, '=(1+' + getCell(DCF_position + 2, 5) + ")*" + getCell(DCF_position + 1, 4), dollar)
    DCF_worksheet.write(DCF_position + 1, 6, '=(1+' + getCell(DCF_position + 2, 6) + ")*" + getCell(DCF_position + 1, 5), dollar)
    DCF_worksheet.write(DCF_position + 1, 7, '=(1+' + getCell(DCF_position + 2, 7) + ")*" + getCell(DCF_position + 1, 6), dollar)
    DCF_worksheet.write(DCF_position + 1, 8, '=(1+' + getCell(DCF_position + 2, 8) + ")*" + getCell(DCF_position + 1, 7), dollar)
    DCF_worksheet.write(DCF_position + 1, 9, '=(1+' + getCell(DCF_position + 2, 9) + ")*" + getCell(DCF_position + 1, 8), dollar)
    DCF_worksheet.write(DCF_position + 1, 10, '=(1+' + getCell(DCF_position + 2, 10) + ")*" + getCell(DCF_position + 1, 9), dollar)
    DCF_worksheet.write(DCF_position + 1, 11, '=(1+' + getCell(DCF_position + 2, 11) + ")*" + getCell(DCF_position + 1, 10), dollar)

    DCF_position += 2


    # DIRECT COSTS


    DCF_worksheet.write(DCF_position + 1, 1, analysis_rows[1][0], dollar) # REVENUES
    DCF_worksheet.write(DCF_position + 1, 2, analysis_rows[1][2], dollar) # 2 YEARS AGO
    DCF_worksheet.write(DCF_position + 1, 3, analysis_rows[1][1], dollar) # LAST YEAR


    DCF_worksheet.write(DCF_position + 2, 1, "% of Sales", dollar)
    DCF_worksheet.write(DCF_position + 2, 2, "=" + getCell(DCF_position + 1, 2) + "/" + getCell(DCF_position - 1, 2), percentage)
    DCF_worksheet.write(DCF_position + 2, 3, "=" + getCell(DCF_position + 1, 3) + "/" + getCell(DCF_position - 1, 3), percentage)

    # projected % of revenue
    DCF_worksheet.write(DCF_position + 2, 4, '=' + getCell(DCF_position + 2, 3), percentage)
    DCF_worksheet.write(DCF_position + 2, 5, '=' + getCell(DCF_position + 2, 4), percentage)
    DCF_worksheet.write(DCF_position + 2, 6, '=' + getCell(DCF_position + 2, 5), percentage)
    DCF_worksheet.write(DCF_position + 2, 7, '=' + getCell(DCF_position + 2, 6), percentage)
    DCF_worksheet.write(DCF_position + 2, 8, '=' + getCell(DCF_position + 2, 7), percentage)
    DCF_worksheet.write(DCF_position + 2, 9, '=' + getCell(DCF_position + 2, 8), percentage)
    DCF_worksheet.write(DCF_position + 2, 10, '=' + getCell(DCF_position + 2, 9), percentage)
    DCF_worksheet.write(DCF_position + 2, 11, '=' + getCell(DCF_position + 2, 10), percentage)

    # resulting costs
    DCF_worksheet.write(DCF_position + 1, 4, '=' + getCell(DCF_position + 2, 4) + "*" + getCell(DCF_position - 1, 4), dollar)
    DCF_worksheet.write(DCF_position + 1, 5, '=' + getCell(DCF_position + 2, 5) + "*" + getCell(DCF_position - 1, 5), dollar)
    DCF_worksheet.write(DCF_position + 1, 6, '=' + getCell(DCF_position + 2, 6) + "*" + getCell(DCF_position - 1, 6), dollar)
    DCF_worksheet.write(DCF_position + 1, 7, '=' + getCell(DCF_position + 2, 7) + "*" + getCell(DCF_position - 1, 7), dollar)
    DCF_worksheet.write(DCF_position + 1, 8, '=' + getCell(DCF_position + 2, 8) + "*" + getCell(DCF_position - 1, 8), dollar)
    DCF_worksheet.write(DCF_position + 1, 9, '=' + getCell(DCF_position + 2, 9) + "*" + getCell(DCF_position - 1, 9), dollar)
    DCF_worksheet.write(DCF_position + 1, 10, '=' + getCell(DCF_position + 2, 10) + "*" + getCell(DCF_position - 1, 10), dollar)
    DCF_worksheet.write(DCF_position + 1, 11, '=' + getCell(DCF_position + 2, 11) + "*" + getCell(DCF_position - 1, 11), dollar)


    DCF_position += 2


    # GROSS PROFIT


    DCF_worksheet.write(DCF_position + 1, 1, analysis_rows[2][0], dollar) # REVENUES
    DCF_worksheet.write(DCF_position + 1, 2, analysis_rows[2][2], dollar) # 2 YEARS AGO
    DCF_worksheet.write(DCF_position + 1, 3, analysis_rows[2][1], dollar) # LAST YEAR


    DCF_worksheet.write(DCF_position + 2, 1, "% of Sales", dollar)
    DCF_worksheet.write(DCF_position + 2, 2, "=" + getCell(DCF_position + 1, 2) + "/" + getCell(DCF_position - 3, 2), percentage)
    DCF_worksheet.write(DCF_position + 2, 3, "=" + getCell(DCF_position + 1, 3) + "/" + getCell(DCF_position - 3, 3), percentage)
    DCF_worksheet.write(DCF_position + 2, 4, "=" + getCell(DCF_position + 1, 4) + "/" + getCell(DCF_position - 3, 4), percentage)
    DCF_worksheet.write(DCF_position + 2, 5, "=" + getCell(DCF_position + 1, 5) + "/" + getCell(DCF_position - 3, 5), percentage)
    DCF_worksheet.write(DCF_position + 2, 6, "=" + getCell(DCF_position + 1, 6) + "/" + getCell(DCF_position - 3, 6), percentage)
    DCF_worksheet.write(DCF_position + 2, 7, "=" + getCell(DCF_position + 1, 7) + "/" + getCell(DCF_position - 3, 7), percentage)
    DCF_worksheet.write(DCF_position + 2, 8, "=" + getCell(DCF_position + 1, 8) + "/" + getCell(DCF_position - 3, 8), percentage)
    DCF_worksheet.write(DCF_position + 2, 9, "=" + getCell(DCF_position + 1, 9) + "/" + getCell(DCF_position - 3, 9), percentage)
    DCF_worksheet.write(DCF_position + 2, 10, "=" + getCell(DCF_position + 1, 10) + "/" + getCell(DCF_position - 3, 10), percentage)
    DCF_worksheet.write(DCF_position + 2, 11, "=" + getCell(DCF_position + 1, 11) + "/" + getCell(DCF_position - 3, 11), percentage)


    # resulting costs
    DCF_worksheet.write(DCF_position + 1, 4, '=' + getCell(DCF_position - 3, 4) + "-" + getCell(DCF_position - 1, 4), dollar)
    DCF_worksheet.write(DCF_position + 1, 5, '=' + getCell(DCF_position - 3, 5) + "-" + getCell(DCF_position - 1, 5), dollar)
    DCF_worksheet.write(DCF_position + 1, 6, '=' + getCell(DCF_position - 3, 6) + "-" + getCell(DCF_position - 1, 6), dollar)
    DCF_worksheet.write(DCF_position + 1, 7, '=' + getCell(DCF_position - 3, 7) + "-" + getCell(DCF_position - 1, 7), dollar)
    DCF_worksheet.write(DCF_position + 1, 8, '=' + getCell(DCF_position - 3, 8) + "-" + getCell(DCF_position - 1, 8), dollar)
    DCF_worksheet.write(DCF_position + 1, 9, '=' + getCell(DCF_position - 3, 9) + "-" + getCell(DCF_position - 1, 9), dollar)
    DCF_worksheet.write(DCF_position + 1, 10, '=' + getCell(DCF_position - 3, 10) + "-" + getCell(DCF_position - 1, 10), dollar)
    DCF_worksheet.write(DCF_position + 1, 11, '=' + getCell(DCF_position - 3, 11) + "-" + getCell(DCF_position - 1, 11), dollar)


    DCF_position += 4

    # Taxation


    DCF_worksheet.write(DCF_position + 1, 1, analysis_rows[3][0], dollar) # REVENUES
    DCF_worksheet.write(DCF_position + 1, 2, analysis_rows[3][2], dollar) # 2 YEARS AGO
    DCF_worksheet.write(DCF_position + 1, 3, analysis_rows[3][1], dollar) # LAST YEAR


    DCF_worksheet.write(DCF_position + 2, 1, "% of Sales", dollar)
    DCF_worksheet.write(DCF_position + 2, 2, "=" + getCell(DCF_position + 1, 2) + "/" + getCell(DCF_position - 7, 2), percentage)
    DCF_worksheet.write(DCF_position + 2, 3, "=" + getCell(DCF_position + 1, 3) + "/" + getCell(DCF_position - 7, 3), percentage)

    # projected % of revenue
    DCF_worksheet.write(DCF_position + 2, 4, '=' + getCell(DCF_position + 2, 3), percentage)
    DCF_worksheet.write(DCF_position + 2, 5, '=' + getCell(DCF_position + 2, 4), percentage)
    DCF_worksheet.write(DCF_position + 2, 6, '=' + getCell(DCF_position + 2, 5), percentage)
    DCF_worksheet.write(DCF_position + 2, 7, '=' + getCell(DCF_position + 2, 6), percentage)
    DCF_worksheet.write(DCF_position + 2, 8, '=' + getCell(DCF_position + 2, 7), percentage)
    DCF_worksheet.write(DCF_position + 2, 9, '=' + getCell(DCF_position + 2, 8), percentage)
    DCF_worksheet.write(DCF_position + 2, 10, '=' + getCell(DCF_position + 2, 9), percentage)
    DCF_worksheet.write(DCF_position + 2, 11, '=' + getCell(DCF_position + 2, 10), percentage)

    # resulting costs
    DCF_worksheet.write(DCF_position + 1, 4, '=' + getCell(DCF_position + 2, 4) + "*" + getCell(DCF_position - 7, 4), dollar)
    DCF_worksheet.write(DCF_position + 1, 5, '=' + getCell(DCF_position + 2, 5) + "*" + getCell(DCF_position - 7, 5), dollar)
    DCF_worksheet.write(DCF_position + 1, 6, '=' + getCell(DCF_position + 2, 6) + "*" + getCell(DCF_position - 7, 6), dollar)
    DCF_worksheet.write(DCF_position + 1, 7, '=' + getCell(DCF_position + 2, 7) + "*" + getCell(DCF_position - 7, 7), dollar)
    DCF_worksheet.write(DCF_position + 1, 8, '=' + getCell(DCF_position + 2, 8) + "*" + getCell(DCF_position - 7, 8), dollar)
    DCF_worksheet.write(DCF_position + 1, 9, '=' + getCell(DCF_position + 2, 9) + "*" + getCell(DCF_position - 7, 9), dollar)
    DCF_worksheet.write(DCF_position + 1, 10, '=' + getCell(DCF_position + 2, 10) + "*" + getCell(DCF_position - 7, 10), dollar)
    DCF_worksheet.write(DCF_position + 1, 11, '=' + getCell(DCF_position + 2, 11) + "*" + getCell(DCF_position - 7, 11), dollar)

    DCF_position += 2



    # Net Income AND Fixed Costs

    # NI HISTORICAL
    DCF_worksheet.write(DCF_position + 1, 1, analysis_rows[4][0], dollar) # REVENUES
    DCF_worksheet.write(DCF_position + 1, 2, analysis_rows[4][2], dollar) # 2 YEARS AGO
    DCF_worksheet.write(DCF_position + 1, 3, analysis_rows[4][1], dollar) # LAST YEAR

    # FC HISTORICAL
    DCF_worksheet.write(DCF_position - 3, 1, "Fixed Costs", dollar)
    DCF_worksheet.write(DCF_position - 3, 2, '=' + getCell(DCF_position - 5, 2) + "-" + getCell(DCF_position + 1, 2) + "-" + getCell(DCF_position - 1, 2), dollar)
    DCF_worksheet.write(DCF_position - 3, 3, '=' + getCell(DCF_position - 5, 3) + "-" + getCell(DCF_position + 1, 3) + "-" + getCell(DCF_position - 1, 3), dollar)

    # FC % revenue

    DCF_worksheet.write(DCF_position - 2, 1, "% of Sales", dollar)
    DCF_worksheet.write(DCF_position - 2, 2, "=" + getCell(DCF_position - 3, 2) + "/" + getCell(DCF_position - 9, 2), percentage)
    DCF_worksheet.write(DCF_position - 2, 3, "=" + getCell(DCF_position - 3, 3) + "/" + getCell(DCF_position - 9, 3), percentage)
    DCF_worksheet.write(DCF_position - 2, 4, "=" + getCell(DCF_position - 3, 4) + "/" + getCell(DCF_position - 9, 4), percentage)
    DCF_worksheet.write(DCF_position - 2, 5, "=" + getCell(DCF_position - 3, 5) + "/" + getCell(DCF_position - 9, 5), percentage)
    DCF_worksheet.write(DCF_position - 2, 6, "=" + getCell(DCF_position - 3, 6) + "/" + getCell(DCF_position - 9, 6), percentage)
    DCF_worksheet.write(DCF_position - 2, 7, "=" + getCell(DCF_position - 3, 7) + "/" + getCell(DCF_position - 9, 7), percentage)
    DCF_worksheet.write(DCF_position - 2, 8, "=" + getCell(DCF_position - 3, 8) + "/" + getCell(DCF_position - 9, 8), percentage)
    DCF_worksheet.write(DCF_position - 2, 9, "=" + getCell(DCF_position - 3, 9) + "/" + getCell(DCF_position - 9, 9), percentage)
    DCF_worksheet.write(DCF_position - 2, 10, "=" + getCell(DCF_position - 3, 10) + "/" + getCell(DCF_position - 9, 10), percentage)
    DCF_worksheet.write(DCF_position - 2, 11, "=" + getCell(DCF_position - 3, 11) + "/" + getCell(DCF_position - 9, 11), percentage)


    DCF_worksheet.write(DCF_position - 3, 4, '=' + getCell(DCF_position - 3, 3), dollar)
    DCF_worksheet.write(DCF_position - 3, 5, '=' + getCell(DCF_position - 3, 4), dollar)
    DCF_worksheet.write(DCF_position - 3, 6, '=' + getCell(DCF_position - 3, 5), dollar)
    DCF_worksheet.write(DCF_position - 3, 7, '=' + getCell(DCF_position - 3, 6), dollar)
    DCF_worksheet.write(DCF_position - 3, 8, '=' + getCell(DCF_position - 3, 7), dollar)
    DCF_worksheet.write(DCF_position - 3, 9, '=' + getCell(DCF_position - 3, 8), dollar)
    DCF_worksheet.write(DCF_position - 3, 10, '=' + getCell(DCF_position - 3, 9), dollar)
    DCF_worksheet.write(DCF_position - 3, 11, '=' + getCell(DCF_position - 3, 10), dollar)

    # net income

    DCF_worksheet.write(DCF_position + 1, 4, '=' + getCell(DCF_position - 5, 4) + "-" + getCell(DCF_position - 3, 4) + "-" + getCell(DCF_position - 1, 4), dollar)
    DCF_worksheet.write(DCF_position + 1, 5, '=' + getCell(DCF_position - 5, 5) + "-" + getCell(DCF_position - 3, 5) + "-" + getCell(DCF_position - 1, 5), dollar)
    DCF_worksheet.write(DCF_position + 1, 6, '=' + getCell(DCF_position - 5, 6) + "-" + getCell(DCF_position - 3, 6) + "-" + getCell(DCF_position - 1, 6), dollar)
    DCF_worksheet.write(DCF_position + 1, 7, '=' + getCell(DCF_position - 5, 7) + "-" + getCell(DCF_position - 3, 7) + "-" + getCell(DCF_position - 1, 7), dollar)
    DCF_worksheet.write(DCF_position + 1, 8, '=' + getCell(DCF_position - 5, 8) + "-" + getCell(DCF_position - 3, 8) + "-" + getCell(DCF_position - 1, 8), dollar)
    DCF_worksheet.write(DCF_position + 1, 9, '=' + getCell(DCF_position - 5, 9) + "-" + getCell(DCF_position - 3, 9) + "-" + getCell(DCF_position - 1, 9), dollar)
    DCF_worksheet.write(DCF_position + 1, 10, '=' + getCell(DCF_position - 5, 10) + "-" + getCell(DCF_position - 3, 10) + "-" + getCell(DCF_position - 1, 10), dollar)
    DCF_worksheet.write(DCF_position + 1, 11, '=' + getCell(DCF_position - 5, 11) + "-" + getCell(DCF_position - 3, 11) + "-" + getCell(DCF_position - 1, 11), dollar)

    DCF_worksheet.write(1,1, '', bold)

    DCF_position += 2

    #
    # print len(ratio_rows1)


    DCF_worksheet.write(DCF_position + 1, 1, 'Free Cash Flow (Projections)', dollar)
    DCF_worksheet.write(DCF_position + 1, 2, '=' + ratio_rows[3][6] + "/100*" + getCell(DCF_position - 11, 2), dollar)
    DCF_worksheet.write(DCF_position + 1, 3, '=' + ratio_rows[3][7] + "/100*" + getCell(DCF_position - 11, 3), dollar)

    num_ebitda_margins = 0
    avg_ebitda_margins = 0

    for r in range (1, len(ratio_rows[3])):
        # print ratio_rows[3][r]
        avg_ebitda_margins += float(ratio_rows[3][r])

    avg_ebitda_margins /= len(ratio_rows[3])

    # print "avg: " + str(avg_ebitda_margins)

    DCF_worksheet.write(DCF_position + 1, 4, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 4), dollar)
    DCF_worksheet.write(DCF_position + 1, 5, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 5), dollar)
    DCF_worksheet.write(DCF_position + 1, 6, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 6), dollar)
    DCF_worksheet.write(DCF_position + 1, 7, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 7), dollar)
    DCF_worksheet.write(DCF_position + 1, 8, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 8), dollar)
    DCF_worksheet.write(DCF_position + 1, 9, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 9), dollar)
    DCF_worksheet.write(DCF_position + 1, 10, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 10), dollar)
    DCF_worksheet.write(DCF_position + 1, 11, '=' + str(avg_ebitda_margins) + "/100*" + getCell(DCF_position - 11, 11), dollar)

    DCF_position += 2

    DCF_worksheet.write(DCF_position + 1, 1, 'WACC', dollar)
    DCF_worksheet.write(DCF_position + 1, 2, '6.29%', dollar)

    DCF_position += 2

    DCF_worksheet.write(DCF_position + 1, 1, 'Terminal Value', dollar)
    DCF_worksheet.write(DCF_position + 1, 2, '=' + getCell(DCF_position - 3, 11) + "*1.03/(" + getCell(DCF_position - 1, 2) +  "-0.03" + ")", dollar)

    DCF_position += 2

    DCF_worksheet.write(DCF_position + 1, 1, 'Enterprise Value', dollar)
    DCF_worksheet.write(DCF_position + 1, 2, '=NPV(' + getCell(DCF_position - 3, 2) + "," + getCell(DCF_position - 5, 4) + ":" + getCell(DCF_position - 5, 11) + ")+" + getCell(DCF_position - 1, 2) + "/(1+" + getCell(DCF_position - 3, 2) + ")^8" , dollar)


    # DCF_worksheet.write(DCF_position + 2, 1, "% of Sales", dollar)
    # DCF_worksheet.write(DCF_position + 2, 2, "=" + getCell(DCF_position + 1, 2) + "/" + getCell(DCF_position - 9, 2), dollar)
    # DCF_worksheet.write(DCF_position + 2, 3, "=" + getCell(DCF_position + 1, 3) + "/" + getCell(DCF_position - 9, 3), dollar)
    # DCF_worksheet.write(DCF_position + 2, 4, "=" + getCell(DCF_position + 1, 4) + "/" + getCell(DCF_position - 9, 4), dollar)
    # DCF_worksheet.write(DCF_position + 2, 5, "=" + getCell(DCF_position + 1, 5) + "/" + getCell(DCF_position - 9, 5), dollar)
    # DCF_worksheet.write(DCF_position + 2, 6, "=" + getCell(DCF_position + 1, 6) + "/" + getCell(DCF_position - 9, 6), dollar)
    # DCF_worksheet.write(DCF_position + 2, 7, "=" + getCell(DCF_position + 1, 7) + "/" + getCell(DCF_position - 9, 7), dollar)
    # DCF_worksheet.write(DCF_position + 2, 8, "=" + getCell(DCF_position + 1, 8) + "/" + getCell(DCF_position - 9, 8), dollar)
    # DCF_worksheet.write(DCF_position + 2, 9, "=" + getCell(DCF_position + 1, 9) + "/" + getCell(DCF_position - 9, 9), dollar)
    # DCF_worksheet.write(DCF_position + 2, 10, "=" + getCell(DCF_position + 1, 10) + "/" + getCell(DCF_position - 9, 10), dollar)
    # DCF_worksheet.write(DCF_position + 2, 11, "=" + getCell(DCF_position + 1, 11) + "/" + getCell(DCF_position - 9, 11), dollar)
    #
    #
    # # resulting costs
    # DCF_worksheet.write(DCF_position + 1, 4, '=' + getCell(DCF_position - 3, 4) + "-" + getCell(DCF_position - 1, 4), dollar)
    # DCF_worksheet.write(DCF_position + 1, 5, '=' + getCell(DCF_position - 3, 5) + "-" + getCell(DCF_position - 1, 5), dollar)
    # DCF_worksheet.write(DCF_position + 1, 6, '=' + getCell(DCF_position - 3, 6) + "-" + getCell(DCF_position - 1, 6), dollar)
    # DCF_worksheet.write(DCF_position + 1, 7, '=' + getCell(DCF_position - 3, 7) + "-" + getCell(DCF_position - 1, 7), dollar)
    # DCF_worksheet.write(DCF_position + 1, 8, '=' + getCell(DCF_position - 3, 8) + "-" + getCell(DCF_position - 1, 8), dollar)
    # DCF_worksheet.write(DCF_position + 1, 9, '=' + getCell(DCF_position - 3, 9) + "-" + getCell(DCF_position - 1, 9), dollar)
    # DCF_worksheet.write(DCF_position + 1, 10, '=' + getCell(DCF_position - 3, 10) + "-" + getCell(DCF_position - 1, 10), dollar)
    # DCF_worksheet.write(DCF_position + 1, 11, '=' + getCell(DCF_position - 3, 11) + "-" + getCell(DCF_position - 1, 11), dollar)



    workbook.close()

    return ticker + ".xlsx"


from flask import Flask
app = Flask(__name__)

@app.route('/test/')
def hello_world():
    return 'Hello World!'




@app.route('/csv/<path>')
def send_js(path):
    print "server"


    filepath = get_statements(path)

    # print filepath
    return send_from_directory('', filepath)


if __name__ == '__main__':
    app.run()

