#!/usr/bin/python
# -*- coding: utf-8 -*-
import xlwt
from selenium import webdriver
from time import sleep
__author__ = 'Aleksandr Jashhuk, Zoer, R5AM, www.r5am.ru'


url = 'http://zerohedge.talking-forex.com/minimal/live/'
headers = ['ASIA', 'EUROPE', 'UK', 'EUROPE/UK', 'FX', 'COMMODITIES', 'US']
base_xls_filename = 'zerohedge_'

driver = webdriver.PhantomJS(executable_path='phantomjs')
driver.set_window_size(1200, 800)
driver.get(url)
sleep(3)


def eu_search():
    global result, flag
    result = driver.find_elements_by_xpath('.//*/td[2]/h4/em')
    flag = False
    for item in result:
        out = item.text
        if 'EU Morning Call' in out:
            flag = True
    return flag


flag = eu_search()
while not flag:
    result = driver.find_elements_by_xpath('.//*[@id="headlines"]/span[1]/a')[0].click()
    sleep(3)
    print '.',
    flag = eu_search()

result2 = driver.find_elements_by_xpath('.//*[contains(@data-headline-subject,'
                                        '"EU Morning Call")]/td[2]')[0].text
driver.close()
driver.quit()

output = result2.split('\n')

my_workbook = xlwt.Workbook('utf8')         # create Excel-file
ws = my_workbook.add_sheet('Sheet 1')       # work with firs sheet (worksheet)

# fonts for headers
header1_style = xlwt.easyxf('font: name Courier, height 360, color-index red, bold on;'
                            'alignment: wrap on, vertical top, horizontal left;')
header2_style = xlwt.easyxf('font: name Courier, height 320, color-index blue, bold on;'
                            'alignment: wrap on, vertical top, horizontal left;')
# font for text
text_style = xlwt.easyxf('font: name Courier, height 320, color-index black, bold off;'
                         'alignment: shrink off, wrap off, vertical top, horizontal left;')

# width of the columns
ws.col(0).width = 14000
ws.col(1).width = 45000

# height of the rows
for counter in range(0, 50):
    ws.row(counter).height = 380


for counter, content in enumerate(output):
    if counter == 0:
        ws.write(counter, 0, content, header1_style)        # header with data
    elif content in headers:
        ws.write(counter+1, 0, content, header2_style)      # headers
    else:
        ws.write(counter, 1, content, text_style)

# file name from date
xls_filename = base_xls_filename + output[0][-2:] + output[0][-5:-3] + output[0][-8:-6] + '.xls'

# save XLS-file
my_workbook.save(xls_filename)

