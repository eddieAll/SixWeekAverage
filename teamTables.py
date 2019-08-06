import requests
import urllib
import time
import bs4
import csv
import lxml.html as lh
import xlwt
from xlwt import Workbook
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select

from selenium import webdriver

reportLink = ['https://www.runningahead.com/logs/202aac0fe44a40c6be1034f835eb5825/reports', 'https://www.runningahead.com/logs/335d3813cbe74978b848f3ee59e158e9/reports', 'https://www.runningahead.com/logs/f03637766ba147ca990ab76668561408', 'https://www.runningahead.com/logs/5fa13f6b33be41278af6a2eef6b4a71a/reports', 'https://www.runningahead.com/logs/05175ec7523c498797871aeb829a9b76/reports', 'https://www.runningahead.com/logs/a16732be55374821a3de7f233adf5c01/reports', 'https://www.runningahead.com/logs/7691c7a9bf2b4e5492848b31c6e2994a/reports', 'https://www.runningahead.com/logs/918ee902b7104d71ba917b385e38a976/reports', 'https://www.runningahead.com/logs/cbaa1bb4f0c9413fbd64a6fa890704a6/reports', 'https://www.runningahead.com/logs/2b0b054092ed47c7bcc421406d6549dd/reports', 'https://www.runningahead.com/logs/6963924766e54417aac3106f3fa82cd1/reports', 'https://www.runningahead.com/logs/cf99c1f25a5941f6abeaeec92d7d0a2e/reports', 'https://www.runningahead.com/logs/ea58371470e642b38ec89c279ec103df/reports', 'https://www.runningahead.com/logs/d15237357016454b9cd009248baa3ef0/reports', 'https://www.runningahead.com/logs/b03940299c884551a89501e337c74ea9/reports', 'https://www.runningahead.com/logs/7c4ff74339184238bb3cac4c9b35c50e/reports', 'https://www.runningahead.com/logs/1d5d5bf1a7d7464a9947302958a0af66/reports']
athleteName = ['Adam Wolfe', 'Andrew Pizzirusso', 'Bill Angelina', 'Blake Samsel', 'Christopher Myers', 'Christian Schaaf', 'Colin Elliot', 'Colm Smith', 'Donato', 'James Teal', 'Liam Coverdale', 'Marc Ramson', 'Sagar Patel', 'Samuel Gerstenbacher', 'Will Schoener', 'Timothy Witmer', 'Zack Kardos']
selectList = [4, 7, 6, 3, 6, 4, 3, 3, 3, 7, 8, 4, 3, 6, 3, 3, 6]
arrayStart = [130, 130, 127, 127, 133, 127, 130, 127, 129, 130, 127, 130, 127, 128, 127, 130, 130]
arrayEnd = [162]
arrayStep = [5]

'''
For testing the length of things

print("Report Link: ", len(reportLink))
print("athleteName: ", len(athleteName))
print("select list: ", len(selectList))
print("Array start: ", len(arrayStart))
for i in range(len(reportLink)):
    print(athleteName[i], selectList[i], arrayStart[i], reportLink[i])

'''

def my_range(start, end, step):
    while start<=end:
        yield start
        start += step

'''
def is_number(n):
    try:
        float(n)
        return True
    except ValueError:
        return False 
'''
        
j = 0
k = 6
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

browser = webdriver.Safari()
browser.get('https://www.runningahead.com/login?redirect=%2flogs')
emailElem = browser.find_element_by_id('ctl00_ctl00_ctl00_SiteContent_PageContent_MainContent_email')
emailElem.send_keys('EMAIL')
passwordElem = browser.find_element_by_id('ctl00_ctl00_ctl00_SiteContent_PageContent_MainContent_password')
passwordElem.send_keys('PASSWORD')
passwordElem.submit()
time.sleep(5)

for t in range(len(reportLink)):
    
    browser.get(reportLink[t])
    sheet1.write(j, 8, athleteName[t])
    
    
    linkElem = browser.find_element_by_link_text('Reports')
    type(linkElem)
    linkElem.click()
    time.sleep(1)
    
    linkElem = browser.find_element_by_link_text('New Search')
    type(linkElem)
    linkElem.click()
    time.sleep(1)
    
    select_element = Select(browser.find_element_by_id('ctl00_ctl00_ctl00_SiteContent_PageContent_TrainingLogContent_SearchForm_GroupBy'))
    select_element.select_by_index(2)
    select_element = Select(browser.find_element_by_id('ctl00_ctl00_ctl00_SiteContent_PageContent_TrainingLogContent_SearchForm_EventType'))
    select_element.select_by_index(selectList[t])
    
    linkElem = browser.find_element_by_id('ctl00_ctl00_ctl00_SiteContent_PageContent_TrainingLogContent_Search_s')
    type(linkElem)
    linkElem.click()
    
    time.sleep(4)
    
    soup = BeautifulSoup(browser.page_source, 'lxml')
    
    text = soup.get_text()
    text = text.split()
    
    
    
    for i in my_range(arrayStart[t], arrayEnd[0], arrayStep[0]):
        if k < 0:
            break
        sheet1.write(j, k, text[i])
        k-=1
    
    j+=2
    k = 6
    wb.save('team.xls')