import time
import unittest

from selenium import webdriver
from selenium.webdriver.support.ui import Select

import csv
import win32com.client as win32

import getpass

import os
from os import path
import shutil

username = getpass.getuser()
print username
src = "C:\\Users\\patimnil\\Downloads\\"
dst = "C:\\TimeSheetReports\\"
# task = "03.1-Off - Development"
task = "02.1-On - Development"
# Test Case Class


def csv_dict_reader(file_obj):
        """
        Read a CSV file using csv.DictReader
        """

        for i in range(4):
            file_obj.next()
        reader = csv.DictReader(file_obj, delimiter=',')
        for line in reader:
            if line["Task"] == task:
                print "INVALID DATA FOUND"
                send_notification()
            else:
                print "CORRECT DATA FOUND"

            print(line["Employee"]),
            print(line["Task"])

            def send_notification():
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = 'Pawan.Kumar@tieto.com'
                mail.CC = 'sachin.nerpagare@tieto.com'
                mail.Subject = 'Please check your timesheet <<Sent through Automated email box>>'
                mail.body = 'This email alert is auto generated. Please do not respond.'
                mail.send


class TimeSheetDemoUnitTest1(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Remote(
            command_executor="http://localhost:4444/wd/hub",
            desired_capabilities={
                "browserName": "chrome",
            })

        self.driver.implicitly_wait(30)
        self.driver.maximize_window()

    def test_PyTimeSheetReportGet(self):
        driver = self.driver
        driver.get("https://mytime.tieto.com/")
        driver.set_page_load_timeout(10)
        driver.maximize_window()
        self.assertIn("My Time", driver.title)
        time.sleep(7)
        elem = driver.find_element_by_xpath("//*[@id='login_windows_button']")
        elem.click()
        time.sleep(4)

        elem = driver.find_element_by_xpath("//span[contains(text(),'Reports')]")
        elem.click()
        time.sleep(7)

        elem = driver.find_element_by_xpath("//*[@id='report_page']/div[2]/ul/li[1]/a")
        elem.click()
        time.sleep(4)

        # select = Select(driver.find_element_by_xpath("// *[ @ id = 'year']"))
        # select.select_by_index('7')

        select_elem = Select(driver.find_element_by_xpath("// *[ @ id = 'year']"))
        # print [o.text for o in select_elem.options]
        select_elem.select_by_visible_text(u'2018')

        month = Select(driver.find_element_by_xpath("//*[@id='month']"))
        # print [p.text for p in month.options]
        month.select_by_visible_text(u'June')

        elem = driver.find_element_by_xpath("// *[ @ id = 'search_box'] / form / input")
        elem.click()
        time.sleep(4)
        filename = "C:\Users\kumarpaw\Downloads\mytime_emp_rh_201806_all_projects.csv"
        if os.path.isfile(filename):
            print "File already Exist.......Deleting Automatically from the directory"
            os.remove(filename)
        elem = driver.find_element_by_xpath("// *[ @ id = 'export'] / input")
        elem.click()
        time.sleep(9)

    def tearDown(self):

        files = [i for i in os.listdir(src) if i.startswith("mytime") and path.isfile(path.join(src, i))]
        for f in files:
            shutil.copy(path.join(src, f), dst)
            with open("C:\\TimeSheetReports\\mytime_emp_rh_201806_all_projects.csv") as f_obj:
                csv_dict_reader(f_obj)
        self.driver.close()
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()
