import csv
import win32com.client as win32
# ----------------------------------------------------------------------
# -*- coding: utf-8 -*-
"""
@author: Nilesh Kumar Patil

"""


def csv_dict_reader(file_obj):

    """
    Read a CSV file using csv.DictReader
    """
    reader = csv.DictReader(file_obj, delimiter=',')
    for line in reader:
        if line["Task"] == "03.1-On - Development":
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
    mail.CC = 'nilesh.a.patil@tieto.com'
    mail.Subject = 'Invalid Data Found <<Sent through Automated email box>>'
    mail.body = 'This email alert is auto generated. Please do not respond.'
    mail.send


# ----------------------------------------------------------------------
if __name__ == "__main__":
    with open("C:\\PycharmProjects\\mytime_emp_rh_201806_all_projects.csv") as f_obj:
        csv_dict_reader(f_obj)


