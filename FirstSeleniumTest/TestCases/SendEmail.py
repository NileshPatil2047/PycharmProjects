import win32com.client as win32


def send_notification():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'Pawan.Kumar@tieto.com'
    mail.CC = 'nilesh.a.patil@tieto.com'
    mail.Subject = 'Sent through Python'
    mail.body = 'This email alert is auto generated. Please do not respond.'
    mail.send

# ----------------------------------------------------------------------


if __name__ == "__main__":
        send_notification()

