# mail server module
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
import os

# time module
import datetime
import time

# pw input module
import getpass

now = datetime.datetime.now()
date = now.strftime("%Y.%m.%d")

# id select
with open(r"C:\Users\user\Desktop\상황실이메일매크로\idlists.txt") as f:
    users = f.readlines()
    totaluser = len(users)
    print(len(users))
    print("Email LIst")
    for idx in range(totaluser):
        users[idx] = users[idx].replace("\n", "")
        print("%d. %s" %(idx, users[idx]))
        
    print("Select your email")
    tmp = int(input())
    while not(0 <= tmp and tmp < totaluser):
        print("Invalid Input")
        print("Select your email")
        tmp = int(input())

    id = users[tmp]
    print("Your email: %s" %id)

me = id

pw = getpass.getpass('Type your pw >')
you = 'sjsj1338@hanyang.ac.kr'

# server setting
naver_server = smtplib.SMTP('smtp.naver.com', 587)
naver_server.starttls()
while True:
    try:
        naver_server.login(me,pw)
        break
    
    except:
        pw = getpass.getpass('Invalid pw! Type your pw > ')

# mail object
msg = MIMEMultipart()
msg['Subject'] = '%s 사회과학대 행정팀입니다.' %date
msg['From'] = me
msg['To'] = you
msg.attach(MIMEText('사회과학대학 강의실 대관 현황입니다.', 'plain'))


excel_path = r'C:\Users\user\Desktop\상활실 이메일파일\%s 관재팀_건물별 공간 개방 요청서(당일).xlsx' %(date)


try:
    # attach excel file
    with open(excel_path, "rb") as fil:
        part = MIMEApplication(
            fil.read(),
            Name=os.path.basename(excel_path)
            )
        part['Content-Disposition'] = 'attachment; filename="%s"' %os.path.basename(excel_path)
        msg.attach(part)

        # close file
        print("send success!")
        time.sleep(2)
        naver_server.sendmail(me, you, msg.as_string())
        naver_server.quit()

except:
    print("There is no Excel file")
    time.sleep(2)    
