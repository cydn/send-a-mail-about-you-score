# 该程序可批量发放邮件。可通过更改文件夹下.xlsx文件内容给相应邮箱发放相应内容。该程序可方便老师/助教通过邮箱发布学生学科成绩
# 不同发件人，不同接收者的情况下，除了更改文件夹下.xlsx文件之外，同时需要改动本程序内部部分：
# 1）请将发件人邮箱账号和发件人序列号做出相应更改。发件人可自行访问邮箱网页版获得该序列号（激活码）
# 2）请在明显标注需要修改的地方，将‘max_row’的值改为n+1，n为学生总数
# 3）请在程序最后将学生总数改至提示的地方
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr

my_sender = '1367744830@qq.com'  # 发件人邮箱账号
my_pass = 'zozcxxzqexafgdfj'  # 发件人序列号（激活码？）总之不是邮箱密码
my_user = []
contents = []

wb = openpyxl.load_workbook('try-xlsxIO.xlsx')
ws = wb.active

###############################################此处须有修改#################################################
for each_row in ws.iter_rows(min_row=2, min_col=2, max_row=3, max_col=2):  # 若有n个学生，将max_row改为n+1
    my_user.append(each_row[0].value)

################################################此处须有修改#################################################
for each_row in ws.iter_rows(min_row=2, min_col=3, max_row=3, max_col=3):  # 若有n个学生，将max_row改为n+1
    contents.append('你的期中考试成绩是'+str(each_row[0].value)+'分')


def mail(k, contains):
    ret = True
    try:
        msg = MIMEText(contains, 'plain', 'utf-8')
        msg['From'] = formataddr(["From TC", my_sender])
        msg['To'] = formataddr(["   ", my_user[k]])
        msg['Subject'] = "试一试用python群发邮件"

        server = smtplib.SMTP_SSL("smtp.qq.com", 465)
        server.login(my_sender, my_pass)
        server.sendmail(my_sender, [my_user[k], ], msg.as_string())
        server.quit()
    except Exception:
        ret = False
    return ret


######################请将（0，2）处的2更改为学生总数##############################
for i in range(0, 2):
    ret = mail(i, contents[i])
    if ret:
        print("邮件发送成功")
    else:
        print("邮件发送失败")
