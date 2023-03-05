# _*_ coding:utf-8 _*_
import chardet
import xlrd
import xlsxwriter

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


#1.read file
data = xlrd.open_workbook('info.xlsx')
classinfo = [] #class information

for sheet in data.sheets():
    dict={'name':sheet.name, 'avgsalary':0}
    sum_salary = 0
    
    #ergodic sheet nrows and sum the salary
    for i in range(sheet.nrows):
        if(i > 1):
            sum_salary = sum_salary + float(sheet.cell(i, 5).value)  # get salary of each column
            
    dict['avgsalary'] = sum_salary /(sheet.nrows - 2)
    classinfo.append(dict)
    
print(classinfo)

#2. write into a new excel
workbook = xlsxwriter.Workbook('newinfo.xlsx')  # structure a new excel workbook
newsheet = workbook.add_worksheet()             # create a worksheet

nameinfo = []
salaryinfo = []

#ergodic classinfo dict
for item in classinfo:
    nameinfo.append(item['name'])
    salaryinfo.append(item['avgsalary'])

# insert column data   
newsheet.write_column('A1', nameinfo)
newsheet.write_column('B1', salaryinfo)

#insert charts
chart = workbook.add_chart({'type':  'column'})   # column type charts
chart.set_title({'name': 'average salary of employments'})
# data resource
chart.add_series({
    'name': 'class',
    'categories': 'Sheet1!$A$1:$A$3',
    'values': 'Sheet1!$B$1:$B$3'
})

newsheet.insert_chart('A7', chart)

workbook.close()

#3. send email
host_server = 'smtp.qq.com'  #主机地址
#发件人邮箱
sender = '1725818634@qq.com'
#发件人密码，授权码
code = 'eyscsthrcfnueega'

#收件人
user1 = 'haiyangfan99@163.com'

#准备邮件数据
#邮件标题
mail_title = '!!!!!!1月份平均就业薪资'
mail_content = '1月份平均就业薪资,请查看具体附件内容'

# with open('newinfo.xlsx', 'rb') as f:
#     enc = chardet.detect(f.read())
#构建附件
attachment = MIMEApplication(open('newinfo.xlsx', 'rb').read())

#附件头部信息
attachment.add_header('Content-Disposition', 'attachment', filename = 'data.xlsx')

#SMTP
smtp = smtplib.SMTP(host_server)
#登录
smtp.login(sender, code)

msg = MIMEMultipart() #带附件的实例
msg['Subject'] = mail_title
msg['from'] = sender
msg['To'] = user1
msg.attach(MIMEText(mail_content))  #邮件正文内容
msg.attach(attachment)
#发送
smtp.sendmail(sender, user1, msg.as_string())

