import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import reportlab.pdfbase.ttfonts #导入reportlab的注册字体
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr
from email import encoders
from email.mime.base import MIMEBase

def getNames(xlsxfile):
    data = xlrd.open_workbook(xlsxfile)
    print (data.sheet_names())
    table = data.sheets()[0] # 打开第一个工作表
    rows = table.nrows  # 获取行数
    names=table.col_values(0)
    emails=table.col_values(1)
    print(names)
    for i in range(rows):
        if i>0:
            xingming=names[i]
            email1=emails[i]
            print(xingming)
            print(email1)
            putNamesInPDF(xingming)
            sendEmail(xingming,email1)
            i +=1

def putNamesInPDF(xm):
    reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.ttfonts.TTFont('song', 'SimSun.ttf')) #注册字体

    packet = io.BytesIO()
    
    can = canvas.Canvas(packet, pagesize=letter)# create a new PDF with Reportlab

    can.setFont('song',18) #设置字体字号
    can.drawString(95, 625, xm)
    can.save()

    
    packet.seek(0)#move to the beginning of the StringIO buffer
    new_pdf = PdfFileReader(packet)
    
    existing_pdf = PdfFileReader(open("original.pdf", "rb"))# read your existing PDF
    output = PdfFileWriter()
    
    page = existing_pdf.getPage(0)# add the "watermark" (which is the new pdf) on the existing page
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    for p in range(1, existing_pdf.getNumPages()):
        page = existing_pdf.getPage(p)
        output.addPage(page)

    
    outputStream = open("%s录取通知书.pdf" % xm, "wb")# finally, write "output" to a real file
    output.write(outputStream)
    outputStream.close()

        

def sendEmail(mingzi,receiver_email):
    mail_user = '' #这里修改Email地址
    mail_pass = '' #这里填写stmp邮箱密码
    
    receiver = [receiver_email]# 接收邮件
    print(receiver)
    
    message = MIMEMultipart() #创建一个带附件的实例
    message['From'] = Header("测试分发通知书", 'utf-8')
    message['To'] =  Header("测试", 'utf-8')
    
    message['Subject'] = Header('测试ceshi', 'utf-8').encode() #subject
    
    
    message.attach(MIMEText('这是一个测试分发邮件', 'plain', 'utf-8'))#邮件正文内容
    '''
    # 构造附件1，传送当前目录下的 test.txt 文件
    att1 = MIMEText(open('%s录取通知书.pdf' % mingzi, 'rb').read(), 'base64', 'utf-8')
    print("文件读取成功")
    att1["Content-Type"] = 'application/octet-stream'
    # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    att1["Content-Disposition"] = 'attachment; filename="通知书.pdf"'
    message.attach(att1)
    '''

    with open('%s录取通知书.pdf' % mingzi, 'rb') as f:
        print('读取成功')
        mime = MIMEBase(_maintype='pdf', _subtype='pdf', filename='tongzhishu.pdf')
        
        mime.add_header('Content-Disposition', 'attachment', filename='tongzhishu.pdf')# 加上必要的头信息:
        mime.add_header('Content-ID', '<0>')
        mime.add_header('X-Attachment-Id', '0')
        
        mime.set_payload(f.read())# 把附件的内容读进来:
        
        encoders.encode_base64(mime)# 用Base64编码:
        
        message.attach(mime)# 添加到MIMEMultipart:
        print('附件已添加')

    
    try:
        smtpObj = smtplib.SMTP_SSL('smtp.163.com', 465)#这里采用163邮箱
        smtpObj.login(mail_user, mail_pass)
        print('登录成功')
        smtpObj.sendmail(mail_user, receiver, message.as_string())
        smtpObj.quit()
        print("邮件发送成功")
    except smtplib.SMTPException:
        print ("Error: 无法发送邮件")



if __name__ == '__main__':
    getNames('name.xls')
