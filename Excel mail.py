import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import xlrd
import xlwt

from datetime import datetime
from datetime import date

from xlwt import Workbook

wb=Workbook()

sheet1=wb.add_sheet('Sheet1')




def get_attachment_path():
    cur_dir=os.getcwd()
    #print(cur_dir)
    file_location=os.path.join('')      #enter the path of the file 
    xl_sample_workbook=xlrd.open_workbook(file_location)
    xl_sample_sheet=xl_sample_workbook.sheet_by_index(0)   #specify the sheet number of the excel sheet
    

    #print(file_location)
    for row_idx in range(1,xl_sample_sheet.nrows):
        value_cell=xl_sample_sheet.cell(row_idx,41).value
        if(value_cell=='Y' or value_cell=='y'):         #checking for a Y which corresponds to yes
            cell_obj_attachment = xl_sample_sheet.cell(row_idx,36).value
            #print("obj_addr",cell_obj_attachment)
            invoice_num=xl_sample_sheet.cell(row_idx,14).value      #different fields extracted from excel
            #return invoice_num
            invoice_date=xl_sample_sheet.cell(row_idx,15).value
            
            dt=datetime.fromordinal(datetime(1900,1,1).toordinal()+ int(invoice_date)-2)
           
            new_date=str(dt)                            #converting date from excel into string
            length=len(new_date)
            
            if(length>10):                                          
                new_date=new_date[0:10]

            
            #print (length)
            #new_date=new_date[::-1]

            
            
            #return invoice_date
            gst_amt=xl_sample_sheet.cell(row_idx,24).value
            #return gst_amt
            cgst=xl_sample_sheet.cell(row_idx,26).value
            #return cgst
            sgst=xl_sample_sheet.cell(row_idx,28).value
            #return sgst
            name_vendor=xl_sample_sheet.cell(row_idx,13).value    #Name of vendor

            gst_vendor=xl_sample_sheet.cell(row_idx,11).value      #GST of vendor

            gst_in=xl_sample_sheet.cell(row_idx,1).value

            
            cell_obj_to_addr = xl_sample_sheet.cell(row_idx,37).value
            ccc1=xl_sample_sheet.cell(row_idx,38).value                     #cc1
            ccc2=xl_sample_sheet.cell(row_idx,39).value                     #cc2
            ccc3=xl_sample_sheet.cell(row_idx,40).value                     #cc3
            print("Mail Sent to:",cell_obj_to_addr)
            print("CC Mail sent to :",'\n',ccc1,'\n',ccc2,'\n',ccc3)
            email_sent=send_mail('From_address',cell_obj_to_addr,cell_obj_attachment,ccc1,ccc2,ccc3,invoice_num,new_date,gst_amt,cgst,sgst,name_vendor,gst_vendor,gst_in)
            print("Email Sent",email_sent)
            #print('Invoice Num',invoice_num)
            #print('Invoice Date',new_date)
            #print('GST Amount',gst_amt)
            #print('CGST',cgst)
            #print('SGST',sgst)

            today=date.today()
            #dt1=datetime.fromordinal(datetime(1900,1,1).toordinal()+ (today) - 2 )
            
            sheet1.write(row_idx,0,cell_obj_to_addr)
            sheet1.write(row_idx,1,today)
            wb.save('status.xls')                               #saving the status onto another excel sheet




            





def send_mail(fromaddr='',toaddr='',attachment=None,cc1=None,cc2=None,cc3=None,i_num=None,i_date=None,amt=None,c=None,s=None,name=None,gst=None,gst_gail=None):         #enter the from and to emails
    #n1=invoice_num.get()
    
    
    
    mail_sent = False
    try:
        msg=MIMEMultipart()
        msg['From'] = fromaddr
                                                
        msg['To']= toaddr

        msg['cc']=cc1+','+cc2+','+cc3      #adding cc in mail
    
        msg['Subject'] = ''  #subject of the mail goes here

        body= ""    #body of the mail goes here






        msg.attach(MIMEText(body,'plain'))

        filename = attachment.split('\\')[-1]               #attaching a file from excel corresponding to a specific row
        print("FILENAME:" ,filename)

        attachment = open(attachment, "rb")

        p=MIMEBase('application','octet-stream')

        p.set_payload((attachment).read())

        encoders.encode_base64(p)

        p.add_header('Content-Disposition',"attachment;filename=%s" %filename)

        msg.attach(p)

        s=smtplib.SMTP('smtp.gmail.com',587)                #gmail server is used

        s.starttls()

        s.login('email','password')                     #email and password goes here

        text=msg.as_string()

        s.send_message(msg)
        mail_sent=True

        s.quit()
        return mail_sent
    except Exception as e:
        print('Error while sending mail',str(e))
        return mail_sent






if __name__ == "__main__":
    get_attachment_path()
            
        
