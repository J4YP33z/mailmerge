import json
import os.path
import xlrd
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

pardir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
filesToProcess = "zz_sg_labels.xlsx"
wb = xlrd.open_workbook(pardir + "\\" + filesToProcess)
ws = wb.sheet_by_index(0)

smtp_user = "***REMOVED***"  # SMTP username used for authentication
smtp_pass = "***REMOVED***"  # SMTP password used for authentication
server = smtplib.SMTP("smtp.sanjaytolani.com", 587)  # e.g. ('in-v3.mailjet.com', 587)
server.starttls()
server.login(smtp_user, smtp_pass)
fromaddr = "***REMOVED***"  # from email address

report = "notifications sent to: \n"

# send notification to customer
for row_idx in range(1, ws.nrows):
    toaddr = ws.cell_value(row_idx, 1)  # destination email address
    name = ws.cell_value(row_idx, 0)
    report += name + ", " + toaddr + "\n"
    msg = MIMEMultipart()
    msg["From"] = fromaddr
    msg["To"] = toaddr
    msg["Subject"] = "Your package has been shipped."  # subject
    body = (
        "Hi "
        + name
        + ", \n\nThank you for your purchase and support. We want to inform you that your parcel has been shipped out and Singpost will deliver it in under 3 working days.\n\nKind Regards,\nDr. Sanjay's Team"
    )  # body
    msg.attach(MIMEText(body, "plain"))
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    print("email sent to", name, "at", toaddr)
    time.sleep(0.2)

# send report to self
toaddr = "***REMOVED***"
msg = MIMEMultipart()
msg["From"] = fromaddr
msg["To"] = "***REMOVED***"
msg["Subject"] = "SG parcel notifications sent."  # subject
msg.attach(MIMEText(report, "plain"))
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
print("summary sent to", toaddr)

server.quit()

