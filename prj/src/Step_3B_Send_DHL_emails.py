import os.path
import smtplib
import time
import csv
import xlrd
import xlsxwriter
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


# parse previously processed
wb = xlrd.open_workbook("dhlDatabase.xlsx")
ws = wb.sheet_by_index(0)
processedRecords = []
oldData = []
for row_idx in range(1, ws.nrows):
    processedRecords.append(
        ws.cell_value(row_idx, 0) + " " + str(ws.cell_value(row_idx, 1))
    )
    oldData.append(
        [
            ws.cell_value(row_idx, 0),
            ws.cell_value(row_idx, 1),
            ws.cell_value(row_idx, 2),
        ]
    )

# parse report from DHL
startFromDate = datetime.date(2019, 10, 30)  # ignore rows earlier than this
pardir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
fileSrc = pardir + "\\dailyReportDHL.xlsx"
wb = xlrd.open_workbook(fileSrc)
ws = wb.sheet_by_index(0)

rawData = []
for row_idx in range(1, ws.nrows):
    if (
        datetime.datetime.strptime(ws.cell_value(row_idx, 3), "%d/%m/%Y").date()
        > startFromDate
    ) and (
        (ws.cell_value(row_idx, 0) + " " + ws.cell_value(row_idx, 2))
        not in processedRecords
    ):
        if ws.cell_value(row_idx, 1) == "":
            print("No email found for", ws.cell_value(row_idx, 0))
        else:
            rawData.append(ws.row_values(row_idx))

smtp_user = "***REMOVED***"  # SMTP username used for authentication
smtp_pass = "***REMOVED***"  # SMTP password used for authentication
server = smtplib.SMTP("smtp.sanjaytolani.com", 587)  # e.g. ('in-v3.mailjet.com', 587)
server.starttls()
server.login(smtp_user, smtp_pass)
fromaddr = "***REMOVED***"  # from email address

report = "notifications sent to: \n"

# send notification to customer
for row in rawData:
    toaddr = row[1]  # destination email address
    name = row[0]
    report += name + ", " + toaddr + ", " + row[2] + "\n"
    msg = MIMEMultipart()
    msg["From"] = fromaddr
    msg["To"] = toaddr
    msg["Subject"] = "Your Delivery Details"  # subject
    body = (
        "Hi "
        + name
        + ",\n\n"
        + "Thank you for your purchase and support. We want to inform you that your parcel has been shipped out.\n"
        + "You may follow up on your order through this tracking number ("
        + row[2]
        + "). Enter your tracking number into this link: http://parcelsapp.com/en.\n\n"
        + "Due to possible complications at the customs, please allow a few more days for the parcel to arrive.\n"
        + "Thank you for your patience and understanding.\n\n"
        + "Kind Regards,\nDr. Sanjay's Team"
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
msg["Subject"] = "DHL parcel notifications sent."  # subject
msg.attach(MIMEText(report, "plain"))
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
print("summary sent to", toaddr)
server.quit()

# write processed records to database
DB_WB = xlsxwriter.Workbook("dhlDatabase.xlsx")
DB_WS = DB_WB.add_worksheet()
DB_WS.write(0, 0, "name")
DB_WS.write(0, 1, "tracking")
DB_WS.write(0, 2, "notification date")
currentRow = 1
for row in oldData:  # write old data
    DB_WS.write(currentRow, 0, row[0])
    DB_WS.write(currentRow, 1, row[1])
    DB_WS.write(currentRow, 2, row[2])
    currentRow += 1
for row in rawData:  # write new data
    DB_WS.write(currentRow, 0, row[0])
    DB_WS.write(currentRow, 1, row[2])
    DB_WS.write(currentRow, 2, datetime.datetime.today().strftime("%d/%m/%Y"))
    currentRow += 1
DB_WB.close()
