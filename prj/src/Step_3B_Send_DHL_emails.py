import os.path
import smtplib
import time
import csv
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


# parse previously processed records
processedRecords = []
with open("dhlDatabase.csv", newline="", encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile, delimiter=",")
    next(reader, None)  # skip headers
    for row in reader:
        processedRecords.append(row[0] + " " + row[1])

print(processedRecords)

# parse report from DHL
startFromDate = datetime.date(2019, 10, 30)  # ignore rows earlier than this
pardir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
fileSrc = pardir + "\\dailyReportDHL.csv"
rawData = []
with open(fileSrc, newline="", encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile, delimiter=",")
    next(reader, None)  # skip headers
    for row in reader:
        if (datetime.datetime.strptime(row[3], "%d/%m/%Y").date() > startFromDate) and (
            (row[0] + " " + row[2]) not in processedRecords
        ):
            if row[1] == "":
                print("No email found for", row[0])
            else:
                rawData.append(row)

print(rawData)

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
with open("dhlDatabase.csv", mode="a", newline="") as dbDHL:
    writer = csv.writer(dbDHL, delimiter=",")
    for row in rawData:
        writer.writerow([row[0], row[2], datetime.date.today()])
