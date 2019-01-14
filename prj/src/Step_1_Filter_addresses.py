import configparser
import csv
import json
import os
import re
import sys

import pycountry
import requests
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# import xlrd


options = webdriver.ChromeOptions()
options.add_argument("headless")
options.add_argument("log-level=3")
browser = webdriver.Chrome(os.getcwd() + "\chromedriver.exe", chrome_options=options)
print("Logging in to CF...")
browser.get("https://app.clickfunnels.com/users/sign_out")  # sign out first
browser.get("https://app.clickfunnels.com/users/sign_in")  # sign in
userNameField = browser.find_element_by_id("user_email")
userNameField.send_keys("***REMOVED***")
pwField = browser.find_element_by_id("user_password")
pwField.send_keys("***REMOVED***")
pwField.send_keys(Keys.ENTER)


# config variables
pardir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
config = configparser.ConfigParser()
config.read(pardir + "\Step_1_config.txt")

# get Country codes from pycountry module and config
countryCodes = {}
for country in pycountry.countries:
    countryCodes[country.name.upper()] = country.alpha_2
for key in config["Additional countries"]:
    countryCodes[str(key).upper()] = (
        config["Additional countries"][str(key).upper()]
    ).upper()


SG_LABELS_WB = xlsxwriter.Workbook(pardir + "\\zz_sg_labels.xlsx")
MY_LABELS_WB = xlsxwriter.Workbook(pardir + "\\zz_my_labels.xlsx")
SG_LABELS_WS = SG_LABELS_WB.add_worksheet()
MY_LABELS_WS = MY_LABELS_WB.add_worksheet()
SG_LABELS_index = 0
MY_LABELS_index = 0


funnels = config.sections()
funnels.remove("Additional countries")
for funnel in funnels:
    URL = config[funnel]["URL"]
    excludedProducts = json.loads(config[funnel]["excludedProducts"])
    bookCode = config[funnel]["bookCode"]
    startingNumber = int(config[funnel]["startingNumber"])
    lastNameProcessed = config[funnel]["lastNameProcessed"]
    lastCountryProcessed = config[funnel]["lastCountryProcessed"]

    print("Generating sales list for " + funnel + "...")
    browser.get(URL)
    elem = browser.find_element_by_xpath("//a[@class='btn btn-default export-link']")
    elem.click()  # generate sales list

    print("Downloading sales list for " + funnel + "...")
    wait = WebDriverWait(browser, 120)  # wait for CF to provide download link
    elem = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[@class='btn btn-primary' and text() = 'Download']")
        )
    )
    r = requests.get(elem.get_attribute("href"), allow_redirects=True)
    fileName = "\\" + funnel + "_sales.csv"
    open(pardir + fileName, "wb").write(r.content)
    print("Download complete!")

    fileSrc = pardir + fileName

    # read csv file from ClickFunnels
    rawData = []
    with open(fileSrc, newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        for row in reader:
            rawData.append(row)

    # remove records that have already been processed
    lastRecordedIndex = 0
    for i, col in enumerate(rawData[0:]):
        if (
            lastNameProcessed.upper() == col[0].strip().upper()
            and lastCountryProcessed.upper() == col[15].upper()
        ):
            lastRecordedIndex = i
    del rawData[: lastRecordedIndex + 1]

    # error handling
    if rawData == []:
        print("NO NEW SALES FOR " + funnel + "!")
        continue

    # save processed name into config
    config[funnel]["lastNameProcessed"] = rawData[-1][0].strip()
    config[funnel]["lastCountryProcessed"] = rawData[-1][15]

    # remove unnecessary products
    for product in excludedProducts:
        rawData = [col for col in rawData if product not in col[17]]

    # output files
    SG_WB = xlsxwriter.Workbook(pardir + "\\" + funnel + "_sg.xlsx")
    MY_WB = xlsxwriter.Workbook(pardir + "\\" + funnel + "_my.xlsx")
    PH_WB = xlsxwriter.Workbook(pardir + "\\" + funnel + "_ph.xlsx")
    OTHERS_WB = xlsxwriter.Workbook(pardir + "\\" + funnel + "_others.xlsx")
    DHL_WB = xlsxwriter.Workbook(pardir + "\\" + funnel + "_DHL.xlsx")
    SG_WS = SG_WB.add_worksheet()
    MY_WS = MY_WB.add_worksheet()
    PH_WS = PH_WB.add_worksheet()
    OTHERS_WS = OTHERS_WB.add_worksheet()
    DHL_WS = DHL_WB.add_worksheet()

    # merge orders going to same address
    while True:
        addresses = []
        for col in rawData[0:]:
            addresses.append(col[11])
        totalCount = 0
        for address in addresses:
            totalCount += addresses.count(address)
            lastIndex = ""
            firstIndex = ""
            if (
                addresses.count(address) != 1
            ):  # each address should only exist once in list
                for i, col in enumerate(rawData[0:]):
                    if col[11] == address:
                        firstIndex = i
                        break  # first index found
                for i, col in enumerate(rawData[0:]):
                    if col[11] == address:
                        lastIndex = i
                rawData[firstIndex][17] += "," + rawData[lastIndex][17]
                del rawData[lastIndex]
                break  # one address merged, restart while loop
        if totalCount == len(addresses):  # no duplicates
            break

    # write headers
    for i, header in enumerate(
        [
            "NAME",
            "EMAIL",
            "TRACKING NO.",
            "PHONE",
            "ADDRESS",
            "CITY",
            "STATE",
            "POSCODE",
            "COUNTRY",
            "ORDERS",
            "STATUS",
            "COST",
        ]
    ):
        SG_WS.write(0, i, header)
        MY_WS.write(0, i, header)
        PH_WS.write(0, i, header)
        OTHERS_WS.write(0, i, header)
        SG_LABELS_WS.write(0, i, header)
        MY_LABELS_WS.write(0, i, header)

    for i, header in enumerate(
        [
            "Pick-up Account Number",
            "Shipment Order ID",
            "Shipping Service Code",
            "Consignee Name",
            "Address Line 1",
            "Address Line 2",
            "Address Line 3",
            "City",
            "State (M)",
            "Postal Code (M)",
            "Destination Country Code",
            "Phone Number",
            "Email Address",
            "Shipment Weight (g)",
            "Currency Code",
            "Total Declared Value",
            "Incoterm",
            "Shipment Description",
            "Content Description",
            "Content Export Description",
            "Content Unit Price",
            "Content Origin Country",
            "Content Quantity",
            "Content Code",
            "Content Indicator",
            "Remarks",
        ]
    ):
        DHL_WS.write(0, i, header)

    SGindex = 1
    MYindex = 1
    PHindex = 1
    OTHERSindex = 1
    DHLindex = 1
    wsOut = ""
    indexOut = ""
    # output by country
    for col in rawData[0:]:
        shipmentOrderID = ""
        if col[15] == "Singapore":
            indexOut = SGindex
            SGindex += 1
            SG_LABELS_index += 1
            wsOut = SG_WS
        elif col[15] in ["Malaysia", "Hong Kong", "Canada", "Iran"]:
            indexOut = MYindex
            MYindex += 1
            MY_LABELS_index += 1
            wsOut = MY_WS
        elif col[15] == "Philippines":
            indexOut = PHindex
            PHindex += 1
            wsOut = PH_WS
        else:
            indexOut = OTHERSindex
            OTHERSindex += 1
            wsOut = OTHERS_WS

            # handle DHL output
            quantityTotal = 0
            tmpOrders = "," + col[17]
            commas = [m.start() for m in re.finditer(",", tmpOrders)]
            Xs = [m.start() for m in re.finditer(" X ", tmpOrders)]
            for startQuantity, endQuantity in zip(commas, Xs):
                quantityTotal += int(tmpOrders[startQuantity + 1 : endQuantity])

            quantity28000 = 0
            if " X 28000 Book" in tmpOrders:
                endIndex = [m.start() for m in re.finditer(" X 28000", tmpOrders)]
                for i in endIndex:
                    startIndex = tmpOrders[:i].rindex(",") + 1
                    quantity28000 += int(tmpOrders[startIndex:i])

            countryCode = countryCodes.get(col[8].upper())
            if countryCode is None:
                print(
                    "Country code not found for:", col[8].upper()
                )  # dbg add file name and row here

            if countryCode in ["SG", "TH", "AU", "GB"]:
                shippingServiceCode = "PLT"
                incoterm = "DDP"
            elif countryCode == "US":
                shippingServiceCode = "PLE"
                incoterm = "DDP"
            else:
                shippingServiceCode = "PPS"
                incoterm = "DDU"

            shipmentOrderID = bookCode + str(startingNumber)

            # All books(250g) except 28000 book(425g)
            weight = (quantityTotal - quantity28000) * 250 + (quantity28000 * 425)

            # RM10 per book, max RM50
            declaredValue = min(quantityTotal * 10, 50)

            for k, content in enumerate(
                [
                    "5345221",
                    shipmentOrderID,
                    shippingServiceCode,
                    col[0][:30],  # name is MAX 30 CHARS
                    col[11],
                    "",
                    "",
                    col[6],
                    col[7],
                    col[9],
                    countryCode,
                    col[10],
                    col[3],
                    weight,
                    "MYR",
                    declaredValue,
                    incoterm,
                    "educational book, perfect bound book",
                    "educational book, perfect bound book",
                    "",
                    declaredValue,
                    "MY",
                    1,
                    shipmentOrderID,
                    "",
                    "",
                ]
            ):
                DHL_WS.write(DHLindex, k, content)
            DHLindex += 1
            startingNumber += 1

        for j, content in enumerate(
            [
                col[0],
                col[3],
                shipmentOrderID,
                col[10],
                col[11],
                col[13],
                col[14],
                col[16],
                col[15],
                col[17],
                "",
                "",
            ]
        ):
            wsOut.write(indexOut, j, content)
            if wsOut == SG_WS:
                SG_LABELS_WS.write(SG_LABELS_index, j, content)
            elif wsOut == MY_WS:
                MY_LABELS_WS.write(MY_LABELS_index, j, content)

    config[funnel]["startingnumber"] = str(startingNumber)

    # cleanup
    SG_WB.close()
    MY_WB.close()
    PH_WB.close()
    OTHERS_WB.close()
    DHL_WB.close()
    print("Output for " + funnel + " complete!")


with open(pardir + "\Step_1_config.txt", "w") as configfile:
    config.write(configfile)
