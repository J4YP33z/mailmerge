import configparser
import json
import os.path
import xlrd

import labels
from reportlab.graphics import shapes
from reportlab.pdfbase.pdfmetrics import stringWidth


# Callback function to draw each label.
def writeLabelInfo(label, width, height, info):
    # (0, 0) is bottom left, (width, height) is top right

    startingFontSize = 12

    x = 5  # starting x coord
    numOfLines = 5  # lines on labels
    verticalSpacing = height / (numOfLines + 2)
    y = height - 20  # starting y coord

    # name
    s = shapes.String(x, y, info[0], fontName="Times-Roman", fontSize=startingFontSize)
    label.add(s)

    # address, shrink font to fit label
    font_size = startingFontSize
    while stringWidth(info[1], "Times-Roman", font_size) > width - 10:
        font_size *= 0.9
    y -= verticalSpacing
    s = shapes.String(x, y, info[1], fontName="Times-Roman", fontSize=font_size)
    label.add(s)

    # city, state, , shrink font to fit label
    tmpStr = info[2] + ", " + info[3]
    # font_size = startingFontSize
    while stringWidth(tmpStr, "Times-Roman", font_size) > width - 10:
        font_size *= 0.9
    y -= verticalSpacing
    s = shapes.String(x, y, tmpStr, fontName="Times-Roman", fontSize=font_size)
    label.add(s)

    # country poscode
    tmpStr = str(info[4]) + " " + str(info[5])
    y -= verticalSpacing
    s = shapes.String(x, y, tmpStr, fontName="Times-Roman", fontSize=font_size)
    label.add(s)

    # orders, shrink font to fit label
    # font_size = startingFontSize
    while stringWidth(info[6], "Times-Roman", font_size) > width - 10:
        font_size *= 0.9
    y -= verticalSpacing
    s = shapes.String(x, y, info[6], fontName="Times-Roman", fontSize=font_size)
    label.add(s)


# config variables
pardir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
config = configparser.ConfigParser()
config.read(pardir + "\Step_2_config.txt")
filesToProcess = json.loads(config.get("CONFIG", "filesToProcess"))


# Create A4 portrait (210mm x 297mm) sheets with labels.
specs = labels.Specification(
    sheet_width=210,  # A4
    sheet_height=297,  # A4
    columns=(config["CONFIG"]["columns"]),
    rows=(config["CONFIG"]["rows"]),
    label_width=(config["CONFIG"]["label_width"]),
    label_height=(config["CONFIG"]["label_height"]),
    corner_radius=2,
)


for countryFile in filesToProcess:
    # Create the sheet with or without border
    sheet = labels.Sheet(specs, writeLabelInfo, border=False)

    # get address data from spreadsheet
    wb = xlrd.open_workbook(pardir + "\\" + countryFile)
    ws = wb.sheet_by_index(0)
    allInfo = []
    rowInfo = []
    for row_idx in range(1, ws.nrows):
        rowInfo.append(ws.cell_value(row_idx, 0))  # name
        rowInfo.append(ws.cell_value(row_idx, 4))  # address
        rowInfo.append(ws.cell_value(row_idx, 5))  # city
        rowInfo.append(ws.cell_value(row_idx, 6))  # state
        rowInfo.append(ws.cell_value(row_idx, 8))  # country
        rowInfo.append(ws.cell_value(row_idx, 7))  # poscode
        rowInfo.append(ws.cell_value(row_idx, 9))  # orders
        allInfo.append(rowInfo[:])
        rowInfo.clear()

    # create actual labels here...
    sheet.add_labels(info for info in allInfo)

    # Save the labels sheet
    sheet.save(pardir + "\\" + countryFile[:-5] + "_labels.pdf")
    print(
        countryFile[:-5]
        + "_labels.pdf created with "
        + str(sheet.label_count)
        + " label(s) on "
        + str(sheet.page_count)
        + " sheet(s)."
    )

