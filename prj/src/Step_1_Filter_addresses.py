import configparser
import csv
import json
import os
import re
import sys

# import xlrd
import xlsxwriter

# countries dict
countries = {
    "AFGHANISTAN": "AF",
    "ÅLAND ISLANDS": "AX",
    "ALBANIA": "AL",
    "ALGERIA": "DZ",
    "AMERICAN SAMOA": "AS",
    "ANDORRA": "AD",
    "ANGOLA": "AO",
    "ANGUILLA": "AI",
    "ANTARCTICA": "AQ",
    "ANTIGUA AND BARBUDA": "AG",
    "ARGENTINA": "AR",
    "ARMENIA": "AM",
    "ARUBA": "AW",
    "AUSTRALIA": "AU",
    "AUSTRIA": "AT",
    "AZERBAIJAN": "AZ",
    "BAHAMAS": "BS",
    "BAHRAIN": "BH",
    "BANGLADESH": "BD",
    "BARBADOS": "BB",
    "BELARUS": "BY",
    "BELGIUM": "BE",
    "BELIZE": "BZ",
    "BENIN": "BJ",
    "BERMUDA": "BM",
    "BHUTAN": "BT",
    "BOLIVIA: PLURINATIONAL STATE OF": "BO",
    "BONAIRE: SINT EUSTATIUS AND SABA": "BQ",
    "BOSNIA AND HERZEGOVINA": "BA",
    "BOTSWANA": "BW",
    "BOUVET ISLAND": "BV",
    "BRAZIL": "BR",
    "BRITISH INDIAN OCEAN TERRITORY": "IO",
    "BRUNEI DARUSSALAM": "BN",
    "BULGARIA": "BG",
    "BURKINA FASO": "BF",
    "BURUNDI": "BI",
    "CAMBODIA": "KH",
    "CAMEROON": "CM",
    "CANADA": "CA",
    "CAPE VERDE": "CV",
    "CAYMAN ISLANDS": "KY",
    "CENTRAL AFRICAN REPUBLIC": "CF",
    "CHAD": "TD",
    "CHILE": "CL",
    "CHINA": "CN",
    "CHRISTMAS ISLAND": "CX",
    "COCOS (KEELING) ISLANDS": "CC",
    "COLOMBIA": "CO",
    "COMOROS": "KM",
    "CONGO": "CG",
    "CONGO: THE DEMOCRATIC REPUBLIC OF THE": "CD",
    "COOK ISLANDS": "CK",
    "COSTA RICA": "CR",
    "CÔTE D'IVOIRE": "CI",
    "CROATIA": "HR",
    "CUBA": "CU",
    "CURAÇAO": "CW",
    "CYPRUS": "CY",
    "CZECH REPUBLIC": "CZ",
    "DENMARK": "DK",
    "DJIBOUTI": "DJ",
    "DOMINICA": "DM",
    "DOMINICAN REPUBLIC": "DO",
    "ECUADOR": "EC",
    "EGYPT": "EG",
    "EL SALVADOR": "SV",
    "EQUATORIAL GUINEA": "GQ",
    "ERITREA": "ER",
    "ESTONIA": "EE",
    "ETHIOPIA": "ET",
    "FALKLAND ISLANDS (MALVINAS)": "FK",
    "FAROE ISLANDS": "FO",
    "FIJI": "FJ",
    "FINLAND": "FI",
    "FRANCE": "FR",
    "FRENCH GUIANA": "GF",
    "FRENCH POLYNESIA": "PF",
    "FRENCH SOUTHERN TERRITORIES": "TF",
    "GABON": "GA",
    "GAMBIA": "GM",
    "GEORGIA": "GE",
    "GERMANY": "DE",
    "GHANA": "GH",
    "GIBRALTAR": "GI",
    "GREECE": "GR",
    "GREENLAND": "GL",
    "GRENADA": "GD",
    "GUADELOUPE": "GP",
    "GUAM": "GU",
    "GUATEMALA": "GT",
    "GUERNSEY": "GG",
    "GUINEA": "GN",
    "GUINEA-BISSAU": "GW",
    "GUYANA": "GY",
    "HAITI": "HT",
    "HEARD ISLAND AND MCDONALD ISLANDS": "HM",
    "HOLY SEE (VATICAN CITY STATE)": "VA",
    "HONDURAS": "HN",
    "HONG KONG": "HK",
    "HUNGARY": "HU",
    "ICELAND": "IS",
    "INDIA": "IN",
    "INDONESIA": "ID",
    "IRAN: ISLAMIC REPUBLIC OF": "IR",
    "IRAQ": "IQ",
    "IRELAND": "IE",
    "ISLE OF MAN": "IM",
    "ISRAEL": "IL",
    "ITALY": "IT",
    "JAMAICA": "JM",
    "JAPAN": "JP",
    "JERSEY": "JE",
    "JORDAN": "JO",
    "KAZAKHSTAN": "KZ",
    "KENYA": "KE",
    "KIRIBATI": "KI",
    "KOREA: DEMOCRATIC PEOPLE'S REPUBLIC OF": "KP",
    "KOREA, REPUBLIC OF": "KR",
    "KUWAIT": "KW",
    "KYRGYZSTAN": "KG",
    "LAO PEOPLE'S DEMOCRATIC REPUBLIC": "LA",
    "LATVIA": "LV",
    "LEBANON": "LB",
    "LESOTHO": "LS",
    "LIBERIA": "LR",
    "LIBYAN ARAB JAMAHIRIYA": "LY",
    "LIECHTENSTEIN": "LI",
    "LITHUANIA": "LT",
    "LUXEMBOURG": "LU",
    "MACAO": "MO",
    "MACEDONIA: THE FORMER YUGOSLAV REPUBLIC OF": "MK",
    "MADAGASCAR": "MG",
    "MALAWI": "MW",
    "MALAYSIA": "MY",
    "MALDIVES": "MV",
    "MALI": "ML",
    "MALTA": "MT",
    "MARSHALL ISLANDS": "MH",
    "MARTINIQUE": "MQ",
    "MAURITANIA": "MR",
    "MAURITIUS": "MU",
    "MAYOTTE": "YT",
    "MEXICO": "MX",
    "MICRONESIA: FEDERATED STATES OF": "FM",
    "MOLDOVA: REPUBLIC OF": "MD",
    "MONACO": "MC",
    "MONGOLIA": "MN",
    "MONTENEGRO": "ME",
    "MONTSERRAT": "MS",
    "MOROCCO": "MA",
    "MOZAMBIQUE": "MZ",
    "MYANMAR": "MM",
    "NAMIBIA": "NA",
    "NAURU": "NR",
    "NEPAL": "NP",
    "NETHERLANDS": "NL",
    "NEW CALEDONIA": "NC",
    "NEW ZEALAND": "NZ",
    "NICARAGUA": "NI",
    "NIGER": "NE",
    "NIGERIA": "NG",
    "NIUE": "NU",
    "NORFOLK ISLAND": "NF",
    "NORTHERN MARIANA ISLANDS": "MP",
    "NORWAY": "NO",
    "OMAN": "OM",
    "PAKISTAN": "PK",
    "PALAU": "PW",
    "PALESTINIAN TERRITORY: OCCUPIED": "PS",
    "PANAMA": "PA",
    "PAPUA NEW GUINEA": "PG",
    "PARAGUAY": "PY",
    "PERU": "PE",
    "PHILIPPINES": "PH",
    "PITCAIRN": "PN",
    "POLAND": "PL",
    "PORTUGAL": "PT",
    "PUERTO RICO": "PR",
    "QATAR": "QA",
    "RÉUNION": "RE",
    "ROMANIA": "RO",
    "RUSSIAN FEDERATION": "RU",
    "RWANDA": "RW",
    "SAINT BARTHÉLEMY": "BL",
    "SAINT HELENA: ASCENSION AND TRISTAN DA CUNHA": "SH",
    "SAINT KITTS AND NEVIS": "KN",
    "SAINT LUCIA": "LC",
    "SAINT MARTIN (FRENCH PART)": "MF",
    "SAINT PIERRE AND MIQUELON": "PM",
    "SAINT VINCENT AND THE GRENADINES": "VC",
    "SAMOA": "WS",
    "SAN MARINO": "SM",
    "SAO TOME AND PRINCIPE": "ST",
    "SAUDI ARABIA": "SA",
    "SENEGAL": "SN",
    "SERBIA": "RS",
    "SEYCHELLES": "SC",
    "SIERRA LEONE": "SL",
    "SINGAPORE": "SG",
    "SINT MAARTEN (DUTCH PART)": "SX",
    "SLOVAKIA": "SK",
    "SLOVENIA": "SI",
    "SOLOMON ISLANDS": "SB",
    "SOMALIA": "SO",
    "SOUTH AFRICA": "ZA",
    "SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS": "GS",
    "SOUTH SUDAN": "SS",
    "SPAIN": "ES",
    "SRI LANKA": "LK",
    "SUDAN": "SD",
    "SURINAME": "SR",
    "SVALBARD AND JAN MAYEN": "SJ",
    "SWAZILAND": "SZ",
    "SWEDEN": "SE",
    "SWITZERLAND": "CH",
    "SYRIAN ARAB REPUBLIC": "SY",
    "TAIWAN, PROVINCE OF CHINA": "TW",
    "TAJIKISTAN": "TJ",
    "TANZANIA: UNITED REPUBLIC OF": "TZ",
    "THAILAND": "TH",
    "TIMOR-LESTE": "TL",
    "TOGO": "TG",
    "TOKELAU": "TK",
    "TONGA": "TO",
    "TRINIDAD AND TOBAGO": "TT",
    "TUNISIA": "TN",
    "TURKEY": "TR",
    "TURKMENISTAN": "TM",
    "TURKS AND CAICOS ISLANDS": "TC",
    "TUVALU": "TV",
    "UGANDA": "UG",
    "UKRAINE": "UA",
    "UNITED ARAB EMIRATES": "AE",
    "UNITED KINGDOM": "GB",
    "UNITED STATES": "US",
    "UNITED STATES MINOR OUTLYING ISLANDS": "UM",
    "URUGUAY": "UY",
    "UZBEKISTAN": "UZ",
    "VANUATU": "VU",
    "VENEZUELA: BOLIVARIAN REPUBLIC OF": "VE",
    "VIETNAM": "VN",
    "VIET NAM": "VN",
    "VIRGIN ISLANDS: BRITISH": "VG",
    "VIRGIN ISLANDS: U.S.": "VI",
    "WALLIS AND FUTUNA": "WF",
    "WESTERN SAHARA": "EH",
    "YEMEN": "YE",
    "ZAMBIA": "ZM",
    "ZIMBABWE": "ZW",
}


# config variables
pardir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))
config = configparser.ConfigParser()
config.read(pardir + "\Step_1_config.txt")
fileSrc = pardir + "\\" + (config["CONFIG"]["RawData"]).strip()
lastNameProcessed = (config["CONFIG"]["lastNameProcessed"]).strip()
lastCountryProcessed = (config["CONFIG"]["lastCountryProcessed"]).strip()
excludedProducts = json.loads(config.get("CONFIG", "excludedProducts"))
bookCode = (config["CONFIG"]["bookCode"]).strip()
startingNumber = int((config["CONFIG"]["startingNumber"]).strip())


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
        lastNameProcessed.upper() == col[0].upper()
        and lastCountryProcessed.upper() == col[15].upper()
    ):
        lastRecordedIndex = i
del rawData[: lastRecordedIndex + 1]

# error handling
if len(rawData) == 0:
    print("NO NEW SALES")
    sys.exit()

# output files
SG_WB = xlsxwriter.Workbook(pardir + "\\zz_sg.xlsx")
MY_WB = xlsxwriter.Workbook(pardir + "\\zz_my.xlsx")
PH_WB = xlsxwriter.Workbook(pardir + "\\zz_ph.xlsx")
OTHERS_WB = xlsxwriter.Workbook(pardir + "\\zz_others.xlsx")
DHL_WB = xlsxwriter.Workbook(pardir + "\\zz_DHL.xlsx")

SG_WS = SG_WB.add_worksheet()
MY_WS = MY_WB.add_worksheet()
PH_WS = PH_WB.add_worksheet()
OTHERS_WS = OTHERS_WB.add_worksheet()
DHL_WS = DHL_WB.add_worksheet()


# file name of csv and last processed name goes here
outFile = open(pardir + "\\zz_lastNameProcessed.txt", "w")
tmpStr = (
    "rawData = "
    + (config["CONFIG"]["RawData"]).strip()
    + "\nlastNameProcessed = "
    + rawData[-1][0]
    + "\nlastCountryProcessed = "
    + rawData[-1][15]
)
outFile.write(tmpStr)


# remove unnecessary products
for product in excludedProducts:
    rawData = [col for col in rawData if product not in col[17]]


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
        if addresses.count(address) != 1:  # each address should only exist once in list
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
LastShipmentOrderID = ""
# output by country
for col in rawData[0:]:
    shipmentOrderID = ""
    if col[15] == "Singapore":
        indexOut = SGindex
        SGindex += 1
        wsOut = SG_WS
    elif col[15] in ["Malaysia", "Hong Kong", "Canada", "Iran"]:
        indexOut = MYindex
        MYindex += 1
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

        countryCode = countries.get(col[8].upper())
        if countryCode is None:
            print("Country code not found for: ", col[8].upper())

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
        LastShipmentOrderID = shipmentOrderID

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

outFile.write("\n" + LastShipmentOrderID)

# cleanup
SG_WB.close()
MY_WB.close()
PH_WB.close()
OTHERS_WB.close()
DHL_WB.close()
outFile.close()
sys.exit()
