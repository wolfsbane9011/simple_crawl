from datetime import datetime
import requests as Requests
from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Font
import time, threading
import config


table_data = dict()

def make_request(url):
    return Requests.get(url, headers=config.REQUEST_HEADERS)

def filter_data(url):
    page = make_request(url)
    tree = etree.HTML(page.text)
    table_data['columns'] = tree.xpath("//table[@id='octable']/thead/tr[2]/th/@title")
    row_data = tree.xpath("//table[@id='octable']/tr")

    for index, row in enumerate(row_data):
        row_values = list()
        for i, col in enumerate(row.xpath("//table[@id='octable']/tr[" + str(index + 1) + "]/td")):
            if len(row_data[index].xpath("//table[@id='octable']/tr" + "[" + str(index + 1) + "]" + "/td[" + str(i + 1) + "]//text()")) == 0 and len(row_data[index].xpath("//table[@id='octable']/tr" + "[" + str(index + 1) + "]" + "/td[" + str(i + 1) + "][not(text())]")) > 0:
                row_values.append("")
            else:
                path = "//table[@id='octable']/tr" + "[" + str(index + 1) + "]" + "/td[" + str(i + 1) + "]//text()"
                row_values.append("".join(row_data[index].xpath(path)).strip())
        table_data['row_' + str(index)] = row_values

def create_worksheet(filename):
    workbook = Workbook()
    sheet = workbook.active

    keys = table_data.keys()

    for index, key in enumerate(keys):
        sheet.append(table_data[key])
        if index == 0 or index == len(keys) - 1:
            style = Font(color='000000', bold=True)
            for cell in sheet[index + 1 : index + 1]:
                cell.font = style

    workbook.save(filename="output/" + filename)
    workbook.close()

def start_crawling():
    print (time.ctime())
    filename = config.FILE_PREFIX + "_" + datetime.now().strftime('%d_%m_%Y_%H_%M_%S') + ".xlsx"
    filter_data(config.URL)
    create_worksheet(filename)

ticker = threading.Event()
while not ticker.wait(config.REPORT_INTERVAL):
    start_crawling()
