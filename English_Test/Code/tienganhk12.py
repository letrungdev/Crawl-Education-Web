from bs4 import BeautifulSoup
import requests
import re
import xlwt
from xlwt import Workbook

wb = Workbook()

types = []


def config(sheet_name):
    sheet = sheet_name
    header_font = xlwt.Font()
    header_font.name = 'Arial'
    header_font.bold = True
    header_style = xlwt.XFStyle()
    header_style.font = header_font
    sheet.col(0).width = 3700
    sheet.col(1).width = 13000
    sheet.col(4).width = 8000
    sheet.col(6).width = 6000
    sheet.col(7).width = 6000
    sheet.col(8).width = 10000
    sheet.write(0, 0, 'Kind', header_style)
    sheet.write(0, 1, 'Text (Câu hỏi chung)', header_style)
    sheet.write(0, 2, 'Text_vn (dịch câu hỏi chung)', header_style)
    sheet.write(0, 3, 'Image (ảnh chung)', header_style)

    col = 3
    for n in range(10):
        question_header = "Question {}".format(str(n + 1))
        sheet.write(0, col + 1, question_header, header_style)
        sheet.write(0, col + 2, 'image', header_style)
        sheet.write(0, col + 3, 'answers', header_style)
        sheet.write(0, col + 4, 'correct-answer', header_style)
        sheet.write(0, col + 5, 'explain vn', header_style)
        sheet.write(0, col + 6, 'explain en', header_style)
        col += 6
    return sheet


def crawl(link, sheet_name):
    sheet = config(sheet_name)
    this_page = requests.get(link)
    soup = BeautifulSoup(this_page.content, 'html.parser')
    items = soup.find_all(class_="quiz-answer-item")
    wb.save('tienganhk12.xlsx')


def insert():
    # lay danh sach link
    # links = []
    # page = requests.get(link_test, headers=headers)
    # soup = BeautifulSoup(page.content, 'html.parser')
    # tab_exam = soup.find(class_="tab-exam")
    # for link in tab_exam.find_all('a'):
    #     links.append(link.get('href') + '/ket-qua')
    # sheets = []
    # sheet_format = "Đề {}"
    #
    # for z, link in enumerate(links):
    # #     sheets.append(sheet_format.format(str(z + 1)))
    # for n in range(len(links)):
    #     print(links[n])
    #     sheet = wb.add_sheet(sheets[n], cell_overwrite_ok=True)
    #     crawl_test(links[n], headers, sheet)
    linkz = "http://tienganhk12.com/ex/1/tieng-anh-tot-nghiep-thpt/q/9103/de-tham-khao-ki-thi-tot-nghiep-trung-hoc-pho-thong-nam-2022---mon-tieng-anh/atid/4499680"
    sheet = wb.add_sheet("test", cell_overwrite_ok = True)
    crawl(linkz, sheet)

