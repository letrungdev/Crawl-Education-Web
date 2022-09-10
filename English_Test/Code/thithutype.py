from bs4 import BeautifulSoup
from xlwt import Workbook
import requests
import re
import xlwt

wb = Workbook()


headersz = {
  'authority': 'thithu.edu.vn',
  'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
  'accept-language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
  'cache-control': 'max-age=0',
  'cookie': '_ga=GA1.3.1862793897.1652769137; _gid=GA1.3.2106382662.1652769137; __gads=ID=ca4e940464b9b4b5-220f3ae33ed30021:T=1652769137:RT=1652769137:S=ALNI_MYxuJE8CpwximKcU3vK9CnUM1O-kg; __gpi=UID=0000056d1ba781a7:T=1652769137:RT=1652769137:S=ALNI_MbbPG4aHA5A4itkAyxVZgXr-Sb_Rw; remember_web_59ba36addc2b2f9401580f014c7f58ea4e30989d=eyJpdiI6Ik9DMHlNNVhrQVlYNlBrcGxKdjdkRFE9PSIsInZhbHVlIjoiYXU1THNaa1dlL0Vac1N0U3drd0tqeW1LcnNQa29vSXlZQVd2WkVPclBDalFLU3ZRcE1JdjFoVzdSeUxyWVRVY1I3U2lJZHpwcS9iMTlvYTRhYlhZZHJGWkVYTzlEQWVUcTlGSktObnhaZEE9IiwibWFjIjoiMDMyY2QyN2E0YWE0YjBjNGYyNmViNmVkYmFkYzgzMDIxNDI5NmU3ZDIxMDdhMGZhMTBmNjNhYWJlNWYxYWE5NiJ9; XSRF-TOKEN=eyJpdiI6IkozYmx1b3IycWg0aVF6MHppdWVLaXc9PSIsInZhbHVlIjoiSTZ4SnFzaWZOdHpLcDUvNFJNWUhhNnVlQjBXY2dTS1hjd2ZLditZakMrNUxJSUhCdkJNYmNDRkY4RHlKeDB6MyIsIm1hYyI6ImZjZGI1MGFhNmM4NTc5OGVkODk0ZGViZWNhNDVjNTlhYWE1OWUxMjQwZDhiMDlmYzViZTBlMGUyMzQ0ZmE4YWYifQ%3D%3D; thi_thu_online_session=eyJpdiI6ImR1NFhxRUtyMXVCcFFrakJncU12MUE9PSIsInZhbHVlIjoiL2Q2S01taTdNTXVBZ2pNNEFNSUw4aWJiRm5KekQ5a3kyZTg0TFFwV1VuUC8xUG85TlJDV2xkYitJVDhiR3pnMSIsIm1hYyI6IjU0MDkwY2ZkYTJhYzIxNWFhYmMzOTk0OTA5OGE4ZjczZmQ4MWI3NzNlZTU0ZDI1YjFjMzg5MDJhYjliYTVkOWMifQ%3D%3D; _gat_gtag_UA_148406271_1=1; XSRF-TOKEN=eyJpdiI6IjRKOFNFYzAyazJSUG5FbEVaWUFnYlE9PSIsInZhbHVlIjoidVFmbWkvbFZhajdXQitpYnFLZkgrK29PNW0rWVFtd1dycjJhK0Z3bkpTbnVPYTVPUGE0S3pOUEd2Y0N1WG9mZSIsIm1hYyI6IjdiZmU5NzFiZTViYzdiZDI1M2NjZDBkNTNkZDFhMDEwODQ3NjZjODYyNWVlNjI3NjcxODcxZGNlYzBjMDMyMWEifQ%3D%3D; thi_thu_online_session=eyJpdiI6Ikt5Q0VYQXZOUDFvWEVrZzRDTE9hUHc9PSIsInZhbHVlIjoiVzlmQUIyQTZoRS95Y24ra3hKR3hEeE5TQm0wZGRUbjhpYTFzdkRkQW9FK21rOHk3b2l0dFRUT29oaHFaa3dtWCIsIm1hYyI6IjViYTEyOWMxYzRjM2NlNDc5ZjU1MjRhODE1ZWEwNzMxNWM4ZjJlNzQwM2NiNGJlYzcwMjFkZjljYjdhYzhlYjEifQ%3D%3D',
  'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="101", "Google Chrome";v="101"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"Windows"',
  'sec-fetch-dest': 'document',
  'sec-fetch-mode': 'navigate',
  'sec-fetch-site': 'cross-site',
  'sec-fetch-user': '?1',
  'upgrade-insecure-requests': '1',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36'
}


def config(sheet_name):
    sheet = sheet_name
    style = xlwt.XFStyle()
    style.alignment.wrap = 1
    style.alignment.VERT_TOP = 0x00  # wrong
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
    sheet.col(9).width = 8000
    sheet.write(0, 0, 'Kind', header_style)
    sheet.write(0, 1, 'Text (Câu hỏi chung)', header_style)
    sheet.write(0, 2, 'Text_vn (dịch câu hỏi chung)', header_style)
    sheet.write(0, 3, 'Image (ảnh chung)', header_style)

    col = 3
    for num in range(10):
        question_header = "Question {}".format(str(num + 1))
        sheet.col(col+1).width = 8000
        sheet.col(col+3).width = 6000
        sheet.col(col+4).width = 6000
        sheet.col(col+5).width = 10000
        sheet.col(col+6).width = 8000
        sheet.write(0, col + 1, question_header, header_style)
        sheet.write(0, col + 2, 'image', header_style)
        sheet.write(0, col + 3, 'answers', header_style)
        sheet.write(0, col + 4, 'correct-answer', header_style)
        sheet.write(0, col + 5, 'explain vn', header_style)
        sheet.write(0, col + 6, 'explain en', header_style)
        col += 6
    return sheet, style


id_form = "dvCau{}"
ids = []
for n in range(10):
    ids.append(id_form.format(str(n+1)))


links = [
         "https://thithu.edu.vn/ket_qua/243",
         "https://thithu.edu.vn/ket_qua/361",
         "https://thithu.edu.vn/ket_qua/362"
         ]


def crawl():
    count = 0
    sheet_name = wb.add_sheet("Antonym", cell_overwrite_ok=True)
    sheet, style = config(sheet_name)
    pattern = '[A-D]'
    for link in links:
        print(link)
        this_page = requests.get(link, headers=headersz)
        soup = BeautifulSoup(this_page.content, 'html.parser')
        for id in ids:
            count += 1
            print(count)
            item = soup.find(id=id)
            answers = ""
            reason = ""
            kind = "Từ trái nghĩa"
            answer_list = item.find_all('span')
            if len(item.find_all('p')) >= 1:
                question = str(item.find_all('p')[1]).replace('<u>', '^').replace('</u>', '$').replace('<strong>', '`').replace(
                    '</strong>', '~')
                ques_soup = BeautifulSoup(question, 'html.parser').get_text().replace('^', '<u>').replace('$', '</u>').replace(
                    '`', '<strong>').replace('~', '</strong>')
                question = ques_soup
                for ans in answer_list[0:(len(answer_list) - 1)]:
                    answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
                if item.find(class_="p-1 ml-2") is not None:
                    reasons = item.find(class_="p-1 ml-2").find_all('p')
                    for rea in reasons:
                        reason += rea.get_text() + '\n'
                    correct_answer = reasons[0].get_text()
                    if "Đáp án" in correct_answer or "Question" in correct_answer:
                        x = re.findall(pattern, correct_answer)
                        if len(x) >= 1:
                            correct_answer = x[len(x) - 1]
                        else:
                            correct_answer = ""
                    else:
                        correct_answer = reasons[len(reasons) - 1].get_text()
                        x = re.findall(pattern, correct_answer)
                        if len(x) >= 1:
                            correct_answer = x[len(x) - 1]
                        else:
                            correct_answer = ""
                    if correct_answer == 'A':
                        correct_answer = str(1)
                    elif correct_answer == 'B':
                        correct_answer = str(2)
                    elif correct_answer == 'C':
                        correct_answer = str(3)
                    elif correct_answer == 'D':
                        correct_answer = str(4)
                    sheet.write(count, 0, kind, style)
                    sheet.write(count, 4, question, style)
                    sheet.write(count, 6, answers, style)
                    sheet.write(count, 7, correct_answer, style)
                    sheet.write(count, 8, reason, style)
                else:
                    count -= 1
            else:
                count -= 1
    wb.save('antonym.xlsx')


crawl()




