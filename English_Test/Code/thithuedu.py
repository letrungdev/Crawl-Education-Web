from bs4 import BeautifulSoup
from xlwt import Workbook
import requests
import re
import xlwt

wb = Workbook()

types = ["the word whose underlined",                                           # phat am               0
         "the word(s) whose underlined part differs",                           # phat am               1

         "the word that differs from the",                                      # trong am              2
         "the main stress different from",                                      # trong am              3

         "the word(s) CLOSEST",                                                 # dong nghia            4
         "CLOSET",                                                              # dong nghia            5
         "the word or phrase CLOSEST",                                          # dong nghia            6
         "the word or phrase that is CLOSEST in meaning to the underlined",     # dong nghia            7
         "the word(s) or phrase(s) CLOSEST in meaning to the underlined part",  # dong nghia            8

         "the word(s) OPPOSITE",                                                # trai nghia            9
         "the word or phrase OPPOSITE",                                         # trai nghia            10
         "the word or phrase that is OPPOSITE in meaning to the underlined",    # trai nghia            11
         "the word(s) or phrase(s) OPPOSITE in meaning to the underlined",      # trai nghia            12

         "following exchanges",                                                 # tinh huong giao tiep  13

         "correct answer to each of the following questions",                   # hoan thanh cau        14

         "correct word or phrase that best fits",                               # hoan thanh doan van   15
         "word or phrase that bests fits",                                      # hoan thanh doan van   16
         "correct word for each of the blanks in",                              # hoan thanh doan van   17
         "the best option for each of the following",                           # hoan thanh doan van   18


         "correct answer to cach of the question",                              # doc hieu              19
         "correct answer to each of the question",                              # doc hieu              20
         "indicate the correct answer to the following question",               # doc hieu              21
         "the correct word for each of the blanks",                             # doc hieu              22
         "correct answer to each of the numbered blanks",                       # doc hieu              23

         "passage and mark the letter A, B, C or D on your answer sheet to indicate the correct answer to each of the",

         "closest in meaning to each of the following questions",               # viet lai cau          25
         "closest in meaning to each of the following sentences",               # viet lai cau          26
         "the sentence that is closest in",                                     # viet lai cau          27
         "CLOSEST in meaning to each of the following sentences",               # viet lai cau          28
         "closest inmeaning to each of the following questions",                # viet lai cau          29

         "needs correction",                                                    # tim loi sai           30
         "sheet to indicate the underlined part that",                          # tim loi sai           31

         "combines each pair",                                                  # ket hop cau           32
         "best joins each of the following pairs"                               # ket hop cau           33
         ]
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
for n in range(50):
    ids.append(id_form.format(str(n+1)))


def crawl(link, headers, sheet_name):
    sheet, style = config(sheet_name)
    this_page = requests.get(link, headers=headers)
    soup = BeautifulSoup(this_page.content, 'html.parser')
    question_kind = []
    text_check = ""
    hoan_thanh_doan_van = 0
    hoan_thanh_doan_van_line = 0
    doc_hieu = 0
    doc_hieu_line = 0
    pattern = '[A-D]'
    for count, id in enumerate(ids):
        count += 1
        item = soup.find(id=id)
        title = item.p.get_text()
        question = ""
        answers = ""
        reason = ""
        text = ""

        # Phat am
        if types[0] in title or types[1] in title:
            kind = "Phát âm"
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                ans = str(ans).replace('<u>', '^').replace('</u>', '$')
                good_soup = BeautifulSoup(ans, 'html.parser')
                ans = good_soup.get_text().lstrip('\n').replace('^', '<u>').replace('$', '</u>')
                answers += str(ans) + '\n'
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
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
            question_kind.append(kind)
        # Trong am
        elif types[2] in title or types[3] in title:
            kind = "Trọng âm"
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
            question_kind.append(kind)

        # Tu dong nghia
        elif types[4] in title or types[5] in title or types[6] in title or types[7] in title or types[8] in title:
            kind = "Từ đồng nghĩa"
            answer_list = item.find_all('span')
            question = str(item.find_all('p')[1]).replace('<u>', '^').replace('</u>', '$').replace('<strong>',
                                                                                                   '`').replace(
                '</strong>', '~')
            ques_soup = BeautifulSoup(question, 'html.parser').get_text().replace('^', '<u>').replace('$',
                                                                                                      '</u>').replace(
                '`', '<strong>').replace('~', '</strong>')
            question = ques_soup
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)
        # Tu trai nghia
        elif types[9] in title or types[10] in title or types[11] in title or types[12] in title:
            kind = "Từ trái nghĩa"
            answer_list = item.find_all('span')

            question = str(item.find_all('p')[1]).replace('<u>', '^').replace('</u>', '$').replace('<strong>',
                                                                                                   '`').replace(
                '</strong>', '~')
            ques_soup = BeautifulSoup(question, 'html.parser').get_text().replace('^', '<u>').replace('$',
                                                                                                      '</u>').replace(
                '`', '<strong>').replace('~', '</strong>')
            question = ques_soup
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)

        # Tinh huong giao tiep
        elif types[13] in title:
            kind = "Tình huống giao tiếp"
            ques_p = item.find(class_="col-12 p-1").find_all('p')[1:]
            for ques in ques_p:
                question += ques.get_text() + '\n'
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)

        # Hoan thanh cau
        elif types[14] in title:
            kind = "Hoàn thành câu"
            question = item.find_all('p')[1].get_text()
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)

        # Hoan thanh doan van
        elif types[17] in title or types[15] in title or types[16] in title or types[18] in title:
            kind = "Hoàn thành đoạn văn"
            texts = item.find_all('p')[0].get_text().split('\xa0')
            for tex in texts:
                if "following passage" not in tex and tex != '' and tex != ' ':
                    text += tex + '\n'
            answer_list = item.find_all('span')
            for ans in answer_list[1:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            if text_check == text:
                question = "(" + str(hoan_thanh_doan_van + 2) + ")"
                sheet.write(hoan_thanh_doan_van_line, 10 + 6 * hoan_thanh_doan_van, question, style)
                sheet.write(hoan_thanh_doan_van_line, 12 + 6 * hoan_thanh_doan_van, answers, style)
                sheet.write(hoan_thanh_doan_van_line, 13 + 6 * hoan_thanh_doan_van, correct_answer, style)
                sheet.write(hoan_thanh_doan_van_line, 14 + 6 * hoan_thanh_doan_van, reason, style)
                hoan_thanh_doan_van += 1
                question_kind.append(kind)
            else:
                text_check = text
                question = "(" + str(1) + ")"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 1, text, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
                hoan_thanh_doan_van_line = count
                question_kind.append(kind)

        # Doc hieu
        elif types[22] in title or types[23] in title or types[19] in title or types[20] in title or types[21] in title or types[24] in title:
            kind = "Đọc hiểu"
            texts = str(item.find_all('p')[0]).replace('<u>', '^').replace('</u>', '$').replace('<strong>',
                                                                                                '`').replace(
                '</strong>', '~')
            text_soup = BeautifulSoup(texts, 'html.parser').get_text().replace('^', '<u>').replace('$', '</u>').replace(
                '`', '<strong>').replace('~', '</strong>')
            texts = text_soup.split('\xa0')
            for tex in texts:
                if 'following passage' not in tex and tex != '' and tex != ' ':
                    text += tex + '\n'
            questions = item.find(class_="col-12 p-1").find_all('p')
            question = str(questions[len(questions) - 1]).replace('<u>', '^').replace('</u>', '$').replace('<strong>',
                                                                                                           '`').replace(
                '</strong>', '~')
            question = BeautifulSoup(question, 'html.parser').get_text().replace('^', '<u>').replace('$',
                                                                                                     '</u>').replace(
                '`', '<strong>').replace('~', '</strong>')

            answer_list = item.find_all('span')
            for ans in answer_list[1:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            if text_check == text:
                sheet.write(doc_hieu_line, 10 + 6 * doc_hieu, question, style)
                sheet.write(doc_hieu_line, 12 + 6 * doc_hieu, answers, style)
                sheet.write(doc_hieu_line, 13 + 6 * doc_hieu, correct_answer, style)
                sheet.write(doc_hieu_line, 14 + 6 * doc_hieu, reason, style)
                doc_hieu += 1
                question_kind.append(kind)
            else:
                text_check = text
                sheet.write(count, 0, kind, style)
                sheet.write(count, 1, text, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
                doc_hieu_line = count
                doc_hieu = 0
                question_kind.append(kind)

        # Viet lai cau
        elif types[25] in title or types[26] in title or types[27] in title or types[28] in title or types[29] in title:
            kind = "Viết lại câu"
            question = item.find_all('p')[1].get_text()
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)

        # Tim loi sai
        elif types[30] in title or types[31] in title:
            kind = "Tìm lỗi sai"
            question = str(item.find_all('p')[1]).replace('<u>', '^').replace('</u>', '$')
            ques_soup = BeautifulSoup(question, 'html.parser').get_text().replace('^', '<u>').replace('$', '</u>')
            question = ques_soup
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)

        # ket hop cau
        elif types[32] in title or types[33] in title:
            kind = "Kết hợp câu"
            question = item.find_all('p')[1].get_text()
            answer_list = item.find_all('span')
            for ans in answer_list[0:(len(answer_list) - 1)]:
                answers += ans.get_text().replace('\xa0', '').lstrip('\n') + '\n'
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
            question_kind.append(kind)
        else:
            print(count)
            print(title)
    wb.save('thithu.xlsx')


def insert():
    links = ["https://thithu.edu.vn/ket_qua/435",
             "https://thithu.edu.vn/ket_qua/434",
             "https://thithu.edu.vn/ket_qua/431",
             "https://thithu.edu.vn/ket_qua/430",
             "https://thithu.edu.vn/ket_qua/429",
             "https://thithu.edu.vn/ket_qua/428",
             "https://thithu.edu.vn/ket_qua/427",
             "https://thithu.edu.vn/ket_qua/357",
             "https://thithu.edu.vn/ket_qua/356",
             "https://thithu.edu.vn/ket_qua/355",
             "https://thithu.edu.vn/ket_qua/354",
             "https://thithu.edu.vn/ket_qua/353",
             "https://thithu.edu.vn/ket_qua/227",
             "https://thithu.edu.vn/ket_qua/223",
             "https://thithu.edu.vn/ket_qua/222",
             "https://thithu.edu.vn/ket_qua/221",
             "https://thithu.edu.vn/ket_qua/220",
             "https://thithu.edu.vn/ket_qua/219",
             "https://thithu.edu.vn/ket_qua/218",
             "https://thithu.edu.vn/ket_qua/217",
             "https://thithu.edu.vn/ket_qua/216",
             "https://thithu.edu.vn/ket_qua/215",
             "https://thithu.edu.vn/ket_qua/214",
             "https://thithu.edu.vn/ket_qua/213",
             "https://thithu.edu.vn/ket_qua/138",
             "https://thithu.edu.vn/ket_qua/137",
             "https://thithu.edu.vn/ket_qua/136",
             "https://thithu.edu.vn/ket_qua/47"]

    sheets = []
    sheet_format = "Đề {}"

    for z, link in enumerate(links):
        sheets.append(sheet_format.format(str(z + 1)))
    for link_in in range(len(links)):
        print(links[link_in])
        sheet = wb.add_sheet(sheets[link_in], cell_overwrite_ok=True)
        crawl(links[link_in], headersz, sheet)


insert()
