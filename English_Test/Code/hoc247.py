from bs4 import BeautifulSoup
from xlwt import Workbook
import xlwt
import requests
import re

wb = Workbook()
types = ["word whose underlined part differs",                              # phat am       0
         "underlined part from the other three in pronunciation",           # phat am       1
         "whose the underlined part that is pronounced",                    # phat am       2

         "word that differs from the other three in",          # trong am      3
         "position of primary stress",                                      # trong am      4

         "CLOSEST in meaning to the underlined",                            # tu dong nghia 5

         "OPPOSITE in meaning to the underlined",                           # tu trai nghia 6

         "Find the mistake"                                                 # tim loi sai   7
         ]

headersz = {
  'authority': 'hoc247.net',
  'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
  'accept-language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
  'cache-control': 'max-age=0',
  'cookie': '_a3rd1494487907=0-9; _a3rd1494487410=; _a3rd1623121950=0-3; __RC=4; _uidcms=1650289434347976653; _fbp=fb.1.1650289434396.1772875420; uplayer_49c0b3d79f9d123d72e249=oheDk1650289435217; _pbjs_userid_consent_data=6683316680106290; _ga=GA1.2.1039984183.1650289434; __R=1; __tb=0; ai_client_id=12100579415.1621388209; au_aid=12100579415; au_gt=1650289437; _ants_services=%5B%22cuid%22%5D; fg_version=3; fg_uuid=ac884818901b0b134279e913ef84fa18; ants_cuids=eyJhZG54cyI6eyJzZXJ2aWNlIjoiYWRueHMiLCJjdWlkIjoiNDY1NzcxOTM2NTc1Nzk2MzQwMyIs%0D%0AInN0YXR1cyI6MSwidGltZSI6MTY1Mjc3Mzg5ODY4MX19; __UF=1%2C4; _ants_utm_v2=eyJzb3VyY2UiOiJsb2NhbGhvc3QiLCJtZWRpdW0iOiJyZWZlcnJlciIsImNhbXBhaWduIjoiIiwi%0D%0AY29udGVudCI6IiIsInRlcm0iOiIiLCJ0eXBlIjoicmVmZXJyZXIiLCJ0aW1lIjoxNjUzMDI5NjUz%0D%0ANjQyLCJjaGVja3N1bSI6ImJHOWpZV3hvYjNOMExYSmxabVZ5Y21WeUxTMHhOalV6TURJNU5qVXpO%0D%0AalF5In0%3D; _gid=GA1.2.848497101.1653224021; firstTime12215_1167180=1653232138700; cdTime12215_1167180=3600; fg_lastUpdate=1653243227661; _a3rd1494487907=0-6; _a3rd1623121950=0-7; __gpi=UID=000004e8c40e3dee:T=1650289436:RT=1653289272:S=ALNI_MbgKHoy6jUSsGLz6cbnvc6fWdECDg; _pk_ses.6553273399.9283=*; _pk_ref.6553273399.9283=%5B%22%22%2C%22%22%2C1653289273%2C%22https%3A%2F%2Fwww.google.com%2F%22%5D; _gat=1; _gat_gtag_UA_93829515_1=1; fg_ucode=2e30e31339dcc45834fe3fc195938680; fg_lastModify=1653289274715; fg_guid=4986247452457670262; __IP=1962943054; Hoc247Namespace=17odcm29hhblvjcai8pk98flf8; _a3rd1494487410=0-8; cto_bundle=h6PHD185U3ZPdmclMkZ2RE1LWjEwM2ZLMEYlMkZQOXFzQXdBM0clMkZiMXF4SHp5UVJGWkczUUt5NW40JTJCeGFCSWxRS2hQckhpeXkzOGh4QlFOZ1FQTEpocFk3N0RvejJSdkxPT0Fpbk0xSXlEbUFkcE13ZG1uUzZLYXA5SEN0SGtaR1p5cU85S1ZhaUVOVVdScDVOM3gwUmJMJTJCUUxlWEV3JTNEJTNE; cto_bidid=M9uKnV9xNGc3dTJFY3FKT2FObmJEWGk0N3poSnQweDR3WVolMkJTeVExMjNaWHFtMWljeDhkQzFaZGI0NTRPN0VLbFBRNkR1TUR6SzJnVEZSMSUyRlVKUUVVUmNtUSUyQmJkRm9PQ2Y1STY3JTJCVGY5dU0zVWE4JTNE; _pk_id.6553273399.9283=e1bba0738fd9f542.1650289436.19.1653289302.1653289273.; an_session=zgzqzizizrzlzlzkzqzizdzizhzizjzjzmzkzqznzizmzdzizlzmzgzhzrzqzhzkzgzdzlzdzizlzmzgzhzrzqzhzkzgzdzizlzmzgzhzrzqzhzkzgzdzjzdzhznzdzhzd2f27zdzjzdzlzmzkzjzl; dgs=1653289301%3A3%3A0; __uif=__ui%3A1%252C4%7C__uid%3A4986247452457670262%7C__create%3A1586247452; cto_bundle=CUl9DF85U3ZPdmclMkZ2RE1LWjEwM2ZLMEYlMkZQd1RMdU00MTNDSlpjWWRBMGp1cEo5QUtEMnQlMkJQWHBhTnVGTGY3V2xNTFFTJTJCNCUyQmZmcTVhYzZrc3dOSFFuZWxTZ1FkbE5EcExKOHRmQmF6enBNZjBCeVFJaVBUZkdJamdsYiUyRmJkcWtCYnVsS05hcmUyM0dnaGZhQW9Ta1ZxaEZ4RUElM0QlM0Q; __gads=ID=5c59e50b80d5ac59:T=1650289436:S=ALNI_MayKuIZ68EbEOpuMUnoCf4hmI6jJg',
  'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="101", "Google Chrome";v="101"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"Windows"',
  'sec-fetch-dest': 'document',
  'sec-fetch-mode': 'navigate',
  'sec-fetch-site': 'none',
  'sec-fetch-user': '?1',
  'upgrade-insecure-requests': '1',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36'
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


def crawl(link, headers, sheet_name):
    sheet, style = config(sheet_name)
    page = requests.get(link, headers=headers)
    soup = BeautifulSoup(page.content, 'html.parser')
    hrefs = []
    text_check = ""
    hoan_thanh_doan_van = 0
    hoan_thanh_doan_van_line = 0
    doc_hieu = 0
    doc_hieu_line = 0
    get_hrefs = soup.find_all(class_="item-cau col-xs-2")
    for get in get_hrefs:
        hrefs.append(get.a.get('href'))

    for count, href in enumerate(hrefs):
        count += 1
        question_link = link.replace("ket-qua-test", "lam-bai-thi") + "?cau=" + str(count)
        question_page = requests.get(question_link, headers=headersz)
        question_soup = BeautifulSoup(question_page.content, 'html.parser')
        ques_item = question_soup.find(class_="fleft 162")
        get_text = str(ques_item).replace('<strong>', '`').replace('<u>', '^').replace('</u>', '$').replace('</strong>',
                                                                                                            '~')
        text = BeautifulSoup(get_text, 'html.parser').get_text().replace('`', '<strong>').replace('^', '<u>').replace(
            '$', '</u>').replace('~', '</strong>')
        answers = ""
        reason = ""
        small_page = requests.get(href, headers=headersz)
        small_soup = BeautifulSoup(small_page.content, 'html.parser')
        title = small_soup.find(class_="lch col-xs-12").find_all(class_="fleft")[1]
        title = str(title).replace('<strong>', '`').replace('<u>', '^').replace('</u>', '$').replace(
            '</strong>',
            '~')
        title = BeautifulSoup(title, 'html.parser').get_text().replace('`', '<strong>').replace('^',
                                                                                                '<u>').replace(
            '$', '</u>').replace('~', '</strong>').replace('\n', '')
        ana_title = title.split('<u>')
        point_title = title.split('.')
        get_answers = small_soup.find_all(class_="item-quest-answer")
        for ans in get_answers:
            ans = ans.find(class_="fleft")
            atod = ans.find(class_="fleft").get_text()
            ans_soup = BeautifulSoup(str(ans.find_all(class_="fleft")[1]).replace('<u>', '^').replace('</u>', '$'),
                                     'html.parser').get_text().strip('\n')
            ans = ans_soup.replace('^', '<u>').replace('$', '</u>')
            answers += atod + ans + '\n'
        correct_answer = small_soup.find(class_="dap-an").find(class_="fleft").strong.get_text()
        if correct_answer == 'A':
            correct_answer = str(1)
        elif correct_answer == 'B':
            correct_answer = str(2)
        elif correct_answer == 'C':
            correct_answer = str(3)
        elif correct_answer == 'D':
            correct_answer = str(4)
        get_reasons = small_soup.find(class_="loigiai").find_all('p')
        for rea in get_reasons:
            reason += rea.get_text() + '\n'
        # Phat am
        if types[0] in title or types[1] in title or types[2] in title:
            kind = "Phát âm"
            sheet.write(count, 0, kind, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
        # Trong am
        elif types[3] in title or types[4] in title:
            kind = "Trọng âm"
            sheet.write(count, 0, kind, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
        # Tu dong nghia
        elif types[5] in title:
            kind = "Từ đồng nghĩa"
            title = small_soup.find(class_="lch col-xs-12").find_all(class_="fleft")[1]
            title = str(title).replace('<strong>', '`').replace('<u>', '^').replace('</u>', '$').replace('</strong>',
                                                                                                         '~')
            title = BeautifulSoup(title, 'html.parser').get_text().replace('`', '<strong>').replace('^', '<u>').replace(
                '$', '</u>').replace('~', '</strong>')
            pattern = ":"
            question = re.split(pattern, title)[1]
            sheet.write(count, 0, kind, style)
            sheet.write(count, 4, question, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
        # Tu trai nghia
        elif types[6] in title:
            kind = "Từ trái nghĩa"
            title = small_soup.find(class_="lch col-xs-12").find_all(class_="fleft")[1]
            title = str(title).replace('<strong>', '`').replace('<u>', '^').replace('</u>', '$').replace('</strong>',
                                                                                                         '~')
            title = BeautifulSoup(title, 'html.parser').get_text().replace('`', '<strong>').replace('^', '<u>').replace(
                '$', '</u>').replace('~', '</strong>')
            pattern = ":"
            question = re.split(pattern, title)[1]
            sheet.write(count, 0, kind, style)
            sheet.write(count, 4, question, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
        # Tim loi sai
        elif types[7] in title:
            kind = "Tìm lỗi sai"
            title = small_soup.find(class_="lch col-xs-12").find_all(class_="fleft")[1]
            title = str(title).replace('<strong>', '`').replace('<u>', '^').replace('</u>', '$').replace('</strong>',
                                                                                                         '~')
            title = BeautifulSoup(title, 'html.parser').get_text().replace('`', '<strong>').replace('^', '<u>').replace(
                '$', '</u>').replace('~', '</strong>')
            pattern = ":"
            question = re.split(pattern, title)[1]
            sheet.write(count, 0, kind, style)
            sheet.write(count, 4, question, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)

        elif text != '\n':
            # Hoan thanh doan van
            if "__" in text:
                if text_check == text:
                    question = "(" + str(hoan_thanh_doan_van + 2) + ")"
                    sheet.write(hoan_thanh_doan_van_line, 10 + 6 * hoan_thanh_doan_van, question, style)
                    sheet.write(hoan_thanh_doan_van_line, 12 + 6 * hoan_thanh_doan_van, answers, style)
                    sheet.write(hoan_thanh_doan_van_line, 13 + 6 * hoan_thanh_doan_van, correct_answer, style)
                    sheet.write(hoan_thanh_doan_van_line, 14 + 6 * hoan_thanh_doan_van, reason, style)
                    hoan_thanh_doan_van += 1
                else:
                    text_check = text
                    kind = "Hoàn thành đoạn văn"
                    question = "(" + str(1) + ")"
                    sheet.write(count, 0, kind, style)
                    sheet.write(count, 1, text, style)
                    sheet.write(count, 4, question, style)
                    sheet.write(count, 6, answers, style)
                    sheet.write(count, 7, correct_answer, style)
                    sheet.write(count, 8, reason, style)
                    hoan_thanh_doan_van_line = count
            # Bai doc hieu
            else:
                if text_check == text:
                    question = question_soup.find_all(class_="fleft")[2].p.get_text()
                    sheet.write(doc_hieu_line, 10 + 6 * doc_hieu, question, style)
                    sheet.write(doc_hieu_line, 12 + 6 * doc_hieu, answers, style)
                    sheet.write(doc_hieu_line, 13 + 6 * doc_hieu, correct_answer, style)
                    sheet.write(doc_hieu_line, 14 + 6 * doc_hieu, reason, style)
                    doc_hieu += 1
                else:
                    text_check = text
                    question = question_soup.find_all(class_="fleft")[3].p.get_text()
                    kind = "Bài đọc hiểu"
                    sheet.write(count, 0, kind, style)
                    sheet.write(count, 1, text, style)
                    sheet.write(count, 4, question, style)
                    sheet.write(count, 6, answers, style)
                    sheet.write(count, 7, correct_answer, style)
                    sheet.write(count, 8, reason, style)
                    doc_hieu_line = count
                    doc_hieu = 0

        # Tu dong nghia va trai nghia
        elif len(ana_title) == 2:
            question = title
            if "~" in reason or "=" in reason:
                kind = "Từ đồng nghĩa"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
            else:
                kind = "Từ trái nghĩa"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)

        # Tim loi sai
        elif len(ana_title) > 3:
            kind = "Tìm lỗi sai"
            sheet.write(count, 0, kind, style)
            sheet.write(count, 4, title, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)

        elif "___" in title or "..." in title:
            if "___.\"" in title or ":\"" in title:
                question = title
                kind = "Tình huống giao tiếp"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
            else:
                question = title
                kind = "Hoàn thành câu"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
        elif len(point_title) > 2:
            if len(point_title[1]) > 0:
                question = title
                kind = "Kết hợp câu"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
            else:
                question = title
                kind = "Viết lại câu"
                sheet.write(count, 0, kind, style)
                sheet.write(count, 4, question, style)
                sheet.write(count, 6, answers, style)
                sheet.write(count, 7, correct_answer, style)
                sheet.write(count, 8, reason, style)
        else:
            kind = "Unknown"
            question = title
            sheet.write(count, 0, kind, style)
            sheet.write(count, 4, question, style)
            sheet.write(count, 6, answers, style)
            sheet.write(count, 7, correct_answer, style)
            sheet.write(count, 8, reason, style)
    sheet.write(54, 1, link, style)
    wb.save("hoc247part2.xlsx")


def insert():
    links = ['https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-bach-dang-lan-2-ktdt12002.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-khuyen-lan-2-ktdt11996.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-anh-xuan-ktdt11987.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phan-van-tri-ktdt11982.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-si-lien-lan-2-ktdt11973.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-luong-van-tuy-ktdt11966.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-chau-van-liem-ktdt11957.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-dong-khoi-ktdt11949.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-tran-ktdt11941.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-si-lien-ktdt11933.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-bui-thi-xuan-ktdt11926.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-tat-thanh-ktdt11919.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-bach-dang-ktdt11914.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-vo-nguyen-giap-ktdt11905.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-to-hien-thanh-ktdt11892.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-xuan-dieu-ktdt11879.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-thai-hoc-ktdt11872.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-chuyen-le-quy-don-ktdt11860.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-nguyen-han-ktdt11853.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-anh-xuan-ktdt11850.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-thu-duc-ktdt11832.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-linh-trung-ktdt11826.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoang-hoa-tham-ktdt11824.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tuy-phuoc-ktdt11822.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tam-quan-ktdt11818.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phu-my-1-ktdt11814.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ly-tu-trong-ktdt11802.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-may-ktdt11800.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-thi-dieu-ktdt11799.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-vo-nguyen-giap-ktdt11793.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-hung-dao-ktdt11788.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-nhan-tong-ktdt11778.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-dai-hanh-ktdt11772.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-thanh-tong-ktdt11722.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-khuyen-ktdt11721.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-khai-tri-ktdt11596.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoa-vang-ktdt11587.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-phu-ktdt11582.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-hien-ktdt11574.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-cam-le-ktdt11561.html',
             'https://hoc247.net/ket-qua-test/de-minh-hoa-ki-thi-tot-nghiep-thpt-nam-2021-mon-tieng-anh-bo-gd-dt-ktdt11068.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-van-cu-ktdt10507.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-gia-thieu-ktdt10496.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phuc-loi-ktdt10486.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-my-dinh-ktdt10479.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-xuan-thuong-ktdt10472.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tay-son-ktdt10466.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-quang-trung-ktdt10464.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-truong-dinh-ktdt10456.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-yen-hoa-ktdt10451.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-cau-giay-ktdt10444.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-khuong-dinh-ktdt10435.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tay-ho-ktdt10429.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-trai-ktdt10421.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-son-tay-ktdt10415.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-van-tam-ktdt10413.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-truong-chinh-ktdt10401.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-hong-phong-ktdt10394.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-gia-dinh-ktdt10391.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-can-vuong-ktdt10386.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-cao-ba-quat-ktdt10376.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-luong-dinh-cua-ktdt10369.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-gia-tu-ktdt10363.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-luong-the-vinh-ktdt10358.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phan-boi-chau-ktdt10352.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-van-lang-ktdt10347.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoa-lu-ktdt10339.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phan-van-tri-ktdt10331.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-thang-long-ktdt10324.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-an-nhon-ktdt10319.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-vo-truong-toan-ktdt10313.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-quang-dinh-ktdt10308.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-viet-duc-ktdt10301.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-loi-ktdt10299.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-quoc-tuan-ktdt10297.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-bui-huu-nghia-ktdt10289.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-cong-tru-ktdt10287.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoang-van-thu-ktdt10281.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-thanh-tong-ktdt10256.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-vo-truong-toan-ktdt10246.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-van-lang-ktdt10245.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hai-ba-trung-ktdt10230.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-pham-hong-thai-ktdt10229.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phan-huy-chu-ktdt10228.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-quang-trung-ktdt10227.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-kim-lien-ktdt10226.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-huynh-thuc-khang-ktdt10225.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-dinh-tien-hoang-ktdt10224.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-dong-da-ktdt10201.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-thang-long-ktdt10199.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-doi-can-ktdt10196.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-binh-trong-ktdt10190.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-quyen-ktdt10187.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-chu-van-an-ktdt10182.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-thi-dinh-ktdt10165.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-thuan-thanh-ktdt10164.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-dong-da-ktdt10152.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hung-vuong-ktdt10151.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-duy-tan-ktdt10141.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-pham-ngu-lao-ktdt10132.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-dinh-bo-linh-ktdt10108.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-quy-cap-ktdt10100.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-lai-ktdt10097.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-thanh-tong-ktdt10093.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-hong-dao-ktdt10089.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoang-van-thu-ktdt10087.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-tran-ktdt10082.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phu-cat-ktdt10076.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-tat-thanh-ktdt10073.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ho-thi-ky-ktdt10070.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-bach-dang-ktdt10065.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoa-binh-ktdt10062.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-vo-thi-sau-ktdt10059.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-van-so-ktdt10047.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-thi-dinh-ktdt10043.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-thi-minh-khai-ktdt10041.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-van-tam-ktdt10040.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-dinh-tien-hoang-ktdt10034.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-le-tan-ktdt10029.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-phan-chau-trinh-ktdt10027.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-khuyen-ktdt10023.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-tran-anh-tong-ktdt10020.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-le-loi-ktdt9997.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-ngo-quyen-ktdt9996.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-hoang-hoa-tham-ktdt9992.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-trung-truc-ktdt9990.html',
             'https://hoc247.net/ket-qua-test/de-thi-thu-thpt-qg-nam-2021-mon-tieng-anh-truong-thpt-nguyen-hue-ktdt9986.html']

    sheets = []
    sheet_format = "Đề {}"

    for z, link in enumerate(links):
        sheets.append(sheet_format.format(str(z + 1)))
    for link_in in range(len(links)):
        print(links[link_in])
        sheet = wb.add_sheet(sheets[link_in], cell_overwrite_ok=True)
        crawl(links[link_in], headersz, sheet)


insert()




