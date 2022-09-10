from bs4 import BeautifulSoup
import requests
import re
import xlwt
from xlwt import Workbook

wb = Workbook()


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


types = [
    "whose underlined part differs from",                           # phat am
    "the word that differs from the rest in the pronunciation",     # phat am
    "is pronounced differently from that of the rest",              # phat am
    "differs from the other three in the position",                 # trong am
    "differs from the rest in the position of the main stress",     # trong am
    "indicate the correct answer to each of the",                   # hoan thanh cau
    "CLOSEST in meaning to the underlined word",                    # tu dong nghia
    "OPPOSITE in meaning to the underlined word",                   # tu trai nghia
    "word or phrase that best fits",                                # hoan thanh doan van
    "indicate the correct word for each of the blanks",             # hoan thanh doan van
    "indicate the underlined part that needs",                      # tim loi sai
    "show the underlined part that needs correction.",              # tim loi sai
    "answer to each of the questions",                              # doc hieu
    "indicate the sentence that is closest in meaning to each of",  # viet lai cau
    "indicate the sentence that best combines each pair",           # ket hop cau
    "following exchanges"  # tinh huong giao tiep
    ]


def crawl_test(linkz, headersz, sheet_name):
    question_kind = []
    sheet = config(sheet_name)
    this_page = requests.get(linkz, headers=headersz)
    soup = BeautifulSoup(this_page.content, 'html.parser')
    items = soup.find_all(class_="quiz-answer-item")
    hoan_thanh_doan_van = 0
    hoan_thanh_doan_van_line = 0
    doc_hieu = 0
    doc_hieu_line = 0
    for count, item in enumerate(items):
        count += 1
        print(count)
        list_answer = []
        question = ""
        ques = str(item.find(class_="question-name"))
        ques_soup = BeautifulSoup(ques, 'html.parser')
        get_list = ques_soup.get_text().replace('\xa0', ' ').split('\n')
        for n in get_list:
            if n != '':
                question = n
                break
        # Phat am
        if types[0] in question or types[1] in question or types[2] in question:
            kind = "Phát âm"
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                answer = str(answer).replace('<u>', '^').replace('</u>', '$')
                ans_soup = BeautifulSoup(answer, 'html.parser')
                answer = ans_soup.get_text().replace('^', '<u>').replace('$', '</u>').strip('\n').replace('\xa0', '')
                list_answer.append(answer + '\n')
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Trong am
        elif types[3] in question or types[4] in question:
            kind = "Trọng âm"
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                list_answer.append(ans + '\n')
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")

            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Hoan thanh cau
        elif types[5] in question and "Mark the letter" in question:
            kind = "Hoàn thành câu"
            closest = 'Mark the letter A, B, C'
            question = [x for x in item.find(class_="question-name").get_text().split('\n') if
                        closest not in x and x != '']
            question = question[0]
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text()
                list_answer.append(ans.lstrip("\n"))
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Tu dong nghia
        elif types[6] in question:
            kind = "Từ đồng nghĩa"
            ques = str(item.find(class_="question-name")).replace('<u>', '^').replace('</u>', '$')
            ques_soup = BeautifulSoup(ques, 'html.parser')
            get_list = ques_soup.get_text().replace('^', '<u>').replace('$', '</u>').replace('\xa0', ' ').split('\n')
            question = [x for x in get_list if x != '' and types[6] not in x][0]
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text()
                list_answer.append(ans.lstrip("\n"))
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Tu trai nghia
        elif types[7] in question:
            kind = "Từ trái nghĩa"
            ques = str(item.find(class_="question-name")).replace('<u>', '^').replace('</u>', '$')
            ques_soup = BeautifulSoup(ques, 'html.parser')
            get_list = ques_soup.get_text().replace('^', '<u>').replace('$', '</u>').replace('\xa0', ' ').split('\n')
            question = [x for x in get_list if x != '' and types[7] not in x][0]

            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text()
                list_answer.append(ans.strip("\n") + "\n")
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Hoan thanh doan van
        elif types[8] in question or types[9] in question:
            kind = "Hoàn thành đoạn văn"
            check = "Read the following passage"
            text = ''
            for x in item.find(class_="question-name").get_text().split('\n'):
                if check not in x and x != '':
                    text += x + '\n'
            text = text.replace('\xa0', '')
            question_1 = "(" + str(1) + ")"
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text()
                list_answer.append(ans.lstrip("\n"))
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 1, text)
            sheet.write(count, 4, question_1)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)
            hoan_thanh_doan_van_line = count

        # Tim loi sai
        elif types[10] in question or types[11] in question:
            kind = "Tìm lỗi sai"
            closest = 'Mark the letter A, B, C'
            question = str(item.find(class_="question-name"))
            text = question.replace('<u>', '^').replace('</u>', '$')
            ques_soup = BeautifulSoup(text, 'html.parser')
            text = ques_soup.get_text().replace('^', '<u>').replace('$', '</u>')

            question = [x for x in text.split('\n') if closest not in x and x != '']
            question = question[0]
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text()
                list_answer.append(ans.lstrip("\n"))
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '').replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Bai doc hieu
        elif types[12] in question:
            kind = "Bài đọc hiểu"
            text = ""
            question = item.find(class_="question-name").get_text().strip()
            paras = question.split("\n")
            ss = len(paras) - 1
            for para in paras[1:ss]:
                text += para + "\n"
            question_1 = paras[ss]
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text().strip('\n').replace('\xa0', '').rstrip() + '\n'
                list_answer.append(ans)
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 1, text)
            sheet.write(count, 4, question_1)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)
            doc_hieu_line = count
            doc_hieu = 0

        # Viet lai cau
        elif types[13] in question:
            kind = "Viết lại câu"
            closest = 'Mark the letter A, B, C'
            question = [x for x in item.find(class_="question-name").get_text().split('\n') if
                        closest not in x and x != '']
            question = question[0]
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                list_answer.append(ans + '\n')
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Ket hop cau
        elif types[14] in question:
            kind = "Kết hợp câu"
            closest = 'Mark the letter A, B, C'
            question = [x for x in item.find(class_="question-name").get_text().split('\n') if
                        closest not in x and x != '']
            question = question[0]
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                list_answer.append(ans + '\n')
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        # Tinh huong giao tiep
        elif ': “_' in question or types[15] in question:
            kind = "Tình huống giao tiếp"
            list_answer = []
            question = ''
            closest = 'Mark the letter A, B, C'
            for x in item.find(class_="question-name").get_text().split('\n'):
                if closest not in x and x != '':
                    question += x + '\n'
            answers = item.find_all(class_="anwser-item")
            for answer in answers:
                ans = answer.get_text()
                list_answer.append(ans.strip('\n') + '\n')
            correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
            if re.match('^A', correct_answer):
                correct_answer = str(1)
            elif re.match('^B', correct_answer):
                correct_answer = str(2)
            elif re.match('^C', correct_answer):
                correct_answer = str(3)
            else:
                correct_answer = str(4)
            reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
            sheet.write(count, 0, kind)
            sheet.write(count, 4, question)
            sheet.write(count, 6, list_answer)
            sheet.write(count, 7, correct_answer)
            sheet.write(count, 8, reason)
            question_kind.append(kind)

        else:
            if question_kind[count-2] == 'Hoàn thành câu':
                list_answer = []
                question = item.find(class_="question-name").get_text().strip()
                kind = "Hoàn thành câu"
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            elif question_kind[count-2] == 'Từ đồng nghĩa':
                ques_name = item.find(class_='question-name')
                list_answer = []
                kind = "Từ đồng nghĩa"
                text = str(ques_name)
                textz = text.replace('<u>', '^').replace('</u>', '$')
                soup = BeautifulSoup(textz, 'html.parser').get_text()
                question = soup.replace('^', '<u>').replace('$', '</u>')
                question = [x for x in question.split('\n') if 'Mark the letter' not in x and x != '']
                question = question[0]
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            elif question_kind[count-2] == 'Từ trái nghĩa':
                ques_name = item.find(class_='question-name')
                list_answer = []
                kind = "Từ trái nghĩa"
                text = str(ques_name)
                textz = text.replace('<u>', '^').replace('</u>', '$')
                soup = BeautifulSoup(textz, 'html.parser').get_text()
                question = soup.replace('^', '<u>').replace('$', '</u>')
                question = [x for x in question.split('\n') if 'Mark the letter' not in x and x != '']
                question = question[0]
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            elif question_kind[count-2] == 'Hoàn thành đoạn văn':
                list_answer = []
                kind = 'Hoàn thành đoạn văn'
                question = "(" + str(hoan_thanh_doan_van + 2) + ")"
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(hoan_thanh_doan_van_line, 10 + 6 * hoan_thanh_doan_van, question)
                sheet.write(hoan_thanh_doan_van_line, 12 + 6 * hoan_thanh_doan_van, list_answer)
                sheet.write(hoan_thanh_doan_van_line, 13 + 6 * hoan_thanh_doan_van, correct_answer)
                sheet.write(hoan_thanh_doan_van_line, 14 + 6 * hoan_thanh_doan_van, reason)
                hoan_thanh_doan_van += 1
                question_kind.append(kind)

            elif question_kind[count-2] == 'Tìm lỗi sai':
                list_answer = []
                kind = "Tìm lỗi sai"
                closest = 'Mark the letter A, B, C'
                question = str(item.find(class_="question-name"))
                text = question.replace('<u>', '^').replace('</u>', '$')
                ques_soup = BeautifulSoup(text, 'html.parser')
                text = ques_soup.get_text().replace('^', '<u>').replace('$', '</u>')
                question = [x for x in text.split('\n') if closest not in x and x != '']
                question = question[0]
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            elif question_kind[count-2] == 'Bài đọc hiểu':
                list_answer = []
                kind = 'Bài đọc hiểu'
                question = item.find(class_="question-name").get_text().strip()
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(doc_hieu_line, 10 + 6 * doc_hieu, question)
                sheet.write(doc_hieu_line, 12 + 6 * doc_hieu, list_answer)
                sheet.write(doc_hieu_line, 13 + 6 * doc_hieu, correct_answer)
                sheet.write(doc_hieu_line, 14 + 6 * doc_hieu, reason)
                doc_hieu += 1
                question_kind.append(kind)

            elif question_kind[count-2] == 'Viết lại câu':
                list_answer = []
                closest = 'Mark the letter A, B, C'
                question = [x for x in item.find(class_="question-name").get_text().split('\n') if
                            closest not in x and x != '']
                question = question[0]
                kind = "Viết lại câu"
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            elif question_kind[count-2] == 'Kết hợp câu':
                list_answer = []
                closest = 'Mark the letter A, B, C'
                question = [x for x in item.find(class_="question-name").get_text().split('\n') if
                            closest not in x and x != '']
                question = question[0]
                kind = "Kết hợp câu"
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            elif question_kind[count-2] == 'Tình huống giao tiếp':
                kind = "Tình huống giao tiếp"
                list_answer = []
                question = ''
                closest = 'Mark the letter A, B, C'
                for x in item.find(class_="question-name").get_text().split('\n'):
                    if closest not in x and x != '':
                        question += x + '\n'
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text()
                    list_answer.append(ans.strip('\n') + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

            else:
                list_answer = []
                closest = 'Mark the letter A, B, C'
                question = [x for x in item.find(class_="question-name").get_text().split('\n') if
                            closest not in x and x != '']
                question = question[0]
                kind = "Viết lại câu"
                answers = item.find_all(class_="anwser-item")
                for answer in answers:
                    ans = answer.get_text().replace('\n', '').replace('\xa0', '')
                    list_answer.append(ans + '\n')
                correct_answer = item.find(class_="anwser-item col-xs-12 d-flex correct").get_text().strip("\n")
                if re.match('^A', correct_answer):
                    correct_answer = str(1)
                elif re.match('^B', correct_answer):
                    correct_answer = str(2)
                elif re.match('^C', correct_answer):
                    correct_answer = str(3)
                else:
                    correct_answer = str(4)
                reason = item.find(class_="question-reason").get_text().lstrip("\n").replace('\xa0', '')
                sheet.write(count, 0, kind)
                sheet.write(count, 4, question)
                sheet.write(count, 6, list_answer)
                sheet.write(count, 7, correct_answer)
                sheet.write(count, 8, reason)
                question_kind.append(kind)

    wb.save('test.xlsx')


headers = {
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
  'Accept-Language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
  'Cache-Control': 'max-age=0',
  'Connection': 'keep-alive',
  'Cookie': 'cross-site-cookie=bar; cross-site-cookie=bar; _ga=GA1.2.1917927615.1650034424; _fbp=fb.1.1652059911518.671998356; _gid=GA1.2.210150947.1652059912; __gads=ID=5ea869403150deeb-229c3e4128d300ef:T=1650034424:RT=1652434308:S=ALNI_MYZhTGK0sK9WB37lRota-H5aYFU0Q; __gpi=UID=00000538d8b08ef4:T=1652060124:RT=1652541357:S=ALNI_MZ_qBxxPTHdfMVeQ7wdLpFGbBAoqQ; _gat_gtag_UA_65991607_2=1; XSRF-TOKEN=eyJpdiI6IlBcL0xUYlBIeGRKend0YUcwS3k2RlFRPT0iLCJ2YWx1ZSI6ImVyeDdHanlBNW1Ecmx1bXF5UUVkYytzMEcxV1lvellKazFRS05ZekNISjVwVzlsVzk1R3FKOWJqR0hQdDZ1ZXkiLCJtYWMiOiJhYmJkNTBjODkwNmFjN2JjYTMyYTEzYmE2MDM0YTE3MjJjMmZkYmMyYjkwM2ExNzRiYWM2ZTg4MTU0YzNmMGQ5In0%3D; khoahocvietjackcom_session=eyJpdiI6Ik9vUEVKSjRJMVRFMnE3eUJZSUxGdFE9PSIsInZhbHVlIjoiTzVDTHIySjlxN3dzQVo0dFwvSVk0RkxaXC92b3RBeUlPcHliV2NIM1FmSjgrUTNnd1BPbmZyYlFFV2I3VE9VWk41ays5Nlwvcm5wOG50Z0Y2WG9PQ21iQ240dldkVzRYK3BKb21LRW5jRjZhWkpVZlZoTVFmZ2tpcnpuTm5SMzZrTEciLCJtYWMiOiI1MmM3OGU1OTU0NDEwOTFhZjJmMWFhMzcxMGE0YWZjMjliODEzMGU2NmFkMDVmMGY5MDRiNjFlN2U5M2M1MTU1In0%3D; XSRF-TOKEN=eyJpdiI6Ilp6V0RsMExLbW1TUmw0RzAxak9KK3c9PSIsInZhbHVlIjoiTFVsZXBHRFVzRnNvWVZDK3hWeDJKczZobXZpTFdUd0pMQUd6UVo4dWhJQUJZMUJuWmhreGxOV21OOW12eXRiWSIsIm1hYyI6IjE2OTk1NmE3ZjIxNzQzMzVmMTkwZTAyMjA4OWY3NzNkZGZkYWNlYTE2MzAxOTM0YWMwNmFkNTBhZDk1YTRhNWYifQ%3D%3D; khoahocvietjackcom_session=eyJpdiI6InhJOEkyZm1PcmRDNGtYNWZTUWk4WGc9PSIsInZhbHVlIjoiU2hUdlJxYjYzQ2gyNXhSV3pydHoxTFwvQm5OSWFnN1BZMWc1Z1ZZVWZEQ0d6STJSWURkZHRVaEFRcGtwWFU0bmliWERHdVJOTWsyWTM1NHdhT0pXeHg0Qytxb0tmait4QUNla2pmTStaZXdBbXcwbklYSmpGaE5MaGJiVWxwSzdUIiwibWFjIjoiZDQzMTJhMDA1ODRhMDNmMmQ2ZGM4YTY5MjI5OWJiNjcxN2M5MjFmMzAxNzFiMjNhMjgxODcwOWI3YzZjMjQzNSJ9',
  'Referer': 'https://khoahoc.vietjack.com/thi-online/bo-35-de-thi-minh-hoa-tieng-anh-co-dap-an-chi-tiet-nam-2022/76519/thi',
  'Sec-Fetch-Dest': 'document',
  'Sec-Fetch-Mode': 'navigate',
  'Sec-Fetch-Site': 'same-origin',
  'Sec-Fetch-User': '?1',
  'Upgrade-Insecure-Requests': '1',
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36',
  'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="101", "Google Chrome";v="101"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"Windows"'
}


def insert():
    # lay danh sach link
    # links = []
    # page = requests.get(link_test, headers=headers)
    # soup = BeautifulSoup(page.content, 'html.parser')
    # tab_exam = soup.find(class_="tab-exam")
    # for link in tab_exam.find_all('a'):
    #     links.append(link.get('href') + '/ket-qua')
    links = ['https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76474/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76478/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76485/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76505/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76515/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76522/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76571/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76613/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76622/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76648/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76684/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76690/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76693/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76698/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76826/ket-qua',
             'https://khoahoc.vietjack.com/thi-online/bo-20-de-thi-thu-thpt-quoc-gia-mon-tieng-anh-co-dap-an-nam-2022-20-de/76841/ket-qua']

    sheets = []
    sheet_format = "Đề {}"

    for z, link in enumerate(links):
        sheets.append(sheet_format.format(str(z + 1)))
    for n in range(len(links)):
        print(links[n])
        sheet = wb.add_sheet(sheets[n], cell_overwrite_ok=True)
        crawl_test(links[n], headers, sheet)


insert()
