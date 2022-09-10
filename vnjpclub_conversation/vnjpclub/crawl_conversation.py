from requests_html import HTMLSession
from bs4 import BeautifulSoup
import xlsxwriter
import re


workbook = xlsxwriter.Workbook('vnjpclub_test.xlsx')
for n in range(50):
    worksheet = workbook.add_worksheet('Bai{}'.format(str(n+1)))
    session = HTMLSession()
    link = 'https://www.vnjpclub.com/minna-no-nihongo/bai-{}-hoi-thoai.html'.format(str(n+1))
    print(link)
    r = session.get(link)
    r.html.render(timeout=100)
    page = r.html.html

    soup = BeautifulSoup(page, 'html.parser')
    contents = soup.find_all('table')
    row = 0
    ja = ''
    vi = ''
    ja_hira = ''
    for content in contents:
        lines = re.findall('<tr>((.|\n)*?)</tr>', str(content))
        for line in lines:
            kanjis = []
            hiraganas = []
            row += 1
            data = BeautifulSoup(str(line), 'html.parser')
            rubies = data.find_all('ruby')
            for ruby in rubies:
                kanji = ''
                hiragana = ''
                kanjis = re.findall('<ruby>(.*?)<rp>', str(ruby))
                if len(kanjis) > 0:
                    kanji = kanjis[0]

                hiraganas = re.findall('<rt>(.*?)</rt>', str(ruby))
                if len(hiraganas) > 0:
                    hiragana = hiraganas[0]
            cells = data.find_all('td')
            name = cells[0].text.replace('\\n', '').replace('\\xa0', '').replace('\\t', '')
            if name != '':
                worksheet.write(row, 0, name)
                for index, cell in enumerate(cells[1:]):
                    cell = str(cell).replace('<u>', '`').replace('</u>', '~')
                    ja = BeautifulSoup(cell, 'html.parser').find_all(class_='candich')
                    if len(ja) > 0:
                        ja_soup = BeautifulSoup(str(ja[0]), 'html.parser')
                        ja = ja[0].text.replace('~', '</u>').replace('`', '<u>').replace('\\t', '').replace('\\u3000', ' ').replace('\\n', '').replace('\\xa0', '')
                        ja_hira = ja
                        rubies = ja_soup.find_all('ruby')
                        for ruby in rubies:
                            kanjis = re.findall('<ruby>(.*?)<rp>', str(ruby))
                            hiraganas = re.findall('<rt>(.*?)</rt>', str(ruby))
                            if len(kanjis) > 0 and len(hiraganas) > 0:
                                ja_hira = ja_hira.replace(kanjis[0], hiraganas[0])
                    else:
                        ja = ''
                        ja_hira = ''
                    vi = BeautifulSoup(cell, 'html.parser').find_all(class_='nddich')
                    if len(vi) > 0:
                        vi = vi[0].text.replace('~', '</u>').replace('`', '<u>').replace('\\t', '').replace('\\u3000', ' ').replace('\\n', '').replace('\\xa0', '')
                    else:
                        vi = ''
            else:
                for index, cell in enumerate(cells[1:]):
                    cell = str(cell).replace('<u>', '`').replace('</u>', '~')
                    ja_add = BeautifulSoup(cell, 'html.parser').find_all(class_='candich')
                    if len(ja_add) > 0:
                        ja_soup = BeautifulSoup(str(ja_add[0]), 'html.parser')
                        ja_add = ja_add[0].text.replace('~', '</u>').replace('`', '<u>').replace('\\t', '').replace('\\u3000', ' ').replace('\\n', '').replace('\\xa0', '')
                        ja_hira_add = ja_add
                        rubies = ja_soup.find_all('ruby')
                        for ruby in rubies:
                            kanjis = re.findall('<ruby>(.*?)<rp>', str(ruby))
                            hiraganas = re.findall('<rt>(.*?)</rt>', str(ruby))
                            if len(kanjis) > 0 and len(hiraganas) > 0:
                                ja_hira_add = ja_hira_add.replace(kanjis[0], hiraganas[0])
                    else:
                        ja_add = ''
                        ja_hira_add = ''
                    vi_add = BeautifulSoup(cell, 'html.parser').find_all(class_='nddich')
                    if len(vi_add) > 0:
                        vi_add = vi_add[0].text.replace('~', '</u>').replace('`', '<u>').replace('\\t', '').replace('\\u3000', ' ').replace('\\n', '').replace('\\xa0', '')
                    else:
                        vi_add = ''
                    ja += '\n' + ja_add
                    ja_hira += '\n' + ja_hira_add
                    vi += '\n' + vi_add
                row -= 1
            worksheet.write(row, 1, ja)
            worksheet.write(row, 3, vi)
            worksheet.write(row, 2, ja_hira)
        row += 1
        worksheet.write(row, 0, '#####')

workbook.close()


