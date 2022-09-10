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


from bs4 import BeautifulSoup
import requests


headers = {
  'authority': 'tracnghiem.net',
  'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
  'accept-language': 'vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5',
  'cache-control': 'max-age=0',
  'cookie': '_uidcms=1652769154832446355; __oagr=true; _ga=GA1.2.1069748650.1652769155; __tb=0; _pbjs_userid_consent_data=6683316680106290; fg_version=3; fg_uuid=ac884818901b0b134279e913ef84fa18; __UF=1%2C4; fg_lastUpdate=1652971073598; _gid=GA1.2.206039740.1653233423; __gpi=UID=0000056d1dc681b1:T=1652769154:RT=1653271591:S=ALNI_MZvi1zXaYcg0HRrcUagXwSnwGwuPQ; fg_ucode=2e30e31339dcc45834fe3fc195938680; fg_lastModify=1653271595507; fg_guid=4986247452457670262; __RC=4; __R=1; __IP=1962943054; _gat_gtag_UA_177772853_1=1; cto_bundle=G4FC7l9jRyUyRmRFMDdqRFdMZkFzJTJCVmF2VEVTZVd6ZzZuSWlzZGdkcm5mZGN5TWpyaDRNbHZGQTVpdiUyQkxVOWhIc1RyV3Q5Q1BneHY3RTB5ZUElMkY3V0FpY2VqdzV5V2t5eEU4NDJ6UTZPamE3amVLazVvN0RoOGFXbjhRbjF3bmprYWNYcDBYdzkzUW5kJTJCcEo2Q1lKR0clMkJ0Z0t1WHclM0QlM0Q; cto_bidid=R6RtGF9kMVlENXlhY0hXMHlVZUZMV2g4U3ZhdzgyNSUyRjdoa290THBIbnNyNWJ3aiUyRkxGa0VlRklia0lxZGpzVldaSWQ1U0NBcG90RmtMR3BZVGJRYzNLU2F2OHl0ckdhVGglMkYlMkZlZlA2cTVVd1ZhcUQ3WHFoMEIwSjRwJTJGVmpUU2RQUWI1MDY; XSRF-TOKEN=eyJpdiI6IjlMdlM5VFhWZ3FTb2hQTHBvOVBtL0E9PSIsInZhbHVlIjoiOEw5K256S201eXBGYUxNamtEejZiMGZnOVZtNWxPOHZHRUdwbmFuMTJXdlFmczQ1cXZUaFZWODlMK0gwamw5TiIsIm1hYyI6IjZlMDAyZTE1NTJjNTJhMzM0MzkwOTk4MGQzMWVkMWIyNWYzMzFiOTliMjI5YmUzZGY1NGNmMjhiMzRiMjAzZDIifQ%3D%3D; tracnghiem_session=eyJpdiI6Ik5uNXBldEdBc2JjbFpYMlh6SU1oYXc9PSIsInZhbHVlIjoiY2tQVFZlbk1CUHBTbjE3b3FTMUl0Q3ZFMEg0TGVyb2RTNVcvWkc2WFl2SWl0cVA0M1BjbG1FUDZ0SEUzdy9mbyIsIm1hYyI6IjA2YzE0MzhmN2ViMWMzMWYxZGYyMGUyYzVlNGIwMTQ3ZWU2N2Q4NDQxMWQ5MjJmNGU5NWI1NTdiNGE3YmYxYWMifQ%3D%3D; __gads=ID=849e9baec59749aa-2223a81c4fd30021:T=1652769154:RT=1653273382:S=ALNI_MYp3cMAthZHaCN5pQRJlX44P_9z3Q; __uif=MWFlZTc1MTdkZmYyZDNmMjhiMjY2OTUzZTQ0YzQ4N2YyOWU4ZDEyMGM2ODUifQ%3D%3D%7C__ui%3A1%252C4%7C__uid%3A4986247452457670262%7C__create%3A1586247452; MgidStorage=%7B%220%22%3A%7B%22svspr%22%3A%22%22%2C%22svsds%22%3A39%7D%2C%22C1216611%22%3A%7B%22page%22%3A20%2C%22time%22%3A1653273384874%7D%7D; XSRF-TOKEN=eyJpdiI6InREQVRxZjVxOGFTUzRKZHY5ZFd0OVE9PSIsInZhbHVlIjoid0Z3b2FhWHhMeUVaTzRyamRqenpHcGhUVDdGL1lYREsvZmVzbnQ5NHV4T0ZtbUp2eHFEcVJiYWNRdFViRnRNdSIsIm1hYyI6IjI3YmQ1MjQyYTJhMWU1MmYzOWZhZjhmODUzMTBiYTY1NDA4MGRiY2RjMDQ3ODJiZGFjZGFhMjUzNGE5MjgzODcifQ%3D%3D; tracnghiem_session=eyJpdiI6IkVoNDBmZXg0akg0U205bmtCVGlHTXc9PSIsInZhbHVlIjoibHo3OE1jVFhWUEtDT2lCSllsQW8vRTNTYUQrUE0xd3hRbTBrSGNFazZQRWxGUGpjcUJIb2dNYlcyNm5pY3JmVyIsIm1hYyI6ImU5ZDJkYmJiZjlkODEwMjZlZTAyMDNlNmRhMzVlNDNhZWNlYzAxMmFmMzRhY2FjZWNlZWIzZjg1ZmYwYjBjOTQifQ%3D%3D',
  'referer': 'https://tracnghiem.net/thptqg/',
  'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="101", "Google Chrome";v="101"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"Windows"',
  'sec-fetch-dest': 'document',
  'sec-fetch-mode': 'navigate',
  'sec-fetch-site': 'same-origin',
  'sec-fetch-user': '?1',
  'upgrade-insecure-requests': '1',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36'
}

link_pages = []
link_pages_format = "https://tracnghiem.net/thptqg/tieng-anh/?p={}"
for n in range(24):
    link_pages.append(link_pages_format.format(str(n+1)))
for link_page in link_pages:
    print(link_page)
    page = requests.get(link_page, headers=headers)
    soup = BeautifulSoup(page.content, 'html.parser')
    tests = soup.find_all(class_="d9Box part-item")
    link_results = []
    for test in tests:
        test = test.replace('qg/', 'qg/ket-qua-lam-bai/')
        link_results.append(test)
    print(link_results)

