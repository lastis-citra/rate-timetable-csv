import csv

import requests
from bs4 import BeautifulSoup
import cloudscraper
import os
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles.alignment import Alignment
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side


# HTMLをダウンロードしてsoupを返す
def download_html(url):
    # html = requests.get(url).text
    scraper = cloudscraper.create_scraper(
        browser={
            'browser': 'chrome',
            'platform': 'windows',
            'desktop': True
        }
    )
    html = scraper.get(url).text
    print(html)

    dir = 'html'
    if not os.path.exists(dir):
        os.makedirs(dir)
    path = os.path.join(dir, 'test.html')

    with open(path, mode='w') as f:
        f.write(html)

    soup = BeautifulSoup(html)
    return soup


# ダウンロード済みのHTMLを読み込んでsoupを返す
def open_cache(path):
    soup = BeautifulSoup(open(path), 'html.parser')
    return soup


# 結果をExcelに出力する
# https://qiita.com/orengepy/items/d10ad53fee5593b29e46
def output_excel(list, types_list):
    wb = Workbook()
    ws = wb.active

    # 最も長い行数を求める
    max_x = 0
    for y in list:
        if max_x < len(y):
            max_x = len(y)

    # 最も長い行に合わせて空白を足す
    list_added = []
    for l in list:
        lack = max_x - len(l)
        l_added = l
        for x in range(lack):
            l_added.append("")
        list_added.append(l_added)

    types_list_added = []
    for types in types_list:
        lack = max_x - len(types)
        types_added = types
        for x in range(lack):
            types_added.append("")
        types_list_added.append(types_added)


    dest_font = Font(name='メイリオ', size=8, color='000000')
    min_font = Font(name='メイリオ', size=11, color='000000')
    dest_align = Alignment(horizontal='center', vertical='top')
    min_align = Alignment(horizontal='center', vertical='center')
    odd_fill = PatternFill(patternType='solid', fgColor='ffffff')
    even_fill = PatternFill(patternType='solid', fgColor='f2f2f2')
    side = Side(style='thin', color='000000')
    dest_border = Border(top=side)
    min_border = Border(bottom=side)

    def set_color(y, x, types_list):
        file_name = './type_color_setting.txt'
        with open(file_name, 'r', errors='replace', encoding="utf_8") as file:
            d = dict(filter(None, csv.reader(file)))

        train_type = types_list[y][x]
        type_color = d[train_type]
        return Font(name=min_font.name, size=min_font.size, color=type_color)

    def write_list_2d(sheet, list_2d, start_row, start_col):
        for y, row in enumerate(list_2d):
            for x, cell in enumerate(row):
                row = start_row + y
                col = start_col + x
                sheet.cell(row=row, column=col, value=list_2d[y][x])

                # 行き先行のフォント設定
                if y % 2 == 0:
                    sheet.cell(row=row, column=col).font = dest_font
                    sheet.cell(row=row, column=col).alignment = dest_align
                    sheet.cell(row=row, column=col).border = dest_border
                # 時刻行のフォント設定
                else:
                    sheet.cell(row=row, column=col).font = set_color(y // 2, x, types_list_added)
                    sheet.cell(row=row, column=col).alignment = min_align
                    sheet.cell(row=row, column=col).border = min_border
                # 白い行
                if y % 4 == 0 or y % 4 == 1:
                    sheet.cell(row=row, column=col).fill = odd_fill
                # 灰色の行
                else:
                    sheet.cell(row=row, column=col).fill = even_fill

    write_list_2d(ws, list, 2, 3)

    wb.save('test.xlsx')
    wb.close()


def create_time_table(table_soup):
    # 行き先を省略して1文字にする
    def replace_dests(dests):
        file_name = './dest_setting.txt'
        with open(file_name, 'r', errors='replace', encoding="utf_8") as file:
            d = dict(filter(None, csv.reader(file)))

        dests_replaced = []
        for dest in dests:
            # 辞書dのkeyに一致するものがあれば，そのvalueで置き換える
            if dest in d:
                dests_replaced.append(d[dest])
            else:
                dests_replaced.append(dest)

        return dests_replaced

    # 当駅始発の場合は●に置換する
    def replace_starts(starts):
        starts_replaced = []
        for start in starts:
            if start == "":
                starts_replaced.append(start)
            else:
                starts_replaced.append("●")

        return starts_replaced

    tds = table_soup.select('tr.ek-hour_line td')
    # tdのlistから偶数番目を取り出して，それぞれのtextを取り出している
    # 時刻のリスト，特に使っていない
    hours = list(map(lambda x: x.text, tds[0::2]))
    print(hours)

    # 種別のリスト
    types_list = []
    # 行き先のリスト
    dests_list = []
    # 分のリスト
    mins_list = []
    # tdのlistから奇数番目を取り出して処理する
    for trains in tds[1::2]:
        types = list(map(lambda x: x['data-tr-type'], trains.select('li.ek-tooltip')))
        types_list.append(types)

        dests = list(map(lambda x: x['data-dest'], trains.select('li.ek-tooltip')))
        starts = list(map(lambda x: x['data-start'], trains.select('li.ek-tooltip')))

        dests_replaced = []
        for dest, start in zip(replace_dests(dests), replace_starts(starts)):
            dests_replaced.append(start + dest)

        dests_list.append(dests_replaced)

        mins = list(map(lambda x: x.text, trains.select('span.time-min')))
        mins_list.append(mins)

    # print(dests_list)
    # print(types_list)
    # print(mins_list)

    result_list = []
    for dests, mins in zip(dests_list, mins_list):
        print(','.join(dests))
        print(','.join(mins))
        result_list.append(dests)
        result_list.append(mins)

    output_excel(result_list, types_list)


def main_function(url, name):
    # soup = download_html(url)
    dir = 'html'
    path = os.path.join(dir, 'test.html')
    soup = open_cache(path)

    # <tr class="ek-hour_line">
    #   <td>07</td>
    #   <td>
    #     <ul>
    #       <li class="ek-tooltip ek-narrow ek-train-tooltip" data-tr-type="普通" data-dest="宇都宮" data-kind_palette="t1" data-start="" data-link01="#">
    #         <a href="#" data-sf="2709" data-tx="570110-16593-2520Y" data-departure="0706" class="tooltip-data ek-train-link t1" ga-event-lbl="GA-TRAC_railway-line-station-pocket_PC_result-time">
    #           <span class="dest means-text" ga-event-lbl="GA-TRAC_railway-line-station-pocket_PC_result-time">[普通]宇</span>
    #           <span class="time-min means-text" ga-event-lbl="GA-TRAC_railway-line-station-pocket_PC_result-time">06</span>
    #         </a>
    #       </li>
    #       <li class="ek-tooltip ek-narrow ek-train-tooltip" data-tr-type="普通" data-dest="宇都宮" data-kind_palette="t1" data-start="" data-link01="#">
    #         <a href="#" data-sf="2709" data-tx="570110-16596-2522Y" data-departure="0750" class="tooltip-data ek-train-link t1" ga-event-lbl="GA-TRAC_railway-line-station-pocket_PC_result-time">
    #           <span class="dest means-text" ga-event-lbl="GA-TRAC_railway-line-station-pocket_PC_result-time">[普通]宇</span>
    #           <span class="time-min means-text" ga-event-lbl="GA-TRAC_railway-line-station-pocket_PC_result-time">50</span>
    #         </a>
    #       </li>
    #     </ul>
    #   </td>
    # </tr>
    tables = soup.select('div.search-result-body')
    create_time_table(tables[0])
    create_time_table(tables[1])


if __name__ == '__main__':
    file_name = './input_url_list.txt'
    with open(file_name, 'r', errors='replace', encoding="utf_8") as file:
        line_list = file.readlines()

    line_count = 0

    for line in line_list:
        line_count += 1
        input_url = line.split(',')[0]
        file_name = line.split(',')[1].replace('\n', '')
        print(line_count, '/', len(line_list))
        print('input_url: ' + input_url)
        main_function(input_url, file_name)
