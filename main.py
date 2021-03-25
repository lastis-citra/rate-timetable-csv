import csv
import os
from datetime import date
from datetime import datetime
import re
import logging
import xlsxwriter

import cloudscraper
from bs4 import BeautifulSoup


logger = logging.getLogger(__name__)


# HTMLをダウンロードしてsoupを返す
def download_html(url, file_path):
    scraper = cloudscraper.create_scraper(
        browser={
            'browser': 'chrome',
            'platform': 'windows',
            'desktop': True
        }
    )
    res = scraper.get(url)
    # print(html)

    # TODO: 保存されるHTMLがSJISになっている
    with open(file_path, mode='w') as f:
        f.write(res.text)

    soup = BeautifulSoup(res.content, 'html.parser')
    return soup


# ダウンロード済みのHTMLを読み込んでsoupを返す
def open_cache(path):
    soup = BeautifulSoup(open(path), 'html.parser')
    return soup


# 結果をExcelに出力する
# https://www.python-izm.com/third_party/excel/xlsxwriter/xlsxwriter_write/
# https://translate.google.com/translate?hl=ja&sl=en&tl=ja&u=https%3A%2F%2Fxlsxwriter.readthedocs.io%2Fformat.html&anno=2&prev=search
# result_listには行き先のリスト，分のリストが交互に入っている
def output_excel(result_list, types_list, wb, color_setting, hours, min_hour,
                 direction, dw, symbol_setting, trains_list):

    ws = wb.add_worksheet(direction + '_' + dw)

    # 開始の時に合うように先頭に足す
    lack = int(hours[0]) - int(min_hour)
    # print('lack: ' + str(lack))
    for i in range(lack):
        result_list.insert(0, list())
        result_list.insert(0, list())
        types_list.insert(0, list())
        trains_list.insert(0, list())

    # 最も長い行数を求める（多めに背景色が塗られていた方が使いやすそうなので30をデフォルトに変更）
    max_x = 30
    for results in result_list:
        if max_x < len(results):
            max_x = len(results)

    # 最も長い行に合わせて空白を足す
    results_list_added = []
    for results in result_list:
        lack = max_x - len(results)
        results_added = results
        for _ in range(lack):
            results_added.append('')
        results_list_added.append(results_added)

    # 最も長い行に合わせて空白を足す
    types_list_added = []
    for types in types_list:
        lack = max_x - len(types)
        types_added = types
        for _ in range(lack):
            types_added.append('')
        types_list_added.append(types_added)

    def create_color_dict():
        with open(color_setting, 'r', errors='replace', encoding="utf_8") as file:
            line_list = file.readlines()

            _d = dict()
            for line in line_list:
                l_dir = line.split(',')[2].replace('\n', '')
                # print(l_dir + '_' + direction)
                if l_dir == direction or l_dir == '':
                    k = line.split(',')[0]
                    v = line.split(',')[1]
                    _d[k] = v
            return _d

    d = create_color_dict()

    symbol_color_dict = dict()

    red_white = wb.add_format()
    red_white.set_font('メイリオ')
    red_white.set_size(8)
    red_white.set_font_color('red')
    red_white.set_align('center')
    red_white.set_align('top')
    red_white.set_top(1)
    red_white.set_bg_color('ffffff')
    symbol_color_dict['red_white'] = red_white

    red_grey = wb.add_format()
    red_grey.set_font('メイリオ')
    red_grey.set_size(8)
    red_grey.set_font_color('red')
    red_grey.set_align('center')
    red_grey.set_align('top')
    red_grey.set_top(1)
    red_grey.set_bg_color('f2f2f2')
    symbol_color_dict['red_grey'] = red_grey

    black_white = wb.add_format()
    black_white.set_font('メイリオ')
    black_white.set_size(8)
    black_white.set_font_color('black')
    black_white.set_align('center')
    black_white.set_align('top')
    black_white.set_top(1)
    black_white.set_bg_color('ffffff')
    symbol_color_dict['black_white'] = black_white

    black_grey = wb.add_format()
    black_grey.set_font('メイリオ')
    black_grey.set_size(8)
    black_grey.set_font_color('black')
    black_grey.set_align('center')
    black_grey.set_align('top')
    black_grey.set_top(1)
    black_grey.set_bg_color('f2f2f2')
    symbol_color_dict['black_grey'] = black_grey

    # 行き先の前方に記号を追加する
    def replace_symbol(_sheet, _row, _col, _y, _x, dest):
        with open(symbol_setting, 'r', errors='replace', encoding="utf_8") as file:
            line_list = file.readlines()
            # 白い行
            if _y % 2 == 0:
                bg_color = 'white'
            # 灰色の行
            else:
                bg_color = 'grey'

            # 行き先が空の場合，write_rich_stringでエラーになるので空にしておく
            if dest == '':
                segments = []
            else:
                segments = [symbol_color_dict['black_' + bg_color], dest]

            for line in line_list:
                attr_name = line.split(',')[0]
                attr_value = line.split(',')[1]
                symbol = line.split(',')[2]
                symbol_color = line.split(',')[3].replace('\n', '')

                trains = trains_list[_y]
                # パディングに使った空の部分は，空の状態でフォント設定だけ入れる
                if len(trains) == 0:
                    return _sheet.write(_row, _col, '', symbol_color_dict['black_' + bg_color])
                lists = list(map(lambda x: x[attr_name], trains.select('li.ek-tooltip')))

                # パディングに使った空の部分は，空の状態でフォント設定だけ入れる
                if _x >= len(lists):
                    return _sheet.write(_row, _col, '', symbol_color_dict['black_' + bg_color])
                if lists[_x] == attr_value:
                    add_segments = [symbol_color_dict[symbol_color + '_' + bg_color], symbol]
                    segments = add_segments + segments

            if len(segments) == 0:
                # [フォント設定, 行き先]の組がない場合，空の状態でフォント設定だけ入れる
                return _sheet.write(_row, _col, '', symbol_color_dict['black_' + bg_color])
            if len(segments) == 2:
                # [フォント設定, 行き先]の組が1組しかない場合もwrite_rich_stringが使えないのでwriteにする
                return _sheet.write(_row, _col, segments[1], segments[0])
            else:
                # write_rich_stringの場合，segmentsに入っているセルのフォーマットは無視されるので，最後にセルのフォーマットだけ足す
                return _sheet.write_rich_string(_row, _col, *segments, symbol_color_dict['black_' + bg_color])

    def set_time_font(_y, _x):
        _y2 = _y // 2
        train_type = types_list[_y2][_x]
        type_color = d[train_type]

        # 白い行
        if _y % 4 == 0 or _y % 4 == 1:
            bg_color = 'ffffff'
        # 灰色の行
        else:
            bg_color = 'f2f2f2'

        time_font = wb.add_format()
        time_font.set_font('メイリオ')
        time_font.set_size(11)
        time_font.set_font_color(type_color)
        time_font.set_align('center')
        time_font.set_align('vcenter')
        time_font.set_bottom(1)
        time_font.set_bg_color(bg_color)

        return time_font

    def write_list_2d(sheet, list_2d, start_row, start_col):
        for y, row in enumerate(list_2d):
            for x, cell in enumerate(row):
                row = start_row + y
                col = start_col + x

                if y % 2 == 0:
                    # segments = replace_symbol(sheet, row, col, y // 2, x, list_2d[y][x])
                    # # segmentが追加されなかった場合
                    # if len(segments) <= 2:
                    #     sheet.write(row, col, list_2d[y][x], set_dest_font(y))
                    # else:
                    #     sheet.write_rich_string(row, col, *segments)

                    replace_symbol(sheet, row, col, y // 2, x, list_2d[y][x])
                # 時刻行のフォント設定
                else:
                    sheet.write(row, col, list_2d[y][x], set_time_font(y, x))

    write_list_2d(ws, results_list_added, 2, 3)
    # セル幅の調整
    ws.set_column('A:AG', 4)


def create_time_table(table_soup, wb, dest_setting, color_setting,
                      symbol_setting, min_hour, direction, dw):
    # 行き先を省略して1文字にする
    def replace_dests(_dests):
        with open(dest_setting, 'r', errors='replace', encoding="utf_8") as file:
            d = dict(filter(None, csv.reader(file)))

        _dests_replaced = []
        for _dest in _dests:
            # 辞書dのkeyに一致するものがあれば，そのvalueで置き換える
            if _dest in d:
                _dests_replaced.append(d[_dest])
            else:
                _dests_replaced.append(_dest)

        return _dests_replaced

    tds = table_soup.select('tr.ek-hour_line td')
    # tdのlistから偶数番目を取り出して，それぞれのtextを取り出している
    hours = list(map(lambda x: x.text, tds[0::2]))
    # print(hours)

    # 種別のリスト
    types_list = []
    # 行き先のリスト
    dests_list = []
    # 分のリスト
    mins_list = []
    # 後で記号を追加するように変数で持っていく
    trains_list = []
    # tdのlistから奇数番目を取り出して処理する
    for trains in tds[1::2]:
        types = list(map(lambda x: x['data-tr-type'], trains.select('li.ek-tooltip')))
        types_list.append(types)

        dests = list(map(lambda x: x['data-dest'], trains.select('li.ek-tooltip')))

        # 行き先を1文字に変換する処理
        dests_replaced = []
        for dest in replace_dests(dests):
            dests_replaced.append(dest)

        dests_list.append(dests_replaced)

        mins = list(map(lambda x: x.text, trains.select('span.time-min')))
        mins_list.append(mins)

        trains_list.append(trains)

    # print(dests_list)
    # print(types_list)
    # print(mins_list)

    result_list = []
    for dests, mins in zip(dests_list, mins_list):
        print(','.join(dests))
        print(','.join(mins))
        result_list.append(dests)
        result_list.append(mins)

    output_excel(result_list, types_list, wb, color_setting, hours, min_hour,
                 direction, dw, symbol_setting, trains_list)


def prepare_soup(url, html_dir, name, dw):
    today = date.today()
    today_string = today.strftime('%Y%m%d')

    html_name = today_string + '_' + name + '_' + dw + '.html'
    html_path = os.path.join(html_dir, html_name)

    # 当日のキャッシュがある場合はキャッシュを利用し，なければダウンロードする
    if os.path.exists(html_path):
        soup = open_cache(html_path)
    else:
        soup = download_html(url, html_path)

    # 時刻表の更新日時を取得
    updated_date_text = soup.select_one('div.date time').text
    updated_date_tuple = re.search(r'(\d+)年(\d+)月(\d+)日現在', updated_date_text).groups()
    u_year = updated_date_tuple[0]
    u_mon = updated_date_tuple[1]
    u_day = updated_date_tuple[2]
    updated_date = datetime(int(u_year), int(u_mon), int(u_day)).strftime('%Y%m%d')
    # print(updated_date)
    return soup, updated_date


def get_each_table(wb, soup, dw, dest_setting, color_setting, symbol_setting, min_hour):
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

    # 上り
    create_time_table(tables[0], wb, dest_setting, color_setting, symbol_setting, min_hour, 'up', dw)

    # 下り（終点などでは片方向しかないため）
    if len(tables) >= 2:
        create_time_table(tables[1], wb, dest_setting, color_setting, symbol_setting, min_hour, 'down', dw)


def main_function(file_name, html_dir, excel_dir, setting_dir):
    with open(file_name, 'r', errors='replace', encoding="utf_8") as file:
        line_list = file.readlines()

    line_count = 0

    for line in line_list:
        line_count += 1
        input_url = line.split(',')[0]
        file_name = line.split(',')[1]
        dest_setting = os.path.join(setting_dir, line.split(',')[2])
        color_setting = os.path.join(setting_dir, line.split(',')[3])
        symbol_setting = os.path.join(setting_dir, line.split(',')[4])
        # 何時から始めるか（上りと下りで開始時刻が違う場合など）
        min_hour = line.split(',')[5].replace('\n', '')
        print(line_count, '/', len(line_list))
        print('input_url: ' + input_url)

        # 平日分
        input_url1 = input_url + '?dw=0'
        dw = 'weekday'
        soup, updated_date = prepare_soup(input_url1, html_dir, file_name, dw)
        excel_name = updated_date + '_' + file_name
        excel_path = os.path.join(excel_dir, excel_name + '.xlsx')

        if not os.path.exists(excel_path):
            wb = xlsxwriter.Workbook(excel_path)
            get_each_table(wb, soup, dw, dest_setting, color_setting, symbol_setting, min_hour)

            # 休日分
            input_url2 = input_url + '?dw=2'
            dw = 'holiday'
            soup, _ = prepare_soup(input_url2, html_dir, file_name, dw)
            get_each_table(wb, soup, dw, dest_setting, color_setting, symbol_setting, min_hour)

            wb.close()


if __name__ == '__main__':
    html_directory = 'html'
    excel_directory = 'excel'
    setting_directory = 'setting'
    input_file_name = './input_url_list.txt'

    if not os.path.exists(html_directory):
        os.makedirs(html_directory)

    if not os.path.exists(excel_directory):
        os.makedirs(excel_directory)

    if not os.path.exists(setting_directory):
        logging.error('There is no setting directory!!!')
        exit(1)

    main_function(input_file_name, html_directory, excel_directory, setting_directory)
