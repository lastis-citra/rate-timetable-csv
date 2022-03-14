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
    res = scraper.get('https://ekitan.com/timetable/railway/line-station/' + url)
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
# TODO: 停車駅が微妙に違うパターンの記号や種別色分けもやりたい
def output_excel(dests_list, mins_list, types_list, wb, color_setting, hours, min_hour,
                 direction, dw, symbol_setting, trains_list):
    print(trains_list[0])

    ws = wb.add_worksheet(direction + '_' + dw)

    # 開始の時に合うように先頭に足す
    lack = int(hours[0]) - int(min_hour)
    # print('lack: ' + str(lack))
    for i in range(lack):
        dests_list.insert(0, list())
        mins_list.insert(0, list())
        types_list.insert(0, list())
        trains_list.insert(0, list())

    # 最も長い行数を求める（多めに背景色が塗られていた方が使いやすそうなので30をデフォルトに変更）
    max_x = 30
    for results in mins_list:
        if max_x < len(results):
            max_x = len(results)

    # 最も長い行に合わせて空白を足す
    def add_space(line):
        l_added = []
        for res in line:
            _lack = max_x - len(res)
            res_added = res
            for _ in range(_lack):
                res_added.append('')
            l_added.append(res_added)
        return l_added

    dests_list_added = add_space(dests_list)
    mins_list_added = add_space(mins_list)
    types_list_added = add_space(types_list)

    time_color_dict = dict()
    time_bg_color_dict = dict()

    def create_color_dict():
        with open(color_setting, 'r', errors='replace', encoding="utf_8") as file:
            line_list = file.readlines()

            for line in line_list:
                l_dir = line.split(',')[3].replace('\n', '')
                # print(l_dir + '_' + direction)
                if l_dir == direction or l_dir == '':
                    k = line.split(',')[0]
                    v1 = line.split(',')[1]
                    v2 = line.split(',')[2]
                    time_color_dict[k] = v1
                    time_bg_color_dict[k] = v2

    create_color_dict()

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
        train_type = types_list_added[_y2][_x]
        type_color = time_color_dict[train_type]
        type_bg_color = time_bg_color_dict[train_type]

        time_font = wb.add_format()

        # 条件フォントのテスト中 TODO:
        trains = trains_list[_y2]
        if len(trains) != 0 and direction == 'down':
            # 各列車のhrefのtx=の値を取り出したい
            lists = list(map(lambda x: x['href'].split('tx=')[1].split('&dw=')[0], trains.select('a')))
            if _x < len(lists) and re.compile('-1[0-9][0-9]M$').search(lists[_x]):
                type_color = 'ff0000'
            elif _x < len(lists) and re.compile('Y$').search(lists[_x]):
                time_font.set_underline(1)

        # 背景色がある種別の場合はそちらを優先する
        if type_bg_color != '':
            bg_color = type_bg_color
        # 白い行
        elif _y % 4 == 0 or _y % 4 == 1:
            bg_color = 'ffffff'
        # 灰色の行
        else:
            bg_color = 'f2f2f2'

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
                    # 行き先行のフォント設定
                    replace_symbol(sheet, row, col, y // 2, x, list_2d[y][x])
                else:
                    # 時刻行のフォント設定
                    sheet.write(row, col, list_2d[y][x], set_time_font(y, x))

    results_list_added = []
    for dests, mins in zip(dests_list_added, mins_list_added):
        results_list_added.append(dests)
        results_list_added.append(mins)
    # print(results_list_added)

    write_list_2d(ws, results_list_added, 2, 3)
    # セル幅の調整
    ws.set_column('A:AG', 4)


def create_time_table(table_soup, dest_setting):
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

    # result_list = []
    # for dests, mins in zip(dests_list, mins_list):
    #     print(','.join(dests))
    #     print(','.join(mins))
    #     result_list.append(dests)
    #     result_list.append(mins)

    return dests_list, mins_list, types_list, trains_list, hours


# time_date: 特定の日付の時刻表を取得したい場合．空文字の場合は今日が基準
def prepare_soup(url, html_dir, name, dw, time_date):
    if time_date == '':
        today = date.today()
        day_string = today.strftime('%Y%m%d')
    else:
        day_string = time_date

    html_name = day_string + '_' + name + '_' + url.split('/')[0] + '_' + dw + '.html'
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


def get_each_table(soup, reverse_flag):
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
    # reverse_flagがTrueの場合は上下のテーブルを逆にする
    if len(tables) >= 2 and reverse_flag:
        table_up = tables[1]
        table_down = tables[0]
    elif len(tables) >= 2:
        table_up = tables[0]
        table_down = tables[1]
    else:
        table_up = tables[0]
        table_down = None

    return table_up, table_down


def join_lists(dests_list, mins_list, types_list, trains_list, hours,
               _dests_list, _mins_list, _types_list, _trains_list, _hours, min_hour):
    # result_listは2列ずつ入っているので注意
    joined_dests_list = []
    joined_mins_list = []
    joined_types_list = []
    joined_trains_list = []
    joined_hours = []
    hours_int = list(map(lambda x: int(x), hours))
    _hours_int = list(map(lambda x: int(x), _hours))

    if len(hours) == 0:
        return _dests_list, _mins_list, _types_list, _trains_list, _hours

    for i in range(int(min_hour), 25):
        i_mod = i
        if i >= 24:
            i_mod = i - 24

        if i_mod in hours_int and i_mod in _hours_int:
            index = hours_int.index(i_mod)
            _index = _hours_int.index(i_mod)

            soup_text = '<td><ul>'
            dests1 = dests_list[index]
            dests2 = _dests_list[_index]
            mins1 = mins_list[index]
            mins2 = _mins_list[_index]
            types1 = types_list[index]
            types2 = _types_list[_index]
            lis1 = trains_list[index].select('li')
            lis2 = _trains_list[_index].select('li')
            joined_dests = []
            joined_mins = []
            joined_types = []

            while len(mins1) > 0 or len(mins2) > 0:
                if len(mins1) > 0 and len(mins2) > 0:
                    # 時刻が小さいものから取り出す
                    if int(mins1[0]) < int(mins2[0]):
                        joined_dests.append(dests1.pop(0))
                        joined_mins.append(mins1.pop(0))
                        joined_types.append(types1.pop(0))
                        soup_text += lis1.pop(0).prettify()
                    elif int(mins1[0]) > int(mins2[0]):
                        joined_dests.append(dests2.pop(0))
                        joined_mins.append(mins2.pop(0))
                        joined_types.append(types2.pop(0))
                        soup_text += lis2.pop(0).prettify()
                    else:
                        # 行き先も時刻も同じものが2つある場合は片方を捨てる
                        if dests1[0] == dests2[0]:
                            dests1.pop(0)
                            mins1.pop(0)
                            types1.pop(0)
                            lis1.pop(0)
                        else:
                            joined_dests.append(dests1.pop(0))
                            joined_mins.append(mins1.pop(0))
                            joined_types.append(types1.pop(0))
                            soup_text += lis1.pop(0).prettify()
                elif len(mins1) > 0:
                    joined_dests.append(dests1.pop(0))
                    joined_mins.append(mins1.pop(0))
                    joined_types.append(types1.pop(0))
                    soup_text += lis1.pop(0).prettify()
                elif len(mins2) > 0:
                    joined_dests.append(dests2.pop(0))
                    joined_mins.append(mins2.pop(0))
                    joined_types.append(types2.pop(0))
                    soup_text += lis2.pop(0).prettify()

            soup_text += '</td></ul>'

            joined_dests_list.append(joined_dests)
            joined_mins_list.append(joined_mins)
            joined_types_list.append(joined_types)
            joined_trains_list.append(BeautifulSoup(soup_text, 'html.parser'))
            joined_hours.append(hours[index])
        elif i_mod in hours_int:
            index = hours_int.index(i_mod)
            joined_dests_list.append(dests_list[index])
            joined_mins_list.append(mins_list[index])
            joined_types_list.append(types_list[index])
            joined_trains_list.append(trains_list[index])
            joined_hours.append(hours[index])
        elif i_mod in _hours_int:
            index = _hours_int.index(i_mod)
            joined_dests_list.append(_dests_list[index])
            joined_mins_list.append(_mins_list[index])
            joined_types_list.append(_types_list[index])
            joined_trains_list.append(_trains_list[index])
            joined_hours.append(_hours[index])
        else:
            joined_dests_list.append(list())
            joined_mins_list.append(list())
            joined_types_list.append(list())
            joined_trains_list.append('')
            joined_hours.append(str(i))
    # print(joined_dests_list)
    # print(joined_mins_list)

    return joined_dests_list, joined_mins_list, joined_types_list, joined_trains_list, joined_hours


def main_function(file_name, html_dir, excel_dir, setting_dir):
    with open(file_name, 'r', errors='replace', encoding="utf_8") as file:
        line_list = file.readlines()

    line_count = 0

    def prepare_join_lists(tables, direction, dw):
        # print(f'tables: {tables}, direction: {direction}, dw: {dw}')
        # 終点などで片方向しか時刻表がない場合はスキップする
        if tables == [None]:
            return

        dests_list = []
        mins_list = []
        types_list = []
        trains_list = []
        hours = []

        for table in tables:
            _dests_list, _mins_list, _types_list, _trains_list, _hours = create_time_table(table, dest_setting)
            dests_list, mins_list, types_list, trains_list, hours = \
                join_lists(dests_list, mins_list, types_list, trains_list, hours,
                           _dests_list, _mins_list, _types_list, _trains_list, _hours, min_hour)
        output_excel(dests_list, mins_list, types_list, wb, color_setting, hours, min_hour,
                     direction, dw, symbol_setting, trains_list)

    for line in line_list:
        line_count += 1
        url_string = line.split(',')[0]
        file_name = line.split(',')[1]
        dest_setting = os.path.join(setting_dir, line.split(',')[2])
        color_setting = os.path.join(setting_dir, line.split(',')[3])
        symbol_setting = os.path.join(setting_dir, line.split(',')[4])
        # 何時から始めるか（上りと下りで開始時刻が違う場合など）
        min_hour = line.split(',')[5]
        # 特定の日時の時刻表を取得する場合
        time_date = line.split(',')[6].replace('\n', '')

        print(line_count, '/', len(line_list))
        print('input_url: ' + url_string)

        # urlが+で繋がっている場合は，soupをいい感じに結合する
        input_urls = url_string.split('+')
        table_weekday_ups = []
        table_weekday_downs = []
        table_holiday_ups = []
        table_holiday_downs = []

        # 特定日の時刻表を取得するとき用
        if time_date == '':
            add_weekday_url_string = '?dw=0'
            add_holiday_url_string = '?dw=2'
            add_path_string = ''
        else:
            add_weekday_url_string = '?dt=' + time_date
            add_holiday_url_string = add_weekday_url_string
            add_path_string = 'date_of_' + time_date + '_'

        # すでに実行済みかどうかを確認するために，一旦soupを取得してpathを調べる
        # pathに更新日時が入っているので，soupを取得しないとpathがわからない
        _, updated_date = prepare_soup(input_urls[0] + add_weekday_url_string, html_dir, file_name, 'weekday', time_date)
        excel_path = os.path.join(excel_dir, updated_date + '_' + add_path_string + file_name + '.xlsx')

        # すでに実行済みのものは除外する
        if not os.path.exists(excel_path):
            for input_url in input_urls:
                # urlが/d2になっている場合は上下を逆にする
                reverse_flag = False
                if '/d2' in input_url:
                    reverse_flag = True

                # 平日分
                soup, updated_date = prepare_soup(input_url + add_weekday_url_string, html_dir, file_name, 'weekday', time_date)
                table_up, table_down = get_each_table(soup, reverse_flag)
                table_weekday_ups.append(table_up)
                table_weekday_downs.append(table_down)

                soup, _ = prepare_soup(input_url + add_holiday_url_string, html_dir, file_name, 'holiday', time_date)
                table_up, table_down = get_each_table(soup, reverse_flag)
                table_holiday_ups.append(table_up)
                table_holiday_downs.append(table_down)

            wb = xlsxwriter.Workbook(excel_path)

            prepare_join_lists(table_weekday_ups, 'up', 'weekday')
            prepare_join_lists(table_weekday_downs, 'down', 'weekday')
            prepare_join_lists(table_holiday_ups, 'up', 'holiday')
            prepare_join_lists(table_holiday_downs, 'down', 'holiday')

            wb.close()
        else:
            logging.warning('Already exist: ' + excel_path)


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
