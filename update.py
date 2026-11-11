# -*- coding: utf-8 -*-
"""台指期全日盤日線自動更新：從期交所下載 TX 近月資料，補進 data.js
全日盤 = 夜盤(前一晚 15:00) + 日盤；開=夜盤開、高低=兩盤合計、收=結算價、量=兩盤合計
用法:  python update.py            # 補到今天
       python update.py --init     # 第一次：從 Excel 匯入歷史
"""
import csv, datetime, json, os, sys, urllib.parse, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(HERE, 'data.json')
DATA_JS = os.path.join(HERE, 'data.js')
XLSX = r'E:\震幅統計2025 (version 1).xlsb.xlsx'
LOG = os.path.join(HERE, 'update.log')

def log(msg):
    line = f'{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {msg}'
    print(line)
    with open(LOG, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

def num(x):
    try:
        return int(round(float(x)))
    except Exception:
        return None

def load_excel():
    import openpyxl
    ws = openpyxl.load_workbook(XLSX, data_only=True).worksheets[0]
    out = []
    for r in ws.iter_rows(min_row=2, max_col=6, values_only=True):
        a, o, h, l, c, v = r
        if a is None or None in (o, h, l, c):
            continue  # 只取四價齊全的列（今天盤中/未來空白列略過）
        if isinstance(a, datetime.datetime):
            d = a.strftime('%Y-%m-%d')
        else:
            s = str(a).strip()
            d = f'{s[:4]}-{s[5:7]}-{s[7:9]}'
        out.append({'d': d, 'o': int(o), 'h': int(h), 'l': int(l), 'c': int(c), 'v': int(v or 0), 'src': 'xls'})
    return out

def fetch_taifex(start, end):
    """回傳 {日期: (開,高,低,收,量)}，只含日盤已收盤的交易日"""
    rows = []
    d = start
    while d <= end:
        e = min(end, (d.replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1))
        body = urllib.parse.urlencode({'down_type': '1', 'commodity_id': 'TX', 'commodity_id2': '',
                                       'queryStartDate': d.strftime('%Y/%m/%d'),
                                       'queryEndDate': e.strftime('%Y/%m/%d')}).encode()
        req = urllib.request.Request('https://www.taifex.com.tw/cht/3/futDataDown', data=body,
                                     headers={'User-Agent': 'Mozilla/5.0'})
        txt = urllib.request.urlopen(req, timeout=60).read().decode('big5', errors='replace')
        rows += [r for r in csv.reader(txt.splitlines())][1:]
        d = e + datetime.timedelta(days=1)
    by = {}
    for r in rows:
        if len(r) < 18:
            continue
        m = r[2].strip()
        if not m.isdigit():      # 排除週選 / 價差
            continue
        by.setdefault(r[0].strip(), {}).setdefault(m, {})[r[17].strip()] = r
    out = {}
    for date, ms in by.items():
        day = [m for m in sorted(ms) if '一般' in ms[m] and num(ms[m]['一般'][3])]
        if not day:
            continue             # 日盤還沒收（例如今天盤中）
        g = ms[day[0]]['一般']   # 近月 = 當天日盤有成交的最近月份
        n = ms[day[0]].get('盤後')
        parts = [x for x in (n, g) if x and num(x[3])]
        settle = num(g[10])
        if settle is None:
            continue             # 結算價還沒出
        out[date.replace('/', '-')] = (num(parts[0][3]), max(num(x[4]) for x in parts),
                                       min(num(x[5]) for x in parts), settle,
                                       sum(num(x[9]) or 0 for x in parts))
    return out

def save(data):
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    with open(DATA_JS, 'w', encoding='utf-8') as f:
        f.write('window.TX_DATA = ' + json.dumps(data, ensure_ascii=False) + ';\n')

def main():
    if '--init' in sys.argv or not os.path.exists(DATA_JSON):
        data = {'bars': load_excel()}
        log(f'從 Excel 匯入 {len(data["bars"])} 筆，最後 {data["bars"][-1]["d"]}')
    else:
        data = json.load(open(DATA_JSON, encoding='utf-8'))
    bars = data['bars']
    last = datetime.date.fromisoformat(bars[-1]['d'])
    today = datetime.date.today()
    try:
        new = fetch_taifex(last + datetime.timedelta(days=1), today) if last < today else {}
    except Exception as ex:
        log(f'下載失敗：{ex}')
        new = {}
    for d in sorted(new):
        o, h, l, c, v = new[d]
        bars.append({'d': d, 'o': o, 'h': h, 'l': l, 'c': c, 'v': v, 'src': 'taifex'})
    data['updated'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    save(data)
    log(f'新增 {len(new)} 筆，目前最後一筆 {bars[-1]["d"]}' + (f'（{", ".join(sorted(new))}）' if new else ''))

if __name__ == '__main__':
    main()
