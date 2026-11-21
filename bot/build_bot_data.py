"""從 index.html 的建案資料 P 產生 LINE 查詢機器人用的資料檔。
用法：python bot/build_bot_data.py   （在 repo 根目錄執行；地圖更新後重跑一次）
產出：bot/index.json（案名清單）、bot/p/<key>.json（每案一檔）
"""
import json, hashlib, os, shutil
s = open('index.html', encoding='utf8').read()
i = s.find('const P=[') + len('const P=')
d = 0
for k in range(i, len(s)):
    if s[k] == '[': d += 1
    elif s[k] == ']':
        d -= 1
        if d == 0: break
P = json.loads(s[i:k + 1])
asof = max(r[0] for p in P for r in p['rw'])
KEEP = ['n', 'd', 'kind', 'addr', 'fl', 'ty', 'done', 'hh', 'sell', 'bld', 'rw']
shutil.rmtree('bot/p', ignore_errors=True); os.makedirs('bot/p')
idx = []
for p in P:
    key = hashlib.md5((p['d'] + '|' + p['n']).encode()).hexdigest()[:10]
    o = {x: p.get(x) for x in KEEP}
    json.dump(o, open(f'bot/p/{key}.json', 'w', encoding='utf8'), ensure_ascii=False, separators=(',', ':'))
    idx.append([p['n'], p['d'], p['kind'], key, len(p['rw'])])
json.dump({'asof': asof, 'list': idx}, open('bot/index.json', 'w', encoding='utf8'), ensure_ascii=False, separators=(',', ':'))
print(len(idx), 'projects, asof', asof)
