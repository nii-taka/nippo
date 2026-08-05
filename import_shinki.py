"""
新規取引先確認Excelファイルからshinki_clientsデータを抽出してindex.htmlに反映する
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r'C:\Users\新居貴弘\Desktop\nippo_git')

import pandas as pd
import datetime
from bcsms_auto_update import _extract_js_var, _replace_js_var, push_to_github, EXCLUDE_PERSONS
from config import GITHUB_TOKEN, GITHUB_REPO

FILES = {
    '本社': r'C:\Users\新居貴弘\Desktop\新規取引先各拠点\新規取引先確認 (56).xlsx',
    '仙台': r'C:\Users\新居貴弘\Desktop\新規取引先各拠点\新規取引先確認 (53).xlsx',
    '東京': r'C:\Users\新居貴弘\Desktop\新規取引先各拠点\新規取引先確認 (54).xlsx',
    '警備': r'C:\Users\新居貴弘\Desktop\新規取引先各拠点\新規取引先確認 (55).xlsx',
}

CUTOFF = datetime.date(2024, 6, 1)  # 2026年6月から2年前

def parse_shinki_excel(path, region):
    df = pd.read_excel(path, sheet_name='地域別', header=None)
    # 行3がヘッダー: 地域,担当者名,契約日,得意先名,戦略予材,...
    # データは行4以降
    data = df.iloc[4:].copy()
    data.columns = range(len(data.columns))

    result = {}  # person -> [{name, 契約日, 戦略}]
    current_person = None

    for _, row in data.iterrows():
        person_cell = str(row[1]).strip() if pd.notna(row[1]) else ''
        name_cell = str(row[3]).strip() if pd.notna(row[3]) else ''
        date_cell = row[2]
        strategy_cell = str(row[4]).strip() if pd.notna(row[4]) else ''

        # 集計行・空行スキップ
        if '集計' in person_cell or '小計' in person_cell:
            continue
        if not name_cell or name_cell == 'nan':
            continue

        # 担当者名の更新（前の行から引き継ぎ）
        if person_cell and person_cell != 'nan':
            current_person = person_cell

        if not current_person:
            continue
        if current_person in EXCLUDE_PERSONS:
            continue

        # 契約日パース
        try:
            if pd.isna(date_cell):
                continue
            if isinstance(date_cell, datetime.datetime):
                cd = date_cell.date()
            else:
                cd = pd.to_datetime(date_cell).date()
        except:
            continue

        # カットオフ以前はスキップ
        if cd < CUTOFF:
            continue

        # 戦略予材
        strategy = '〇' if strategy_cell == '〇' else ''

        if current_person not in result:
            result[current_person] = []

        result[current_person].append({
            'name': name_cell,
            '契約日': cd.strftime('%Y/%m/%d'),
            '戦略': strategy
        })

    # 契約日降順ソート
    for p in result:
        result[p].sort(key=lambda x: x['契約日'], reverse=True)

    print(f'  {region}: {sum(len(v) for v in result.values())}件 / {len(result)}名')
    return result

# 全地域パース
shinki_all = {}
for region, path in FILES.items():
    print(f'[{region}] 読み込み中...')
    shinki_all[region] = parse_shinki_excel(path, region)

# index.html 更新
repo_path = r'C:\Users\新居貴弘\Desktop\nippo_git'
html_path = repo_path + r'\index.html'

with open(html_path, 'r', encoding='utf-8') as f:
    content = f.read()

raw = _extract_js_var(content, 'RAW')
all_regions = _extract_js_var(content, 'ALL_REGIONS')

# RAW（本社）
raw['shinki_clients'] = shinki_all.get('本社', raw.get('shinki_clients', {}))
content = _replace_js_var(content, 'RAW', raw)

# ALL_REGIONS（全地域）
for region in ['本社', '東京', '警備', '仙台']:
    if region in all_regions:
        all_regions[region]['shinki_clients'] = shinki_all.get(region, {})
content = _replace_js_var(content, 'ALL_REGIONS', all_regions)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(content)

print('[OK] index.html 更新完了')

# GitHubプッシュ
import subprocess, os
os.chdir(repo_path)
subprocess.run(['git', 'add', 'index.html'], check=True)
subprocess.run(['git', 'commit', '-m', 'data: 新規取引先確認Excelから手動インポート'], check=True)
subprocess.run(['git', 'push'], check=True)
print('[OK] GitHubプッシュ完了')
