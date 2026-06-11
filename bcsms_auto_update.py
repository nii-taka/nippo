"""
NIPPO自動更新スクリプト
BCSMSから作業日報Excelを自動取得 → GitHubにプッシュ → NIPPOに反映

【セットアップ手順】
1. Python 3.8以上をインストール
2. 必要ライブラリをインストール:
   pip install selenium webdriver-manager pandas openpyxl gitpython requests

3. このファイルと同じフォルダに config.py を作成（後述）

4. Windowsタスクスケジューラで毎朝8:00に実行設定
"""

import sys
import os
import datetime

# Windows環境でのUTF-8出力設定
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
import time
import glob
import shutil
import json
import re
import subprocess
import tempfile

# ============================================================
# 設定（config.pyに書いてください）
# ============================================================
try:
    import config
    BCSMS_URL     = config.BCSMS_URL
    BCSMS_ID      = config.BCSMS_ID
    BCSMS_PW      = config.BCSMS_PW
    GITHUB_TOKEN  = config.GITHUB_TOKEN
    GITHUB_REPO   = config.GITHUB_REPO   # "nii-taka/nippo"
    DOWNLOAD_DIR  = config.DOWNLOAD_DIR  # Excelの保存先
except ImportError:
    print("[ERROR] config.py が見つかりません。README を参照してください。")
    sys.exit(1)

# ============================================================
# 地域リスト（本社→東京→警備→仙台の順で取得）
# ============================================================
REGIONS = ['本社', '東京', '警備', '仙台']

REGION_PERSONS = {
    '本社': ['井部 辰悟','新居 貴弘','園田 奎治','山田 勝生','久保木 奎伍','古谷 将紀','小野寺 陽丈','槇戸 康博','今本 椋介'],
    '東京': ['田牧 光','樋口 正寛','目黒 大地','北村 翔太','尾崎 高大','冨賀見 匠太'],
    '警備': ['奥田 宇紀(警備)','淺井 康太郎(警備)','柿本 直也'],
    '仙台': ['髙野 良成(仙台)','芳賀 誉士弥','庄司 陸','山田 一成'],
}

# 反映しない担当者（退職者など）
EXCLUDE_PERSONS = ['小寺 崚太', '小寺 崚太(警備)', '阪岡 直樹', '退)阪岡 直樹']

SHIN_CUTOFF = datetime.date(2024, 6, 1)
JUN_HOLIDAYS_PATTERN = {
    6: [6, 7, 13, 14, 20, 21, 27, 28],  # 6月の土日
}


def get_holidays(year, month):
    """その月の土日リスト"""
    import calendar
    holidays = []
    for day in range(1, calendar.monthrange(year, month)[1] + 1):
        dt = datetime.date(year, month, day)
        if dt.weekday() >= 5:
            holidays.append(dt)
    return holidays


def download_excel_from_bcsms(start_date=None, end_date=None):
    """BCSMSにログインしてExcelをダウンロード"""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    today = datetime.date.today()
    if start_date is None:
        start_date = today.replace(day=1).strftime('%Y/%m/%d')
    if end_date is None:
        end_date = today.strftime('%Y/%m/%d')

    print(f"[DL] BCSMSから取得: {start_date} 〜 {end_date}")

    # ダウンロード先を一時フォルダに設定
    dl_dir = os.path.abspath(DOWNLOAD_DIR)
    os.makedirs(dl_dir, exist_ok=True)

    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')

    # Selenium 4.6+ の selenium-manager で自動的にChromeDriverを管理
    driver = webdriver.Chrome(options=options)

    # ダウンロード先をCDPで設定（experimental_optionより確実）
    driver.execute_cdp_cmd('Page.setDownloadBehavior', {
        'behavior': 'allow',
        'downloadPath': dl_dir,
    })
    wait = WebDriverWait(driver, 20)

    try:
        # ログイン
        driver.get(BCSMS_URL)
        time.sleep(2)
        driver.find_element(By.NAME, 'U001').send_keys(BCSMS_ID)
        driver.find_element(By.NAME, 'U002').send_keys(BCSMS_PW)
        driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="ログイン"]').click()
        time.sleep(3)
        print("[OK] ログイン完了")
        print(f"[DEBUG] ログイン後URL: {driver.current_url}")

        # メニュー：随時出力処理（フレーム対応）
        def find_and_click_text(driver, text, timeout=20):
            """テキストを含む要素をフレームも含めて検索してクリック"""
            from selenium.common.exceptions import TimeoutException as TE
            # メインフレームで試す
            try:
                driver.switch_to.default_content()
                el = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(),'{text}')]"))
                )
                el.click()
                return True
            except TE:
                pass
            # 各フレームで試す
            frames = driver.find_elements(By.TAG_NAME, 'frame') + driver.find_elements(By.TAG_NAME, 'iframe')
            print(f"[DEBUG] フレーム数: {len(frames)}")
            for i, frame in enumerate(frames):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(frame)
                    el = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(),'{text}')]"))
                    )
                    el.click()
                    return True
                except Exception:
                    pass
            return False

        if not find_and_click_text(driver, '随時出力'):
            # ページソースをデバッグ出力（先頭2000文字）
            print(f"[DEBUG] ページソース: {driver.page_source[:2000]}")
            raise Exception("随時出力ボタンが見つかりません")
        time.sleep(2)
        print("[OK] 随時出力処理クリック")

        # 作業日報
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'作業日報') or contains(text(),'作 業 日 報')]"))).click()
        time.sleep(2)
        print("[OK] 作業日報クリック")

        # 出力選択：担当者別
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio']"))).click()
        # 担当者別ラジオボタンを選択
        radios = driver.find_elements(By.XPATH, "//input[@type='radio']")
        for r in radios:
            try:
                label = r.find_element(By.XPATH, "following-sibling::*[1]").text
                if '担当者' in label:
                    r.click()
                    break
            except:
                pass

        # 日付設定
        date_inputs = driver.find_elements(By.XPATH, "//input[@type='text']")
        for inp in date_inputs:
            val = inp.get_attribute('value')
            if '/' in str(val) and len(str(val)) == 10:
                # 開始日
                inp.clear()
                inp.send_keys(start_date)
                break

        # 終了日を今日に設定
        date_inputs2 = driver.find_elements(By.XPATH, "//input[@type='text']")
        found_start = False
        for inp in date_inputs2:
            val = inp.get_attribute('value')
            if '/' in str(val) and len(str(val)) == 10:
                if not found_start:
                    found_start = True
                    inp.clear()
                    inp.send_keys(start_date)
                else:
                    inp.clear()
                    inp.send_keys(end_date)
                    break

        time.sleep(1)

        # Excel出力ボタン（F10）
        excel_btn = driver.find_element(By.XPATH, "//input[contains(@value,'Excel') or contains(@value,'excel')]")
        excel_btn.click()
        print("[OK] Excel出力クリック")

        # ダウンロード完了を待つ（最大60秒）
        time.sleep(5)
        before_time = time.time() - 10  # クリック前の時刻（余裕を持たせる）
        for _ in range(60):
            files = glob.glob(os.path.join(dl_dir, '*.xlsx'))
            # 直近に作成されたファイルを探す
            new_files = [f for f in files if os.path.getmtime(f) > before_time]
            if new_files:
                latest = max(new_files, key=os.path.getmtime)
                print(f"[OK] ダウンロード完了: {latest}")
                return latest
            time.sleep(1)

        print("[ERROR] ダウンロードタイムアウト")
        return None

    finally:
        driver.quit()


def calc_seasonal_gyoshu_all():
    """2024/07-09 と 2025/07-09 の業種別集計をBCSMSからダウンロードして全地域分を返す
    戻り値: {region: {'2024/07': {'解体': X, ...}, ...}, ...}
             + special key '__all__' for 本社（RAW用）
    """
    import pandas as pd
    target_periods = [
        ('2024/07/01', '2024/07/31', '2024/07'),
        ('2024/08/01', '2024/08/31', '2024/08'),
        ('2024/09/01', '2024/09/30', '2024/09'),
        ('2025/07/01', '2025/07/31', '2025/07'),
        ('2025/08/01', '2025/08/31', '2025/08'),
        ('2025/09/01', '2025/09/30', '2025/09'),
    ]
    # 全地域の結果を蓄積
    per_region = {r: {} for r in REGIONS}

    for start, end, key in target_periods:
        print(f"[季節] {key} のダウンロード開始...")
        excel_path = download_excel_from_bcsms(start_date=start, end_date=end)
        if not excel_path:
            print(f"[WARN] {key} ダウンロード失敗、スキップ")
            continue
        try:
            df = pd.read_excel(excel_path, sheet_name='作業日報')
            df.columns = [str(c).strip() for c in df.columns]
            if '担当者名' not in df.columns or '業種' not in df.columns or '人工合計' not in df.columns:
                print(f"[WARN] {key}: 必須列が見つかりません")
                continue
            df = df[~df['担当者名'].isin(EXCLUDE_PERSONS)]
            for region, persons in REGION_PERSONS.items():
                df_r = df[df['担当者名'].isin(persons)]
                gyoshu_totals = {}
                for g, grp in df_r.groupby('業種'):
                    gyoshu_totals[str(g)] = round(float(grp['人工合計'].sum()), 1)
                if gyoshu_totals:
                    per_region[region][key] = gyoshu_totals
            print(f"[OK] {key}: 集計完了")
        except Exception as e:
            print(f"[WARN] {key} 処理エラー: {e}")
    return per_region


def process_excel(excel_path):
    """ExcelからNIPPO用データを集計"""
    import pandas as pd

    df = pd.read_excel(excel_path, sheet_name='作業日報')
    df['伝票日付'] = pd.to_datetime(df['伝票日付'])
    df['契約日'] = pd.to_datetime(df['契約日'], errors='coerce')
    df['date'] = df['伝票日付'].dt.date

    today = datetime.date.today()
    all_dates = sorted(df['date'].unique())
    holidays = get_holidays(today.year, today.month)
    worked_dates = [d for d in all_dates if d not in holidays]
    holiday_dates_str = [d.strftime('%Y/%m/%d') for d in holidays]

    def kbn(row):
        if pd.isna(row['契約日']): return '空欄'
        return '新規' if row['契約日'].date() >= SHIN_CUTOFF else '既存'
    df['kbn'] = df.apply(kbn, axis=1)

    def calc_ps(df_r):
        result = {}
        for name, pd_ in df_r.groupby('担当者名'):
            shin = pd_[pd_['kbn']=='新規']['人工合計'].sum()
            kison = pd_[pd_['kbn']=='既存']['人工合計'].sum()
            null = pd_[pd_['kbn']=='空欄']['人工合計'].sum()
            total = pd_['人工合計'].sum()
            hol = pd_[pd_['date'].isin(holidays)]['人工合計'].sum()
            result[str(name)] = {
                "新規_人工":round(float(shin),1),"新規_売上":round(float(pd_[pd_['kbn']=='新規']['売上合計'].sum()),1),
                "既存_人工":round(float(kison),1),"既存_売上":round(float(pd_[pd_['kbn']=='既存']['売上合計'].sum()),1),
                "空欄_人工":round(float(null),1),"空欄_売上":round(float(pd_[pd_['kbn']=='空欄']['売上合計'].sum()),1),
                "総合_人工":round(float(total),1),"総合_売上":round(float(pd_['売上合計'].sum()),1),
                "月間社数":int(pd_['得意先CD'].nunique()),"月間現場数":int(pd_['現場名'].nunique()),
                "非稼働日_人工":round(float(hol),1)
            }
        return result

    def calc_ranking(ps):
        return {
            "総合":sorted([{"name":n,"value":d["総合_人工"]} for n,d in ps.items()],key=lambda x:-x['value']),
            "新規":sorted([{"name":n,"value":d["新規_人工"]} for n,d in ps.items()],key=lambda x:-x['value']),
            "非稼働日":sorted([{"name":n,"value":d["非稼働日_人工"]} for n,d in ps.items()],key=lambda x:-x['value']),
            "非稼働日一覧":holiday_dates_str
        }

    def calc_daily(df_r, persons):
        result = {}
        for d in all_dates:
            dstr = d.strftime('%Y/%m/%d')
            day_df = df_r[df_r['date']==d]
            result[dstr] = {p: round(float(day_df[day_df['担当者名']==p]['人工合計'].sum()),1) for p in persons}
        return result

    def calc_daily_all(df_r):
        result = {}
        for d in all_dates:
            dstr = d.strftime('%Y/%m/%d')
            day_df = df_r[df_r['date']==d]
            result[dstr] = {'新規':round(float(day_df[day_df['kbn']=='新規']['人工合計'].sum()),1),'既存':round(float(day_df[day_df['kbn']=='既存']['人工合計'].sum()),1),'空欄':round(float(day_df[day_df['kbn']=='空欄']['人工合計'].sum()),1)}
        return result

    def calc_person_daily(df_r):
        result = {}
        for person, pg in df_r.groupby('担当者名'):
            result[str(person)] = {}
            for d in all_dates:
                dstr = d.strftime('%Y/%m/%d')
                day_df = pg[pg['date']==d]
                result[str(person)][dstr] = {'新規':round(float(day_df[day_df['kbn']=='新規']['人工合計'].sum()),1),'既存':round(float(day_df[day_df['kbn']=='既存']['人工合計'].sum()),1),'空欄':round(float(day_df[day_df['kbn']=='空欄']['人工合計'].sum()),1)}
        return result

    def calc_industry(df_r):
        result = {'全体':{}}
        for g, grp in df_r.groupby('業種'):
            result['全体'][g] = {'新規':round(float(grp[grp['kbn']=='新規']['人工合計'].sum()),1),'既存':round(float(grp[grp['kbn']=='既存']['人工合計'].sum()),1),'空欄':round(float(grp[grp['kbn']=='空欄']['人工合計'].sum()),1)}
        for person, pg in df_r.groupby('担当者名'):
            result[str(person)] = {}
            for g, grp in pg.groupby('業種'):
                result[str(person)][g] = {'新規':round(float(grp[grp['kbn']=='新規']['人工合計'].sum()),1),'既存':round(float(grp[grp['kbn']=='既存']['人工合計'].sum()),1),'空欄':round(float(grp[grp['kbn']=='空欄']['人工合計'].sum()),1)}
        return result

    def calc_clients(df_r):
        result = {'全体':[]}
        grp = df_r.groupby(['得意先CD','得意先名'])['人工合計'].sum().reset_index().sort_values('人工合計',ascending=False).head(10)
        ts = df_r['人工合計'].sum()
        result['全体'] = [{'name':str(r['得意先名']),'value':round(float(r['人工合計']),1),'pct':round(float(r['人工合計'])/ts*100,1) if ts>0 else 0} for _,r in grp.iterrows()]
        for person, pg in df_r.groupby('担当者名'):
            pg2 = pg.groupby(['得意先CD','得意先名'])['人工合計'].sum().reset_index().sort_values('人工合計',ascending=False).head(10)
            ps2 = pg['人工合計'].sum()
            result[str(person)] = [{'name':str(r['得意先名']),'value':round(float(r['人工合計']),1),'pct':round(float(r['人工合計'])/ps2*100,1) if ps2>0 else 0} for _,r in pg2.iterrows()]
        return result

    def calc_daily_detail(df_r):
        result = {}
        for person, pg in df_r.groupby('担当者名'):
            result[str(person)] = {}
            for d in all_dates:
                dstr = d.strftime('%Y/%m/%d')
                day_df = pg[pg['date']==d]
                if len(day_df)==0: continue
                tok_list = []
                for (cd,name), cg in day_df.groupby(['得意先CD','得意先名']):
                    v = round(float(cg['人工合計'].sum()),1)
                    k = '新規' if (cg['kbn']=='新規').any() else ('既存' if (cg['kbn']=='既存').any() else '空欄')
                    tok_list.append({'name':str(name),'人工':v,'区分':k})
                tok_list.sort(key=lambda x:-x['人工'])
                result[str(person)][dstr] = {'総合':round(float(day_df['人工合計'].sum()),1),'新規':round(float(day_df[day_df['kbn']=='新規']['人工合計'].sum()),1),'社数':int(day_df['得意先CD'].nunique()),'現場数':int(day_df['現場名'].nunique()),'得意先':tok_list}
        return result

    dates_str = [d.strftime('%Y/%m/%d') for d in all_dates]
    worked_str = [d.strftime('%Y/%m/%d') for d in worked_dates]
    data_range = f"{today.year}年{today.month}月1日 〜 {today.month}月{today.day}日"

    # 2024/07-09の契約 → 2026/07-09に新規ステータス失効
    def calc_shinki_expire(df_all):
        expire_map = {7: '2024/07', 8: '2024/08', 9: '2024/09'}
        result = {}
        for month, expire_ym in expire_map.items():
            mask = (df_all['契約日'].dt.year == 2024) & (df_all['契約日'].dt.month == month) & df_all['契約日'].notna()
            df_m = df_all[mask]
            if len(df_m) == 0:
                continue
            regions_data = {}
            for region, persons in REGION_PERSONS.items():
                df_r = df_m[df_m['担当者名'].isin(persons)]
                if len(df_r) == 0:
                    continue
                persons_data = {}
                for person, pg in df_r.groupby('担当者名'):
                    clients = []
                    seen = set()
                    for (cd, name), cg in pg.groupby(['得意先CD','得意先名']):
                        if name in seen: continue
                        seen.add(name)
                        kt = cg['契約日'].iloc[0]
                        clients.append({'name':str(name),'契約日':kt.strftime('%Y/%m/%d') if pd.notna(kt) else ''})
                    if clients:
                        persons_data[str(person)] = clients
                if persons_data:
                    regions_data[region] = persons_data
            if regions_data:
                result[expire_ym] = regions_data
        return result

    shinki_expire = calc_shinki_expire(df)

    new_data = {}
    for region, persons in REGION_PERSONS.items():
        df_r = df[df['地域']==region]
        ps = calc_ps(df_r)
        ps_ord = {k:ps[k] for k in persons if k in ps}
        new_data[region] = {
            'targets':persons,'dates':dates_str,'worked_dates':worked_str,
            'person_summary':ps_ord,'ranking':calc_ranking(ps_ord),
            'daily':calc_daily(df_r,persons),'daily_all':calc_daily_all(df_r),
            'person_daily':calc_person_daily(df_r),'industry':calc_industry(df_r),
            'clients':calc_clients(df_r),'daily_detail':calc_daily_detail(df_r),
            'region_total':{'総合_人工':round(float(df_r['人工合計'].sum()),1),'新規_人工':round(float(df_r[df_r['kbn']=='新規']['人工合計'].sum()),1),'非稼働日_人工':round(float(df_r[df_r['date'].isin(holidays)]['人工合計'].sum()),1)}
        }
        print(f"  {region}: 総合={new_data[region]['region_total']['総合_人工']}")

    return new_data, data_range, shinki_expire


def _find_js_var_prefix(content, var_name):
    """const VAR_NAME = or const VAR_NAME= の実際のプレフィックスを返す"""
    for candidate in [f'const {var_name} = ', f'const {var_name}=', f'const {var_name} =']:
        for line in content.splitlines():
            if line.startswith(candidate):
                return candidate
    raise ValueError(f'{var_name} が index.html に見つかりません')


def _extract_js_var(content, var_name):
    """const VAR_NAME = {...}; の行からJSONを抽出（行ベース）"""
    prefix = _find_js_var_prefix(content, var_name)
    for line in content.splitlines():
        if line.startswith(prefix):
            json_str = line[len(prefix):]
            if json_str.endswith(';'):
                json_str = json_str[:-1]
            return json.loads(json_str)
    raise ValueError(f'{var_name} が index.html に見つかりません')


def _replace_js_var(content, var_name, new_obj):
    """const VAR_NAME = {...}; の行をまるごと置換（= の後ろにスペースを統一）"""
    prefix = _find_js_var_prefix(content, var_name)
    new_line = f'const {var_name} = ' + json.dumps(new_obj, ensure_ascii=False) + ';'
    lines = content.splitlines(keepends=True)
    for i, line in enumerate(lines):
        if line.startswith(prefix):
            eol = '\r\n' if line.endswith('\r\n') else '\n'
            lines[i] = new_line + eol
            return ''.join(lines)
    raise ValueError(f'{var_name} が index.html に見つかりません')


def merge_expiry(old_expiry, shinki_expire, region):
    """既存のexpiryにshinki_expireをマージ（重複除去）"""
    merged = {ym: dict(v) for ym, v in old_expiry.items()}
    for ym, regions_data in shinki_expire.items():
        if region not in regions_data:
            continue
        if ym not in merged:
            merged[ym] = {}
        existing_persons = merged[ym].get(region, {})
        for person, clients in regions_data[region].items():
            if person not in existing_persons:
                existing_persons[person] = clients
            else:
                existing_names = {c['name'] for c in existing_persons[person]}
                for c in clients:
                    if c['name'] not in existing_names:
                        existing_persons[person].append(c)
        merged[ym][region] = existing_persons
    return merged


def _is_valid_seasonal(data):
    """seasonal_gyoshu が year/month キー形式かチェック"""
    if not data:
        return False
    for k in data.keys():
        if '/' in k and len(k) == 7:  # '2024/07' 形式
            return True
    return False


def _get_valid_seasonal(old_data, new_per_region_data):
    """有効な seasonal_gyoshu を返す（正しい形式ならそのまま、なければ空）"""
    if _is_valid_seasonal(old_data):
        return old_data
    if new_per_region_data is not None:
        return new_per_region_data
    return {}


def update_index_html(new_data, data_range, repo_path, shinki_expire=None):
    """index.htmlのRAWとALL_REGIONSを更新"""
    html_path = os.path.join(repo_path, 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 既存の保持すべきキーを引き継ぐ
    raw_old = _extract_js_var(content, 'RAW')
    all_old = _extract_js_var(content, 'ALL_REGIONS')

    # seasonal_gyoshu が不正な形式（担当者名キー）なら再ダウンロードして計算
    need_seasonal = not _is_valid_seasonal(raw_old.get('seasonal_gyoshu', {}))
    if need_seasonal:
        for region in REGIONS:
            if not _is_valid_seasonal(all_old.get(region, {}).get('seasonal_gyoshu', {})):
                need_seasonal = True
                break
        else:
            need_seasonal = False
    seasonal_per_region = {}
    if need_seasonal:
        print("[INFO] seasonal_gyoshu を再取得します（初回のみ）...")
        seasonal_per_region = calc_seasonal_gyoshu_all()

    honsha = new_data['本社']
    raw_new = {
        'data_range': data_range,
        'targets': honsha['targets'], 'dates': honsha['dates'], 'worked_dates': honsha['worked_dates'],
        'person_summary': honsha['person_summary'], 'ranking': honsha['ranking'],
        'daily': honsha['daily'], 'daily_all': honsha['daily_all'],
        'person_daily': honsha['person_daily'], 'industry': honsha['industry'],
        'clients': honsha['clients'], 'daily_detail': honsha['daily_detail'],
        'region_total': honsha['region_total'],
        'history': raw_old.get('history', []),
        'expiry': merge_expiry({}, shinki_expire or {}, '本社'),
        'seasonal_gyoshu': _get_valid_seasonal(raw_old.get('seasonal_gyoshu', {}), seasonal_per_region.get('本社')),
        'shinki_clients': raw_old.get('shinki_clients', {}),
    }
    content = _replace_js_var(content, 'RAW', raw_new)

    all_new = {}
    for region in ['本社','東京','警備','仙台']:
        all_new[region] = new_data[region]
        all_new[region]['history'] = all_old.get(region,{}).get('history',[])
        all_new[region]['expiry'] = merge_expiry({}, shinki_expire or {}, region)
        all_new[region]['seasonal_gyoshu'] = _get_valid_seasonal(all_old.get(region,{}).get('seasonal_gyoshu',{}), seasonal_per_region.get(region))
        all_new[region]['shinki_clients'] = all_old.get(region,{}).get('shinki_clients',{})
    content = _replace_js_var(content, 'ALL_REGIONS', all_new)

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"[OK] index.html 更新完了")


def push_to_github(repo_path, data_range):
    """GitHubにプッシュ（競合時はリモートを基準にリセットして再適用）"""
    today = datetime.date.today().strftime('%Y/%m/%d')
    os.chdir(repo_path)

    remote_url = f'https://nii-taka:{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git'
    subprocess.run(['git', 'config', 'user.email', 'nii-taka@users.noreply.github.com'], check=True)
    subprocess.run(['git', 'config', 'user.name', 'nii-taka'], check=True)

    # リモートの最新を取得
    fetch_result = subprocess.run(['git', 'fetch', remote_url, 'main:refs/remotes/origin/main'],
                                  capture_output=True, text=True)
    if fetch_result.returncode != 0:
        print("[ERROR] fetch失敗:", fetch_result.stderr)

    # ローカルにコミットがあればリセット（前回のコミットを取り消し）
    log = subprocess.run(['git', 'log', '--oneline', 'origin/main..HEAD'], capture_output=True, text=True)
    if log.stdout.strip():
        subprocess.run(['git', 'reset', '--soft', 'origin/main'], check=True)

    subprocess.run(['git', 'add', 'index.html'], check=True)

    # 変更があればコミット
    diff = subprocess.run(['git', 'diff', '--cached', '--quiet'], capture_output=True)
    if diff.returncode != 0:
        subprocess.run(['git', 'commit', '-m', f'data: {today} 自動更新'], check=True)
    else:
        print("[OK] 変更なし（スキップ）")
        return

    result = subprocess.run(['git', 'push', remote_url, 'main'], capture_output=True, text=True)
    if result.returncode == 0:
        print("[OK] GitHubプッシュ完了")
    else:
        print(f"[ERROR] プッシュ失敗: {result.stderr}")


# ============================================================
# main（LINEワークス通知付き、未設定の場合はスキップ）
# ============================================================
def main():
    import traceback
    try:
        from lineworks_notify import send_lineworks
    except ImportError:
        send_lineworks = None
    start = datetime.datetime.now()
    print(f"[START] NIPPO自動更新開始 {start.strftime('%Y/%m/%d %H:%M:%S')}")

    repo_path = os.path.dirname(os.path.abspath(__file__))

    try:
        # 1. BCSMSからExcelダウンロード（当日分が既にあればスキップ）
        today_str = datetime.date.today().strftime('%Y-%m-%d')
        existing = [f for f in glob.glob(os.path.join(DOWNLOAD_DIR, '*作業日報*.xlsx'))
                    if today_str in f or datetime.date.fromtimestamp(os.path.getmtime(f)) == datetime.date.today()]
        if existing:
            excel_path = max(existing, key=os.path.getmtime)
            print(f"[OK] 当日Excel既存: {os.path.basename(excel_path)}")
        else:
            excel_path = download_excel_from_bcsms()
        if not excel_path:
            raise Exception("Excelのダウンロードに失敗しました（BCSMSエラー、当日ファイルなし）")

        # 2. Excelを集計
        print("[DATA] データ集計中...")
        new_data, data_range, shinki_expire = process_excel(excel_path)

        # 3. index.html更新
        update_index_html(new_data, data_range, repo_path, shinki_expire)

        # 4. GitHubプッシュ（GitHub Actions環境ではワークフロー側で実行）
        if not os.environ.get('GITHUB_ACTIONS'):
            push_to_github(repo_path, data_range)

        # 5. 成功通知
        today = datetime.date.today()
        honsha_total = new_data['本社']['region_total']['総合_人工']
        honsha_shin  = new_data['本社']['region_total']['新規_人工']
        tokyo_total  = new_data['東京']['region_total']['総合_人工']
        worked_days  = len(new_data['本社']['worked_dates'])

        msg = f"""[OK] NIPPO 自動更新完了

 {today.strftime('%Y年%m月%d日')}（稼働{worked_days}日目）

[DATA] 本日時点の実績
━━━━━━━━━━━━
 本社
  総合: {honsha_total:,.1f} 人工
  新規: {honsha_shin:,.1f} 人工

 東京
  総合: {tokyo_total:,.1f} 人工

━━━━━━━━━━━━
▶ https://nii-taka.github.io/nippo/"""

        # 業務時間内（7:00〜20:00）のみ通知
        now_hour = datetime.datetime.now().hour
        if send_lineworks and 7 <= now_hour < 17:
            send_lineworks(msg)
        elif send_lineworks:
            print(f"[INFO] 業務時間外のため通知スキップ ({now_hour}時)")
        print(f"[OK] 完了! {datetime.datetime.now().strftime('%H:%M:%S')}")

    except Exception as e:
        # エラー通知（時間帯問わず送信）
        err_msg = f"""[ERROR] NIPPO 自動更新失敗

{datetime.date.today().strftime('%Y年%m月%d日')} のデータ取得に失敗しました。

エラー内容:
{str(e)}

手動でExcelをアップロードしてください。
https://github.com/nii-taka/nippo"""
        print(f"[ERROR] エラー: {e}")
        traceback.print_exc()
        if send_lineworks:
            send_lineworks(err_msg)
        sys.exit(1)


if __name__ == '__main__':
    main()

