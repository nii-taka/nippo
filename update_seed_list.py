import sys, io, glob, os, json, subprocess, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r'C:\Users\新居貴弘\Desktop\nippo_git')
from config import GITHUB_TOKEN, GITHUB_REPO

repo_path = r'C:\Users\新居貴弘\Desktop\nippo_git'

# ── 最新の種リストJSONを自動検出（Desktop・Downloads から） ──
search_dirs = [
    r'C:\Users\新居貴弘\Desktop',
    r'C:\Users\新居貴弘\Downloads',
]
candidates = []
for d in search_dirs:
    candidates += glob.glob(os.path.join(d, 'strategy_seeds*.json'))
if not candidates:
    print('[ERROR] strategy_seeds*.json が見つかりません（my-strategy.htmlで「エクスポート」してください）')
    sys.exit(1)
src = max(candidates, key=os.path.getmtime)
print(f'[FILE] 使用ファイル: {src}')

with open(src, encoding='utf-8') as f:
    data = json.load(f)
print(f'[DATA] {len(data)}件')

dst = os.path.join(repo_path, 'strategy_seeds.json')
with open(dst, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ── GitHubにプッシュ ──
today = datetime.date.today().strftime('%Y/%m/%d')
os.chdir(repo_path)
remote_url = f'https://nii-taka:{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git'
subprocess.run(['git', 'config', 'user.email', 'nii-taka@users.noreply.github.com'], check=True)
subprocess.run(['git', 'config', 'user.name', 'nii-taka'], check=True)

fetch_result = subprocess.run(['git', 'fetch', remote_url, 'main:refs/remotes/origin/main'],
                              capture_output=True, text=True)
if fetch_result.returncode != 0:
    print("[ERROR] fetch失敗:", fetch_result.stderr)

log = subprocess.run(['git', 'log', '--oneline', 'origin/main..HEAD'], capture_output=True, text=True)
if log.stdout.strip():
    subprocess.run(['git', 'reset', '--soft', 'origin/main'], check=True)

subprocess.run(['git', 'add', 'strategy_seeds.json'], check=True)

diff = subprocess.run(['git', 'diff', '--cached', '--quiet'], capture_output=True)
if diff.returncode != 0:
    subprocess.run(['git', 'commit', '-m', f'data: {today} 種リスト更新'], check=True)
else:
    print("[OK] 変更なし（スキップ）")
    sys.exit(0)

result = subprocess.run(['git', 'push', remote_url, 'main'], capture_output=True, text=True)
if result.returncode == 0:
    print("[OK] GitHubプッシュ完了")
else:
    print(f"[ERROR] プッシュ失敗: {result.stderr}")
