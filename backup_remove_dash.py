import shutil
import datetime
import os
import stat
from pathlib import Path
ROOT = Path(__file__).resolve().parent
src = ROOT / 'DASH'
if not src.exists():
    print('DASH not found')
    raise SystemExit(1)
now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
backup = ROOT / f'DASH_backup_{now}'
print(f'Copying {src} -> {backup}')
shutil.copytree(src, backup)
print('Copy complete')

def _on_rm_error(func, path, exc_info):
    # Clear read-only bit and retry
    try:
        os.chmod(path, stat.S_IWRITE)
    except Exception:
        pass
    try:
        func(path)
    except Exception as exc:
        print('Failed to remove', path, exc)

shutil.rmtree(src, onerror=_on_rm_error)
print('Removed original DASH')
print('BACKUP=' + str(backup.name))
print('\nRemaining root entries:')
for p in sorted(os.listdir(ROOT)):
    print(p)
