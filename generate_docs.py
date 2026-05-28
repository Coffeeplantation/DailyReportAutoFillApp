import sys, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, '.')
from app import _build_manual, _build_report

os.makedirs('dist_docs', exist_ok=True)

with open('dist_docs/manual.xlsx', 'wb') as f:
    f.write(_build_manual().read())
print('manual.xlsx OK')

with open('dist_docs/report.xlsx', 'wb') as f:
    f.write(_build_report().read())
print('report.xlsx OK')
