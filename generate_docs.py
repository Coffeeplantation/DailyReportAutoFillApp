import sys, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, '.')
from build_pptx import _build_manual_pptx, _build_report_pptx

os.makedirs('dist_docs', exist_ok=True)

with open('dist_docs/manual.pptx', 'wb') as f:
    f.write(_build_manual_pptx().read())
print('manual.pptx OK')

with open('dist_docs/report.pptx', 'wb') as f:
    f.write(_build_report_pptx().read())
print('report.pptx OK')
