"""Execute existing notebook loading cells, then export one slide locally."""
import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from slide2_export import build_slide2


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--month', default='сентябрь')
    parser.add_argument('--output', type=Path, default=ROOT / 'output/NL_пациенты_2026_слайд2.pptx')
    args = parser.parse_args()
    if Path.cwd() != ROOT:
        parser.error('Запускайте команду из корня проекта.')
    nb = json.loads((ROOT / 'excel_loading.ipynb').read_text(encoding='utf-8'))
    scope = {}
    for cell in nb['cells']:
        if cell['cell_type'] == 'code' and cell.get('metadata', {}).get('papagun_stage') != 'presentation':
            exec(compile(''.join(cell['source']), 'excel_loading.ipynb', 'exec'), scope)
    build_slide2(scope['quantity_df'], scope['quantity_provenance'], month=args.month, output=args.output)
    print('Created one editable slide from quantity_df; source files unchanged.')
