"""Keep only what the sector-target golden test needs from run_workbook.py output."""
import json, sys
data = json.load(open(sys.argv[1]))
out = [{'scenario': d['scenario'],
        'workbook': {'sectors': {k: {'G': v['G']} for k, v in d['workbook']['sectors'].items()},
                     'portfolio': {'finalPortfolioTotal': d['workbook']['portfolio']['finalPortfolioTotal']}}}
       for d in data]
json.dump(out, open(sys.argv[2], 'w'), indent=1)
print(f'{len(out)} fixtures -> {sys.argv[2]}')
