from pathlib import Path
import json
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / 'Laporan_Format_Baris_Garis.xlsx'
OUT = ROOT / 'data' / 'ownership.json'

def clean(v):
    if pd.isna(v): return None
    return str(v).strip()

df = pd.read_excel(INPUT, dtype=object)
# The pandas workbook uses paired rows: holder summary followed by broker codes.
result = {}
i = 0
while i < len(df):
    row = df.iloc[i]
    stock = clean(row.get('SAHAM'))
    holder = clean(row.get('PEMEGANG SAHAM'))
    if stock and holder and stock.lower() != 'nan':
        stock = stock.upper()
        porto = clean(row.get('PORTO'))
        change = clean(row.get('CHANGE %'))
        brokers = []
        if i + 1 < len(df):
            nxt = df.iloc[i + 1]
            nxt_stock = clean(nxt.get('SAHAM'))
            nxt_holder = clean(nxt.get('PEMEGANG SAHAM'))
            if not nxt_stock and nxt_holder:
                brokers = [x.strip() for x in nxt_holder.split(',') if x.strip()]
                i += 1
        result.setdefault(stock, []).append({
            'stock': stock, 'holder': holder, 'porto': porto,
            'changeText': change, 'brokerCodes': brokers,
            'source': 'pandas:Laporan_Format_Baris_Garis.xlsx'
        })
    i += 1
payload = {
    'ok': True,
    'engine': 'pandas',
    'generatedFrom': INPUT.name,
    'generatedAt': pd.Timestamp.utcnow().isoformat(),
    'stocks': result,
}
OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
print(f'Wrote {OUT} for {len(result)} stocks')
