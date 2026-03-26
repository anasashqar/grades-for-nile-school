"""فحص شامل - يقرأ كل ملف إكسل عبر عملية فرعية بترميز صحيح"""
import csv
import os
import subprocess
import json
import sys
import tempfile

print("=" * 60)
print("   فحص شامل سريع لصحة البيانات والمعدلات")
print("=" * 60)

errors = []
warnings = []

# --- قراءة CSV ---
csv_file = r"c:\Users\sa\Desktop\فحص المعدل\جميع_الطلاب.csv"
with open(csv_file, 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    csv_headers = reader.fieldnames
    csv_rows = list(reader)

print(f"\nأعمدة CSV ({len(csv_headers)}): {csv_headers}")
print(f"عدد صفوف CSV: {len(csv_rows)}")

csv_by_source = {}
for row in csv_rows:
    src = row['الملف_المصدر']
    if src not in csv_by_source:
        csv_by_source[src] = []
    csv_by_source[src].append(row)

input_dir = r"c:\Users\sa\Desktop\فحص المعدل\ملفات الاكسل\الطلاب"

# سكريبت فرعي - يكتب النتيجة في ملف بدلاً من stdout
reader_code = '''
import pandas as pd, json, sys
fp = sys.argv[1]
out = sys.argv[2]
df = pd.read_excel(fp, engine='openpyxl')
result = {
    'columns': list(df.columns),
    'count': len(df),
    'data': df.fillna('__NULL__').astype(str).values.tolist()
}
with open(out, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)
'''

reader_path = os.path.join(os.path.dirname(csv_file), '_reader.py')
with open(reader_path, 'w', encoding='utf-8') as f:
    f.write(reader_code)

xlsx_files = sorted([f for f in os.listdir(input_dir) if f.endswith('.xlsx')])
print(f"عدد ملفات الإكسل: {len(xlsx_files)}")

print("\n" + "=" * 60)
print("الفحص ١: مقارنة الصفوف والقيم مع كل ملف إكسل")
print("=" * 60)

total_excel = 0
all_excel_students = []
tmp_json = os.path.join(os.path.dirname(csv_file), '_tmp_result.json')

for xf in xlsx_files:
    fname = xf.replace('.xlsx', '')
    full_path = os.path.join(input_dir, xf)
    
    # حذف ملف مؤقت قديم
    if os.path.exists(tmp_json):
        os.remove(tmp_json)
    
    proc = subprocess.run(
        [sys.executable, reader_path, full_path, tmp_json],
        capture_output=True, timeout=120
    )
    
    if proc.returncode != 0 or not os.path.exists(tmp_json):
        errors.append(f"فشل قراءة {xf}")
        print(f"  ✗ فشل قراءة {xf}")
        continue
    
    with open(tmp_json, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    excel_cols = data['columns']
    excel_count = data['count']
    excel_rows = data['data']
    
    total_excel += excel_count
    
    csv_subset = csv_by_source.get(fname, [])
    csv_count = len(csv_subset)
    
    count_ok = excel_count == csv_count
    if not count_ok:
        errors.append(f"عدد مختلف: {fname} Excel={excel_count} CSV={csv_count}")
    
    # فهرس CSV بهوية الطالب
    csv_by_id = {r['هوية الطالب']: r for r in csv_subset}
    
    value_errors = 0
    for row_vals in excel_rows:
        row_dict = dict(zip(excel_cols, row_vals))
        sid = row_dict.get('هوية الطالب', '')
        if '.' in sid:
            sid = sid.split('.')[0]
        
        csv_match = csv_by_id.get(sid)
        if not csv_match:
            errors.append(f"مفقود في CSV: ID={sid} من {fname}")
            value_errors += 1
            continue
        
        for col in excel_cols:
            if col in ['النقطة التعليمية', 'الشعبة', 'هوية الطالب', 'اسم الطالب']:
                continue
            
            ev = row_dict[col]
            cv = csv_match.get(col, '')
            
            if ev == '__NULL__' and (cv == '' or cv is None):
                continue
            if ev == '__NULL__' and cv != '':
                errors.append(f"ID={sid} '{col}': Excel=فارغ CSV={cv}")
                value_errors += 1
                continue
            if ev != '__NULL__' and (cv == '' or cv is None):
                errors.append(f"ID={sid} '{col}': Excel={ev} CSV=فارغ")
                value_errors += 1
                continue
            
            try:
                if abs(float(ev) - float(cv)) > 0.01:
                    errors.append(f"ID={sid} '{col}': Excel={ev} CSV={cv}")
                    value_errors += 1
            except (ValueError, TypeError):
                if ev.strip() != cv.strip():
                    errors.append(f"ID={sid} '{col}': Excel='{ev}' CSV='{cv}'")
                    value_errors += 1
        
        all_excel_students.append((sid, fname, row_dict, excel_cols))
    
    status = "✓" if count_ok and value_errors == 0 else "✗"
    extra = f" ({value_errors} اختلاف)" if value_errors > 0 else ""
    print(f"  {status} {fname}: Excel={excel_count} CSV={csv_count} cols={len(excel_cols)}{extra}")

print(f"\n  مجموع: Excel={total_excel} CSV={len(csv_rows)}")
if total_excel == len(csv_rows):
    print("  ✓ المجموع متطابق")
else:
    errors.append(f"مجموع Excel={total_excel} CSV={len(csv_rows)}")

# ============================================================
print("\n" + "=" * 60)
print("الفحص ٢: إعادة حساب المعدل والتحقق")
print("=" * 60)

grade_cols = [
    'الرياضيات', 'العلوم الحياتية / تربية وطنية',
    'اللغة الإنجليزية', 'اللغة العربية',
    'الأحياء', 'الفيزياء', 'الكيمياء',
    'الغة العربية', 'التاريخ', 'الجغرافيا'
]

avg_errors = 0
for row in csv_rows:
    grades = []
    for col in grade_cols:
        val = row.get(col, '')
        if val != '' and val is not None:
            try:
                grades.append(float(val))
            except ValueError:
                pass
    
    expected = round(sum(grades) / len(grades), 2) if grades else None
    actual_str = row.get('المعدل', '')
    actual = float(actual_str) if actual_str else None
    
    if expected is None and actual is None:
        continue
    if (expected is None) != (actual is None):
        errors.append(f"معدل: {row['اسم الطالب']} متوقع={expected} فعلي={actual}")
        avg_errors += 1
        continue
    if abs(expected - actual) > 0.01:
        errors.append(f"معدل خاطئ: {row['اسم الطالب']} متوقع={expected} فعلي={actual}")
        avg_errors += 1

print(f"  فحص {len(csv_rows)} معدل -> أخطاء: {avg_errors}")
if avg_errors == 0:
    print("  ✓ جميع المعدلات صحيحة رياضياً")

# ============================================================
print("\n" + "=" * 60)
print("الفحص ٣: عمودي اللغة العربية")
print("=" * 60)

both = sum(1 for r in csv_rows if r.get('اللغة العربية', '') and r.get('الغة العربية', ''))
if both == 0:
    print("  ✓ لا تداخل بين العمودين")
else:
    print(f"  ✗ {both} طالب لديهم عربية مزدوجة!")
    errors.append(f"{both} طالب بعربية مزدوجة")

for fname in sorted(csv_by_source.keys()):
    sub = csv_by_source[fname]
    a1 = sum(1 for r in sub if r.get('اللغة العربية', ''))
    a2 = sum(1 for r in sub if r.get('الغة العربية', ''))
    tag = " ← 'الغة' فقط" if a1 == 0 and a2 > 0 else ""
    if a1 > 0 and a2 > 0: tag = " ⚠⚠⚠"
    print(f"    {fname}: عربية={a1} غة={a2}{tag}")

# ============================================================
print("\n" + "=" * 60)
print("الفحص ٤: عينات عشوائية (معدل من الإكسل مباشرة)")
print("=" * 60)

import random
random.seed(77)
samples = random.sample(all_excel_students, min(10, len(all_excel_students)))

for sid, fname, rd, hds in samples:
    gcols = [h for h in hds if h not in ['النقطة التعليمية', 'الشعبة', 'هوية الطالب', 'اسم الطالب']]
    gvals = []
    for c in gcols:
        v = rd.get(c, '__NULL__')
        if v != '__NULL__':
            try: gvals.append(float(v))
            except: pass
    e_avg = round(sum(gvals)/len(gvals), 2) if gvals else None
    
    cm = next((r for r in csv_rows if r['هوية الطالب'] == sid and r['الملف_المصدر'] == fname), None)
    c_avg = float(cm['المعدل']) if cm and cm.get('المعدل', '') else None
    
    ok = "✓" if e_avg and c_avg and abs(e_avg - c_avg) < 0.01 else "✗"
    name = rd.get('اسم الطالب', '?')
    print(f"  {ok} {name} | {fname}")
    print(f"     مواد: {gcols} -> علامات: {gvals}")
    print(f"     إكسل={e_avg} | CSV={c_avg}")
    if ok == "✗":
        errors.append(f"عينة: {name} e={e_avg} c={c_avg}")

# ============================================================
print("\n" + "=" * 60)
print("الفحص ٥: قيم شاذة")
print("=" * 60)

for col in grade_cols:
    vals = [float(r[col]) for r in csv_rows if r.get(col, '')]
    if vals:
        out = [v for v in vals if v < 0 or v > 100]
        s = "✓" if not out else "✗"
        print(f"  {s} {col}: {len(vals)} قيمة [{min(vals):.0f}-{max(vals):.0f}]")

ids = [r['هوية الطالب'] for r in csv_rows]
dup_ids = set(x for x in ids if ids.count(x) > 1)
if dup_ids:
    print(f"  ⚠ {len(dup_ids)} هوية مكررة")
    warnings.append(f"{len(dup_ids)} هوية مكررة")
else:
    print(f"  ✓ لا هويات مكررة")

# ============================================================
print("\n" + "=" * 60)
print("       التقرير النهائي")
print("=" * 60)

if errors:
    print(f"\n  ❌ أخطاء ({len(errors)}):")
    for i, e in enumerate(errors[:30], 1):
        print(f"    {i}. {e}")
    if len(errors) > 30:
        print(f"    ... و {len(errors)-30} أخرى")
else:
    print("\n  ✅ لا توجد أخطاء")

if warnings:
    print(f"\n  ⚠ تحذيرات ({len(warnings)}):")
    for i, w in enumerate(warnings, 1):
        print(f"    {i}. {w}")
else:
    print("  ✅ لا توجد تحذيرات")

print("\n" + "=" * 60)
if not errors:
    print("  ✅✅✅ النتيجة: كل شيء صحيح ومتطابق ✅✅✅")
else:
    print(f"  ❌ النتيجة: {len(errors)} خطأ يجب المعالجة")
print("=" * 60)

# تنظيف
for f in [reader_path, tmp_json]:
    if os.path.exists(f):
        os.remove(f)
