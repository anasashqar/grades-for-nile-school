import pandas as pd
import numpy as np
import os
import glob

print("=" * 70)
print("       فحص شامل ودقيق لصحة البيانات والمعدلات")
print("=" * 70)

input_dir = r"c:\Users\sa\Desktop\فحص المعدل\ملفات الاكسل\الطلاب"
csv_file = r"c:\Users\sa\Desktop\فحص المعدل\جميع_الطلاب.csv"

errors = []
warnings = []

# ============================================================
# الفحص ١: قراءة كل ملف إكسل أصلي والتحقق من عدد الصفوف
# ============================================================
print("\n" + "=" * 70)
print("الفحص ١: مقارنة عدد الصفوف بين ملفات الإكسل وملف CSV")
print("=" * 70)

excel_files = sorted(glob.glob(os.path.join(input_dir, "*.xlsx")))
csv_df = pd.read_csv(csv_file, encoding='utf-8-sig')

total_excel_rows = 0
for file_path in excel_files:
    filename = os.path.basename(file_path).replace('.xlsx', '')
    excel_df = pd.read_excel(file_path, engine='openpyxl')
    excel_count = len(excel_df)
    csv_count = len(csv_df[csv_df['الملف_المصدر'] == filename])
    total_excel_rows += excel_count
    
    status = "✓" if excel_count == csv_count else "✗ خطأ!"
    if excel_count != csv_count:
        errors.append(f"عدد الصفوف مختلف في {filename}: Excel={excel_count}, CSV={csv_count}")
    print(f"  {status} {filename}: Excel={excel_count}, CSV={csv_count}")

print(f"\n  المجموع: Excel={total_excel_rows}, CSV={len(csv_df)} (بدون عمود المعدل)")
if total_excel_rows != len(csv_df):
    errors.append(f"المجموع الكلي مختلف: Excel={total_excel_rows}, CSV={len(csv_df)}")
    print("  ✗ المجموع الكلي غير متطابق!")
else:
    print("  ✓ المجموع الكلي متطابق")

# ============================================================
# الفحص ٢: التحقق من أن كل قيمة في كل ملف إكسل موجودة بشكل صحيح في CSV
# ============================================================
print("\n" + "=" * 70)
print("الفحص ٢: مقارنة القيم فردياً بين كل إكسل وCSV")
print("=" * 70)

value_mismatches = 0
files_checked = 0

for file_path in excel_files:
    filename = os.path.basename(file_path).replace('.xlsx', '')
    excel_df = pd.read_excel(file_path, engine='openpyxl')
    csv_subset = csv_df[csv_df['الملف_المصدر'] == filename].copy()
    
    # مقارنة بناءً على هوية الطالب
    file_mismatches = 0
    
    for _, excel_row in excel_df.iterrows():
        student_id = excel_row['هوية الطالب']
        csv_match = csv_subset[csv_subset['هوية الطالب'] == student_id]
        
        if len(csv_match) == 0:
            errors.append(f"طالب مفقود في CSV: {student_id} من {filename}")
            file_mismatches += 1
            continue
        
        if len(csv_match) > 1:
            warnings.append(f"طالب مكرر في CSV: {student_id} من {filename} ({len(csv_match)} مرات)")
        
        csv_row = csv_match.iloc[0]
        
        # مقارنة كل عمود مشترك
        for col in excel_df.columns:
            if col not in csv_df.columns:
                continue
            
            excel_val = excel_row[col]
            csv_val = csv_row[col]
            
            # كلاهما فارغ - OK
            if pd.isna(excel_val) and pd.isna(csv_val):
                continue
            
            # أحدهما فارغ والآخر لا
            if pd.isna(excel_val) != pd.isna(csv_val):
                errors.append(f"قيمة مختلفة للطالب {student_id} عمود '{col}': Excel={excel_val}, CSV={csv_val}")
                file_mismatches += 1
                continue
            
            # مقارنة القيم الرقمية
            try:
                if abs(float(excel_val) - float(csv_val)) > 0.01:
                    errors.append(f"قيمة رقمية مختلفة للطالب {student_id} عمود '{col}': Excel={excel_val}, CSV={csv_val}")
                    file_mismatches += 1
            except (ValueError, TypeError):
                # مقارنة نصية
                if str(excel_val).strip() != str(csv_val).strip():
                    errors.append(f"قيمة نصية مختلفة للطالب {student_id} عمود '{col}': Excel='{excel_val}', CSV='{csv_val}'")
                    file_mismatches += 1
    
    value_mismatches += file_mismatches
    status = "✓" if file_mismatches == 0 else f"✗ {file_mismatches} اختلاف"
    print(f"  {status} {filename}")
    files_checked += 1

print(f"\n  تم فحص {files_checked} ملف، إجمالي الاختلافات: {value_mismatches}")

# ============================================================
# الفحص ٣: التحقق من أعمدة كل ملف إكسل
# ============================================================
print("\n" + "=" * 70)
print("الفحص ٣: فحص أعمدة كل ملف إكسل الأصلي")
print("=" * 70)

for file_path in excel_files:
    filename = os.path.basename(file_path).replace('.xlsx', '')
    excel_df = pd.read_excel(file_path, engine='openpyxl')
    cols = list(excel_df.columns)
    print(f"  {filename}: {cols}")

# ============================================================
# الفحص ٤: التحقق من حساب المعدل
# ============================================================
print("\n" + "=" * 70)
print("الفحص ٤: إعادة حساب المعدل والتحقق من صحته")
print("=" * 70)

grade_columns = [
    'الرياضيات',
    'العلوم الحياتية / تربية وطنية',
    'اللغة الإنجليزية',
    'اللغة العربية',
    'الأحياء',
    'الفيزياء',
    'الكيمياء',
    'الغة العربية',
    'التاريخ',
    'الجغرافيا'
]

avg_mismatches = 0
avg_missing = 0

for idx, row in csv_df.iterrows():
    grades = []
    for col in grade_columns:
        if col in csv_df.columns:
            val = row[col]
            if pd.notna(val):
                grades.append(float(val))
    
    if grades:
        expected_avg = round(sum(grades) / len(grades), 2)
    else:
        expected_avg = None
    
    actual_avg = row['المعدل']
    
    if expected_avg is None and pd.isna(actual_avg):
        continue
    
    if expected_avg is None and pd.notna(actual_avg):
        errors.append(f"صف {idx+2}: معدل موجود رغم عدم وجود علامات: {actual_avg}")
        avg_mismatches += 1
        continue
    
    if pd.isna(actual_avg):
        errors.append(f"صف {idx+2}: معدل مفقود رغم وجود {len(grades)} علامات")
        avg_missing += 1
        continue
    
    if abs(float(actual_avg) - expected_avg) > 0.01:
        errors.append(f"صف {idx+2} ({row['اسم الطالب']}): المعدل المحسوب={expected_avg}, المعدل في CSV={actual_avg}")
        avg_mismatches += 1

print(f"  إجمالي الطلاب: {len(csv_df)}")
print(f"  معدلات خاطئة: {avg_mismatches}")
print(f"  معدلات مفقودة: {avg_missing}")
if avg_mismatches == 0 and avg_missing == 0:
    print("  ✓ جميع المعدلات صحيحة")
else:
    print("  ✗ يوجد أخطاء في المعدلات")

# ============================================================
# الفحص ٥: تحقق من المعدل بناءً على نوع الشعبة مباشرة من الإكسل
# ============================================================
print("\n" + "=" * 70)
print("الفحص ٥: تحقق من مطابقة المعدل مع الملف الأصلي (عينة عشوائية)")
print("=" * 70)

# اختيار 5 طلاب عشوائيين من كل نوع ملف للتحقق اليدوي
import random
random.seed(42)

sample_files = random.sample(excel_files, min(5, len(excel_files)))
for file_path in sample_files:
    filename = os.path.basename(file_path).replace('.xlsx', '')
    excel_df = pd.read_excel(file_path, engine='openpyxl')
    
    # اختيار طالب عشوائي
    sample_row = excel_df.iloc[0]
    student_id = sample_row['هوية الطالب']
    student_name = sample_row['اسم الطالب']
    
    # استخراج العلامات من الإكسل
    excel_grades = {}
    for col in excel_df.columns:
        if col not in ['النقطة التعليمية', 'الشعبة', 'هوية الطالب', 'اسم الطالب']:
            val = sample_row[col]
            if pd.notna(val):
                try:
                    excel_grades[col] = float(val)
                except:
                    pass
    
    excel_avg = round(sum(excel_grades.values()) / len(excel_grades.values()), 2) if excel_grades else None
    
    # استخراج من CSV
    csv_match = csv_df[csv_df['هوية الطالب'] == student_id]
    if len(csv_match) > 0:
        csv_avg = csv_match.iloc[0]['المعدل']
        
        print(f"\n  ملف: {filename}")
        print(f"  طالب: {student_name} (ID: {student_id})")
        print(f"  العلامات من الإكسل: {excel_grades}")
        print(f"  المعدل المحسوب من الإكسل: {excel_avg}")
        print(f"  المعدل في CSV: {csv_avg}")
        
        if excel_avg is not None and abs(float(csv_avg) - excel_avg) > 0.01:
            print(f"  ✗ غير متطابق!")
            errors.append(f"معدل غير متطابق للطالب {student_name}")
        else:
            print(f"  ✓ متطابق")

# ============================================================
# الفحص ٦: فحص القيم الشاذة والبيانات المفقودة
# ============================================================
print("\n" + "=" * 70)
print("الفحص ٦: فحص القيم الشاذة والبيانات المفقودة")
print("=" * 70)

# فحص أن جميع العلامات بين 0 و 100
for col in grade_columns:
    if col in csv_df.columns:
        valid_grades = csv_df[col].dropna()
        out_of_range = valid_grades[(valid_grades < 0) | (valid_grades > 100)]
        if len(out_of_range) > 0:
            errors.append(f"عمود '{col}': {len(out_of_range)} قيم خارج النطاق 0-100")
            print(f"  ✗ عمود '{col}': {len(out_of_range)} قيم خارج النطاق")
        else:
            non_empty = len(valid_grades)
            print(f"  ✓ عمود '{col}': {non_empty} قيمة، جميعها بين 0-100")

# فحص أن المعدل بين 0 و 100
out_range_avg = csv_df['المعدل'].dropna()
out_range_avg = out_range_avg[(out_range_avg < 0) | (out_range_avg > 100)]
if len(out_range_avg) > 0:
    errors.append(f"عمود 'المعدل': {len(out_range_avg)} قيم خارج النطاق")
    print(f"  ✗ المعدل: {len(out_range_avg)} قيم خارج النطاق")
else:
    print(f"  ✓ المعدل: جميع القيم بين 0-100")

# فحص هل يوجد طلاب بدون أي علامة
no_grades = csv_df[csv_df['المعدل'].isna()]
if len(no_grades) > 0:
    warnings.append(f"{len(no_grades)} طالب بدون أي علامة")
    print(f"  ⚠ {len(no_grades)} طالب بدون أي علامة")
else:
    print(f"  ✓ جميع الطلاب لديهم علامات")

# فحص هل يوجد هويات مكررة
duplicated_ids = csv_df[csv_df.duplicated(subset=['هوية الطالب'], keep=False)]
if len(duplicated_ids) > 0:
    unique_dup = duplicated_ids['هوية الطالب'].nunique()
    warnings.append(f"{unique_dup} هوية طالب مكررة ({len(duplicated_ids)} صف)")
    print(f"  ⚠ {unique_dup} هوية طالب مكررة:")
    for sid in duplicated_ids['هوية الطالب'].unique():
        matches = csv_df[csv_df['هوية الطالب'] == sid]
        names = matches['اسم الطالب'].tolist()
        sources = matches['الملف_المصدر'].tolist()
        print(f"    ID {sid}: {list(zip(names, sources))}")
else:
    print(f"  ✓ لا توجد هويات مكررة")

# ============================================================
# الفحص ٧: فحص أن عمود "اللغة العربية" و "الغة العربية" لا يتداخلان
# ============================================================
print("\n" + "=" * 70)
print("الفحص ٧: فحص أعمدة اللغة العربية (هناك عمودان!)")
print("=" * 70)

both_arabic = csv_df[csv_df['اللغة العربية'].notna() & csv_df['الغة العربية'].notna()]
if len(both_arabic) > 0:
    print(f"  ✗ تحذير خطير: {len(both_arabic)} طالب لديهم قيم في كلا العمودين!")
    print(f"    هذا يعني أن المعدل قد يحسب اللغة العربية مرتين!")
    errors.append(f"{len(both_arabic)} طالب لديهم 'اللغة العربية' و 'الغة العربية' معاً - احتمال حساب مزدوج!")
    # عرض أمثلة
    for _, row in both_arabic.head(3).iterrows():
        print(f"    الطالب: {row['اسم الطالب']}, اللغة العربية={row['اللغة العربية']}, الغة العربية={row['الغة العربية']}, الملف={row['الملف_المصدر']}")
else:
    print(f"  ✓ لا يوجد تداخل - كل طالب لديه عمود واحد فقط للغة العربية")

# فحص ما إذا كانت القيم في "الغة العربية" هي نفسها "اللغة العربية" لبعض الملفات
print(f"\n  توزيع العمودين حسب الملف المصدر:")
for src in csv_df['الملف_المصدر'].unique():
    subset = csv_df[csv_df['الملف_المصدر'] == src]
    has_arabic1 = subset['اللغة العربية'].notna().sum()
    has_arabic2 = subset['الغة العربية'].notna().sum()
    if has_arabic1 > 0 or has_arabic2 > 0:
        indicator = ""
        if has_arabic1 > 0 and has_arabic2 > 0:
            indicator = " ⚠ كلاهما!"
        print(f"    {src}: 'اللغة العربية'={has_arabic1}, 'الغة العربية'={has_arabic2}{indicator}")

# ============================================================
# التقرير النهائي
# ============================================================
print("\n" + "=" * 70)
print("       التقرير النهائي")
print("=" * 70)

if errors:
    print(f"\n  ❌ أخطاء ({len(errors)}):")
    for i, e in enumerate(errors, 1):
        print(f"    {i}. {e}")
else:
    print("\n  ✅ لا توجد أخطاء")

if warnings:
    print(f"\n  ⚠ تحذيرات ({len(warnings)}):")
    for i, w in enumerate(warnings, 1):
        print(f"    {i}. {w}")
else:
    print("\n  ✅ لا توجد تحذيرات")

print("\n" + "=" * 70)
if not errors:
    print("  النتيجة: ✅ جميع البيانات صحيحة ومتطابقة")
else:
    print(f"  النتيجة: ❌ تم العثور على {len(errors)} خطأ يجب معالجته")
print("=" * 70)
