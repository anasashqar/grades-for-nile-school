import pandas as pd
import os
import glob

# مسار ملفات الإكسل
input_dir = r"c:\Users\sa\Desktop\فحص المعدل\ملفات الاكسل\الطلاب"
output_file = r"c:\Users\sa\Desktop\فحص المعدل\جميع_الطلاب.csv"

# جمع كل ملفات xlsx
excel_files = glob.glob(os.path.join(input_dir, "*.xlsx"))
print(f"تم العثور على {len(excel_files)} ملف إكسل")

all_dfs = []

for i, file_path in enumerate(sorted(excel_files)):
    filename = os.path.basename(file_path)
    print(f"  [{i+1}/{len(excel_files)}] قراءة: {filename}")
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        # إضافة عمود اسم الملف المصدر
        df['الملف_المصدر'] = filename.replace('.xlsx', '')
        all_dfs.append(df)
        print(f"    -> {len(df)} صف، {len(df.columns)} عمود")
    except Exception as e:
        print(f"    خطأ: {e}")

if all_dfs:
    # دمج كل الملفات
    merged = pd.concat(all_dfs, ignore_index=True)
    # حفظ كملف CSV مع ترميز UTF-8 BOM للتوافق مع Excel
    merged.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"\nتم الدمج بنجاح!")
    print(f"  إجمالي الصفوف: {len(merged)}")
    print(f"  الأعمدة: {list(merged.columns)}")
    print(f"  الملف: {output_file}")
    print(f"\nأول 5 صفوف:")
    print(merged.head().to_string())
else:
    print("لم يتم قراءة أي ملف!")
