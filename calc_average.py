import pandas as pd
import numpy as np

# قراءة الملف المدمج
input_file = r"c:\Users\sa\Desktop\فحص المعدل\جميع_الطلاب.csv"
output_file = r"c:\Users\sa\Desktop\فحص المعدل\جميع_الطلاب.csv"

df = pd.read_csv(input_file, encoding='utf-8-sig')

# تعريف أعمدة العلامات
grade_columns = [
    'الرياضيات',
    'العلوم الحياتية / تربية وطنية',
    'اللغة الإنجليزية',
    'اللغة العربية',
    'الأحياء',
    'الفيزياء',
    'الكيمياء',
    'الغة العربية',  # عمود عربية خاص بالعاشر
    'التاريخ',
    'الجغرافيا'
]

# التأكد من أن أعمدة العلامات رقمية
for col in grade_columns:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# حساب المعدل لكل طالب بناءً على المواد المتوفرة (غير الفارغة)
def calc_average(row):
    grades = []
    for col in grade_columns:
        if col in df.columns:
            val = row[col]
            if pd.notna(val):
                grades.append(val)
    if grades:
        return round(sum(grades) / len(grades), 2)
    return None

df['المعدل'] = df.apply(calc_average, axis=1)

# حفظ الملف مع المعدل
df.to_csv(output_file, index=False, encoding='utf-8-sig')

print("تم حساب المعدل بنجاح!")
print(f"إجمالي الطلاب: {len(df)}")
print(f"\nتوزيع المعدلات:")
print(f"  90-100 (ممتاز):    {len(df[df['المعدل'] >= 90])} طالب")
print(f"  80-89  (جيد جداً): {len(df[(df['المعدل'] >= 80) & (df['المعدل'] < 90)])} طالب")
print(f"  70-79  (جيد):      {len(df[(df['المعدل'] >= 70) & (df['المعدل'] < 80)])} طالب")
print(f"  60-69  (مقبول):    {len(df[(df['المعدل'] >= 60) & (df['المعدل'] < 70)])} طالب")
print(f"  أقل من 60:         {len(df[df['المعدل'] < 60])} طالب")
print(f"\nأعلى معدل:  {df['المعدل'].max()}")
print(f"أقل معدل:   {df['المعدل'].min()}")
print(f"متوسط المعدلات: {df['المعدل'].mean():.2f}")

print(f"\nعينة من النتائج:")
sample = df[['اسم الطالب', 'الشعبة', 'المعدل']].head(10)
print(sample.to_string(index=False))

print(f"\nالملف: {output_file}")
