import csv, json, os

csv_file = r"c:\Users\sa\Desktop\فحص المعدل\جميع_الطلاب.csv"
with open(csv_file, 'r', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f))

# بيانات مضغوطة - فقط الاسم والشعبة والمعدل (بدون المواد)
students = {}
for r in rows:
    students[r['هوية الطالب']] = [
        r['اسم الطالب'],
        r['الشعبة'],
        round(float(r['المعدل']), 1)
    ]

output = r"c:\Users\sa\Desktop\فحص المعدل\site\data.js"
with open(output, 'w', encoding='utf-8') as f:
    f.write('const S=' + json.dumps(students, ensure_ascii=False, separators=(',', ':')) + ';')

print(f"تم: {len(students)} طالب, حجم: {os.path.getsize(output)/1024:.1f} KB")
