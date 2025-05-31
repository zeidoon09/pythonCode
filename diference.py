import pandas as pd
import os

os.system(command='cls')

# 1️⃣ مسیر فایل اکسل
file_path = './140402_ReportLastMovementOver10.xlsx'
#
# # 2️⃣ بارگذاری فایل اکسل
excel_file = pd.ExcelFile(file_path)
#
# # نمایش نام شیت‌ها (برای اطمینان)
print(excel_file.sheet_names)
#
# 3️⃣ بارگذاری داده‌ها از شیت‌های مشخص‌شده
df_1 = excel_file.parse('Report3')  # نام شیت اول
df_2 = excel_file.parse('Dashboard')  # نام شیت دوم

# # 4️⃣ استخراج ستون کلیدی (مثلاً MESC) به صورت یکتا
mesc_1 = df_1['MESC'].astype(str).unique()
mesc_2 = df_2['MESC'].astype(str).unique()
#
# # 5️⃣ پیدا کردن مقادیر فقط موجود در یکی از شیت‌ها
only_in_Report3 = set(mesc_1) - set(mesc_2)
only_in_Dashboard = set(mesc_2) - set(mesc_1)
#
print(f"فقط در Report3 و نه در Dashboard: {len(only_in_Report3)} رکورد")
print(f"فقط در Dashboard و نه در Report3: {len(only_in_Dashboard)} رکورد")

# # 6️⃣ ذخیره نتایج در فایل اکسل خروجی (اختیاری)
df_only_in_Report3 = df_1[df_1['MESC'].astype(str).isin(only_in_Report3)]
df_only_in_Dashboard = df_2[df_2['MESC'].astype(str).isin(only_in_Dashboard)]
#
with pd.ExcelWriter('output_differences.xlsx') as writer:
    df_only_in_Report3.to_excel(writer, sheet_name='OnlyInSheetReport3', index=False)
    df_only_in_Dashboard.to_excel(writer, sheet_name='OnlyInSheetDashboard', index=False)

print("✅ فایل output_differences.xlsx ساخته شد.")