import gspread
from googleapiclient.discovery import build

# استبدل هذه القيم
SPREADSHEET_ID = "H8mWnUA3VEQhEljSimsaUtMzmBCVj96wnoOJznAlz4"
SHEET_NAME = "Offline_Orders"

# إنشاء خدمة Google Sheets
gc = gspread.service_account()

# فتح ملف Google Sheet
sheet = gc.open_by_id(SPREADSHEET_ID)

# الوصول إلى ورقة العمل المحددة
worksheet = sheet.worksheet(SHEET_NAME)

# جلب جميع القيم كصيغ
cell_values = worksheet.get_all_values(value_render_option='FORMULA')

# تكرار كل سطر ومعالجة البيانات
for row in cell_values:
    # مثال: طباعة كل سطر
    print(row)

