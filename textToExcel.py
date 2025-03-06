import pandas as pd
from datetime import datetime

df = pd.read_excel('1402BLSTMPrediction.xlsx')
dates = df['Date'].values

def revertDates(date_str):
    if pd.isna(date_str):
        return date_str
    # تبدیل به رشته
    date_str = str(date_str)
    # تبدیل رشته به شیء datetime و سپس به فرمت جدید
    date_obj = datetime.strptime(date_str, "%m/%d/%Y")
    return date_obj.strftime("%Y/%m/%d")



revertDates(dates)
# خواندن داده ها از فایل اکسل

# تقسیم متن بر اساس کاما و انتخاب مقدار مورد نظر
df.to_excel('converted_dates_excel.xlsx', index=False)



