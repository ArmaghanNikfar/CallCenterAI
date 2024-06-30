import pandas as pd

# خواندن داده ها از فایل اکسل
df = pd.read_excel('excel.xlsx')

# تقسیم متن بر اساس کاما و انتخاب مقدار مورد نظر
df['new2'] = df['A1'].apply(lambda x: x.split(',')[2])


# ذخیره داده ها در فایل اکسل
df.to_excel('split.xlsx', index=False)
