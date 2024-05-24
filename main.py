import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pandas as pd

# Настройка доступа к Google Sheets API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'Путь к json-ключу сервисного аккаунта', scope)
client = gspread.authorize(creds)

# Открытие таблицы по её ID
spreadsheet_id = 'ID-таблицы'
sheet = client.open_by_key(spreadsheet_id).sheet1

# Чтение данных из таблицы
data = sheet.get_all_records(expected_headers=["Имя и Фамилия", "Группа", "Баллы", "Дата", "За что", "Комментарий"])
for row in data:
    row['Имя и Фамилия'] = row['Имя и Фамилия'].strip()

# Запрос числа занятий у пользователя
num_classes_week1 = int(input("Введите число занятий за первую контрольную неделю: "))
num_classes_week2 = int(input("Введите число занятий за вторую контрольную неделю: "))

# Добавление столбцов для баллов за контрольные недели и подсчёта пропущенных часов
for row in data:
    date_str = row['Дата']
    date_obj = datetime.strptime(date_str, '%d.%m.%Y')
    target_date = datetime(2024, 3, 18)

    if date_obj <= target_date:
        row['Баллы 1 неделя'] = row['Баллы'] * 2
        row['Неделя'] = 1
    else:
        row['Баллы 2 неделя'] = row['Баллы'] * 2
        row['Неделя'] = 2

    row['Посещения'] = 1 if 'посещение' in row['За что'].lower() else 0

# Преобразование данных в DataFrame библиотеки pandas
df = pd.DataFrame(data)

# Группировка по имени и фамилии студента, группе и суммирование баллов и посещений за каждую неделю
grouped_df = df.groupby(['Имя и Фамилия', 'Группа'], as_index=False).agg({
    'Баллы 1 неделя': 'sum',
    'Баллы 2 неделя': 'sum',
    'Баллы': 'sum',
    'Посещения': 'sum'
})

# Расчёт пропущенных часов
grouped_df_week1 = df[df['Неделя'] == 1].groupby(['Имя и Фамилия', 'Группа'], as_index=False)['Посещения'].sum()
grouped_df_week1['Пропущенные часы за первую контрольную неделю'] = (
            (num_classes_week1 - grouped_df_week1['Посещения']) * 4.85).astype(int)

grouped_df_week2 = df[df['Неделя'] == 2].groupby(['Имя и Фамилия', 'Группа'], as_index=False)['Посещения'].sum()
grouped_df_week2['Пропущенные часы за вторую контрольную неделю'] = (
            (num_classes_week2 - grouped_df_week2['Посещения']) * 4.85).astype(int)

# Объединение данных о пропущенных часах с основным DataFrame
grouped_df = grouped_df.merge(
    grouped_df_week1[['Имя и Фамилия', 'Группа', 'Пропущенные часы за первую контрольную неделю']],
    on=['Имя и Фамилия', 'Группа'], how='left')
grouped_df = grouped_df.merge(
    grouped_df_week2[['Имя и Фамилия', 'Группа', 'Пропущенные часы за вторую контрольную неделю']],
    on=['Имя и Фамилия', 'Группа'], how='left')

# Заполнение пропущенных значений 0 (в случае если студент не пропускал занятий)
grouped_df['Пропущенные часы за первую контрольную неделю'] = grouped_df[
    'Пропущенные часы за первую контрольную неделю'].fillna(0).astype(int)
grouped_df['Пропущенные часы за вторую контрольную неделю'] = grouped_df[
    'Пропущенные часы за вторую контрольную неделю'].fillna(0).astype(int)

# Ограничение баллов за недели до максимума 50
grouped_df['Баллы 1 неделя'] = grouped_df['Баллы 1 неделя'].apply(lambda x: min(x, 50))
grouped_df['Баллы 2 неделя'] = grouped_df['Баллы 2 неделя'].apply(lambda x: min(x, 50))

# Открытие таблицы для обновления
existing_spreadsheet_id = 'ID-таблицы'
existing_sheet = client.open_by_key(existing_spreadsheet_id).sheet1

# Обновление данных в таблице
existing_sheet.clear()  # Очистка существующих данных перед обновлением
existing_sheet.update([grouped_df.columns.tolist()] + grouped_df.values.tolist())
print(grouped_df)
print("Данные успешно обновлены в таблице.")
