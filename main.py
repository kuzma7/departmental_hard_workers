import pandas as pd

# Параметры файлов
source_file_1 = 'Состав кафедры.xlsx'
source_file_2 = 'Проект нагрузки.xlsx'
target_file = 'План по нагрузке.xlsx'
header_file = 'Нагруз без лишних данных.xlsx'


# Словарь с нагрузкой для каждой должности
load_by_position = {
    'Должность': 'Допустимая нагрузка',
    'Заведующий кафедрой': 800,
    'Профессор': 800,
    'Доцент': 850,
    'Старший преподаватель': 900,
    'Ассистент': 900
}

try:
    df_header = pd.read_excel(header_file, header=None)  # Пропускаем 2 строки, читаем 4
    header = df_header.values.tolist()  # Получаем строки с 3 по 6 как список строк
except Exception as e:
    print(f"Ошибка при обработке файла '{header_file}': {e}")
    header = []

# Чтение данных из первого файла ("Состав кафедры")
try:
    df_orders = pd.read_excel(source_file_1, header=None, usecols=[1, 3])  # Чтение колонок: ФИО и должность
    df_orders = df_orders.dropna()  # Удаляем строки с NaN
    df_orders = df_orders.drop_duplicates()  # Удаляем дубликаты
    df_orders = df_orders[~((df_orders[1] == 'Федосенко Юрий Семенович') & (df_orders[3] == 'Профессор'))]
    df_orders = df_orders[~(df_orders[1] == 'Вакантно')]


    # Добавляем новый столбец с нагрузкой
    df_orders = df_orders.copy()  # Избегаем SettingWithCopyWarning
    df_orders['Нагрузка'] = df_orders[3].map(load_by_position)

    # Преобразуем данные первого файла в список строк
    rows_orders = df_orders.values.tolist()
except Exception as e:
    print(f"Ошибка при обработке файла '{source_file_1}': {e}")
    rows_orders = []

# Чтение данных из второго файла ("Проект нагрузки")
try:
    df_nagruzka = pd.read_excel(source_file_2, header=None, skiprows=8)  # Пропускаем первые 8 строк

    # Преобразуем данные второго файла в список строк
    rows_nagruzka = df_nagruzka.values.tolist()
except Exception as e:
    print(f"Ошибка при обработке файла '{source_file_2}': {e}")
    rows_nagruzka = []



combined_rows = []
if header:
    combined_rows.append(header[0])  # Добавляем первую строку заголовков (если она есть)

# Добавляем строки из "Проект нагрузки" и "Состав кафедры"
for nagruzka_row in rows_nagruzka:
    combined_rows.append(nagruzka_row)  # Добавляем строку из "Проект нагрузки"
    combined_rows.extend(rows_orders)  # Добавляем все строки из "Состав кафедры"
    combined_rows.append('\n')

# Создаём DataFrame из комбинированных строк
combined_df = pd.DataFrame(combined_rows)

# Сохраняем данные в итоговый файл
with pd.ExcelWriter(target_file, engine='openpyxl', mode='w') as writer:
    combined_df.to_excel(writer, index=False, header=False)

print(f"Данные успешно записаны в файл '{target_file}'.")