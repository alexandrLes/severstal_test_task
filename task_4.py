import pandas as pd

# Загрузка данных из файла
file_path = 'task_4_1.xlsx'
df = pd.read_excel(file_path)

# Разделение столбца "Период" на два отдельных столбца
df[['Период начала', 'Период окончания']] = df['Период'].str.split(' - ', expand=True)


# Преобразование строк дат в datetime
df['Период начала'] = pd.to_datetime(df['Период начала'], format='%Y.%m.%d')
df['Период окончания'] = pd.to_datetime(df['Период окончания'], format='%Y.%m.%d')


# Определение периодов
period1_start, period1_end = pd.to_datetime("2023-01-01"), pd.to_datetime("2023-03-01")
period2_start, period2_end = pd.to_datetime("2023-04-01"), pd.to_datetime("2023-06-01")

# Фильтрация данных по периодам
df_period1 = df[(df['Период начала'] >= period1_start) & (df['Период окончания'] <= period1_end)]
df_period2 = df[(df['Период начала'] >= period2_start) & (df['Период окончания'] <= period2_end)]

print(df_period1.head())
print(df_period2.head())

# Объединение данных по продуктам
merged_df = pd.merge(df_period1, df_period2, on='Продукт', suffixes=('_период1', '_период2'), how='outer')

print(merged_df.head())

# Рассчет критичных позиций
def is_critical(row):
    if pd.notnull(row['ВГ_период1']) and pd.notnull(row['ВГ_период2']):
        # Оба периода
        if abs(row['ВГ_период1'] - row['ВГ_период2']) > 5 and row['ВГ_период2'] < 90:
            return 1
        else:
            return 0
    elif pd.notnull(row['ВГ_период1']):
        # Только первый период
        return 1 if row['ВГ_период1'] < 90 else 0
    elif pd.notnull(row['ВГ_период2']):
        # Только второй период
        return 1 if row['ВГ_период2'] < 90 else 0
    else:
        return 0

merged_df['Критичная позиция'] = merged_df.apply(is_critical, axis=1)

# Выбор необходимых столбцов для финального результата
final_columns = ['Продукт', 'Период_период1', 'ВГ_период1', 'Период_период2', 'ВГ_период2', 'Критичная позиция']
final_df = merged_df[final_columns]

# Сохранение обновленного DataFrame в новый файл
output_file_path = 'updated_task_4.xlsx'
final_df.to_excel(output_file_path, index=False)

print(f"Данные успешно сохранены в {output_file_path}")
