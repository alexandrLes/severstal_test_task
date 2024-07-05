import pandas as pd
import random
from datetime import datetime, timedelta
from sqlalchemy import create_engine, text

# Параметры генерации
NUM_ROWS = 10000
regions = ['North', 'South', 'East', 'West']
start_date = datetime(2023, 1, 1)
end_date = datetime(2024, 6, 30)

def random_date(start, end):
    """
    Функция для генерации случайной даты
    """
    return start + timedelta(days=random.randint(0, (end - start).days))

# Генерация данных
data = {
    'Region': [random.choice(regions) for _ in range(NUM_ROWS)],
    'Sale Date': [random_date(start_date, end_date) for _ in range(NUM_ROWS)],
    'Sales Amount': [round(random.uniform(100, 10000), 2) for _ in range(NUM_ROWS)]
}

df = pd.DataFrame(data)
df.to_excel('sales_data.xlsx', index=False)

# Загрузка Excel файла обратно в Python
df_loaded = pd.read_excel('sales_data.xlsx')

# Создание подключения к SQLite
engine = create_engine('sqlite:///sales_data.db')

# Загрузка данных в SQL таблицу
df_loaded.to_sql('sales', con=engine, index=False, if_exists='replace')

# Параметры для фильтрации
filter_date = '2023-12-31'

# SQL-запрос
query = text('''
SELECT 
    Region, 
    SUM([Sales Amount]) as Total_Sales, 
    CASE 
        WHEN SUM([Sales Amount]) > 50000 THEN 'High'
        WHEN SUM([Sales Amount]) BETWEEN 20000 AND 50000 THEN 'Medium'
        ELSE 'Low'
    END as Sales_Category
FROM sales
WHERE [Sale Date] <= :filter_date
GROUP BY Region
ORDER BY Total_Sales DESC
''')

# Выполнение запроса
result_df = pd.read_sql_query(query, con=engine, params={'filter_date': filter_date})

# Сохранение итогового dataframe в новый Excel файл
result_df.to_excel('aggregated_sales_data.xlsx', index=False)
