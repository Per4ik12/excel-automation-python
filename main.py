import pandas as pd
import matplotlib.pyplot as plt

# Входной файл
input_file = 'input_sales.xlsx'
output_file = 'report_sales.xlsx'
chart_file = 'top5_chart.png'

# Чтение данных + обработка отсутствия файла (НОВОЕ: try/except с генерацией тестового)
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"Файл {input_file} не найден. Создаю тестовый пример...")
    # Тестовые данные (можно изменить)
    data = {
        'Товар': ['A', 'B', 'C', 'D', 'E', 'F', 'G'],
        'Количество': [10, 5, 8, 2, 7, 15, 3],
        'Цена': [100, 200, 150, 300, 50, 80, 400]
    }
    df = pd.DataFrame(data)
    df.to_excel(input_file, index=False)
    print(f"Тестовый файл создан: {input_file}")

# Нормализация колонок (ПОЛНАЯ ВЕРСИЯ, без обрезки)
columns = df.columns.str.strip().str.lower()  # нормализуем
rename_map = {}
if 'товар' not in columns:
    for col in columns:
        if col.startswith('товар') or 'товар' in col:
            rename_map[df.columns[columns.get_loc(col)]] = 'Товар'
if 'количество' not in columns:
    for col in columns:
        if col.startswith('оличество') or col.startswith('количество') or 'количество' in col:
            rename_map[df.columns[columns.get_loc(col)]] = 'Количество'
if 'цена' not in columns:
    for col in columns:
        if col.startswith('ена') or col.startswith('цена') or 'цена' in col:
            rename_map[df.columns[columns.get_loc(col)]] = 'Цена'

if rename_map:
    df = df.rename(columns=rename_map)
    print("Колонки нормализованы:", rename_map)

# Расчёты
df['Выручка'] = df['Количество'] * df['Цена']
total_revenue = df['Выручка'].sum()
top_5 = df.sort_values('Выручка', ascending=False).head(5)

print(f"Общая выручка: {total_revenue:,} руб.")
print("Топ-5 товаров по выручке:")
print(top_5[['Товар', 'Выручка']])

# График топ-5 (НОВОЕ: числа над столбцами)
plt.figure(figsize=(10, 6))
bars = plt.bar(top_5['Товар'], top_5['Выручка'], color='#4CAF50')
plt.title('Топ-5 товаров по выручке')
plt.xlabel('Товар')
plt.ylabel('Выручка (руб.)')
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Добавляем числа над столбцами (НОВОЕ)
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height + 20,
             f'{int(height)}', ha='center', va='bottom')

plt.tight_layout()
plt.savefig(chart_file)
plt.close()

# Сохранение отчёта
with pd.ExcelWriter(output_file) as writer:
    df.to_excel(writer, sheet_name='Все данные', index=False)
    top_5.to_excel(writer, sheet_name='Топ-5', index=False)

print(f"Отчёт готов: {output_file} + график {chart_file}")