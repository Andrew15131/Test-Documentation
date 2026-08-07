#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_testplan.py

Скрипт генерирует Excel-файл TestPlan_v1.0.xlsx с тест-планом.
Требования:
  pip install openpyxl

Использование:
  python3 scripts/generate_testplan.py

Файл будет сохранён в рабочей директории как TestPlan_v1.0.xlsx
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Создаем книгу
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "TestPlan"

# Заголовки
headers = ["ID", "Раздел", "Тип тестирования", "Описание", "Приоритет", "Ответственный", "Статус", "Спринт"]
ws.append(headers)

# Данные
data = [
    ["TP-001", "Модуль авторизации", "Функциональное", "Проверка входа с валидными учетными данными", "Высокий", "Иванов И.", "Выполнен", "Sprint 1"],
    ["TP-002", "Модуль авторизации", "Функциональное", "Проверка входа с невалидным паролем", "Высокий", "Иванов И.", "Выполнен", "Sprint 1"],
    ["TP-003", "Модуль авторизации", "Функциональное", "Проверка восстановления пароля", "Средний", "Петров П.", "В работе", "Sprint 2"],
    ["TP-004", "Модуль авторизации", "UI/UX", "Проверка адаптивности страницы входа", "Низкий", "Сидорова А.", "Запланирован", "Sprint 3"],
    ["TP-005", "Модуль оплаты", "Функциональное", "Проверка успешной оплаты через карту", "Высокий", "Петров П.", "Выполнен", "Sprint 1"],
    ["TP-006", "Модуль оплаты", "Интеграционное", "Проверка взаимодействия с платежным шлюзом", "Высокий", "Иванов И.", "В работе", "Sprint 2"],
    ["TP-007", "Модуль оплаты", "Функциональное", "Проверка обработки ошибки при недостатке средств", "Средний", "Сидорова А.", "Запланирован", "Sprint 3"],
    ["TP-008", "Модуль оплаты", "Безопасность", "Проверка шифрования данных при транзакции", "Высокий", "Петров П.", "В работе", "Sprint 2"],
    ["TP-009", "Модуль печати", "Функциональное", "Проверка загрузки файла для печати", "Высокий", "Иванов И.", "Выполнен", "Sprint 1"],
    ["TP-010", "Модуль печати", "Функциональное", "Проверка выбора параметров печати", "Средний", "Петров П.", "Запланирован", "Sprint 3"],
    ["TP-011", "Модуль печати", "Производите��ьность", "Проверка времени обработки большого файла", "Низкий", "Сидорова А.", "Запланирован", "Sprint 4"],
    ["TP-012", "Модуль профиля", "Функциональное", "Проверка редактирования личных данных", "Средний", "Иванов И.", "Выполнен", "Sprint 2"],
    ["TP-013", "Модуль профиля", "Функциональное", "Проверка загрузки аватара", "Низкий", "Петров П.", "Запланирован", "Sprint 4"],
    ["TP-014", "Системный", "Регрессионное", "Проверка критических сценариев после релиза", "Высокий", "Сидорова А.", "В работе", "Sprint 2"],
    ["TP-015", "Системный", "Приемочное (UAT)", "Проверка соответствия требованиям заказчика", "Высокий", "Иванов И.", "Запланирован", "Sprint 4"],
]

for row in data:
    ws.append(row)

# ===== СТИЛИЗАЦИЯ =====
# Настройка шрифта и выравнивания для заголовков
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for col in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment

# Автоширина колонок
for col in range(1, len(headers) + 1):
    column_letter = get_column_letter(col)
    ws.column_dimensions[column_letter].width = 22

# Настройка ячеек с данными
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

for row in range(2, len(data) + 2):
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=row, column=col)
        cell.border = thin_border
        cell.alignment = data_alignment
        
        # Цветная заливка для статуса
        if col == 7:
            status = cell.value
            if status == "Выполнен":
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif status == "В работе":
                cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            elif status == "Запланирован":
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# Фиксация первой строки
ws.freeze_panes = "A2"

# Сохранение
wb.save("TestPlan_v1.0.xlsx")
print("✅ Файл TestPlan_v1.0.xlsx успешно создан!")
