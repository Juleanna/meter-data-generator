#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Виправлена версія генератора без помилок
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import pandas as pd
import random
import calendar

# Спроба імпорту openpyxl
try:
    import openpyxl
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils.dataframe import dataframe_to_rows
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

class SimpleCalendar:
    def __init__(self, parent):
        self.parent = parent
        self.selected_date = datetime.now()
        self.calendar_window = None
        
        # Українські назви
        self.month_names = ["", "Січень", "Лютий", "Березень", "Квітень", "Травень", 
                           "Червень", "Липень", "Серпень", "Вересень", 
                           "Жовтень", "Листопад", "Грудень"]
    
    def open_calendar(self, date_var, button):
        if self.calendar_window:
            self.calendar_window.destroy()
            
        self.calendar_window = tk.Toplevel(self.parent)
        self.calendar_window.title("Календар")
        self.calendar_window.geometry("300x350")
        self.calendar_window.resizable(False, False)
        
        # Позиціонування
        try:
            button_x = button.winfo_rootx()
            button_y = button.winfo_rooty() + button.winfo_height()
            self.calendar_window.geometry(f"300x350+{button_x}+{button_y}")
        except:
            pass
        
        self.current_year = self.selected_date.year
        self.current_month = self.selected_date.month
        self.date_var = date_var
        
        self.create_calendar_content()
    
    def create_calendar_content(self):
        # Заголовок
        header = tk.Frame(self.calendar_window)
        header.pack(pady=10)
        
        tk.Button(header, text="◄", width=3, command=self.prev_month).pack(side=tk.LEFT)
        
        month_label = tk.Label(header, text=f"{self.month_names[self.current_month]} {self.current_year}",
                              font=("Arial", 12, "bold"))
        month_label.pack(side=tk.LEFT, padx=20)
        
        tk.Button(header, text="►", width=3, command=self.next_month).pack(side=tk.LEFT)
        
        # Календар
        cal_frame = tk.Frame(self.calendar_window)
        cal_frame.pack(pady=10)
        
        # Дні тижня
        days = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд"]
        for i, day in enumerate(days):
            tk.Label(cal_frame, text=day, font=("Arial", 9, "bold"), width=4).grid(row=0, column=i)
        
        # Дні місяця
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        for week_num, week in enumerate(cal, 1):
            for day_num, day in enumerate(week):
                if day == 0:
                    tk.Label(cal_frame, text="", width=4).grid(row=week_num, column=day_num)
                else:
                    btn = tk.Button(cal_frame, text=str(day), width=3, height=1,
                                   command=lambda d=day: self.select_date(d))
                    
                    # Виділення поточної дати
                    today = datetime.now()
                    if (day == today.day and 
                        self.current_month == today.month and 
                        self.current_year == today.year):
                        btn.config(bg='lightgreen')
                    
                    # Виділення вибраної дати
                    if (day == self.selected_date.day and 
                        self.current_month == self.selected_date.month and 
                        self.current_year == self.selected_date.year):
                        btn.config(bg='lightblue')
                    
                    btn.grid(row=week_num, column=day_num, padx=1, pady=1)
        
        # Кнопки
        btn_frame = tk.Frame(self.calendar_window)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="Сьогодні", command=self.select_today).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Закрити", command=self.close_calendar).pack(side=tk.LEFT, padx=5)
    
    def prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.update_calendar()
    
    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.update_calendar()
    
    def update_calendar(self):
        for widget in self.calendar_window.winfo_children():
            widget.destroy()
        self.create_calendar_content()
    
    def select_date(self, day):
        self.selected_date = datetime(self.current_year, self.current_month, day)
        self.date_var.set(self.selected_date.strftime("%d.%m.%Y"))
        self.close_calendar()
    
    def select_today(self):
        today = datetime.now()
        self.current_year = today.year
        self.current_month = today.month
        self.selected_date = today
        self.date_var.set(self.selected_date.strftime("%d.%m.%Y"))
        self.close_calendar()
    
    def close_calendar(self):
        if self.calendar_window:
            self.calendar_window.destroy()
            self.calendar_window = None

class MeterDataGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ Генератор даних лічильників")
        self.root.geometry("750x580")
        self.root.configure(bg='#f8f9fa')
        
        self.create_widgets()
    
    def create_widgets(self):
        # Головний контейнер
        main = tk.Frame(self.root, bg='#f8f9fa', padx=10, pady=8)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title = tk.Label(main, text="⚡ Генератор даних лічильників", 
                        font=('Arial', 14, 'bold'), bg='#f8f9fa', fg='#2c3e50')
        title.pack(pady=(0, 8))
        
        # Контейнер форми
        form = tk.Frame(main, bg='white', relief='solid', bd=1, padx=10, pady=8)
        form.pack(fill=tk.BOTH, expand=True)
        
        # РЯД 1: Дата + Час + Номер + Тип
        row1 = tk.Frame(form, bg='white')
        row1.pack(fill=tk.X, pady=(0, 8))
        
        # Дата
        date_frame = tk.Frame(row1, bg='white')
        date_frame.pack(side=tk.LEFT)
        
        tk.Label(date_frame, text="📅 Дата:", font=('Arial', 9, 'bold'), 
                bg='white').pack(anchor=tk.W)
        
        date_input = tk.Frame(date_frame, bg='white')
        date_input.pack()
        
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        date_entry = tk.Entry(date_input, textvariable=self.date_var, state="readonly", 
                             width=9, font=('Arial', 9))
        date_entry.pack(side=tk.LEFT)
        
        self.calendar = SimpleCalendar(form)
        cal_btn = tk.Button(date_input, text="📅", width=2, height=1, font=('Arial', 8),
                           command=lambda: self.calendar.open_calendar(self.date_var, cal_btn))
        cal_btn.pack(side=tk.LEFT, padx=(1, 0))
        
        # Час
        time_frame = tk.Frame(row1, bg='white')
        time_frame.pack(side=tk.LEFT, padx=(15, 0))
        
        tk.Label(time_frame, text="🕐 Час:", font=('Arial', 9, 'bold'), 
                bg='white').pack(anchor=tk.W)
        
        time_input = tk.Frame(time_frame, bg='white')
        time_input.pack()
        
        self.hour_var = tk.StringVar(value="00")
        self.minute_var = tk.StringVar(value="00")
        
        hour_spin = tk.Spinbox(time_input, from_=0, to=23, width=3, textvariable=self.hour_var,
                              font=('Arial', 9))
        hour_spin.pack(side=tk.LEFT)
        
        tk.Label(time_input, text=":", bg='white', font=('Arial', 10)).pack(side=tk.LEFT)
        
        minute_spin = tk.Spinbox(time_input, values=[f"{i:02d}" for i in range(0, 60, 10)], 
                                width=3, textvariable=self.minute_var, font=('Arial', 9))
        minute_spin.pack(side=tk.LEFT)
        
        # Номер лічільника
        num_frame = tk.Frame(row1, bg='white')
        num_frame.pack(side=tk.LEFT, padx=(15, 0))
        
        tk.Label(num_frame, text="🔢 Номер:", font=('Arial', 9, 'bold'), 
                bg='white').pack(anchor=tk.W)
        self.meter_entry = tk.Entry(num_frame, width=8, font=('Arial', 9))
        self.meter_entry.insert(0, "001")
        self.meter_entry.pack()
        
        # Тип лічільника
        type_frame = tk.Frame(row1, bg='white')
        type_frame.pack(side=tk.LEFT, padx=(15, 0))
        
        tk.Label(type_frame, text="⚙️ Тип:", font=('Arial', 9, 'bold'), 
                bg='white').pack(anchor=tk.W)
        self.meter_type = tk.StringVar(value="1-фазний")
        type_combo = ttk.Combobox(type_frame, textvariable=self.meter_type, 
                                 values=["1-фазний", "3-фазний"], state="readonly", 
                                 width=9, font=('Arial', 9))
        type_combo.pack()
        
        # РЯД 2: Напруга
        row2 = tk.Frame(form, bg='white')
        row2.pack(fill=tk.X, pady=(0, 8))
        
        # Мін напруга
        min_frame = tk.Frame(row2, bg='white')
        min_frame.pack(side=tk.LEFT)
        
        tk.Label(min_frame, text="⚡ Мін. (В):", font=('Arial', 9, 'bold'), 
                bg='white').pack(anchor=tk.W)
        self.min_volt = tk.Entry(min_frame, width=10, font=('Arial', 9))
        self.min_volt.insert(0, "220.00")
        self.min_volt.pack()
        
        # Макс напруга
        max_frame = tk.Frame(row2, bg='white')
        max_frame.pack(side=tk.LEFT, padx=(20, 0))
        
        tk.Label(max_frame, text="⚡ Макс. (В):", font=('Arial', 9, 'bold'), 
                bg='white').pack(anchor=tk.W)
        self.max_volt = tk.Entry(max_frame, width=10, font=('Arial', 9))
        self.max_volt.insert(0, "240.00")
        self.max_volt.pack()
        
        # РЯД 3: Кнопки
        row3 = tk.Frame(form, bg='white')
        row3.pack(fill=tk.X, pady=(0, 8))
        
        self.gen_btn = tk.Button(row3, text="⚡ Генерувати", font=('Arial', 10, 'bold'), 
                                bg='#007bff', fg='white', padx=15, pady=5,
                                command=self.generate_data)
        self.gen_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_btn = tk.Button(row3, text="💾 Зберегти Excel", font=('Arial', 10, 'bold'),
                                 bg='#28a745', fg='white', padx=15, pady=5,
                                 command=self.save_excel)
        self.save_btn.pack(side=tk.LEFT)
        
        # РЯД 4: Прогрес
        row4 = tk.Frame(form, bg='white')
        row4.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(row4, text="📊 Прогрес:", font=('Arial', 8, 'bold'), 
                bg='white').pack(anchor=tk.W)
        
        self.progress = ttk.Progressbar(row4, length=320, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(1, 0))
        
        # РЯД 5: Статус
        row5 = tk.Frame(form, bg='#e9ecef', relief='solid', bd=1, pady=3)
        row5.pack(fill=tk.X)
        
        self.status = tk.Label(row5, text="✅ Готово до роботи", 
                              font=('Arial', 9, 'bold'), fg='#28a745', bg='#e9ecef')
        self.status.pack()
        
        self.data = None
    
    def generate_data(self):
        try:
            # Валідація
            date_str = self.date_var.get()
            start_date = datetime.strptime(date_str, "%d.%m.%Y")
            start_time = datetime.combine(start_date.date(), 
                                        datetime.strptime(f"{self.hour_var.get()}:{self.minute_var.get()}", 
                                                        "%H:%M").time())
            
            meter_num = self.meter_entry.get().strip()
            if not meter_num:
                raise ValueError("Номер лічільника не може бути пустим")
            
            min_volt = float(self.min_volt.get())
            max_volt = float(self.max_volt.get())
            
            if min_volt >= max_volt:
                raise ValueError("Мін. напруга повинна бути менше макс.")
            
            is_3phase = self.meter_type.get() == "3-фазний"
            
            # Генерація
            self.status.config(text="🔄 Генерація даних...", fg='#007bff')
            self.progress['value'] = 0
            self.root.update()
            
            data = []
            current_time = start_time
            
            for i in range(1200):
                row = {
                    'Номер лічільника': meter_num,
                    'Дата': current_time.strftime("%Y-%m-%d"),
                    'Час': current_time.strftime("%H:%M"),
                    'Фаза A': round(random.uniform(min_volt, max_volt), 2)
                }
                
                if is_3phase:
                    row['Фаза B'] = round(random.uniform(min_volt, max_volt), 2)
                    row['Фаза C'] = round(random.uniform(min_volt, max_volt), 2)
                
                data.append(row)
                current_time += timedelta(minutes=10)
                
                # Оновлення прогресу
                if i % 60 == 0:
                    progress_val = (i / 1200) * 100
                    self.progress['value'] = progress_val
                    self.status.config(text=f"🔄 Генерація: {i}/1200 ({progress_val:.0f}%)")
                    self.root.update()
            
            self.data = pd.DataFrame(data)
            self.progress['value'] = 100
            self.status.config(text=f"✅ Згенеровано {len(data)} записів!", fg='#28a745')
            
        except Exception as e:
            messagebox.showerror("Помилка", str(e))
            self.status.config(text="❌ Помилка генерації", fg='#dc3545')
            self.progress['value'] = 0
    
    def save_excel(self):
        if self.data is None:
            messagebox.showwarning("Увага", "Спочатку згенеруйте дані")
            return
        
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Помилка", "Потрібно встановити openpyxl:\npip install openpyxl")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if not file_path:
                return
            
            self.status.config(text="💾 Збереження...", fg='#007bff')
            self.progress['value'] = 20
            self.root.update()
            
            # Збереження Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Збереження даних
                self.data.to_excel(writer, sheet_name='Дані', index=False)
                
                self.progress['value'] = 50
                self.status.config(text="💾 Форматування...")
                self.root.update()
                
                workbook = writer.book
                worksheet = writer.sheets['Дані']
                
                # Автоширина колонок
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    worksheet.column_dimensions[column_letter].width = max_length + 2
                
                self.progress['value'] = 70
                self.status.config(text="📊 Створення діаграм...")
                self.root.update()
                
                # Створення діаграми
                self.create_simple_chart(workbook)
                
                self.progress['value'] = 100
                self.root.update()
            
            self.status.config(text="✅ Файл збережено!", fg='#28a745')
            
            # Запит на відкриття
            if messagebox.askyesno("Успіх", f"Файл збережено:\n{file_path}\n\nВідкрити?"):
                try:
                    import os
                    os.startfile(file_path)
                except:
                    pass
                    
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка збереження:\n{str(e)}")
            self.status.config(text="❌ Помилка збереження", fg='#dc3545')
    
    def create_simple_chart(self, workbook):
        """Створення простої діаграми без проблемних функцій"""
        try:
            # Створення листа діаграми
            chart_sheet = workbook.create_sheet(title="Діаграма")
            
            # Підготовка даних (використовуємо 'h' замість 'H')
            self.data['DateTime'] = pd.to_datetime(self.data['Дата'] + ' ' + self.data['Час'])
            hourly_data = self.data.groupby(self.data['DateTime'].dt.floor('h')).agg({
                'Фаза A': ['min', 'max', 'mean']
            }).round(2)
            
            is_three_phase = 'Фаза B' in self.data.columns
            
            if is_three_phase:
                hourly_data_b = self.data.groupby(self.data['DateTime'].dt.floor('h')).agg({
                    'Фаза B': ['min', 'max', 'mean']
                }).round(2)
                hourly_data_c = self.data.groupby(self.data['DateTime'].dt.floor('h')).agg({
                    'Фаза C': ['min', 'max', 'mean']
                }).round(2)
            
            # Створення таблиці даних
            chart_data = []
            for i, (timestamp, row) in enumerate(hourly_data.iterrows()):
                chart_row = {
                    'Час': timestamp.strftime('%H:%M'),
                    'Фаза A мін': row[('Фаза A', 'min')],
                    'Фаза A макс': row[('Фаза A', 'max')],
                    'Фаза A сер': row[('Фаза A', 'mean')]
                }
                
                if is_three_phase:
                    chart_row.update({
                        'Фаза B мін': hourly_data_b.iloc[i][('Фаза B', 'min')],
                        'Фаза B макс': hourly_data_b.iloc[i][('Фаза B', 'max')],
                        'Фаза B сер': hourly_data_b.iloc[i][('Фаза B', 'mean')],
                        'Фаза C мін': hourly_data_c.iloc[i][('Фаза C', 'min')],
                        'Фаза C макс': hourly_data_c.iloc[i][('Фаза C', 'max')],
                        'Фаза C сер': hourly_data_c.iloc[i][('Фаза C', 'mean')]
                    })
                
                chart_data.append(chart_row)
            
            # Записуємо дані на лист
            chart_df = pd.DataFrame(chart_data)
            for r_idx, row in enumerate(dataframe_to_rows(chart_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    chart_sheet.cell(row=r_idx, column=c_idx, value=value)
            
            # Створення діаграми
            chart = LineChart()
            chart.title = "Аналіз напруги по годинах"
            chart.style = 2
            chart.x_axis.title = 'Проміжок часу (години)'
            chart.y_axis.title = 'Напруга (В)'
            chart.width = 20
            chart.height = 12
            
            # Додавання даних до діаграми
            data_range = Reference(chart_sheet, min_col=2, min_row=1, 
                                  max_col=len(chart_df.columns), max_row=len(chart_df) + 1)
            cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=len(chart_df) + 1)
            
            chart.add_data(data_range, titles_from_data=True)
            chart.set_categories(cats)
            
            # Розміщення діаграми
            chart_sheet.add_chart(chart, "A15")
            
        except Exception as e:
            print(f"Помилка створення діаграми: {e}")
            # Продовжуємо без діаграми

def main():
    root = tk.Tk()
    app = MeterDataGenerator(root)
    
    # Центрування для маленького екрану
    root.update_idletasks()
    w = root.winfo_reqwidth()
    h = root.winfo_reqheight()
    
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    
    x = (screen_w - w) // 2
    y = max(10, (screen_h - h) // 2 - 50)
    
    root.geometry(f"{w}x{h}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()