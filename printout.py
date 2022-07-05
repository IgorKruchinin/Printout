# -*- coding: utf-8 -*-
import pandas as pd
import sys
import os.path
import os
#from fpdf import FPDF
#import tempfile
#import win32api
#import win32print
from win32printing import Printer
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
from tkcalendar import DateEntry
#import keyboard
pd.set_option("display.max_columns", None)
pd.set_option("display.max.colwidth", None)
pd.set_option("display.max_rows", 20)

path = "data.xlsx"

if os.path.exists(path):
	df = pd.read_excel("data.xlsx")
	#df.Дата = pd.to_datetime(df.Дата, format='%d.%m.%Y')
	df = df.fillna('	')
else:
	#if input("Вы хотите использовать стороннюю дазу данных (в формате Excel)?, (Д/Н)") == 'Д':
		#df = pd.read_excel(input("Введите путь до файла Excel: "))
	#else:
	df = pd.DataFrame({"Дата":[], "Серия":[], "Номер":[], "Секция":[], "Ремонт":[], "Вид испытаний":[], "Мощность НР кВт":[], "Мощность АР кВт":[], "ВШ1 вкл/выкл А":[], "ВШ2 вкл/выкл А":[], "Наддув атм":[], "Время ч":[], "Расход топлива л":[], "Замечания":[], "Мастер р/и":[]})
stack = []
b'\xe9\x80\x80'.decode('utf-8')
#b'\xd0'.decode('latin-1')
b'\xd1'.decode('latin-1')


#class PDF(FPDF):
#
#	colontitle = None
#
#	def header(self):
#		"""Оформление верхнего контитула каждого листа"""
#		# Настройка шрифта: Sans, bold, размер 15 пунктов
#		self.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
#		self.set_font("DejaVu", '', 15)
#		# Вычисление ширины заголовка 
#		# и установка положения курсора
#		width = self.get_string_width(self.colontitle) + 6
#		self.set_x((210 - width) / 2)
#		# Настройка цветов для рамки, фона и текста
#		self.set_draw_color(255, 255, 255)
#		self.set_fill_color(255, 255, 255)
#		self.set_text_color(0, 0, 0)
#		# Настройка толщины рамки (1 mm)
#		self.set_line_width(1)
#		# вывод текста, переданного в `colontitle`
#		self.cell(width, 9, self.colontitle, 1, 1, "C", True)
#		# Выполнение разрыва строки в 10 мм
#		self.ln(10)
#
#	def footer(self):
#		"""Оформление нижнего контитула каждого листа"""
#		# Устанавливаем курсор на 1,5 см от нижнего края
#		self.set_y(-15)
#		self.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
#		# Настройка шрифта: Sans, italic, 8
#		self.set_font("DejaVu", "", 8)
#		# Установка цвета текста на серый:
#		self.set_text_color(128)
#		# вывод номера страницы
#		#self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")
#
#	def chapter_title(self, label):
#		"""Оформление главы документа"""
#		# Настройка шрифта: Sans 12
#		self.set_font("DejaVuSans", '', 12)
#		# Настройка цвета фона
#		self.set_fill_color(200, 220, 255)
#		# Печать названия главы
#		self.cell(0, 6, f"{label}", 0, 1, "L", True)
#		# Выполнение разрыва строки на 4 мм
#		self.ln(4)
#
#	def chapter_body(self, txt):
#		"""Чтение файла главы и вывод его в PDF-документ"""
#		# Настройка шрифта: Times, размер 12 пунктов
#		self.set_font("DejaVu", '', size=12)
#		# Печать текста:
#		self.multi_cell(0, 5, txt)
#		# Выполнение разрыва строки:
#		self.ln()
#		# надпись 'Конец главы' выделяем курсивом
#		#self.set_font(style="I")
#		#self.cell(0, 5, "(Конец главы)")
#
#	def print_chapter(self, txt):
#		"""Печать одной главы документа"""
#		self.add_page()
#		#self.chapter_title(title)
#		self.chapter_body(txt)
#

def main_window():
	def printer():
		global df
		chapter_font = {
			"height": 32,
		}
		font = {
			"height": 12,
		}
		df = df.fillna("		")
		text = "Дата: {0:10s}	  Серия: {1:10s}	 Номер: {2:7d}	  Секция: {3:3s}	Ремонт: {4:9s}\nВид испытаний: {5:50s}\nМощность НР кВт: {6:5s}	   ВШ1 вкл/выкл А: {7:30s}	  Наддув атм: {8:3s}\nМощность АР кВт: {9:5s}	 ВШ2 вкл/выкл А: {10:30s}	 Время ч: {11:3s}\nРасход топлива л: {12:7s}\nЗамечания:\n{13:750s}\n		 Мастер р/и: {14:25s}".format(date_i.get(), series_i.get(), int(float(num_i.get())), section_i.get(), repair_i.get(), check_i.get(), pnr_i.get(), vsh1_i.get(), nadduv_i.get(), par_i.get(), vsh2_i.get(), time_i.get(), rash_i.get(), notes_i.get(1.0, 'end'), master_i.get())
		with Printer(linegap=1) as printer:
			printer.text("					  Реостатные испытания", font_config=chapter_font)
			printer.text(text, font_config=font)
		#file = tempfile.mktemp(".txt")
		#open(file, 'w').write(text)
		#win32api.ShellExecute(0, "printto", file, '"%s"' % win32print.GetDefaultPrinter(), ".", 0)
	def save():
		global df, path
		#date = Date(date_i.get_date())
		#pd.Timestamp(year=date.year, month=date.month, day=date.day).date()
		dct_1 = {"Дата":pd.Timestamp(date_i.get_date()).date(), "Серия":series_i.get(), "Номер":int(float(num_i.get())), "Секция":section_i.get(), "Ремонт":repair_i.get(), "Вид испытаний":check_i.get(), "Мощность НР кВт":pnr_i.get(), "Мощность АР кВт":par_i.get(), "ВШ1 вкл/выкл А":vsh1_i.get(), "ВШ2 вкл/выкл А":vsh2_i.get(), "Наддув атм":nadduv_i.get(), "Время ч":time_i.get(), "Расход топлива л":rash_i.get(), "Замечания":notes_i.get("1.0", "end"), "Мастер р/и":master_i.get()}
		df = df.append(dct_1, ignore_index=True)
		df.to_excel(path, index=False)
		date_i.delete(0, "end")
		series_i.delete(0, "end")
		num_i.delete(0, "end")
		section_i.delete(0, "end")
		repair_i.delete(0, "end")
		check_i.delete(0, "end")
		par_i.delete(0, "end")
		pnr_i.delete(0, "end")
		vsh1_i.delete(0, "end")
		vsh2_i.delete(0, "end")
		nadduv_i.delete(0, "end")
		time_i.delete(0, "end")
		rash_i.delete(0, "end")
		notes_i.delete(1.0, "end")
		master_i.delete(0, "end")

	def select_from_tab():
		def reset_table():
			global df
			table.delete(*table.get_children())
			#table.column("#0", width=0, stretch='NO')
			#table.heading("#0", text="", anchor='center')
			for colname in list(df.columns.values):
				table.column(colname, width=80, anchor='center')
				table.heading(colname, text=colname, anchor='center')
			for index, row in df.iterrows():
				table.insert(parent='', index='end', iid=index, text='', values=row.to_list())

		def filter_reset_func():
			global df
			df = stack.pop()
			reset_table()

		def filter_date_func():
			global df
			stack.append(df)
			df = df[df.Дата == pd.to_datetime(pd.Timestamp(filter_date_fld_i.get_date()), format='%d.%m.%Y')]
			reset_table()
		
		def filter_series_func():
			global df
			stack.append(df)
			df = df[df.Серия == filter_series_fld_i.get()]
			reset_table()

		def filter_num_func():
			global df
			stack.append(df)
			df = df[df.Номер == int(filter_num_fld_i.get())]
			reset_table()

		def filter_section_func():
			global df
			stack.append(df)
			df = df[df.Секция == filter_section_fld_i.get()]
			reset_table()
		
		def filter_check_func():
			global df
			stack.append(df)
			df = df[df["Вид испытаний"] == filter_check_fld_i.get()]
			reset_table()

		def filter_dates_func():
			global df
			stack.append(df)
			#dd1 = str(filter_dates_fld_i1.get_date()).split('.')
			#dd2 = str(filter_dates_fld_i2.get_date()).split('.')
			d1 = pd.to_datetime(pd.Timestamp(filter_dates_fld_i1.get_date()))
			d2 = pd.to_datetime(pd.Timestamp(filter_dates_fld_i2.get_date()))
			df = df[(d1 <= pd.to_datetime(df.Дата)) & (pd.to_datetime(df.Дата) <= d2)]
			reset_table()

		def select(event):
			for selection in table.selection():
				item = table.item(selection)
				date_i.delete(0, "end")
				series_i.delete(0, "end")
				num_i.delete(0, "end")
				section_i.delete(0, "end")
				repair_i.delete(0, "end")
				check_i.delete(0, "end")
				par_i.delete(0, "end")
				pnr_i.delete(0, "end")
				vsh1_i.delete(0, "end")
				vsh2_i.delete(0, "end")
				nadduv_i.delete(0, "end")
				time_i.delete(0, "end")
				rash_i.delete(0, "end")
				notes_i.delete(1.0, "end")
				master_i.delete(0, "end")

				date_i.insert(0, str(pd.Timestamp(item["values"][0]).strftime('%d.%m.%Y')))
				series_i.insert(0, item["values"][1])
				num_i.insert(0, int(float(item["values"][2])))
				section_i.insert(0, item["values"][3])
				repair_i.insert(0, item["values"][4])
				check_i.insert(0, item["values"][5])
				try:
					par_i.insert(0, int(float(item["values"][6])))
				except ValueError:
					par_i.insert(0, "		 ")
				try:
					pnr_i.insert(0, int(float(item["values"][7])))
				except ValueError:
					pnr_i.insert(0, "		 ")
				vsh1_i.insert(0, item["values"][8])
				vsh2_i.insert(0, item["values"][9])
				nadduv_i.insert(0, item["values"][10])
				try:
					time_i.insert(0, int(float(item["values"][11])))
				except ValueError:
					time_i.insert(0, "		  ")
				try:
					rash_i.insert(0, int(float(item["values"][12])))
				except ValueError:
					rash_i.insert(0, "		  ")
				notes_i.insert(1.0, item["values"][13])
				master_i.insert(0, item["values"][14])
		global df
		select_win = tk.Tk()
		select_win.geometry("2000x750")
		#select_win.attributes("-fullscreen", True)
		select_win.title("Выберите нужный элемент")
		f1 = tk.Frame(select_win)
		f1.pack(side="top")
		filter_date_fld_l = tk.Label(f1, text="По одной дате: ")
		filter_date_fld_l.pack(side="left")
		filter_date_fld_i = DateEntry(f1, locale='ru_RU', date_pattern='dd.MM.yyyy')
		filter_date_fld_i.pack(side="left")
		filter_date_fld_btn = tk.Button(f1, text="Применить", command=filter_date_func)
		filter_date_fld_btn.pack(side="left")
		filter_series_fld_l = tk.Label(f1, text="По серии: ")
		filter_series_fld_l.pack(side="left")
		filter_series_fld_i = tk.Entry(f1, width=10)
		filter_series_fld_i.pack(side="left")
		filter_series_fld_btn = tk.Button(f1, text="Применить", command=filter_series_func)
		filter_series_fld_btn.pack(side="left")
		filter_num_fld_l = tk.Label(f1, text="По номеру: ")
		filter_num_fld_l.pack(side="left")
		filter_num_fld_i = tk.Entry(f1, width=10)
		filter_num_fld_i.pack(side="left")
		filter_num_fld_btn = tk.Button(f1, text="Применить", command=filter_num_func)
		filter_num_fld_btn.pack(side="left")
		filter_dates_fld_l = tk.Label(f1, text="По диапазону дат: ")
		filter_dates_fld_l.pack(side="left")
		filter_dates_fld_i1 = DateEntry(f1, locale='ru_RU', date_pattern='dd.MM.yyyy')
		filter_dates_fld_i1.pack(side="left")
		filter_dates_fld_i2 = DateEntry(f1, locale='ru_RU', date_pattern='dd.MM.yyyy')
		filter_dates_fld_i2.pack(side="left")
		filter_dates_fld_btn = tk.Button(f1, text="Применить", command=filter_dates_func)
		filter_dates_fld_btn.pack(side="left")
		filter_section_fld_l = tk.Label(f1, text="По секции: ")
		filter_section_fld_l.pack(side="left")
		filter_section_fld_i = tk.Entry(f1, width=10)
		filter_section_fld_i.pack(side="left")
		filter_section_fld_btn = tk.Button(f1, text="Применить", command=filter_section_func)
		filter_section_fld_btn.pack(side="left")
		filter_reset_btn = tk.Button(f1, text="Отменить фильтр", command=filter_reset_func)
		filter_reset_btn.pack(side="left")
		frame = tk.Frame(select_win)
		frame.pack(side="top", fill="both", expand=1)
		scrollx = tk.Scrollbar(frame, orient="horizontal")
		scrolly = tk.Scrollbar(frame, orient="vertical")
		table = ttk.Treeview(frame, xscrollcommand=scrollx.set, yscrollcommand=scrolly.set)
		table['columns'] = list(df.columns)
		table.column("#0", width=0, stretch='NO')
		table.heading("#0", text="", anchor='center')
		for colname in list(df.columns.values):
			table.column(colname, width=80, anchor='center')
			table.heading(colname, text=colname, anchor='center')
		for index, row in df.iterrows():
			table.insert(parent='', index='end', iid=index, text='', values=row.to_list())
		table.bind("<<TreeviewSelect>>", select, "+")
		scrollx.config(command=table.xview)
		scrollx.pack(side="bottom", fill="x")
		scrolly.config(command=table.yview)
		scrolly.pack(side="right", fill="y")
		table.pack(fill="both", expand=1)
		

#	def to_pdf():
#		global df
#		pdf = PDF()
#		df = df.fillna('	  ')
#		pdf.colontitle = "Реостатные испытания"
#		try:
#			pdf.print_chapter(txt=f"Дата: {date_i.get()}		 Серия: {series_i.get()}		 Номер: {int(float(num_i.get()))}		 Секция: {section_i.get()}		  Ремонт: {repair_i.get()}\nВид испытаний: {check_i.get()}\nМощность НР кВт: {pnr_i.get()}		  ВШ1 вкл/выкл А: {vsh1_i.get()}		Наддув атм: {nadduv_i.get()}\nМощность АР кВт: {par_i.get()}		ВШ2 вкл/выкл А: {vsh2_i.get()}		  Время ч: {time_i.get()}\nРасход топлива л: {rash_i.get()}\nЗамечания:\n{notes_i.get(1.0, 'end')}\n		Мастер р/и: {master_i.get()}")
#			pdf.output("Распечатка.pdf")
#			print("Готово!")
#			os.system("Распечатка.pdf")
#		except KeyError:
#			print("Нет такого пункта")

	def openfile():
		global df
		path = fd.askopenfilename(title="Открыть файл", initialdir="C:\\", filetypes = (("Excel", "*.xlsx"), ("Excel 98", "*.xls"), ("Любой", "*")))
		df = pd.read_excel(path)
		df.Дата = pd.to_datetime(df.Дата, format='%d.%m.%Y')
		df = df.fillna('	  ')

	def savefile():
		global df
		df.to_excel(fd.asksaveasfilename(title="Сохранить файл", defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"), ("Другой", "*"))), index=False)
		
	def clear():
				date_i.delete(0, "end")
				series_i.delete(0, "end")
				num_i.delete(0, "end")
				section_i.delete(0, "end")
				repair_i.delete(0, "end")
				check_i.delete(0, "end")
				par_i.delete(0, "end")
				pnr_i.delete(0, "end")
				vsh1_i.delete(0, "end")
				vsh2_i.delete(0, "end")
				nadduv_i.delete(0, "end")
				time_i.delete(0, "end")
				rash_i.delete(0, "end")
				notes_i.delete(1.0, "end")
				master_i.delete(0, "end")

	win = tk.Tk()
	win.geometry("750x360")
	win.title("Распечатка")
	fr0 = tk.Frame(win)
	fr0.pack(side="top")
	menu_openfile_btn = tk.Button(fr0, text="Открыть", command=openfile)
	menu_openfile_btn.pack(side="left")
	menu_savefile_btn = tk.Button(fr0, text="Экспорт в Excel", command=savefile)
	menu_savefile_btn.pack(side="left")
	select_from_tab_btn = tk.Button(fr0, text="Подставить из таблицы", command=select_from_tab)
	select_from_tab_btn.pack(side="left")
	lbl = tk.Label(win, text="Реостатные Испытания", font=("Arial Bold", 23))
	lbl.pack(side="top")
	fr1 = tk.Frame(win)
	fr1.pack(side="top")
	date_l = tk.Label(fr1, text="Дата: ")
	date_l.pack(side="left")
	date_i = DateEntry(fr1, locale='ru_RU', date_pattern='dd.MM.yyyy')
	#date_i = tk.Entry(fr1, width=10)
	#date_i.bind("<Button-1>", calendar)
	date_i.pack(side="left")
	#dct_1 = {"Дата": date_i.get()}
	series_l = tk.Label(fr1, text="Серия: ")
	series_l.pack(side="left")
	series_i = tk.Entry(fr1, width=10)
	series_i.pack(side="left")
	num_l = tk.Label(fr1, text="Номер: ")
	num_l.pack(side="left")
	num_i = tk.Entry(fr1, width=10)
	num_i.pack(side="left")
	section_l = tk.Label(fr1, text="Секция: ")
	section_l.pack(side="left")
	section_i = tk.Entry(fr1, width=10)
	section_i.pack(side="left")
	repair_l = tk.Label(fr1, text="Ремонт: ")
	repair_l.pack(side="left")
	repair_i = tk.Entry(fr1, width=10)
	repair_i.pack(side="left")
	fr2 = tk.Frame(win)
	fr2.pack(side="top")
	check_l = tk.Label(fr2, text="Вид испытаний: ")
	check_l.pack(side="left")
	check_i = tk.Entry(fr2, width=77)
	check_i.pack(side="left")
	fr3 = tk.Frame(win)
	fr3.pack(side="top")
	pnr_l = tk.Label(fr3, text="Мощность НР: ")
	pnr_l.pack(side="left")
	pnr_i = tk.Entry(fr3, width=10)
	pnr_i.pack(side="left")
	vsh1_l = tk.Label(fr3, text="ВШ1 вкл/выкл: ")
	vsh1_l.pack(side="left")
	vsh1_i = tk.Entry(fr3, width=10)
	vsh1_i.pack(side="left")
	nadduv_l = tk.Label(fr3, text="Наддув атм: ")
	nadduv_l.pack(side="left")
	nadduv_i = tk.Entry(fr3, width=10)
	nadduv_i.pack(side="left")
	fr4 = tk.Frame(win)
	fr4.pack(side="top")
	par_l = tk.Label(fr4, text="Мощность АР: ")
	par_l.pack(side="left")
	par_i = tk.Entry(fr4, width=10)
	par_i.pack(side="left")
	vsh2_l = tk.Label(fr4, text="ВШ2 вкл/выкл: ")
	vsh2_l.pack(side="left")
	vsh2_i = tk.Entry(fr4, width=10)
	vsh2_i.pack(side="left")
	time_l = tk.Label(fr4, text="Время ч: ")
	time_l.pack(side="left")
	time_i = tk.Entry(fr4, width=10)
	time_i.pack(side="left")
	fr5 = tk.Frame(win)
	fr5.pack(side="top")
	rash_l = tk.Label(fr5, text="Расход топлива л: ")
	rash_l.pack(side="left")
	rash_i = tk.Entry(fr5, width=10)
	rash_i.pack(side="left")
	fr6 = tk.Frame(win)
	fr6.pack(side="top")
	notes_l = tk.Label(fr6, text="Замечания: ")
	notes_l.pack(side="left")
	notes_i = tk.Text(fr6, width=75, height=7)
	notes_i.pack(side="left")
	fr7 = tk.Frame(win)
	fr7.pack(side="top")
	master_l = tk.Label(fr7, text="Мастер: ")
	master_l.pack(side="left")
	master_i = tk.Entry(fr7, width=20)
	master_i.pack(side="left")
	fr8 = tk.Frame(win)
	fr8.pack(side="top")
	clear_btn = tk.Button(fr8, text="Очистить", command=clear)
	clear_btn.pack(side="left")
	save_btn = tk.Button(fr8, text="Сохранить", command=save)
	save_btn.pack(side="left")
	make_rasp_btn = tk.Button(fr8, text="Печать", command=printer)
	make_rasp_btn.pack(side="left")
	win.mainloop()

#def main_cli():
#	def readstring(msg=""):
#		print(msg)
#		ch = sys.stdin.read(1)
#		string = ""
#		exitsym = False
#		#keyboard.add_hotkey('Ctrl + O', exitsym=True)
#		while not exitsym and ch != '!':
#			string += ch
#			ch = sys.stdin.read(1)
#		return string
#	while (1):
#		print("Используйте следующие команды (цифры):")
#		print("0 - Выход (Не забудьте сделать сохранение!)")
#		print("1 - Добавить новый пункт")
#		print("2 - Экспорт в Excel")
#		print("3 - Вывести таблицу на экран (не распечатку!)")
#		print("4 - Установить фильтры")
#		print("5 - Назад (после применения фильтров)")
#		print("6 - Сделать распечатку в PDF")
#		print("7 - Сохранить базу данных")
#		print("8 - Открыть базу данных из файла Excel")
#		print("9 - Удалить элемент")
#		cmd = int(input("Введите команду: "))
#		if cmd == 0:
#			break
#		elif cmd == 1:
#			dct = {"Дата":pd.Timestamp(input("Дата (чч.мм.гггг)")), "Серия":input("Серия: "), "Номер":int(input("Номер: ")), "Секция":input("Секция: "), "Ремонт":input("Ремонт: "), "Вид испытаний":input("Вид испытаний: "), "Мощность НР кВт":input("Мощность НР кВт: "), "Мощность АР кВт":input("Мощность АР кВт: "), "ВШ1 вкл/выкл А":input("ВШ1 вкл/выкл А: "), "ВШ2 вкл/выкл А":input("ВШ2 вкл/выкл А: "), "Наддув атм":input("Наддув атм: "), "Время ч":input("Время ч: "), "Расход топлива л":input("Расход топлива л: ")}
#			dct.update({"Замечания":readstring("Замечания (в конце ввода введите \"!\" и нажмите \"Enter\"): "), "Мастер р/и":input("Мастер р/и: ")})
#			df = df.append(dct, ignore_index=True)
#		elif cmd == 2:
#			filename = input("Введите имя файла для экспорта в Excel: ") + ".xlsx"
#			df.to_excel(filename, index=False)
#		elif cmd == 3:
#			print(df)
#		elif cmd == 4:
#			print("Как вы хотите отфильтровать данные?")
#			print("1 - по серии")
#			print("2 - по дате")
#			print("3 - по виду ремонта")
#			print("4 - по диапазону дат")
#			print("5 - по номеру")
#			print()
#			print("После применения фильтра, вы будете переключены на отфильтрованную таблицу (никаких изменений в исходной таблице на диске не произойдёт, если вы сами не сохраните таблицу (все последующие команды будут отностится к отфилтрованной таблице), чтобы переключится на предыдущую таблицу, введите команду 5 в главном меню")
#			filtercmd = int(input("Введите номер нужного фильтра: "))
#			if filtercmd == 1:
#				stack.append(df)
#				df = df[df.Серия == input("Серия: ")]
#			elif filtercmd == 2:
#				stack.append(df)
#				df = df[df.Дата == pd.Timestamp(input("Дата: "))]
#			elif filtercmd == 3:
#				stack.append(df)
#				df = df[df["Вид ремонта"] == input("Вид ремонта: ")]
#			elif filtercmd == 4:
#				stack.append(df)
#				d1 = pd.Timestamp(input("Дата 1: "))
#				d2 = pd.Timestamp(input("Дата 2: "))
#				#df = pd.DataFrame({"Дата":[], "Серия":[], "Номер":[], "Секция":[], "Ремонт":[], "Вид испытаний":[], "Мощность НР кВт":[], "Мощность АР кВт":[], "ВШ1 вкл/выкл А":[], "ВШ2 вкл/выкл А":[], "Время ч":[], "Расход топлива л":[], "Замечания":[], "Мастер р/и":[]})
#				df = df[(d1 <= df.Дата) & (df["Дата"] <= d2)]
#				#for i in range(len(stack[-1])):
#					#dd, dm, dy = map(int, stack[-1].loc[i]['Дата'].split('.'))
#					#if d1 <= Date(date=dd, month=dm, year=dy) <= d2:
#						#dct = {"Дата":stack[-1].loc[i]['Дата'], "Серия":stack[-1].loc[i]['Серия'], "Номер":stack[-1].loc[i]['Номер'], "Секция":stack[-1].loc[i]['Секция'], "Ремонт":stack[-1].loc[i]['Ремонт'], "Вид испытаний":stack[-1].loc[i]['Вид испытаний'], "Мощность НР кВт":stack[-1].loc[i]['Мощность НР кВт'], "Мощность АР кВт":stack[-1].loc[i]['Мощность АР кВт'], "ВШ1 вкл/выкл А":stack[-1].loc[i]['ВШ1 вкл/выкл А'], "ВШ2 вкл/выкл А":stack[-1].loc[i]['ВШ2 вкл/выкл А'], "Время ч":stack[-1].loc[i]['Время ч'], "Расход топлива л":stack[-1].loc[i]['Расход топлива л'], "Замечания":stack[-1].loc[i]['Замечания'], "Мастер р/и":stack[-1].loc[i]['Мастер р/и']}
#						##print(dct)
#						#df = df.append(dct, ignore_index=True)
#						#df = df.append(stack[-1].loc[i], ignore_index=True)
#			elif filtercmd == 5:
#				stack.append(df)
#				df = df[df.Номер == int(input("Введите номер: "))]
#			elif filtercmd == 0:
#				continue
#		elif cmd == 5:
#			try:
#				df = stack.pop()
#			except IndexError:
#				print("Возможно, вы уже находитесь в исходной таблице")
#		elif cmd == 6:
#			#pdf = FPDF()
#			#pdf.add_page()
#			#pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
#			#pdf.set_font('DejaVu', '', 14)
#			#pdf.cell(30, 10, "									  Реостатные испытания\n")
#			#pdf.cell(ln=0, h=5.0, align='L', w=0, txt=f"Дата: {df['Дата']} Серия: {df['Серия']} Номер: {df['Номер']} Секция: {df['Секция']} Ремонт: {df['Ремонт']} Вид испытаний: {df['Вид испытаний']} Мощность НР кВт: {df['Мощность НР кВт']} Мощность АР кВт: {df['Мощность АР кВт']}\n ВШ1 вкл/выкл А: {df['ВШ1 вкл/выкл А']} ВШ2 вкл/выкл А: {df['ВШ2 вкл/выкл А']} Время ч: {df['Время ч']}\n Расход топлива л: {df['Расход топлива л']}\n", border=0)
#			#pdf.output("Распечатка.pdf", "F")
#			print("Выберете нужный элемент, введите его индекс (индексация начинается с нуля) (-1 - отмена)")
#			print(df)
#			idx = int(input("Введите индекс нужного элемента: "))
#			if idx != -1:
#				pdf = PDF()
#				df = df.fillna('	  ')
#				pdf.colontitle = "Реостатные испытания"
#				try:
#					pdf.print_chapter(txt=f"Дата: {df.loc[idx]['Дата'].date()}		   Серия: {df.loc[idx]['Серия']}		 Номер: {int(float(df.loc[idx]['Номер']))}		  Секция: {df.loc[idx]['Секция']}		 Ремонт: {df.loc[idx]['Ремонт']}\nВид испытаний: {df.loc[idx]['Вид испытаний']}\nМощность НР кВт: {df.loc[idx]['Мощность НР кВт']}		  ВШ1 вкл/выкл А: {df.loc[idx]['ВШ1 вкл/выкл А']}		 Наддув атм: {df.loc[idx]['Наддув атм']}\nМощность АР кВт: {df.loc[idx]['Мощность АР кВт']}		   ВШ2 вкл/выкл А: {df.loc[idx]['ВШ2 вкл/выкл А']}		  Время ч: {df.loc[idx]['Время ч']}\nРасход топлива л: {df.loc[idx]['Расход топлива л']}\nЗамечания:\n{df.loc[idx]['Замечания']}\n		  Мастер р/и: {df.loc[idx]['Мастер р/и']}")
#					pdf.output("Распечатка.pdf")
#					print("Готово!")
#					os.system(".\Распечатка.pdf")
#				except KeyError:
#					print("Нет такого пункта")
#		elif cmd == 7:
#			df.to_excel(path, index=False)
#		elif cmd == 8:
#			path = input("Введите путь до файла .xlsx (вместе с расширением): ")
#			df = pd.read_excel(path)
#		elif cmd == 9:
#			print(df)
#			df = df.drop(labels=[int(input("Введите индекс элемента для удаления: "))])
#

main_window()
