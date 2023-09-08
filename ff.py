import xlwings as wb

import ttkbootstrap as tb
from tkinter import filedialog
from tkinter import *
import os

from openpyxl.utils import get_column_letter

class My_Gui():

	def __init__(self,root):
		
		self.lis_type_search_cell = ["Hàng Cuối Data","Data Cuối Cùng","Hàng Cuối(Ngắt Dòng)","Data Cuối(Ngắt Dòng)"]

		self.values_spinbox_column = [get_column_letter(int(i)) for i in range(1, 18279)] # Các Row Excel 

		self.widget()

		self.balance(root)




	def widget(self):


		

		self.spin_box = tb.Spinbox(bootstyle="danger",width=2, from_=1, to=18279, font=("Arial", 11),state="readonly",

			justify="center",values=self.values_spinbox_column)

		self.spin_box.grid(row=0,column=1,stick="e",padx=5,pady=5)

		self.spin_box.set("A")

		self.my_button_zero = tb.Button(text="Log File",bootstyle="outline",command=self.log_file)

		self.my_button_zero.grid(row=0,column=0,stick="w",padx=5,pady=5)

		

		

		self.my_label = tb.Label(text="Nằm Ở Cell", font=("Helvetica", 18))

		self.my_label.grid(row=1,column=0,columnspan=2)

		self.my_combobox_sheet =tb.Combobox(bootstyle="dark",state="readonly",justify="center",width=7)

		self.my_combobox_sheet.grid(row=1,column=0,stick="w",padx=5)



		self.my_entry = tb.Entry(bootstyle="danger",foreground='gray')

		self.my_entry.grid(row=2,column=0,columnspan=2)

		self.my_entry.insert(0, "Nhập Data...")

		self.my_entry.bind('<FocusOut>', self.on_focus_out_data)

		self.my_entry.bind('<FocusIn>', self.on_entry_click_data)


		self.my_combobox =tb.Combobox(bootstyle="dark",state="readonly",justify="center",values =self.lis_type_search_cell)

		self.my_combobox.grid(row=3,column=0,columnspan=2)

		self.my_combobox.current(0)

		

		self.my_button_one = tb.Button(text="Enter Data",bootstyle="success-outline",command=self.enter)

		self.my_button_one.grid(row=4,column=0,stick="w",pady=20,padx=40)

		self.my_button_two = tb.Button(text="Tìm Addres",bootstyle="outline-warning",command=self.search)

		self.my_button_two.grid(row=4,column=1,stick="e",pady=20,padx=40)

	def log_file(self):

		self.filename = filedialog.askopenfilename(title="Chọn File Excel Data",filetypes=[("Excel files", ".xlsx .xlsm .xltx .xltm .xlt")])

		if self.filename =="":

			pass
		
		self.name_file_data = os.path.basename(self.filename)


		self.open_file_excel = wb.Book(self.filename)

		lis_sheet = [x.name for x in self.open_file_excel.sheets]

		self.my_combobox_sheet["values"] = lis_sheet

		self.my_combobox_sheet.current(0)


	def enter(self):

		try:
			self.search()

			self.sht = self.open_file_excel.sheets[self.my_combobox_sheet.get()]

			if self.my_entry.get() =="":

				pass

			else:


				self.sht.range(self.spin_box.get()+str(self.last_row)).value = self.my_entry.get()

		except AttributeError:

			print("lỗi")

		

	def search(self):

		try:

			self.sht = self.open_file_excel.sheets[self.my_combobox_sheet.get()]

			if self.my_combobox.get() == self.lis_type_search_cell[0]:

				self.last_row = self.sht.range(self.spin_box.get() + str(self.sht.cells.last_cell.row)).end('up').row

				if self.sht.range(self.spin_box.get()+str(self.last_row)).value ==None or "":

					pass
				else:

					self.last_row +=1


			elif self.my_combobox.get() == self.lis_type_search_cell[1]:

				self.last_row = self.sht.range(self.spin_box.get() + str(self.sht.cells.last_cell.row)).end('up').row


			elif self.my_combobox.get() == self.lis_type_search_cell[2]:


				self.last_row = self.sht.range(self.spin_box.get()+"1").end('down').row

				if self.sht.range(self.spin_box.get()+str(self.last_row)).value ==None or "":

					if self.last_row == 1048576: #________________trường hợp nầy quét từ dưới lên nên đặt 1 điều kiện

						self.last_row =1

					pass
				else:


						self.last_row +=1

			elif self.my_combobox.get() == self.lis_type_search_cell[3]:

				self.last_row = self.sht.range(self.spin_box.get()+"1").end('down').row

				if self.last_row == 1048576: #________________trường hợp nầy quét từ dưới lên nên đặt 1 điều kiện

					self.last_row =1


			self.my_label.config(text ="Nằm Ở Cell:\n  " + self.spin_box.get()+str(self.last_row))
		except AttributeError:

			print("lỗi")


	def on_focus_out_data(self, event):

		if self.my_entry.get() == "":
			
			self.my_entry.insert(0, "Nhập Data...")
			
			self.my_entry.config(foreground="gray")
			

	def on_entry_click_data(self, event):
		
		if self.my_entry.get() == "Nhập Data...":
			
			self.my_entry.delete(0, "end")
			
			self.my_entry.config(foreground="yellow")

	def balance(self,root):



		lis_Lableframedata_gui2 = [self.my_label,self.my_entry,self.my_button_one,self.my_button_two]

		for widget in lis_Lableframedata_gui2:
			Grid.rowconfigure(root, widget, weight=1, minsize=0)
			Grid.columnconfigure(root, widget, weight=1, minsize=0)


if __name__ == "__main__":

	root = tb.Window(themename="vapor")


	root.geometry('350x250')

	root.title("Tìm Cell Cuối Excel")

	My_Gui(root)
	root.mainloop()

	













