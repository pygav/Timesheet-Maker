from openpyxl import *
from timesheetmaker import *
from tkinter import *
from PIL import ImageTk,Image

from copy import copy

root = Tk()
root.withdraw()
logo = ImageTk.PhotoImage(Image.open('D://lnbtimesheetassembler/imgs/logo.png'))

welcome = Toplevel()
welcome.iconbitmap('D://lnbtimesheetassembler/iconwindow.ico')
welcome.title("Light & Breuning Timesheet Assembler")

welcome.geometry('1000x500+420+200')

welcome.configure(bg='white')

set_logo = Label(welcome,image=logo, borderwidth=0).place(x=170, y=50)

enter_name_label = Label(welcome, text="Enter employee name below:", font="Tahoma 20 bold", bg='white', justify = 'center').place(x=295,y=200)

credit_to_creator = Label(welcome, text='Created by Hardware/Software Tech Gavin Truex', font="Tahoma 12 bold", bg='white', fg='blue', justify = 'center').place(x=285, y=400)


def windowCloseWarning():

	close_warning_window = Toplevel()
	close_warning_window.iconbitmap('D://lnbtimesheetassembler/iconwindow.ico')
	close_warning_window.title("Warning!")

	close_warning_window.geometry('300x150+790+350')

	close_warning_window.configure(bg='white')

	ask_if_want_to_quit = Label(close_warning_window, text='Are you sure you want to exit?\nAll unsaved work will be lost', font='Tahoma 12 bold', bg='white').place(x=20 , y= 25)

	confirm_quit = Button(close_warning_window, text = 'Yes, Exit', font='Tahoma 12', bg='white', command = root.destroy).place(x=70, y=100)

	cancel_quit = Button(close_warning_window, text='Cancel',font='Tahoma 12', bg='white', command = close_warning_window.destroy).place(x=170, y=100)

def finishAndSave(name, day, date, previous_window):
	
	
	for i in range (1, mr +1):
		for j in range (1, mc + 1):
			c=ws1.cell(row=i, column=j)
			ws2.cell(row=i, column=j).value=c.value

	

	wb2.save('Created Timesheet.xlsx')

	root.destroy()

def askIfFinished(name, day, date, previous_window):

	ask_if_finish_window = Toplevel()
	ask_if_finish_window.iconbitmap('D://lnbtimesheetassembler/iconwindow.ico')
	ask_if_finish_window.geometry('300x150+790+350')

	ask_if_finish_window.title("Warning!")

	ask_if_finish_window.configure(bg='white')

	ask_if_want_to_quit = Label(ask_if_finish_window, text='Are you sure you are finished?\nYou cannot go back and\nchange data when finished', font='Tahoma 12 bold', bg='white').place(x=20 , y= 25)

	confirm_quit = Button(ask_if_finish_window, text = 'Yes, Finish', font='Tahoma 12', bg='white' ,command = lambda:finishAndSave(name, day, date, previous_window)).place(x=70, y=100)

	cancel_quit = Button(ask_if_finish_window, text='Cancel',font='Tahoma 12', bg='white', command = ask_if_finish_window.destroy).place(x=170, y=100)


	











def additionalEvents(name, day, date, previous_window, first_event_check, timestop_cell, jobcode_cell, job_descrip_cell,order_cell, end_odo_cell, event_number, event_info, order_number, start_odo, end_odo, code, start_time, end_time):

	event_number = event_number + 1

	
	
	if first_event_check:

		ws1['H3'].value = start_odo

		print(ws1['H3'].value)

		ws1['H5']  = end_odo

		print(ws1['H5'].value)

		ws1['E5'] = order_number

		print(ws1['E5'].value)

		ws1['F5'].value = code

		print(ws1['F5'].value)

		ws1['G5'] = event_info

		print(ws1['G5'].value)

		ws1['A5'] = start_time

		print(ws1['A5'].value)

		ws1['B5'] = end_time

		print(ws1['B5'].value)
		first_event_is_true = False
	
	else:
		
		first_event_is_true = False

		end_odo_cell = end_odo_cell + 1

		timestop_cell = timestop_cell + 1

		job_descrip_cell = job_descrip_cell +1

		jobcode_cell = jobcode_cell + 1

		order_cell = order_cell + 1

		ws1[f'H{end_odo_cell}'].value = end_odo

		print(ws1[f'H{end_odo_cell}'].value)

		ws1[f'E{order_cell}'].value = order_number

		print(ws1[f'E{order_cell}'].value)

		ws1[f'F{jobcode_cell}'].value = code

		print(ws1[f'F{jobcode_cell}'].value)

		print(ws1['F6'].value)

		ws1[f'G{job_descrip_cell}'].value = event_info

		print(ws1[f'G{job_descrip_cell}'].value)

		ws1[f'B{timestop_cell}'].value = end_time

		print(ws1[f'B{timestop_cell}'].value)



	first_event_is_true = False

	additional_events_window = Toplevel(root)
	additional_events_window.iconbitmap('D://lnbtimesheetassembler/iconwindow.ico')
	previous_window.destroy()

	additional_events_window.protocol("WM_DELETE_WINDOW", windowCloseWarning)

	additional_events_window.title(f"Light & Breuning Timesheet Assembler for {name}")

	additional_events_window.geometry('1000x500+420+200')

	additional_events_window.configure(bg='white')

	tell_user_event = Label(additional_events_window, text = f"Events: Event {event_number}",font = 'Tahoma 26 bold', bg='white')

	tell_user_event.place(x=10, y=15)

	prompt_user_to_enter_detail = Label(additional_events_window, text = f"Describe what happened in event below:",font = 'Tahoma 16', bg='white')

	prompt_user_to_enter_detail.place(x=65, y=160)

	emp_enter_event = Text(additional_events_window, height=10, width=50, borderwidth = 6)

	emp_enter_event.place(x=60, y=200)

	prompt_emp_work_order = Label(additional_events_window, text = "WO Num. :", font = 'Tahoma 12 bold', bg='white')

	prompt_emp_work_order.place(x=485, y=75)

	get_wo_info = Entry(additional_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

	get_wo_info.place(x=580, y=74)                            

	prompt_emp_odometer_end = Label(additional_events_window, text = "End Odometer:", font = 'Tahoma 12 bold', bg='white')

	prompt_emp_odometer_end.place(x=485, y=150)

	get_end_odo_info = Entry(additional_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

	get_end_odo_info.place(x=625, y=148)

	get_wo_info = Entry(additional_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

	get_wo_info.place(x=580, y=74)

	code_list = ['W','L','A','B','C','D','E','F','G','H','I','J','K','M','N','O','P','Q','R','S','T','U','V','X','Y','Z']

	code_node = StringVar()

	code_node.set("Job Code")

	job_code_drop = OptionMenu(additional_events_window, code_node, *code_list)

	job_code_drop.place(x= 600, y = 200)

	job_code_drop.configure(font = 'Tahoma 18 bold')

	event_time = ['06:00:00','06:15:00','06:30:00','06:45:00','07:00:00','07:15:00','07:30:00','07:45:00','08:00:00','08:15:00','08:30:00','08:45:00','09:00:00','09:15:00','09:30:00','09:45:00','10:00:00','10:15:00','10:30:00','10:45:00','11:00:00','11:15:00','11:30:00','11:45:00',
			'12:00:00','12:15:00','12:30:00','12:45:00','13:00:00','13:15:00','13:30:00','13:45:00','14:00:00','14:15:00','14:30:00','14:45:00','15:00:00','15:15:00','15:30:00','15:45:00','16:00:00','16:15:00','16:30:00','16:45:00','17:00:00','17:15:00','17:30:00','17:45:00','18:00:00',
			'18:15:00','18:30:00','18:45:00','19:00:00','19:15:00','19:30:00','19:45:00','20:00:00','20:15:00','20:30:00','20:45:00','21:00:00','21:15:00','21:30:00','21:45:00','22:00:00','22:15:00','22:30:00','22:45:00','23:00:00',]

	event_start_time = StringVar()

	event_start_time.set(end_time)

	event_start_time_drop = OptionMenu(additional_events_window,event_start_time, *event_time)

	event_start_time_drop.config(font='Tahoma 18 bold')

	event_start_time_drop.place(x=800, y=200)

	event_stop_time = StringVar()

	event_stop_time.set('Stop Time')

	event_stop_time_drop = OptionMenu(additional_events_window,event_stop_time, *event_time)

	event_stop_time_drop.config(font='Tahoma 18 bold')

	event_stop_time_drop.place(x=800, y=300)

	go_to_next_event = Button(additional_events_window, text = "Next", font = 'Tahoma 18 bold', borderwidth = 3, command = lambda: additionalEvents(name, day, date, additional_events_window, first_event_is_true, timestop_cell, jobcode_cell, job_descrip_cell,order_cell, end_odo_cell, event_number, emp_enter_event.get(1.0, "end-1c"), get_wo_info.get(), 0, get_end_odo_info.get(), code_node.get(), event_start_time.get(), event_stop_time.get()))

	go_to_next_event.place(x=760, y=425)

	go_to_finish = Button(additional_events_window, text = "Finish", font = 'Tahoma 18 bold', borderwidth = 3, command = lambda: askIfFinished(name, day, date, additional_events_window))

	go_to_finish.place(x=870, y=425)



def startEventEntry(previous_window, name, day, date_day, date_month, date_year):

	if day == 'Weekday' or date_month == 'Month' or date_day == 'Day' or date_year == 'Year':
		warning_invalid_dates_days = Label(previous_window, text = 'Sorry, but you left some blanks, please continue filling out the day and date for the timesheet', font = 'Tahoma 12 bold', fg='red', bg='white')
		
		warning_invalid_dates_days.place(x=10, y = 460)

	else:
		event_ticker = 1



		previous_window.destroy()

		start_events_window = Toplevel(root)
		start_events_window.protocol("WM_DELETE_WINDOW", windowCloseWarning)
		start_events_window.iconbitmap('D://lnbtimesheetassembler/iconwindow.ico')
		ws1['C3'].value = day

		print(ws1['C3'].value)

		present_date = f'{date_month}/{date_day}/{date_year}'

		ws1['A3'].value = present_date

		print(ws1['A3'].value)

		start_events_window.title(f"Light & Breuning Timesheet Assembler for {name}")

		start_events_window.geometry('1000x500+420+200')

		start_events_window.configure(bg='white')

		tell_user_event = Label(start_events_window, text = f"Events: Event {event_ticker}",font = 'Tahoma 26 bold', bg='white')

		tell_user_event.place(x=10, y=15)

		prompt_user_to_enter_detail = Label(start_events_window, text = f"Describe what happened in event below:",font = 'Tahoma 16', bg='white')

		prompt_user_to_enter_detail.place(x=65, y=160)

		emp_enter_event = Text(start_events_window, height=10, width=50, borderwidth = 6)

		emp_enter_event.place(x=60, y=200)

		prompt_emp_work_order = Label(start_events_window, text = "WO Num. :", font = 'Tahoma 12 bold', bg='white')

		prompt_emp_work_order.place(x=485, y=75)

		get_wo_info = Entry(start_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

		get_wo_info.place(x=580, y=74)

		prompt_emp_odometer_start = Label(start_events_window, text = "Starting Odometer:", font = 'Tahoma 12 bold', bg='white')

		prompt_emp_odometer_start.place(x=485, y=115)

		get_start_odo_info = Entry(start_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

		get_start_odo_info.place(x=650, y=114)

		prompt_emp_odometer_end = Label(start_events_window, text = "End Odometer:", font = 'Tahoma 12 bold', bg='white')

		prompt_emp_odometer_end.place(x=485, y=150)

		get_end_odo_info = Entry(start_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

		get_end_odo_info.place(x=625, y=148)

		get_wo_info = Entry(start_events_window, width=20, font = 'Tahoma 12', borderwidth=5)

		get_wo_info.place(x=580, y=74)

		code_list = ['W','L','A','B','C','D','E','F','G','H','I','J','K','M','N','O','P','Q','R','S','T','U','V','X','Y','Z']

		code_node = StringVar()

		code_node.set("Job Code")

		job_code_drop = OptionMenu(start_events_window, code_node, *code_list)

		job_code_drop.configure(font = 'Tahoma 18 bold')

		job_code_drop.place(x= 600, y = 200)

		

		event_time = ['06:00:00','06:15:00','06:30:00','06:45:00','07:00:00','07:15:00','07:30:00','07:45:00','08:00:00','08:15:00','08:30:00','08:45:00','09:00:00','09:15:00','09:30:00','09:45:00','10:00:00','10:15:00','10:30:00','10:45:00','11:00:00','11:15:00','11:30:00','11:45:00',
			'12:00:00','12:15:00','12:30:00','12:45:00','13:00:00','13:15:00','13:30:00','13:45:00','14:00:00','14:15:00','14:30:00','14:45:00','15:00:00','15:15:00','15:30:00','15:45:00','16:00:00','16:15:00','16:30:00','16:45:00','17:00:00','17:15:00','17:30:00','17:45:00','18:00:00',
			'18:15:00','18:30:00','18:45:00','19:00:00','19:15:00','19:30:00','19:45:00','20:00:00','20:15:00','20:30:00','20:45:00','21:00:00','21:15:00','21:30:00','21:45:00','22:00:00','22:15:00','22:30:00','22:45:00','23:00:00',]

		event_start_time = StringVar()

		event_start_time.set('Start Time')

		event_start_time_drop = OptionMenu(start_events_window,event_start_time, *event_time)

		event_start_time_drop.config(font='Tahoma 18 bold')

		event_start_time_drop.place(x=800, y=200)

		event_stop_time = StringVar()

		event_stop_time.set('Stop Time')

		event_stop_time_drop = OptionMenu(start_events_window,event_stop_time, *event_time)

		event_stop_time_drop.config(font='Tahoma 18 bold')

		event_stop_time_drop.place(x=800, y=300)

		first_event_is_true = True

		go_to_next_event = Button(start_events_window, text = "Next", font = 'Tahoma 18 bold', borderwidth = 3, command = lambda: additionalEvents(name, day, present_date, start_events_window, first_event_is_true, 5, 5, 5, 5,5, event_ticker, emp_enter_event.get(1.0, "end-1c"), get_wo_info.get(), get_start_odo_info.get(), get_end_odo_info.get(), code_node.get(), event_start_time.get(), event_stop_time.get()))

		go_to_next_event.place(x=760, y=425)

		go_to_finish = Button(start_events_window, text = "Finish", font = 'Tahoma 18 bold', borderwidth = 3, state = DISABLED)

		go_to_finish.place(x=870, y=425)

		first_event = True
 








def giveDateDay():


	give_name = enter_name_input.get()

	if give_name == '':

		warning_no_name = Label(welcome, text = 'Please enter a name to continue', font = 'Tahoma 12 bold', fg='red', bg='white')

		warning_no_name.place(x=360, y=300)

	else:

		ws1['E3'].value = give_name

		print(ws1['E3'].value)

		welcome.destroy()

		specify_date_and_day = Toplevel(root)
		specify_date_and_day.iconbitmap('D://lnbtimesheetassembler/iconwindow.ico')
		specify_date_and_day.title(f"Light & Breuning Timesheet Assembler for {give_name}")

		specify_date_and_day.geometry('1000x500+420+200')

		specify_date_and_day.configure(bg='white')

		specify_date_and_day.configure(bg='white')

		welcome_user = Label(specify_date_and_day, text = f'Welcome, {give_name}', font = 'Tahoma 26 bold', bg='white').place(x= 30, y = 25)

		tell_user_to_select_day = Label(specify_date_and_day, text = 'Please select the day\nof the week you would like\nto make a timesheet for:', font='Tahoma 16', bg='white').place(x= 375, y=100)

		day_selection = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
		
		day_node = StringVar()

		day_node.set("Weekday")

		day_selection_drop = OptionMenu(specify_date_and_day, day_node, *day_selection)

		day_selection_drop.config(font='Tahoma 18 bold')

		day_selection_drop.place(x=430, y=200)

		tell_user_to_select_date = Label(specify_date_and_day, text = 'Please select the date\nof specified day:', font='Tahoma 16', bg='white').place(x= 395, y=280)

		date_month_selection = ['01', '02', '03', '04','05', '06','07', '08', '09', '10','11','12']

		date_month_node = StringVar()

		date_month_node.set('Month')

		date_month_selection_drop = OptionMenu(specify_date_and_day,date_month_node, *date_month_selection)

		date_month_selection_drop.config(font='Tahoma 18 bold')

		date_month_selection_drop.place(x=305, y=350)

		date_day_selection = ['01', '02', '03', '04','05', '06','07', '08', '09', '10','11','12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']

		date_day_node = StringVar()

		date_day_node.set('Day')

		date_day_selection_drop = OptionMenu(specify_date_and_day,date_day_node, *date_day_selection)

		date_day_selection_drop.config(font='Tahoma 18 bold')

		date_day_selection_drop.place(x=460, y=350)

		date_Year_selection = ['2022', '2023', '2024', '2025','2026', '2027', '2028','2029', '2030', '2030','2031','2032','2033','2034','2035','2036','2037','2038','2039','2040']

		date_Year_node = StringVar()

		date_Year_node.set('Year')

		date_Year_selection_drop = OptionMenu(specify_date_and_day,date_Year_node, *date_Year_selection)

		date_Year_selection_drop.config(font='Tahoma 18 bold')

		date_Year_selection_drop.place(x=600, y=350)

		go_to_event_entry = Button(specify_date_and_day, text='Next', font = "Tahoma 22", borderwidth = 5, command = lambda: startEventEntry(specify_date_and_day, give_name, day_node.get(), date_day_node.get(), date_month_node.get(), date_Year_node.get())).place(x=800, y=420)





enter_name_input = Entry(welcome, width=40, font= "Tahoma 18", bg='white', borderwidth=3)

enter_name_input.place(x=240, y=250)

name_entered_go_to_next = Button(welcome, text='Next', font = "Tahoma 22", command = giveDateDay).place(x=455, y=330)












if __name__ == '__main__':

	root.mainloop()
