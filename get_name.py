from tkinter import *
import runpy


RED = '#ff4d4d'

def close():
	try:
		new.destroy()
	except:
		pass
	file = runpy.run_path('Stocks manager.py')


def write():
	
	file = open("files\\company.txt","w+")
	file.write(name.get()+'\n')
	file.write(lin1.get()+'\n')
	file.write(lin2.get()+'\n')
	file.close()
	new.destroy()
	file = runpy.run_path('Stocks manager.py')

new = Tk()
new.iconbitmap(r'images\\icon.ico')
WIDTH = new.winfo_screenwidth()
HEIGHT = new.winfo_screenheight()
new.protocol("WM_DELETE_WINDOW", close)
name = StringVar()
lin1 = StringVar()
lin2 = StringVar()


try:
	file = open('files\\company.txt', 'r+')
	line = file.readlines()
	name.set(line[0].split('\n')[0])
	if len(line[1]) > 0:
		lin1.set(line[1].split('\n')[0])
		lin2.set(line[2].split('\n')[0])
	# else:
	# 	lin1.set(line[2])
	# 	lin2.set(line[4])
	file.close()
except Exception as e:
	print(e)


new.geometry('500x300+%d+%d'%(WIDTH // 2 - 250, HEIGHT // 2 - 150))

new.title("Company name")
new.configure(bg = RED)

Label(new, text = 'Company name:' , bg = RED, font = ('times',20)).grid(row = 1, column = 1, pady = 10)
Entry(new,textvariable = name, font = ('times',15)).grid(row = 1, column = 2)

Label(new, text = 'Adress line 1:' , bg = RED, font = ('times',20)).grid(row = 2, column = 1, pady = 10)
Entry(new,textvariable = lin1, font = ('times',15)).grid(row = 2, column = 2)

Label(new, text = 'Adress line 2:' , bg = RED, font = ('times',20)).grid(row = 3, column = 1, pady = 10)
Entry(new,textvariable = lin2, font = ('times',15)).grid(row = 3, column = 2)

Button(new, text = "Submit" ,  bg = "#FF1A1A" , command = write).place(x = 230 , y = 230)

new.mainloop()