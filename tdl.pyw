import tkinter as t
import xlrd
import xlwt
import random

item_list=[]
list_name="tdl.xls"
motto_name="mn.xls"

def load_motto():
	try:
		data = xlrd.open_workbook(motto_name)
	except FileNotFoundError as F:
		return None
	sheet = data.sheets()[0]
	nrows = sheet.nrows
	items=sheet.col_values(0)
	if (items==[]):
		return None
	random.shuffle(items)
	return items[0]

def erase(parent,name):
	item_list.remove(name)
	parent.pack_forget()
	
def h1(e,temp):
	add_item(e.get())
	temp.withdraw()
	
def get_name():
	temp=t.Toplevel()
	f=t.Frame(temp)
	f.pack()
	l=t.Label(f,text="请输入新事项名称")
	e=t.Entry(f)
	l.pack()
	e.pack()
	b=t.Button(f,text="确定",command=lambda:h1(e,temp))
	b.pack()
	
	
def add_item(name):
	name="    "+name+"    "
	frame=t.Frame()
	frame.pack()
	i=t.Label(frame,text=name)
	i.pack(side="left")
	b=t.Button(frame,command=lambda:erase(frame,name),text="删除")
	b.pack(side="right")
	item_list.append(name)
	
def head(sentencetence):
	if(sentence==None):
		sentence="默认格言"
	f=t.Frame(width=1200,height=1600)
	f.pack()
	h=t.Label(f,text=sentence)
	h.pack()
	b=t.Button(f,text="添加新事项",command=get_name)
	b.pack(side="left")
	s=t.Button(f,text="保存",command=save_list)
	s.pack(side="right")

def save_list():
	item_book=xlwt.Workbook()
	item_sheet=item_book.add_sheet("items")
	for i in range(len(item_list)):
		item_sheet.write(i,0,item_list[i])
	item_book.save(list_name)
	
def load_list():
	data = xlrd.open_workbook(list_name)
	sheet = data.sheets()[0]
	nrows = sheet.nrows
	items=sheet.col_values(0)
	for item in items:
		add_item(item)
		
def create_root():
	root=t.Tk()
	root.title("To Do List")
	root["height"]=1600
	root["width"]=1200
	return root

def main():
	root=create_root()
	head(load_motto())
	load_list()
	root.mainloop()
	save_list()
	
main()