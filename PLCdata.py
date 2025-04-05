#############################################
#TDC 3-35-25
#
##############################################
from time import sleep
from tkinter import *
import tkinter.font
from plcconn import *
import re
from pycomm3 import LogixDriver, CIPDriver
from datetime import datetime
import os
from openpyxl import Workbook
import openpyxl



win = Tk()
win.title("datacollection")
win.geometry('1280x1024')
#win.state("zoomed")

data = []
IPpath = ''
arraylist = []

heading_font = tkinter.font.Font(family="Helvetica", size=10, weight="bold")
timems = 10000
readok = False
readvalues = []
#testtags = ["tag1", "tag2", "motor_on", "input1", "output1", "monday", "friday"] # for testing
msg = StringVar()
msg.set('Errors go Here')

win.grid_columnconfigure(0, weight=1)
#win.grid_rowconfigure(0, weight=1)

#------------------Frames for easier gui row-column setup------------------------#
frame = Frame(win, width=700, height=50)
frame.grid(row=0, column=0, columnspan=4, sticky='w')

frame1 = Frame(win, width=600, height=35, bg='beige')
frame1.grid(row=1, column=0, columnspan=4, sticky='w')

frame2 = Frame(win, width=100, height=200, bg='orange')
frame2.grid(row=3, column=0, padx=10, pady=10, columnspan=1, sticky='w')

frame3 = Frame(win, width=500, height=35)
frame3.grid(row=2, column=0, columnspan=4, sticky='w')

frame4 = Frame(win, width=500, height=35, bg='red')
frame4.grid(row=4, column=0, padx=10, pady=10, columnspan=3, sticky='s')

frame5 = Frame(win, width=300, height=200, bg='beige')
frame5.grid(row=3, column=1, sticky='w')
#---------------------------------------------------------------------------------#

def getip():
     global IPpath
     IPpath = Pathip.get()

     if is_valid_ip(IPpath):
        tblbl = Label(frame1, text=f'{IPpath} is Valid and Saved!', font='Arial', fg='Green')
        tblbl.grid(row=0, column=5, sticky='w')
     else:
        tblbl = Label(frame1, text='Invalid IP Address! Try Again!', font='Arail', fg="Red")
        tblbl.grid(row=0, column=5, sticky='w')

     Pathip.delete(0, END)
 
def is_valid_ip(ip):
    # Regular expression for a valid IPv4 address
    pattern = re.compile(r'^((25[0-5]|2[0-4][0-9]|[0-1]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[0-1]?[0-9][0-9]?)$')
    return pattern.match(ip) is not None
          
      
def discoverplc():# gets all tags from plc
    global IPpath
    try:
     with LogixDriver(IPpath, init_program_tags=True) as plc:
            # Get all controller and program tags
            tag_list = plc.get_tag_list(program='*')      

           # Insert tags into the text box
            for tag in tag_list:
                text_box_gettags.insert(END, tag['tag_name'] + "\n")
    except Exception as e:
        msg.set(f"Cannot get tags from PLC, check connection. Error: {e}")
     
     
#def testxcel(): for testing change writeexcel on discover button to testxcel
    #msg.set("made it") testing
  #  excelwrite()

#-----------------------------------------------------------------------------------------------------------------#
def read_plc_a(): #Read array from PLC 
  global IPpath, arraylist, timems, readok

  if not readok: 
        return  # Exit the function without continuing
  intervalms()
  readlinelbl.config(text=" ")
  
  try:
   
      array_to_read = entry_Array.get()
      with LogixDriver(IPpath) as plc:
        #array syntax nameofarray{#of elements} example: testarray{5}
        read_plc_array = plc.read(array_to_read) # read 5 elements starting at 0 from an array
        arraylist = read_plc_array.value
        #print(arraylist) for debug   
        excelwrite()
        readlinelbl.config(text=", ".join(map(str, arraylist))) #print to label on gui to see string being read
        arraylist.clear()# clear the array or it keeps growing!!
        if readok:
          win.after(timems, read_plc_a)
  except Exception as e:
        msg.set(f"Cannot connect to PLC. Error: {e}") 

#-----------------------------------------------------------------------------------------------------------------#
def read_plc_s():#read string from plc
   global IPpath, arraylist, timems, readok

   if not readok: 
        return  # Exit the function without continuing
   intervalms()
   readlinelbl.config(text=" ")
   
   try:
     
       string_to_read = entry_string.get()
       with LogixDriver(IPpath) as plc:
         read_plc_string = plc.read(string_to_read)  #read values from plc
         arraylist.append(read_plc_string.value)
        # print(arraylist) for debug
         excelwrite()
         readlinelbl.config(text=", ".join(map(str, arraylist))) #print to label on gui to see string being read
         arraylist.clear()
         if readok:
           win.after(timems, read_plc_s)
   except Exception as e:
        msg.set(f"Cannot connect to PLC. Error: {e}") 

#-----------------------------------------------------------------------------------------------------------------#
def read_plc_u(): #read udt from plc
    global IPpath, arraylist, timems, readok

    if not readok:  # Check if logging should stop
        return  # Exit the function without continuing
    
    udt_to_read = entry_udt.get()
    intervalms()
    readlinelbl.config(text=" ")
    
    try:
       
          with LogixDriver(IPpath) as plc:
            read_udt = plc.read(udt_to_read) 
            if read_udt and hasattr(read_udt, "value"):  # Ensure it's valid
                value_list = list(read_udt.value.values())  # Extract the dictionary and convert to a list
                
                for value in value_list:
                    if isinstance(value, list):  # If value is a list, extend instead of append
                        arraylist.extend(value)    
                    else:
                        arraylist.append(value)
          # Add current datetime to the list
          current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
          arraylist.append(current_datetime)                   
          #print(arraylist) for debugging
          excelwrite()
          readlinelbl.config(text=", ".join(map(str, arraylist))) #print to label on gui to see string being read
          arraylist.clear()
          if readok:
            win.after(timems, read_plc_u)
    except Exception as e:
        msg.set(f"Cannot connect to PLC. Error: {e}")
        
#-----------------------------------------------------------------------------------------------------------------#        
def plc_read_m(): #read multitag from plc
    global IPpath, arraylist, timems, readok
    if not readok: 
        return  # Exit the function without continuing
    intervalms()
    readlinelbl.config(text=" ")
    multilist = []
   
    try:    
         multi_to_read = multi_tag.get()
         multilist = [item.strip() for item in multi_to_read.split(',')]# need to separate list or it thinks its one string!!
       
         with LogixDriver(IPpath) as plc: 
           for i in multilist:     
             multiread = plc.read(i)  # read tags 
             arraylist.append(multiread.value) # append to our writing list to excel
        #print(arraylist) for debugging
         excelwrite()
         readlinelbl.config(text=", ".join(map(str, arraylist))) #print to label on gui to see string being read
         arraylist.clear()
         if readok:
           win.after(timems, plc_read_m)
    except Exception as e:
        msg.set(f"Cannot connect to PLC. Error: {e}")        
      
#-----------------------------------------------------------------------------------------------------------------#       
def excelwrite():
  global data, file2, arraylist
# this is your headers for excel sheet only populates if new sheet started
  dataheader = ['testheader1','testheader2','testheader3','testheader4','testheader5','DateTime']#headers for new sheet

  current_datetime = datetime.now().strftime("%Y_%m_%d")

  #change folder location when get up and running to save location of file to be sent
  #file2 = "C:\\Users\\name\\documents\\test.xlsx" + 
  file2 = "C:\\Users\\name\\documents\\test_" + current_datetime + ".xlsx"
  try:
    if os.path.isfile(file2):  # if file already exists append to existing file
     wb = openpyxl.load_workbook(file2)  # load workbook if already exists
     ws = wb.active
     # append the data results to the current excel file
     ws.append(arraylist)
     wb.save(file2)  # save workbook
     wb.close()
    
    else:  # create the excel file if doesn't already exist
        wb = Workbook()
        ws = wb.active
        ws.append(dataheader)
        ws.append(data) 
        wb.save(file2)  # save workbook
        wb.close()
  except Exception as e:
        msg.set(f"Cannot connect to Excel. Error: {e}")           
#-----------------------------------------------------------------------------------------------------------------#
#       
def intervalms():
    global timems
    interval = int(timeentryms.get()) 
    if interval > 0:
      timems = interval * 1000
    else:
       timems = 10000 # default 10sec

#-----------------------------------------------------------------------------------------------------------------#      
def stoplogging():
    global readok
    readok = False  

def startlogging():
    global readok

    try:
      if entry_udt.get().strip():  # Check if not empty
         readok = True
         read_plc_u()
      if entry_Array.get().strip():  
         readok = True
         read_plc_a()
      if entry_string.get().strip():  
         readok = True
         read_plc_s()
      if multi_tag.get().strip():  
         readok = True
         plc_read_m()
    except Exception as e:
        msg.set(f"Is entry box empty? Error: {e}")         

#-------------------GUI ITEMS---------------------------------

ipalbl = Label(frame, text="IP Address: ", font='arial')
ipalbl.grid(row=0, column=0, padx=5, pady = 10, sticky='e')

Pathip = Entry(frame, width=20)
Pathip.grid(row=0, column=1, padx=52, pady=5, sticky='w')

Saveip = Button(frame,text="Save IP", pady=8, width=15, command=getip)
Saveip.grid(row=0, column=2, padx=0, sticky='w')


#Tags we are watching for data collection
entry_Array = Entry( frame5, width=70, bg='aqua' )
entry_Array.grid(row=0, column=1, padx=10, pady=30, sticky='nw')

e_arraylabel = Label(frame5, text='syntax: nameofarray{#of elements to read} example: testarray{5}', font= heading_font)
e_arraylabel.grid(row=0, column=1, padx=10, pady=5, sticky='nw')

multi_tag = Entry( frame5, width=70, bg='aqua' )
multi_tag.grid(row=3, column=1, padx=10, pady=25, sticky='w')

multilbl = Label(frame5, text="syntax: tag1, tag2, tag3, etc....", font=heading_font)
multilbl.grid(row=3, column=1, padx=10, pady=0, sticky='nw')

entry_udt = Entry( frame5, width=70, bg='aqua' )
entry_udt.grid(row=2, column=1, padx=10, pady=30, sticky='w')

entry_string = Entry(frame5, width=70, bg='aqua')
entry_string.grid(row=1, column=1, padx=10, pady=25, sticky='w')

stringlbl = Label(frame5, text="Enter a single tag", font=heading_font)
stringlbl.grid(row=1, column=1, padx=10, sticky='nw')

gettaglbl = Label(frame3, text="Global PLC Tag List", fg='blue')
gettaglbl.grid(row=0, column=0, padx=1, pady=5, sticky='n')

#watchtaglbl = Label(frame3, text='Watching Tags', fg='blue')
#watchtaglbl.grid(row=0, column=2, padx=170, sticky='sw')

text_box_gettags = Text(frame2, height=20, width=30, bg='beige')
text_box_gettags.grid(row=0, column=0, columnspan=1, padx=20, pady=10, sticky='w')

# GET ALL plc tags
gettagbtn = Button(frame3, width=15, height=2, text='Get PLC Tags', command=discoverplc)
gettagbtn.grid(row=1, column=0, padx=90, pady=0, sticky='s')


# Button to move tag from get to watch
array_button = Button(frame5, width=12, height=2, text="Read Array", command=read_plc_a)
array_button.grid(row=0, column=0, padx=5, pady=10, sticky='w')

string_button = Button(frame5, width=12, height=2, text="Read String", command=read_plc_s)
string_button.grid(row=1, column=0, padx=5, pady=0, sticky='w')

udt_button = Button(frame5, width=12, height=2, text="UDT", command=startlogging)
udt_button.grid(row=2, column=0, padx=5, pady=20, sticky='w')

mult_button = Button(frame5, width=12, height=2, text="Read Multiple", command=plc_read_m)
mult_button.grid(row=3, column=0, padx=5, pady=5, sticky='w')

errorlbl = Label(frame, textvariable=msg, width=90, fg='Red', font='arial')
errorlbl.grid(row=0, column=3, padx=20, sticky='w')

stoplogxl = Button(frame3, width=15, height=2, text="Stop Logging", command=stoplogging)
stoplogxl.grid(row=1, column=3, padx=50, sticky='w')

readlinelbl = Label(frame4, text='array to read')
readlinelbl.grid(row=0, column=0, padx=10, sticky='w')
#readlinelbl.config(text=", ".join(map(str, arraylist)))  # Convert list items to a string and print to gui

timeentryms = Entry(frame1, width=20)
timeentryms.grid(row=0, column=1, padx=2, sticky='w')

timelbl = Label(frame1, text="Interval in seconds: ", font='arial')
timelbl.grid(row=0, column=0,padx=5,sticky='w')

#------------------------------------------------------------------------------)

win.mainloop()



