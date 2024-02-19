# -*- coding: utf-8 -*-
"""
Created on Mon Jul 19 10:31:06 2021

@author: GanesVe1
"""

import tkinter as tk #for using tkinter widgets like buttons and indicators
from tkinter import ttk #for combobox and progress bar
from tkinter import messagebox #for message box
from reportlab.lib.pagesizes import letter #for test report
from reportlab.pdfgen import canvas #for lines and writing in pdf
from datetime import datetime #for current date and time
import os #for force exit in a function
from tkinter import filedialog #for opening a file

root = tk.Tk()

now1 = datetime.now()
dot_string1 = now1.strftime("%d-%m-%y--%H-%M-%S")
f= open('C:/526-A16693-001/events/'+dot_string1+".txt","a+")

now2 =now1.strftime("%H:%M:%S")
f.write("the program is started    " + now2 + "\n")

import xlwings as xw #to automate excel sheet
root.title("520-A16693-001 REV A") 

wb = xw.Book('INTERMEDIATE.xlsx') #intermediate excel sheet
sht1 = wb.sheets['Sheet1'] 

we = tk.Label(root,text='GUI to automate testing for EtherCAT PCBAs/controllers',font=("Helvetica", 14),padx=5,pady=5)
we.grid(row=0,column=0,columnspan=3, sticky ='n')

global t
t=0

def openWindow():
    global f
    if (f.close==True):
        f= open('C:/526-A16693-001/events/'+dot_string1+".txt","a+")
    print("user/serial number is selected")
    f.write("user/serial number is selected\n") #for event log purpose
    newWindow = tk.Toplevel(root) #sub UI from the parent UI
    newWindow.geometry("220x100") #size for sub UI
    newWindow.title("Entry") #title
    lab = tk.Label(newWindow, text = "Part Number") #labelling for part number
    lab.grid(row=0,column=0)
    lab1 = tk.Label(newWindow, text = "User Name")#labelling for user name
    lab1.grid(row=1,column=0)
    lab2 = tk.Label(newWindow, text = "Serial Number") #labelling for serial number
    lab2.grid(row=2,column=0)
    ent0 = ttk.Combobox(newWindow, width = 17) #selecting the part number using combobox
    ent0['values']= ('853-A16693-001',
                    '853-A16693-002') #values needed in the part number
    ent0.grid(row=0,column=1)
    ent0.current()
    ent1 = tk.Entry(newWindow) #entry for user name
    ent1.grid(row=1,column=1)
    ent2 = tk.Entry(newWindow) #entry for serial number
    ent2.grid(row=2,column=1)
    
    def window():
        print("part no., serial no., are entered and saved")
        f.write("part no., serial no., are entered and saved \n") #eventlog entry
        global part
        part = ent1.get() #variable for user name
        global serial 
        serial = ent2.get() #variable for serial number
        global port
        port = ent0.get() #variale for port number
        if (port == '853-A16693-002'): #selection part number 2
            for n in btns[72:96]: 
                n.configure(bg='rosy brown')
                n.configure(state='disabled')
                global t
                t=1
            for m in ent[72:96]:
                m.configure(bg='rosy brown')
                m.configure(state = 'disabled')               
            for k in gop[9:13]:
                k.configure(state = 'disabled')    
                #t=1
            for l in my_entries[9:13]:
                l.configure(state = 'disabled')   
            messagebox.showwarning("showwarning", " 72:95 DI/DO ARE DISABLED")
        else: #selection of part number 1
             for n in btns[72:96]:
                 n.configure(state='normal')
                 n.configure(bg='white')
                 t=0
             for m in ent[72:96]:
                 m.configure(state = 'normal')
                 m.configure(bg='white')
             for k in gop[9:13]:
                 k.configure(state = 'normal')
                 #t=0
             for l in my_entries[9:13]:
                 l.configure(state = 'normal')
        newWindow.destroy()
           
    enter = tk.Button(newWindow, text = "Enter", command =window) #button relating the command
    enter.grid(row=3,column=1)

User = tk.Button(root, text="User/Serial Number", width = 15,padx=5,pady=5,bg='gold',command = openWindow)
User.grid (row=0,column=0, sticky = 'n' + 'w',padx = 50, pady=50)
global host
def manual():
    global host
    host=1
    print("manual mode is selected")
    f.write("manual mode is selected\n") #for eventlog update
    messagebox.showinfo("showinfo", "Manual Mode: Please control maually")

etr=tk.IntVar()
etr.set(0)

chk1=tk.Checkbutton(root,text='Manual',command = manual, var=etr)
chk1.grid(row=0,column=0, sticky='n' + 'e',padx = 65, pady=40)

running = False
counter = 1


def auto():
    global host
    host=0
    print("automatic mode is selected")
    f.write("automatic mode is selected\n")
    def progress():
        global prog
        prog = tk.Toplevel(root)
        prog.geometry('210x50')
        prog.title('running')
        tk.Label(prog, text="Test in Progress...").pack()
        progres=ttk.Progressbar(prog,orient='horizontal',length=100,mode='determinate')
        progres.pack()
        global saw
        saw =0
        def seesaw():
            global saw
            progres['value']=saw
            prog.update_idletasks()
            saw = saw + 8
            root.after(1000,seesaw)
        seesaw()
        
        def _end():
            prog.destroy()
        prog.after(18000, _end)
        
        
    def report():
        
        if(sht2.range('I2').value=='PASS'):
            messagebox.showinfo("showinfo", "ALL TEST CASES PASSED")
        else:
            ADDIN=''
            for vbn in range(2,27):
                if (sht2.range('H'+str(vbn)).value=='FAIL'):
                    ADDIN = ADDIN + '\n' + sht2.range('E'+str(vbn)).value +' ' 
            messagebox.showerror("showerror", "FAILED\nPLEASE CHECK "+ADDIN )
        now = datetime.now() #current date and time
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        dot_string = now.strftime("%d-%m-%y--%H-%M-%S")
        can = canvas.Canvas('C:/526-A16693-001/results/'+port+'-'+ serial+'--'+ dot_string +'.pdf', pagesize=letter) #file name
        can.setLineWidth(.3) #line width in the PDF
        can.setFont('Helvetica-Bold', 24) #font size
        can.drawString(200,750,'TEST REPORT') #heading
        can.line(200,747,370,747) #underline the test report heading
                
        can.setFont('Helvetica', 12)
        can.drawString(40,700,'PN : ' + port) #part number from GUI
        can.drawString(40,675,'SN : ' + serial) #serial number
        can.drawString(390,700,"DATE&TIME: " + dt_string) #date and time
        can.drawString(390,675,"USER NAME: " + part) #user name
        can.line(30,640,580,640)
                
        can.setFont('Helvetica-Bold', 12)
        can.drawString(170,610,'TEST CASES')
        can.drawString(410,610,'OVERALL TEST RESULT')
        can.line(30,590,580,590) 
        can.line(30,640,30,80)
        can.line(580,640,580,80)
        can.line(390,640,390,80)
        can.line(30,80,580,80)
        #tabular column
        can.line(30,565,390,565)
        can.line(30,545,390,545)
        can.line(30,525,390,525)
        can.line(30,505,390,505)
        can.line(30,485,390,485)
        can.line(30,465,390,465)
        can.line(30,445,390,445)
        can.line(30,425,390,425)
        can.line(30,405,390,405)
        can.line(30,385,390,385)
        can.line(30,365,390,365)
        can.line(30,345,390,345)
        can.line(30,325,390,325)
        can.line(30,305,390,305)
        can.line(30,285,390,285)
        can.line(30,265,390,265)
        can.line(30,245,390,245)
        can.line(30,225,390,225)
        can.line(30,205,390,205)
        can.line(30,185,390,185)
        can.line(30,165,390,165)
        can.line(30,145,390,145)
        can.line(30,125,390,125)
        can.line(30,105,390,105)
        

        
        can.setFont('Helvetica', 12)
        can.drawString(40,570, 'CASE 1')
        if (sht2.range('H2').value=='FAIL'): #if pass then green color fg or red color fg
            can.setFillColor('red')
        else :
            can.setFillColor('green')
        can.drawString(330,570,': ' +sht2.range('H2').value)
        can.setFillColor('black')
        can.drawString(40,550, 'CASE 2')
        if (sht2.range('H3').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green')       
        can.drawString(330,550, ': ' +sht2.range('H3').value)
        can.setFillColor('black')
        can.drawString(40,530, 'CASE 3')
        if (sht2.range('H4').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,530,': ' +sht2.range('H4').value)
        can.setFillColor('black')
        can.drawString(40,510, 'CASE 4')
        if (sht2.range('H5').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,510,': ' +sht2.range('H5').value)
        can.setFillColor('black')
        can.drawString(40,490, 'CASE 5')
        if (sht2.range('H6').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,490,': ' +sht2.range('H6').value)
        can.setFillColor('black')
        can.drawString(40,470, 'CASE 6')
        if (sht2.range('H7').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,470,': ' +sht2.range('H7').value) 
        can.setFillColor('black')
        can.drawString(40,450, 'CASE 7')
        if (sht2.range('H8').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,450,': ' +sht2.range('H8').value)
        can.setFillColor('black')
        can.drawString(40,430, 'CASE 8')
        if (sht2.range('H9').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,430,': ' +sht2.range('H9').value)
        can.setFillColor('black')
        can.drawString(40,410, 'CASE 9')
        if (sht2.range('H10').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,410,': ' +sht2.range('H10').value)
        can.setFillColor('black')
        can.drawString(40,390, 'CASE10')
        if (sht2.range('H11').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,390,': ' +sht2.range('H11').value)
        can.setFillColor('black')
        can.drawString(40,370, 'CASE11')
        if (sht2.range('H12').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,370,': ' +sht2.range('H12').value)
        can.setFillColor('black')
        can.drawString(40,350, 'CASE12')
        if (sht2.range('H13').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(330,350,': ' +sht2.range('H13').value)
        if (sht2.range('I2').value=='FAIL'):
            can.setFillColor('red')
        else :
            can.setFillColor('green') 
        can.drawString(470,570,sht2.range('I2').value)
        can.setFillColor('black')
        can.drawString(40,330, 'CASE13 ')
        can.drawString(40,310, 'CASE14 ')
        can.drawString(40,290, 'CASE15 ')
        can.drawString(40,270, 'CASE16 ')
        can.drawString(40,250, 'CASE17 ')
        can.drawString(40,230, 'CASE18 ')
        can.drawString(40,210, 'CASE19 ')
        can.drawString(40,190, 'CASE20 ')
        can.drawString(40,170, 'CASE21 ')
        can.drawString(40,150, 'CASE22 ')
        can.drawString(40,130, 'CASE23 ')
        can.drawString(40,110, 'CASE24 ')
        can.drawString(40,90, 'CASE25 ')
        kyt=330
        vyt=14
        while (kyt>70):
            if (sht2.range('H' + str(vyt)).value=='FAIL'):
                can.setFillColor('red')
            else :
                can.setFillColor('green') 
            can.drawString(330,kyt, ': ' + sht2.range('H' +str(vyt)).value)
            kyt=kyt-20
            vyt=vyt+1
        can.setFillColor('black')    
        
        can.setFont('Helvetica', 8)
        can.drawString(80,570,'      ('+sht2.range('G2').value+')')
        can.drawString(80,550,'      ('+sht2.range('G3').value+')')
        can.drawString(80,530,'      ('+sht2.range('G4').value+')')
        can.drawString(80,510,'      ('+sht2.range('G5').value+')')
        can.drawString(80,490,'      ('+sht2.range('G6').value+')')
        can.drawString(80,470,'      ('+sht2.range('G7').value+')') 
        can.drawString(80,450,'      ('+sht2.range('G8').value+')')
        can.drawString(80,430,'      ('+sht2.range('G9').value+')')
        can.drawString(80,410,'      ('+sht2.range('G10').value+')')
        can.drawString(80,390,'      ('+sht2.range('G11').value+')')
        can.drawString(80,370,'      ('+sht2.range('G12').value+')')
        can.drawString(80,350,'      ('+sht2.range('G13').value+')')
        can.drawString(80,330,'      ('+sht2.range('G14').value+')')
        can.drawString(80,310,'      ('+sht2.range('G15').value+')')
        can.drawString(80,290,'      ('+sht2.range('G16').value+')')
        can.drawString(80,270,'      ('+sht2.range('G17').value+')')
        can.drawString(80,250,'      ('+sht2.range('G18').value+')')
        can.drawString(80,230,'      ('+sht2.range('G19').value+')')
        can.drawString(80,210,'      ('+sht2.range('G20').value+')')
        can.drawString(80,190,'      ('+sht2.range('G21').value+')')
        can.drawString(80,170,'      ('+sht2.range('G22').value +')')
        can.drawString(80,150,'      ('+sht2.range('G23').value+')')
        can.drawString(80,130,'      ('+sht2.range('G24').value+')')
        can.drawString(80,110,'      ('+sht2.range('G25').value+')')
        can.drawString(80,90,'       ('+sht2.range('G26').value+')' )
            

        can.setFont('Helvetica-Bold', 12)
        can.drawString(450,35,'SOFTWARE PN & REV') 
        can.drawString(450,20,'520-A16693-001 REV A')
        can.save() #save the pdf
        f.write("test report is generated\n")
        f.write("the program is stoped    " + now2 + "\n")
        f.close()
        wb.close()
        wb1.close()
        path = 'C:/526-A16693-001/results/' +port+'-'+serial+'--'+ dot_string +'.pdf'
        os.system(path) #to open the pdf
        root.destroy()

        
        
    global num # calling the input for  DI's automatic test from input excel file 
    num = 2 
    global ryt #indexing
    ryt=0
    global jt #indexing
    jt=-1
    wb1 = xw.Book('INPUT.xlsx') #input for automatic mode
    sht2 = wb1.sheets['Sheet1'] 
    checkvar = tk.IntVar(value=1) #variable for selecting checkbotton automatically
    checkvar2 = tk.IntVar(value=0) #variable for deselecting checkbotton automatically
    global xq #indexing for checkbuttons
    global abc #extracting input for AO's from excel
    abc=17
    global dfe #storing the AO values in the intermediate excel file
    dfe = 2
    xq = 7
    global ghi #indexing for AO list
    ghi = 0
    def count_two():
        global num # calling the input for  DI's automatic test from input excel file
        global xq # indexing for checkbuttons
        global abc,dfe,ghi #using the global variable
        if running:
            global counter
            counter += 1
            fgt = tk.StringVar() #varaible for AO's
            #setting the varaible value from excel input for automation
            fgt.set(str(sht2.range('A'+ str(abc)).options(numbers=int).value))
            #storing the value in the list
            entry[ghi].configure(textvariable = fgt)
            #check whether the analog value is between 0-10V
            if(int(fgt.get())<=10 and int(fgt.get())>=0): 
                sht1.range('B' + str(dfe)).value = fgt.get() 
                sht2.range('B' + str(abc)).value = fgt.get()
            else: #any value out of range
                sht1.range('B' + str(dfe)).value = 'invalid'
                sht2.range('B' + str(abc)).value = 'invalid'
            if(ghi!=0): #erase the previous test case
                entry[ghi-1].delete(0,'end') 
                sht1.range('B' + str(dfe-1)).value = 0
            abc=abc+1
            ghi=ghi+1
            dfe=dfe+1
            input1=str(sht2.range('A' + str(num)).options(numbers=int).value) #input for DI's
            if(input1!='None'): 
                global ryt #indexing
                gop[ryt].configure(text=input1) 
                ryt = ryt + 1
                nyt=int(str(input1), 16) #converting into hexadecimal
                byt=''
                ayt = 1
                #the logic for converting hexadecimal into binary is the same as manual mode
                while nyt > 0:
                    global jt
                    jt = jt+1
                    byt = str(nyt % 2)
                    if(byt=='1'):                    
                        btns[xq-jt].configure(text=byt)
                        btns[xq-jt].configure(variable = checkvar)
                        l1[xq-jt]=1
                        save()
                    else:                   
                        btns[xq-jt].configure(text=byt)
                        btns[xq-jt].configure(variable = checkvar2)
                        l1[xq-jt]=0
                        save()
                    if (xq-jt>7): #for clearing the previous case
                        btns[xq-jt-8].configure(text='0')
                        btns[xq-jt-8].configure(variable = checkvar2)
                        l1[xq-jt-8]=0
                        save()
                    nyt = nyt >> 1 
                    ayt = ayt + 1
                    
                xq = xq + 16
                while ayt < 9:
                    jt = jt + 1
                    btns[xq-jt].configure(text='0')
                    btns[xq-jt].configure(variable = checkvar2)
                    l1[xq-jt]=0
                    ayt = ayt + 1 
                sht2.range('B'+ str(num)).value = input1
                num=num+1
                    
                root.after(1000,count_two) #continue this loop after every 10 seconds
            else: #for clearing the values once all the test cases runs
                entry[ghi-1].delete(0,'end') #analog inputs
                sht1.range('B' + str(dfe-1)).value = 0 #digital inputs
                for wr in gop:
                    wr.configure(text='0x0')
                grt=0
                for gr in btns :
                    gr.configure(variable = checkvar2)
                    gr.configure(text = '0')
                    l1[grt]=0
                    grt+=1
                    save()
                root.after (1000,report)
                        
    def _start():
        print("start is selected")
        f.write("start is selected\n")
        global running
        running=True #to run the program until stop is presssed
        count_two() #calling the function for automatic testing of cases
                
    
    def _end():
        print("stop is selected")
        f.write("stop is selected\n")
        global running  #stop the program if it is running
        running = False
        global counter 
        counter=1
        prog.destroy()
        for wr in gop: #to make the buttons zero
            wr.configure(text='0x0')
        grt=0
        for gr in btns : #to make the checkbuttons zero
            gr.configure(variable = checkvar2)
            gr.configure(text = '0')
            l1[grt]=0
            grt+=1
            save() #calls the funtion save in order to make the DI's zero
        for kop in entry:
            kop.delete(0,'END')
    
    
    start = tk.Button(root, text="Start the Test", width = 15,padx=5,pady=5,bg='gold', 
                      command =lambda:[ progress(),_start()])
    start.grid (row=0,column=0,columnspan=2, sticky = 'n' +'e', pady=50, padx =240)
    
    stop = tk.Button(root, text="Stop the test", width = 15,padx=5,pady=5,bg='gold', command = _end)
    stop.grid (row=0,column=1, sticky = 'n' + 'e',padx = 50, pady=50)
    
    messagebox.showinfo("showinfo", "Automatic Mode : User will not have any control")

etr1=tk.IntVar()
etr1.set(0)
    

chk1=tk.Checkbutton(root,text='Automatic',command = auto,var=etr1 )
chk1.grid(row=0,column=0, sticky='n' + 'e',padx = 50, pady=60)

def phot():
    print('test report button is selected')
    f.write("test report button is selected\n")
    filename=filedialog.askopenfilename(initialdir = "C:/526-A16693-001/results")
    import webbrowser
    webbrowser.open_new(filename)

test = tk.Button(root, text="Test Report", width = 15,padx=5,pady=5,bg='gold',command =phot)
test.grid (row=0,column=2, sticky = 'n'+'e', pady=80,padx=150 )

    
DO = tk.LabelFrame(root,text='DO',bg='grey',padx=5,pady=5)
DO.grid(row=0,column=1, sticky = 'n', pady =100,padx=15)
    
c = 0 #for alternate rows
d = 0 #for printing DO0 to DO95 above the checkbuttons
while (c < 24):
    for b in range(9,17): #as first 8 columns are used by DI's, next 8 columns for DO's
        tk.Label(DO,text = "DO " + str(d),bg='grey').grid(row=c, column=b)
        d = d + 1
    c = c + 2
 
#for manual use of check buttons
#selecting or deselecting the buttons     
def action(button):
    global host
    if (host==1):
        f.write('DO'+ str(button) +' is and read back by DI' +str(button)+ ' \n')#for updating in eventlog condition
        l1[button] = 1
    if (btns[button].cget('text') == '0'):
        btns[button].configure(text="1")
        save()    
    else:
        btns[button].configure(text='0')
        l1[button] = 0
        save()
  
btn_nr = -1
btns = [] #list for storing the checkbutton
l1 = [0]*96 #list for storing binary values 
z=1 #alternatively placing the checkbuttons 
while (z < 24):
    for y in range(9, 17):
        btn_nr += 1 #indexing in the list
        btns.append(tk.Checkbutton(DO,text='0',width=1,bg='white',command=lambda z=btn_nr: action(z)))
        #defining the checkbuttons
        btns[btn_nr].grid(row=z,column=y,padx=2)
    z = z + 2

gop=[] #empty list storing hexadecimal values
  
def save():#for converting binary to hexadecimal value
    #calling 8 values at a time in the binary list 
    u=0 #initial value of binary list
    v=7 #final value of a hexadecimal number in a binary list
    r=1 #alternate rows of hexadecimal buttons
    s=2 #for storing the hexadecimal value in a column in the excel sheet 
    while (v<96):
        binary =""
        for i in (l1[u:v+1]): #for appending 8 binary values to convert into single hexadecimal value 
            binary += str(i) #appending the binary values list in a string
        decimal = int(binary, 2) #converting the binary string into integer
        hexadecimal = hex(decimal) #convert the integer value into hexadecimal value
        #storing the hexadecimal value in the button
        DO1=tk.Button(DO,width=5, text=hexadecimal,bg='skyblue')
        DO1.grid(row=r,column=17,padx=2)
        gop.append(DO1) #storing the hexadecimal value in a list
        sht1.range('A'+ str(s)).value = hexadecimal #storing the hexadecimal values in excel sheet
        u = u + 8
        v = v + 8
        r = r + 2
        s = s + 1
save()#calling the funtion for the first time in order to create the buttons(blue buttons) 

AO = tk.LabelFrame(root,text='AO',bg='grey',padx=5,pady=5)
AO.grid(row=0,column=2,sticky='n',pady =150,padx=15)


def callback(event): #space bar event to occur
    g=2
    for h in entry:
        if(int(h.get())<=10 and int(h.get())>=0): #for values in between 0-10V
            sht1.range('B'+ str(g)).value = h.get() 
            if(host==1):
                f.write('AO' +str(entry.index(h))+' is entered and read back by AI'+str(entry.index(h))+'\n')#for event log
        else: #for values not between 0-10V
            sht1.range('B'+ str(g)).value = 'invalid'
            if(host==1):
                f.write('AO' +str(entry.index(h))+' is entered and read back by AI'+str(entry.index(h))+'\n')#for event log
        g = g + 1
        
entry= []#list for AO entry
c=0 #alternating
d=0 #printing AO0 to AO31
while (c < 8):
    for b in range(18,26): 
        tk.Label(AO,text = "AO " + str(d),bg='grey').grid(row=c, column=b)
        AO1 = tk.Entry(AO, width=5, borderwidth=7) #creating entry
        AO1.grid(row=c+1, column=b) 
        entry.append(AO1) #appending into list
        AO1.bind("<Return>",callback) #after scape bar the value is stored
        d = d + 1
    c = c + 2

AI = tk.LabelFrame(root,text='AI',bg='grey',padx=5,pady=5)
AI.grid(row=0,column=2,sticky='e',pady = 445,padx=25)

entries=[]#list to store the AI buttons
c=0 #alternating rows for placing label
d=0 #displaying AI0 to AI31
i=-1 #index of the list
s=2 #to call a cell values in a column in excel

while (c < 8): 
    for b in range(27,35): #next 8 columns in GUI
        i += 1 
        tk.Label(AI,text = "AI " + str(d),bg='grey').grid(row=c, column=b) #labelling
        data = sht1.range('C'+ str(s)).options(numbers =int).value #retreving data from excel
        entries.append(tk.Button(AI,width=4,padx=3, text =data)) #appending in list
        entries[i].grid(row=c+1, column=b,padx=2)  
        s = s + 1
        d = d + 1
    c = c + 2
    
def update(): #printing from the list to button
    i = 0 #indexing with the list
    s = 2 #cells in the excel
    for i in entries:
        rev = sht1.range('C'+ str(s)).options().value 
        if (rev=='invalid'): #if the voltage value is not between 0-10 V
            i.configure(text = 'invalid')
            i.configure(bg='sandy brown')

        else : #if voltage is between 0-10 V
            i.configure(text = str(int(rev))+'V')
            i.configure(bg='white')
        s = s+1
    root.after(1000, update)
    
update() #calling the function first time for creating buttons

DI = tk.LabelFrame(root,text='DI',bg='grey',padx=5,pady=5)
DI.grid(row=0,column=0, sticky='n',pady =100,padx=5)

#as its a 2D array variable 'b' and 'c' are used define the rows and columns
b = 0 #for printing 8 label of radio button in the same row
c = 0 #the label should be printed in alternate row
d = 0 #for printing DI0 to DI95
while (c < 24):#as there 96 DO's and 8 DO'S in each row so we need 12 rows in total (alternatively)
    for b in range(8): # 8 DO's in each row
        tk.Label(DI,text = "DI " + str(d),bg='grey').grid(row=c, column=b)#defining a label
        d = d + 1
    c = c + 2 #printing the labels alternatively

my_entries=[]#list for storing the button which displays the hexadecimal value in the button
i,w,q,o,j = 1,1,1,-1,-1
ent = []#list for storing the 96 radio buttons
box = tk.BooleanVar() #defining the boolean variable for radiobuttion to select and deselect
box.set(True)#deselection the radio button by default
for i in range (2,14):#calling the D column from excel sheet for hexadecimal value
    ini_string = str(sht1.range('D'+ str(i)).options(numbers=int).value)
    if (ini_string=="None"): 
        break 
    else:
        o+=1 # for list indexes
        my_entries.append(tk.Button(DI,width=5, text=ini_string,bg='pink'))#defining hexadecimal buttons
        my_entries[o].grid(row=w, column=8,padx=2)
        w= w+2 #alternate rows of buttons
        a = 1
        x = 7
        n = int(ini_string, 16) #converting hexadecimal in decimal
        bStr = '' #empty string for storing the converted hexadecimal to binary
        while n > 0: 
            j+=1 #index for radiobuttons list
            bStr = str(n % 2)#decimal number divided bt 2 (prime factorisation for binary)
            ent.append(tk.Radiobutton(DI,text=bStr,var=box,width=1))#assigning the value to radiobutton
            ent[j].grid(row=q,column=x,padx=2)
            x=x-1 #prime fatorisation should be reversed in order
            n = n >> 1 #binary shifting
            a = a + 1 #in order to add zeros when there are no 8 binary numbers
        while a < 9:
            j+=1
            ent.append(tk.Radiobutton(DI,text='0',var=box,width=1))
            ent[j].grid(row=q,column=x,padx=2)
            x=x-1
            a = a + 1
        q = q + 2 #printing radiobuttons in alternate rows



 
def update1():
    i = 0
    s = 2
    w = 1
    q = 1
    j=-1
    for i in my_entries:
        i.configure(text = sht1.range('D'+ str(s)).options(numbers =int).value )
        data = sht1.range('D'+ str(s)).options(numbers =int).value
        s= s+1
        w= w+2
        a = 1
        x = 7
        n = int(str(data), 16) 
        bStr = ''
        while n > 0:
            j+=1
            bStr = str(n % 2)
            if(bStr=='1'):
                ent[j].configure(text=bStr)
                ent[j].configure(selectcolor = 'green')
                ent[j].configure(bg = 'light green')
            else:
                ent[j].configure(text=bStr)
                ent[j].configure(selectcolor = 'white')
                ent[j].configure(bg = 'white')
            x=x-1
            n = n >> 1 
            a = a + 1
        while a < 9:
            j+=1
            ent[j].configure(text='0')
            if (t==1):
                if (j>71):
                    ent[j].configure(bg='rosy brown')                
            else:
                ent[j].configure(selectcolor = 'white')
                ent[j].configure(bg = 'white')
            x=x-1
            a = a + 1
        q = q + 2            
    root.after(1000, update1)
    
    
update1()

root.mainloop()