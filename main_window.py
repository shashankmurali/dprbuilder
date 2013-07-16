import Tkinter as Tk
import MySQLdb as mdb
import xlwt
import xlrd
from xlutils.copy import copy
import sys


class OtherFrame(Tk.Toplevel):
    def __init__(self, original,value,district):

        self.original_frame = original
        Tk.Toplevel.__init__(self)
        self.geometry("900x600")
        self.title(value + " - " + district)

        # A canvas to add the scroll bar.
        self.canvas = Tk.Canvas(self, borderwidth=0)
        self.frame = Tk.Frame(self.canvas)
        self.vsb = Tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((4,4), window=self.frame, anchor="nw", 
                                  tags="self.frame")

        self.frame.bind("<Configure>", self.OnFrameConfigure)

        self.populate(value,district)

    def populate(self,table_name,district):
        #Labels for each of the entry fields
        code_label = Tk.Label(self.frame,text="Code")
        code_label.grid(row=0,column=0)
        code_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="red")
        
        description_label = Tk.Label(self.frame,text="Description")
        description_label.grid(row=0,column=1)
        description_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="red")

        unit_label = Tk.Label(self.frame,text="Unit")
        unit_label.grid(row=0,column=2)
        unit_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="red")

        quantity_label = Tk.Label(self.frame,text="Quantity")
        quantity_label.grid(row=0,column=3)
        quantity_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="red")

        rate_label = Tk.Label(self.frame,text="Rate")
        rate_label.grid(row=0,column=4)
        rate_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="red")

        amount_label = Tk.Label(self.frame,text="Amount")
        amount_label.grid(row=0,column=5)
        amount_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="red")

        #Entry fields - Set one
        self.code_entry_1 = Tk.Entry(self.frame,width = 10)
        self.code_entry_1.grid(row=1,column=0)
                
        self.description_entry_1 = Tk.Text(self.frame,width = 35,height=5)
        self.description_entry_1.grid(row=1,column=1)
        self.description_entry_1.config(wrap=Tk.WORD)
             
        self.unit_entry_1 = Tk.Entry(self.frame,width = 10)
        self.unit_entry_1.grid(row=1,column=2)

        self.quantity_entry_1 = Tk.Entry(self.frame,width = 10)
        self.quantity_entry_1.grid(row=1,column=3)

        self.rate_entry_1 = Tk.Entry(self.frame,width = 10)
        self.rate_entry_1.grid(row=1,column=4)

        self.amount_entry_1 = Tk.Entry(self.frame,width = 10)
        self.amount_entry_1.grid(row=1,column=5)

        #Upon hitting return after entering the code, this below line calls the fill_entries function which populates other entry and text fields
        self.code_entry_1.bind('<Return>',lambda event, code_entry = self.code_entry_1,description_entry = self.description_entry_1,
                               unit_entry = self.unit_entry_1,quantity_entry = self.quantity_entry_1,rate_entry = self.rate_entry_1,
                               amount_entry = self.amount_entry_1,table_name = table_name,
                               district=district: self.fill_entries(event,code_entry,description_entry,unit_entry,quantity_entry,rate_entry,
                                                                    amount_entry,table_name,district))

        #Entry fields - Set two
        self.code_entry_2 = Tk.Entry(self.frame,width = 10)
        self.code_entry_2.grid(row=2,column=0)
        
        self.description_entry_2 = Tk.Text(self.frame,width = 35,height=5)
        self.description_entry_2.grid(row=2,column=1)
        self.description_entry_2.config(wrap=Tk.WORD)

        self.unit_entry_2 = Tk.Entry(self.frame,width = 10)
        self.unit_entry_2.grid(row=2,column=2)

        self.quantity_entry_2 = Tk.Entry(self.frame,width = 10)
        self.quantity_entry_2.grid(row=2,column=3)

        self.rate_entry_2 = Tk.Entry(self.frame,width = 10)
        self.rate_entry_2.grid(row=2,column=4)

        self.amount_entry_2 = Tk.Entry(self.frame,width = 10)
        self.amount_entry_2.grid(row=2,column=5)

        self.code_entry_2.bind('<Return>',lambda event, code_entry = self.code_entry_2,description_entry = self.description_entry_2,
                               unit_entry = self.unit_entry_2,quantity_entry = self.quantity_entry_2,rate_entry = self.rate_entry_2,
                               amount_entry = self.amount_entry_2,table_name = table_name,
                               district = district: self.fill_entries(event,code_entry,description_entry,unit_entry,quantity_entry,rate_entry,
                                                                      amount_entry,table_name,district))

        #list of variables for code entry values
        code = []
        code.append(self.code_entry_1)
        code.append(self.code_entry_2)

        #list of variables for description entry values     
        description = []
        description.append(self.description_entry_1)
        description.append(self.description_entry_2)

        #list of variables for unit entry values
        unit = []
        unit.append(self.unit_entry_1)
        unit.append(self.unit_entry_2)

        #list of variables for quantity entry values
        quantity = []
        quantity.append(self.quantity_entry_1)
        quantity.append(self.quantity_entry_2)
        
        #list of variables for rate entry values
        rate = []
        rate.append(self.rate_entry_1)
        rate.append(self.rate_entry_2)

        #list of variables for amount entry values
        amount = []
        amount.append(self.amount_entry_1)
        amount.append(self.amount_entry_2)
        
        #Button to add a new row
        self.num_rows=2
        add_row_btn = Tk.Button(self.frame, text="Add Row", command=lambda: self.add_new_row(code,description,unit,quantity,rate,amount,table_name,district))
        add_row_btn.grid(row=1,column=6)

        #Button to close this window and open the root window
        close_btn = Tk.Button(self.frame, text="Back", command=self.onClose)
        close_btn.grid(row=1,column=7)

        #Button to write all the data onto an excel doc
        submit_btn = Tk.Button(self.frame, text="Submit",command= lambda: self.openFrame(code,description,unit,quantity,rate,amount,table_name,self.num_rows
                                                                                         ,self.original_frame))
        submit_btn.grid(row=2,column=6)
        
    def OnFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    #----------------------------------------------------------------------
    #The below function takes in variables of different entry fields and is bound to the "Return Key" on the code entry field.
    #It then inserts data into other entry fields.

    def fill_entries(self,event,code_entry,description_entry,unit_entry,quantity_entry,rate_entry,amount_entry,table_name,district):
        rows = self.open_file(table_name)
        code_entered = code_entry.get()
        for i in range(0,len(rows)):
            count = 0
            if code_entered == rows[i][0]:
                description_entry.delete(1.0,Tk.END)
                description_entry.insert(Tk.END,rows[i][1])
                unit_entry.delete(0,Tk.END)
                unit_entry.insert(Tk.END,rows[i][2])
                rate_entry.delete(0,Tk.END)
                rate_entry.insert(Tk.END,rows[i][3])
                count = 1
                break
        if count==1:
            quantity_entry.bind('<Return>',lambda event, quantity_entry = quantity_entry, rate_entry = rate_entry,
                                    amount_entry = amount_entry,district = district: self.amount_calc(event,quantity_entry,rate_entry,amount_entry,district))
            
                    
                    
                     
    def amount_calc(self,event,quantity_entry,rate_entry,amount_entry,district):
        cost_index = {"Trivandrum":1.4295,"Pathanamthitta":1.3960,"Kottayam":1.3960,"Kollam":1.3826,"Aleppey":1.4698,"Munnar":1.4899,
                      "Thodupuzha,Koothattukulam & Manimalakunnu":1.3960,"Idukki & Nedumkandam":1.4832,"Ernakulam":1.4228,"Sai Punnamada":1.5101,
                      "Calicut":1.3087,"Trichur":1.2819,"Mahe":1.3087,"Kannur":1.3087,"Kasargod":1.3020,"Nadapuram":1.3087,"Palakkad":1.2752
                      ,"Malappuram":1.2953,"Wayanad":1.3087}
        quantity = quantity_entry.get()
        quantity = int(quantity)
        rate = rate_entry.get()
        amount = cost_index[district] * float(rate) * quantity
        amount_entry.delete(0,Tk.END)
        amount_entry.insert(Tk.END,amount)                         
                                        

    def open_file(self,table_name):
        f = open(sys.path[0]+'/'+'database_tables/'+table_name+'.txt')
        table_data=[]
        for line in f:
            line_split = line.split('\t')
            #print line_split
            a = line_split[3]
            #print line_split[0],a
            pos1 = a.find('\\')
            pos2 = a.find('\\r')
            #print pos2
            if pos2==None:
                small = pos1
            else:
                if pos1 < pos2:
                    small = pos1
                else:
                    small = pos2 - 1
            a = a[:small]
            line_split[3] = a
            table_data.append(line_split)
        return table_data
                    
    #----------------------------------------------------------------------
    #Connects to the database and returns a list of data retrieved from a table.

    def query(self,table_name):
        con = mdb.connect('localhost', 'shashank', 'shashank', 'dprdb');
        with con:
            cur = con.cursor()
            cur.execute("SELECT * FROM "+table_name)
            rows=[]
            for i in range(cur.rowcount):
                row = cur.fetchone()
                rows.append([row[0],row[1],row[2],str(row[3])])
        return rows
    #----------------------------------------------------------------------
    def onClose(self):
        """"""
        self.destroy()
        self.original_frame.show()

   
    #----------------------------------------------------------------------
    def openFrame(self,code,description,unit,quantity,rate,amount,table_name,num_rows,original_frame):
        """"""
        self.hide()
        subFrame = popupWindow(self,code,description,unit,quantity,rate,amount,table_name,self.num_rows,original_frame)

    def hide(self):
        """"""
        self.withdraw()

    def show(self):
        """"""
        self.update()
        self.deiconify()
        
    #----------------------------------------------------------------------
    def add_new_row(self,code,description,unit,quantity,rate,amount,table_name,district):
        #for code
        e = Tk.Entry(self.frame,width=10)
        e.grid(row=(self.num_rows+1),column=0)
        code.append(e)

        #for description
        e = Tk.Text(self.frame,width=35,height=5)
        e.grid(row=(self.num_rows+1),column=1)
        e.config(wrap=Tk.WORD)
        description.append(e)
        
        
        #for unit
        e = Tk.Entry(self.frame,width=10)
        e.grid(row=(self.num_rows+1),column=2)
        #unit[self.num_rows] = e
        unit.append(e) 
        
        #for quantity
        e = Tk.Entry(self.frame,width=10)
        e.grid(row=(self.num_rows+1),column=3)
        #quantity[self.num_rows] = e
        quantity.append(e)
        
        #for rate
        e = Tk.Entry(self.frame,width=10)
        e.grid(row=(self.num_rows+1),column=4)
        #rate[self.num_rows] = e
        rate.append(e)

        #for amount
        e = Tk.Entry(self.frame,width=10)
        e.grid(row=(self.num_rows+1),column=5)
        #amount[self.num_rows] = e
        amount.append(e)
        
        code[self.num_rows].bind('<Return>',lambda event, code_entry = code[self.num_rows],description_entry = description[self.num_rows],
                                 unit_entry = unit[self.num_rows], quantity_entry = quantity[self.num_rows],rate_entry = rate[self.num_rows],
                                 amount_entry = amount[self.num_rows],table_name = table_name,
                                 district=district: self.fill_entries(event,code_entry,description_entry,unit_entry,quantity_entry,
                                                                      rate_entry,amount_entry,table_name,district))
                                                                                                                 
        self.num_rows+= 1

class popupWindow(Tk.Toplevel):
    def __init__(self,parent,code,description,unit,quantity,rate,amount,table_name,num_rows,original_frame):
        self.original_frame = parent
        self.code = code
        self.description = description
        self.unit = unit
        self.quantity = quantity
        self.rate = rate
        self.amount = amount
        self.table_name = table_name
        self.num_rows = num_rows
        self.original_frame = original_frame
        Tk.Toplevel.__init__(self)
        self.geometry("300x100")
        self.title("FileName")
        #self.frame = Tk.Frame(parent)
        self.l=Tk.Label(self,text="Enter Filename")
        self.l.pack()
        a = Tk.StringVar()
        a.set(".xls")
        self.e=Tk.Entry(self)
        self.e.config(textvariable=a)
        self.e.pack()
        self.b=Tk.Button(self,text='Ok',command=self.onClose)
        self.b.pack()

    def onClose(self):
        """"""
        self.filename = self.e.get()
        self.write_doc()
        self.destroy()
        subFrame = successDialog(self,self.filename,self.table_name,self.original_frame)

    def hide(self):
        """"""
        self.withdraw()

    def write_doc(self):
        try:
            workbook = xlrd.open_workbook(sys.path[0]+'/'+self.filename)
            wb = copy(workbook)
            sheet1 = wb.add_sheet(self.table_name)
            sheet1.write(0, 0, "Code")
            sheet1.write(0, 1, "Description")
            sheet1.write(0, 2, "Unit")
            sheet1.write(0, 3, "Quantity")
            sheet1.write(0, 4, "Rate")
            sheet1.write(0, 5, "Amount")
            e=2
            for i in range(0,self.num_rows):
                sheet1.write(e,0,self.code[i].get())
                sheet1.write(e,1,self.description[i].get(1.0,Tk.END))
                sheet1.write(e,2,self.unit[i].get())
                sheet1.write(e,3,self.quantity[i].get())
                sheet1.write(e,4,self.rate[i].get())
                sheet1.write(e,5,self.amount[i].get())
                e = e+1
            wb.save(self.filename)

        except IOError:
            book = xlwt.Workbook(encoding="utf-8")
            sheet1 = book.add_sheet(self.table_name)
            sheet1.write(0, 0, "Code")
            sheet1.write(0, 1, "Description")
            sheet1.write(0, 2, "Unit")
            sheet1.write(0, 3, "Quantity")
            sheet1.write(0, 4, "Rate")
            sheet1.write(0, 5, "Amount")
            e=2
            for i in range(0,self.num_rows):
                sheet1.write(e,0,self.code[i].get())
                sheet1.write(e,1,self.description[i].get(1.0,Tk.END))
                sheet1.write(e,2,self.unit[i].get())
                sheet1.write(e,3,self.quantity[i].get())
                sheet1.write(e,4,self.rate[i].get())
                sheet1.write(e,5,self.amount[i].get())
                e = e+1
            book.save(self.filename)

              
        
class successDialog(Tk.Toplevel):
    def __init__(self,parent,filename,tablename,original_frame):
        self.filename = filename
        self.tablename = tablename
        self.original_frame = original_frame
        Tk.Toplevel.__init__(self)
        self.geometry("600x50")
        self.title("Success!")
        self.l=Tk.Label(self,text="Congratulations! Your DPR for " +
                        self.tablename + " is successfully saved onto the file "
                        + self.filename + ".")
        self.l.pack()

        close_btn = Tk.Button(self, text="Back", command=self.onClose)
        close_btn.pack()

    def onClose(self):
        """"""
        self.destroy()
        self.original_frame.show()
        

class MyApp(object):
    """"""
 
    #----------------------------------------------------------------------
    def __init__(self, parent):
        """Constructor"""
        self.root = parent
        self.root.title("Detailed Project Report Builder")
        self.frame = Tk.Frame(parent)
        self.frame.grid()

        #Drop Down Menu for selecting category
        var = Tk.StringVar(self.frame)
        var.set("Earth_Work")
        option = Tk.OptionMenu(self.frame, var,"Carriage_of_Materials", "Earth_Work", "Mortars", "Concrete_Work",
                               "Reinforced_Cement_Concrete","Brick_Work","Stone_Work",
                               "Marble_and_Granite Work","Wood_and_PVC Work","Steel_Work",
                               "Flooring","Roofing","Finishing","Repairs_to_Building",
                               "Dismantling_and_Demolishing", "Road_Work", "Sanitary_Installations",
                               "Water_Supply", "Drainage","Pile_Work","Aluminium_Work",
                               "Water_Proofing","Horticulture_and_Landscaping","Rain_Water_Harvesting_and_Tubewells",
                               "Conservation_of_Heritage_Buildings"
                               )
        option.grid(row=0,column=2,columnspan=3)

        dvar = Tk.StringVar(self.frame)
        dvar.set("Trivandrum")
        district = Tk.OptionMenu(self.frame,dvar,"Trivandrum","Pathanamthitta","Kottayam","Kollam",
                                 "Aleppey","Munnar","Thodupuzha,Koothattukulam & Manimalakunnu",
                                 "Idukki & Nedumkandam","Ernakulam","Sai Punnamada","Calicut",
                                 "Trichur","Mahe","Kannur","Kasargod","Nadapuram"
                                 ,"Palakkad","Malappuram","Wayanad")
        district.grid(row=1,column=2,columnspan=3)

        category_label = Tk.Label(self.frame,text="Select a category:")
        category_label.grid(row=0,column=1)
        category_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="blue")

        district_label = Tk.Label(self.frame,text="Select the district:")
        district_label.grid(row=1,column=1)
        district_label.config(font="Verdana",relief="groove",borderwidth="3",foreground="blue")

        #Button opens a new window for entering the actual DPR" 
        btn = Tk.Button(self.frame, text="Start", command=lambda: self.openFrame(option,district))
        btn.grid(row=2,column=3)
 
    #----------------------------------------------------------------------
    def hide(self):
        """"""
        self.root.withdraw()
 
    #----------------------------------------------------------------------
    def openFrame(self,option,district):
        """"""
        #print sys.path[0]
        option_value = option.cget("text")
        district_value = district.cget("text")
        self.hide()
        subFrame = OtherFrame(self,option_value,district_value)
 
    #----------------------------------------------------------------------
    def show(self):
        """"""
        self.root.update()
        self.root.deiconify()
 
#----------------------------------------------------------------------
if __name__ == "__main__":
    root = Tk.Tk()
    root.geometry("400x300")
    app = MyApp(root)
    root.mainloop()

