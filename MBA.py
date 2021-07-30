import winsound
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfilename
import csv
from tkinter import messagebox
import xlsxwriter as xl
import datetime
import os


desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
#print(type(desktop))

bg_lbl_login= 'black'
fg_lbl_login= 'white'

bg_lbl_home= 'lightblue'
fg_lbl_home= 'black'


duration = 300  # milliseconds
freq = 4000  # Hz
#winsound.Beep(freq, duration)


form=tk.Tk()

form.geometry('597x380+300+130')
form.resizable(False,False)
form.title('Market Basket Analysis')
form.iconbitmap('icon1.ico')


login_frame = Frame(form, bg='black' , width=30 )

#form.config(background='darkblue')

Frame(login_frame , height=80, width=80 ).grid(row=1 , column=0)


tk.Label(login_frame , text='Log in'  , bg = bg_lbl_login, fg=fg_lbl_login  , font=('consolas ', 35)  ,height=1 ).grid(row=0 ,column=0,columnspan=5  , padx=250 )
tk.Label(login_frame ,text='username:', bg=bg_lbl_login ,fg=fg_lbl_login , font='consolas 15' , width=13  ,anchor=E ).grid(row=2 , column=0 , sticky = E)
pass_lbl=tk.Label(login_frame ,text='password:', bg=bg_lbl_login ,fg=fg_lbl_login , font='consolas 15' , width=13 ,anchor=E ).grid(row=3 , column=0 , sticky = E)

canvas_home=Canvas(login_frame, width=800, height=500)
login_bg_image=PhotoImage(file='login_ps_yellow.png')
canvas_home.grid(row=0, column=0, columnspan=10, rowspan=10)
canvas_home.create_image(0, 0, image=login_bg_image, anchor=NW)

### title lable
#tk.Label(login_frame , text='Log in'  , bg = bg_lbl_login, fg=fg_lbl_login  , font=('consolas ', 35)  ,height=1 ).grid(row=0 ,column=0,columnspan=5  , padx=250 )

### username's lable and entry
user= StringVar()
user.set('hossam')
#tk.Label(login_frame ,text='username:', bg=bg_lbl_login ,fg=fg_lbl_login , font='consolas 15' , width=13  ,anchor=E ).grid(row=2 , column=0 , sticky = E)
name_txt=ttk.Entry(login_frame,font='consolas 13' , textvariable=user)
name_txt.grid(row=2 , column=1)

### password's lable and entry
passw=StringVar()
passw.set('25888')
#pass_lbl=tk.Label(login_frame ,text='password:', bg=bg_lbl_login ,fg=fg_lbl_login , font='consolas 15' , width=13 ,anchor=E ).grid(row=3 , column=0 , sticky = E)
pass_txt=tk.Entry(login_frame,font='consolas 13' , textvariable=passw , show='*')
pass_txt.grid(row=3 , column=1)

def login():


    if user.get() =='hossam' and passw.get() =='25888':
        #winsound.Beep(freq, 50)
        winsound.Beep(freq, 50)
        file_path.set('')
        btn_rules.grid_forget()
        export_or_not.set(False)
        supp.set(4.0)
        conf.set(4.0)

        #messagebox.showinfo('yes', 'valid username and passowrd')
        login_frame.grid_forget()
        home_frame.grid(row=0, column=0, columnspan=10, rowspan=10)
        form.geometry('800x500')
        form.columnconfigure(0,weight=1)
        form.rowconfigure(1,weight=1)
        form.resizable(True,True)

        if home_frame.grid_location(0, 0):
            form.configure(menu=menu_bar)
    else:
        messagebox.showerror( 'Error' , 'invalid username or password')
        user.set('')
        passw.set('')
        name_txt.focus()

def forget():
    user.set('hossam')
    passw.set('25888')
    login()

### login button
btn=tk.Button( login_frame , text='Login' , bg='#ff0077',fg='white' , width=7   , font = 'tahoma 15 '  , command=login )
btn.grid(row=4 , column=1,pady=(30,0) )


btn=tk.Button( login_frame , text='Forget Password' , bg='lightgray',fg='black' , width=15   , font = 'tahoma 10 '  , command=forget ).grid(row=4,column=2 , pady=30 , sticky=SW)

#def ext():
#    exit(0)
#tk.Button(login_frame , text ='exit' , bg = '#ff5050' , fg = 'white' , anchor =CENTER ,command = ext , font =' tahoma 13'  ).grid( row = 0 , column=0, sticky=NW  )
#lucida

login_frame.grid(row = 0 , column =0, columnspan=10 , rowspan=10)



####                          ####
####       HOME       ####
####                          ####



home_frame = Frame(form )

tk.Label(home_frame, text='Market Basket Analysis', bg = bg_lbl_home , fg=fg_lbl_home, font=('consolas ', 30), anchor=CENTER).grid(row=0, column=1, pady=30, columnspan=2)
tk.Label(home_frame, text='Select File:',  bg=bg_lbl_home , fg=fg_lbl_home , font='consolas 14', anchor=CENTER).grid(row=2, column=0)
tk.Label(home_frame, text='min_support:', bg=bg_lbl_home, fg=fg_lbl_home , font='consolas 13', anchor=E).grid(row=5, column=1, sticky=E)
tk.Label(home_frame, text='min_confidence:', bg=bg_lbl_home, fg=fg_lbl_home , font='consolas 13', anchor=E).grid(row=6, column=1, sticky=E)



canvas_home=Canvas(home_frame, width=800, height=500)
home_bg_image=PhotoImage(file='home_ps_mod.png')
canvas_home.grid(row=0, column=0, columnspan=10, rowspan=10)
canvas_home.create_image(0, 0, image=home_bg_image, anchor=NW)


#p=PhotoImage(file='.png')


#tk.Label(home_frame, text='Market Basket Analysis', bg = bg_lbl_home , fg=fg_lbl_home, font=('consolas ', 30), anchor=CENTER).grid(row=0, column=1, pady=30, columnspan=2)

fnt_entry=('arial', 12)
fnt2=('arial', 15)

#home_frame.config(background='darkblue')

def browse():
    f = askopenfilename(initialdir = "",title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))
    file_path.set(f)

file_path=tk.StringVar()

#tk.Label(home_frame, text='Select File:',  bg=bg_lbl_home , fg=fg_lbl_home , font='consolas 14', anchor=CENTER).grid(row=2, column=0)



txt_path=ttk.Entry(home_frame, font='arial 11' , textvariable=file_path , width =40)
btn_slct=tk.Button(home_frame , text = 'Browse..' ,font='arial 13 ', bg='deeppink' , fg='white' , command=browse)
txt_path.grid(row=2, column=1,columnspan=2)
btn_slct.grid(row=2 , column=3 )


def preview():
    try:
        path = file_path.get()
        with open(path, 'r') as file:
            records = list(csv.reader(file))

        prev = Toplevel()

        # prev.geometry('500x300+400+180')
        prev.geometry('+400+180')
        prev_bg = 'lightblue'
        prev.title('Preview')
        prev.configure(bg=prev_bg)
        prev.resizable(False, False)

        for i in range(10):
            pr = ' [ '
            for j in records[i]:
                pr += j + '  ,  '
            pr = pr[:-3] + ' ] '
            Label(prev, text='Transaction: ' + str(i + 1), bg=prev_bg, fg='darkblue', font='tahoma 8', borderwidth="3").grid(row=i , column=0 , sticky=W)
            Label(prev, text=pr, bg=prev_bg, fg='darkblue', font='tahoma ', borderwidth=5,).grid(row=i, column=1, sticky=W)

        Button(prev, text='Close', bg='#ff4040', fg='white', font='tahoma 11 bold' , command=prev.destroy, width=10 , height=1).grid(columnspan=2, pady=15)
        prev.focus()
        # messagebox.showinfo('sample' ,pr)
    except:
        messagebox.showerror('File error', 'Select a valid csv file ')

btn_prev=tk.Button(home_frame, text='Preview ►'  , bg='limegreen' , fg='white' ,width=9 , height=1 , font= 'arial 13 bold' , command= preview)
btn_prev.grid(row=4, pady=(0 , 30) ,padx=(120, 0) ,  columnspan=6)

supp=DoubleVar()
supp.set(4.0)
#tk.Label(home_frame, text='min_support:', bg=bg_lbl_home, fg=fg_lbl_home , font='consolas 13', anchor=E).grid(row=5, column=1, sticky=E)
txt_support=tk.Entry(home_frame, font=fnt_entry, width=7 , textvariable=supp)
txt_support.grid(row=5 , column=2,sticky=W)

conf= tk.IntVar()
conf.set(4.0)
#tk.Label(home_frame, text='min_confidence:', bg=bg_lbl_home, fg=fg_lbl_home , font='consolas 13', anchor=E).grid(row=6, column=1, sticky=E)
txt_confidence=tk.Entry(home_frame ,textvariable=conf , font=fnt_entry,width=7 )
txt_confidence.grid(row=6 , column=2,sticky=W)

btn_rules = tk.Button(home_frame, text='view rules', bg='orange', fg='white', font='consolas 12 bold')


def analyze():

    try:
        path=file_path.get()
        min_support=supp.get() / 100
        min_confidence=conf.get() / 100

        rules = apriori(path, min_support, min_confidence)

        def view_rules ():
            rule_view = Toplevel()
            rule_view.geometry('+300+150')
            rule_view.title('Results')
            rule_bg = 'lightblue'
            rule_view.configure(bg=rule_bg)
            #rule_view.resizable(False, False)
            rule_view.rowconfigure([i for i in range(len(rules[0]))] ,weight=1)
            rule_view.columnconfigure([0,1,2,3,4,5,6] ,weight=1)
            style = ttk.Style()
            style.theme_use('default')
            style.configure("black.Horizontal.TProgressbar", background='yellow')


            #Label(rule_view, text='Rule: ►' , bg=rule_bg, fg='black' , font='tahoma 11'  ).grid(row=0, column=0,columnspan=2 , padx=50 )
            Label(rule_view, text='Antecedent: ' , bg=rule_bg, fg='red' , font='tahoma 13 bold'  ).grid(row=0, column=0 , padx=40 ,pady=10)
            Label(rule_view, text='Consequent: ' , bg=rule_bg, fg='red' , font='tahoma 13 bold'  ).grid(row=0, column=1 , padx=40  ,pady=10)

            Label(rule_view, text='Support: %' , bg=rule_bg, fg='red' , font='tahoma 13 bold' ).grid(row=0, column=2,padx=40, columnspan=2  ,pady=10)
            Label(rule_view, text='Confidence: %' , bg=rule_bg, fg='red' , font='tahoma 13 bold' ).grid(row=0, column=4, padx=40, columnspan=2  ,pady=10)
            for i in range (15):  #(len(rules[0]) ):
                # if i%2 ==1:
                #     rule_bg='cyan'
                # else:
                #     rule_bg='lightblue'
                ro =i+1
                Label(rule_view , text=str (rules[0][i][:-1] ) , bg=rule_bg, fg='black', font='tahoma 12').grid(row=ro, column=0,  pady=5)
                Label(rule_view, text=str(rules[0][i][-1]), bg=rule_bg, fg='black', font='tahoma 12').grid(row=ro, column=1,  pady=5)

                Label(rule_view , text=round(rules[1][i] , 4 )  , bg=rule_bg , fg='black' , font='tahoma 12' ).grid(row=ro , column=3, pady=5)
                ttk.Progressbar(rule_view, length=100, style='black.Horizontal.TProgressbar', value=rules[1][i] ).grid(row=ro , column=2, pady=5)

                Label(rule_view , text=round(rules[2][i] ,4 ) , bg=rule_bg , fg='black' , font='tahoma 12' ).grid(row=ro , column=5, pady=5)
                ttk.Progressbar(rule_view, length=100, style='black.Horizontal.TProgressbar', value=rules[2][i] ).grid(row=ro , column=4, pady=5)


            Button(rule_view, text='Close', bg='#ff4040', fg='white', font='tahoma 11 bold' , command=rule_view.destroy, width=10 , height=1).grid(columnspan=10, pady=15)

        if export_or_not.get() == True:
            dt = datetime.datetime.now()
            name = str(dt.year) + str(dt.month) + str(dt.day) + str(dt.hour) + str(dt.minute) + str(dt.second)

            workbook = xl.Workbook('%s/MBA_output%s.xlsx'%(desktop,name))
            workseet = workbook.add_worksheet()

            workseet.write(0, 0, 'Antecedent:')
            workseet.write(0, 1, 'Consequent:')

            workseet.write(0, 2, 'Support: %')
            workseet.write(0, 3, 'Confidence: %')

            for line in range(len( rules[0] )):
                workseet.write(line + 1, 0, str(rules[0] [line][:-1]))
                workseet.write(line + 1, 1, str(rules[0] [line][-1]))

                workseet.write(line + 1, 2, rules [1] [line])
                workseet.write(line + 1, 3, rules [2] [line])

            workbook.close()
            messagebox.showinfo('Done','Output saved in:\n%s/MBA_output%s.xlsx'%(desktop,name))

        btn_rules.configure(command=view_rules)
        btn_rules.grid(row=7 , column=1, pady=1, columnspan=6 ,sticky=E)



        #for i in rules:
          #  print(i)
    except:
        btn_rules.grid_forget()
        messagebox.showerror('Error' ,'select valid; file path, min_support, and min_confidence' )


btn_analyze=tk.Button(home_frame ,  text='◄ Analyze ►' , bg='brown' , fg='white' , font='consolas 14 bold' ,command=analyze  )
btn_analyze.grid(row=7 , column=1,pady=30, columnspan=2)
export_or_not=BooleanVar()

cbx_export=tk.Checkbutton(home_frame ,text='Export Results' , bg='#3bb3c3', fg='black' ,font='consolas 12 bold' , variable = export_or_not    )
cbx_export.grid(row=7 , column=0,pady=30 ,sticky=E )


def ext():
    exit(0)
#tk.Button(home_frame , text ='exit' , bg = '#ff5050' , fg = 'white' , anchor =CENTER ,command = ext , font =' tahoma 13'  ).grid( row = 0 , column=0, sticky=NW  )




menu_bar = Menu(home_frame)
def func_new():
    file_path.set('')
    supp.set(0.2)
    conf.set(0.2)
    btn_rules.grid_forget()

def logout():
    user.set('')
    passw.set('')
    home_frame.grid_forget()
    login_frame.grid()
    form.geometry('597x380+300+130')
    form.resizable(False, False)

file = Menu(menu_bar)
file.add_command(label='new' , command=func_new)
file.add_command(label='open file' , command=browse)
file.add_command(label='Logout' , command=logout)
file.add_command(label='Exit' , command=ext)
menu_bar.add_cascade(label='File', menu=file)

view=Menu(menu_bar)
view.add_command(label='font')
view.add_command(label='mode')
menu_bar.add_cascade(label='View', menu=view)

def hlp():
    messagebox.showinfo('Help' , 'first a csv file path should be entered, '
                                 '\n\nbrowse button: to select the file directly, '
                                 '\n\npreview button: to preview data from your file, '
                                 '\n\nmin_support and min_confidence:'
                                 '\n\tare used for applying apriori algorithm, '
                                 '\n\nand analyze button: to apply apriory algorithm' )
#tk.Button(home_frame , text ='help' , bg = 'lightblue' , fg = 'black' , anchor =CENTER ,command = help , font =' tahoma 13 '  ).grid( row = 0 , column=0,padx=42, sticky=NW  )

help=Menu(menu_bar)
help.add_command(label='help' , command=hlp)
menu_bar.add_cascade(label='Help', menu=help)


def about_us():
    messagebox.showinfo('About us' , 'Team H2I2A2\nfrom faculty of computers and information\nat fayoum university\n\n\ncontact us at: hh1340@fayoum.edu.eg')
about=Menu(menu_bar)
about.add_command(label='about us' , command=about_us)
menu_bar.add_cascade(label='About', menu=about)





def apriori(path, m_supp, m_conf):
    with open(path, 'r') as file:
        data_set = list(csv.reader(file))


    # data_set = [['eggs', 'sugar', 'tea', 'bread', 'soda'],
    #             ['bread', 'cheese', 'eggs', 'juice'],
    #             ['soda', 'biscuit', 'juice'],
    #             ['bread', 'eggs', 'soda', 'cheese'],
    #             ['tea', 'sugar', 'bread', 'soda'],
    #             ['biscuit', 'tea'],
    #             ['juice', 'biscuit'],
    #             ['tea', 'biscuit', 'soda'],
    #             ['juice', 'bread', 'biscuit', 'soda'],
    #             ['sugar', 'tea', 'eggs'],
    #             ['sugar', 'tea'],
    #             ['bread', 'cheese'],
    #             ['sugar', 'tea'],
    #             ['biscuit', 'juice'],
    #             ['sugar', 'tea', 'juice', 'biscuit']
    #             ]


    ### Function to count the support_n
    def support_calc(data_set, list):
        counter = 0
        check_all = 0
        returned = []
        for group in list:
            for record0 in data_set:
                for element in group:
                    if element in record0:
                        check_all += 1
                if check_all == len(group):
                    counter += 1
                check_all = 0
            returned.append(counter / len(data_set))
            counter = 0
        return returned


    ### Function to Concatinate to get the next candidate
    def conc(lst):
        returned = []
        for base in lst[:len(lst) - 1]:
            for add in lst[lst.index(base) + 1:]:
                m = [x for x in add if x not in base]
                for item in m:
                    if (base + [item]) not in returned:
                        returned.append(base + [item])
        return returned

    ### Remove item_sets less than minimum support and its support count
    def remove_lt_min_supp(items_sets, support_n , min_supp ):
        x = 0
        while (x < len(items_sets)):
            if support_n[x] < min_supp: #or confidence_n[x] < min_conf:
                del items_sets[x]
                del support_n[x]
                #del confidence_n[x]
                x -= 1
            x += 1

    ### getting all item_sets from the dataset for the first time
    item_sets = []
    for record in data_set:
        for item in record:
            if [item] not in item_sets:
                item_sets.append([item])
    ###count number of item in all records
    support_n = support_calc(data_set, item_sets)

    ### Remove item_sets less than minimum support and its support_n from C1 to get L1

    remove_lt_min_supp(item_sets, support_n, m_supp)


    s= support_n
    t= item_sets




    rules = []
    supp_of_rules = []

    z = 2
    m=0
    while not item_sets == []:
        ### getting C2
        item_sets = conc(item_sets)

        ### Calculate support of Cn
        support_n = support_calc(data_set, item_sets)

        ### Remove item_sets less than minimum support and its support_n from C2 to get L2
        remove_lt_min_supp(item_sets, support_n , m_supp )

        t+=item_sets
        s+= support_n



        rules += item_sets
        supp_of_rules += support_n

        z += 1

    ### End Of LOOP
    #############################################################################
    #############################################################################

    all_supports=[]
    all_confidences = []
    all_rules = []

    for group in rules  :
        for consequent in group:
            sup_a_b =round( s[t.index(group)] ,5 )

            b = group[:group.index(consequent)] + group[group.index(consequent) + 1:]
            sup_a = s[t.index(b)]
            conf =round( sup_a_b/sup_a , 5 )
            if conf >= m_conf:
                all_supports.append(sup_a_b*100)
                all_confidences.append(conf*100)
                all_rules.append(b + [consequent])

    #print(all_rules)
    #print(all_supports)
    #print(all_confidences)

    n = len(all_confidences)
    for i in range(n):
        for j in range(0, n - i - 1):

            if all_confidences[j] < all_confidences[j + 1]:
                all_confidences[j], all_confidences[j + 1] = all_confidences[j + 1], all_confidences[j]
                all_supports[j], all_supports[j + 1] = all_supports[j + 1], all_supports[j]
                all_rules[j], all_rules[j + 1] = all_rules[j + 1], all_rules[j]



    return [all_rules , all_supports , all_confidences]
    ##*******-------*******-------*******-------*******-------*******-------*******-------*******
    ### MadeBy (" Hossam Zaki ")///





form.mainloop()

#mail1.place(height=70, width=400, x=83, y=109)








