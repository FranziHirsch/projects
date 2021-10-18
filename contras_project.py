import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import messagebox
import pandas as pd
import os
import time
import threading
import queue
import pprint
from rapidfuzz import fuzz
from rapidfuzz import process, utils
import numpy as np
import itertools






#re-write the code using OOP
#Also figure out the progress bar - this shuld only appear when the button is clicked

# Define a class
# Want this class to inherit from tkinter, so put in brackets tk.TK. If I don't want my class to inherit from another
# class, don't put brackets
class contra_app(tk.Tk):

    # Initialise method - thigs placed under this method will always run when I run the class
    # Is like booting up a PC - what do I always want to run immediately when I run my class "PC" - e.g. load up keyboard, mouse etc.
    # I don't want, e.g. for music to play immediately when I start my PC, so this play_music would be another method
    # Into init, pass:
        # self (is convention to call "self", could be calle anything). With this parameter, you can access any attributes/ method in the class

    def __init__(self, *args, **kwargs):



        # Initialise tkinter
        # This creates a root window to place everything into
        tk.Tk.__init__(self, *args, **kwargs)

        # Create string variables for use in widgets. Create using methods below.
        self.input_file_name=tk.StringVar()
        self.output_directory=tk.StringVar()
        self.list_col=tk.StringVar()
        self.list_col2=tk.StringVar()
        self.list_col3=tk.StringVar()

        # I want to create a progress bar below. Will set progress bar "maximum" to x. x is an int - indicates how many units must be filled to fill the progress bar
        self.int_var=tk.IntVar()


        # # df = pd.read_excel("C:/Users/franz/Desktop/python/contras/Input_Data/sample_dataset.xlsx")
        # df = pd.read_excel(self.input_file_name["textvariable"]) #FIX THIS
        # self.list_columns = df.columns.to_list()
        # print(self.list_columns)

        # Replace the default title in the root window
        self.title("Contra Wizard")

        # Create a canvas to place everything into
        canvas=tk.Canvas(self, height=500, width=800, cursor="arrow", bg="white")
        canvas.pack()

        # # Create a header label
        header = tk.Label(self, bg="white", text="Contra Wizard", font=("Calibri", 23, "bold"))
        header.place(rely=0.03, relx=0.02, relheight=0.1, relwidth=0.3)

        # Create 2 frames
        frame_left = tk.Frame(self, bg="#EAEBEC")
        frame_left.place(rely=0.15, relx=0.05, relheight=0.67, relwidth=0.44)

        frame_right = tk.Frame(self, bg="#EAEBEC")
        frame_right.place(rely=0.15, relx=0.5, relheight=0.67, relwidth=0.44)

        # Populate left hand frame

        # Input file header - this will contain the instruction to the user to upload their input file
        input_file_header = tk.Label(frame_left, bg="#EAEBEC", text="1. Upload input file", font=("Calibri", 10, "bold italic"),anchor="w", justify="left")
        input_file_header.place(rely=0.1, relx=0.02, relheight=0.1, relwidth=0.3)

        # Entry box - this will show the name of the file the user selected
        self.input_file_entry = tk.Entry(frame_left, bg="white", bd=2, font=("Calibri", 10), textvariable=self.get_input_file())
        self.input_file_entry.place(rely=0.2, relx=0.02, relheight=0.1, relwidth=0.65)

        # Button - this will open the directory box
        input_file_button = tk.Button(frame_left, text="Browse", bg="white", font=("Calibri", 10), command=self.open_sourcefile)
        input_file_button.place(rely=0.2, relx=0.7, relheight=0.1, relwidth=0.25)

        # Output name header - this will instruct the user to enter a name for their output file
        output_name_label = tk.Label(frame_left, bg="#EAEBEC", text="2. Enter output file name",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        output_name_label.place(rely=0.3, relx=0.02, relheight=0.1, relwidth=0.4)

        # Entry box - This is where the user enters the name of their output file
        self.output_file_name = tk.Entry(frame_left, bg="white", bd=2, font=("Calibri", 10))
        self.output_file_name.place(rely=0.4, relx=0.02, relheight=0.1, relwidth=0.93)

        # Output dir header - This instructs the user to select where they want to save their file
        output_dir_header = tk.Label(frame_left, bg="#EAEBEC", text="3. Select output directory",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        output_dir_header.place(rely=0.5, relx=0.02, relheight=0.1, relwidth=0.5)

        # Entry box - this will show the output directory the user picks
        self.output_directory_name = tk.Entry(frame_left, bg="white", bd=2, font=("Calibri", 10), textvariable=self.get_output_dir())
        self.output_directory_name.place(rely=0.6, relx=0.02, relheight=0.1, relwidth=0.65)

        # Button - this will open the directory box
        output_dir_button = tk.Button(frame_left, text="Browse", bg="white", font=("Calibri", 10), command=self.output_dir)
        output_dir_button.place(rely=0.6, relx=0.7, relheight=0.1, relwidth=0.25)

        # Populate right hand frame

        # Refresh button - this will populate the column entry boxes
        self.refresh_button = tk.Button(frame_right, text="Refresh columns", bg="white", font=("Calibri", 8), justify="center", command=self.get_columns)
        self.refresh_button.place(rely=0.04, relx=0.7, relheight=0.06, relwidth=0.25)

        # Amounts header - this is where the user specifes which column contains the amounts
        amounts_header = tk.Label(frame_right, bg="#EAEBEC", text="4. Specify 'Amounts' column",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        amounts_header.place(rely=0.1, relx=0.02, relheight=0.1, relwidth=0.94)

        # # List of columns
        self.list_col = tk.Listbox(frame_right, bg="white", bd=2, font=("Calibri", 10), selectmode="single",yscrollcommand=True, exportselection=0)
        self.list_col.place(rely=0.2, relx=0.02, relheight=0.1, relwidth=0.94)

        # Date header - this is where the user specifes which column contains the dates
        date_header = tk.Label(frame_right, bg="#EAEBEC", text="5. Specify 'Date' column",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        date_header.place(rely=0.3, relx=0.02, relheight=0.1, relwidth=0.94)

        # # List of columns
        self.list_col2 = tk.Listbox(frame_right, bg="white", bd=2, font=("Calibri", 10), selectmode="single",yscrollcommand=True, exportselection=0)
        self.list_col2.place(rely=0.4, relx=0.02, relheight=0.1, relwidth=0.94)

        # Description header - this is where the user specifes which column contains the description
        # FUTURE DEVELOPMENT POINT: USER CAN SELECT MORE THAN ONE COLUMN
        descr_header = tk.Label(frame_right, bg="#EAEBEC", text="6. Specify 'Description' column",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        descr_header.place(rely=0.5, relx=0.02, relheight=0.1, relwidth=0.94)

        # # List of columns
        self.list_col3 = tk.Listbox(frame_right, bg="white", bd=2, font=("Calibri", 10), selectmode="single",yscrollcommand=True, exportselection=0)
        self.list_col3.place(rely=0.6, relx=0.02, relheight=0.1, relwidth=0.94)

        # Fuzzy button - instructs the user to select if they wish the tool to detect fuzzy contras
        fuzzy_label = tk.Label(frame_right, bg="#EAEBEC", text="7. Detect fuzzy contras?",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        fuzzy_label.place(rely=0.7, relx=0.02, relheight=0.1, relwidth=0.94)

        self.var = tk.IntVar()
        yes_button = tk.Radiobutton(frame_right, text="Yes", variable=self.var, value=1, bg="#EAEBEC",font=("Calibri", 10), anchor="w", justify="left")
        yes_button.place(rely=0.77, relx=0.02, relheight=0.1, relwidth=0.2)

        no_button = tk.Radiobutton(frame_right, text="No", variable=self.var, value=2, bg="#EAEBEC",font=("Calibri", 10), anchor="w", justify="left")
        no_button.place(rely=0.77, relx=0.25, relheight=0.1, relwidth=0.4)

        # Split button - instructs the user to select if they wish the tool to detect split contras
        split_label = tk.Label(frame_right, bg="#EAEBEC", text="8. Detect split contras?",font=("Calibri", 10, "bold italic"), anchor="w", justify="left")
        split_label.place(rely=0.7, relx=0.5, relheight=0.1, relwidth=0.94)

        self.var2 = tk.IntVar()
        yes_button = tk.Radiobutton(frame_right, text="Yes", variable=self.var2, value=1, bg="#EAEBEC",font=("Calibri", 10), anchor="w", justify="left")
        yes_button.place(rely=0.77, relx=0.5, relheight=0.1, relwidth=0.2)

        no_button = tk.Radiobutton(frame_right, text="No", variable=self.var2, value=2, bg="#EAEBEC",font=("Calibri", 10), anchor="w", justify="left")
        no_button.place(rely=0.77, relx=0.75, relheight=0.1, relwidth=0.4)

        # Create a progress bar widget to assure the user something is happening when they run the code
        # Mode "Determinate" means the widget shows an indicator that moves from beginning to end under program control.
        # maximum=4 means 4 units must be filled to fill the progress bar (i.e. 4 methods must run)
        progress = Progressbar(frame_right, orient="horizontal", length=100, mode='determinate', maximum=4)
        # variable is an option within Progressbar. Indicates the current value of the indicator. This is how I access it:
        progress["variable"]=self.int_var
        progress.place(rely=0.85, relx=0.4, relheight=0.1, relwidth=0.25)




        # Run button - to execute the code
        self.run_button = tk.Button(frame_right, text="Run", bg="white", font=("Calibri", 10), command=self.start_thread)
        self.run_button.place(rely=0.85, relx=0.7, relheight=0.1, relwidth=0.25)








    # Class methods

    def open_sourcefile(self):
        # Allows user to select and import a single Excel file
        # This creates a dialogue box. I can specify the file types/ extensions the user can select. Would need to enter each extension that is allowed (separate by a space)
        input_file= filedialog.askopenfilename(title="Select input file",filetypes=[("Excel files", "*.xlsx *.xls")])
        # We set up input_file_name as a string variable under __init__()
        self.input_file_name.set(input_file)
        return input_file

    def get_input_file(self):
        # Create the label text for entry box input_file_name
        return self.input_file_name

    def get_input_name(self):
        # This stores the input file name selected by the user in a variable so I can work with it
        input_name=self.input_file_name.get()
        return input_name

    def get_output_name(self):
        # This stores the output file name selected by the user in a variable so I can work with it
        output_name=self.output_file_name.get()
        return output_name

    def output_dir(self):
        # Allow user to select an output directory and store it in global var called folder_path
        output_dir_path = filedialog.askdirectory(title="Select output directory")
        self.output_directory.set(output_dir_path)
        return output_dir_path

    def get_output_dir(self):
        # Create the label text for entry box output_dir_name
        return self.output_directory

    def get_output_dir_name(self):
        # This stores the output directory selected by the user in a variable so I can work with it
        output_dir_name=self.output_directory_name.get()
        return output_dir_name

    def get_columns(self):
        # This function takes the columns from the input file selected by the user and returns them to the entry boxes
        # Note: This function will have to wait to be executed until the user selects the input file
        # Else the method self.get_input_name() returns None & this will throw an error
        thread_columns=threading.Thread(target=self.get_input_file)
        thread_columns.start()

        # wait here for the result to be available before continuing
        thread_columns.join()

        # Perform this action
        df = pd.read_excel(self.get_input_name())
        self.list_columns=df.columns.to_list()
        # for c in self.list_columns:
        #     self.list_col.insert(tk.END, c)
        #     self.list_col2.insert(tk.END, c)
        #     self.list_col3.insert(tk.END, c)
        for self.c in self.list_columns:
            self.list_col.insert(tk.END, self.c)
            self.list_col2.insert(tk.END, self.c)
            self.list_col3.insert(tk.END, self.c)
        return self.list_col, self.list_col2, self.list_col3

    def get_descr_col(self):
        # This function takes the value selected by the user & stores this in a variable
        sel=self.list_col3.curselection()
        descr_col=self.list_columns[sel[0]]
        return descr_col

    def get_date_col(self):
        # This function takes the value selected by the user & stores this in a variable
        # curselection returns a tuple. The first element represents the position of the cursor
        sel=self.list_col2.curselection()
        date_col=self.list_columns[sel[0]]
        return date_col

    def get_amt_col(self):
        # This function takes the value selected by the user & stores this in a variable
        sel = self.list_col.curselection()
        amt_col=self.list_columns[sel[0]]
        return amt_col





    # Progress bar

    def start_thread(self):
        # What is this? A Thread is a part of a process (i.e. the program that is being executed). Simple programs only need one threat, however this means
        # that if there is a barrier (e.g. program waiting for user input), the whole program is held up by this barrier
        # If I split the codes into multuple threads, then parts of the code can run whilst another part of the code waits for the barrier to be lifted (e.g. user provides input)
        # Why do I need to split my code into threads in this instance?
        # tkinter takes possession of the main thread (how? If I execute mainloop(), then tkinter takes possesison of the main thread), so if I have work
        # intensive functions I need to run, this will interfere with the mainloop, hence my functions should always be outside the mainloop

        # Grey out the button after it is pressed by changing "state" (i.e. an option within button widget) like this
        self.run_button["state"]="disable"

        #Empty the progressbar by sessing intvar to 0
        self.int_var.set(0)

        # Create a new method "secondary_thread". Call threading library and class "Thread"
        # The Thread class represents an activity that is run in a separate thread of control.
        # target is the callable object to be invoked by the run() method. "arbitrary" is created below - this calls the methods
        self.secondary_thread=threading.Thread(target=self.arbitrary)
        # Start the secondary thread's activity
        self.secondary_thread.start()
        # After 50 ms, execute function check_que.
        self.after(50, self.check_que)

    def check_que(self):
        # The secondary thread puts information into a queue (see below). This function accesses this queue
        while True:
            try:
                # get_nowait(): Return an item from the queue if the queue is not empty
                x=que.get_nowait()
            except queue.Empty:
                self.after(25, self.check_que)
                break
            else: # continue from the try suite
                # Set int_var to the relevant item in the queue
                self.int_var.set(x)
                # After the last item has been reached, unfreeze the button
                if x==4:
                    self.run_button["state"]="normal"
                    # I also want to show a pop-up button when the code is finished
                    messagebox.showinfo("Status","Done!")
                    # button = tk.Button(popup, text="Okay", command=popup.destroy)
                    # button.pack()
                    break




    #Contra detection

    # Read in dataset
    def read_dataset(self):

        # At the moment, this can only read in Excel files
        self.df = pd.read_excel(self.get_input_name())

    def contra_loop(self, df):

        # Sleep pauses program for x seconds
        time.sleep(0.3)

        # Define a function which identifies perfect & fuzzy contras
        def perfect_and_fuzzy_contras(idx1, idx2, amount1, amount2, description1, description2, date1, date2):

            perf_contra = "Perfect Contra"
            fuzz_contra = "Fuzzy Contra"
            oth = "Other"

            which_fuzzy_button_is_selected = self.var.get() #1 is "Yes", 2 is "No"

            if which_fuzzy_button_is_selected == 1:
                if (idx1 != idx2) and (amount2 == -amount1) and (description2 == description1) and (date2 > date1):
                    return perf_contra
                elif (idx1 != idx2) and (amount2 == -amount1) and fuzz.ratio(description2, description1) >= 80 and (
                        date2 > date1):
                    return fuzz_contra
                else:
                    return oth

            elif which_fuzzy_button_is_selected == 2:
                if (idx1 != idx2) and (amount2 == -amount1) and (description2 == description1) and (date2 > date1):
                    return perf_contra
                else:
                    return oth

            else:
                if (idx1 != idx2) and (amount2 == -amount1) and (description2 == description1) and (date2 > date1):
                    return perf_contra
                else:
                    return oth


        amt_col= self.list_col.curselection()[0]+1 #This returns the index. Note: I must add 1 because the first column will be the index - want to skip that
        date_col = self.list_col2.curselection()[0]+1
        descr_col = self.list_col3.curselection()[0]+1

        # Create an empty dict where I will store all rows which are perfect contras. The key will be the row index and the value will be "perfect contra"
        results_dict = {}

        # Create an empty dict where I will store the pair rows. The key will be the row index and the value will be a list with the indexes of the pair rows
        pairs_dict = {}

        # Iterate up the dataframe
        for row1 in self.df.itertuples():
            if row1.Index not in results_dict:
            # This is to ensure that when a contra pair has been identified, this does not get overridden during the next iteration
                for row2 in self.df.itertuples():
                    if row2.Index not in results_dict:
                        result=perfect_and_fuzzy_contras(row1.Index, row2.Index, row1[amt_col],row2[amt_col], row1[descr_col],row2[descr_col],row1[date_col],row2[date_col])
                        if result == "Perfect Contra" or result == "Fuzzy Contra":
                                pairs_dict[row1.Index]=[row1.Index,row2.Index]
                                pairs_dict[row2.Index]=[row1.Index,row2.Index]
                                results_dict[row1.Index]=result
                                results_dict[row2.Index]=result
                                break


        # Create 2 new columns where to display the results
        self.df["Contras"] = ""
        self.df["Contra_Pairs"] = ""

        # Now insert the values into the correct rows using the "index" key
        for r, v in results_dict.items():
            self.df["Contras"].iloc[r] = v

        for r, v in pairs_dict.items():
            self.df["Contra_Pairs"].iloc[r] = v

        return self.df


    def split_contras(self, df):

        # Sleep pauses program for x seconds
        time.sleep(0.6)

        which_split_button_is_selected = self.var2.get()  # 1 is "Yes", 2 is "No"

        amt_col= self.get_amt_col()

        if which_split_button_is_selected == 1:
            # Create a copy of the amount column
            self.df["Amount2"] = self.df[amt_col]
            # Sets all values in col "Functional_Amount2" where the value in col "Contras" is "Fuzzy Contra" or "Perfect Contra" to Contra
            # This is so that they don't get taken into account when looking for split contras
            self.df["Amount2"][self.df.Contras == "Perfect Contra"] = "Contra"
            self.df["Amount2"][self.df.Contras == "Fuzzy Contra"] = "Contra"


            numbers = self.df["Amount2"].to_list()

            # This is a list of indexes where every index represents a number from list "numbers"
            indexes = [idx for idx, num in enumerate(numbers)]

            # Create a new list where every number is paired with its index
            idx_num = list(zip(indexes, numbers))

            # Take out all the "nans"
            idx_num = [n for n in idx_num if n[1] != "Contra"]
            # print(idx_num)

            # Create an empty list in which to store the results
            results = set()

            # Whenever we identify a contra combination, we cannot use any of the numbers/ indexes in another contra combinatinn
            # So we will add these to a list "remove" and any entries in that list can't be used for a new combination
            remove= set()

            # I want to check all combinations of numbers from 3 numbers to the len of the idx_no
            # Formula first checks all combinations of 3 numbers, then 4, 5, etc.
            # Approach #1: Identify the first combination of numbers which nets to Nil as a split contra. Do not use the
            # numbers for any other combinations
            # Approach #2: Identify all possible combinations of contras

            # Approach #1
            # Iterate through len idx_num from 3 to get the numbers from 3 to the len of idx_um
            for i in range(3, len(idx_num)):
                #NEED TO PUT THE TEST WHETHER THE ITEM IS IN "REMOVE"
                # Determine all combinations of elements of list num_idx which have i members
                combinations = list(itertools.combinations(idx_num, i))
                # indeces = [[s[0] for s in seq] for seq in combinations]
                # print(indeces)
                # numbers = [[s[1] for s in seq] for seq in combinations]
                for seq in combinations:
                    for s in seq:
                        if s[0] not in remove:
                # Iterate through combinations
                # for seq in combinations:
                #     print(seq)
                    # Extract the indexes to a separate list
                    # indexes = [s[0] for s in seq]
                    # Are any of the indexes in the list remove?
                    # check = any(index in indexes for index in remove)
                    # If not, then extract all the numbers into a sep. list "Sum" and calculate the sum
                    # if check == False:
                            Sum=[]
                            Sum=[s[1] for s in seq]
                        # s = sum(Sum)
                        # If the sum is 0, then add the index to list "remove" and the combination to list "results"
                        # if s == 0:
                            if sum(Sum)==0:
                            # for i in indexes:
                            #     for i in indeces:
                                remove.add(s[0])
                            results.add((s[0],s[1]))


            for contras in results:
                for c in contras:
                    self.df["Contra_Pairs"].iloc[c[0]] =[c[0] in c for c in contras]
                    self.df["Contras"].iloc[c[0]] = "Split Contra"



            self.df.drop(columns=["Amount2"], inplace=True)

        else:
            pass

        return self.df

    def save_output(self, df):
        # Sleep pauses program for x seconds
        time.sleep(0.9)

        os.chdir(self.get_output_dir_name())

        df.to_excel(f"{self.get_output_name()}.xlsx")




    def arbitrary(self):
        # Queue is a built-in module in Python - used in threading
        # What is a queue? Stores items on a first in- first out basis
        # Put test_function into queue as item 1
        self.contra_loop(self.read_dataset())
        que.put(1)
        # Put test_function2 into queue as item 2
        self.split_contras(self.df)
        que.put(2)
        # Put test_function3 into queue as item 3
        self.save_output(self.df)
        que.put(3)
        # Put test_function4 into queue as item 4
        que.put(4)









que = queue.Queue()
app=contra_app()
# app.set_root(500, 800)
# Execute root window
tk.mainloop()
