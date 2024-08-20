import tkinter as tk
import numpy as np
from tkinter import filedialog, ttk,messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# Creating the Tkinter window
root = tk.Tk()
root.title("Student Result Analysis")
root.geometry("1000x1200")
root.configure(background='#FDFD96')

# scrollbar
sb = ttk.Scrollbar(root)

perdata = pd.DataFrame()

# Adding a figure to show the graph in the window
fig = Figure()
a = fig.add_subplot()
canvas = FigureCanvasTkAgg(fig, master=root)
plot_widget = canvas.get_tk_widget()

# Global variable to store the file path
file_path = ""


def clearToTextInput():
    T.delete("1.0", "end")


def english_marks():
    a.clear()
    d = pd.read_excel(file_path, index_col=0)
    clearToTextInput()
    T.insert(tk.INSERT, "English.\n\n")
    bins = [39, 59, 69, 74, 79, 90, 95, 100]
    a.clear()
    perdata["English"] = d['English'].value_counts(bins=bins, sort=False)
    avg_eng = d['English'].values.mean()
    gradedata = pd.DataFrame(index=["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2"])
    gradedata["English"] = d['GR1'].value_counts()
    English = d['English'].values
    value, count = np.unique(English, return_counts=True)
    english_data = dict(zip(value, count))
    a.set_xlabel('Percentage in English', fontsize=15)
    a.set_ylabel('Number of Students', fontsize=15)
    a.set_title('RESULT ANALYSIS', fontsize=20)
    a.plot(english_data.keys(), english_data.values())
    fig.canvas.draw()
    dic_eng = dict(zip(d.index, d['English'].values))
    max_eng = max(d['English'].values)
    for i_eng in dic_eng.keys():
        if dic_eng[i_eng] == max_eng:
            T.insert(tk.INSERT, str(i_eng) + ",")
    T.insert(tk.INSERT, "Highest marks in English: " + str(max_eng) + ".\n\n")
    T.insert(tk.INSERT, "Average marks in English:" + str("%.2f" % avg_eng) + ".\n\n")
    T.insert(tk.INSERT, gradedata)


def Maths_marks():
    a.clear()
    d = pd.read_excel(file_path, index_col=0)
    clearToTextInput()
    T.insert(tk.INSERT, "\tMathematics.\n")
    bins = [39, 59, 69, 74, 79, 90, 95, 100]
    perdata["Maths"] = d['Mathematics'].value_counts(bins=bins, sort=False)
    avg_math = d['Mathematics'].values.mean()
    gradedata = pd.DataFrame(index=["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2"])
    gradedata["Maths"] = d['GR2'].value_counts()
    Mathematics = d['Mathematics'].values
    value, count = np.unique(Mathematics, return_counts=True)
    maths_data = dict(zip(value, count))
    a.set_xlabel('Percentage in Maths', fontsize=15)
    a.set_ylabel('Number of Students', fontsize=15)
    a.set_title('RESULT ANALYSIS', fontsize=20)
    a.plot(maths_data.keys(), maths_data.values())
    fig.canvas.draw()
    dic_maths = dict(zip(d.index, d['Mathematics'].values))
    max_maths = max(d['Mathematics'].values)
    for i_maths in dic_maths.keys():
        if dic_maths[i_maths] == max_maths:
            T.insert(tk.INSERT, str(i_maths) + ",")
    T.insert(tk.INSERT, "highest marks in Mathematics: " + str(max_maths) + ".\n\n")
    T.insert(tk.INSERT, "Average marks in Maths: " + str(avg_math))
    T.insert(tk.INSERT, gradedata)


def Physics_marks():
    a.clear()
    d = pd.read_excel(file_path, index_col=0)
    clearToTextInput()
    T.insert(tk.INSERT, "Physics" + ".\n")
    bins = [39, 59, 69, 74, 79, 90, 95, 100]
    perdata = pd.DataFrame()
    perdata["Physics"] = d['Physics'].value_counts(bins=bins, sort=False)
    avg_phy = d['Physics'].values.mean()
    gradedata = pd.DataFrame(index=["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2"])
    gradedata["Physics"] = d['GR3'].value_counts()
    Physics = d['Physics'].values
    value, count = np.unique(Physics, return_counts=True)
    Physics_data = dict(zip(value, count))
    a.set_xlabel('Percentage in Physics', fontsize=15)
    a.set_ylabel('Number of Students', fontsize=15)
    a.set_title('RESULT ANALYSIS', fontsize=20)
    a.plot(Physics_data.keys(), Physics_data.values())
    fig.canvas.draw()
    dic_Physics = dict(zip(d.index, d['Physics'].values))
    max_Physics = max(d['Physics'].values)
    for i_Physics in dic_Physics.keys():
        if dic_Physics[i_Physics] == max_Physics:
            T.insert(tk.INSERT, str(i_Physics) + ",")
    T.insert(tk.INSERT, "highest marks in Physics: " + str(max_Physics) + ".\n\n")
    T.insert(tk.INSERT, "Average marks in Physics: " + str(avg_phy))
    T.insert(tk.INSERT, gradedata)


def Chem_marks():
    a.clear()
    d = pd.read_excel(file_path, index_col=0)
    clearToTextInput()
    T.insert(tk.INSERT, "Chemistry" + ".\n")
    bins = [39, 59, 69, 74, 79, 90, 95, 100]
    perdata = pd.DataFrame()
    perdata["Chemistry"] = d['Chemistry'].value_counts(bins=bins, sort=False)
    avg_chem = d['Chemistry'].values.mean()
    gradedata = pd.DataFrame(index=["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2"])
    gradedata["Chemistry"] = d['GR4'].value_counts()
    Chemistry = d['Chemistry'].values
    value, count = np.unique(Chemistry, return_counts=True)
    Chemistry_data = dict(zip(value, count))
    a.set_xlabel('Percentage in Chemistry', fontsize=15)
    a.set_ylabel('Number of Students', fontsize=15)
    a.set_title('RESULT ANALYSIS', fontsize=20)
    a.plot(Chemistry_data.keys(), Chemistry_data.values())
    fig.canvas.draw_idle()
    dic_Chemistry = dict(zip(d.index, d['Chemistry'].values))
    max_Chemistry = max(d['Chemistry'].values)
    for i_Chemistry in dic_Chemistry.keys():
        if dic_Chemistry[i_Chemistry] == max_Chemistry:
            T.insert(tk.INSERT, str(i_Chemistry) + ",")
    T.insert(tk.INSERT, "highest marks in Chemistry: " + str(max_Chemistry) + ".\n\n")
    T.insert(tk.INSERT, "Average marks in chemistry: " + str(avg_chem))
    T.insert(tk.INSERT, gradedata)


def CS_marks():
    a.clear()
    d = pd.read_excel(file_path, index_col=0)
    clearToTextInput()
    T.insert(tk.INSERT, "CS" + ".\n")
    bins = [39, 59, 69, 74, 79, 90, 95, 100]
    perdata = pd.DataFrame()
    perdata["CS"] = d['CS'].value_counts(bins=bins, sort=False)
    avg_cs = d['CS'].values.mean()
    gradedata = pd.DataFrame(index=["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2"])
    gradedata["Computer Science"] = d['GR5'].value_counts()
    CS = d['CS'].values
    value, count = np.unique(CS, return_counts=True)
    CS_data = dict(zip(value, count))
    a.set_xlabel('Percentage in Computer Science', fontsize=15)
    a.set_ylabel('Number of Students', fontsize=15)
    a.set_title('RESULT ANALYSIS', fontsize=20)
    a.plot(CS_data.keys(), CS_data.values())
    fig.canvas.draw()
    dic_CS = dict(zip(d.index, d['CS'].values))
    max_CS = max(d['CS'].values)
    for i_CS in dic_CS.keys():
        if dic_CS[i_CS] == max_CS:
            T.insert(tk.INSERT, str(i_CS) + ",")
    T.insert(tk.INSERT, "highest marks in Computer Science: " + str(max_CS) + ".\n\n")
    T.insert(tk.INSERT, "Average marks in computer Science: " + str(avg_cs))
    T.insert(tk.INSERT, gradedata)


# Main program where calculations are done and results are printed/plotted
def main():
    a.clear()
    # declaring path of file choose
    global file_path

    # Here the variable d reads the selected file
    d = pd.read_excel(file_path, index_col=0)

    # sum of all five subjects
    total = d['per'].values

    # selects distinct values of total marks scored by the student
    v, c = np.unique(total, return_counts=True)

    # creating a bar graph of all the marks of students
    t = dict(zip(v, c))
    a.bar(t.keys(), t.values())
    a.set_xlabel('Percentage', fontsize=15)
    a.set_ylabel('Number of Students', fontsize=15)
    a.set_title('RESULT ANALYSIS', fontsize=20)
    fig.canvas.draw_idle()

    avg_total = d['per'].values.mean()

    dic_total = dict(zip(d.index, d['Total marks'].values))
    max_total = max(d['Total marks'].values)

    bins = [39, 59, 69, 74, 79, 90, 95, 100]

    # it will print number of students who have received marks within a particualr range
    print(d['per'].value_counts(bins=bins, sort=False))

    T.insert(tk.INSERT, "\n")
    T.insert(tk.INSERT, "Max marks obtained: " + str(max_total) + ".\n\n")
    for i_total in dic_total.keys():
        if dic_total[i_total] == max_total:
            T.insert(tk.INSERT, "By:" + str(i_total) + ".\n\n")

    above_avg = 0
    above_pass = 0
    pass_marks = 300
    for j in d['Total marks'].values:
        if j > avg_total:
            above_avg += 1
    for k in d['Total marks'].values:
        if k >= pass_marks:
            above_pass += 1

    T.insert(tk.INSERT,
             " The pass marks are " + str(pass_marks) + " and " + str(above_pass) + " students have passed.\n\n")
    gradedata = pd.DataFrame(index=["A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2"])
    gradedata["English"] = d['GR1'].value_counts()
    gradedata["Maths"] = d['GR2'].value_counts()
    gradedata["Physics"] = d['GR3'].value_counts()
    gradedata["Chemistry"] = d['GR4'].value_counts()
    gradedata["Computer Science"] = d['GR5'].value_counts()
    T.insert(tk.INSERT, gradedata)
    T.config(state='normal')


# Function to choose the file
def choose_file():
    global file_path

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("Excel Files", "*.csv")])

    e.config(text=file_path)
    b.configure(state="normal")
    english.configure(state="normal")
    Computers.configure(state="normal")
    Chemistry.configure(state="normal")
    Physics.configure(state="normal")
    Maths.configure(state="normal")


# Function to perform the analysis

def perform_operation():
    T.delete("1.0", "end")
    main()


if __name__ == "__main__":
    # Creating different Tkinter widgets and defining their positions in the tkinter window
    enter = tk.Label(root, text='Choose Your file', font=11, bg="#FDFD96")
    enter.grid(row=2, column=0)

    e = tk.Label(root, text="No file selected")
    e.grid(row=2, column=3)

    choose_button = tk.Button(root, text="Choose File", command=choose_file)
    choose_button.grid(row=2, column=2)

    b = tk.Button(root, text='Go', command=perform_operation, state='disabled')
    b.grid(row=2, column=4)

    # TO DISPLAY THE GRAPH OF ENGLISH MARKS AND STUDENTS WHO HAVE THE HIGHEST MARKS IN ENGLISH
    english = tk.Button(root, text='English', command=english_marks, state='disabled')
    english.grid(row=30, column=2)

    # TO DISPLAY THE GRAPH OF COMPUTER SCIENCE MARKS AND STUDENTS WHO HAVE THE HIGHEST MARKS IN COMPUTER SCIENCE
    Computers = tk.Button(root, text='Computer Science', command=CS_marks, state='disabled')
    Computers.grid(row=30, column=3)

    # TO DISPLAY THE GRAPH OF MATHS MARKS AND STUDENTS WHO HAVE THE HIGHEST MARKS IN MATHS
    Maths = tk.Button(root, text='Mathematics', command=Maths_marks, state='disabled')
    Maths.grid(row=30, column=4)

    # TO DISPLAY THE GRAPH OF PHYSICS MARKS AND STUDENTS WHO HAVE THE HIGHEST MARKS IN PHYSICS
    Physics = tk.Button(root, text='Physics', command=Physics_marks, state='disabled')
    Physics.grid(row=30, column=5)

    # TO DISPLAY THE GRAPH OF CHEMISTRY MARKS AND STUDENTS WHO HAVE THE HIGHEST MARKS IN CHEMISTRY
    Chemistry = tk.Button(root, text='Chemistry', command=Chem_marks, state='disabled', )
    Chemistry.grid(row=30, column=6)

    plot_widget.grid(row=5, column=1, rowspan=15, columnspan=55)

    # this button will close the whole window
    b1 = tk.Button(root, text='Quit', command=root.destroy)
    b1.grid(row=100, column=50)

    # text entry box to display information
    T = tk.Text(root, fg="blue", bg="white", width=60, height=40, borderwidth=2)
    T.grid(row=6, column=0)
root.mainloop()
