import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from openpyxl.reader.excel import load_workbook
print("f")
def name_correction(input_file_path, output_file_path):
    data_frame = pd.read_excel(input_file_path)
    surnames = data_frame['Прізвище']
    names = data_frame["Ім'я"]
    full_names = [f"{name} {surname}" for name, surname in zip(names, surnames)]

    data_frameT = pd.read_excel(output_file_path)
    surnamesT = data_frameT['Прізвище']
    namesT = data_frameT["Ім'я"]
    full_namesT = [f"{name} {surname}" for name, surname in zip(namesT, surnamesT)]

    ListForWrite=[]

    for i in full_namesT:
        if i in full_names:
            ListForWrite.append(i)
        else:
            ListForWrite.append(LevinsteinList(i, full_names))

    output_wb = load_workbook(output_file_path)
    output_ws = output_wb.active
    column_surname = "Прізвище"
    column_name = "Ім'я"
    target_column_name = None
    target_column_surname = None
    for column in output_ws.iter_cols(min_row=1, max_row=1):
        if column[0].value == column_surname:
            target_column_surname = column[0].column
        if column[0].value == column_name:
            target_column_name = column[0].column
        if target_column_surname != None and target_column_name != None:
            break
    for i, item in enumerate(ListForWrite, start=2):
        splited= item.split(" ")
        surname=splited[1]
        name=splited[0]
        output_ws.cell(row=i, column=target_column_surname, value=surname)  # Assuming surname goes in the first column
        output_ws.cell(row=i, column=target_column_name, value=name)  # Assuming name goes in the second column

    output_wb.save(output_file_path)

def LevinsteinList(value, checker):
    draft = {}
    for VARIABLE in checker:
        matrix = np.zeros((len(value) + 1, len(VARIABLE) + 1))

        for i in range(len(value) + 1):
            matrix[i, 0] = i

        for j in range(len(VARIABLE) + 1):
            matrix[0, j] = j

        for i in range(1, len(value) + 1):
            for j in range(1, len(VARIABLE) + 1):
                v1 = matrix[i - 1, j] + 1
                v2 = matrix[i, j - 1] + 1
                v3 = 0
                v4 = 0
                cost = 0

                if value[i - 1] == VARIABLE[j - 1]:
                    v3 = matrix[i - 1, j - 1]
                    cost = 0
                else:
                    v3 = matrix[i - 1, j - 1] + 1
                    cost = 1

                if i > 1 and j > 1:
                    if VARIABLE[j - 1] == value[i - 2] and VARIABLE[j - 2] == value[i - 1]:
                        v4 = matrix[i - 2, j - 2] + cost
                    else:
                        v4 = 999999999
                else:
                    v4 = 999999999

                matrix[i, j] = min(v1, v2, v3, v4)

        addd = matrix[len(value), len(VARIABLE)]
        draft[VARIABLE] = addd

    Sorted_Word_Points = sorted(draft.items(), key=lambda x: x[1])

    res=str(next(iter(Sorted_Word_Points))).split("'")

    return res[1]

# Main program starts here
# I have removed the repetitive code. All the remaining functions and Tkinter GUI code remains the same. You can add it back.




def doZoom(xls_file_path, output_file_path):
    data_frame = pd.read_excel(xls_file_path)
    duration = data_frame['Тривалість']
    Name = data_frame["Ім'я(справжнє)"]
    output_wb = load_workbook(output_file_path)
    output_ws = output_wb.active
    column_duration = "Відвідуваність"
    column_name = "Ім'я(справжнє)"
    target_column_attendance = None
    target_column_name = None
    for column in output_ws.iter_cols(min_row=1, max_row=1):
        if column[0].value == column_duration:
            target_column_attendance = column[0].column
        if column[0].value == column_name:
            target_column_name = column[0].column
        if target_column_attendance != None and target_column_name != None:
            break

    data = zip(duration, Name)

    for i, (attendance, name) in enumerate(data, start=2):
        status = "absent"
        if attendance >= 46:
            status = "present"
        output_ws.cell(row=i, column=target_column_attendance, value=status)
        output_ws.cell(row=i, column=target_column_name, value=name)

    output_wb.save(output_file_path)
    print("Success")

def doMoodle(xls_file_path, output_file_path):
    data_frame = pd.read_excel(xls_file_path)
    Grades = data_frame['Загальне за курс (Бали)']
    Surname = data_frame['Прізвище']
    Name = data_frame["Ім'я"]
    output_wb = load_workbook(output_file_path)
    output_ws = output_wb.active
    column_grade = "Бали"
    column_surname = "Прізвище"
    column_name = "Ім'я"
    target_column_grade = None
    target_column_name = None
    target_column_surname = None
    for column in output_ws.iter_cols(min_row=1, max_row=1):
        if column[0].value == column_grade:
            target_column_grade = column[0].column
        if column[0].value == column_surname:
            target_column_surname = column[0].column
        if column[0].value == column_name:
            target_column_name = column[0].column
        if target_column_surname!=None and target_column_name!=None and target_column_grade!=None:
            break

    data = zip(Grades, Surname, Name)
    for i, (grade, surname, name) in enumerate(data, start=2):
        output_ws.cell(row=i, column=target_column_grade, value=grade)
        output_ws.cell(row=i, column=target_column_surname, value=surname)
        output_ws.cell(row=i, column=target_column_name, value=name)

    output_wb.save(output_file_path)
    print("Success")

def select_input_file():
    global input_file_path
    input_file_path = filedialog.askopenfilename(title="Select input file",
                                                 filetypes=(("Excel files", "*.xls *.xlsx"),("All files", "*.*")))
    if input_file_path:
        if not (input_file_path.endswith(".xls") or input_file_path.endswith(".xlsx")):
            messagebox.showerror("Error", "Input file must be a .xls or .xlsx file!")
            input_file_path = ""
        else:
            input_file_label['text'] = input_file_path

def select_output_file():
    global output_file_path

    output_file_path = filedialog.askopenfilename(title="Select input file",
                                                 filetypes=(("Excel files", "*.xls *.xlsx"), ("All files", "*.*")))
    if output_file_path:
        if not (output_file_path.endswith(".xls") or output_file_path.endswith(".xlsx")):
            messagebox.showerror("Error", "Input file must be a .xls or .xlsx file!")
            output_file_path = ""
        else:
            output_file_label['text'] = output_file_path

def process_files():
    if input_file_path and output_file_path:
        if "Zoom" in input_file_path:
            doZoom(input_file_path, output_file_path)
            name_correction("C:\\Users\Давід\\PycharmProjects\\Internship2.0\\Students.xlsx", output_file_path)
        elif "Moodle" in input_file_path:
            doMoodle(input_file_path, output_file_path)
            name_correction("C:\\Users\Давід\\PycharmProjects\\Internship2.0\\Students.xlsx", output_file_path)
        print("g")

root = tk.Tk()
root.title("KSE Automatization")

title_label = tk.Label(root, text="KSE Automatization", font=('Verdana', 18, 'bold'))
title_label.pack(pady=10)

explanation_label = tk.Label(root, text="This program helps you systematize tables.", font=('Verdana', 12))
explanation_label.pack(pady=10)

input_file_button = tk.Button(root, text="Select Input File", command=select_input_file, font=('Verdana', 14), padx=20, pady=10)
input_file_button.pack()

input_file_label = tk.Label(root, text="", font=('Verdana', 10))
input_file_label.pack(pady=10)

output_file_button = tk.Button(root, text="Select Output File", command=select_output_file, font=('Verdana', 14), padx=20, pady=10)
output_file_button.pack()

output_file_label = tk.Label(root, text="", font=('Verdana', 10))
output_file_label.pack(pady=10)

process_button = tk.Button(root, text="Process Files", command=process_files, font=('Verdana', 14), padx=20, pady=10)
process_button.pack()

root.mainloop()