import tkinter as tk
from tkinter import filedialog, Text, Label, Entry, messagebox
import pandas as pd

# v1.03 let's try to add variable for  different column,

def integerify(input_text): #on Win11 result somehow shows in float. So added integerifying function if available
    try:
        return int(input_text)
    except:
        return input_text


# Global variable to store the dataframe
df = None

desigPaitentID = [] # 첫 선언 global variable
def IDseeker(input_sheets, givenDocName, col_index): # col_index 는 list array를 받아야할듯. 그냥 integer도 될테고,,?(걍 int는 안되는듯) [7] 이렇게 주면 된다. 
    desigPaitentID = []
    for sheet_name, data in input_sheets.items():
        print(f"Data from sheet: {sheet_name}")
        if "진료조" in sheet_name or "응콜" in sheet_name: # 진료조 sheet 에서만 에러가나서 예외처리하도록
            continue
        try:
            for col_indice in col_index:
                for i_counter in range(len(data.iloc[:, col_indice])):
                    patientID = integerify(data.iloc[i_counter, 1]) # patientID(intovetID)
                    patientID2 = integerify(data.iloc[i_counter, 2]) # infinit ID
                    docName = data.iloc[i_counter, col_indice] # 담당 (column H == 7, cause A == 0), 안과탭도 동일. H가 담당.

                    if docName == givenDocName and pd.isna(patientID) == False:
                        print(str(patientID) + ", " + str(docName))
                        desigPaitentID.append(patientID)

                    if docName == givenDocName and pd.isna(patientID2) == False:
                        print(str(patientID2) + ", " + str(docName))
                        desigPaitentID.append(patientID2)
        except IndexError:
            # Handle the error if the 8th column does not exist
            # pass # 다른건 다 비슷한데 최태웅 검사했을때 pass와 print 시 순서가 달라... 이거는 왜그런지 좀 연구해보자. #이거라서 다른게 아니라... 그냥 할때마다 좀 달라지나봄
            print(f"Error: Sheet '{sheet_name}' does not have an 8th column.") # 여기는 있어야한다 혹은 pass나 continue
    print(desigPaitentID)
    print(len(desigPaitentID))
    desigPaitentID_noDuple = list(set(desigPaitentID)) # 애초에 중복제거를 하고 리턴해준다
    print(desigPaitentID_noDuple)
    print(len(desigPaitentID_noDuple))
    return desigPaitentID_noDuple


def load_excel():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        df = pd.read_excel(file_path, sheet_name=None)
        messagebox.showinfo("Information", "Excel file loaded successfully.")
        

def search_text():
    inputDocName = entry.get()
    if df is not None and inputDocName:
        IDs = IDseeker(df, inputDocName, [4,5,7]) # col_index 주의
        IDs_string = ', '.join(map(str, IDs))
        IDs_string_noEnter = IDs_string.replace("\n", "") # \n 뜨는 에러 제거
        text_area.delete('1.0', tk.END)  # Clear the text area before displaying new content
        text_area.insert(tk.END, IDs_string_noEnter)  # Display the first column
    else:
        messagebox.showwarning("Warning", "Please load an Excel file and enter the name of the radiologist to search.")


def copy_to_clipboard():
    root.clipboard_clear()
    root.clipboard_append(text_area.get("1.0", tk.END))



root = tk.Tk()
root.title("Case without study comment finder")
root.geometry("450x620")

w=tk.Label(root,text="\n <How To Use>\n1. Input 진료순서표 as an excel file .\n2, Enter the name of the radiologist for your search.\n3. Patient IDs of the selected radiologist will be shown.\n4. Copy and paste these IDs into \"Patient ID\" section of INFINITT browser. \n5. Ensure the \"Study Comment\" section is set to \"null\".\n6. Cases without study comment will be displayed.")
w.pack()

open_file_btn = tk.Button(root, text="1. Open Excel File", font='sans 10 bold', fg='green', command=load_excel)
open_file_btn.pack(pady=20)


# Entry widget for input
Label(root, text="Enter Radiologist to Search:").pack(pady=(20, 0))
entry = Entry(root)
entry.pack(pady=5)

# Search Button
search_button = tk.Button(root, text="2. Search", font='sans 10 bold', command=search_text)
search_button.pack()


text_area = Text(root, bd=1, height=15, width=50)
text_area.pack(pady=20)

copy_btn = tk.Button(root,anchor='e', text="3. Copy to Clipboard", command=copy_to_clipboard)
copy_btn.pack(pady=10)


def popup():
    messagebox.showinfo("Credit", "For fellow radiologists. \n\n             DEC.19.2023 Chief")

button2 = tk.Button(root, text='Credit', command=popup)
button2.pack(side=tk.LEFT, padx=2, pady=2)


button_quit = tk.Button(root, text="Exit Program",  command=root.quit)
button_quit.pack(side=tk.RIGHT, padx=2, pady=2)


root.mainloop()
# CT, MRI 탭은 F column이 담당. (index == 5)
# 외부초진 탭은 E column이 담당.(index == 4)

'''
search_text = entry.get()
if not search_text:  # Check if the search text is not empty
    text_area.delete('1.0', tk.END)
    text_area.insert(tk.END, "Please enter text to search.\n")
    return
'''

'''
def popup():
    # Create a top-level window
    popup_window = Toplevel()
    popup_window.title("Credit")

    # Create a Label for the main message
    main_message = Label(popup_window, text="For my fellow radiologists.")
    main_message.pack(fill='x', expand=True)

    # Create a Label for "Chief" and align it to the right
    chief_label = Label(popup_window, text="Chief", anchor='e', justify='right')
    chief_label.pack(fill='both', expand=True)

button = tk.Button(root, text="Show Popup", command=popup)
button.pack()
'''
