import pandas as pd
import tkinter as tk
from win32api import GetSystemMetrics

EXCEL_FILE = 'materiaalcontrole.xlsx'
KLASSEN = pd.ExcelFile(EXCEL_FILE).sheet_names
KLAS = None
TODAY = pd.Timestamp.today().strftime('%Y-%m-%d')
SCALE_FACTOR = GetSystemMetrics(1) / 1080  # Reference height is 1080px, adjust for others

# Create the main window
root = tk.Tk()
root.attributes("-fullscreen", True)
root.bind("f", lambda event: root.attributes("-fullscreen", not root.attributes("-fullscreen")))
root.bind("<Escape>", lambda event: root.attributes("-fullscreen", False))


def clear_root():
    """
    Clear the root window of all widgets.
    """
    for widget in root.winfo_children():
        widget.destroy()
    for i in range(root.grid_size()[0]):
        root.grid_rowconfigure(i, weight=0)

def select_klas():
    """
    Create a slection menu listing all classes.
    """
    global KLAS
    KLAS = None

    clear_root()

    # create a button for each class
    for ii, klas in enumerate(KLASSEN):
        frame = tk.Frame(root)
        frame.grid(row=ii, column=0)
        button = tk.Button(frame, text=klas, font=("Helvetica", int(25*SCALE_FACTOR)), bg='green', width=60, height=3, relief=tk.FLAT, 
                            command=lambda k=klas: open_klas(k))
        button.pack()
    
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=0)
    for i in range(root.grid_size()[1]):
        root.grid_rowconfigure(i, weight=1)

def open_klas(klas):
    global KLAS
    KLAS = klas

    clear_root()

    # Structurise the GUI layout
    top_frame = tk.Frame(root)
    bottom_left_frame = tk.Frame(root)
    bottom_right_frame = tk.Frame(root)
    top_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
    bottom_left_frame.grid(row=1, column=0, sticky="nsew")
    bottom_right_frame.grid(row=1, column=1, sticky="nsew")

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=2)
    root.grid_rowconfigure(0, weight=5)
    root.grid_rowconfigure(1, weight=1)
    
    # Return button
    button = tk.Button( bottom_left_frame, 
                        text=klas, 
                        font=("Helvetica", int(20*SCALE_FACTOR)), 
                        bg='green', 
                        width=30, 
                        height=3, 
                        relief=tk.FLAT, 
                        command=select_klas)
    button.pack()

    # List all students in selected class
    students = pd.read_excel(EXCEL_FILE, sheet_name=klas)

    ncols = 4
    nrows = (len(students) - 1) // ncols + 1  # Calculate number of rows needed
    buttons = []
    for ii in range(len(students)):
        if TODAY in students.columns and not pd.isna(students.at[ii, TODAY]):
            colour = 'green' if students.at[ii, TODAY] == 0 else 'orange' if students.at[ii, TODAY] == 1 else 'red'
        else:
            colour = 'green'

        student_frame = tk.Frame(top_frame)
        student_frame.grid(row=ii % nrows, column=ii // nrows)
        button = tk.Button( student_frame, 
                            text=students.at[ii,'Voornaam'], 
                            font=("Helvetica", int(18*SCALE_FACTOR)), 
                            bg=colour, 
                            width=25, 
                            height=2, 
                            relief=tk.FLAT)
        button.config(command=lambda b=button: change_colour(b))
        button.pack()
        buttons.append(button)

    for i in range(ncols):
        top_frame.grid_columnconfigure(i, weight=1)
    for i in range(nrows):
        top_frame.grid_rowconfigure(i, weight=1)

    # Confirm button
    confirm_button = tk.Button( bottom_right_frame, 
                                text="Bevestigen", 
                                font=("Helvetica", int(20*SCALE_FACTOR)),
                                bg='green',
                                width=50, 
                                height=3, 
                                relief=tk.FLAT, 
                                command=lambda buttons=buttons: save_and_show_overview(data=get_data_from_buttons(buttons)))
    confirm_button.pack()

def get_data_from_buttons(buttons):
    data = []
    for ii in range(len(buttons)):
        if buttons[ii]['bg'] == 'green':
            data.append(0)
        elif buttons[ii]['bg'] == 'orange':
            data.append(1)
        elif buttons[ii]['bg'] == 'red':
            data.append(2)
        else:
            raise ValueError(f"Invalid button color: {buttons[ii]['bg']}")
    return data
    
def save_and_show_overview(data):
    """
    Save the added data to the Excel file and visualize the sheet.
    """
    save_data(data)
    show_overview()

def save_data(data):
    """
    Save the data to the Excel file.
    """
    # Save the updates to the Excel sheet, keeping the other sheets intact
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for sheet in KLASSEN:
            students = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
            if sheet == KLAS:
                students[TODAY] = data
            students.to_excel(writer, sheet_name=sheet, index=False)

def show_overview():
    clear_root()

    # Structurise the GUI layout
    top_frame = tk.Frame(root)
    # bottom_left_frame = tk.Frame(root)
    bottom_right_frame = tk.Frame(root)
    top_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
    # bottom_left_frame.grid(row=1, column=0, sticky="nsew")
    bottom_right_frame.grid(row=1, column=1, sticky="nsew")

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=2)
    root.grid_rowconfigure(0, weight=5)
    root.grid_rowconfigure(1, weight=1)
    
    # # Return button
    # button = tk.Button( bottom_left_frame, 
    #                     text=KLAS, 
    #                     font=("Helvetica", int(20*SCALE_FACTOR)), 
    #                     bg='green', 
    #                     width=30, 
    #                     height=3, 
    #                     relief=tk.FLAT, 
    #                     command=open_klas(KLAS))
    # button.pack()

    def assign_penalty(button):
        """
        Assign a penalty to a student by increasing their 'Middagstudies' count by 1.
        """
        voornaam = button['text'].split('     ')[-1].strip()
        idx = students.index[students['Voornaam'] == voornaam][0]
        green_count = 3 * (students.at[idx,'Middagstudies'] + 1) - students.iloc[idx, 2:].sum()
        if green_count == 0:
            # increase the counter by 1
            students.at[idx, 'Middagstudies'] += 1
        elif green_count == 3 and students.at[idx, 'Middagstudies'] > 0:
            # reduce the counter by 1
            students.at[idx, 'Middagstudies'] -= 1
        else:
            # do nothing
            return

        # Save the update to the Excel sheet
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for sheet in KLASSEN:
                if sheet == KLAS:
                    tab = students
                else:
                    tab = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
                tab.to_excel(writer, sheet_name=sheet, index=False)
        show_overview()
        
    # List all students in selected class
    students = pd.read_excel(EXCEL_FILE, sheet_name=KLAS)

    ncols = 4
    nrows = (len(students) - 1) // ncols + 1  # Calculate number of rows needed
    buttons = []
    for ii in range(len(students)):
        green_count = 3 * (students.at[ii,'Middagstudies'] + 1) - students.iloc[ii, 2:].sum()

        student_frame = tk.Frame(top_frame)
        student_frame.grid(row=ii % nrows, column=ii // nrows)

        # penalty button
        button = tk.Button( student_frame, 
                            text=f"[{students.at[ii,'Middagstudies']}]     {students.at[ii,'Voornaam']}", 
                            font=("Helvetica", int(18*SCALE_FACTOR)), 
                            width=20, 
                            height=2, 
                            relief=tk.RAISED)
        button.config(command=lambda b=button: assign_penalty(button=b))
        button.pack()
        buttons.append(button)
        
        # life bars
        for jj in [1,2,3]:
            colour = 'red' if jj > green_count else 'green'
            bar_frame = tk.Frame(student_frame, bg=colour, width=int(SCALE_FACTOR*95), height=int(SCALE_FACTOR*30))
            bar_frame.pack(side=tk.LEFT, padx=1)

    for i in range(ncols):
        top_frame.grid_columnconfigure(i, weight=1)
    for i in range(nrows):
        top_frame.grid_rowconfigure(i, weight=1)

    # Confirm button
    confirm_button = tk.Button( bottom_right_frame, 
                                text="Afsluiten", 
                                font=("Helvetica", int(20*SCALE_FACTOR)),
                                bg='green',
                                width=50, 
                                height=3, 
                                relief=tk.FLAT, 
                                command=root.destroy)
    confirm_button.pack()

def change_colour(button):
    """
    Change the color of the button. The colour cycles through green, orange and red.
    """
    if button['bg'] == 'green':
        button['bg'] = 'orange'
    elif button['bg'] == 'orange':
        button['bg'] = 'red'
    elif button['bg'] == 'red':
        button['bg'] = 'green'
    else:
        raise ValueError(f"Invalid button color: {button['bg']}")

select_klas()
root.mainloop()
