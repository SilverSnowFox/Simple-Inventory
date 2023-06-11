import tkinter as tk
import pandas as pd
from tkinter import ttk, filedialog


# ============== Global Variables ==============

FILENAME = str()
COLUMNS = list()

# ============== Functions ==============

def reload_file():
    global FILENAME
    if FILENAME:
        try:
            clear_and_load_file(filename=FILENAME, message="Spreadsheet Reloaded")
            print(f'{FILENAME} reloaded.')
        except ValueError:
            label.config(text="File Could Not Be Opened")
        except FileNotFoundError:
            label.config(text="File Not Found")


def open_file():
    file = filedialog.askopenfilename(title="Open a File", filetype=(("xlxs files", ".*xlsx"),("All Files", "*.")))
    global FILENAME
    if file:
        try:
            FILENAME = r"{}".format(file)

            clear_and_load_file(filename=FILENAME, message="Spreadsheet Opened")
            print(f'{FILENAME} opened.')
        except ValueError:
            label.config(text="File Could Not Be Opened")
        except FileNotFoundError:
            label.config(text="File Not Found")


def clear_treeview():
    tree.delete(*tree.get_children())
    pass


def clear_and_load_file(filename: str, message: str):
    df = pd.read_excel(filename, header=0)
    label.config(text=message)

    # Clear all the previous data in tree
    clear_treeview()

    # Add new data in Treeview widget
    global COLUMNS 
    COLUMNS = list(df.columns)
    tree["column"] = COLUMNS
    tree["show"] = "headings"


    # For Headings iterate over the columns + an ID column
    for col in tree["column"]:
        tree.heading(col, text=col, command = lambda c = col: selc(c))   # TODO: Add the function to all the rows CAS to open SigmaAldrich for that CAS
        tree.column(col, anchor='center', width=100, stretch='no')

    # Put Data in Rows
    df_rows = df.to_numpy().tolist()
    for rows in df_rows:
        for i in range(len(rows)):
            if pd.isnull(rows[i]):
                rows[i] = ''

        print(rows)
        tree.insert("", "end", values=rows)

    tree.bind('<Double-1>', item_selected)

    tree.pack(anchor='e')

# TODO: To complete
def close_file():
    unavailable()

# TODO: To complete
def add_entry():
    # Building the Toplevel window
    new_window = tk.Toplevel(win)
    new_window.geometry('500x250')
    new_window.minsize(500, 150)
    new_window.title("Adding Entry")

    entry_bottom_frame = tk.Frame(new_window, height=50)
    entry_bottom_frame.pack(anchor='s', expand=False, fill='x', side='bottom')

    entry_top_frame = tk.Frame(new_window, height=100)
    entry_top_frame.pack(anchor='n', expand=True, fill='both', side='top')


    # Getting the columns from the main
    global COLUMNS
    entries = []

    # Creating a frame for each heading so that the Entry and the Label are aligned
    for col in COLUMNS:

        frm = tk.Frame(entry_top_frame)
        frm.pack(anchor='n', side='left')

        lb = tk.Label(frm, text=col)
        lb.pack()

        ent = tk.Entry(frm)
        ent.pack()

        entries.append(ent)
    

    def insert_main_tree():
        entry_values = []
        is_empty = True
        for entry in entries:
            value = entry.get()
            if not len(value) == 0:
                is_empty = False
            entry_values.append(value)

            entry.delete(0, 'end')
        
        if is_empty:
            top = tk.Toplevel()
            top.title("Warning")
            top.geometry("160x100")
            top.resizable(False, False)
            
            label = tk.Label(top, text="All entries cannot be empty!")
            label.pack(pady=20)
            
            ok_button = tk.Button(top, text="OK", command=top.destroy)
            ok_button.pack()
        else:
            tree.insert("", "end", values=entry_values)
        
    tk.Button(entry_bottom_frame, text='Add Entry', command=insert_main_tree).pack(anchor='center', expand=False, fill='none', side='top', padx=10, pady=20)
    


# TODO: Adds an excel sheet to the tree
def append_sheet():
    unavailable()

# TODO: To complete, make right click
def edit_entry():
    unavailable()

# TODO: To complete, make right click
def remove_entry():
    unavailable()

# TODO: Change selc to sorting the Treeview
def selc(c):
    print('c '+ str(c))

# TODO: Check what to do when selecting row
def item_selected(event):
    for selected_item in tree.selection():
        item = tree.item(selected_item)
        record = item['values']
        print(record)
        # show a message
        #tk.messagebox.showinfo(title='Information', message=','.join(str(record)))

# Temporary
def unavailable():
    tk.messagebox.showinfo(title='Warning', message='Feature unavailable.')

# ============== Tkinter App ==============

win = tk.Tk()
win.geometry('1000x500')
win.minsize(650, 300)
win.title('Excel Reader')
try:
    icon = tk.PhotoImage(file = 'img.png')
    win.iconphoto(False, icon)
except:
    print("Icon not loaded.")

"""Creating the Menu"""

mainMenu = tk.Menu(win)
win.config(menu=mainMenu)

fileMenu = tk.Menu(master = mainMenu, tearoff = False)
mainMenu.add_cascade(label = 'File', menu = fileMenu)
fileMenu.add_command(label = 'Open File', command=open_file)
fileMenu.add_command(label = 'Reload File', command=reload_file)
fileMenu.add_command(label = 'Close File', command=close_file)

entryMenu = tk.Menu(master = mainMenu, tearoff = False)
mainMenu.add_cascade(label = 'Entry', menu = entryMenu)
entryMenu.add_command(label = 'Add Entry', command=add_entry)
entryMenu.add_command(label = 'Append Sheet', command=append_sheet)


mainMenu.add_command(label = 'Help')

"""Left Menu and Window"""

leftFrame = tk.Frame(win, width=200, bg='#7FC4F5')
leftFrame.pack(anchor='w', fill='y', expand=False, side='left')
leftFrame.pack_propagate(False)

label = tk.Label(leftFrame, text='', bg='#7FC4F5', padx=10, pady=20)
label.pack(anchor='center', expand=False, fill='x', side='top')
# TODO: Might need to put the buttons inside a frame to add spacing between them 
tk.Button(leftFrame, text='Commit Changes', command=unavailable).pack(anchor='center', expand=False, fill='none', side='top')   # Commits all the changes to the sheet, saving to excel
tk.Button(leftFrame, text='Button 2', command=unavailable).pack(anchor='center', expand=False, fill='none', side='top')

"""Main Treeview"""

tree = ttk.Treeview(win, selectmode='extended')

# Adding the scrollbar before placing the tree so the tree doesn't overlap it.
scrollbary = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
tree.configure(yscroll=scrollbary.set)
scrollbary.pack(anchor='e', fill='y', expand=False, side='right')

scrollbarx = ttk.Scrollbar(win, orient='horizontal', command=tree.xview)
tree.configure(xscroll=scrollbarx.set)
scrollbarx.pack(anchor='s', fill='x', expand=False, side='bottom')

# Packing the tree after the scrollbars are loaded to not overlap it and fill all the area.
tree.pack(anchor='nw', fill='both', expand=True, side='left')


win.mainloop()