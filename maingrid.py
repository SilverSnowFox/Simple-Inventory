#!/usr/bin/env py

"""A simple chemical inventory program centred around excel files for ease of use.

Author: @SilverSnowFox
"""

import tkinter as tk
import pandas as pd
import webbrowser
import copy
from tkinter import ttk, filedialog


# ============== Global Variables ==============

FILENAME = str()
COLUMNS = list()
SHEETNAME = str()
BACKGROUND_HEX = '#7FC4F5'

# ============== Functions ==============

def reload_file():

    global FILENAME

    if FILENAME:
        try:
            clear_and_load_file()
            label.config(text = "Spreadsheet Reloaded")
            print(f'{FILENAME} reloaded.')

        except ValueError:
            label.config(text="File Could Not Be Opened")
            print(f'{FILENAME} could not be opened.')

        except FileNotFoundError:
            label.config(text="File Not Found")
            print(f'{FILENAME} could not be found.')


def open_file():

    global FILENAME

    file = filedialog.askopenfilename(title="Open a File", filetype=(("xlxs files", ".*xlsx"),("All Files", "*.")))

    if file:
        try:
            FILENAME = r"{}".format(file)

            clear_and_load_file()
            label.config(text = "Spreadsheet Opened")
            print(f'{FILENAME} opened.')

        except ValueError:
            label.config(text="File Could Not Be Opened")
            print(f'{FILENAME} could not be opened.')

        except FileNotFoundError:
            label.config(text="File Not Found")
            print(f'{FILENAME} could not be found.')


def clear_treeview(tree: ttk.Treeview):
    tree.delete(*tree.get_children())


def clear_and_load_file():
    """Clears the main window Tree and inserts values of the newly selected sheet."""

    global SHEETNAME
    global COLUMNS

    xl = pd.ExcelFile(FILENAME)
    SHEETNAME = xl.sheet_names[0]   # Only gets the first sheet
    df = xl.parse(SHEETNAME)

    # Clear all the previous data in tree
    clear_treeview(tree=tree)

    # Add new data in Treeview widget
    COLUMNS = list(df.columns)
    tree["column"] = COLUMNS
    tree["show"] = "headings"

    # For Headings iterate over the columns + an ID column
    for col in tree["column"]:
        tree.heading(col, text=col, command = lambda _col = col: treeview_sort_column(tree, _col, False))
        tree.column(col, anchor='center', width=100, stretch='no')

    # Put Data in Rows
    df_rows = df.to_numpy().tolist()

    for rows in df_rows:
        for i in range(len(rows)):
            if pd.isnull(rows[i]):
                rows[i] = ''

        tree.insert("", "end", values=rows)

    tree.bind('<Double-1>', item_selected)
    tree.bind('<Button-3>', popup)

    tree.pack(anchor='e')


def add_entry():
    # Building the Toplevel window
    new_window = tk.Toplevel(win)
    new_window.geometry('500x100')
    new_window.minsize(500, 100)
    new_window.title("Adding Entry")

    entry_bottom_frame = tk.Frame(new_window, height=50)
    entry_bottom_frame.pack(anchor='s', expand=False, fill='x', side='bottom')

    entry_top_frame = tk.Frame(new_window, height=50)
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
    
    # Making sure it doesn't close the Toplevel when adding an entry to add multiple entries
    def insert_main_tree():
        entry_values = []
        is_empty = True
        for entry in entries:
            value = entry.get()
            if not len(value) == 0:
                is_empty = False
            entry_values.append(value)

        if is_empty:
            top = tk.Toplevel()
            top.title("Warning")
            top.geometry("160x100")
            top.resizable(False, False)
            
            label_empty = tk.Label(top, text = "All entries cannot be empty!")
            label_empty.pack(pady=20)
            
            ok_button = tk.Button(top, text = "OK", command = top.destroy)
            ok_button.pack()
        else:
            tree.insert("", "end", values = entry_values)
            label.config(text = "Entry Added")
            
    tk.Button(entry_bottom_frame, text='Add Entry', command=insert_main_tree).pack(anchor='center', expand=False, fill='none', side='top', padx=10, pady=20)
    

def append_sheet():
    
    # Getting the file
    file = filedialog.askopenfilename(title="Open a File", filetype=(("xlxs files", ".*xlsx"),("All Files", "*.")))
    if file:
        try:
            filename = r"{}".format(file)
            xl = pd.ExcelFile(filename)
            sheetname = xl.sheet_names[0]
            df = xl.parse(sheetname)
            print(f'{filename} opened.')

        except ValueError:
            label.config(text="File Could Not Be Opened")
            print(f'{filename} could not be opened.')

        except FileNotFoundError:
            label.config(text="File Not Found")
            print(f'{filename} not found.')

    # Building the Toplevel window
    new_window = tk.Toplevel(win)
    new_window.geometry('500x250')
    new_window.minsize(500, 150)
    new_window.title("Appending Sheet")

    entry_bottom_frame = tk.Frame(new_window, height=50)
    entry_bottom_frame.pack(anchor='s', expand=False, fill='x', side='bottom')

    entry_top_frame = tk.Frame(new_window, height=100)
    entry_top_frame.pack(anchor='n', expand=True, fill='both', side='top')

    appending_tree = ttk.Treeview(entry_top_frame, selectmode='extended')
    appending_tree["column"] = list(df.columns)
    appending_tree["show"] = "headings"

    # For Headings iterate over the columns + an ID column
    for col in appending_tree["column"]:
        appending_tree.heading(col, text=col, command = lambda c = col: treeview_sort_column(c))   # TODO: Add the function to all the rows CAS to open SigmaAldrich for that CAS
        appending_tree.column(col, anchor='center', width=100, stretch='no')

    # Put Data in Rows
    df_rows = df.to_numpy().tolist()

    for rows in df_rows:
        for i in range(len(rows)):
            if pd.isnull(rows[i]):
                rows[i] = ''

        appending_tree.insert("", "end", values=rows)

    # Adding the scrollbar before placing the tree so the tree doesn't overlap it.
    appending_scrollbary = ttk.Scrollbar(entry_top_frame, orient='vertical', command=appending_tree.yview)
    appending_tree.configure(yscroll=appending_scrollbary.set)
    appending_scrollbary.pack(anchor='e', fill='y', expand=False, side='right')

    appending_scrollbarx = ttk.Scrollbar(entry_top_frame, orient='horizontal', command=appending_tree.xview)
    appending_tree.configure(xscroll=appending_scrollbarx.set)
    appending_scrollbarx.pack(anchor='s', fill='x', expand=False, side='bottom')

    appending_tree.pack(anchor='n', fill='both', expand=True, side='left')

    def insert_main_tree():
        global COLUMNS

        lower_main_columns = [col.lower() for col in COLUMNS]
        lower_appending_columns = [col.lower() for col in appending_tree['column']]

        if lower_main_columns == lower_appending_columns:
            for selected_item in appending_tree.get_children():
                tree.insert("", "end", values=tree.item(selected_item)['values'])
            
            new_window.destroy()
        else:
            tk.messagebox.showinfo(title='Warning', message='Sheet columns don\'t match with main.')

    tk.Button(entry_bottom_frame, text='Confirm', command=insert_main_tree).pack(anchor='center', expand=False, fill='none', side='top', padx=10, pady=20)


def edit_entry():
    for selected_item in tree.selection():
        val = tree.item(selected_item)['values']
        
        # Building the Toplevel window
        new_window = tk.Toplevel(win)
        new_window.geometry('500x100')
        new_window.minsize(500, 100)
        new_window.title("Editing Entry")

        entry_bottom_frame = tk.Frame(new_window, height=50)
        entry_bottom_frame.pack(anchor='s', expand=False, fill='x', side='bottom')

        entry_top_frame = tk.Frame(new_window, height=50)
        entry_top_frame.pack(anchor='n', expand=True, fill='both', side='top')

        # Getting the columns from the main
        global COLUMNS
        entries = []

        # Creating a frame for each heading so that the Entry and the Label are aligned
        for i in range(len(COLUMNS)):

            frm = tk.Frame(entry_top_frame)
            frm.pack(anchor='n', side='left')

            lb = tk.Label(frm, text=COLUMNS[i])
            lb.pack()

            ent = tk.Entry(frm)
            ent.insert(0, val[i])
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
            
            if is_empty:
                top = tk.Toplevel()
                top.title("Warning")
                top.geometry("160x100")
                top.resizable(False, False)
                
                label_empty = tk.Label(top, text="All entries cannot be empty!")
                label_empty.pack(pady=20)
                
                ok_button = tk.Button(top, text="OK", command=top.destroy)
                ok_button.pack()
            else:
                tree.insert("", "end", values=entry_values)
                label.config(text="Entry Edited")
                new_window.destroy()
            
        tk.Button(entry_bottom_frame, text='Edit Entry', command=insert_main_tree).pack(anchor='center', expand=False, fill='none', side='top', padx=10, pady=20)


def remove_entry():
    [tree.delete(selected_item) for selected_item in tree.selection()]


def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)

    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    # reverse sort next time
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))


def search_sigma_cas():

    matching = [s for s in COLUMNS if "cas" in s.lower()]
    if len(matching) == 0:
        tk.messagebox.showinfo(title='Warning', message='No CAS column found.')
    elif len(matching) > 1:
        tk.messagebox.showinfo(title='Warning', message='More than one (1) CAS columns found.')
    else:
        indx = COLUMNS.index(matching[0])

        for selected_item in tree.selection():
            cas = tree.item(selected_item)['values'][indx]
            webbrowser.open_new_tab(f"https://www.sigmaaldrich.com/CA/en/search/{cas}?focus=products&page=1&perpage=30&sort=relevance&term={cas}&type=cas_number")


def search_sigma_name():

    matching = [s for s in COLUMNS if "name" in s.lower()]
    if len(matching) == 0:
        tk.messagebox.showinfo(title='Warning', message='No name column found.')
    elif len(matching) > 1:
        tk.messagebox.showinfo(title='Warning', message='More than one (1) name columns found.')
    else:
        indx = COLUMNS.index(matching[0])

        for selected_item in tree.selection():
            name = tree.item(selected_item)['values'][indx]
            name_other = copy.deepcopy(name)
            name.lower().replace(' ', '-')
            name_other.lower().replace(' ', '%20')
            webbrowser.open_new_tab(f"https://www.sigmaaldrich.com/CA/en/search/{name}?focus=products&page=1&perpage=30&sort=relevance&term={name_other}&type=product_name")


def item_selected():
    [print(tree.item(selected_item)['values']) for selected_item in tree.selection]


def unselect():
    [tree.selection_remove(item) for item in tree.selection()]


def commit():

    rows = [tree.item(row, 'values') for row in tree.get_children()]
    rows = list(map(list, rows))

    temp_df = pd.DataFrame(rows, columns=COLUMNS)
    temp_df.to_excel(FILENAME, SHEETNAME, header=COLUMNS, index=False)

    tk.messagebox.showinfo(title='Commit', message='File saved.')


def popup(event):
    iid = tree.identify_row(event.y)
    if iid:
        tree.selection_set(iid)
        tree_contextMenu.post(event.x_root, event.y_root)

# TODO: To make
def help():
    unavailable()

# Temporary
def unavailable():
    tk.messagebox.showinfo(title='Warning', message='Feature unavailable.')

# ============== Tkinter App ==============

win = tk.Tk()
win.geometry('1000x500')
win.minsize(650, 300)
win.title('Excel Reader')

"""Creating the Top Menu"""

mainMenu = tk.Menu(win)
win.config(menu=mainMenu)

fileMenu = tk.Menu(master = mainMenu, tearoff = False)
mainMenu.add_cascade(label = 'File', menu = fileMenu)
fileMenu.add_command(label = 'Open File', command=open_file)
fileMenu.add_command(label = 'Reload File', command=reload_file)

entryMenu = tk.Menu(master = mainMenu, tearoff = False)
mainMenu.add_cascade(label = 'Entry', menu = entryMenu)
entryMenu.add_command(label = 'Add Entry', command=add_entry)
entryMenu.add_command(label = 'Add From Sheet', command=append_sheet)
entryMenu.add_command(label = 'Delete Entry', command=remove_entry)

SigmaMenuTop = tk.Menu(master = entryMenu, tearoff=False)
entryMenu.add_cascade(label='Search SigmaAldrich', menu=SigmaMenuTop)
SigmaMenuTop.add_command(label='CAS', command=search_sigma_cas)
SigmaMenuTop.add_command(label='Name', command=search_sigma_name)

mainMenu.add_command(label = 'Help', command=unavailable)    # TODO: To Make

"""Main TreeView Context Menu"""

tree_contextMenu = tk.Menu(win, tearoff=0)
tree_contextMenu.add_command(label='Edit', command=edit_entry)
tree_contextMenu.add_command(label='Delete', command=remove_entry)
tree_contextMenu.add_command(label='Unselect', command=unselect)

SigmaMenu = tk.Menu(master = tree_contextMenu, tearoff=False)
tree_contextMenu.add_cascade(label='Search Sigma', menu=SigmaMenu)
SigmaMenu.add_command(label='CAS', command=search_sigma_cas)
SigmaMenu.add_command(label='Name', command=search_sigma_name)

"""Left Menu and Window"""

leftFrame = tk.Frame(win, width=170, bg=BACKGROUND_HEX)
leftFrame.pack(anchor='w', fill='y', expand=False, side='left')
leftFrame.pack_propagate(False)

label = tk.Label(leftFrame, text='', bg=BACKGROUND_HEX, padx=10, pady=20)
label.pack(anchor='center', expand=False, fill='x', side='top')
tk.Button(leftFrame, text='Commit Changes', command=commit).pack(anchor='center', expand=False, fill='none', side='top')   # Commits all the changes to the sheet, saving to excel

"""Main Treeview"""

tree = ttk.Treeview(win, selectmode='extended')

# Adding the scrollbar before placing the tree so the tree doesn't overlap it.
scrollbary = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
tree.configure(yscroll=scrollbary.set)
scrollbary.pack(anchor='e', fill='y', expand=False, side='right')

scrollbarx = ttk.Scrollbar(win, orient='horizontal', command=tree.xview)
tree.configure(xscroll=scrollbarx.set)
scrollbarx.pack(anchor='s', fill='x', expand=False, side='bottom')

tree.pack(anchor='nw', fill='both', expand=True, side='left')


if __name__ == '__main__':
    win.mainloop()