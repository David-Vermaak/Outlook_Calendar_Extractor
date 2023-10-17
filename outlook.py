import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import win32com.client
import sv_ttk


def select_date_range():

    # Label and Entry widgets for start and end dates
    start_date_label = ttk.Label(content_frame, text="\nStart Date (YYYY-MM-DD):")
    start_date_label.pack()
    start_date_entry = ttk.Entry(content_frame)
    start_date_entry.pack()

    end_date_label = ttk.Label(content_frame, text="\nEnd Date (YYYY-MM-DD):")
    end_date_label.pack()
    end_date_entry = ttk.Entry(content_frame)
    end_date_entry.pack()
    

    # Button to trigger event extraction
    extract_button = ttk.Button(content_frame, text="Extract Calendar Events", command=lambda: get_data(start_date_entry, end_date_entry))
    extract_button.pack(pady=15)


def get_data(start_date_entry, end_date_entry):

    start_date = start_date_entry.get()
    end_date = end_date_entry.get()

    clear_window()
    extract_calendar_events(start_date, end_date)



def extract_calendar_events(start_date, end_date):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # 9 represents the calendar folder
        appointments = calendar.Items
        start_filter = f"[Start] >= '{start_date} 00:00 AM'"
        end_filter = f"[End] <= '{end_date} 11:59 PM'"
        events = appointments.Restrict(start_filter + " AND " + end_filter)
        display_data(events)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to retrieve calendar events: {str(e)}")
        return []

def display_data(events):

    # Create a ttk.Treeview widget
    tree = ttk.Treeview(content_frame, columns=("Subject", "Start Time"), show="headings")

    # Define column headings
    tree.heading("Subject", text="Subject")
    tree.heading("Start Time", text="Start Time")

    # Set column widths
    tree.column("Subject", width=200)
    tree.column("Start Time", width=200)

    # Insert data from the events into the Treeview
    for event in events:
        print(f"Subject: {event.Subject}")
        print(f"Start Time: {event.Start}")
        print(f"End Time: {event.End}")
        print(f"Location: {event.Location}")
        print(f"Description: {event.Body}\n")
        print("\n////////////////////////////////////////////////////////////////\n")
        tree.insert("", "end", values=(event.Subject, event.Start))

    # Create horizontal scrollbar
    hsb = ttk.Scrollbar(content_frame, orient="horizontal", command=tree.xview)
    tree.configure(xscrollcommand=hsb.set)

    # Pack the Treeview and scrollbar
    tree.pack(fill="both", expand=True)
    hsb.pack(side="bottom", fill="x")




#information textbox
def info():

    clear_window()

    title = "Program Information"
    info_text = """
        Welcome to the Outlook Calendar Extractor
        How to Use:

        1.)Enter the date range you want to get events from
        2.)push the button
        

    """
    
    # Title Label
    textbox_label = ttk.Label(content_frame, text=title)
    textbox_label.pack(pady=15)
    
    frame = ttk.Frame(content_frame)
    frame.pack(fill="both", expand=False)
    
    text_area = tk.Text(frame, wrap="none", font=menu_font)
    text_area.pack(padx=15, pady=15, fill="both", expand=True)
    
    # Inserting Text which is read only
    text_area.insert(tk.INSERT, info_text)


def clear_window():
    # Destroy everything except Menu, content frame, and boolian values
    for widget in content_frame.winfo_children():
        if widget != Menu_button and widget != Info_button and widget != Exit_button and widget != switch and widget != content_frame:
            widget.destroy()


#stops program
def exit_program():
    root.destroy()




#Main program config:

root = tk.Tk()

#window header title
root.title("Outlook Calendar Event Extractor")



#setting tkinter window size (fullscreen windowed)
root.state('zoomed')

# Set the minimum width and height for the window
root.minsize(400, 300)


menu_font = ("Arial", 14)

# Create a style for the Menu button
menu_button_style = ttk.Style()
menu_button_style.configure("Menu.TButton", font=menu_font, foreground='white', background='#2589bd')

# Create a style for the Info button
info_button_style = ttk.Style()
info_button_style.configure("Toggle.TButton", font=menu_font, foreground='white', background='#5C946E')

# Create a style for the Exit button
exit_button_style = ttk.Style()
exit_button_style.configure("Toggle.TButton", font=menu_font, foreground='white', background='#B3001B')


# Add Main Menu button at the top left of the window
Menu_button = ttk.Button(root, text="Menu", command=select_date_range, style="Accent.TButton")
Menu_button.place(x=15, y=12)

# Add an Info button
Info_button = ttk.Button(root, text="Info", command=info, style="Info.TButton")
Info_button.place(x=15, y=66)  # Position below Menu_button


# Add an Exit button
Exit_button = ttk.Button(root, text="Exit", command=exit_program, style="Exit.TButton")
Exit_button.place(x=15, y=120)  # Position below Info_button


# Add a darkmode toggle switch
switch = ttk.Checkbutton(text="Light mode", style="Switch.TCheckbutton", command=sv_ttk.toggle_theme)
switch.place(x=15, y=260)  # Position below toggle switch


# Create a content frame for the main content area
content_frame = ttk.Frame(root)
content_frame.place(x=175, y=5, relwidth=.8, relheight=.95)  # Use relative dimensions for expansion


# Bind the window close event to the exit_program function
root.protocol("WM_DELETE_WINDOW", exit_program)



import ctypes as ct
#dark titlebar - ONLY WORKS IN WINDOWS 11!!
def dark_title_bar(window):
    """
    MORE INFO:
    https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/ne-dwmapi-dwmwindowattribute
    """
    window.update()
    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
    set_window_attribute = ct.windll.dwmapi.DwmSetWindowAttribute
    get_parent = ct.windll.user32.GetParent
    hwnd = get_parent(window.winfo_id())
    rendering_policy = DWMWA_USE_IMMERSIVE_DARK_MODE
    value = 2
    value = ct.c_int(value)
    set_window_attribute(hwnd, rendering_policy, ct.byref(value), ct.sizeof(value))


dark_title_bar(root)

#trying to get the taskbar icon to work
myappid = 'File.Parser.V2' # arbitrary string
ct.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


# Theme
sv_ttk.set_theme("dark")

# Initial function
select_date_range()

#this is the loop that keeps the window persistent 
root.mainloop()
