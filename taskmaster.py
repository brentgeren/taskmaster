import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import ctypes
from ctypes import byref, c_int, sizeof, windll, c_bool, Structure, c_uint
from create_icon import create_checkmark_icon
import csv
from datetime import datetime

TASKS_FILE = "tasks.json"

class WINDOWCOMPOSITIONATTRIBDATA(Structure):
    _fields_ = [
        ('Attrib', c_int),
        ('Data', c_int),
        ('SizeOfData', c_uint)
    ]

def set_window_dark_title_bar(window):
    """
    Set the title bar to dark mode on Windows using composition attributes
    """
    try:
        # Define constants
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        WCA_USEDARKMODECOLORS = 26

        # Try the newer Windows 11 method first
        set_window_attribute = windll.dwmapi.DwmSetWindowAttribute
        get_parent = windll.user32.GetParent
        hwnd = get_parent(window.winfo_id())
        rendering_policy = c_int(2)
        set_window_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                           byref(rendering_policy), sizeof(rendering_policy))

        # Try the Windows 10 method as fallback
        data = WINDOWCOMPOSITIONATTRIBDATA()
        data.Attrib = WCA_USEDARKMODECOLORS
        data.SizeOfData = sizeof(c_int)
        data.Data = c_int(1)
        
        set_window_composition_attribute = windll.user32.SetWindowCompositionAttribute
        set_window_composition_attribute(hwnd, byref(data))
    except:
        pass

class TaskMaster:
    def __init__(self, root):
        self.root = root
        
        # Create and set the application icon
        try:
            icon_path = create_checkmark_icon()
            # Set icon for the window title bar
            self.root.iconbitmap(icon_path)
            # Set icon for the taskbar/application (Windows specific)
            if os.name == 'nt':
                self.root.wm_iconbitmap(default=icon_path)
        except Exception as e:
            print(f"Could not set custom icon: {e}")
        
        # Set dark theme colors
        self.bg_color = "#2e2e2e"
        self.fg_color = "#ffffff"
        self.configure_dark_theme()
        
        # Basic window setup
        self.root.title("TaskMaster")
        self.root.geometry("800x600")
        self.root.configure(bg=self.bg_color)
        
        # Enable dark title bar on Windows immediately
        if os.name == 'nt':
            self.root.update_idletasks()
            set_window_dark_title_bar(self.root)
            # Schedule a quick focus cycle
            self.root.after(250, self.quick_focus_cycle)

        self.tasks = []
        self.filter_state = "all"  # all, active, completed
        self.sort_column = None
        self.sort_reverse = False
        
        self.create_widgets()
        
        # Configure the tree tag after tree is created
        self.tree.tag_configure("completed", foreground="#888888", font=("TkDefaultFont", 12, "overstrike"))
        
        self.load_tasks()
        self.center_window()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Keyboard Shortcuts
        self.root.bind("<Control-n>", lambda e: self.add_simple_task())
        self.root.bind("<Control-N>", lambda e: self.add_detailed_task())  # Ctrl+Shift+N
        self.root.bind("<Control-l>", lambda e: self.clear_completed_tasks())
        self.root.bind("<Control-L>", lambda e: self.clear_all_tasks())  # Ctrl+Shift+L
        self.root.bind("<Control-e>", lambda e: self.export_to_csv())
        self.root.bind("<Control-1>", lambda e: self.set_filter("all"))
        self.root.bind("<Control-2>", lambda e: self.set_filter("active"))
        self.root.bind("<Control-3>", lambda e: self.set_filter("completed"))
        self.root.bind("<Control-Return>", self.toggle_selected_completion)
        self.root.bind("<Delete>", self.delete_selected_task)
        self.root.bind("<Control-d>", self.delete_selected_task)

    def quick_focus_cycle(self):
        """Quickly cycle focus to force title bar update"""
        try:
            # Create a temporary transparent window
            temp = tk.Toplevel(self.root)
            temp.attributes('-alpha', 0.0)  # Make it invisible
            temp.geometry('1x1+0+0')  # Make it tiny
            temp.overrideredirect(True)  # Remove window decorations
            temp.focus_force()  # Force focus to temp window
            self.root.after(1, lambda: [self.root.focus_force(), temp.destroy()])  # Return focus and destroy temp
        except:
            pass  # If anything fails, just ignore it

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"800x600+{x}+{y}")

    def create_widgets(self):
        # Top frame for buttons and search bar
        top_frame = tk.Frame(self.root, bg=self.bg_color)
        top_frame.pack(side=tk.TOP, pady=5, fill=tk.X)

        # Add task buttons with tooltips
        add_simple = self.add_button(top_frame, "+ Add Simple Task", self.add_simple_task)
        add_simple.bind('<Enter>', lambda e: self.show_tooltip(add_simple, "Add a simple task (Ctrl+N)"))
        add_simple.bind('<Leave>', lambda e: self.hide_tooltip())

        add_detailed = self.add_button(top_frame, "+ Add Detailed Task", self.add_detailed_task)
        add_detailed.bind('<Enter>', lambda e: self.show_tooltip(add_detailed, "Add a detailed task (Ctrl+Shift+N)"))
        add_detailed.bind('<Leave>', lambda e: self.hide_tooltip())

        clear_completed = self.add_button(top_frame, "Clear Completed Tasks", self.clear_completed_tasks)
        clear_completed.bind('<Enter>', lambda e: self.show_tooltip(clear_completed, "Clear completed tasks (Ctrl+L)"))
        clear_completed.bind('<Leave>', lambda e: self.hide_tooltip())

        clear_all = self.add_button(top_frame, "Clear All Tasks", self.clear_all_tasks)
        clear_all.bind('<Enter>', lambda e: self.show_tooltip(clear_all, "Clear all tasks (Ctrl+Shift+L)"))
        clear_all.bind('<Leave>', lambda e: self.hide_tooltip())

        export_btn = self.add_button(top_frame, "Export to CSV", self.export_to_csv)
        export_btn.bind('<Enter>', lambda e: self.show_tooltip(export_btn, "Export tasks to CSV (Ctrl+E)"))
        export_btn.bind('<Leave>', lambda e: self.hide_tooltip())

        # Create tooltip label
        self.tooltip = tk.Label(self.root, bg="yellow", fg="black", relief="solid", borderwidth=1)
        self.tooltip.pack_forget()  # Hide initially

        # Filter buttons
        filter_frame = tk.Frame(self.root, bg=self.bg_color)
        filter_frame.pack(side=tk.TOP, pady=5)
        
        show_all = self.add_button(filter_frame, "Show All", lambda: self.set_filter("all"))
        show_all.bind('<Enter>', lambda e: self.show_tooltip(show_all, "Show all tasks (Ctrl+1)"))
        show_all.bind('<Leave>', lambda e: self.hide_tooltip())
        
        show_active = self.add_button(filter_frame, "Show Active", lambda: self.set_filter("active"))
        show_active.bind('<Enter>', lambda e: self.show_tooltip(show_active, "Show active tasks (Ctrl+2)"))
        show_active.bind('<Leave>', lambda e: self.hide_tooltip())
        
        show_completed = self.add_button(filter_frame, "Show Completed", lambda: self.set_filter("completed"))
        show_completed.bind('<Enter>', lambda e: self.show_tooltip(show_completed, "Show completed tasks (Ctrl+3)"))
        show_completed.bind('<Leave>', lambda e: self.hide_tooltip())

        # Search bar
        search_frame = tk.Frame(self.root, bg=self.bg_color)
        search_frame.pack(side=tk.TOP, pady=5, fill=tk.X)
        search_label = tk.Label(search_frame, text="Search:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12))
        search_label.pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=30, bg=self.bg_color, fg=self.fg_color, insertbackground=self.fg_color, font=("TkDefaultFont", 12))
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.update_task_list())

        self.task_frame = tk.Frame(self.root, bg=self.bg_color)
        self.task_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.task_frame, columns=("Completed", "Task", "Due Date", "Priority"), show='headings')
        self.tree.heading("Completed", text="✓", command=lambda: self.sort_tasks("completed"))
        self.tree.heading("Task", text="Task", command=lambda: self.sort_tasks("text"))
        self.tree.heading("Due Date", text="Due Date", command=lambda: self.sort_tasks("due"))
        self.tree.heading("Priority", text="Priority", command=lambda: self.sort_tasks("priority"))

        self.tree.column("Completed", width=40, anchor='center')
        self.tree.column("Task", width=300)
        self.tree.column("Due Date", width=150)
        self.tree.column("Priority", width=100)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<Double-1>", self.toggle_task_completion)
        self.tree.bind("<Button-3>", self.edit_task)

    def add_button(self, parent, text, command):
        btn = tk.Button(parent, text=text, command=command, bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12))
        btn.pack(side=tk.LEFT, padx=5)
        return btn

    def set_filter(self, state):
        self.filter_state = state
        self.update_task_list()

    def sort_tasks(self, key):
        self.sort_reverse = not self.sort_reverse if self.sort_column == key else False
        self.sort_column = key
        
        if key == "priority":
            # Custom priority sorting (Low -> Normal -> High)
            priority_order = {"Low": 1, "Normal": 2, "High": 3}
            self.tasks.sort(key=lambda x: priority_order.get(x.get(key, ""), 0), reverse=self.sort_reverse)
        else:
            # Default string sorting for other columns
            self.tasks.sort(key=lambda x: str(x.get(key, "")), reverse=self.sort_reverse)
        
        self.update_task_list()

    def add_simple_task(self):
        self.open_task_window(simple=True)

    def add_detailed_task(self):
        self.open_task_window(simple=False)

    def open_task_window(self, simple=False):
        window = tk.Toplevel(self.root)
        window.title("New Task")
        window.geometry("400x300")
        window.configure(bg=self.bg_color)
        self.center_child(window)
        window.grab_set()
        window.focus_force()

        tk.Label(window, text="Task:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12)).pack(pady=5)
        task_entry = tk.Entry(window, width=40, bg=self.bg_color, fg=self.fg_color, insertbackground=self.fg_color, font=("TkDefaultFont", 12))
        task_entry.pack(pady=5)
        task_entry.focus()

        if not simple:
            tk.Label(window, text="Due Date:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12)).pack(pady=5)
            date_frame = tk.Frame(window, bg=self.bg_color)
            date_frame.pack()

            def update_days(*args):
                try:
                    month = int(month_cb.get())
                    year = int(year_cb.get())
                    
                    # Get the last day of the selected month
                    if month in [4, 6, 9, 11]:
                        max_days = 30
                    elif month == 2:
                        # Check for leap year
                        if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                            max_days = 29
                        else:
                            max_days = 28
                    else:
                        max_days = 31
                    
                    # Update days combobox
                    current_day = day_cb.get()
                    day_cb['values'] = [f"{i:02d}" for i in range(1, max_days + 1)]
                    
                    # If current day is greater than max_days, set to last day of month
                    if int(current_day) > max_days:
                        day_cb.set(f"{max_days:02d}")
                except ValueError:
                    pass

            month_cb = ttk.Combobox(date_frame, values=[f"{i:02d}" for i in range(1, 13)], width=5, font=("TkDefaultFont", 12))
            day_cb = ttk.Combobox(date_frame, values=[f"{i:02d}" for i in range(1, 32)], width=5, font=("TkDefaultFont", 12))
            year_cb = ttk.Combobox(date_frame, values=[str(i) for i in range(2024, 2100)], width=7, font=("TkDefaultFont", 12))

            # Set today's date as default
            today = datetime.now()
            month_cb.set(f"{today.month:02d}")
            day_cb.set(f"{today.day:02d}")
            year_cb.set(str(today.year))

            month_cb.pack(side=tk.LEFT)
            day_cb.pack(side=tk.LEFT, padx=5)
            year_cb.pack(side=tk.LEFT)

            # Bind the update function to the month and year comboboxes
            month_cb.bind('<<ComboboxSelected>>', update_days)
            year_cb.bind('<<ComboboxSelected>>', update_days)

            tk.Label(window, text="Priority:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12)).pack(pady=5)
            priority_cb = ttk.Combobox(window, values=["Low", "Normal", "High"], state="readonly", font=("TkDefaultFont", 12))
            priority_cb.set("Normal")
            priority_cb.pack(pady=5)

        def submit(event=None):
            task_text = task_entry.get()
            if not task_text.strip():
                messagebox.showwarning("Input Error", "Task cannot be empty.")
                return

            due = f"{month_cb.get()}-{day_cb.get()}-{year_cb.get()}" if not simple else ""
            prio = priority_cb.get() if not simple else ""

            task = {"text": task_text, "due": due, "priority": prio, "completed": False}
            self.tasks.append(task)
            self.update_task_list()
            window.destroy()

        window.bind("<Return>", submit)  # Bind Enter key to submit
        submit_btn = tk.Button(window, text="Submit", command=lambda: submit(), bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12))
        submit_btn.pack(pady=10)

    def toggle_task_completion(self, event):
        selected = self.tree.focus()
        if not selected:
            return
        task_text = self.tree.item(selected)['values'][1]
        self.toggle_completion(task_text)

    def toggle_selected_completion(self, event=None):
        selected = self.tree.focus()
        if not selected:
            return
        task_text = self.tree.item(selected)['values'][1]
        self.toggle_completion(task_text)

    def toggle_completion(self, task_text):
        for task in self.tasks:
            if task['text'] == task_text:
                task['completed'] = not task.get('completed', False)
                break
        self.update_task_list()

    def delete_selected_task(self, event=None):
        selected = self.tree.focus()
        if not selected:
            return
        task_text = self.tree.item(selected)['values'][1]
        self.tasks = [task for task in self.tasks if task['text'] != task_text]
        self.update_task_list()

    def edit_task(self, event):
        selected = self.tree.focus()
        if not selected:
            return
        item = self.tree.item(selected)
        old_text = item['values'][1]
        
        # Find the task in our tasks list
        task = next((t for t in self.tasks if t['text'] == old_text), None)
        if not task:
            return
            
        is_detailed = bool(task.get('due') or task.get('priority'))
        
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Task")
        edit_window.geometry("400x300" if is_detailed else "400x150")
        edit_window.configure(bg=self.bg_color)
        self.center_child(edit_window)
        edit_window.grab_set()
        edit_window.focus_force()

        tk.Label(edit_window, text="Edit Task:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12)).pack(pady=5)
        new_entry = tk.Entry(edit_window, width=40, bg=self.bg_color, fg=self.fg_color, insertbackground=self.fg_color, font=("TkDefaultFont", 12))
        new_entry.insert(0, old_text)
        new_entry.pack(pady=5)
        new_entry.focus()

        date_frame = None
        month_cb = None
        day_cb = None
        year_cb = None
        priority_cb = None

        if is_detailed:
            tk.Label(edit_window, text="Due Date:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12)).pack(pady=5)
            date_frame = tk.Frame(edit_window, bg=self.bg_color)
            date_frame.pack()

            def update_days(*args):
                try:
                    month = int(month_cb.get())
                    year = int(year_cb.get())
                    
                    # Get the last day of the selected month
                    if month in [4, 6, 9, 11]:
                        max_days = 30
                    elif month == 2:
                        # Check for leap year
                        if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                            max_days = 29
                        else:
                            max_days = 28
                    else:
                        max_days = 31
                    
                    # Update days combobox
                    current_day = day_cb.get()
                    day_cb['values'] = [f"{i:02d}" for i in range(1, max_days + 1)]
                    
                    # If current day is greater than max_days, set to last day of month
                    if int(current_day) > max_days:
                        day_cb.set(f"{max_days:02d}")
                except ValueError:
                    pass

            month_cb = ttk.Combobox(date_frame, values=[f"{i:02d}" for i in range(1, 13)], width=5, font=("TkDefaultFont", 12))
            day_cb = ttk.Combobox(date_frame, values=[f"{i:02d}" for i in range(1, 32)], width=5, font=("TkDefaultFont", 12))
            year_cb = ttk.Combobox(date_frame, values=[str(i) for i in range(2024, 2100)], width=7, font=("TkDefaultFont", 12))

            # Set current date values
            if task.get('due'):
                try:
                    month, day, year = task['due'].split('-')
                    month_cb.set(month)
                    day_cb.set(day)
                    year_cb.set(year)
                except:
                    # If there's any error parsing the date, use today's date
                    today = datetime.now()
                    month_cb.set(f"{today.month:02d}")
                    day_cb.set(f"{today.day:02d}")
                    year_cb.set(str(today.year))
            else:
                # Use today's date as default
                today = datetime.now()
                month_cb.set(f"{today.month:02d}")
                day_cb.set(f"{today.day:02d}")
                year_cb.set(str(today.year))

            month_cb.pack(side=tk.LEFT)
            day_cb.pack(side=tk.LEFT, padx=5)
            year_cb.pack(side=tk.LEFT)

            # Bind the update function to the month and year comboboxes
            month_cb.bind('<<ComboboxSelected>>', update_days)
            year_cb.bind('<<ComboboxSelected>>', update_days)

            tk.Label(edit_window, text="Priority:", bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12)).pack(pady=5)
            priority_cb = ttk.Combobox(edit_window, values=["Low", "Normal", "High"], state="readonly", font=("TkDefaultFont", 12))
            priority_cb.set(task.get('priority', "Normal"))
            priority_cb.pack(pady=5)

        def save_edit(event=None):
            new_text = new_entry.get()
            if new_text:
                for t in self.tasks:
                    if t['text'] == old_text:
                        t['text'] = new_text
                        if is_detailed:
                            t['due'] = f"{month_cb.get()}-{day_cb.get()}-{year_cb.get()}"
                            t['priority'] = priority_cb.get()
                        break
                self.update_task_list()
                edit_window.destroy()

        edit_window.bind("<Return>", save_edit)
        save_btn = tk.Button(edit_window, text="Save", command=lambda: save_edit(), bg=self.bg_color, fg=self.fg_color, font=("TkDefaultFont", 12))
        save_btn.pack(pady=10)

    def update_task_list(self):
        search_text = self.search_var.get().lower().strip()

        self.tree.delete(*self.tree.get_children())
        for task in self.tasks:
            if self.filter_state == "active" and task.get('completed'):
                continue
            if self.filter_state == "completed" and not task.get('completed'):
                continue
            if search_text and search_text not in task['text'].lower():
                continue

            completed_mark = "✓" if task.get('completed') else ""
            tags = ("completed",) if task.get('completed') else ()
            self.tree.insert("", tk.END, values=(completed_mark, task['text'], task['due'], task['priority']), tags=tags)

    def clear_all_tasks(self, event=None):
        if not self.tasks:
            messagebox.showinfo("Info", "No tasks to clear.")
            return
            
        if messagebox.askyesno("Confirm", "Are you sure you want to clear ALL tasks? This cannot be undone."):
            self.tasks.clear()
            self.update_task_list()

    def clear_completed_tasks(self, event=None):
        if not any(task.get("completed", False) for task in self.tasks):
            messagebox.showinfo("Info", "No completed tasks to clear.")
            return
            
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all completed tasks?"):
            self.tasks = [task for task in self.tasks if not task.get("completed", False)]
            self.update_task_list()

    def configure_dark_theme(self):
        style = ttk.Style()
        style.theme_use("default")
        
        # Configure Treeview
        style.configure("Treeview",
                      background=self.bg_color,
                      foreground=self.fg_color,
                      fieldbackground=self.bg_color,
                      font=("TkDefaultFont", 12))  # Main task list font
        style.configure("Treeview.Heading",
                      background=self.bg_color,
                      foreground=self.fg_color,
                      font=("TkDefaultFont", 12))  # Column headers font
        
        # Configure buttons and other widgets font
        button_font = ("TkDefaultFont", 12)
        entry_font = ("TkDefaultFont", 12)
        
        # Configure Combobox
        style.configure("TCombobox",
                      selectbackground=self.bg_color,
                      selectforeground=self.fg_color,
                      fieldbackground=self.bg_color,
                      background=self.bg_color,
                      foreground=self.fg_color,
                      font=entry_font)
        
        style.map("TCombobox",
                 fieldbackground=[("readonly", self.bg_color)],
                 selectbackground=[("readonly", self.bg_color)],
                 selectforeground=[("readonly", self.fg_color)],
                 background=[("readonly", self.bg_color)],
                 foreground=[("readonly", self.fg_color)])
        
        # Set option add for combobox popdown
        self.root.option_add('*TCombobox*Listbox.background', self.bg_color)
        self.root.option_add('*TCombobox*Listbox.foreground', self.fg_color)
        self.root.option_add('*TCombobox*Listbox.selectBackground', "#4a4a4a")
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.fg_color)

    def on_close(self):
        self.save_tasks()
        self.root.destroy()

    def save_tasks(self):
        with open(TASKS_FILE, 'w') as f:
            json.dump(self.tasks, f)

    def load_tasks(self):
        if os.path.exists(TASKS_FILE):
            with open(TASKS_FILE, 'r') as f:
                self.tasks = json.load(f)
                self.update_task_list()

    def center_child(self, window):
        self.root.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()

        window.update_idletasks()
        win_width = window.winfo_width()
        win_height = window.winfo_height()

        x = root_x + (root_width - win_width) // 2
        y = root_y + (root_height - win_height) // 2
        window.geometry(f"+{x}+{y}")

    def show_tooltip(self, widget, text):
        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 20
        
        self.tooltip.configure(text=text)
        self.tooltip.place(x=x, y=y)

    def hide_tooltip(self):
        self.tooltip.place_forget()

    def export_to_csv(self, event=None):
        if not self.tasks:
            messagebox.showinfo("Info", "No tasks to export.")
            return
            
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Tasks to CSV"
            )
            
            if not file_path:  # User cancelled
                return
                
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                # Define headers with proper capitalization
                headers = ['Task', 'Completed', 'Due', 'Priority']
                writer = csv.DictWriter(csvfile, fieldnames=headers)
                writer.writeheader()
                
                # Write rows with field name mapping
                for task in self.tasks:
                    writer.writerow({
                        'Task': task['text'],
                        'Completed': task['completed'],
                        'Due': task['due'],
                        'Priority': task['priority']
                    })
                
            messagebox.showinfo("Success", f"Tasks exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export tasks: {str(e)}")

if __name__ == '__main__':
    root = tk.Tk()
    app = TaskMaster(root)
    root.mainloop()