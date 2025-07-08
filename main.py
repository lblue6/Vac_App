import tkinter as tk
from ttkbootstrap import Style, Button, Treeview, OptionMenu, Label, Entry, Frame
from ttkbootstrap.dialogs import Messagebox
from tkinter import ttk, filedialog
import datetime
import sqlite3
from PIL import Image, ImageTk
import os
from pdf2image import convert_from_path

# Database setup
db_file = os.path.join(os.path.dirname(__file__), "employees.db")
conn = sqlite3.connect(db_file)
cursor = conn.cursor()

expected_columns = ["id", "name", "employee_number", "status", "anniversary", "days_taken", "days_available", "document_path"]
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='employees'")
table_exists = cursor.fetchone()

if not table_exists:
    cursor.execute('''CREATE TABLE employees (
                        id INTEGER PRIMARY KEY,
                        employee_number INTEGER, 
                        name TEXT,
                        status TEXT, 
                        anniversary DATE, 
                        days_taken INTEGER, 
                        days_available INTEGER,
                        document_path TEXT)''')
    conn.commit()
    print("Table 'employees' created with correct schema.")
else:
    cursor.execute("PRAGMA table_info(employees)")
    existing_columns = [col[1] for col in cursor.fetchall()]
    for col in expected_columns:
        if col not in existing_columns:
            if col == "id":
                print("Error: Table exists but 'id' column is missing. Consider resetting the database.")
            else:
                col_type = "INTEGER" if col in ["employee_number", "days_taken", "days_available"] else "TEXT"
                cursor.execute(f"ALTER TABLE employees ADD COLUMN {col} {col_type}")
                print(f"Added missing column: {col}")
    conn.commit()

cursor.execute("SELECT id, anniversary FROM employees")
for row in cursor.fetchall():
    emp_id, anniversary = row
    if anniversary and '-' in anniversary:
        new_anniversary = anniversary.replace('-', '/')
        cursor.execute("UPDATE employees SET anniversary = ? WHERE id = ?", (new_anniversary, emp_id))
conn.commit()
print("Migrated existing dates to YYYY/MM/DD format.")

def calculate_vacation_days(anniversary):
    today = datetime.datetime.now()
    years_of_service = (today - anniversary).days / 365.25
    if years_of_service < 2:
        annual_days = 5
    elif 3 <= years_of_service < 5:
        annual_days = 10
    elif 6 <= years_of_service < 9:
        annual_days = 15
    else:
        annual_days = 20
    return int(years_of_service * annual_days)

class SplashScreen:
    def __init__(self, root):
        self.root = root
        self.root.geometry("840x540")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 840) // 2
        y = (screen_height - 540) // 2
        self.position = (x, y)
        self.root.geometry(f"840x540+{x}+{y}")

        try:
            img = Image.open("brand2.png")
            img = img.resize((840, 540), Image.Resampling.LANCZOS)
            self.photo = ImageTk.PhotoImage(img)
            self.label = tk.Label(self.root, image=self.photo)
            self.label.pack()
        except FileNotFoundError:
            self.label = tk.Label(self.root, text="Splash Screen\n(Image 'brand2.png' not found)",
                                  font=("Arial", 20), bg="black", fg="white")
            self.label.pack(expand=True)

        self.root.after(5000, self.close_splash)

    def close_splash(self):
        self.root.destroy()
        main_root = tk.Tk()
        app = VacationApp(main_root, position=self.position)
        app.root.mainloop()

class VacationApp:
    def __init__(self, root, position=None):
        self.delete_doc_btn = None
        self.root = root
        self.root.title("Employee Vacation Tracker")
        self.root.configure(background="lightgrey")

        if position:
            x, y = position
            self.root.geometry(f"830x440+{x}+{y}")
        else:
            self.root.geometry("830x440")

        self.root.resizable(False, False)

        self.style = Style(theme='sandstone')
        self.style.configure("Small.TMenubutton", width=7)
        self.style.configure("Treeview", background="whitesmoke", fieldbackground="white", foreground="black")
        self.style.configure("Treeview.Heading", font=("Arial", 12), background="darkblue", foreground="azure",
                            borderwidth=0, relief="flat")
        self.style.map("Treeview.Heading", background=[("disabled", "royal blue")], foreground=[("disabled", "white")])

        self.style.configure("Custom.TButton", background="firebrick", foreground="black", borderwidth=0,
                            font=("Arial", 12), padding=5)
        self.style.map("Custom.TButton", background=[("disabled", "lightcoral"), ("disabled", "#a9a9a9")],
                      foreground=[("disabled", "black"), ("disabled", "#d3d3d3")])
        self.style.configure("danger.Toolbutton", background=self.style.colors.danger, foreground="white", borderwidth=0)
        self.style.configure("primary.Toolbutton", background=self.style.colors.primary, foreground="white", borderwidth=0)

        input_frame = tk.Frame(root, background="white")
        input_frame.pack(side="top", anchor="n", pady=0)

        self.root.configure(bg="dimgrey")
        self.tree = ttk.Treeview(root, columns=("Name", "#", "Status", "Anniversary", "Days Taken", "Days Available", "Document"),
                                 show="headings")
        self.tree.heading("Name", text="Employee Name")
        self.tree.heading("#", text="#")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Anniversary", text="Anniversary Date")
        self.tree.heading("Days Taken", text="Days Taken")
        self.tree.heading("Days Available", text="Days Available")
        self.tree.heading("Document", text="Document")

        self.tree.column("Name", width=160, anchor="w", stretch=False)
        self.tree.column("#", width=40, anchor="center", stretch=False)
        self.tree.column("Status", width=100, anchor="center", stretch=False)
        self.tree.column("Anniversary", width=140, anchor="center", stretch=False)
        self.tree.column("Days Taken", width=140, anchor="center", stretch=False)
        self.tree.column("Days Available", width=140, anchor="center", stretch=False)
        self.tree.column("Document", width=120, anchor="center", stretch=False)

        self.tree.pack( pady=0,fill="x" )

        # Disable interactive column resizing
        def block_resize(event):
            if self.tree.identify_region(event.x, event.y) == "separator":
                return "break"  # Prevent column resizing
            return None

        self.tree.bind("<Button-1>", block_resize)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", lambda e: self.tree.selection_set(self.tree.identify_row(e.y)))

        input_fields_frame = tk.Frame(root, background="gainsboro")
        input_fields_frame.pack(pady=5, anchor="n", fill="x")

        self.labels = [
            tk.Label(input_fields_frame, text="Employee Number:", bg="gainsboro", fg="black", font=("Arial", 12)),
            tk.Label(input_fields_frame, text="Employee Name:", bg="gainsboro", fg="black", font=("Arial", 12)),
            tk.Label(input_fields_frame, text="Status:", bg="gainsboro", fg="black", font=("Arial", 12)),
            tk.Label(input_fields_frame, text="Anniversary Date:", bg="gainsboro", fg="black", font=("Arial", 12)),
            tk.Label(input_fields_frame, text="Day Grab:", bg="gainsboro", fg="black", font=("Arial", 12))
        ]

        vcmd_num = (self.root.register(self.validate_employee_number), '%P')
        self.employee_number_entry = tk.Entry(input_fields_frame, bg="whitesmoke", borderwidth=0,
                                              highlightthickness=0, width=6, fg="black", validate="key",
                                              validatecommand=vcmd_num, insertbackground="black")

        vcmd_name = (self.root.register(self.validate_employee_name), '%P')
        self.name_entry = tk.Entry(input_fields_frame, bg="whitesmoke", borderwidth=0,
                                   highlightthickness=0, width=15, fg="black", validate="key",
                                   validatecommand=vcmd_name, insertbackground="black")

        self.status_var = tk.StringVar(value="Company")
        self.status_menu = ttk.OptionMenu(input_fields_frame, self.status_var, "Company", "Temp", "Company",
                                          style="Small.TMenubutton", command=self.update_status)

        vcmd_date = (self.root.register(self.validate_anniversary_date), '%P')
        self.anniversary_entry = tk.Entry(input_fields_frame, bg="whitesmoke", borderwidth=0,
                                          highlightthickness=0, width=10, fg="gray", justify="center",
                                          validate="key", validatecommand=vcmd_date, insertbackground="black")
        self.anniversary_entry.insert(0, "YYYY/MM/DD")
        self.anniversary_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.anniversary_entry.bind("<FocusOut>", self.on_entry_focus_out)

        # Input fields (labels, entries, status menu)
        self.labels[0].grid(row=0, column=4, padx=(0, 5), sticky="w")
        self.employee_number_entry.grid(row=0, column=5, padx=(0, 10), pady=2)
        self.labels[1].grid(row=0, column=2, padx=(0, 5), sticky="w")
        self.name_entry.grid(row=0, column=3, padx=(0, 10), pady=2)
        self.labels[2].grid(row=0, column=0, padx=(0, 2), sticky="e")
        self.status_menu.grid(row=0, column=1, padx=(0, 10), pady=2)
        self.labels[3].grid(row=0, column=6, padx=(0, 5), sticky="w")
        self.anniversary_entry.grid(row=0, column=7, padx=(0, 10), pady=2)

        # Horizontal separator
        separator = ttk.Separator(input_fields_frame, orient="horizontal", style="primary.TSeparator")
        separator.grid(row=1, column=0, columnspan=10, sticky="nsew", pady=(10, 5), padx=(0))

        separator = ttk.Separator(input_fields_frame, orient="horizontal", style="primary.TSeparator")
        separator.grid(row=4, column=0, columnspan=10, sticky="nsew", pady=(10), padx=(0))

        # Buttons and slider
        self.add_employee_btn = Button(input_fields_frame, text="Add Employee", command=self.add_employee,
                                      style="danger.Outline.Toolbutton")
        self.add_employee_btn.grid(row=2, column=7, padx=5, pady=6)
        self.upload_document_btn = ttk.Button(input_fields_frame, text="Upload Doc", command=self.upload_document,
                                              style="danger.Outline.Toolbutton")
        self.upload_document_btn.grid(row=2, column=1, padx=(10, 5), pady=6)
        self.preview_btn = ttk.Button(input_fields_frame, text="Preview", command=self.show_preview,
                                      style="danger.Toolbutton", state="disabled")
        self.preview_btn.grid(row=3, column=1, padx=0, pady=0)

        self.data_btn = ttk.Button(input_fields_frame, text="Save", command=self.print_database,
                                    style="primary.Toolbutton")
        self.data_btn.grid(row=5, column=1, padx=5, pady=28)

        self.delete_employee_btn = Button(input_fields_frame, text="Delete", command=self.delete_employee,
                                          style="danger.Toolbutton", state="disabled")
        self.delete_employee_btn.grid(row=3, column=7, padx=5, pady=6)

        self.refresh_days_btn = ttk.Button(input_fields_frame, text="Refresh", command=self.refresh_days,
                                           style="primary.Toolbutton")
        self.refresh_days_btn.grid(row=5, column=7, padx=5, pady=28)

        self.labels[4].grid(row=2, column=3, columnspan=1, pady=5, padx=10, sticky="e")
        self.days_slider = ttk.Scale(input_fields_frame, from_=0, to=0, orient="horizontal",
                                     command=self.on_slider_change, state="disabled")
        self.days_slider.grid(row=3, column=3, columnspan=2, padx=(0, 5), pady=5)

        self.version_label = tk.Label(root, text="Version 1.2", font=("Arial", 12), fg="black")
        self.version_label.place(relx=0.48, rely=0.96, anchor="s")

        self.selected_employee_id = None
        self.preview_window = None
        self.doc_selector = None
        self.preview_label = None
        self.zoom_level = 1.0
        self.load_data()

        self.update_employee_number_state("Company")
        self.root.update()
        self.root.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}")

    def show_centered_messagebox(self, title, message, msg_type="show_error"):
        """Show a centered Messagebox relative to the main application window."""
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # Calculate the center position
        dialog_width = 300  # Approximate width of the dialog
        dialog_height = 150  # Approximate height of the dialog
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        # Call the appropriate Messagebox method based on msg_type
        if msg_type == "show_error":
            return Messagebox.show_error(title=title, message=message, parent=self.root, position=(x, y))
        elif msg_type == "show_info":
            return Messagebox.show_info(title=title, message=message, parent=self.root, position=(x, y))
        elif msg_type == "yesno":
            return Messagebox.yesno(title=title, message=message, parent=self.root, position=(x, y))
        return None

    def update_treeview_style(self, status):
        """Update the Treeview selection background based on status."""
        if status == "Temp":
            self.style.map("Treeview", background=[("selected", "#a9a9a9")])
        else:
            self.style.map("Treeview", background=[("selected", "#4a72a6")])

    @staticmethod
    def validate_employee_number(new_value):
        if new_value == "":
            return True
        if not new_value.isdigit():
            return False
        if len(new_value) > 3:
            return False
        return True

    @staticmethod
    def validate_employee_name(new_value):
        if new_value == "":
            return True
        return all(char.isalpha() or char.isspace() for char in new_value)

    @staticmethod
    def validate_anniversary_date(new_value):
        if new_value == "":
            return True
        if new_value == "YYYY/MM/DD":
            return True
        if len(new_value) > 10:
            return False
        return all(char.isdigit() or char == "/" for char in new_value)

    def on_entry_focus_in(self, _):
        if self.anniversary_entry.get() == "YYYY/MM/DD":
            self.anniversary_entry.delete(0, tk.END)
            self.anniversary_entry.config(fg="black")

    def on_entry_focus_out(self, _):
        if not self.anniversary_entry.get():
            self.anniversary_entry.insert(0, "YYYY/MM/DD")
            self.anniversary_entry.config(fg="gray")

    def update_employee_number_state(self, status):
        if status == "Temp":
            self.employee_number_entry.delete(0, tk.END)
            self.employee_number_entry.config(state="disabled")
        else:
            self.employee_number_entry.config(state="normal")
        self.update_treeview_style(status)

    def update_status(self, *_):
        if self.selected_employee_id:
            new_status = self.status_var.get()
            try:
                cursor.execute("UPDATE employees SET status = ? WHERE id = ?", (new_status, self.selected_employee_id))
                conn.commit()
                current_values = self.tree.item(self.selected_employee_id, "values")
                self.tree.item(self.selected_employee_id,
                               values=(current_values[0], current_values[1], new_status, *current_values[3:]))
                self.update_employee_number_state(new_status)
            except sqlite3.Error as e:
                self.show_centered_messagebox(title="Database Error", message=f"Error updating status: {e}", msg_type="show_error")

    def add_employee(self):
        status = self.status_var.get()
        employee_number_str = self.employee_number_entry.get().strip()

        if status == "Temp":
            employee_number = None
        else:
            if not employee_number_str:
                self.show_centered_messagebox(title="Error", message="Employee Number cannot be empty for Company status!", msg_type="show_error")
                return
            if not employee_number_str.isdigit() or len(employee_number_str) > 3:
                self.show_centered_messagebox(title="Error", message="Employee Number must be a number with maximum 3 digits!", msg_type="show_error")
                return
            employee_number = int(employee_number_str)
            # Check for duplicate employee number
            cursor.execute("SELECT id FROM employees WHERE employee_number = ? AND status = 'Company'", (employee_number,))
            if cursor.fetchone():
                self.show_centered_messagebox(title="Error", message="Employee Number already exists!", msg_type="show_error")
                return

        name = self.name_entry.get().strip()
        anniversary_str = self.anniversary_entry.get().strip()
        if not name or not anniversary_str or anniversary_str == "YYYY/MM/DD":
            self.show_centered_messagebox(title="Error", message="Name and Anniversary Date cannot be empty!", msg_type="show_error")
            return

        try:
            anniversary = datetime.datetime.strptime(anniversary_str, "%Y/%m/%d")
            total_days = calculate_vacation_days(anniversary)
            days_taken = 0
            days_available = total_days - days_taken

            cursor.execute("SELECT MAX(id) FROM employees")
            max_id = cursor.fetchone()[0]
            employee_id = 1 if max_id is None else max_id + 1

            cursor.execute(
                "INSERT INTO employees (id, name, employee_number, status, anniversary, days_taken, days_available, document_path) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (employee_id, name, employee_number, status, anniversary_str, days_taken, days_available, None)
            )
            conn.commit()

            employee_number_display = "" if status == "Temp" else employee_number
            self.tree.insert("", "end", iid=employee_id,
                            values=(name, employee_number_display, status, anniversary_str, days_taken, days_available, ""))
            self.clear_entries()
        except ValueError:
            self.show_centered_messagebox(title="Error", message="Invalid date format! Use YYYY/MM/DD.", msg_type="show_error")
        except sqlite3.Error as e:
            self.show_centered_messagebox(title="Database Error", message=f"Error adding employee: {e}", msg_type="show_error")

    def upload_document(self):
        if not self.selected_employee_id:
            self.show_centered_messagebox(title="Error", message="Please select an employee first!", msg_type="show_error")
            return

        file_path = filedialog.askopenfilename(filetypes=[("PDF/JPEG Files", "*.pdf *.jpg *.jpeg")])
        if not file_path:
            return

        doc_name = os.path.basename(file_path)
        new_doc_entry = f"{doc_name}|{file_path}"

        cursor.execute("SELECT document_path FROM employees WHERE id = ?", (self.selected_employee_id,))
        current_docs = cursor.fetchone()[0]

        if current_docs:
            updated_docs = f"{current_docs};{new_doc_entry}"
        else:
            updated_docs = new_doc_entry

        cursor.execute("UPDATE employees SET document_path = ? WHERE id = ?", (updated_docs, self.selected_employee_id))
        conn.commit()

        current_values = self.tree.item(self.selected_employee_id, "values")
        self.tree.item(self.selected_employee_id, values=(*current_values[:-1], doc_name))

    def on_tree_select(self, _):
        selected_item = self.tree.selection()
        if selected_item:
            self.selected_employee_id = int(selected_item[0])
            self.days_slider.config(state="normal")
            self.delete_employee_btn.config(state="normal")
            self.preview_btn.config(state="normal")
            cursor.execute("SELECT status, days_taken, days_available, anniversary FROM employees WHERE id = ?",
                           (self.selected_employee_id,))
            current_status, days_taken, days_available, anniversary_str = cursor.fetchone()
            self.status_var.set(current_status)
            anniversary = datetime.datetime.strptime(anniversary_str, "%Y/%m/%d")
            total_days = calculate_vacation_days(anniversary)
            self.days_slider.config(from_=0, to=total_days)
            self.days_slider.set(days_taken)
            self.update_employee_number_state(current_status)
        else:
            self.selected_employee_id = None
            self.days_slider.config(state="disabled", from_=0, to=0)
            self.delete_employee_btn.config(state="disabled")
            self.preview_btn.config(state="disabled")
            self.close_preview()
            self.status_var.set("Company")
            self.update_treeview_style("Company")

    def on_double_click(self, event):
        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column or not self.selected_employee_id:
            return

        col_index = int(column[1:]) - 1
        current_values = self.tree.item(item, "values")
        value_to_edit = current_values[col_index]

        if col_index == 1 and not value_to_edit:
            return

        x, y, width, height = self.tree.bbox(item, column)

        if col_index == 2:
            status_var = tk.StringVar(value=value_to_edit)
            menu = ttk.OptionMenu(self.root, status_var, value_to_edit, "Company", "Temp",
                                  command=lambda new_val: self.save_status_edit(item, new_val))
            menu.place(x=x + self.tree.winfo_x(), y=y + self.tree.winfo_y(), width=width, height=height)
            menu.focus_set()
            menu.bind("<FocusOut>", lambda _: menu.destroy())
        else:
            entry = tk.Entry(self.root, borderwidth=1, highlightthickness=1)
            entry.place(x=x + self.tree.winfo_x(), y=y + self.tree.winfo_y(), width=width, height=height)
            entry.insert(0, value_to_edit)
            entry.focus_set()

            def save_edit(_):
                new_value = entry.get().strip()
                if new_value:
                    try:
                        if col_index == 0:
                            cursor.execute("UPDATE employees SET name = ? WHERE id = ?", (new_value, item))
                            self.tree.item(item, values=(new_value, *current_values[1:]))
                        elif col_index == 1:
                            if not new_value.isdigit() or len(new_value) > 3:
                                self.show_centered_messagebox(title="Error", message="Employee Number must be a number with max 3 digits!", msg_type="show_error")
                                return
                            new_employee_number = int(new_value)
                            # Check for duplicate employee number
                            cursor.execute("SELECT id FROM employees WHERE employee_number = ? AND id != ? AND status = 'Company'", (new_employee_number, item))
                            if cursor.fetchone():
                                self.show_centered_messagebox(title="Error", message="Employee Number already exists!", msg_type="show_error")
                                return
                            cursor.execute("UPDATE employees SET employee_number = ? WHERE id = ?", (new_employee_number, item))
                            self.tree.item(item, values=(current_values[0], new_value, *current_values[2:]))
                        elif col_index == 3:
                            datetime.datetime.strptime(new_value, "%Y/%m/%d")
                            cursor.execute("UPDATE employees SET anniversary = ? WHERE id = ?", (new_value, item))
                            self.tree.item(item, values=(current_values[0], current_values[1], current_values[2], new_value, *current_values[4:]))
                        elif col_index == 6:
                            cursor.execute("SELECT document_path FROM employees WHERE id = ?", (item,))
                            doc_path = cursor.fetchone()[0]
                            doc_list = doc_path.split(";") if doc_path else []
                            if doc_list:
                                old_doc = doc_list[-1]
                                file_path = old_doc.split("|")[1]
                                doc_list[-1] = f"{new_value}|{file_path}"
                                updated_docs = ";".join(doc_list)
                                cursor.execute("UPDATE employees SET document_path = ? WHERE id = ?", (updated_docs, item))
                                self.tree.item(item, values=(*current_values[:-1], new_value))

                        conn.commit()
                    except ValueError as e:
                        self.show_centered_messagebox(title="Error", message=f"Invalid input: {str(e)}", msg_type="show_error")
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)

    def save_status_edit(self, item, new_status):
        try:
            cursor.execute("UPDATE employees SET status = ? WHERE id = ?", (new_status, item))
            conn.commit()
            current_values = self.tree.item(item, "values")
            self.tree.item(item, values=(current_values[0], current_values[1], new_status, *current_values[3:]))
            self.status_var.set(new_status)
            self.update_employee_number_state(new_status)
        except sqlite3.Error as e:
            self.show_centered_messagebox(title="Database Error", message=f"Error updating status: {e}", msg_type="show_error")

    def show_preview(self):
        self.close_preview()
        if not self.selected_employee_id:
            return

        cursor.execute("SELECT document_path FROM employees WHERE id = ?", (self.selected_employee_id,))
        doc_path = cursor.fetchone()[0]
        if not doc_path:
            return

        doc_list = doc_path.split(";")
        if not doc_list:
            return

        self.preview_window = tk.Toplevel(self.root)
        self.preview_window.title("Document Preview")
        self.preview_window.transient(self.root)
        self.preview_window.protocol("WM_DELETE_WINDOW", self.close_preview)

        doc_names = [doc.split("|")[0] for doc in doc_list]
        self.doc_selector = ttk.Combobox(self.preview_window, values=doc_names, state="readonly")
        self.doc_selector.pack(pady=5)
        self.doc_selector.current(len(doc_names) - 1)
        self.doc_selector.bind("<<ComboboxSelected>>", self.update_preview)

        zoom_frame = tk.Frame(self.preview_window)
        zoom_frame.pack(pady=5)

        self.delete_doc_btn = ttk.Button(zoom_frame, text="Delete", command=self.delete_current_doc, style="danger.Toolbutton")
        self.delete_doc_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="+", width=2, command=self.zoom_in, style="dark.Outline.Toolbutton").pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="−", width=2, command=self.zoom_out, style="dark.Outline.Toolbutton").pack(side=tk.LEFT, padx=2)

        self.preview_label = tk.Label(self.preview_window)
        self.preview_label.pack()
        self.zoom_level = 1.0
        self.update_preview(None)

    def delete_current_doc(self):
        if not self.selected_employee_id or not self.doc_selector:
            return

        if self.show_centered_messagebox(title="Confirm Delete", message="Are you sure you want to delete this document?", msg_type="yesno") != "Yes":
            return

        selected_idx = self.doc_selector.current()
        if selected_idx < 0:
            self.show_centered_messagebox(title="Error", message="No document selected!", msg_type="show_error")
            return

        try:
            cursor.execute("SELECT document_path FROM employees WHERE id = ?", (self.selected_employee_id,))
            doc_path = cursor.fetchone()[0]
            if not doc_path:
                return

            doc_list = doc_path.split(";")
            if selected_idx >= len(doc_list):
                return

            del doc_list[selected_idx]
            updated_docs = ";".join(doc_list) if doc_list else None

            cursor.execute("UPDATE employees SET document_path = ? WHERE id = ?", (updated_docs, self.selected_employee_id))
            conn.commit()

            # Update Treeview
            current_values = self.tree.item(self.selected_employee_id, "values")
            if updated_docs:
                new_doc_name = updated_docs.split(";")[-1].split("|")[0]
            else:
                new_doc_name = ""
            self.tree.item(self.selected_employee_id, values=(*current_values[:-1], new_doc_name))

            if not doc_list:
                self.close_preview()
                self.show_centered_messagebox(title="Success", message="Document deleted successfully!", msg_type="show_info")
                return

            doc_names = [doc.split("|")[0] for doc in doc_list]
            self.doc_selector['values'] = doc_names
            self.doc_selector.current(0)
            self.update_preview(None)
            self.show_centered_messagebox(title="Success", message="Document deleted successfully!", msg_type="show_info")
        except sqlite3.Error as e:
            self.show_centered_messagebox(title="Database Error", message=f"Error deleting document: {e}", msg_type="show_error")

    def zoom_in(self):
        self.zoom_level = min(self.zoom_level + 0.2, 3.0)
        self.update_preview(None)

    def zoom_out(self):
        self.zoom_level = max(self.zoom_level - 0.2, 0.2)
        self.update_preview(None)

    def update_preview(self, _):
        if not self.preview_window or not self.selected_employee_id:
            return

        cursor.execute("SELECT document_path FROM employees WHERE id = ?", (self.selected_employee_id,))
        doc_path = cursor.fetchone()[0]
        doc_list = doc_path.split(";")
        selected_idx = self.doc_selector.current()
        doc_name, file_path = doc_list[selected_idx].split("|", 1)

        try:
            if file_path.lower().endswith(('.jpg', '.jpeg')):
                img = Image.open(file_path)
            elif file_path.lower().endswith('.pdf'):
                images = convert_from_path(file_path, first_page=1, last_page=1)
                if images:
                    img = images[0]
                else:
                    self.preview_label.config(image=None, text="PDF is empty")
                    return
            else:
                self.preview_label.config(image=None, text="Unsupported file format")
                return

            base_size = 600
            new_size = int(base_size * self.zoom_level)
            img.thumbnail((new_size, new_size))
            photo = ImageTk.PhotoImage(img)
            self.preview_label.config(image=photo, text="")
            self.preview_window.image = photo
        except Exception as e:
            self.preview_label.config(image=None, text=f"Error loading preview: {str(e)}")

    def close_preview(self):
        if self.preview_window:
            self.preview_window.destroy()
            self.preview_window = None
            self.doc_selector = None
            self.preview_label = None

    def on_slider_change(self, value):
        if not self.selected_employee_id:
            return

        try:
            new_days_taken = int(float(value))
            cursor.execute("SELECT anniversary FROM employees WHERE id = ?", (self.selected_employee_id,))
            anniversary_str = cursor.fetchone()[0]
            anniversary = datetime.datetime.strptime(anniversary_str, "%Y/%m/%d")
            total_days = calculate_vacation_days(anniversary)
            if new_days_taken > total_days:
                new_days_taken = total_days
            new_days_available = total_days - new_days_taken
            cursor.execute("UPDATE employees SET days_taken = ?, days_available = ? WHERE id = ?",
                           (new_days_taken, new_days_available, self.selected_employee_id))
            conn.commit()
            current_values = self.tree.item(self.selected_employee_id, "values")
            self.tree.item(self.selected_employee_id, values=(
                current_values[0], current_values[1], current_values[2], current_values[3],
                new_days_taken, new_days_available, current_values[6]))
        except sqlite3.Error as e:
            self.show_centered_messagebox(title="Database Error", message=f"Error adjusting days: {e}", msg_type="show_error")

    def delete_employee(self):
        if not self.selected_employee_id:
            self.show_centered_messagebox(title="Error", message="Please select an employee to delete!", msg_type="show_error")
            return

        if self.show_centered_messagebox(title="Confirm Delete", message="Are you sure you want to delete this employee?", msg_type="yesno") == "Yes":
            try:
                cursor.execute("DELETE FROM employees WHERE id = ?", (self.selected_employee_id,))
                conn.commit()
                self.tree.delete(self.selected_employee_id)
                self.selected_employee_id = None
                self.days_slider.config(state="disabled", from_=0, to=0)
                self.delete_employee_btn.config(state="disabled")
                self.preview_btn.config(state="disabled")
                self.close_preview()
                self.status_var.set("Company")
                self.update_treeview_style("Company")
                self.show_centered_messagebox(title="Success", message="Employee deleted successfully!", msg_type="show_info")
            except sqlite3.Error as e:
                self.show_centered_messagebox(title="Database Error", message=f"Error deleting employee: {e}", msg_type="show_error")

    def load_data(self, sort_by_last_name=False):
        try:
            self.tree.delete(*self.tree.get_children())
            query = "SELECT id, employee_number, name, status, anniversary, days_taken, days_available, document_path FROM employees"
            if sort_by_last_name:
                rows = cursor.execute(query).fetchall()
                rows.sort(key=lambda x: x[2].split()[-1] if x[2].split() else x[2])
            else:
                rows = cursor.execute(query + " ORDER BY id").fetchall()

            for row in rows:
                employee_id, employee_number, name, status, anniversary, days_taken, days_available, doc_path = row
                anniversary_date = datetime.datetime.strptime(anniversary, "%Y/%m/%d")
                total_days = calculate_vacation_days(anniversary_date)
                updated_available = total_days - days_taken
                if updated_available != days_available:
                    cursor.execute("UPDATE employees SET days_available = ? WHERE id = ?", (updated_available, employee_id))
                    conn.commit()
                doc_name = ""
                if doc_path and ";" in doc_path:
                    doc_name = doc_path.split(";")[-1].split("|")[0]
                elif doc_path:
                    doc_name = doc_path.split("|")[0]
                employee_number_str = "" if employee_number is None else str(employee_number)
                self.tree.insert("", "end", iid=employee_id,
                                values=(name, employee_number_str, status, anniversary, days_taken, updated_available, doc_name))
        except sqlite3.Error as e:
            self.show_centered_messagebox(title="Database Error", message=f"Error loading data: {e}", msg_type="show_error")

    def refresh_days(self):
        self.load_data(sort_by_last_name=True)

    def print_database(self):
        try:
            # Get selected employees from Treeview
            selected_items = self.tree.selection()
            if not selected_items:
                self.show_centered_messagebox(title="Error", message="Please select at least one employee!",
                                              msg_type="show_error")
                return

            # Prepare the report header
            output = "Selected Employees Report\n"
            output += "=" * 120 + "\n"
            header = f"{'Name':<20} {'#':^10} {'Status':^15} {'Anniversary':^20} {'Days Taken':^15} {'Days Available':^15} {'Document':^25}"
            output += header + "\n"
            output += "=" * 120 + "\n"

            # Fetch data for each selected employee
            for employee_id in selected_items:
                cursor.execute(
                    "SELECT employee_number, name, status, anniversary, days_taken, days_available, document_path FROM employees WHERE id = ?",
                    (employee_id,)
                )
                row = cursor.fetchone()
                if row:
                    employee_number, name, status, anniversary, days_taken, days_available, doc_path = row
                    employee_number_str = "" if employee_number is None else str(employee_number)
                    doc_name = ""
                    if doc_path and ";" in doc_path:
                        doc_name = doc_path.split(";")[-1].split("|")[0]
                    elif doc_path:
                        doc_name = doc_path.split("|")[0]

                    line = f"{name:<20} {employee_number_str:^10} {status:^15} {anniversary:^20} {days_taken:^15} {days_available:^15} {doc_name:^25}"
                    output += line + "\n"

            output += "=" * 120 + "\n"

            # Display the report in a new window
            print_window = tk.Toplevel(self.root)
            print_window.title("Selected Employees Report")
            print_window.geometry("740x400")

            text_widget = tk.Text(print_window, wrap="none", font=("Courier", 10))
            text_widget.insert(tk.END, output)
            text_widget.config(state="disabled")
            text_widget.pack(expand=True, fill="both", padx=5, pady=5)

            def save_to_file():
                file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                         filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
                if file_path:
                    with open(file_path, "w") as f:
                        f.write(output)
                    self.show_centered_messagebox(title="Success",
                                                  message="Selected employees report saved successfully!",
                                                  msg_type="show_info")
                    print_window.destroy()

            save_btn = ttk.Button(print_window, text="Save to File", command=save_to_file, style="danger.Toolbutton")
            save_btn.pack(pady=10)

        except sqlite3.Error as e:
            self.show_centered_messagebox(title="Database Error", message=f"Error generating report: {e}",
                                          msg_type="show_error")

    def clear_entries(self):
        self.employee_number_entry.config(state="normal")
        self.employee_number_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        self.anniversary_entry.delete(0, tk.END)
        self.anniversary_entry.insert(0, "YYYY/MM/DD")
        self.anniversary_entry.config(fg="gray")
        self.status_var.set("Company")
        self.update_employee_number_state("Company")

if __name__ == "__main__":
    splash_root = tk.Tk()
    splash = SplashScreen(splash_root)
    splash_root.mainloop()

conn.close()