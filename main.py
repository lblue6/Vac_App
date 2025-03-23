import tkinter as tk
from ttkbootstrap import Style, Button, Treeview, OptionMenu, Label, Entry, Frame
from tkinter import ttk, messagebox, filedialog
import datetime
import sqlite3
from PIL import Image, ImageTk
import os
from pdf2image import convert_from_path

# Database setup (unchanged)
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
        self.root.geometry("840x440")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 840) // 2
        y = (screen_height - 440) // 2
        self.root.geometry(f"840x440+{x}+{y}")

        try:
            img = Image.open("brand.png")
            img = img.resize((840, 440), Image.Resampling.LANCZOS)
            self.photo = ImageTk.PhotoImage(img)
            self.label = tk.Label(self.root, image=self.photo)
            self.label.pack()
        except FileNotFoundError:
            self.label = tk.Label(self.root, text="Splash Screen\n(Image 'brand.png' not found)",
                                  font=("Arial", 20), bg="black", fg="white")
            self.label.pack(expand=True)

        self.root.after(5000, self.close_splash)

    def close_splash(self):
        self.root.destroy()
        main_root = tk.Tk()
        app = VacationApp(main_root)
        app.root.mainloop()

class VacationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Employee Vacation Tracker")
        self.root.configure(background="lightgrey")
        self.root.geometry("840x440")
        self.root.resizable(False, False)

        style = Style(theme='sandstone')
        style.configure("Small.TMenubutton", width=7)
        style.configure("Treeview", background="whitesmoke", fieldbackground="white", foreground="black")
        style.configure("Treeview.Heading", font=("Arial", 12), background="darkblue", foreground="azure",
                        borderwidth=0, relief="flat")
        style.map("Treeview.Heading", background=[("disabled", "royal blue")], foreground=[("disabled", "white")])
        style.map("Treeview", background=[("selected", "#4a72a6")])

        style.configure("Custom.TButton", background="firebrick", foreground="black", borderwidth=0,
                        font=("Arial", 12), padding=5)
        style.map("Custom.TButton", background=[("disabled", "lightcoral"), ("disabled", "#a9a9a9")],
                  foreground=[("disabled", "black"), ("disabled", "#d3d3d3")])
        style.configure("danger.Toolbutton", background=style.colors.danger, foreground="white", borderwidth=0)
        style.configure("primary.Toolbutton", background=style.colors.primary, foreground="white", borderwidth=0)

        input_frame = tk.Frame(root, background="white")
        input_frame.pack(side="top", anchor="n", pady=0)

        self.root.configure(bg="dimgrey")
        self.tree = ttk.Treeview(root, columns=("ID_Name", "#", "Status", "Anniversary", "Days Taken", "Days Available", "Document"),
                                 show="headings")
        self.tree.heading("ID_Name", text="Employee Name")
        self.tree.heading("#", text="#")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Anniversary", text="Anniversary Date")
        self.tree.heading("Days Taken", text="Days Taken")
        self.tree.heading("Days Available", text="Days Available")
        self.tree.heading("Document", text="Document")

        self.tree.column("ID_Name", width=160, anchor="w", stretch=False)
        self.tree.column("#", width=50, anchor="center", stretch=False)
        self.tree.column("Status", width=100, anchor="center", stretch=False)
        self.tree.column("Anniversary", width=140, anchor="center", stretch=False)
        self.tree.column("Days Taken", width=140, anchor="center", stretch=False)
        self.tree.column("Days Available", width=140, anchor="center", stretch=False)
        self.tree.column("Document", width=120, anchor="center", stretch=False)

        self.tree.pack(pady=0, fill="x")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.on_double_click)

        input_fields_frame = tk.Frame(root, background="gainsboro")
        input_fields_frame.pack(pady=5)

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

        self.labels[0].grid(row=0, column=0, padx=(0, 2), sticky="e")
        self.employee_number_entry.grid(row=0, column=1, padx=(0, 10), pady=2)
        self.labels[1].grid(row=0, column=2, padx=(0, 5), sticky="w")
        self.name_entry.grid(row=0, column=3, padx=(0, 10), pady=2)
        self.labels[2].grid(row=0, column=4, padx=(0, 5), sticky="w")
        self.status_menu.grid(row=0, column=5, padx=(0, 10), pady=2)
        self.labels[3].grid(row=0, column=6, padx=(0, 5), sticky="w")
        self.anniversary_entry.grid(row=0, column=7, padx=(0, 10), pady=2)

        self.add_employee_btn = Button(input_fields_frame, text="Add Employee", command=self.add_employee,
                                       style="danger.Outline.Toolbutton")
        self.add_employee_btn.grid(row=1, column=7, padx=5, pady=10)

        self.delete_employee_btn = Button(input_fields_frame, text="Delete",
                                          command=self.delete_employee,
                                          style="danger.Toolbutton",
                                          state="disabled")
        self.delete_employee_btn.grid(row=2, column=7, padx=5, pady=10)

        self.labels[4].grid(row=2, column=3, columnspan=1, pady=5, padx=10, sticky="e")
        self.minus_btn = ttk.Button(input_fields_frame, text="−", width=2, command=self.decrease_days, state="disabled",
                                    style="dark.Outline.Toolbutton")
        self.minus_btn.grid(row=3, column=3, columnspan=1, padx=(0, 5), pady=5)
        self.plus_btn = ttk.Button(input_fields_frame, text="+", width=2, command=self.increase_days, state="disabled",
                                   style="dark.Outline.Toolbutton")
        self.plus_btn.grid(row=3, column=3, columnspan=3, padx=(0, 5), pady=5)

        self.upload_document_btn = ttk.Button(input_fields_frame, text="Upload Doc", command=self.upload_document,
                                              style="danger.Toolbutton")
        self.upload_document_btn.grid(row=3, column=7, padx=5, pady=5)

        self.refresh_days_btn = ttk.Button(input_fields_frame, text="Refresh", command=self.refresh_days,
                                           style="danger.Toolbutton")
        self.refresh_days_btn.grid(row=4, column=7, padx=5, pady=20)

        # En el __init__ de VacationApp, después de input_fields_frame
        self.refresh_days_btn.grid( row=4,column=7,padx=5,pady=20 )

        self.print_btn = ttk.Button( root,text="Save",command=self.print_database,
                                     style="primary.Toolbutton" )

        self.print_btn = ttk.Button(root, text="Save", command=self.print_database,
                                    style="primary.Toolbutton")
        self.print_btn.place(x=10, y=370)

        self.version_label = tk.Label(root, text="Version 1.2", font=("Arial", 12), fg="black")
        self.version_label.place(relx=0.48, rely=0.93, anchor="s")

        self.selected_employee_id = None
        self.preview_window = None
        self.doc_selector = None
        self.preview_label = None
        self.zoom_level = 1.0
        self.load_data()

        self.update_employee_number_state("Company")
        self.root.update()
        self.root.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}")

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

    def update_employee_number_state(self, *args):
        selected_status = args[0]
        if selected_status == "Temp":
            self.employee_number_entry.delete(0, tk.END)
            self.employee_number_entry.config(state="disabled")
        else:
            self.employee_number_entry.config(state="normal")

    def update_status(self, *_):
        if self.selected_employee_id:
            new_status = self.status_var.get()
            try:
                cursor.execute("UPDATE employees SET status = ? WHERE id = ?", (new_status, self.selected_employee_id))
                conn.commit()
                current_values = self.tree.item(self.selected_employee_id, "values")
                self.tree.item(self.selected_employee_id, values=(current_values[0], current_values[1], new_status, *current_values[3:]))
                self.update_employee_number_state(new_status)
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Error updating status: {e}")

    def add_employee(self):
        status = self.status_var.get()
        employee_number_str = self.employee_number_entry.get().strip()

        if status == "Temp":
            employee_number = None
        else:
            if not employee_number_str:
                messagebox.showerror("Error", "Employee Number cannot be empty for Company status!")
                return
            if not employee_number_str.isdigit() or len(employee_number_str) > 3:
                messagebox.showerror("Error", "Employee Number must be a number with maximum 3 digits!")
                return
            employee_number = int(employee_number_str)

        name = self.name_entry.get().strip()
        anniversary_str = self.anniversary_entry.get().strip()
        if not name or not anniversary_str or anniversary_str == "YYYY/MM/DD":
            messagebox.showerror("Error", "Name and Anniversary Date cannot be empty!")
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

            display_name = f"     {employee_id}. {name}"
            employee_number_display = "" if status == "Temp" else employee_number
            self.tree.insert("", "end", iid=employee_id,
                             values=(display_name, employee_number_display, status, anniversary_str, days_taken, days_available, ""))
            self.clear_entries()
        except ValueError:
            messagebox.showerror("Error", "Invalid date format! Use YYYY/MM/DD.")
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Error adding employee: {e}")

    def upload_document(self):
        if not self.selected_employee_id:
            messagebox.showerror("Error", "Please select an employee first!")
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
            self.plus_btn.config(state="normal")
            self.minus_btn.config(state="normal")
            self.delete_employee_btn.config(state="normal")
            cursor.execute("SELECT status FROM employees WHERE id = ?", (self.selected_employee_id,))
            current_status = cursor.fetchone()[0]
            self.status_var.set(current_status)
            self.update_employee_number_state(current_status)
        else:
            self.selected_employee_id = None
            self.plus_btn.config(state="disabled")
            self.minus_btn.config(state="disabled")
            self.delete_employee_btn.config(state="disabled")
            self.close_preview()
            self.status_var.set("Company")

    def on_double_click(self, event):
        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column or not self.selected_employee_id:
            return

        col_index = int(column[1:]) - 1
        if col_index in [4, 5]:  # Days Taken and Days Available
            messagebox.showinfo("Info", "Use the '+' or '−' buttons to adjust days.")
            return

        current_values = self.tree.item(item, "values")
        value_to_edit = current_values[col_index]

        if col_index == 6 and value_to_edit:  # Document column
            self.show_preview()
            return

        if col_index == 1 and not value_to_edit:
            return

        x, y, width, height = self.tree.bbox(item, column)

        if col_index == 2:  # Status column
            status_var = tk.StringVar(value=value_to_edit)
            menu = ttk.OptionMenu(self.root, status_var, value_to_edit, "Company", "Temp",
                                  command=lambda new_val: self.save_status_edit(item, new_val))
            menu.place(x=x + self.tree.winfo_x(), y=y + self.tree.winfo_y(), width=width, height=height)
            menu.focus_set()
            menu.bind("<FocusOut>", lambda _: menu.destroy())
        else:
            entry = tk.Entry(self.root, borderwidth=1, highlightthickness=1)
            entry.place(x=x + self.tree.winfo_x(), y=y + self.tree.winfo_y(), width=width, height=height)
            if col_index == 0:
                value_to_edit = value_to_edit.strip().split(". ", 1)[1]
            entry.insert(0, value_to_edit)
            entry.focus_set()

            def save_edit(_):
                new_value = entry.get().strip()
                if new_value:
                    try:
                        if col_index == 0:  # Employee Name
                            cursor.execute("UPDATE employees SET name = ? WHERE id = ?", (new_value, item))
                            new_display_name = f"     {item}. {new_value}"
                            self.tree.item(item, values=(new_display_name, *current_values[1:]))
                        elif col_index == 1:  # Employee Number
                            if not new_value.isdigit() or len(new_value) > 3:
                                messagebox.showerror("Error", "Employee Number must be a number with max 3 digits!")
                                return
                            cursor.execute("UPDATE employees SET employee_number = ? WHERE id = ?", (int(new_value), item))
                            self.tree.item(item, values=(current_values[0], new_value, *current_values[2:]))
                        elif col_index == 3:  # Anniversary
                            datetime.datetime.strptime(new_value, "%Y/%m/%d")
                            cursor.execute("UPDATE employees SET anniversary = ? WHERE id = ?", (new_value, item))
                            self.tree.item(item, values=(current_values[0], current_values[1], current_values[2], new_value, *current_values[4:]))
                        elif col_index == 6:  # Document
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
                        messagebox.showerror("Error", f"Invalid input: {str(e)}")
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
            messagebox.showerror("Database Error", f"Error updating status: {e}")

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
        ttk.Button(zoom_frame, text="+", width=2, command=self.zoom_in, style="dark.Outline.Toolbutton").pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="−", width=2, command=self.zoom_out, style="dark.Outline.Toolbutton").pack(side=tk.LEFT, padx=2)

        self.preview_label = tk.Label(self.preview_window)
        self.preview_label.pack()
        self.zoom_level = 1.0
        self.update_preview(None)

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

            base_size = 400
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

    def adjust_days(self, change):
        if not self.selected_employee_id:
            return
        try:
            cursor.execute("SELECT days_taken, days_available, anniversary FROM employees WHERE id = ?",
                           (self.selected_employee_id,))
            row = cursor.fetchone()
            if not row:
                return
            days_taken, days_available, anniversary_str = row
            anniversary = datetime.datetime.strptime(anniversary_str, "%Y/%m/%d")
            total_days = calculate_vacation_days(anniversary)

            new_days_taken = max(0, min(total_days, days_taken + change))
            new_days_available = total_days - new_days_taken

            cursor.execute("UPDATE employees SET days_taken = ?, days_available = ? WHERE id = ?",
                           (new_days_taken, new_days_available, self.selected_employee_id))
            conn.commit()

            current_values = self.tree.item(self.selected_employee_id, "values")
            self.tree.item(self.selected_employee_id, values=(
                current_values[0], current_values[1], current_values[2], current_values[3],
                new_days_taken, new_days_available, current_values[6]))
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Error adjusting days: {e}")

    def increase_days(self):
        self.adjust_days(1)

    def decrease_days(self):
        self.adjust_days(-1)

    def delete_employee(self):
        if not self.selected_employee_id:
            messagebox.showerror("Error", "Please select an employee to delete!")
            return

        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this employee?"):
            try:
                cursor.execute("DELETE FROM employees WHERE id = ?", (self.selected_employee_id,))
                conn.commit()
                self.tree.delete(self.selected_employee_id)
                self.selected_employee_id = None
                self.plus_btn.config(state="disabled")
                self.minus_btn.config(state="disabled")
                self.delete_employee_btn.config(state="disabled")
                self.close_preview()
                self.status_var.set("Company")
                messagebox.showinfo("Success", "Employee deleted successfully!")
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Error deleting employee: {e}")

    def load_data(self):
        try:
            self.tree.delete(*self.tree.get_children())
            for row in cursor.execute(
                    "SELECT id, employee_number, name, status, anniversary, days_taken, days_available, document_path FROM employees ORDER BY name"):
                employee_id, employee_number, name, status, anniversary, days_taken, days_available, doc_path = row
                display_name = f"     {employee_id}. {name}"
                anniversary_date = datetime.datetime.strptime(anniversary, "%Y/%m/%d")
                total_days = calculate_vacation_days(anniversary_date)
                updated_available = total_days - days_taken
                if updated_available != days_available:
                    cursor.execute("UPDATE employees SET days_available = ? WHERE id = ?",
                                   (updated_available, employee_id))
                    conn.commit()
                doc_name = ""
                if doc_path and ";" in doc_path:
                    doc_name = doc_path.split(";")[-1].split("|")[0]
                elif doc_path:
                    doc_name = doc_path.split("|")[0]
                employee_number_str = "" if employee_number is None else str(employee_number)
                self.tree.insert("", "end", iid=employee_id,
                                 values=(display_name, employee_number_str, status, anniversary, days_taken, updated_available, doc_name))
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Error loading data: {e}")

    def refresh_days(self):
        self.load_data()

    def print_database(self):
        try:
            output = "Employee Database Report\n"
            output += "=" * 120 + "\n"
            # Adjusted header with left-aligned ID.Name
            header = f"{'ID.Name':<20} {'#':^10} {'Status':^15} {'Anniversary':^20} {'Days Taken':^15} {'Days Available':^15} {'Document':^25}"
            output += header + "\n"
            output += "=" * 120 + "\n"

            for row in cursor.execute(
                    "SELECT id, employee_number, name, status, anniversary, days_taken, days_available, document_path FROM employees ORDER BY name"):
                employee_id, employee_number, name, status, anniversary, days_taken, days_available, doc_path = row
                # Consistent left margin for ID.Name (5 spaces before the number)
                display_name = f"     {employee_id}. {name}"  # 5 spaces before ID
                employee_number_str = "" if employee_number is None else str(employee_number)
                doc_name = ""
                if doc_path and ";" in doc_path:
                    doc_name = doc_path.split(";")[-1].split("|")[0]
                elif doc_path:
                    doc_name = doc_path.split("|")[0]

                # Left-aligned ID.Name with consistent spacing
                line = f"{display_name:<20} {employee_number_str:^10} {status:^15} {anniversary:^20} {days_taken:^15} {days_available:^15} {doc_name:^25}"
                output += line + "\n"

            output += "=" * 120 + "\n"

            print_window = tk.Toplevel(self.root)
            print_window.title("Print Database")
            print_window.geometry("800x400")

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
                    messagebox.showinfo("Success", "Database report saved successfully!")
                    print_window.destroy()

            save_btn = ttk.Button(print_window, text="Save to File", command=save_to_file, style="danger.Toolbutton")
            save_btn.pack(pady=10)

        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"Error printing database: {e}")

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