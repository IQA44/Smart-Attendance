import os
import time
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import simpledialog, messagebox, ttk

from PIL import Image, ImageTk

import pandas as pd  

import serial
import serial.tools.list_ports

from app.constants import (
    APP_NAME,
    SCHOOL_NAME,
    STAGES,
    DEPARTMENTS,
    PROGRAM_STORAGE,
    STUDENTS_FOLDER,
    ATTENDANCE_FOLDER,
    CARDS_FILE,
    STAGES_FILE,
    RECORDS_FOLDER,
    BAUD_RATE,
    WINDOW_WIDTH,
    WINDOW_HEIGHT,
)

from app.storage import (
    create_folders,
    load_data,
    save_data,
    get_stage_department_path,
    load_stage_department_students,
    save_stage_department_students,
    load_stage_department_attendance,
    save_stage_department_attendance,
)

from app.export import export_data



_AR_FIRST_NAMES = [
    "عبدالرحمن", "محمد", "علي", "حسين", "مصطفى", "حيدر", "أحمد", "عمر", "سعد", "ياسر",
    "كريم", "وليد", "فارس", "زياد", "رامي", "باسل", "قيس", "ناصر", "وسام", "مهدي",
    "جلال", "سليم", "شادي", "صفاء", "كرار", "نبيل", "سامر", "زهير", "حارث", "مازن"
]
_AR_LETTERS = list("ابتثجحخدذرزسشصضطظعغفقكلمنهوي")


def _fake_name(seed_idx: int) -> str:
    fn = _AR_FIRST_NAMES[seed_idx % len(_AR_FIRST_NAMES)]
    l1 = _AR_LETTERS[(seed_idx * 3 + 5) % len(_AR_LETTERS)]
    l2 = _AR_LETTERS[(seed_idx * 7 + 11) % len(_AR_LETTERS)]
    return f"{fn} {l1}.{l2}"


def ensure_default_fake_students(min_count_per_class: int = 10):
    """
    إذا ملفات الطلاب فاضية: يضيف 10 طلاب وهميين لكل مرحلة/قسم.
    """
    for stage in STAGES:
        for department in DEPARTMENTS:
            students = load_stage_department_students(stage, department)
            if students and len(students) >= min_count_per_class:
                continue

            existing_names = {s.get("name") for s in students}
            base_seed = abs(hash(f"{stage}|{department}")) % 10_000

            while len(students) < min_count_per_class:
                idx = base_seed + len(students)
                name = _fake_name(idx)
                if name in existing_names:
                    base_seed += 1
                    continue
                students.append({"name": name, "stage": stage, "department": department})
                existing_names.add(name)

            students.sort(key=lambda x: x["name"])
            save_stage_department_students(stage, department, students)


def normalize_uid(uid: str) -> str:
    return uid.strip().upper().replace(" ", "")


def auto_detect_port():
    ports = list(serial.tools.list_ports.comports())
    for port in ports:
        if any(keyword in port.description for keyword in ["Arduino", "CH340", "USB Serial"]):
            return port.device
    if ports:
        return ports[0].device
    return None



create_folders()

stages_data = load_data(STAGES_FILE, {})
if not stages_data:
    for st in STAGES:
        stages_data[st] = list(DEPARTMENTS)
    save_data(STAGES_FILE, stages_data)

all_stage_departments = []
for stg, deps in stages_data.items():
    for dep in deps:
        all_stage_departments.append(f"{stg} - {dep}")

card_students = load_data(CARDS_FILE, {})

ensure_default_fake_students(min_count_per_class=10)

all_students = []
for st in STAGES:
    for dep in DEPARTMENTS:
        all_students.extend(load_stage_department_students(st, dep))



class AttendanceApp:
    def __init__(self, master):
        self.master = master
        master.title(f"{APP_NAME} - {SCHOOL_NAME}")
        master.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

        master.configure(bg="white")
        master.tk_setPalette(background='white', foreground='black')
        master.option_add('*Font', 'Arial 10')
        master.option_add('*justify', 'center')
        master.option_add('*Text.direction', 'rtl')
        master.option_add('*Combobox*Listbox.justify', 'center')

        try:
            right_img = Image.open("Zaytoon_Vocational_Logo.png").resize((120, 120))
            self.right_image = ImageTk.PhotoImage(right_img)
            right_label = tk.Label(master, image=self.right_image, bg="white")
            right_label.place(relx=0.95, rely=0.01, anchor="ne")
        except FileNotFoundError:
            pass
        except Exception:
            pass

        try:
            left_img = Image.open("Karkh1_Vocational_Edu_Dept_Logo.png").resize((100, 100))
            self.left_image = ImageTk.PhotoImage(left_img)
            left_label = tk.Label(master, image=self.left_image, bg="white")
            left_label.place(relx=0.05, rely=0.05, anchor="nw")
        except FileNotFoundError:
            pass
        except Exception:
            pass

        self.card_mode_running = False
        self.card_thread = None
        self.selected_port = None
        self.last_uid = None
        self.current_stage = tk.StringVar(value=STAGES[0] if STAGES else "")
        self.current_department = tk.StringVar(value=DEPARTMENTS[0] if DEPARTMENTS else "")
        self.ser = None
        self.showing_records = False
        self.reader_enabled = False

        title_label = tk.Label(
            master,
            text=f"{SCHOOL_NAME}",
            font=("Arial", 18, "bold"),
            fg="#2c3e50",
            bg="white",
            pady=10
        )
        title_label.place(relx=0.5, rely=0.05, anchor="n")

        self.clock_label = tk.Label(master, font=("Arial", 14, "bold"), fg="black", bg="white")
        self.clock_label.place(relx=0.5, rely=0.11, anchor="n")
        self.update_clock()

        center_frame = tk.Frame(master, bg="white")
        center_frame.place(relx=0.5, y=80, anchor="n", width=800, relheight=0.75)

        system_label = tk.Label(center_frame, text="Welcome to the system",
                                font=("Arial", 16, "bold"), fg="#3498db", bg="white")
        system_label.pack(pady=5)

        subtitle_label = tk.Label(center_frame, text="اهلا بك في نظام إدارة الحضور والانصراف",
                                  font=("Arial", 14, "italic"), fg="#f39c12", bg="white")
        subtitle_label.pack()

        stage_dept_frame = tk.Frame(center_frame, bg="white")
        stage_dept_frame.pack(pady=10)

        tk.Label(stage_dept_frame, text="اختر المرحلة:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3)

        self.stage_dropdown = ttk.Combobox(
            stage_dept_frame,
            textvariable=self.current_stage,
            values=STAGES,
            font=("Arial", 11),
            state="readonly",
            width=15
        )
        self.stage_dropdown.grid(row=0, column=1, padx=3, pady=3)
        self.stage_dropdown.bind("<<ComboboxSelected>>", self.on_stage_selected)

        tk.Label(stage_dept_frame, text="اختر القسم:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=2, padx=3, pady=3)

        self.department_dropdown = ttk.Combobox(
            stage_dept_frame,
            textvariable=self.current_department,
            values=DEPARTMENTS,
            font=("Arial", 11),
            state="readonly",
            width=15
        )
        self.department_dropdown.grid(row=0, column=3, padx=3, pady=3)
        self.department_dropdown.bind("<<ComboboxSelected>>", self.on_department_selected)

        tk.Button(
            stage_dept_frame, text="إدارة الأقسام", font=("Arial", 10),
            command=self.manage_departments, bg="#e74c3c", fg="white",
            relief="raised", bd=1
        ).grid(row=0, column=4, padx=3, pady=3)

        main_buttons_frame = tk.Frame(center_frame, bg="white")
        main_buttons_frame.pack(pady=10)

        search_frame = tk.Frame(main_buttons_frame, bg="white")
        search_frame.pack(pady=3)

        tk.Label(search_frame, text="اسم الطالب:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3)

        self.search_entry = tk.Entry(
            search_frame, font=("Arial", 11), justify="center",
            bg="#ecf0f1", fg="#2c3e50", width=20
        )
        self.search_entry.grid(row=0, column=1, padx=3, pady=3)
        self.search_entry.bind("<Return>", lambda event: self.search_student())

        button_style = {"font": ("Arial", 11), "relief": "raised", "bd": 1, "padx": 8, "pady": 3}

        tk.Button(
            search_frame, text="بحث 🔍", **button_style,
            command=self.search_student, bg="#3498db", fg="white"
        ).grid(row=0, column=2, padx=3, pady=3)

        tk.Button(
            search_frame, text="تسجيل حضور/انصراف 📝", **button_style,
            command=self.direct_register, bg="#f39c12", fg="white"
        ).grid(row=0, column=3, padx=3, pady=3)

        control_buttons_frame = tk.Frame(main_buttons_frame, bg="white")
        control_buttons_frame.pack(pady=8)

        tk.Button(
            control_buttons_frame, text="إضافة طالب ➕", **button_style,
            command=self.add_student, bg="#27ae60", fg="white"
        ).grid(row=0, column=0, padx=3, pady=3)

        self.toggle_btn = tk.Button(
            control_buttons_frame, text="عرض السجلات 📋", **button_style,
            command=self.toggle_display, bg="#9b59b6", fg="white"
        )
        self.toggle_btn.grid(row=0, column=1, padx=3, pady=3)

        self.reader_btn = tk.Button(
            control_buttons_frame, text="تشغيل القارئ 🔌", **button_style,
            command=self.toggle_reader, bg="#16a085", fg="white"
        )
        self.reader_btn.grid(row=0, column=2, padx=3, pady=3)

        extra_buttons_frame = tk.Frame(main_buttons_frame, bg="white")
        extra_buttons_frame.pack(pady=8)

        tk.Button(
            extra_buttons_frame, text="تصدير البيانات 💾", **button_style,
            command=self.export_data, bg="#d35400", fg="white"
        ).grid(row=0, column=0, padx=3, pady=3)

        tk.Button(
            extra_buttons_frame, text="البحث عن أيام الغياب 📊", **button_style,
            command=self.search_absence_days, bg="#8e44ad", fg="white"
        ).grid(row=0, column=1, padx=3, pady=3)

        tk.Button(
            extra_buttons_frame, text="نقل طالب/طلاب 🔄", **button_style,
            command=self.move_students, bg="#e74c3c", fg="white"
        ).grid(row=0, column=2, padx=3, pady=3)

        self.display_frame = tk.Frame(center_frame, bg="white")
        self.display_frame.pack(fill="both", expand=True, pady=10, padx=15)

        self.create_student_list()
        self.create_records_tree()

        self.current_matches = []
        self.search_student()
        self.show_student_list()

    def update_clock(self):
        current_time = datetime.now().strftime("%I:%M:%S %p")
        self.clock_label.config(text=current_time)
        self.master.after(1000, self.update_clock)

    
    def create_student_list(self):
        self.student_frame = tk.Frame(self.display_frame, bg="white")

        self.student_listbox = tk.Listbox(
            self.student_frame, font=("Arial", 11), height=6,
            bg="#ecf0f1", fg="#2c3e50", justify="center",
            selectbackground="#3498db", selectforeground="white"
        )

        student_scrollbar = tk.Scrollbar(self.student_frame, orient="vertical", command=self.student_listbox.yview)
        self.student_listbox.configure(yscrollcommand=student_scrollbar.set)

        self.student_listbox.pack(side="left", fill="both", expand=True)
        student_scrollbar.pack(side="right", fill="y")

        self.student_listbox.bind("<Double-Button-1>", lambda event: self.direct_register())

    def create_records_tree(self):
        self.tree_frame = tk.Frame(self.display_frame, bg="white")

        self.records_tree = ttk.Treeview(
            self.tree_frame,
            columns=("الانصراف", "الحضور", "التاريخ", "القسم", "المرحلة", "الاسم"),
            show="headings",
            height=10
        )

        self.records_tree.heading("الاسم", text="اسم الطالب", anchor='center')
        self.records_tree.heading("المرحلة", text="المرحلة", anchor='center')
        self.records_tree.heading("القسم", text="القسم", anchor='center')
        self.records_tree.heading("التاريخ", text="التاريخ", anchor='center')
        self.records_tree.heading("الحضور", text="الحضور", anchor='center')
        self.records_tree.heading("الانصراف", text="الانصراف", anchor='center')

        self.records_tree.column("الاسم", width=150, anchor="center")
        self.records_tree.column("المرحلة", width=80, anchor="center")
        self.records_tree.column("القسم", width=80, anchor="center")
        self.records_tree.column("التاريخ", width=80, anchor="center")
        self.records_tree.column("الحضور", width=80, anchor="center")
        self.records_tree.column("الانصراف", width=80, anchor="center")

        self.records_tree["displaycolumns"] = ("الانصراف", "الحضور", "التاريخ", "القسم", "المرحلة", "الاسم")

        tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.records_tree.yview)
        self.records_tree.configure(yscrollcommand=tree_scrollbar.set)

        self.records_tree.pack(side="left", fill="both", expand=True)
        tree_scrollbar.pack(side="right", fill="y")

    def toggle_display(self):
        if self.showing_records:
            self.show_student_list()
            self.toggle_btn.config(text="عرض السجلات 📋")
        else:
            self.show_records()
            self.toggle_btn.config(text="عرض الطلاب 👥")
        self.showing_records = not self.showing_records

    def show_student_list(self):
        self.tree_frame.pack_forget()
        self.student_frame.pack(fill="both", expand=True)
        self.search_student()

    def show_records(self):
        self.student_frame.pack_forget()
        self.tree_frame.pack(fill="both", expand=True)
        self.update_records_display()

    def update_records_display(self):
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)

        selected_stage = self.current_stage.get()
        selected_department = self.current_department.get()
        stage_department_attendance = load_stage_department_attendance(selected_stage, selected_department)

        for student_key, dates in stage_department_attendance.items():
            name, stage, department = student_key.split("|")
            for date, records in dates.items():
                attendance_times = [rec['time'] for rec in records if rec['type'] == 'حضور']
                departure_times = [rec['time'] for rec in records if rec['type'] == 'انصراف']

                attendance_str = "\n".join(attendance_times) if attendance_times else "غياب"
                departure_str = "\n".join(departure_times) if departure_times else "لم ينصرف"

                self.records_tree.insert(
                    "", "end",
                    values=(departure_str, attendance_str, date, department, stage, name)
                )

    def toggle_reader(self):
        if not self.reader_enabled:
            auto_port = auto_detect_port()
            if auto_port is None:
                messagebox.showerror("❌ خطأ", "لم يتم التعرف على جهاز الاردوينو تلقائيًا.\nيرجى اختيار المنفذ يدويًا.")
                chosen_port = self.choose_serial_port()
                if not chosen_port:
                    return
                self.selected_port = chosen_port
            else:
                self.selected_port = auto_port
                messagebox.showinfo("✅ نجح", f"تم التعرف على جهاز الاردوينو تلقائيًا على المنفذ {auto_port}")

            self.reader_enabled = True
            self.reader_btn.config(text="إطفاء القارئ ⏹️", bg="#e74c3c")
            messagebox.showinfo("✅ نجح", "تم تشغيل القارئ بنجاح")

            self.start_card_mode()
        else:
            self.reader_enabled = False
            self.reader_btn.config(text="تشغيل القارئ 🔌", bg="#16a085")
            self.stop_card_mode()
            messagebox.showinfo("✅ نجح", "تم إيقاف القارئ بنجاح")

    def start_card_mode(self):
        if not self.card_mode_running:
            self.card_mode_running = True
            self.card_thread = threading.Thread(target=self.card_mode_worker, daemon=True)
            self.card_thread.start()

    def stop_card_mode(self):
        self.card_mode_running = False
        if self.ser and self.ser.is_open:
            self.ser.close()

    def choose_serial_port(self):
        ports = list(serial.tools.list_ports.comports())
        if not ports:
            messagebox.showerror("❌ خطأ", "لا يوجد أجهزة USB متصلة.")
            return None

        port_window = tk.Toplevel(self.master)
        port_window.title("اختر منفذ USB")
        port_window.configure(bg="white")
        port_window.option_add('*Font', 'Arial 10')
        port_window.option_add('*justify', 'center')

        tk.Label(port_window, text="اختر منفذ USB:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").pack(pady=8)

        listbox_frame = tk.Frame(port_window, bg="white")
        listbox_frame.pack(pady=8, padx=8, fill="both", expand=True)

        listbox = tk.Listbox(listbox_frame, font=("Arial", 11), justify="center",
                             bg="#ecf0f1", fg="#2c3e50", selectbackground="#3498db", height=6)
        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)

        for port in ports:
            listbox.insert(tk.END, f"{port.device} - {port.description}")

        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        selected = {}

        def select_and_close():
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                selected['port'] = ports[index].device
                port_window.destroy()
            else:
                messagebox.showwarning("⚠️ تنبيه", "يرجى اختيار منفذ.")

        tk.Button(port_window, text="اختيار", command=select_and_close,
                  font=("Arial", 11), bg="#3498db", fg="white",
                  relief="raised", bd=1).pack(pady=8)

        self.master.wait_window(port_window)
        return selected.get('port')

    def card_mode_worker(self):
        try:
            self.ser = serial.Serial(self.selected_port, BAUD_RATE, timeout=1)
        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror("❌ خطأ", f"تعذر فتح منفذ البطاقة:\n{e}"))
            self.stop_card_mode()
            return

        while self.card_mode_running and self.reader_enabled:
            try:
                if self.ser.in_waiting:
                    uid_line = self.ser.readline().decode("utf-8").strip()
                    if uid_line:
                        uid = normalize_uid(uid_line)
                        self.master.after(0, self.process_card, uid)
                time.sleep(0.1)
            except Exception as e:
                self.master.after(0, lambda: messagebox.showerror("❌ خطأ", f"خطأ في قراءة البطاقة:\n{e}"))
                break

        if self.ser and self.ser.is_open:
            self.ser.close()

    def on_stage_selected(self, event):
        self.search_student()
        if self.showing_records:
            self.update_records_display()

    def on_department_selected(self, event):
        self.search_student()
        if self.showing_records:
            self.update_records_display()

    def manage_departments(self):
        manage_window = tk.Toplevel(self.master)
        manage_window.title("إدارة الأقسام")
        manage_window.geometry("400x350")
        manage_window.configure(bg="white")
        manage_window.option_add('*Font', 'Arial 10')
        manage_window.option_add('*justify', 'center')

        tk.Label(manage_window, text="إدارة الأقسام الدراسية", font=("Arial", 14, "bold"),
                 fg="#f39c12", bg="white").pack(pady=8)

        listbox_frame = tk.Frame(manage_window, bg="white")
        listbox_frame.pack(pady=8, padx=8, fill="both", expand=True)

        tk.Label(listbox_frame, text="الأقسام الحالية:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").pack(pady=3)

        listbox = tk.Listbox(listbox_frame, font=("Arial", 11), height=8, justify="center",
                             bg="#ecf0f1", fg="#2c3e50", selectbackground="#3498db")

        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)

        for department in DEPARTMENTS:
            listbox.insert(tk.END, department)

        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = tk.Frame(manage_window, bg="white")
        btn_frame.pack(pady=8)

        def add_department():
            department = simpledialog.askstring("إضافة قسم", "أدخل اسم القسم الجديد:")
            if department and department not in DEPARTMENTS:
                DEPARTMENTS.append(department)
                self.department_dropdown['values'] = DEPARTMENTS
                listbox.insert(tk.END, department)
                messagebox.showinfo("نجاح", f"تم إضافة القسم {department} بنجاح")
            elif department in DEPARTMENTS:
                messagebox.showwarning("تحذير", "هذا القسم موجود مسبقاً!")

        def remove_department():
            selection = listbox.curselection()
            if selection:
                department = listbox.get(selection[0])
                if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف القسم {department}؟"):
                    try:
                        DEPARTMENTS.remove(department)
                    except ValueError:
                        pass
                    self.department_dropdown['values'] = DEPARTMENTS
                    listbox.delete(selection[0])
                    messagebox.showinfo("نجاح", f"تم حذف القسم {department} بنجاح")
            else:
                messagebox.showwarning("تحذير", "يرجى اختيار قسم للحذف!")

        tk.Button(btn_frame, text="إضافة قسم", font=("Arial", 10),
                  command=add_department, bg="#27ae60", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=3)

        tk.Button(btn_frame, text="حذف قسم", font=("Arial", 10),
                  command=remove_department, bg="#e74c3c", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=3)

        tk.Button(btn_frame, text="إغلاق", font=("Arial", 10),
                  command=manage_window.destroy, bg="#7f8c8d", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=3)

    def search_student(self):
        search_text = self.search_entry.get().strip()
        self.student_listbox.delete(0, tk.END)

        stage_department_students = load_stage_department_students(
            self.current_stage.get(),
            self.current_department.get()
        )

        if not search_text:
            self.current_matches = stage_department_students
            for idx, student in enumerate(stage_department_students, 1):
                self.student_listbox.insert(tk.END, f"{student['name']} - {idx}")
            return

        self.current_matches = [
            s for s in stage_department_students
            if search_text.lower() in s["name"].lower()
        ]

        if not self.current_matches:
            self.student_listbox.insert(tk.END, "لا يوجد طالب بهذا الاسم في هذه المرحلة والقسم.")
        else:
            for idx, student in enumerate(self.current_matches, 1):
                self.student_listbox.insert(tk.END, f"{student['name']} - {idx}")

    def add_student(self):
        add_window = tk.Toplevel(self.master)
        add_window.title("إضافة طالب جديد")
        add_window.geometry("350x200")
        add_window.configure(bg="white")
        add_window.option_add('*Font', 'Arial 10')
        add_window.option_add('*justify', 'center')
        add_window.option_add('*Entry.justify', 'center')

        tk.Label(add_window, text="الاسم الثلاثي:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3, sticky="e")

        name_entry = tk.Entry(add_window, font=("Arial", 11), justify="center", bg="#ecf0f1")
        name_entry.grid(row=0, column=1, padx=3, pady=3)

        tk.Label(add_window, text="المرحلة:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=1, column=0, padx=3, pady=3, sticky="e")

        stage_var = tk.StringVar(add_window)
        stage_var.set(self.current_stage.get())
        stage_dropdown = ttk.Combobox(add_window, textvariable=stage_var,
                                      values=STAGES, font=("Arial", 11), state="readonly")
        stage_dropdown.grid(row=1, column=1, padx=3, pady=3)

        tk.Label(add_window, text="القسم:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=2, column=0, padx=3, pady=3, sticky="e")

        department_var = tk.StringVar(add_window)
        department_var.set(self.current_department.get())
        department_dropdown = ttk.Combobox(add_window, textvariable=department_var,
                                           values=DEPARTMENTS, font=("Arial", 11), state="readonly")
        department_dropdown.grid(row=2, column=1, padx=3, pady=3)

        def save_student():
            name = name_entry.get().strip()
            stage = stage_var.get()
            department = department_var.get()

            if not name:
                messagebox.showwarning("⚠️ تنبيه", "يرجى إدخال اسم الطالب.")
                return

            stage_department_students = load_stage_department_students(stage, department)

            for student in stage_department_students:
                if student["name"] == name:
                    messagebox.showinfo("⚠️ تنبيه", "الطالب موجود مسبقًا في هذه المرحلة والقسم!")
                    return

            new_student = {"name": name, "stage": stage, "department": department}
            stage_department_students.append(new_student)
            stage_department_students.sort(key=lambda x: x["name"])

            if save_stage_department_students(stage, department, stage_department_students):
                global all_students
                all_students = []
                for stg in STAGES:
                    for dept in DEPARTMENTS:
                        all_students.extend(load_stage_department_students(stg, dept))

                messagebox.showinfo("✅ نجح", f"تم إضافة الطالب {name} بنجاح")
                add_window.destroy()
                self.search_student()
                if self.showing_records:
                    self.update_records_display()

        def cancel():
            add_window.destroy()

        btn_frame = tk.Frame(add_window, bg="white")
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)

        tk.Button(btn_frame, text="حفظ", font=("Arial", 11),
                  command=save_student, bg="#27ae60", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=8)

        tk.Button(btn_frame, text="إلغاء", font=("Arial", 11),
                  command=cancel, bg="#e74c3c", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=8)

        name_entry.focus_set()

    def direct_register(self):
        if self.student_listbox.curselection():
            selected_index = self.student_listbox.curselection()[0]
            if selected_index < len(self.current_matches):
                student = self.current_matches[selected_index]
                self.record_attendance(student)
                return

        search_text = self.search_entry.get().strip()
        if not search_text:
            messagebox.showwarning("⚠️ تنبيه", "يرجى إدخال اسم الطالب.")
            return

        stage_department_students = load_stage_department_students(
            self.current_stage.get(),
            self.current_department.get()
        )
        matches = [s for s in stage_department_students if search_text.lower() == s["name"].lower()]

        if not matches:
            messagebox.showwarning("⚠️ تنبيه", "الطالب غير موجود في هذه المرحلة والقسم. يرجى إضافته أولاً.")
            return
        elif len(matches) > 1:
            messagebox.showwarning("⚠️ تنبيه", "هناك أكثر من طالب لهذا الاسم. يرجى اختياره من القائمة.")
            return
        else:
            student = matches[0]

        self.record_attendance(student)
        self.search_entry.delete(0, tk.END)

    def record_attendance(self, student, operation=None):
        stage_department_attendance = load_stage_department_attendance(student["stage"], student["department"])

        student_key = f"{student['name']}|{student['stage']}|{student['department']}"
        current_time = datetime.now().strftime("%I:%M %p")
        current_date = datetime.now().strftime("%Y-%m-%d")

        if student_key not in stage_department_attendance:
            stage_department_attendance[student_key] = {}

        if current_date not in stage_department_attendance[student_key]:
            stage_department_attendance[student_key][current_date] = []

        if operation is None:
            if not stage_department_attendance[student_key][current_date]:
                op = "حضور"
            else:
                last_action = stage_department_attendance[student_key][current_date][-1]['type']
                op = "انصراف" if last_action == "حضور" else "حضور"
        else:
            op = operation

        stage_department_attendance[student_key][current_date].append({'type': op, 'time': current_time})

        if save_stage_department_attendance(student["stage"], student["department"], stage_department_attendance):
            if self.showing_records:
                self.update_records_display()

    def export_data(self):
        threading.Thread(target=self.export_data_thread, daemon=True).start()

    def export_data_thread(self):
        try:
            export_data()
            self.master.after(0, self.stop_card_mode)
            if self.showing_records:
                self.master.after(0, self.update_records_display)
        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror("❌ خطأ", f"فشل في تصدير البيانات:\n{str(e)}"))
            print(traceback.format_exc())

    def choose_name_option(self):
        option_window = tk.Toplevel(self.master)
        option_window.title("اختيار طريقة تسجيل الاسم")
        option_window.configure(bg="white")
        option_window.option_add('*Font', 'Arial 10')
        option_window.option_add('*justify', 'center')

        tk.Label(option_window, text="اختر طريقة تسجيل الاسم:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").pack(pady=8)

        result = {}

        def choose_existing():
            result['option'] = "existing"
            option_window.destroy()

        def choose_new():
            result['option'] = "new"
            option_window.destroy()

        tk.Button(option_window, text="اختيار اسم موجود", font=("Arial", 11),
                  command=choose_existing, bg="#3498db", fg="white",
                  relief="raised", bd=1).pack(pady=3)

        tk.Button(option_window, text="إدخال اسم جديد", font=("Arial", 11),
                  command=choose_new, bg="#27ae60", fg="white",
                  relief="raised", bd=1).pack(pady=3)

        self.master.wait_window(option_window)
        return result.get('option')

    def prompt_for_existing_name(self):
        select_window = tk.Toplevel(self.master)
        select_window.title("اختيار اسم موجود")
        select_window.geometry("500x400")
        select_window.configure(bg="white")
        select_window.option_add('*Font', 'Arial 10')
        select_window.option_add('*justify', 'center')

        tk.Label(select_window, text="اختر اسم الطالب:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").pack(pady=8)

        search_frame = tk.Frame(select_window, bg="white")
        search_frame.pack(pady=3, fill="x", padx=8)

        tk.Label(search_frame, text="بحث:", font=("Arial", 10),
                 fg="#2c3e50", bg="white").pack(side="left", padx=3)

        search_entry = tk.Entry(search_frame, font=("Arial", 10), justify="center", bg="#ecf0f1")
        search_entry.pack(side="left", fill="x", expand=True, padx=3)

        listbox_frame = tk.Frame(select_window, bg="white")
        listbox_frame.pack(pady=8, padx=8, fill="both", expand=True)

        listbox = tk.Listbox(listbox_frame, font=("Arial", 10), justify="center",
                             bg="#ecf0f1", fg="#2c3e50", selectbackground="#3498db", height=8)

        scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)

        global all_students
        all_students = []
        for stg in STAGES:
            for dept in DEPARTMENTS:
                all_students.extend(load_stage_department_students(stg, dept))

        for student in all_students:
            listbox.insert(tk.END, f"{student['name']} ({student['stage']} - {student['department']})")

        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def search_students_key(*args):
            search_text = search_entry.get().strip().lower()
            listbox.delete(0, tk.END)
            if not search_text:
                for student in all_students:
                    listbox.insert(tk.END, f"{student['name']} ({student['stage']} - {student['department']})")
                return
            for student in all_students:
                if search_text in student['name'].lower():
                    listbox.insert(tk.END, f"{student['name']} ({student['stage']} - {student['department']})")

        search_entry.bind("<KeyRelease>", search_students_key)

        result = {}

        def select_name():
            selection = listbox.curselection()
            if selection:
                selected_text = listbox.get(selection[0])
                name = selected_text.split(" (")[0]
                stage_dept = selected_text.split(" (")[1].replace(")", "")
                stage, department = stage_dept.split(" - ")
                result['student'] = {"name": name, "stage": stage, "department": department}
                select_window.destroy()
            else:
                messagebox.showwarning("⚠️ تنبيه", "يرجى اختيار اسم.")

        btn_frame = tk.Frame(select_window, bg="white")
        btn_frame.pack(pady=8)

        tk.Button(btn_frame, text="اختيار", font=("Arial", 11),
                  command=select_name, bg="#3498db", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=8)

        tk.Button(btn_frame, text="إلغاء", font=("Arial", 11),
                  command=select_window.destroy, bg="#e74c3c", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=8)

        self.master.wait_window(select_window)
        return result.get('student')

    def prompt_for_new_name(self):
        new_window = tk.Toplevel(self.master)
        new_window.title("إضافة طالب جديد")
        new_window.geometry("350x250")
        new_window.configure(bg="white")
        new_window.option_add('*Font', 'Arial 10')
        new_window.option_add('*justify', 'center')
        new_window.option_add('*Entry.justify', 'center')

        tk.Label(new_window, text="الاسم الثلاثي:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3, sticky="e")

        name_entry = tk.Entry(new_window, font=("Arial", 11), justify="center", bg="#ecf0f1")
        name_entry.grid(row=0, column=1, padx=3, pady=3)

        tk.Label(new_window, text="المرحلة:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=1, column=0, padx=3, pady=3, sticky="e")

        stage_var = tk.StringVar(new_window)
        stage_var.set(self.current_stage.get())
        stage_dropdown = ttk.Combobox(new_window, textvariable=stage_var,
                                      values=STAGES, font=("Arial", 11), state="readonly")
        stage_dropdown.grid(row=1, column=1, padx=3, pady=3)

        tk.Label(new_window, text="القسم:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=2, column=0, padx=3, pady=3, sticky="e")

        department_var = tk.StringVar(new_window)
        department_var.set(self.current_department.get())
        department_dropdown = ttk.Combobox(new_window, textvariable=department_var,
                                           values=DEPARTMENTS, font=("Arial", 11), state="readonly")
        department_dropdown.grid(row=2, column=1, padx=3, pady=3)

        result = {}

        def save_and_close():
            name = name_entry.get().strip()
            stage = stage_var.get()
            department = department_var.get()

            if not name:
                messagebox.showwarning("⚠️ تنبيه", "يرجى إدخال اسم الطالب.")
                return

            stage_department_students = load_stage_department_students(stage, department)
            for student in stage_department_students:
                if student["name"] == name:
                    messagebox.showinfo("⚠️ تنبيه", "الطالب موجود مسبقًا في هذه المرحلة والقسم!")
                    return

            result['student'] = {"name": name, "stage": stage, "department": department}
            new_window.destroy()

        def cancel():
            new_window.destroy()

        btn_frame = tk.Frame(new_window, bg="white")
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)

        tk.Button(btn_frame, text="حفظ", font=("Arial", 11),
                  command=save_and_close, bg="#27ae60", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=8)

        tk.Button(btn_frame, text="إلغاء", font=("Arial", 11),
                  command=cancel, bg="#e74c3c", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=8)

        self.master.wait_window(new_window)
        return result.get('student')

    def process_card(self, uid):
        global card_students

        if uid in card_students:
            student_data = card_students[uid]
            self.record_attendance(student_data)
            return

        if not messagebox.askyesno("❓ بطاقة غير مسجلة", f"البطاقة ({uid}) غير مسجلة.\nهل تريد تسجيلها؟"):
            return

        option = self.choose_name_option()
        if option == "existing":
            student = self.prompt_for_existing_name()
            if not student:
                return
        elif option == "new":
            student = self.prompt_for_new_name()
            if not student:
                return
        else:
            return

        card_students[uid] = student
        if save_data(CARDS_FILE, card_students):
            found = False
            stage_department_students = load_stage_department_students(student["stage"], student["department"])
            for s in stage_department_students:
                if s["name"] == student["name"]:
                    found = True
                    break

            if not found:
                stage_department_students.append(student)
                stage_department_students.sort(key=lambda x: x["name"])
                save_stage_department_students(student["stage"], student["department"], stage_department_students)

                global all_students
                all_students = []
                for stg in STAGES:
                    for dept in DEPARTMENTS:
                        all_students.extend(load_stage_department_students(stg, dept))

            messagebox.showinfo("✅ نجح", f"تم ربط البطاقة بالطالب {student['name']} بنجاح")
            self.record_attendance(student)

    def search_absence_days(self):
        absence_window = tk.Toplevel(self.master)
        absence_window.title("البحث عن أيام الغياب")
        absence_window.geometry("450x250")
        absence_window.configure(bg="white")
        absence_window.option_add('*Font', 'Arial 10')
        absence_window.option_add('*justify', 'center')
        absence_window.option_add('*Entry.justify', 'center')

        absence_window.transient(self.master)
        absence_window.grab_set()
        absence_window.focus_set()

        tk.Label(absence_window, text="البحث عن أيام الغياب", font=("Arial", 14, "bold"),
                 fg="#f39c12", bg="white").pack(pady=8)

        class_frame = tk.Frame(absence_window, bg="white")
        class_frame.pack(pady=3)

        tk.Label(class_frame, text="اختر المرحلة:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3)

        stage_var = tk.StringVar(value=self.current_stage.get())
        stage_dropdown = ttk.Combobox(class_frame, textvariable=stage_var,
                                      values=STAGES, font=("Arial", 11), state="readonly")
        stage_dropdown.grid(row=0, column=1, padx=3, pady=3)

        tk.Label(class_frame, text="اختر القسم:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=2, padx=3, pady=3)

        department_var = tk.StringVar(value=self.current_department.get())
        department_dropdown = ttk.Combobox(class_frame, textvariable=department_var,
                                           values=DEPARTMENTS, font=("Arial", 11), state="readonly")
        department_dropdown.grid(row=0, column=3, padx=3, pady=3)

        student_frame = tk.Frame(absence_window, bg="white")
        student_frame.pack(pady=3)

        tk.Label(student_frame, text="اسم الطالب:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3)

        student_var = tk.StringVar()
        student_entry = tk.Entry(student_frame, textvariable=student_var, font=("Arial", 11),
                                 justify="center", bg="#ecf0f1")
        student_entry.grid(row=0, column=1, padx=3, pady=3)

        search_btn = tk.Button(
            absence_window, text="بدء البحث 🔍", font=("Arial", 11),
            command=lambda: threading.Thread(
                target=self.calculate_absence_days,
                args=(stage_var.get(), department_var.get(), student_var.get(), absence_window),
                daemon=True
            ).start(),
            bg="#16a085", fg="white", relief="raised", bd=1
        )
        search_btn.pack(pady=8)

        self.progress_label = tk.Label(absence_window, text="", font=("Arial", 10),
                                       fg="#2c3e50", bg="white")
        self.progress_label.pack(pady=3)

        self.result_label = tk.Label(absence_window, text="", font=("Arial", 12, "bold"),
                                     fg="#f39c12", bg="white")
        self.result_label.pack(pady=8)

    def calculate_absence_days(self, stage, department, student_name, window):
        if not student_name.strip():
            self.master.after(0, lambda: messagebox.showwarning("⚠️ تنبيه", "يرجى إدخال اسم الطالب."))
            return

        stage_department_students = load_stage_department_students(stage, department)
        matches = [s for s in stage_department_students if student_name.strip().lower() == s["name"].lower()]

        if not matches:
            self.master.after(0, lambda: messagebox.showwarning("⚠️ تنبيه", "الطالب غير موجود في هذه المرحلة والقسم."))
            return

        student = matches[0]
        class_folder = os.path.join(RECORDS_FOLDER, stage, department)

        if not os.path.exists(class_folder):
            self.master.after(0, lambda: messagebox.showinfo("⚠️ تنبيه", "لا توجد سجلات لهذه المرحلة والقسم."))
            return

        excel_files = [f for f in os.listdir(class_folder) if f.endswith('.xlsx')]
        total_files = len(excel_files)

        if total_files == 0:
            self.master.after(0, lambda: messagebox.showinfo("⚠️ تنبيه", "لا توجد سجلات لهذه المرحلة والقسم."))
            return

        absence_count = 0

        for i, file_name in enumerate(excel_files):
            file_path = os.path.join(class_folder, file_name)
            self.master.after(0, lambda i=i: self.progress_label.config(text=f"جارٍ البحث في الملف {i+1}/{total_files}"))

            try:
                df = pd.read_excel(file_path)
                student_row = df[df['اسم'] == student['name']]
                if not student_row.empty:
                    attendance_status = student_row['الحضور'].values[0]
                    if attendance_status == "غياب":
                        absence_count += 1
            except Exception:
                pass

        self.master.after(0, lambda: self.progress_label.config(text="اكتمل البحث"))
        result_text = f"عدد أيام الغياب للطالب {student['name']}: {absence_count} من أصل {total_files} يوم"
        self.master.after(0, lambda: self.result_label.config(text=result_text))

    def move_students(self):
        move_window = tk.Toplevel(self.master)
        move_window.title("نقل الطلاب بين الأقسام")
        move_window.geometry("550x450")
        move_window.configure(bg="white")
        move_window.option_add('*Font', 'Arial 10')
        move_window.option_add('*justify', 'center')

        tk.Label(move_window, text="نقل الطلاب بين الأقسام", font=("Arial", 14, "bold"),
                 fg="#f39c12", bg="white").pack(pady=8)

        source_frame = tk.Frame(move_window, bg="white")
        source_frame.pack(pady=3)

        tk.Label(source_frame, text="المرحلة المصدر:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3)

        source_stage_var = tk.StringVar(value=self.current_stage.get())
        source_stage_dropdown = ttk.Combobox(source_frame, textvariable=source_stage_var,
                                             values=STAGES, font=("Arial", 11), state="readonly")
        source_stage_dropdown.grid(row=0, column=1, padx=3, pady=3)

        tk.Label(source_frame, text="القسم المصدر:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=2, padx=3, pady=3)

        source_department_var = tk.StringVar(value=self.current_department.get())
        source_department_dropdown = ttk.Combobox(source_frame, textvariable=source_department_var,
                                                  values=DEPARTMENTS, font=("Arial", 11), state="readonly")
        source_department_dropdown.grid(row=0, column=3, padx=3, pady=3)

        target_frame = tk.Frame(move_window, bg="white")
        target_frame.pack(pady=3)

        tk.Label(target_frame, text="المرحلة الهدف:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=0, padx=3, pady=3)

        target_stage_var = tk.StringVar(value=STAGES[0] if STAGES else "")
        target_stage_dropdown = ttk.Combobox(target_frame, textvariable=target_stage_var,
                                             values=STAGES, font=("Arial", 11), state="readonly")
        target_stage_dropdown.grid(row=0, column=1, padx=3, pady=3)

        tk.Label(target_frame, text="القسم الهدف:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").grid(row=0, column=2, padx=3, pady=3)

        target_department_var = tk.StringVar(value=DEPARTMENTS[0] if DEPARTMENTS else "")
        target_department_dropdown = ttk.Combobox(target_frame, textvariable=target_department_var,
                                                  values=DEPARTMENTS, font=("Arial", 11), state="readonly")
        target_department_dropdown.grid(row=0, column=3, padx=3, pady=3)

        tk.Label(move_window, text="اختر الطلاب للنقل:", font=("Arial", 12),
                 fg="#2c3e50", bg="white").pack(pady=3)

        list_frame = tk.Frame(move_window, bg="white")
        list_frame.pack(pady=3, fill="both", expand=True)

        scroll_frame = tk.Frame(list_frame, bg="white")
        scroll_frame.pack(fill="both", expand=True, padx=8)

        canvas = tk.Canvas(scroll_frame, bg="white")
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        student_vars = {}

        def load_students_ui(*args):
            for widget in scrollable_frame.winfo_children():
                widget.destroy()

            stage = source_stage_var.get()
            department = source_department_var.get()
            students = load_stage_department_students(stage, department)

            if not students:
                tk.Label(scrollable_frame, text="لا يوجد طلاب في هذه المرحلة والقسم",
                         font=("Arial", 10), fg="#2c3e50", bg="white").pack(pady=8)
                return

            student_vars.clear()
            for student in students:
                var = tk.BooleanVar()
                student_vars[student['name']] = var

                frame = tk.Frame(scrollable_frame, bg="white")
                frame.pack(fill="x", pady=1)

                tk.Label(frame, text=student['name'], font=("Arial", 10),
                         fg="#2c3e50", bg="white").pack(side="right", padx=(0, 8))
                tk.Checkbutton(frame, variable=var, bg="white").pack(side="right")

        source_stage_dropdown.bind("<<ComboboxSelected>>", load_students_ui)
        source_department_dropdown.bind("<<ComboboxSelected>>", load_students_ui)

        load_students_ui()

        def move_selected_students():
            global card_students

            source_stage = source_stage_var.get()
            source_department = source_department_var.get()
            target_stage = target_stage_var.get()
            target_department = target_department_var.get()

            if source_stage == target_stage and source_department == target_department:
                messagebox.showwarning("تحذير", "لا يمكن نقل الطلاب إلى نفس المرحلة والقسم!")
                return

            selected_students = [name for name, var in student_vars.items() if var.get()]
            if not selected_students:
                messagebox.showwarning("تحذير", "يرجى اختيار طالب واحد على الأقل!")
                return

            if not messagebox.askyesno(
                "تأكيد",
                f"هل أنت متأكد من نقل {len(selected_students)} طالب من {source_stage} - {source_department} إلى {target_stage} - {target_department}؟"
            ):
                return

            source_students = load_stage_department_students(source_stage, source_department)
            target_students = load_stage_department_students(target_stage, target_department)
            target_names = {s['name'] for s in target_students}

            moved_students = []
            skipped_duplicates = []
            for student_name in selected_students:
                student_to_move = None
                for student in source_students:
                    if student['name'] == student_name:
                        student_to_move = student
                        break

                if student_to_move:
                    if student_name in target_names:
                        skipped_duplicates.append(student_name)
                        continue

                    target_students.append({"name": student_name, "stage": target_stage, "department": target_department})
                    moved_students.append(student_name)
                    target_names.add(student_name)

                    for uid, student_data in card_students.items():
                        if (student_data['name'] == student_name and
                                student_data['stage'] == source_stage and
                                student_data['department'] == source_department):
                            card_students[uid]['stage'] = target_stage
                            card_students[uid]['department'] = target_department

            if moved_students:
                source_students = [s for s in source_students if s['name'] not in moved_students]

                save_stage_department_students(source_stage, source_department, source_students)
                save_stage_department_students(target_stage, target_department, target_students)
                save_data(CARDS_FILE, card_students)

                global all_students
                all_students = []
                for stg in STAGES:
                    for dept in DEPARTMENTS:
                        all_students.extend(load_stage_department_students(stg, dept))

                messagebox.showinfo("✅ نجح", f"تم نقل {len(moved_students)} طالب بنجاح")
                move_window.destroy()

                self.search_student()
                if self.showing_records:
                    self.update_records_display()
            else:
                if skipped_duplicates:
                    messagebox.showwarning("تحذير", "الطلاب المحددون موجودون مسبقًا في القسم الهدف!")
                else:
                    messagebox.showwarning("تحذير", "لم يتم نقل أي طالب!")

        btn_frame = tk.Frame(move_window, bg="white")
        btn_frame.pack(pady=8)

        tk.Button(btn_frame, text="نقل المختارين", font=("Arial", 11),
                  command=move_selected_students, bg="#27ae60", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=3)

        tk.Button(btn_frame, text="إلغاء", font=("Arial", 11),
                  command=move_window.destroy, bg="#e74c3c", fg="white",
                  relief="raised", bd=1).pack(side="left", padx=3)


def run_app():
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()

