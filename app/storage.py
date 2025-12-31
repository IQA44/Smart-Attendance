import json
import os
from tkinter import messagebox

from app.constants import (
    PROGRAM_STORAGE,
    STUDENTS_FOLDER,
    ATTENDANCE_FOLDER,
    CARDS_FILE,
    STAGES_FILE,
    RECORDS_FOLDER
)


def create_folders():
    folders = [
        PROGRAM_STORAGE,
        STUDENTS_FOLDER,
        ATTENDANCE_FOLDER,
        RECORDS_FOLDER
    ]
    for folder in folders:
        if not os.path.exists(folder):
            os.makedirs(folder)


def load_data(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    else:
        
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(default, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        return default


def save_data(filepath, data):
    try:
        folder = os.path.dirname(filepath)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل في حفظ البيانات:\n{str(e)}")
        return False


def get_stage_department_path(stage, department):
    return f"{stage}_{department}"


def load_stage_department_students(stage, department):
    path = get_stage_department_path(stage, department)
    class_file = os.path.join(STUDENTS_FOLDER, f"{path}.json")
    return load_data(class_file, [])

def save_stage_department_students(stage, department, students):
    path = get_stage_department_path(stage, department)
    class_file = os.path.join(STUDENTS_FOLDER, f"{path}.json")
    return save_data(class_file, students)


def load_stage_department_attendance(stage, department):
    path = get_stage_department_path(stage, department)
    attendance_file = os.path.join(ATTENDANCE_FOLDER, f"{path}.json")
    return load_data(attendance_file, {})

def save_stage_department_attendance(stage, department, attendance):
    path = get_stage_department_path(stage, department)
    attendance_file = os.path.join(ATTENDANCE_FOLDER, f"{path}.json")
    return save_data(attendance_file, attendance)


def initialize_storage():
    create_folders()

    
    load_data(CARDS_FILE, {})

       
    load_data(STAGES_FILE, {})

