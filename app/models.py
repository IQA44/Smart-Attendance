from dataclasses import dataclass
from typing import List, Dict

@dataclass
class Student:
    name: str
    stage: str
    department: str



@dataclass
class AttendanceEntry:
    type: str        
    time: str      


@dataclass
class DailyAttendance:
    date: str
    records: List[AttendanceEntry]



@dataclass
class StudentAttendance:
    student_key: str
    days: Dict[str, List[AttendanceEntry]]

