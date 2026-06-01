import sys
import os

with open('oop-new.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_init = """            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT, subject TEXT, date TEXT,
                started_at TEXT, closed_at TEXT, total_registered INTEGER DEFAULT 0
            )""")"""
new_init = """            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT, subject TEXT, date TEXT,
                started_at TEXT, closed_at TEXT, total_registered INTEGER DEFAULT 0
            )""")
            cursor.execute(\"\"\"
            CREATE TABLE IF NOT EXISTS students (
                id TEXT PRIMARY KEY, name TEXT, registered_at TEXT
            )\"\"\")"""
content = content.replace(old_init, new_init)

old_sm_init = """class StudentManager:
    def __init__(self, folder_path="students"):
        self.folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), folder_path)"""
new_sm_init = """class StudentManager:
    def __init__(self, db, folder_path="students"):
        self.db = db
        self.folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), folder_path)"""
content = content.replace(old_sm_init, new_sm_init)

old_add = """        with open(filepath, "wb") as f: f.write(image_bytes)
        return True, f"✅ Student {name} added successfully as {filename}!\""""
new_add = """        with open(filepath, "wb") as f: f.write(image_bytes)
        
        try:
            from datetime import datetime
            conn = self.db.connect()
            conn.cursor().execute("INSERT OR REPLACE INTO students (id, name, registered_at) VALUES (?, ?, ?)", 
                                  (student_id, name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
        except Exception as e: print('DB Error', e)
        
        return True, f"✅ Student {name} added successfully as {filename}!\""""
content = content.replace(old_add, new_add)

old_del = """    def delete_student(self, filename):
        filepath = os.path.join(self.folder_path, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            return True
        return False"""
new_del = """    def delete_student(self, filename, student_id=None):
        filepath = os.path.join(self.folder_path, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            if student_id:
                try:
                    conn = self.db.connect()
                    remaining = [f for f in os.listdir(self.folder_path) if f.startswith(f"{student_id}-")]
                    if not remaining:
                        conn.cursor().execute("DELETE FROM students WHERE id=?", (student_id,))
                        conn.commit()
                    conn.close()
                except: pass
            return True
        return False"""
content = content.replace(old_del, new_del)

old_app_init = """class AttendanceApp:
    def __init__(self):
        self.db = DatabaseManager()
        self.student_manager = StudentManager()
        self.recognizer = FaceRecognitionSystem()"""
new_app_init = """class AttendanceApp:
    def __init__(self):
        self.db = DatabaseManager()
        self.student_manager = StudentManager(self.db)
        
        try:
            from datetime import datetime
            conn = self.db.connect()
            cursor = conn.cursor()
            for s in self.student_manager.get_unique_students():
                cursor.execute("INSERT OR IGNORE INTO students (id, name, registered_at) VALUES (?, ?, ?)", 
                               (s['id'], s['name'], datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
        except: pass
        
        self.recognizer = FaceRecognitionSystem()"""
content = content.replace(old_app_init, new_app_init)

old_del_call = """                    if st.button("❌ Delete", key=f"del_{student['id']}", width='stretch'):
                        for s in students:
                            if s['id'] == student['id']: self.student_manager.delete_student(s['filename'])"""
new_del_call = """                    if st.button("❌ Delete", key=f"del_{student['id']}", width='stretch'):
                        for s in students:
                            if s['id'] == student['id']: self.student_manager.delete_student(s['filename'], student['id'])"""
content = content.replace(old_del_call, new_del_call)

with open('oop-new.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Student DB patch applied!")
