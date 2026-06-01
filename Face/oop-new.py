
import streamlit as st
st.set_page_config(page_title="Biometric Attendance System", page_icon="👤", layout="wide", initial_sidebar_state="expanded")

import cv2
import face_recognition
import numpy as np
import os
import sqlite3
from datetime import datetime
import pandas as pd
import sys
import time
import threading
import hashlib
import pickle
import io
import base64
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# --- Fix for Windows Asyncio ConnectionResetError ---
if sys.platform.startswith("win"):
    try:
        from functools import wraps
        from asyncio.proactor_events import _ProactorBasePipeTransport
        def silence_winerror_10054(func):
            @wraps(func)
            def wrapper(self, *args, **kwargs):
                try: return func(self, *args, **kwargs)
                except ConnectionResetError as e:
                    if e.winerror in (10054, 10053, 121): pass
                    else: raise
            return wrapper
        _ProactorBasePipeTransport._call_connection_lost = silence_winerror_10054(_ProactorBasePipeTransport._call_connection_lost)
    except Exception: pass
# ----------------------------------------------------

def safe_rerun():
    if hasattr(st, 'rerun'): st.rerun()
    else: st.experimental_rerun()

def inject_custom_css():
    st.markdown("""
    <style>
    :root {
        --primary: #00C9FF;
        --secondary: #92FE9D;
        --danger: #FF6B6B;
        --warning: #FFD93D;
        --bg-dark: #0f2027;
        --card-bg: rgba(255, 255, 255, 0.07);
    }
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
    
    .stApp {
        background: linear-gradient(135deg, var(--bg-dark), #203a43, #2c5364);
        color: #ffffff;
    }
    .glass-card {
        background: var(--card-bg);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 20px;
        transition: all 0.2s ease;
    }
    .glass-card:hover { transform: translateY(-3px); box-shadow: 0 10px 20px rgba(0,0,0,0.3); }
    .header-banner {
        background: linear-gradient(90deg, var(--primary), var(--secondary));
        border-radius: 16px;
        padding: 20px 30px;
        color: #1a1a1a;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0, 201, 255, 0.2);
    }
    .pulsing-dot {
        height: 15px; width: 15px; border-radius: 50%;
        display: inline-block; margin-right: 10px;
    }
    .dot-green { background-color: #00FF00; box-shadow: 0 0 10px #00FF00; animation: pulse 1.5s infinite; }
    .dot-red { background-color: #FF0000; box-shadow: 0 0 10px #FF0000; }
    
    @keyframes pulse {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(0, 255, 0, 0.7); }
        70% { transform: scale(1); box-shadow: 0 0 0 10px rgba(0, 255, 0, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(0, 255, 0, 0); }
    }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background: var(--card-bg); border-radius: 10px 10px 0 0; }
    .stTabs [aria-selected="true"] { background: linear-gradient(90deg, var(--primary), var(--secondary)); color: #1a1a1a !important; font-weight: 800; }
    </style>
    """, unsafe_allow_html=True)

def get_image_base64(path):
    try:
        with open(path, "rb") as img_file: return base64.b64encode(img_file.read()).decode('utf-8')
    except: return ""

# =========================
# DATABASE CLASS
# =========================
class DatabaseManager:
    def __init__(self, db_name="attendance.db"):
        self.db_name = db_name
        self.init_db()

    def connect(self): return sqlite3.connect(self.db_name)

    def init_db(self):
        conn = None
        try:
            conn = self.connect()
            cursor = conn.cursor()
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS attendance (
                id TEXT, name TEXT, date TEXT, time TEXT, subject TEXT
            )""")
            for col in ["subject TEXT DEFAULT 'General'", "time_leave TEXT", "status TEXT"]:
                try: cursor.execute(f"ALTER TABLE attendance ADD COLUMN {col}")
                except: pass
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS subjects (
                name TEXT PRIMARY KEY, start_time TEXT, end_time TEXT
            )""")
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT, subject TEXT, date TEXT,
                started_at TEXT, closed_at TEXT, total_registered INTEGER DEFAULT 0
            )""")
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id TEXT PRIMARY KEY, name TEXT, registered_at TEXT
            )""")
            conn.commit()
        except sqlite3.Error as e: st.error(f"❌ DB init error: {e}")
        finally:
            if conn: conn.close()

    def clear_db(self):
        conn = None
        try:
            conn = self.connect()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM attendance")
            cursor.execute("DELETE FROM sessions")
            conn.commit()
        except sqlite3.Error as e: st.error(f"❌ Error clearing DB: {e}")
        finally:
            if conn: conn.close()

    def save(self, student_id, name, subject):
        conn = None
        try:
            conn = self.connect()
            cursor = conn.cursor()
            now = datetime.now()
            date_now, time_now = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")

            cursor.execute("SELECT * FROM attendance WHERE id=? AND date=? AND subject=?", (student_id, date_now, subject))
            if cursor.fetchone():
                cursor.execute("UPDATE attendance SET time_leave=? WHERE id=? AND date=? AND subject=?", (time_now, student_id, date_now, subject))
                conn.commit()
                return "updated"

            status = 'present'
            if subject != 'General':
                cursor.execute("SELECT start_time FROM subjects WHERE name=?", (subject,))
                subj_row = cursor.fetchone()
                if subj_row:
                    try:
                        start_dt = datetime.strptime(f"{date_now} {subj_row[0]}:00", "%Y-%m-%d %H:%M:%S")
                        if (now - start_dt).total_seconds() / 60.0 > 10: status = 'late'
                    except ValueError: pass

            cursor.execute(
                "INSERT INTO attendance (id, name, date, time, subject, time_leave, status) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (student_id, name, date_now, time_now, subject, time_now, status)
            )
            conn.commit()
            return "inserted"
        except sqlite3.Error as e:
            st.error(f"❌ Error saving attendance: {e}")
            return False
        finally:
            if conn: conn.close()

    def close_session(self, subject, all_students):
        conn = None
        try:
            conn = self.connect()
            cursor = conn.cursor()
            now = datetime.now()
            date_now, time_now = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")
            
            cursor.execute("SELECT id, status FROM attendance WHERE date=? AND subject=?", (date_now, subject))
            attended_ids = {row[0]: row[1] for row in cursor.fetchall()}
            
            absent_count = 0
            for student in all_students:
                if student['id'] not in attended_ids:
                    cursor.execute(
                        "INSERT INTO attendance (id, name, date, time, subject, time_leave, status) VALUES (?, ?, ?, ?, ?, NULL, 'absent')",
                        (student['id'], student['name'], date_now, time_now, subject)
                    )
                    absent_count += 1
                    
            cursor.execute(
                "INSERT INTO sessions (subject, date, started_at, closed_at, total_registered) VALUES (?, ?, ?, ?, ?)",
                (subject, date_now, time_now, time_now, len(all_students))
            )
            conn.commit()
            
            return {
                "present": sum(1 for s in attended_ids.values() if s == 'present'),
                "late": sum(1 for s in attended_ids.values() if s == 'late'),
                "absent": absent_count
            }
        except sqlite3.Error as e: return None
        finally:
            if conn: conn.close()

    def load(self):
        conn = None
        try:
            conn = self.connect()
            return pd.read_sql_query("SELECT * FROM attendance ORDER BY date DESC, time DESC", conn)
        except sqlite3.Error: return pd.DataFrame()
        finally:
            if conn: conn.close()

    def get_excel_data(self, filtered_df=None):
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        def format_sheet(ws, is_summary=False, title_row=None, header_fill="1F3864", alt_color="DCE6F1"):
            from openpyxl.utils import get_column_letter
            for row in ws.iter_rows():
                for cell in row: cell.font = Font(name='Calibri', size=11)
            
            start_row = 1
            if title_row:
                ws.insert_rows(1)
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
                title_cell = ws.cell(row=1, column=1, value=title_row)
                title_cell.font = Font(name='Calibri', size=14, bold=True)
                title_cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                title_cell.alignment = Alignment(horizontal='center', vertical='center')
                start_row = 2
                
            header_font = Font(name='Calibri', size=12, bold=True, color="FFFFFF")
            h_fill = PatternFill(start_color=header_fill, end_color=header_fill, fill_type="solid")
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=start_row, column=col_idx)
                cell.font = header_font
                cell.fill = h_fill
                
            alt_fill = PatternFill(start_color=alt_color, end_color=alt_color, fill_type="solid")
            for row_idx in range(start_row + 1, ws.max_row + 1):
                if row_idx % 2 == (0 if start_row==1 else 1):
                    for col_idx in range(1, ws.max_column + 1):
                        ws.cell(row=row_idx, column=col_idx).fill = alt_fill
                        
            for col_idx, col in enumerate(ws.columns, start=1):
                max_length = 0
                column = get_column_letter(col_idx)
                for cell in col:
                    try: max_length = max(max_length, len(str(cell.value)))
                    except: pass
                ws.column_dimensions[column].width = (max_length + 2)
                
            ws.freeze_panes = f"A{start_row + 1}"
            ws.auto_filter.ref = f"A{start_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
            
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ws.oddFooter.center.text = f"Generated by Biometric Attendance System — {ts}"
            ws.evenFooter.center.text = f"Generated by Biometric Attendance System — {ts}"

        if filtered_df is not None:
            ws1 = wb.create_sheet("Filtered Records")
            ws1.append(filtered_df.columns.tolist())
            for r in filtered_df.values.tolist(): ws1.append(r)
            format_sheet(ws1)
            
            today = datetime.now().strftime("%Y-%m-%d")
            today_df = filtered_df[filtered_df['date'] == today]
            ws2 = wb.create_sheet("Today's Attendance")
            ws2.append(today_df.columns.tolist())
            for r in today_df.values.tolist(): ws2.append(r)
            format_sheet(ws2, title_row=f"Attendance Report — {today}")
        else:
            df = self.load()
            ws1 = wb.create_sheet("All Records")
            if not df.empty:
                ws1.append(df.columns.tolist())
                for r in df.values.tolist(): ws1.append(r)
            format_sheet(ws1)
            
            today = datetime.now().strftime("%Y-%m-%d")
            today_df = df[df['date'] == today] if not df.empty else pd.DataFrame(columns=df.columns if not df.empty else [])
            ws2 = wb.create_sheet("Today's Attendance")
            ws2.append(today_df.columns.tolist())
            for r in today_df.values.tolist(): ws2.append(r)
            format_sheet(ws2, title_row=f"Attendance Report — {today}")
            
            conn = self.connect()
            sessions_df = pd.read_sql_query("SELECT subject, COUNT(*) as total_sessions FROM sessions GROUP BY subject", conn)
            att_df = pd.read_sql_query("SELECT id, name, subject, COUNT(*) as attended FROM attendance WHERE status IN ('present', 'late') GROUP BY id, name, subject", conn)
            
            ws3 = wb.create_sheet("Attendance Summary")
            ws3.append(["Student ID", "Student Name", "Subject", "Sessions Held", "Days Attended", "Attendance %"])
            if not sessions_df.empty and not att_df.empty:
                merged = pd.merge(att_df, sessions_df, on='subject', how='left')
                merged['Attendance %'] = (merged['attended'] / merged['total_sessions'])
                for _, row in merged.iterrows():
                    ws3.append([row['id'], row['name'], row['subject'], row['total_sessions'], row['attended'], row['Attendance %']])
                    
                for row_idx in range(2, ws3.max_row + 1):
                    cell = ws3.cell(row=row_idx, column=6)
                    cell.number_format = '0.0%'
                    val = cell.value
                    if isinstance(val, (int, float)):
                        if val < 0.75:
                            cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                            cell.font = Font(color="FF0000", bold=True, name='Calibri', size=11)
                        elif val <= 0.90:
                            cell.fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
                        else:
                            cell.fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
            format_sheet(ws3, header_fill="375623")
            
            absent_df = pd.read_sql_query("SELECT id, name, subject, date FROM attendance WHERE status='absent' ORDER BY date DESC, subject ASC", conn)
            if not absent_df.empty:
                ws4 = wb.create_sheet("Absent Students")
                ws4.append(absent_df.columns.tolist())
                for r in absent_df.values.tolist(): ws4.append(r)
                format_sheet(ws4, header_fill="C00000")
            conn.close()
                
        wb.save(output)
        return output.getvalue()

    def add_subject(self, name, start_time, end_time):
        if start_time == end_time:
            st.error("❌ Start time and end time cannot be the same.")
            return False
        conn = None
        try:
            conn = self.connect()
            conn.cursor().execute("REPLACE INTO subjects (name, start_time, end_time) VALUES (?, ?, ?)", 
                           (name, start_time.strftime("%H:%M"), end_time.strftime("%H:%M")))
            conn.commit()
            return True
        except sqlite3.Error as e: return False
        finally:
            if conn: conn.close()

    def get_subjects(self):
        conn = None
        try:
            conn = self.connect()
            return pd.read_sql_query("SELECT * FROM subjects ORDER BY start_time", conn)
        except sqlite3.Error: return pd.DataFrame()
        finally:
            if conn: conn.close()
                
    def delete_subject(self, name):
        conn = None
        try:
            conn = self.connect()
            conn.cursor().execute("DELETE FROM subjects WHERE name=?", (name,))
            conn.commit()
            return True
        except sqlite3.Error: return False
        finally:
            if conn: conn.close()


# =========================
# STUDENT MANAGER CLASS
# =========================
class StudentManager:
    def __init__(self, db, folder_path="students"):
        self.db = db
        self.folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), folder_path)
        if not os.path.exists(self.folder_path): os.makedirs(self.folder_path)

    def add_student(self, name, student_id, image_bytes, force_overwrite=False):
        existing = []
        if os.path.exists(self.folder_path):
            for f in os.listdir(self.folder_path):
                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')) and f.startswith(f"{student_id}-"):
                    existing.append(f)
                    
        if existing and not force_overwrite:
            return False, f"⚠️ Student ID {student_id} is already registered."
        if len(existing) >= 3:
            return False, "Maximum 3 photos per student reached."
            
        count = len(existing) + 1
        filename = f"{student_id}-{name}-{count}.jpg"
        filepath = os.path.join(self.folder_path, filename)
        with open(filepath, "wb") as f: f.write(image_bytes)
        
        try:
            conn = self.db.connect()
            conn.cursor().execute("INSERT OR REPLACE INTO students (id, name, registered_at) VALUES (?, ?, ?)", 
                                  (student_id, name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
        except: pass
        
        return True, f"✅ Student {name} added successfully as {filename}!"

    def get_all_students(self):
        students = []
        if not os.path.exists(self.folder_path): return students
        for img_name in os.listdir(self.folder_path):
            if img_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                raw_name = os.path.splitext(img_name)[0]
                parts = raw_name.split("-")
                if len(parts) >= 3 and parts[-1].isdigit():
                    sid, sname = parts[0], "-".join(parts[1:-1])
                elif len(parts) >= 2:
                    sid, sname = parts[0], "-".join(parts[1:])
                else: sid, sname = raw_name, raw_name
                students.append({"id": sid, "name": sname, "filename": img_name})
        return students

    def get_unique_students(self):
        students_dict = {}
        for s in self.get_all_students(): students_dict[s['id']] = s
        return list(students_dict.values())

    def delete_student(self, filename, student_id=None):
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
        return False


# =========================
# FACE RECOGNITION CLASS
# =========================
class FaceRecognitionSystem:
    def __init__(self):
        self.known_encodings, self.names = self.load_faces()

    def load_faces(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "students")
        if not os.path.exists(path): os.makedirs(path)

        image_extensions = (".jpg", ".jpeg", ".png", ".bmp")
        files_info = []
        for img_name in os.listdir(path):
            if img_name.lower().endswith(image_extensions):
                files_info.append((img_name, os.path.getmtime(os.path.join(path, img_name))))
                
        files_info.sort()
        cache_key_str = str(files_info).encode('utf-8')
        current_hash = hashlib.md5(cache_key_str).hexdigest()
        
        cache_file = os.path.join(path, ".encoding_cache.pkl")
        if os.path.exists(cache_file):
            try:
                with open(cache_file, "rb") as f:
                    cache_data = pickle.load(f)
                    if cache_data.get('hash') == current_hash:
                        return cache_data['encodings'], cache_data['names']
            except: pass

        encodings, names = [], []
        for img_name, _ in files_info:
            img = cv2.imread(os.path.join(path, img_name))
            if img is None: continue
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            # Use 10 jitters for a much more robust encoding
            enc = face_recognition.face_encodings(img, num_jitters=10, model='large')
            if len(enc) > 0:
                encodings.append(enc[0])
                names.append(os.path.splitext(img_name)[0])

        with open(cache_file, "wb") as f:
            pickle.dump({'hash': current_hash, 'encodings': encodings, 'names': names}, f)
        return encodings, names

    def reload_faces(self):
        self.known_encodings, self.names = self.load_faces()

    def recognize_with_dist(self, face_encoding, threshold=0.5):
        if not self.known_encodings: return "Unknown", None
        distances = face_recognition.face_distance(self.known_encodings, face_encoding)
        best_match = np.argmin(distances)
        dist = distances[best_match]
        if dist <= threshold: return self.names[best_match], dist
        return "Unknown", dist

    def parse_name_id(self, raw_name):
        parts = raw_name.split("-")
        if len(parts) >= 3 and parts[-1].isdigit():
            return parts[0].strip(), "-".join(parts[1:-1]).strip()
        elif len(parts) >= 2:
            return parts[0].strip(), "-".join(parts[1:]).strip()
        return raw_name.strip(), raw_name.strip()


# =========================
# CAMERA CLASS
# =========================
class CameraSystem:
    def __init__(self, recognizer, db, app):
        self.recognizer = recognizer
        self.db = db
        self.app = app

    def run(self):
        idx = st.session_state.get("camera_index", 0)
        cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
        if not cap.isOpened():
            st.error("❌ Cannot open camera. Please check your device.")
            return

        session_recorded = {}
        process_this_frame = True
        threshold = st.session_state.get("face_threshold", 0.5)

        try:
            while st.session_state.camera_on:
                ret, frame = cap.read()
                if not ret: break

                small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
                rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

                if process_this_frame:
                    face_locations = face_recognition.face_locations(rgb_small_frame)
                    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
                    face_names, face_distances = [], []
                    for enc in face_encodings:
                        name, dist = self.recognizer.recognize_with_dist(enc, threshold)
                        face_names.append(name)
                        face_distances.append(dist)
                        
                process_this_frame = not process_this_frame
                current_time_str = datetime.now().strftime('%H:%M:%S')

                for (top, right, bottom, left), name, dist in zip(face_locations, face_names, face_distances):
                    top *= 4; right *= 4; bottom *= 4; left *= 4
                    color = (0, 255, 0) if name != "Unknown" else (0, 0, 255)
                    cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                    cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)

                    if name != "Unknown" and dist is not None:
                        conf = max(0.0, 1.0 - dist)
                        fill_w = int((right - left) * conf)
                        c_bar = (0,255,0) if conf > 0.6 else (0,255,255) if conf > 0.4 else (0,0,255)
                        cv2.rectangle(frame, (left, bottom + 5), (left + fill_w, bottom + 15), c_bar, cv2.FILLED)
                        cv2.rectangle(frame, (left, bottom + 5), (right, bottom + 15), (255,255,255), 1)

                    subject = st.session_state.get('current_subject', 'General')
                    if "live_feed" not in st.session_state: st.session_state.live_feed = []

                    if name == "Unknown":
                        if time.time() - st.session_state.get('last_unk', 0) > 5:
                            st.session_state.live_feed.append({"name": "Unknown", "time": current_time_str})
                            st.session_state.last_unk = time.time()
                    else:
                        record_key = f"{name}_{subject}"
                        if st.session_state.get('auto_scan', True) or st.session_state.capture:
                            student_id, student_name = self.recognizer.parse_name_id(name)
                            if time.time() - session_recorded.get(record_key, 0) > 10:
                                status = self.db.save(student_id, student_name, subject)
                                if status == "inserted":
                                    st.session_state.live_feed.append({
                                        "name": student_name, "id": student_id, 
                                        "subject": subject, "time": current_time_str
                                    })
                                session_recorded[record_key] = time.time()
                
                st.session_state.capture = False
                
                if hasattr(self.app, 'feed_placeholder'):
                    html = ""
                    for item in reversed(st.session_state.live_feed[-10:]):
                        if item['name'] == 'Unknown':
                            html += f'<div style="background: rgba(255,0,0,0.2); border-left: 4px solid #FF6B6B; padding: 10px; margin-bottom: 10px; border-radius: 4px;">⚠️ Unknown Face Detected<br><small>{item["time"]}</small></div>'
                        else:
                            html += f'<div style="background: rgba(0,255,0,0.1); border-left: 4px solid #00C9FF; padding: 10px; margin-bottom: 10px; border-radius: 4px;">✅ <b>{item["name"]}</b> ({item["id"]})<br><small>{item["subject"]} • {item["time"]}</small></div>'
                    self.app.feed_placeholder.markdown(html, unsafe_allow_html=True)

                if hasattr(self.app, 'frame_placeholder'):
                    self.app.frame_placeholder.image(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB), width='stretch')

        finally:
            threading.Thread(target=cap.release).start()


# =========================
# MAIN APP CLASS
# =========================
class AttendanceApp:
    def __init__(self):
        self.db = DatabaseManager()
        self.student_manager = StudentManager(self.db)
        
        # Sync DB with folder
        try:
            conn = self.db.connect()
            cursor = conn.cursor()
            for s in self.student_manager.get_unique_students():
                cursor.execute("INSERT OR IGNORE INTO students (id, name, registered_at) VALUES (?, ?, ?)", 
                               (s['id'], s['name'], datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
        except: pass
        
        self.recognizer = FaceRecognitionSystem()
        self.camera = CameraSystem(self.recognizer, self.db, self)
        self._init_session_state()

    def _init_session_state(self):
        defaults = { "camera_on": False, "capture": False, "confirm_clear": False, "live_feed": [] }
        for k, v in defaults.items():
            if k not in st.session_state: st.session_state[k] = v

    def get_current_subject(self):
        df = self.db.get_subjects()
        if df.empty: return "General"
        current_time = datetime.now().strftime("%H:%M")
        for _, row in df.iterrows():
            if row['start_time'] <= current_time <= row['end_time']: return row['name']
        return "General"

    def sidebar(self):
        hour = datetime.now().hour
        greeting = "Good Morning" if hour < 12 else "Good Afternoon" if hour < 18 else "Good Evening"
        
        st.sidebar.markdown(f"""
        <div class="glass-card" style="text-align: center; margin-bottom: 20px;">
            <h2 style="margin: 0; font-size: 1.5rem; color: #fff;">👋 {greeting}</h2>
            <p style="margin: 5px 0 0 0; opacity: 0.8; color: #fff;">{datetime.now().strftime('%A, %b %d, %Y')}</p>
        </div>
        """, unsafe_allow_html=True)
        
        df = self.db.load()
        today = datetime.now().strftime("%Y-%m-%d")
        today_df = df[df['date'] == today] if not df.empty else pd.DataFrame()
        
        try: sessions_count = pd.read_sql_query("SELECT COUNT(*) FROM sessions", self.db.connect()).iloc[0,0]
        except: sessions_count = 0
            
        st.sidebar.markdown('<div class="glass-card">', unsafe_allow_html=True)
        c1, c2 = st.sidebar.columns(2)
        c1.metric("✅ Today", len(today_df))
        c2.metric("👥 Students", len(self.student_manager.get_unique_students()))
        c3, c4 = st.sidebar.columns(2)
        c3.metric("📅 Sessions", sessions_count)
        c4.metric("🗂️ Records", len(df))
        st.sidebar.markdown('</div>', unsafe_allow_html=True)

        st.sidebar.markdown('<div class="glass-card"><h4>⚙️ Camera Controls</h4>', unsafe_allow_html=True)
        st.sidebar.selectbox("Camera Index", [0, 1, 2], key="camera_index")
        st.sidebar.slider("Recognition Sensitivity", 0.3, 0.7, key="face_threshold")
        
        if st.sidebar.button("▶️ START CAMERA", width='stretch'):
            st.session_state.camera_on = True
            safe_rerun()
        if st.sidebar.button("⏹️ STOP CAMERA", width='stretch'):
            st.session_state.camera_on = False
            safe_rerun()
        if st.sidebar.button("📸 MANUAL SCAN", width='stretch'):
            if st.session_state.camera_on: st.session_state.capture = True
            else: st.sidebar.warning("⚠️ Please start camera first.")
        st.sidebar.markdown('</div>', unsafe_allow_html=True)

    def render_header(self):
        status_class = "dot-green" if st.session_state.camera_on else "dot-red"
        status_text = "Camera ON" if st.session_state.camera_on else "Camera OFF"
        st.markdown(f"""
        <div class="header-banner">
            <div>
                <h1 style="margin: 0; color: #1a1a1a;">👤 Biometric Attendance System</h1>
                <h4 style="margin: 0; opacity: 0.8; color: #1a1a1a;">{datetime.now().strftime('%Y-%m-%d %H:%M')}</h4>
            </div>
            <div style="font-weight: bold; font-size: 1.2rem; display: flex; align-items: center;">
                <span class="pulsing-dot {status_class}"></span>{status_text}
            </div>
        </div>
        """, unsafe_allow_html=True)

    def camera_section(self):
        colL, colR = st.columns([2, 1])
        
        subjects_df = self.db.get_subjects()
        subject_list = ["General"] + (subjects_df['name'].tolist() if not subjects_df.empty else [])
        auto_subject = self.get_current_subject()
        default_idx = subject_list.index(auto_subject) if auto_subject in subject_list else 0
        
        with colL:
            st.markdown('<div class="glass-card" style="padding: 10px;">', unsafe_allow_html=True)
            st.markdown('<div class="video-container">', unsafe_allow_html=True)
            self.frame_placeholder = st.empty()
            st.markdown('</div>', unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                subject = st.selectbox("Select Subject", subject_list, index=default_idx)
                st.session_state.current_subject = subject
            with c2:
                st.write(" ")
                st.session_state.auto_scan = st.toggle("🤖 Auto-Scan", value=True)
            with c3:
                st.write(" ")
                if st.session_state.camera_on:
                    if st.button("🔒 Close Session", width='stretch'):
                        res = self.db.close_session(st.session_state.current_subject, self.student_manager.get_unique_students())
                        if res: st.toast(f"Session closed — Present: {res['present']}, Late: {res['late']}, Absent: {res['absent']}")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with colR:
            st.markdown('<div class="glass-card"><h3 style="margin-top:0;">📋 Who\'s Here</h3>', unsafe_allow_html=True)
            self.feed_placeholder = st.empty()
            st.markdown('</div>', unsafe_allow_html=True)
            
        if st.session_state.camera_on:
            self.camera.run()
        else:
            self.frame_placeholder.info("📷 Camera is OFF — Press START in the sidebar to begin")

    def student_management_section(self):
        st.markdown('<div class="glass-card"><h3 style="margin-top:0;">➕ Enroll New Student (3 Photos)</h3>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1: new_id = st.text_input("Student ID", placeholder="e.g. 001")
        with c2: new_name = st.text_input("Student Name", placeholder="e.g. Ahmed")
        with c3: upload_method = st.radio("Photo Method", ["Auto-Capture (Webcam)", "Upload Images"], horizontal=True)
        
        force_overwrite = st.checkbox("Force overwrite / Add extra photo")

        if upload_method == "Upload Images":
            uploaded_files = st.file_uploader("Upload up to 3 Student Photos", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
            if st.button("Save Uploaded Photos", width='stretch'):
                if new_id and new_name and uploaded_files:
                    if len(uploaded_files) > 3:
                        st.error("Please upload maximum 3 photos.")
                    else:
                        for f in uploaded_files:
                            self.student_manager.add_student(new_name, new_id, f.getvalue(), force_overwrite)
                        self.recognizer.reload_faces()
                        st.success(f"✅ {len(uploaded_files)} photos saved for {new_name}!")
                        time.sleep(1)
                        safe_rerun()
                else: st.error("❌ Please provide ID, Name and Image(s).")
        else:
            if st.button("📸 Start Auto-Capture (3 Photos)", width='stretch'):
                if new_id and new_name:
                    cap = cv2.VideoCapture(st.session_state.get("camera_index", 0), cv2.CAP_DSHOW)
                    if not cap.isOpened(): st.error("❌ Cannot open camera.")
                    else:
                        placeholder = st.empty()
                        photos_taken = 0
                        st.info("Looking for 1 face... slightly turn your head for each photo to get better angles!")
                        
                        while photos_taken < 3:
                            # Countdown sequence
                            directions = ["Straight", "Slightly Left", "Slightly Right"]
                            for countdown in [3, 2, 1]:
                                start_time = time.time()
                                while time.time() - start_time < 1.0:
                                    ret, frame = cap.read()
                                    if not ret: break
                                    display = frame.copy()
                                    msg = f"Photo {photos_taken+1}/3: Look {directions[photos_taken]} in {countdown}..."
                                    cv2.putText(display, msg, (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 255), 2)
                                    placeholder.image(cv2.cvtColor(display, cv2.COLOR_BGR2RGB))
                            
                            # Capture sequence
                            captured_this_step = False
                            while not captured_this_step:
                                ret, frame = cap.read()
                                if not ret: break
                                
                                small = cv2.resize(frame, (0,0), fx=0.25, fy=0.25)
                                rgb = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
                                locs = face_recognition.face_locations(rgb)
                                
                                display = frame.copy()
                                if len(locs) == 1:
                                    top, right, bottom, left = [v * 4 for v in locs[0]]
                                    cv2.rectangle(display, (left, top), (right, bottom), (0, 255, 0), 2)
                                    cv2.putText(display, "✅ SNAP!", (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 1.2, (0, 255, 0), 3)
                                    placeholder.image(cv2.cvtColor(display, cv2.COLOR_BGR2RGB))
                                    
                                    _, buffer = cv2.imencode('.jpg', frame)
                                    self.student_manager.add_student(new_name, new_id, buffer.tobytes(), True)
                                    photos_taken += 1
                                    captured_this_step = True
                                    time.sleep(1.0) # Hold the confirmation screen
                                else:
                                    msg = "Looking for exactly 1 face..." if len(locs) == 0 else "Multiple faces detected!"
                                    cv2.putText(display, msg, (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)
                                    placeholder.image(cv2.cvtColor(display, cv2.COLOR_BGR2RGB))
                                    time.sleep(0.1)
                                
                        cap.release()
                        placeholder.empty()
                        self.recognizer.reload_faces()
                        st.success(f"✅ Successfully captured 3 photos for {new_name}!")
                        time.sleep(1)
                        safe_rerun()
                else: st.error("❌ Please provide ID and Name first.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.subheader("📋 Registered Students")
        students = self.student_manager.get_all_students()
        unique = self.student_manager.get_unique_students()
        photo_counts = {s['id']: sum(1 for x in students if x['id'] == s['id']) for s in unique}
        
        if unique:
            cols = st.columns(2)
            for i, student in enumerate(unique):
                with cols[i % 2]:
                    img_path = os.path.join(self.student_manager.folder_path, student['filename'])
                    img_b64 = get_image_base64(img_path)
                    img_html = f'<img src="data:image/jpeg;base64,{img_b64}" style="width:80px;height:80px;border-radius:50%;object-fit:cover;">' if img_b64 else '👤'
                    
                    st.markdown(f"""
                    <div class="glass-card" style="display:flex; align-items:center; gap: 20px; margin-bottom: 10px; padding: 15px;">
                        {img_html}
                        <div style="flex-grow: 1;">
                            <h4 style="margin:0;">{student['name']}</h4>
                            <p style="opacity:0.7; margin:0;">ID: {student['id']}</p>
                            <p style="margin:5px 0;">📸 {photo_counts[student['id']]}/3 photos</p>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button("❌ Delete", key=f"del_{student['id']}", width='stretch'):
                        for s in students:
                            if s['id'] == student['id']: self.student_manager.delete_student(s['filename'], student['id'])
                        self.recognizer.reload_faces()
                        safe_rerun()
        else: st.info("No students registered yet.")

    def subjects_management_section(self):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="glass-card"><h3 style="margin-top:0;">➕ Add New Subject</h3>', unsafe_allow_html=True)
            new_subj_name = st.text_input("Subject Name", placeholder="e.g. Physics 101")
            t_col1, t_col2 = st.columns(2)
            with t_col1: start_time = st.time_input("Start Time")
            with t_col2: end_time = st.time_input("End Time")
            if st.button("Save Subject", width='stretch'):
                if new_subj_name:
                    if self.db.add_subject(new_subj_name, start_time, end_time):
                        st.success(f"✅ Subject '{new_subj_name}' added!")
                        safe_rerun()
                else: st.error("Please provide a subject name.")
            st.markdown('</div>', unsafe_allow_html=True)
                    
        with col2:
            st.subheader("📚 Configured Subjects")
            df = self.db.get_subjects()
            if df.empty: st.info("No subjects configured yet.")
            else:
                for _, row in df.iterrows():
                    now = datetime.now().strftime("%H:%M")
                    b_color = "var(--secondary)" if row['start_time'] <= now <= row['end_time'] else "var(--primary)" if now < row['start_time'] else "grey"
                    st.markdown(f"""
                    <div class="glass-card" style="border-left: 5px solid {b_color}; padding: 15px; display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <h3 style="margin:0;">{row['name']}</h3>
                            <p style="margin:5px 0; opacity:0.8;">🕒 {row['start_time']} ➔ {row['end_time']}</p>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button("❌", key=f"del_subj_{row['name']}"):
                        if self.db.delete_subject(row['name']): safe_rerun()

    def records_section(self):
        st.subheader("🗂️ Attendance Database")
        df = self.db.load()
        if df.empty:
            st.info("No attendance records found.")
            return

        col1, col2 = st.columns(2)
        with col1: date_filter = st.date_input("Filter by Date", value=None)
        subjects = ["All"] + list(df.get('subject', pd.Series(['General'])).dropna().unique())
        with col2: subject_filter = st.selectbox("Filter by Subject", subjects)
            
        if date_filter: df = df[df['date'] == date_filter.strftime("%Y-%m-%d")]
        if subject_filter != "All": df = df[df.get('subject', 'General') == subject_filter]

        # Top Summary Metrics
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        m1, m2, m3 = st.columns(3)
        m1.metric("🟢 Present Today", len(df[df['status'] == 'present']))
        m2.metric("🟡 Late Today", len(df[df['status'] == 'late']))
        m3.metric("🔴 Absent Today", len(df[df['status'] == 'absent']))
        st.markdown('</div>', unsafe_allow_html=True)

        if df.empty: st.warning("No records match your filters.")
        else:
            view_mode = st.radio("View Mode", ["Raw Data", "By Subject", "By Student", "📊 Attendance Summary"], horizontal=True)
            if view_mode == "Raw Data":
                st.dataframe(df.style.apply(lambda x: ['background: rgba(255,255,255,0.05)' if i%2==0 else '' for i in range(len(x))], axis=0), width='stretch')
            elif view_mode == "By Subject":
                st.bar_chart(df.groupby('subject').size())
                st.dataframe(df.groupby(['subject', 'date']).size().reset_index(name='Total Present'), width='stretch')
            elif view_mode == "By Student":
                st.dataframe(df.groupby(['name', 'id', 'subject']).size().reset_index(name='Total Days Attended'), width='stretch')
            elif view_mode == "📊 Attendance Summary":
                sessions_df = pd.read_sql_query("SELECT subject, COUNT(*) as total_sessions FROM sessions GROUP BY subject", self.db.connect())
                att_df = pd.read_sql_query("SELECT id, name, subject, COUNT(*) as attended FROM attendance WHERE status IN ('present', 'late') GROUP BY id, name, subject", self.db.connect())
                if not sessions_df.empty and not att_df.empty:
                    merged = pd.merge(att_df, sessions_df, on='subject', how='left')
                    merged['Attendance %'] = (merged['attended'] / merged['total_sessions'] * 100).fillna(0).round(1)
                    threshold = st.number_input("Highlight Below %", value=75.0, step=1.0)
                    st.dataframe(merged.style.apply(lambda row: ['background-color: rgba(255, 0, 0, 0.3)'] * len(row) if row['Attendance %'] < threshold else [''] * len(row), axis=1), width='stretch')
                else: st.info("No session data available yet. Close a session to generate summaries.")
            
        st.divider()
        colA, colB = st.columns(2)
        with colA:
            if st.session_state.confirm_clear:
                st.warning("⚠️ This will permanently delete ALL records. Click again to confirm.")
                if st.button("🚨 Confirm Delete All Records"):
                    self.db.clear_db()
                    st.session_state.confirm_clear = False
                    safe_rerun()
            else:
                if st.button("🗑 Clear All Attendance Records"):
                    st.session_state.confirm_clear = True
                    safe_rerun()
                
        with colB:
            c1, c2 = st.columns(2)
            with c1:
                st.spinner("Preparing Export...")
                full_excel = self.db.get_excel_data()
                st.download_button("📊 Export Excel Report", data=full_excel, file_name=f"attendance_report_{datetime.now().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width='stretch')
            with c2:
                filtered_excel = self.db.get_excel_data(df)
                st.download_button("📊 Export Filtered Data", data=filtered_excel, file_name=f"filtered_attendance_{datetime.now().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width='stretch')

    def run(self):
        inject_custom_css()
        self.sidebar()
        self.render_header()
        
        tab1, tab2, tab3, tab4 = st.tabs(["📷 Live Attendance", "👥 Student Management", "📘 Subjects Management", "📊 Attendance Records"])
        with tab1: self.camera_section()
        with tab2: self.student_management_section()
        with tab3: self.subjects_management_section()
        with tab4: self.records_section()

# =========================
if __name__ == "__main__":
    AttendanceApp().run()
