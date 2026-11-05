import os
import sys
import cv2
import datetime
import sqlite3
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw
import pandas as pd
from pyzbar import pyzbar

# ==========================================
# 0. DIRECTORIES & DATABASE PATHS
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(BASE_DIR, "database")
STUDENTS_DIR = os.path.join(DATABASE_DIR, "data_students")
DB_PATH = os.path.join(DATABASE_DIR, "attendance.db")
SHEET_PATH = os.path.join(DATABASE_DIR, "attendance.xlsx")
CSV_PATH = os.path.join(DATABASE_DIR, "attendance.csv")

os.makedirs(STUDENTS_DIR, exist_ok=True)

# OpenCV Haar Cascade Classifier for Face Detection
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')


# ==========================================
# 1. ROBUST BACKEND DATABASE ENGINE (SQLite + Excel + CSV)
# ==========================================
def get_db_connection():
    """Returns a SQLite connection with row factory enabled."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_database():
    """Initializes SQLite database tables and syncs with existing files."""
    conn = get_db_connection()
    cursor = conn.cursor()

    # Table for registered students
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            student_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            department TEXT DEFAULT 'General',
            registered_at TEXT NOT NULL,
            face_image_path TEXT,
            id_card_image_path TEXT
        )
    """)

    # Table for attendance records
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT NOT NULL,
            student_name TEXT NOT NULL,
            date TEXT NOT NULL,
            time TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Present',
            confidence REAL DEFAULT 0.0,
            verification_method TEXT DEFAULT 'Face Scan',
            UNIQUE(student_id, date)
        )
    """)
    conn.commit()

    # Migrate any existing text files in data_students into SQLite
    try:
        for fname in os.listdir(STUDENTS_DIR):
            if fname.endswith("_info.txt"):
                s_id = fname.replace("_info.txt", "")
                fpath = os.path.join(STUDENTS_DIR, fname)
                s_name = "Student"
                reg_date = str(datetime.date.today())
                with open(fpath, "r", encoding="utf-8") as f:
                    for line in f:
                        if line.startswith("Name:"):
                            s_name = line.split(":", 1)[1].strip()
                        elif line.startswith("Date:"):
                            reg_date = line.split(":", 1)[1].strip()

                face_img = os.path.join(STUDENTS_DIR, f"{s_id}_face.jpg")
                id_img = os.path.join(STUDENTS_DIR, f"{s_id}_id.jpg")
                cursor.execute("""
                    INSERT OR IGNORE INTO students (student_id, name, department, registered_at, face_image_path, id_card_image_path)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    str(s_id),
                    s_name,
                    'General',
                    reg_date,
                    face_img if os.path.exists(face_img) else None,
                    id_img if os.path.exists(id_img) else None
                ))
        conn.commit()
    except Exception as e:
        print(f"Migration note: {e}")

    # Migrate existing excel sheet if present
    if os.path.exists(SHEET_PATH):
        try:
            df = pd.read_excel(SHEET_PATH)
            if not df.empty and "Student ID" in df.columns and "Date" in df.columns:
                for _, row in df.iterrows():
                    s_id = str(row.get("Student ID", "")).strip()
                    s_name = str(row.get("Student Name", "Student")).strip()
                    s_date = str(row.get("Date", "")).split()[0].strip()
                    s_time = str(row.get("Time", "00:00:00")).strip()
                    s_status = str(row.get("Status", "Present")).strip()
                    if s_id and s_date:
                        cursor.execute("""
                            INSERT OR IGNORE INTO attendance (student_id, student_name, date, time, status, confidence, verification_method)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        """, (s_id, s_name, s_date, s_time, s_status, 1.0, 'Legacy Import'))
                conn.commit()
        except Exception as e:
            print(f"Excel import note: {e}")

    conn.close()
    sync_db_to_excel_and_csv()

def sync_db_to_excel_and_csv():
    """Syncs the attendance table to attendance.xlsx and attendance.csv safely."""
    try:
        conn = get_db_connection()
        df = pd.read_sql_query(
            "SELECT student_id AS [Student ID], student_name AS [Student Name], date AS [Date], time AS [Time], status AS [Status], confidence AS [Confidence], verification_method AS [Method] FROM attendance ORDER BY id ASC",
            conn
        )
        conn.close()

        # Save to CSV
        df.to_csv(CSV_PATH, index=False)

        # Save to Excel
        try:
            df.to_excel(SHEET_PATH, index=False)
        except PermissionError:
            print(f"[Notice] {SHEET_PATH} is open in Excel. Synced to CSV and SQLite.")
    except Exception as e:
        print(f"Database sync warning: {e}")

def db_register_student(student_id, name, department='General', face_path=None, id_path=None):
    """Registers a student into SQLite database and writes metadata text file."""
    conn = get_db_connection()
    cursor = conn.cursor()
    reg_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        INSERT OR REPLACE INTO students (student_id, name, department, registered_at, face_image_path, id_card_image_path)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (str(student_id), name, department, reg_date, face_path, id_path))
    conn.commit()
    conn.close()

    meta_path = os.path.join(STUDENTS_DIR, f"{student_id}_info.txt")
    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            f.write(f"ID: {student_id}\nName: {name}\nDepartment: {department}\nDate: {reg_date}\n")
    except Exception:
        pass

def db_get_student(student_id):
    """Retrieves student details by ID."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM students WHERE student_id = ?", (str(student_id),))
    row = cursor.fetchone()
    conn.close()
    if row:
        return dict(row)
    return None

def db_get_all_students():
    """Retrieves list of all registered students."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM students ORDER BY student_id ASC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_delete_student(student_id):
    """Deletes a student and their image files."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM students WHERE student_id = ?", (str(student_id),))
    conn.commit()
    conn.close()

    for suffix in ["_face.jpg", "_id.jpg", "_info.txt"]:
        p = os.path.join(STUDENTS_DIR, f"{student_id}{suffix}")
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass

def db_mark_attendance(student_id, student_name, confidence=1.0, method="Face Scan"):
    """
    Marks attendance into SQLite database and syncs to Excel/CSV.
    Guarantees no duplicate entries per day.
    """
    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")

    conn = get_db_connection()
    cursor = conn.cursor()

    # Check for duplicate entry today
    cursor.execute(
        "SELECT time FROM attendance WHERE student_id = ? AND date = ?",
        (str(student_id), date_str)
    )
    existing = cursor.fetchone()
    if existing:
        conn.close()
        return False, f"Attendance already marked today at {existing['time']} for {student_name}!"

    try:
        cursor.execute("""
            INSERT INTO attendance (student_id, student_name, date, time, status, confidence, verification_method)
            VALUES (?, ?, ?, ?, 'Present', ?, ?)
        """, (str(student_id), student_name, date_str, time_str, round(float(confidence), 2), method))
        conn.commit()
        conn.close()

        sync_db_to_excel_and_csv()
        return True, f"Attendance Successfully Marked!\n{student_name} ({student_id})\nTime: {time_str}"
    except Exception as e:
        conn.close()
        return False, f"Database Error: {str(e)}"

def db_get_attendance_records(date=None):
    """Fetches attendance records from database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    if date:
        cursor.execute("SELECT * FROM attendance WHERE date = ? ORDER BY id DESC", (str(date),))
    else:
        cursor.execute("SELECT * FROM attendance ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_get_stats():
    """Returns total students, today's attendance count, and date."""
    conn = get_db_connection()
    cursor = conn.cursor()
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    cursor.execute("SELECT COUNT(*) AS total FROM students")
    total_students = cursor.fetchone()["total"]

    cursor.execute("SELECT COUNT(*) AS today_count FROM attendance WHERE date = ?", (today_str,))
    today_attendance = cursor.fetchone()["today_count"]

    conn.close()
    return {
        "total_students": total_students,
        "today_attendance": today_attendance,
        "date": today_str
    }

# Initialize database on module load
init_database()


# ==========================================
# 2. ADVANCED FACIAL RECOGNITION & BIOMETRICS
# ==========================================
def extract_and_normalize_face(frame_bgr, target_size=(160, 160)):
    """
    Detects face in image, crops with margin, normalizes lighting using CLAHE,
    and returns (normalized_face_gray, cropped_bgr, bbox).
    """
    if frame_bgr is None:
        return None, None, None

    gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY) if len(frame_bgr.shape) == 3 else frame_bgr.copy()
    faces = face_cascade.detectMultiScale(gray, scaleFactor=1.15, minNeighbors=4, minSize=(60, 60))

    if len(faces) > 0:
        x, y, w, h = max(faces, key=lambda f: f[2] * f[3])
        pad_w = int(w * 0.12)
        pad_h = int(h * 0.12)
        y1 = max(0, y - pad_h)
        y2 = min(gray.shape[0], y + h + pad_h)
        x1 = max(0, x - pad_w)
        x2 = min(gray.shape[1], x + w + pad_w)

        crop_gray = gray[y1:y2, x1:x2]
        crop_bgr = frame_bgr[y1:y2, x1:x2] if len(frame_bgr.shape) == 3 else None
        bbox = (x, y, w, h)
    else:
        crop_gray = gray
        crop_bgr = frame_bgr
        bbox = (0, 0, gray.shape[1], gray.shape[0])

    crop_resized = cv2.resize(crop_gray, target_size)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    norm_face = clahe.apply(crop_resized)

    return norm_face, crop_bgr, bbox

def compare_faces_robust(norm_face_a, norm_face_b):
    """
    Multi-metric composite facial similarity matcher:
    1. Normalized Cross-Correlation Template Matching (cv2.TM_CCOEFF_NORMED)
    2. Edge / Gradient Feature Matching (Sobel)
    3. Normalized Luminance-Invariant Histogram Correlation
    Returns a unified score in range [0.0, 1.0].
    """
    if norm_face_a is None or norm_face_b is None:
        return 0.0

    # 1. Template Match Score
    t_match = cv2.matchTemplate(norm_face_a, norm_face_b, cv2.TM_CCOEFF_NORMED)[0][0]
    t_score = max(0.0, float(t_match))

    # 2. Gradient / Edge Structure Score
    grad_a = cv2.Sobel(norm_face_a, cv2.CV_32F, 1, 1, ksize=3)
    grad_b = cv2.Sobel(norm_face_b, cv2.CV_32F, 1, 1, ksize=3)
    cv2.normalize(grad_a, grad_a, 0, 1, cv2.NORM_MINMAX)
    cv2.normalize(grad_b, grad_b, 0, 1, cv2.NORM_MINMAX)
    g_match = cv2.matchTemplate(grad_a, grad_b, cv2.TM_CCOEFF_NORMED)[0][0]
    g_score = max(0.0, float(g_match))

    # 3. Histogram Distribution Score
    h1 = cv2.calcHist([norm_face_a], [0], None, [64], [0, 256])
    h2 = cv2.calcHist([norm_face_b], [0], None, [64], [0, 256])
    cv2.normalize(h1, h1, 0, 1, cv2.NORM_MINMAX)
    cv2.normalize(h2, h2, 0, 1, cv2.NORM_MINMAX)
    hist_score = max(0.0, float(cv2.compareHist(h1, h2, cv2.HISTCMP_CORREL)))

    # Composite weighted similarity
    final_score = (0.50 * t_score) + (0.30 * g_score) + (0.20 * hist_score)
    return round(float(final_score), 4)

def match_face_1_to_1(registered_img_path, captured_norm_face, threshold=0.45):
    """Verifies a captured face against a specific registered student's photo."""
    if not os.path.exists(registered_img_path):
        return False, 0.0

    reg_bgr = cv2.imread(registered_img_path)
    if reg_bgr is None:
        return False, 0.0

    reg_norm, _, _ = extract_and_normalize_face(reg_bgr)
    score = compare_faces_robust(reg_norm, captured_norm_face)
    return (score >= threshold), score

def identify_face_1_to_N(captured_norm_face, threshold=0.45):
    """
    1:N Automatic Facial Identification:
    Compares captured face against all registered students in the database.
    Returns (student_id, student_name, best_score, is_match).
    """
    students = db_get_all_students()
    if not students:
        return None, None, 0.0, False

    best_score = 0.0
    best_student = None

    for stu in students:
        s_id = stu["student_id"]
        face_path = stu.get("face_image_path")
        if not face_path or not os.path.exists(face_path):
            face_path = os.path.join(STUDENTS_DIR, f"{s_id}_face.jpg")

        if os.path.exists(face_path):
            reg_bgr = cv2.imread(face_path)
            if reg_bgr is not None:
                reg_norm, _, _ = extract_and_normalize_face(reg_bgr)
                score = compare_faces_robust(reg_norm, captured_norm_face)
                if score > best_score:
                    best_score = score
                    best_student = stu

    if best_student and best_score >= threshold:
        return best_student["student_id"], best_student["name"], best_score, True
    elif best_student:
        return best_student["student_id"], best_student["name"], best_score, False
    return None, None, 0.0, False

def create_placeholder_image(width, height, text="Connecting Camera..."):
    """Creates a clean placeholder Tkinter PhotoImage with exact pixel dimensions."""
    img = Image.new("RGB", (width, height), color=(15, 23, 42))
    draw = ImageDraw.Draw(img)
    draw.rectangle([(1, 1), (width - 2, height - 2)], outline=(51, 65, 85), width=2)
    draw.text((width // 2, height // 2), text, fill=(148, 163, 184), anchor="mm")
    return ImageTk.PhotoImage(img)


# ==========================================
# 3. REGISTRATION WINDOW (CLEAN, PROPORTIONAL LAYOUT)
# ==========================================
class RegisterWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Register Student - Face & ID Portal")
        self.geometry("740x590")
        self.minsize(720, 560)
        self.configure(bg="#0F172A")

        self.student_id_var = tk.StringVar()
        self.student_name_var = tk.StringVar()
        self.department_var = tk.StringVar(value="Computer Science")
        self.status_var = tk.StringVar(value="Connecting camera... Please wait a moment.")

        self.captured_face_img = None
        self.captured_id_img = None
        self.current_frame = None
        self.is_running = True
        self.timer_id = None
        self.cap = None

        self.build_ui()
        self.start_camera_thread()
        self.update_webcam()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def start_camera_thread(self):
        """Initializes the webcam in a non-blocking background thread."""
        def _connect():
            try:
                cap = cv2.VideoCapture(0)
                if cap.isOpened():
                    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                    self.cap = cap
                    if self.is_running:
                        self.status_var.set("Camera connected. Align face and click '📸 Scan & Capture Face'")
                else:
                    if self.is_running:
                        self.status_var.set("Camera not detected. You can use '📁 Browse Face File'.")
            except Exception as e:
                print(f"Camera init error: {e}")

        threading.Thread(target=_connect, daemon=True).start()

    def build_ui(self):
        # Header banner
        header = tk.Frame(self, bg="#1E293B", height=45)
        header.pack(fill=tk.X)
        lbl_title = tk.Label(header, text="Student Registration Portal", font=("Arial", 14, "bold"), fg="#38BDF8", bg="#1E293B")
        lbl_title.pack(pady=8)

        # Form Inputs Frame
        form = tk.Frame(self, bg="#1E293B", padx=12, pady=6)
        form.pack(fill=tk.X, padx=15, pady=6)

        tk.Label(form, text="Student ID *:", font=("Arial", 10, "bold"), fg="#E2E8F0", bg="#1E293B").grid(row=0, column=0, padx=5, pady=4, sticky="e")
        self.entry_id = tk.Entry(form, textvariable=self.student_id_var, font=("Arial", 10), width=13, bg="#334155", fg="white", insertbackground="white")
        self.entry_id.grid(row=0, column=1, padx=5, pady=4)

        tk.Label(form, text="Name *:", font=("Arial", 10, "bold"), fg="#E2E8F0", bg="#1E293B").grid(row=0, column=2, padx=5, pady=4, sticky="e")
        self.entry_name = tk.Entry(form, textvariable=self.student_name_var, font=("Arial", 10), width=16, bg="#334155", fg="white", insertbackground="white")
        self.entry_name.grid(row=0, column=3, padx=5, pady=4)

        tk.Label(form, text="Department:", font=("Arial", 10, "bold"), fg="#E2E8F0", bg="#1E293B").grid(row=0, column=4, padx=5, pady=4, sticky="e")
        self.entry_dept = tk.Entry(form, textvariable=self.department_var, font=("Arial", 10), width=15, bg="#334155", fg="white", insertbackground="white")
        self.entry_dept.grid(row=0, column=5, padx=5, pady=4)

        # Main Center Container (Camera on Left, Previews on Right)
        center_frame = tk.Frame(self, bg="#0F172A")
        center_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=4)

        # Fixed-size Camera Container (480x320 pixels)
        self.cam_box = tk.Frame(center_frame, width=480, height=320, bg="#020617", relief=tk.SOLID, borderwidth=1)
        self.cam_box.pack_propagate(False)
        self.cam_box.pack(side=tk.LEFT, padx=(0, 10))

        self.video_placeholder = create_placeholder_image(480, 320, "Connecting Camera...")
        self.video_label = tk.Label(self.cam_box, image=self.video_placeholder, bg="#020617")
        self.video_label.pack(fill=tk.BOTH, expand=True)

        # Previews Sidebar Container (Fixed 180px width)
        side_box = tk.LabelFrame(center_frame, text="Captured Images", font=("Arial", 9, "bold"), fg="#94A3B8", bg="#1E293B", padx=8, pady=6)
        side_box.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Face Preview Box
        tk.Label(side_box, text="Face Photo:", font=("Arial", 9), fg="#CBD5E1", bg="#1E293B").pack(anchor="w")
        self.face_box = tk.Frame(side_box, width=120, height=110, bg="#334155", relief=tk.SUNKEN, borderwidth=1)
        self.face_box.pack_propagate(False)
        self.face_box.pack(pady=3)
        self.face_placeholder = create_placeholder_image(120, 110, "No Face Yet")
        self.lbl_face_preview = tk.Label(self.face_box, image=self.face_placeholder, bg="#334155")
        self.lbl_face_preview.pack(fill=tk.BOTH, expand=True)

        # ID Card Preview Box
        tk.Label(side_box, text="ID Card (Optional):", font=("Arial", 9), fg="#CBD5E1", bg="#1E293B").pack(anchor="w", pady=(4, 0))
        self.id_box = tk.Frame(side_box, width=120, height=90, bg="#334155", relief=tk.SUNKEN, borderwidth=1)
        self.id_box.pack_propagate(False)
        self.id_box.pack(pady=3)
        self.id_placeholder = create_placeholder_image(120, 90, "Optional ID")
        self.lbl_id_preview = tk.Label(self.id_box, image=self.id_placeholder, bg="#334155")
        self.lbl_id_preview.pack(fill=tk.BOTH, expand=True)

        # Status Label Bar
        self.lbl_status = tk.Label(self, textvariable=self.status_var, font=("Arial", 10, "bold"), fg="#38BDF8", bg="#0F172A")
        self.lbl_status.pack(pady=4)

        # Bottom Action Buttons
        btn_frame = tk.Frame(self, bg="#0F172A", pady=6)
        btn_frame.pack(fill=tk.X, padx=15)

        self.btn_capture_face = tk.Button(
            btn_frame,
            text="📸 Scan & Capture Face",
            font=("Arial", 10, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self.capture_face
        )
        self.btn_capture_face.pack(side=tk.LEFT, padx=4)

        btn_upload_face = tk.Button(
            btn_frame,
            text="📁 Browse Face File",
            font=("Arial", 9),
            bg="#475569",
            fg="white",
            padx=8,
            pady=6,
            cursor="hand2",
            command=self.upload_face_file
        )
        btn_upload_face.pack(side=tk.LEFT, padx=4)

        self.btn_capture_id = tk.Button(
            btn_frame,
            text="🪪 Capture ID Card",
            font=("Arial", 9),
            bg="#475569",
            fg="white",
            padx=8,
            pady=6,
            cursor="hand2",
            command=self.capture_id_card
        )
        self.btn_capture_id.pack(side=tk.LEFT, padx=4)

        self.btn_save = tk.Button(
            btn_frame,
            text="💾 Save in Database",
            font=("Arial", 10, "bold"),
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            padx=14,
            pady=6,
            cursor="hand2",
            state=tk.DISABLED,
            command=self.save_data
        )
        self.btn_save.pack(side=tk.LEFT, padx=6)

        btn_cancel = tk.Button(
            btn_frame,
            text="Close",
            font=("Arial", 10),
            bg="#DC2626",
            fg="white",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self.on_close
        )
        btn_cancel.pack(side=tk.RIGHT, padx=4)

    def update_webcam(self):
        if not self.is_running:
            return

        if self.cap and self.cap.isOpened():
            ret, frame = self.cap.read()
            if ret and frame is not None:
                self.current_frame = frame.copy()
                display_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

                # Real-time face detection overlay
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = face_cascade.detectMultiScale(gray, scaleFactor=1.15, minNeighbors=4, minSize=(60, 60))
                for (x, y, w, h) in faces:
                    cv2.rectangle(display_frame, (x, y), (x + w, y + h), (34, 197, 94), 2)
                    cv2.putText(display_frame, "Face Ready", (x, y - 8),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.55, (34, 197, 94), 2)

                img = Image.fromarray(display_frame).resize((480, 320))
                imgtk = ImageTk.PhotoImage(image=img)
                self.video_label.imgtk = imgtk
                self.video_label.configure(image=imgtk)

        if self.is_running:
            self.timer_id = self.after(30, self.update_webcam)

    def capture_face(self):
        s_id = self.student_id_var.get().strip()
        s_name = self.student_name_var.get().strip()

        if not s_id or not s_name:
            messagebox.showwarning("Incomplete Form", "Please fill in Student ID and Name first.")
            self.entry_id.focus()
            return

        if self.current_frame is None:
            messagebox.showerror("Camera Error", "No frame from webcam. You can also use 'Browse Face File'.")
            return

        gray = cv2.cvtColor(self.current_frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.15, minNeighbors=4, minSize=(60, 60))

        if len(faces) == 0:
            messagebox.showwarning("No Face Detected", "No clear face detected in the frame.\nPlease look straight into the camera and try again.")
            return

        self.captured_face_img = self.current_frame.copy()

        # Generate thumbnail preview
        x, y, w, h = max(faces, key=lambda f: f[2]*f[3])
        face_crop = self.current_frame[y:y+h, x:x+w]
        preview_rgb = cv2.cvtColor(face_crop, cv2.COLOR_BGR2RGB)
        thumb = Image.fromarray(preview_rgb).resize((120, 110))
        thumb_tk = ImageTk.PhotoImage(image=thumb)
        self.lbl_face_preview.imgtk = thumb_tk
        self.lbl_face_preview.configure(image=thumb_tk)

        self.status_var.set("Face captured successfully! Now click '💾 Save in Database'.")
        self.btn_save.configure(state=tk.NORMAL, bg="#10B981")
        messagebox.showinfo("Face Captured", "Face photo captured successfully! Click 'Save in Database' to complete.")

    def upload_face_file(self):
        s_id = self.student_id_var.get().strip()
        s_name = self.student_name_var.get().strip()

        if not s_id or not s_name:
            messagebox.showwarning("Incomplete Form", "Please fill in Student ID and Name first.")
            self.entry_id.focus()
            return

        file_path = filedialog.askopenfilename(
            title="Select Student Photo",
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.bmp")]
        )
        if file_path:
            img_bgr = cv2.imread(file_path)
            if img_bgr is not None:
                self.captured_face_img = img_bgr
                preview_rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
                thumb = Image.fromarray(preview_rgb).resize((120, 110))
                thumb_tk = ImageTk.PhotoImage(image=thumb)
                self.lbl_face_preview.imgtk = thumb_tk
                self.lbl_face_preview.configure(image=thumb_tk)

                self.status_var.set("Face photo loaded from file! Click '💾 Save in Database'.")
                self.btn_save.configure(state=tk.NORMAL, bg="#10B981")

    def capture_id_card(self):
        if self.current_frame is None:
            messagebox.showerror("Camera Error", "Webcam is not ready.")
            return

        self.captured_id_img = self.current_frame.copy()
        preview_rgb = cv2.cvtColor(self.captured_id_img, cv2.COLOR_BGR2RGB)
        thumb = Image.fromarray(preview_rgb).resize((120, 90))
        thumb_tk = ImageTk.PhotoImage(image=thumb)
        self.lbl_id_preview.imgtk = thumb_tk
        self.lbl_id_preview.configure(image=thumb_tk)
        self.status_var.set("ID Card photo captured!")

    def save_data(self):
        s_id = self.student_id_var.get().strip()
        s_name = self.student_name_var.get().strip()
        dept = self.department_var.get().strip() or "General"

        if not s_id or not s_name:
            messagebox.showwarning("Incomplete Form", "Student ID and Name are required.")
            return

        if self.captured_face_img is None:
            messagebox.showwarning("Missing Face", "Please click 'Scan & Capture Face' or upload a photo first.")
            return

        face_path = os.path.join(STUDENTS_DIR, f"{s_id}_face.jpg")
        id_path = os.path.join(STUDENTS_DIR, f"{s_id}_id.jpg")

        cv2.imwrite(face_path, self.captured_face_img)

        if self.captured_id_img is not None:
            cv2.imwrite(id_path, self.captured_id_img)
        else:
            cv2.imwrite(id_path, self.captured_face_img)

        try:
            db_register_student(s_id, s_name, dept, face_path, id_path)
            messagebox.showinfo(
                "Registration Successful",
                f"Student Registered Successfully!\n\nID: {s_id}\nName: {s_name}\nDepartment: {dept}\n\nStored in SQLite Database & Excel records."
            )
            self.on_close()
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to save to database: {str(e)}")

    def on_close(self):
        self.is_running = False
        if self.timer_id:
            try:
                self.after_cancel(self.timer_id)
            except Exception:
                pass
        if self.cap and self.cap.isOpened():
            try:
                self.cap.release()
            except Exception:
                pass
        self.destroy()


# ==========================================
# 4. ATTENDANCE SCANNER WINDOW (CLEAN, PROPORTIONAL LAYOUT)
# ==========================================
class AttendanceWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Smart Attendance Verification - Face & ID Scanner")
        self.geometry("740x590")
        self.minsize(720, 550)
        self.configure(bg="#0F172A")

        self.selected_id = None
        self.selected_name = None
        self.status_var = tk.StringVar(value="Connecting camera... Please wait a moment.")
        self.current_frame = None
        self.is_running = True
        self.timer_id = None
        self.cap = None
        self.last_qr_scanned = ""
        self.last_qr_time = 0

        self.build_ui()
        self.start_camera_thread()
        self.update_webcam()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def start_camera_thread(self):
        """Initializes the webcam in a non-blocking background thread."""
        def _connect():
            try:
                cap = cv2.VideoCapture(0)
                if cap.isOpened():
                    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                    self.cap = cap
                    if self.is_running:
                        self.status_var.set("Camera connected. Click '🔍 Scan Face & Mark Attendance' (or enter/scan ID)")
                else:
                    if self.is_running:
                        self.status_var.set("Camera not detected. You can use '📁 Verify from Photo File'.")
            except Exception as e:
                print(f"Camera init error: {e}")

        threading.Thread(target=_connect, daemon=True).start()

    def build_ui(self):
        # Header banner
        header = tk.Frame(self, bg="#0F766E", height=45)
        header.pack(fill=tk.X)
        lbl_title = tk.Label(header, text="Mark Attendance Verification", font=("Arial", 14, "bold"), fg="white", bg="#0F766E")
        lbl_title.pack(pady=8)

        # Status Bar
        status_frame = tk.Frame(self, bg="#134E4A", padx=10, pady=5)
        status_frame.pack(fill=tk.X, padx=15, pady=4)
        self.lbl_status = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Arial", 10, "bold"),
            fg="#5EEAD4",
            bg="#134E4A"
        )
        self.lbl_status.pack()

        # Center Camera Container (Fixed 480x300 pixels)
        self.cam_box = tk.Frame(self, width=480, height=300, bg="#020617", relief=tk.SOLID, borderwidth=1)
        self.cam_box.pack_propagate(False)
        self.cam_box.pack(pady=6)

        self.video_placeholder = create_placeholder_image(480, 300, "Connecting Camera...")
        self.video_label = tk.Label(self.cam_box, image=self.video_placeholder, bg="#020617")
        self.video_label.pack(fill=tk.BOTH, expand=True)

        # ID Lookup & Mode Bar
        id_frame = tk.Frame(self, bg="#1E293B", padx=10, pady=6)
        id_frame.pack(fill=tk.X, padx=15, pady=4)

        tk.Label(id_frame, text="Student ID (Optional):", font=("Arial", 9, "bold"), fg="#E2E8F0", bg="#1E293B").grid(row=0, column=0, padx=4)
        self.entry_id = tk.Entry(id_frame, font=("Arial", 10), width=12, bg="#334155", fg="white", insertbackground="white")
        self.entry_id.grid(row=0, column=1, padx=4)
        self.entry_id.bind("<Return>", lambda event: self.handle_manual_id())

        btn_find = tk.Button(
            id_frame,
            text="Verify ID",
            font=("Arial", 9, "bold"),
            bg="#0284C7",
            fg="white",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.handle_manual_id
        )
        btn_find.grid(row=0, column=2, padx=4)

        btn_clear_id = tk.Button(
            id_frame,
            text="Clear",
            font=("Arial", 9),
            bg="#64748B",
            fg="white",
            padx=6,
            pady=3,
            cursor="hand2",
            command=self.clear_selected_id
        )
        btn_clear_id.grid(row=0, column=3, padx=4)

        self.lbl_mode = tk.Label(
            id_frame,
            text="Mode: Direct 1:N Facial Scan (Matches all registered students)",
            font=("Arial", 9, "italic"),
            fg="#FDE047",
            bg="#1E293B"
        )
        self.lbl_mode.grid(row=0, column=4, padx=10, sticky="w")

        # Action Buttons Bar
        ctrl_frame = tk.Frame(self, bg="#0F172A", pady=6)
        ctrl_frame.pack(fill=tk.X, padx=15)

        # MAIN SCAN FACE BUTTON - ALWAYS ACTIVE & PROMINENT
        self.btn_scan_face = tk.Button(
            ctrl_frame,
            text="🔍 Scan Face & Mark Attendance",
            font=("Arial", 11, "bold"),
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            padx=16,
            pady=8,
            cursor="hand2",
            command=self.scan_and_mark_attendance
        )
        self.btn_scan_face.pack(side=tk.LEFT, padx=6)

        btn_browse_photo = tk.Button(
            ctrl_frame,
            text="📁 Verify from Photo File",
            font=("Arial", 9),
            bg="#475569",
            fg="white",
            padx=10,
            pady=8,
            cursor="hand2",
            command=self.verify_from_file
        )
        btn_browse_photo.pack(side=tk.LEFT, padx=6)

        btn_close = tk.Button(
            ctrl_frame,
            text="Close",
            font=("Arial", 10),
            bg="#DC2626",
            fg="white",
            padx=14,
            pady=8,
            cursor="hand2",
            command=self.on_close
        )
        btn_close.pack(side=tk.RIGHT, padx=6)

    def handle_manual_id(self):
        s_id = self.entry_id.get().strip()
        if not s_id:
            messagebox.showwarning("Input Required", "Please enter a Student ID to search.")
            return

        student = db_get_student(s_id)
        if not student:
            face_img = os.path.join(STUDENTS_DIR, f"{s_id}_face.jpg")
            if os.path.exists(face_img):
                student = {"student_id": s_id, "name": "Student"}
            else:
                messagebox.showerror("Student Not Found", f"Student ID '{s_id}' was not found in the database!\nPlease register the student first.")
                return

        self.selected_id = student["student_id"]
        self.selected_name = student["name"]
        self.lbl_mode.configure(
            text=f"Mode: 1:1 Verification for {self.selected_name} ({self.selected_id})",
            fg="#86EFAC"
        )
        self.status_var.set(f"Selected: {self.selected_name} ({self.selected_id}). Click 'Scan Face' to verify!")

    def clear_selected_id(self):
        self.selected_id = None
        self.selected_name = None
        self.entry_id.delete(0, tk.END)
        self.lbl_mode.configure(
            text="Mode: Direct 1:N Facial Scan (Matches all registered students)",
            fg="#FDE047"
        )
        self.status_var.set("Mode reset to 1:N Automatic Face Recognition. Click 'Scan Face & Mark Attendance' anytime!")

    def update_webcam(self):
        if not self.is_running:
            return

        if self.cap and self.cap.isOpened():
            ret, frame = self.cap.read()
            if ret and frame is not None:
                self.current_frame = frame.copy()
                display_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

                # Real-time QR Code scanning
                now_ts = datetime.datetime.now().timestamp()
                try:
                    decoded = pyzbar.decode(frame)
                    for obj in decoded:
                        qr_data = obj.data.decode("utf-8").strip()
                        x, y, w, h = obj.rect
                        cv2.rectangle(display_frame, (x, y), (x + w, y + h), (245, 158, 11), 2)
                        cv2.putText(display_frame, f"QR: {qr_data}", (x, y - 6),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (245, 158, 11), 2)

                        if qr_data and (qr_data != self.last_qr_scanned or (now_ts - self.last_qr_time) > 4):
                            self.last_qr_scanned = qr_data
                            self.last_qr_time = now_ts
                            self.entry_id.delete(0, tk.END)
                            self.entry_id.insert(0, qr_data)
                            self.handle_manual_id()
                except Exception:
                    pass

                # Live face detection box
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = face_cascade.detectMultiScale(gray, scaleFactor=1.15, minNeighbors=4, minSize=(60, 60))
                for (x, y, w, h) in faces:
                    target_label = self.selected_name if self.selected_name else "Face Detected"
                    color = (34, 197, 94) if self.selected_name else (56, 189, 248)
                    cv2.rectangle(display_frame, (x, y), (x + w, y + h), color, 2)
                    cv2.putText(display_frame, target_label, (x, y - 8),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.55, color, 2)

                img = Image.fromarray(display_frame).resize((480, 300))
                imgtk = ImageTk.PhotoImage(image=img)
                self.video_label.imgtk = imgtk
                self.video_label.configure(image=imgtk)

        if self.is_running:
            self.timer_id = self.after(30, self.update_webcam)

    def scan_and_mark_attendance(self):
        if self.current_frame is None:
            messagebox.showerror("Camera Error", "No frame received from webcam. Please check your camera.")
            return

        norm_face, crop_bgr, bbox = extract_and_normalize_face(self.current_frame)
        if norm_face is None or bbox[2] < 50:
            messagebox.showwarning("No Face Detected", "No face was detected in front of the camera!\nPlease face the camera and try again.")
            return

        self._process_face_recognition(norm_face)

    def verify_from_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Face Photo for Attendance",
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.bmp")]
        )
        if file_path:
            img_bgr = cv2.imread(file_path)
            if img_bgr is not None:
                norm_face, _, bbox = extract_and_normalize_face(img_bgr)
                if norm_face is None or bbox[2] < 40:
                    messagebox.showwarning("No Face Found", "No face detected in selected image file.")
                    return
                self._process_face_recognition(norm_face)

    def _process_face_recognition(self, norm_face):
        # MODE 1: 1:1 Verification against selected ID
        if self.selected_id:
            face_path = os.path.join(STUDENTS_DIR, f"{self.selected_id}_face.jpg")
            if not os.path.exists(face_path):
                messagebox.showerror("Missing Biometric", f"No face template registered for ID '{self.selected_id}'.")
                return

            matched, score = match_face_1_to_1(face_path, norm_face, threshold=0.45)
            if matched:
                success, msg = db_mark_attendance(
                    self.selected_id,
                    self.selected_name,
                    confidence=score,
                    method="1:1 Face Verification"
                )
                if success:
                    messagebox.showinfo("Success", f"{msg}\nConfidence: {score*100:.1f}%")
                    self.on_close()
                else:
                    messagebox.showwarning("Notice", msg)
            else:
                messagebox.showerror(
                    "Verification Failed",
                    f"Face does NOT match registered student '{self.selected_name}' ({self.selected_id})!\nMatch Score: {score*100:.1f}% (Required: >= 45%)"
                )

        # MODE 2: 1:N Automatic Identification
        else:
            all_students = db_get_all_students()
            if not all_students:
                messagebox.showwarning(
                    "No Registered Students",
                    "There are no students registered in the database yet.\nPlease register at least one student first!"
                )
                return

            stu_id, stu_name, score, is_match = identify_face_1_to_N(norm_face, threshold=0.45)

            if is_match and stu_id:
                success, msg = db_mark_attendance(
                    stu_id,
                    stu_name,
                    confidence=score,
                    method="1:N Auto Face Scan"
                )
                if success:
                    messagebox.showinfo(
                        "Attendance Marked!",
                        f"Student Identified: {stu_name} ({stu_id})\nConfidence: {score*100:.1f}%\n\n{msg}"
                    )
                    self.on_close()
                else:
                    messagebox.showwarning("Notice", msg)
            else:
                score_msg = f"{score*100:.1f}%" if score > 0 else "0%"
                closest_name = f" (Closest match: {stu_name})" if stu_name else ""
                messagebox.showerror(
                    "Face Not Recognized",
                    f"Face not recognized in database.{closest_name}\nConfidence: {score_msg} (Required: >= 45%).\n\nTips:\n1. Look straight into camera.\n2. Or enter your Student ID in the box and click 'Verify ID'."
                )

    def on_close(self):
        self.is_running = False
        if self.timer_id:
            try:
                self.after_cancel(self.timer_id)
            except Exception:
                pass
        if self.cap and self.cap.isOpened():
            try:
                self.cap.release()
            except Exception:
                pass
        self.destroy()


# ==========================================
# 5. DATABASE VIEWER & LOGS WINDOW (PROPORTIONAL LAYOUT)
# ==========================================
class DatabaseViewerWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Backend Database & Attendance Records Viewer")
        self.geometry("820x580")
        self.minsize(780, 520)
        self.configure(bg="#0F172A")

        self.build_ui()
        self.load_data()

    def build_ui(self):
        header = tk.Frame(self, bg="#1E293B", height=45)
        header.pack(fill=tk.X)
        lbl_title = tk.Label(header, text="Backend Database & Attendance Records", font=("Arial", 14, "bold"), fg="#38BDF8", bg="#1E293B")
        lbl_title.pack(pady=8)

        stats = db_get_stats()
        metrics = tk.Frame(self, bg="#1E293B", padx=10, pady=6)
        metrics.pack(fill=tk.X, padx=15, pady=6)

        self.lbl_stat_total = tk.Label(metrics, text=f"👥 Total Students: {stats['total_students']}", font=("Arial", 10, "bold"), fg="#38BDF8", bg="#1E293B")
        self.lbl_stat_total.grid(row=0, column=0, padx=10)

        self.lbl_stat_today = tk.Label(metrics, text=f"✅ Today's Present: {stats['today_attendance']}", font=("Arial", 10, "bold"), fg="#4ADE80", bg="#1E293B")
        self.lbl_stat_today.grid(row=0, column=1, padx=10)

        self.lbl_stat_date = tk.Label(metrics, text=f"📅 Date: {stats['date']}", font=("Arial", 10), fg="#94A3B8", bg="#1E293B")
        self.lbl_stat_date.grid(row=0, column=2, padx=10)

        btn_sync = tk.Button(metrics, text="🔄 Sync Excel", font=("Arial", 9, "bold"), bg="#0284C7", fg="white", padx=8, pady=2, command=self.sync_and_refresh)
        btn_sync.grid(row=0, column=3, padx=10)

        btn_open_folder = tk.Button(metrics, text="📁 Open Folder", font=("Arial", 9), bg="#475569", fg="white", padx=8, pady=2, command=self.open_db_folder)
        btn_open_folder.grid(row=0, column=4, padx=6)

        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=6)

        # Tab 1: Attendance Records
        tab_att = tk.Frame(notebook, bg="#0F172A")
        notebook.add(tab_att, text="  Attendance Records  ")

        att_cols = ("ID", "Student ID", "Student Name", "Date", "Time", "Status", "Confidence", "Method")
        self.tree_att = ttk.Treeview(tab_att, columns=att_cols, show="headings", height=13)
        for col in att_cols:
            self.tree_att.heading(col, text=col)
            width = 60 if col in ("ID", "Status", "Confidence") else 110
            self.tree_att.column(col, width=width, anchor="center")

        scroll_att = ttk.Scrollbar(tab_att, orient=tk.VERTICAL, command=self.tree_att.yview)
        self.tree_att.configure(yscroll=scroll_att.set)
        self.tree_att.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_att.pack(side=tk.RIGHT, fill=tk.Y)

        # Tab 2: Registered Students
        tab_stu = tk.Frame(notebook, bg="#0F172A")
        notebook.add(tab_stu, text="  Registered Students  ")

        stu_cols = ("Student ID", "Name", "Department", "Registered At", "Face Template")
        self.tree_stu = ttk.Treeview(tab_stu, columns=stu_cols, show="headings", height=13)
        for col in stu_cols:
            self.tree_stu.heading(col, text=col)
            width = 110 if col != "Registered At" else 140
            self.tree_stu.column(col, width=width, anchor="center")

        scroll_stu = ttk.Scrollbar(tab_stu, orient=tk.VERTICAL, command=self.tree_stu.yview)
        self.tree_stu.configure(yscroll=scroll_stu.set)
        self.tree_stu.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_stu.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom Action Bar
        bottom = tk.Frame(self, bg="#0F172A", pady=6)
        bottom.pack(fill=tk.X, padx=15)

        btn_delete_stu = tk.Button(
            bottom,
            text="Delete Selected Student",
            font=("Arial", 9),
            bg="#EF4444",
            fg="white",
            padx=10,
            pady=4,
            command=self.delete_selected_student
        )
        btn_delete_stu.pack(side=tk.LEFT, padx=5)

        btn_refresh = tk.Button(
            bottom,
            text="Refresh Tables",
            font=("Arial", 9, "bold"),
            bg="#10B981",
            fg="white",
            padx=12,
            pady=4,
            command=self.load_data
        )
        btn_refresh.pack(side=tk.RIGHT, padx=5)

    def load_data(self):
        stats = db_get_stats()
        self.lbl_stat_total.configure(text=f"👥 Total Students: {stats['total_students']}")
        self.lbl_stat_today.configure(text=f"✅ Today's Present: {stats['today_attendance']}")
        self.lbl_stat_date.configure(text=f"📅 Date: {stats['date']}")

        for item in self.tree_att.get_children():
            self.tree_att.delete(item)

        records = db_get_attendance_records()
        for r in records:
            conf_str = f"{r.get('confidence', 1.0)*100:.0f}%" if r.get('confidence') else "100%"
            self.tree_att.insert("", tk.END, values=(
                r.get("id"),
                r.get("student_id"),
                r.get("student_name"),
                r.get("date"),
                r.get("time"),
                r.get("status"),
                conf_str,
                r.get("verification_method", "Face Scan")
            ))

        for item in self.tree_stu.get_children():
            self.tree_stu.delete(item)

        students = db_get_all_students()
        for s in students:
            s_id = s.get("student_id")
            face_img = s.get("face_image_path")
            has_face = "Ready" if (face_img and os.path.exists(face_img)) or os.path.exists(os.path.join(STUDENTS_DIR, f"{s_id}_face.jpg")) else "Missing"
            self.tree_stu.insert("", tk.END, values=(
                s_id,
                s.get("name"),
                s.get("department", "General"),
                s.get("registered_at"),
                has_face
            ))

    def sync_and_refresh(self):
        sync_db_to_excel_and_csv()
        self.load_data()
        messagebox.showinfo("Synced", f"Database successfully synced to:\n- {SHEET_PATH}\n- {CSV_PATH}\n- {DB_PATH}")

    def delete_selected_student(self):
        selected = self.tree_stu.selection()
        if not selected:
            messagebox.showwarning("Select Student", "Please select a student from the 'Registered Students' tab to delete.")
            return

        values = self.tree_stu.item(selected[0], "values")
        s_id = values[0]
        s_name = values[1]

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete student '{s_name}' ({s_id}) from the database?"):
            db_delete_student(s_id)
            self.load_data()
            messagebox.showinfo("Deleted", f"Student {s_name} ({s_id}) deleted.")

    def open_db_folder(self):
        try:
            if sys.platform == "win32":
                os.startfile(DATABASE_DIR)
            else:
                subprocess.Popen(["xdg-open", DATABASE_DIR])
        except Exception as e:
            messagebox.showinfo("Folder Path", f"Database folder is at:\n{DATABASE_DIR}")


# ==========================================
# 6. MAIN APPLICATION PORTAL
# ==========================================
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smart Attendance Management Portal")
        self.geometry("540x480")
        self.configure(bg="#0F172A")
        self.resizable(False, False)

        self.build_ui()

    def build_ui(self):
        header = tk.Frame(self, bg="#1E293B", pady=18)
        header.pack(fill=tk.X)

        title = tk.Label(header, text="Smart Attendance Portal", font=("Arial", 18, "bold"), fg="#38BDF8", bg="#1E293B")
        title.pack()

        subtitle = tk.Label(header, text="AI Facial Recognition & ID Biometric Verification", font=("Arial", 10), fg="#94A3B8", bg="#1E293B")
        subtitle.pack(pady=3)

        body = tk.Frame(self, bg="#0F172A", pady=16)
        body.pack(expand=True)

        btn_register = tk.Button(
            body,
            text="1. Register Student (Face & ID)",
            font=("Arial", 11, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            width=28,
            height=2,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.open_register
        )
        btn_register.pack(pady=7)

        btn_attendance = tk.Button(
            body,
            text="2. Mark Attendance (Scan Face / ID)",
            font=("Arial", 11, "bold"),
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            width=28,
            height=2,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.open_attendance
        )
        btn_attendance.pack(pady=7)

        btn_database = tk.Button(
            body,
            text="3. View Database & Attendance Logs",
            font=("Arial", 11, "bold"),
            bg="#0891B2",
            fg="white",
            activebackground="#0e7490",
            activeforeground="white",
            width=28,
            height=2,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.open_database_viewer
        )
        btn_database.pack(pady=7)

        stats = db_get_stats()
        footer_text = f"SQLite: Connected | Registered: {stats['total_students']} | Today's Attendance: {stats['today_attendance']}"
        self.lbl_footer = tk.Label(
            self,
            text=footer_text,
            font=("Arial", 9),
            fg="#64748B",
            bg="#0F172A"
        )
        self.lbl_footer.pack(side=tk.BOTTOM, pady=10)

    def open_register(self):
        win = RegisterWindow(self)
        self.wait_window(win)
        self.refresh_footer()

    def open_attendance(self):
        win = AttendanceWindow(self)
        self.wait_window(win)
        self.refresh_footer()

    def open_database_viewer(self):
        win = DatabaseViewerWindow(self)
        self.wait_window(win)
        self.refresh_footer()

    def refresh_footer(self):
        stats = db_get_stats()
        self.lbl_footer.configure(
            text=f"SQLite: Connected | Registered: {stats['total_students']} | Today's Attendance: {stats['today_attendance']}"
        )


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
