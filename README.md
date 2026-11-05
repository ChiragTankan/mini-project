# Smart Attendance Management System

A Python-based multimodal attendance verification system powered by Tkinter GUI, OpenCV computer vision, QR/barcode ID scanning, and a robust dual-layer database backend (SQLite + Excel + CSV).

## Folder Structure
```
smart_attendance_system/
├── main.py                    # Main application GUI & complete workflow
├── presentation_content.md    # 8-slide presentation outline & speaker notes for defense
├── requirements.txt           # Python dependencies
├── database/
│   ├── attendance.db          # ACID-compliant SQLite backend database
│   ├── attendance.xlsx        # Excel attendance records (Date, Time, Status, Confidence)
│   ├── attendance.csv         # CSV attendance export
│   └── data_students/         # Registered student portraits and ID snapshots (.jpg)
└── README.md                  # Project setup and guide
```

## Features

1. **Student Registration (`1. Register Student`)**:
   - Live face detection with CLAHE lighting normalization.
   - **Scan & Capture Face**: Direct 1-click facial biometric capture with thumbnail preview.
   - Optional ID card capture.
   - Saves records into SQLite `students` database table and stores photos in `database/data_students/`.

2. **Attendance Marking (`2. Mark Attendance`)**:
   - **Always-Enabled "Scan Face & Mark Attendance" Button**:
     - **Mode 1 (Direct 1:N Facial Recognition)**: Simply face the camera and click the button; the system compares your face against all registered students, identifies you with confidence score, and logs your attendance.
     - **Mode 2 (1:1 Verified Biometric)**: Scan student QR/barcode or enter Student ID manually, then verify against that specific student profile.
   - Real-time feedback with green bounding box and recognized name overlay.
   - Daily duplicate prevention preventing double logging on the same day.

3. **Database & Attendance Viewer (`3. View Database & Attendance Logs`)**:
   - Built-in GUI dashboard to inspect, filter, and verify records.
   - View Attendance History and Registered Students in interactive tables.
   - One-click sync to `attendance.xlsx` and `attendance.csv`.
   - Delete student records with automatic file cleanup.

## Setup Instructions

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   *(Note: On Linux/Ubuntu, if pyzbar prompts for a missing C library, run: `sudo apt-get install libzbar0`)*

2. Run the application:
   ```bash
   python main.py
   ```
