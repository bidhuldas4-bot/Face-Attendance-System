"""
FACE ATTENDANCE SYSTEM
======================
A desktop attendance application for Windows.

Technology : Python 3.13, Tkinter (GUI), OpenCV (webcam + face detection/recognition),
             NumPy, Pillow, OpenPyXL (Excel).
Recognition: OpenCV LBPH recognizer (needs the package "opencv-contrib-python").

HOW THE PROGRAM IS ORGANISED (read top to bottom):
    1. Settings ............ passwords, file names, thresholds
    2. Helper functions .... folders, JSON files, buttons, windows
    3. FaceEngine .......... face detection + training + recognition
    4. AttendanceBook ...... reads/writes attendance.xlsx
    5. CameraView .......... shows the webcam INSIDE a Tkinter label
    6. Windows ............. LoginWindow, TeacherDashboard, AddStudentWindow,
                             ManualAttendanceWindow, HistoryWindow
    7. FaceAttendanceApp ... the main window (student attendance)
"""

import json
import os
import re
import sys
import time
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# ---------------------------------------------------------------------------
# Import the extra packages. If one is missing, tell the user clearly.
# ---------------------------------------------------------------------------
try:
    import cv2
    import numpy as np
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    from PIL import Image, ImageTk
except ImportError as import_error:
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "Missing package",
        f"A required package is missing:\n\n{import_error}\n\n"
        "Open Command Prompt in the project folder and run:\n\n"
        "pip install -r requirements.txt",
    )
    sys.exit(1)


# ===========================================================================
# 1. SETTINGS
# ===========================================================================
TEACHER_PASSWORD = "1234"            # <-- change the teacher password HERE only

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CASCADE_FILE = os.path.join(BASE_DIR, "haarcascade_frontalface_default.xml")
KNOWN_FACES_DIR = os.path.join(BASE_DIR, "known_faces")
DATA_DIR = os.path.join(BASE_DIR, "data")
TRAINER_FILE = os.path.join(DATA_DIR, "trainer.yml")
LABELS_FILE = os.path.join(DATA_DIR, "labels.json")
STUDENTS_FILE = os.path.join(DATA_DIR, "students.json")
EXCEL_FILE = os.path.join(BASE_DIR, "attendance.xlsx")

SAMPLES_NEEDED = 20                  # face photos captured per student
CAPTURE_INTERVAL = 0.25              # seconds between two captured samples
FACE_SIZE = (200, 200)               # every face image is resized to this
CONFIDENCE_LIMIT = 70.0              # LBPH distance: LOWER = better match.
#                                      If good students are shown as UNKNOWN, raise it (e.g. 80).
#                                      If wrong people are accepted, lower it (e.g. 55).
STABLE_FRAMES = 3                    # same result needed in N frames in a row
HOLD_SECONDS = 3.0                   # how long the result stays after the face leaves

CAM_W, CAM_H = 520, 420              # size of the main camera area
DATE_FORMAT = "%d-%m-%Y"             # worksheet names, e.g. 29-09-2026
TIME_FORMAT = "%H:%M:%S"

# Colours
BG = "#f2f4f8"
HEADER_BG = "#1f3a5f"
GREEN, RED, ORANGE, GREY, BLUE = "#1e8e3e", "#d93025", "#e37400", "#555555", "#2d6cdf"

CASCADE_HELP = (
    "haarcascade_frontalface_default.xml is missing.\n"
    "Please place it in the project folder:\n"
    f"{BASE_DIR}\n\n"
    "Download it with this command (Command Prompt, inside the project folder):\n\n"
    "curl -L -o haarcascade_frontalface_default.xml "
    "https://raw.githubusercontent.com/opencv/opencv/master/data/haarcascades/"
    "haarcascade_frontalface_default.xml"
)

CONTRIB_HELP = (
    "OpenCV face recognition (cv2.face) is NOT available.\n\n"
    "You need the package 'opencv-contrib-python'.\n"
    "Only ONE OpenCV package may be installed, so run these commands:\n\n"
    "pip uninstall -y opencv-python opencv-python-headless opencv-contrib-python\n"
    "pip install opencv-contrib-python\n\n"
    "Then start the program again."
)


# ===========================================================================
# 2. HELPER FUNCTIONS
# ===========================================================================
def ensure_folders():
    """Create the folders the program needs (does nothing if they exist)."""
    os.makedirs(KNOWN_FACES_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)


def load_json(path, default):
    """Read a JSON file. Return `default` if it is missing or damaged."""
    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, ValueError):
        return default


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=4, ensure_ascii=False)


def load_students():
    """Return the students dictionary: {"ST001": {"name": "Rahul", "label": 1}}"""
    data = load_json(STUDENTS_FILE, {})
    return data if isinstance(data, dict) else {}


def imwrite_safe(path, image):
    """Save a JPG (works even if the folder path has non-English letters)."""
    ok, buffer = cv2.imencode(".jpg", image)
    if ok:
        buffer.tofile(path)
    return ok


def imread_gray_safe(path):
    """Read an image as grayscale (works with any folder path). None on failure."""
    try:
        data = np.fromfile(path, dtype=np.uint8)
        return cv2.imdecode(data, cv2.IMREAD_GRAYSCALE)
    except OSError:
        return None


def make_button(parent, text, command, bg=BLUE, width=16):
    """A coloured button (plain tk.Button so we can colour it on Windows)."""
    return tk.Button(
        parent, text=text, command=command, bg=bg, fg="white",
        activebackground=bg, activeforeground="white",
        disabledforeground="#cccccc", font=("Segoe UI", 10, "bold"),
        width=width, relief="flat", cursor="hand2", padx=6, pady=6,
    )


def build_camera_area(parent, width, height, text):
    """Create the BLACK camera area. Returns (frame, label).
    The webcam picture is later drawn inside the label."""
    frame = tk.Frame(parent, width=width + 4, height=height + 4, bg="black",
                     highlightbackground="#555555", highlightthickness=2)
    frame.pack_propagate(False)
    label = tk.Label(frame, text=text, bg="black", fg="white",
                     font=("Segoe UI", 16, "bold"))
    label.pack(fill="both", expand=True)
    return frame, label


def fit_and_center(window, min_width=0):
    """Size a window to fit its content and put it in the middle of the screen."""
    window.update_idletasks()
    width = max(window.winfo_reqwidth(), min_width)
    height = window.winfo_reqheight()
    x = (window.winfo_screenwidth() - width) // 2
    y = max((window.winfo_screenheight() - height) // 2 - 20, 0)
    window.geometry(f"{width}x{height}+{x}+{y}")


# ===========================================================================
# 3. FACE ENGINE  (detection, training, recognition)
# ===========================================================================
class FaceEngine:
    """Everything that touches OpenCV face detection / recognition."""

    def __init__(self):
        self.cascade = None
        self.cascade_error = None
        self.recognizer = None
        self.recognizer_error = None
        self.trained = False
        self.label_map = {}          # {1: "ST001", 2: "ST002"}

        self._load_cascade()
        self._create_recognizer()
        self.load_model()

    # ---- setup -----------------------------------------------------------
    def _load_cascade(self):
        """Load the LOCAL Haar cascade file (we never use cv2.data.haarcascades)."""
        if not os.path.isfile(CASCADE_FILE):
            self.cascade_error = CASCADE_HELP
            return
        cascade = cv2.CascadeClassifier(CASCADE_FILE)
        if cascade.empty():
            self.cascade_error = (
                "haarcascade_frontalface_default.xml could not be loaded.\n"
                "The file may be damaged. Download it again:\n\n" + CASCADE_HELP
            )
            return
        self.cascade = cascade

    def _create_recognizer(self):
        """Check that cv2.face.LBPHFaceRecognizer_create exists."""
        if hasattr(cv2, "face") and hasattr(cv2.face, "LBPHFaceRecognizer_create"):
            self.recognizer = cv2.face.LBPHFaceRecognizer_create()
        else:
            self.recognizer_error = CONTRIB_HELP

    def problem(self):
        """Return a text describing a setup problem, or None if all is fine."""
        return self.cascade_error or self.recognizer_error

    # ---- model -----------------------------------------------------------
    def load_model(self):
        """Load data/trainer.yml (or rebuild it from the saved photos)."""
        students = load_students()
        self.label_map = {info["label"]: sid for sid, info in students.items()}
        self.trained = False
        if self.recognizer is None or not students:
            return
        if os.path.isfile(TRAINER_FILE):
            try:
                self.recognizer.read(TRAINER_FILE)
                self.trained = True
                return
            except cv2.error:
                pass                      # damaged file -> rebuild below
        self.train()

    def train(self):
        """Train the recognizer from every image in known_faces/.
        Returns (True/False, message)."""
        if self.recognizer is None:
            return False, self.recognizer_error

        students = load_students()
        faces, labels = [], []
        for sid, info in students.items():
            folder = os.path.join(KNOWN_FACES_DIR, sid)
            if not os.path.isdir(folder):
                continue
            for filename in sorted(os.listdir(folder)):
                if not filename.lower().endswith(".jpg"):
                    continue
                image = imread_gray_safe(os.path.join(folder, filename))
                if image is None:
                    continue
                faces.append(cv2.resize(image, FACE_SIZE))
                labels.append(info["label"])

        self.label_map = {info["label"]: sid for sid, info in students.items()}
        save_json(LABELS_FILE, {str(k): v for k, v in self.label_map.items()})

        if not faces:
            self.trained = False
            if os.path.isfile(TRAINER_FILE):
                os.remove(TRAINER_FILE)
            return False, "No face images found to train."

        self.recognizer.train(faces, np.array(labels, dtype=np.int32))
        self.recognizer.write(TRAINER_FILE)
        self.trained = True
        return True, f"Trained with {len(faces)} images of {len(set(labels))} student(s)."

    # ---- detection / recognition -----------------------------------------
    def detect(self, gray):
        """Return a list of (x, y, w, h) faces, biggest face first."""
        if self.cascade is None:
            return []
        found = self.cascade.detectMultiScale(
            gray, scaleFactor=1.2, minNeighbors=5, minSize=(80, 80))
        boxes = [tuple(int(v) for v in box) for box in found]
        boxes.sort(key=lambda b: b[2] * b[3], reverse=True)
        return boxes

    @staticmethod
    def prepare_face(gray, box):
        """Cut the face out, resize it and improve contrast."""
        x, y, w, h = box
        face = cv2.resize(gray[y:y + h, x:x + w], FACE_SIZE)
        return cv2.equalizeHist(face)

    def identify(self, gray, box):
        """Return (student_id, confidence). student_id is None if unknown."""
        if not self.trained:
            return None, 999.0
        face = self.prepare_face(gray, box)
        label, confidence = self.recognizer.predict(face)
        student_id = self.label_map.get(label)
        if student_id is None or confidence > CONFIDENCE_LIMIT:
            return None, confidence
        return student_id, confidence


# ===========================================================================
# 4. ATTENDANCE BOOK  (Excel)
# ===========================================================================
class AttendanceBook:
    """attendance.xlsx: one worksheet per date (named dd-mm-yyyy)."""

    HEADERS = ["Student ID", "Student Name", "Time", "Status", "Method"]
    LOCKED_MSG = ("attendance.xlsx cannot be saved.\n"
                  "Please close it in Excel and try again.")

    def __init__(self, path):
        self.path = path
        self.init_error = self.ensure_file()

    @staticmethod
    def today_name():
        return datetime.now().strftime(DATE_FORMAT)

    def _write_header(self, sheet):
        sheet.append(self.HEADERS)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column, width in zip("ABCDE", (14, 24, 12, 12, 12)):
            sheet.column_dimensions[column].width = width

    def ensure_file(self):
        """Create attendance.xlsx (with today's sheet) if missing. Returns error text or None."""
        if os.path.isfile(self.path):
            return None
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = self.today_name()
            self._write_header(sheet)
            workbook.save(self.path)
            return None
        except PermissionError:
            return self.LOCKED_MSG
        except OSError as error:
            return f"Could not create attendance.xlsx:\n{error}"

    def mark(self, student_id, name, method):
        """Add a Present row for today.
        Returns (status, message) where status is 'marked', 'duplicate', 'locked' or 'error'."""
        error = self.ensure_file()
        if error:
            return "error", error
        try:
            workbook = load_workbook(self.path)
            sheet_name = self.today_name()
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
            else:
                sheet = workbook.create_sheet(sheet_name)   # new day -> new sheet
                self._write_header(sheet)

            # Duplicate check: is this student already in today's sheet?
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None and \
                        str(row[0]).strip().upper() == student_id.upper():
                    return "duplicate", "Attendance already marked for today."

            sheet.append([student_id, name, datetime.now().strftime(TIME_FORMAT),
                          "Present", method])
            workbook.save(self.path)
            return "marked", "Attendance marked successfully."
        except PermissionError:
            return "locked", self.LOCKED_MSG
        except Exception as exc:                      # damaged file etc.
            return "error", f"Could not write attendance:\n{exc}"

    def list_dates(self):
        """Return all worksheet names that are dates, newest first."""
        try:
            workbook = load_workbook(self.path, read_only=True)
            names = list(workbook.sheetnames)
            workbook.close()
        except Exception:
            return []
        dates = []
        for name in names:
            try:
                dates.append((datetime.strptime(name, DATE_FORMAT), name))
            except ValueError:
                pass
        dates.sort(reverse=True)
        return [name for _, name in dates]

    def read_date(self, sheet_name):
        """Return (rows, error). rows is a list of 5-item tuples."""
        try:
            workbook = load_workbook(self.path, read_only=True)
            sheet = workbook[sheet_name]
            rows = []
            for row in sheet.iter_rows(min_row=2, max_col=5, values_only=True):
                if row and row[0] not in (None, ""):
                    rows.append(tuple("" if v is None else v for v in row))
            workbook.close()
            return rows, None
        except PermissionError:
            return None, "attendance.xlsx is locked. Please close it in Excel."
        except Exception as exc:
            return None, f"Could not read attendance:\n{exc}"


# ===========================================================================
# 5. CAMERA VIEW  (webcam shown inside a Tkinter Label)
# ===========================================================================
class CameraView:
    """Opens the webcam and keeps drawing frames into a Tkinter label using after()."""

    def __init__(self, master, label, on_frame, width, height, placeholder):
        self.master = master          # any widget (used for .after)
        self.label = label            # the black label
        self.on_frame = on_frame      # function(frame) -> frame with drawings
        self.width, self.height = width, height
        self.placeholder = placeholder
        self.cap = None
        self.running = False
        self.after_id = None
        self.photo = None
        self.fail_count = 0

    def start(self):
        """Open the camera. Returns True on success, False if it cannot be opened."""
        if self.running:
            return True
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)       # DirectShow = fast on Windows
        if not cap.isOpened():
            cap.release()
            cap = cv2.VideoCapture(0)                  # fallback
        if not cap.isOpened():
            cap.release()
            return False
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 800)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 400)
        self.cap = cap
        self.running = True
        self.fail_count = 0
        self._update()
        return True

    def _update(self):
        """Read one frame, draw it, and schedule the next call."""
        if not self.running:
            return
        ok, frame = self.cap.read()
        if ok:
            self.fail_count = 0
            frame = cv2.flip(frame, 1)                          # mirror view
            frame = cv2.resize(frame, (self.width, self.height))
            try:
                result = self.on_frame(frame)
                if result is not None:
                    frame = result
            except Exception:
                traceback.print_exc()                           # never crash the loop
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = Image.fromarray(frame_rgb)
            self.photo = ImageTk.PhotoImage(image=image)
            try:
                self.label.configure(image=self.photo, text="")
                self.label.image = self.photo                   # keep a reference!
            except tk.TclError:
                self.stop()
                return
        else:
            self.fail_count += 1
            if self.fail_count > 60:                            # camera unplugged?
                self.stop()
                messagebox.showerror(
                    "Camera Error",
                    "Camera could not be opened.\nPlease check your webcam.")
                return
        self.after_id = self.master.after(30, self._update)

    def stop(self):
        """Stop and release the webcam, and show the black placeholder again."""
        self.running = False
        if self.after_id is not None:
            try:
                self.master.after_cancel(self.after_id)
            except tk.TclError:
                pass
            self.after_id = None
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        self.photo = None
        try:
            self.label.configure(image="", text=self.placeholder)
            self.label.image = None
        except tk.TclError:
            pass                                                # window already closed


# ===========================================================================
# 6. WINDOWS
# ===========================================================================
class LoginWindow(tk.Toplevel):
    """Asks for the teacher password."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Teacher Login")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(app.root)

        tk.Label(self, text="TEACHER LOGIN", font=("Segoe UI", 14, "bold"),
                 bg=BG, fg=HEADER_BG).pack(padx=40, pady=(20, 10))
        tk.Label(self, text="Password:", bg=BG, font=("Segoe UI", 10)).pack()
        self.password_var = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.password_var, show="*", width=24)
        self.entry.pack(pady=8)
        self.entry.focus_set()
        self.entry.bind("<Return>", lambda event: self.login())
        make_button(self, "LOGIN", self.login, GREEN, 14).pack(pady=(4, 20))

        fit_and_center(self)
        self.grab_set()

    def login(self):
        if self.password_var.get() == TEACHER_PASSWORD:
            self.destroy()
            self.app.open_dashboard()
        else:
            messagebox.showerror("Login Failed", "Incorrect teacher password.", parent=self)
            self.password_var.set("")
            self.entry.focus_set()


class TeacherDashboard(tk.Toplevel):
    """The teacher's menu."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.child = None                    # the currently open sub-window
        self.title("Teacher Dashboard")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close)

        tk.Label(self, text="TEACHER DASHBOARD", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=HEADER_BG).pack(padx=50, pady=(20, 14))
        make_button(self, "ADD STUDENT", self.open_add_student, BLUE, 30).pack(pady=5)
        make_button(self, "MANUAL ATTENDANCE", self.open_manual, GREEN, 30).pack(pady=5)
        make_button(self, "VIEW ATTENDANCE HISTORY", self.open_history, "#8e44ad", 30).pack(pady=5)
        make_button(self, "BACK", self.close, GREY, 30).pack(pady=(5, 20))
        fit_and_center(self)

    def _open_child(self, window_class):
        """Open a sub-window, but only one at a time."""
        if self.child is not None and self.child.winfo_exists():
            self.child.lift()
            return
        self.child = window_class(self, self.app)

    def open_add_student(self):
        problem = self.app.engine.problem()
        if problem:
            messagebox.showerror("Setup problem", problem, parent=self)
            return
        self._open_child(AddStudentWindow)

    def open_manual(self):
        if not load_students():
            messagebox.showwarning(
                "No students",
                "No students registered.\nTeacher must add a student first.", parent=self)
            return
        self._open_child(ManualAttendanceWindow)

    def open_history(self):
        self._open_child(HistoryWindow)

    def close(self):
        if self.child is not None and self.child.winfo_exists():
            self.child.close()               # releases the camera if needed
        self.destroy()
        self.app.on_dashboard_closed()


class AddStudentWindow(tk.Toplevel):
    """Register a new student: type ID + name, then capture 20 face photos."""

    def __init__(self, dashboard, app):
        super().__init__(dashboard)
        self.app = app
        self.samples = []                    # captured face images (kept in memory)
        self.capturing = False
        self.last_capture = 0.0
        self.title("Add Student")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close)

        # --- form -----------------------------------------------------------
        form = tk.Frame(self, bg=BG)
        form.pack(padx=20, pady=(15, 5))
        tk.Label(form, text="Student ID:", bg=BG, font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="e", padx=5, pady=4)
        tk.Label(form, text="Student Name:", bg=BG, font=("Segoe UI", 10, "bold")).grid(
            row=1, column=0, sticky="e", padx=5, pady=4)
        self.id_var, self.name_var = tk.StringVar(), tk.StringVar()
        self.id_entry = ttk.Entry(form, textvariable=self.id_var, width=30)
        self.id_entry.grid(row=0, column=1, pady=4)
        ttk.Entry(form, textvariable=self.name_var, width=30).grid(row=1, column=1, pady=4)

        # --- camera area ----------------------------------------------------
        self.cam_w, self.cam_h = 480, 360
        cam_frame, self.video_label = build_camera_area(
            self, self.cam_w, self.cam_h, "CAMERA OFF")
        cam_frame.pack(padx=20, pady=8)
        self.camera = CameraView(self, self.video_label, self.process_frame,
                                 self.cam_w, self.cam_h, "CAMERA OFF")

        self.progress_var = tk.StringVar(value="Step 1: enter ID and name, then START CAMERA.")
        tk.Label(self, textvariable=self.progress_var, bg=BG, fg=HEADER_BG,
                 font=("Segoe UI", 10, "bold")).pack()

        # --- buttons --------------------------------------------------------
        buttons = tk.Frame(self, bg=BG)
        buttons.pack(pady=10)
        make_button(buttons, "START CAMERA", self.start_camera, GREEN, 15).pack(side="left", padx=5)
        make_button(buttons, "CAPTURE FACE", self.start_capture, BLUE, 15).pack(side="left", padx=5)
        make_button(buttons, "SAVE STUDENT", self.save_student, "#8e44ad", 15).pack(side="left", padx=5)
        make_button(self, "CLOSE", self.close, GREY, 15).pack(pady=(0, 15))

        fit_and_center(self)
        self.id_entry.focus_set()

    # --- camera -------------------------------------------------------------
    def start_camera(self):
        if not self.camera.start():
            messagebox.showerror(
                "Camera Error",
                "Camera could not be opened.\nPlease check your webcam.", parent=self)
            return
        self.progress_var.set("Step 2: face the camera, then click CAPTURE FACE.")

    def process_frame(self, frame):
        """Called for every camera frame: draw a box and capture samples."""
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = self.app.engine.detect(gray)
        count = len(self.samples)
        if faces:
            box = faces[0]
            x, y, w, h = box
            color = (0, 200, 0) if self.capturing else (0, 255, 255)
            if self.capturing and time.time() - self.last_capture >= CAPTURE_INTERVAL:
                self.samples.append(self.app.engine.prepare_face(gray, box))
                self.last_capture = time.time()
                count = len(self.samples)
                if count >= SAMPLES_NEEDED:
                    self.capturing = False
                    self.progress_var.set(
                        f"Capture complete ({count}/{SAMPLES_NEEDED}). Now click SAVE STUDENT.")
                else:
                    self.progress_var.set(f"Capturing face {count}/{SAMPLES_NEEDED}")
            cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)
        elif self.capturing:
            self.progress_var.set(
                f"No face detected - look at the camera ({count}/{SAMPLES_NEEDED})")
        return frame

    # --- steps --------------------------------------------------------------
    def _read_form(self):
        """Check the ID and name. Returns (id, name) or None."""
        student_id = self.id_var.get().strip().upper()
        name = self.name_var.get().strip()
        if not student_id or not name:
            messagebox.showwarning("Missing data",
                                   "Please enter both Student ID and Student Name.", parent=self)
            return None
        if not re.fullmatch(r"[A-Z0-9_-]+", student_id):
            messagebox.showwarning(
                "Invalid ID",
                "Student ID may only contain letters, numbers, - and _\n(example: ST001).",
                parent=self)
            return None
        return student_id, name

    def start_capture(self):
        if not self.camera.running:
            messagebox.showinfo("Camera off", "Please click START CAMERA first.", parent=self)
            return
        if self._read_form() is None:
            return
        self.samples = []
        self.capturing = True
        self.last_capture = 0.0
        self.progress_var.set(f"Capturing face 0/{SAMPLES_NEEDED} - slowly turn your head a little")

    def save_student(self):
        form = self._read_form()
        if form is None:
            return
        student_id, name = form
        if len(self.samples) < SAMPLES_NEEDED:
            messagebox.showwarning(
                "Not enough photos",
                f"Please capture {SAMPLES_NEEDED} face samples first.\n"
                f"Captured so far: {len(self.samples)}", parent=self)
            return

        students = load_students()
        if student_id in students:
            if not messagebox.askyesno(
                    "Student exists",
                    f"{student_id} is already registered as {students[student_id]['name']}.\n"
                    "Replace the face photos and name?", parent=self):
                return
            label = students[student_id]["label"]
        else:
            label = max((s["label"] for s in students.values()), default=0) + 1

        try:
            # Photos are saved by STUDENT ID (never by name): known_faces/ST001/1.jpg ...
            folder = os.path.join(KNOWN_FACES_DIR, student_id)
            os.makedirs(folder, exist_ok=True)
            for old in os.listdir(folder):
                if old.lower().endswith(".jpg"):
                    os.remove(os.path.join(folder, old))
            for number, image in enumerate(self.samples, start=1):
                imwrite_safe(os.path.join(folder, f"{number}.jpg"), image)

            students[student_id] = {"name": name, "label": label}
            save_json(STUDENTS_FILE, students)
            ok, message = self.app.engine.train()
        except OSError as exc:
            messagebox.showerror("Save error", f"Could not save the student:\n{exc}", parent=self)
            return

        if not ok:
            messagebox.showerror("Training error", message, parent=self)
            return
        messagebox.showinfo("Student saved",
                            f"{name} ({student_id}) was saved.\n{message}", parent=self)
        self.samples = []
        self.id_var.set("")
        self.name_var.set("")
        self.progress_var.set("Saved! Enter another student or click CLOSE.")
        self.id_entry.focus_set()

    def close(self):
        self.capturing = False
        self.camera.stop()
        self.destroy()


class ManualAttendanceWindow(tk.Toplevel):
    """Teacher picks a student and marks him/her present."""

    def __init__(self, dashboard, app):
        super().__init__(dashboard)
        self.app = app
        self.title("Manual Attendance")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.close)

        self.students = load_students()
        self.ids = sorted(self.students)
        values = [f"{sid} - {self.students[sid]['name']}" for sid in self.ids]

        tk.Label(self, text="MANUAL ATTENDANCE", font=("Segoe UI", 14, "bold"),
                 bg=BG, fg=HEADER_BG).pack(padx=40, pady=(20, 10))
        tk.Label(self, text="Select Student:", bg=BG, font=("Segoe UI", 10)).pack()
        self.combo = ttk.Combobox(self, values=values, state="readonly", width=32)
        self.combo.pack(pady=8)
        if values:
            self.combo.current(0)
        make_button(self, "MARK PRESENT", self.mark_present, GREEN, 18).pack(pady=8)
        make_button(self, "CLOSE", self.close, GREY, 18).pack(pady=(0, 20))
        fit_and_center(self, 380)

    def mark_present(self):
        index = self.combo.current()
        if index < 0:
            messagebox.showwarning("No student", "Please select a student.", parent=self)
            return
        student_id = self.ids[index]
        name = self.students[student_id]["name"]
        status, message = self.app.book.mark(student_id, name, "Teacher")
        if status == "marked":
            messagebox.showinfo("Success", message, parent=self)
        elif status == "duplicate":
            messagebox.showwarning("Already marked", message, parent=self)
        else:
            messagebox.showerror("Error", message, parent=self)

    def close(self):
        self.destroy()


class HistoryWindow(tk.Toplevel):
    """Shows the attendance of one chosen date in a table."""

    COLUMNS = ("Student ID", "Student Name", "Time", "Status", "Method")

    def __init__(self, dashboard, app):
        super().__init__(dashboard)
        self.app = app
        self.title("Attendance History")
        self.configure(bg=BG)
        self.protocol("WM_DELETE_WINDOW", self.close)

        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=15, pady=12)
        tk.Label(top, text="Select Date:", bg=BG, font=("Segoe UI", 10, "bold")).pack(side="left")
        self.date_var = tk.StringVar()
        self.combo = ttk.Combobox(top, textvariable=self.date_var, state="readonly", width=16)
        self.combo.pack(side="left", padx=8)
        make_button(top, "VIEW", self.show_date, BLUE, 10).pack(side="left")

        table_frame = tk.Frame(self, bg=BG)
        table_frame.pack(fill="both", expand=True, padx=15)
        self.tree = ttk.Treeview(table_frame, columns=self.COLUMNS, show="headings", height=12)
        for column, width in zip(self.COLUMNS, (100, 180, 90, 90, 90)):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, anchor="center")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.total_var = tk.StringVar(value="Total Present: 0")
        tk.Label(self, textvariable=self.total_var, bg=BG, fg=GREEN,
                 font=("Segoe UI", 13, "bold")).pack(pady=8)
        make_button(self, "CLOSE", self.close, GREY, 14).pack(pady=(0, 15))

        self.load_dates()
        fit_and_center(self)

    def load_dates(self):
        dates = self.app.book.list_dates()
        self.combo["values"] = dates
        if dates:
            self.combo.current(0)
            self.show_date()
        else:
            self.total_var.set("No attendance records yet.")

    def show_date(self):
        sheet_name = self.date_var.get()
        if not sheet_name:
            messagebox.showinfo("No date", "No date selected.", parent=self)
            return
        rows, error = self.app.book.read_date(sheet_name)
        if error:
            messagebox.showerror("Error", error, parent=self)
            return
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", "end", values=row)
        present = sum(1 for row in rows if str(row[3]).strip().lower() == "present")
        self.total_var.set(f"Total Present: {present}")

    def close(self):
        self.destroy()


# ===========================================================================
# 7. MAIN WINDOW  (student attendance)
# ===========================================================================
class FaceAttendanceApp:
    def __init__(self, root):
        self.root = root
        ensure_folders()
        self.engine = FaceEngine()
        self.book = AttendanceBook(EXCEL_FILE)
        self.dashboard = None

        # Recognition state (used to avoid flicker and repeated messages)
        self.candidate_key = None       # who we think we see right now
        self.candidate_count = 0        # in how many frames in a row
        self.confirmed_key = None       # who is already shown/handled
        self.last_face_time = 0.0
        self.marked_date = ""           # students already marked today (cache)
        self.marked_ids = set()

        self.build_ui()
        self.camera = CameraView(root, self.video_label, self.process_frame,
                                 CAM_W, CAM_H, "LIVE CAMERA")
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.refresh_status()
        root.after(300, self.show_startup_problems)

    # ---- building the screen -------------------------------------------------
    def build_ui(self):
        root = self.root
        root.title("Face Attendance System")
        root.configure(bg=BG)
        root.resizable(True, True)
        root.state("zoomed")
        root.bind("<F11>", lambda event: root.attributes("-fullscreen", not root.attributes("-fullscreen")))
        root.bind("<Escape>", lambda event: root.attributes("-fullscreen", False))

        header = tk.Frame(root, bg=HEADER_BG, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="FACE ATTENDANCE SYSTEM", font=("Segoe UI", 18, "bold"),
                 bg=HEADER_BG, fg="white").pack(expand=True)
        self.login_btn = make_button(header, "Teacher Login", self.open_login, "#e67e22", 14)
        self.login_btn.place(relx=1.0, x=-15, rely=0.5, anchor="e")      # top-right corner

        body = tk.Frame(root, bg=BG)
        body.pack(padx=20, pady=12)

        cam_frame, self.video_label = build_camera_area(body, CAM_W, CAM_H, "LIVE CAMERA")
        cam_frame.pack()

        buttons = tk.Frame(body, bg=BG)
        buttons.pack(pady=10)
        self.start_btn = make_button(buttons, "START CAMERA", self.start_camera, GREEN, 18)
        self.start_btn.pack(side="left", padx=8)
        self.stop_btn = make_button(buttons, "STOP CAMERA", self.stop_camera, RED, 18)
        self.stop_btn.pack(side="left", padx=8)
        self.stop_btn.config(state="disabled")

        # Result panel: Student Name / Student ID / Status
        self.name_var = tk.StringVar(value="-")
        self.id_var = tk.StringVar(value="-")
        self.status_var = tk.StringVar(value="-")
        info = tk.Frame(body, bg=BG)
        info.pack(pady=(4, 0))
        self.value_labels = []
        for column, (caption, variable) in enumerate(
                [("Student Name", self.name_var), ("Student ID", self.id_var),
                 ("Status", self.status_var)]):
            tk.Label(info, text=caption, bg=BG, fg=GREY,
                     font=("Segoe UI", 9)).grid(row=0, column=column, padx=25)
            value = tk.Label(info, textvariable=variable, bg=BG, fg=HEADER_BG,
                             font=("Segoe UI", 15, "bold"), width=16)
            value.grid(row=1, column=column, padx=10)
            self.value_labels.append(value)

        self.message_label = tk.Label(body, text="", bg=BG, fg=GREY, wraplength=CAM_W,
                                      height=3, font=("Segoe UI", 11, "bold"))
        self.message_label.pack(pady=(6, 0))

        root.update_idletasks()
        fit_and_center(root, min_width=780)

    def set_result(self, name, student_id, status, message, color):
        """Update the result panel below the camera."""
        self.name_var.set(name)
        self.id_var.set(student_id)
        self.status_var.set(status)
        for label in self.value_labels:
            label.config(fg=color)
        self.message_label.config(text=message, fg=color)

    def refresh_status(self):
        """Show a helpful message depending on whether students exist."""
        if not load_students():
            self.set_result("-", "-", "-",
                            "No students registered yet.\n"
                            "Please use Teacher Login \u2192 Add Student.", ORANGE)
        else:
            self.set_result("-", "-", "-",
                            "Ready. Click START CAMERA to scan your face.", GREY)

    def show_startup_problems(self):
        """Explain missing files/packages as soon as the window is visible."""
        if self.engine.cascade_error:
            messagebox.showerror("Haar cascade missing", self.engine.cascade_error)
        if self.engine.recognizer_error:
            messagebox.showerror("OpenCV face module missing", self.engine.recognizer_error)
        if self.book.init_error:
            messagebox.showerror("Excel error", self.book.init_error)

    # ---- camera buttons -----------------------------------------------------
    def start_camera(self):
        problem = self.engine.problem()
        if problem:
            messagebox.showerror("Setup problem", problem)
            return
        if not self.camera.start():
            messagebox.showerror("Camera Error",
                                 "Camera could not be opened.\nPlease check your webcam.")
            return
        self.reset_recognition_state()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        if not load_students():
            self.refresh_status()
        else:
            self.set_result("-", "-", "Scanning...", "Look at the camera.", GREY)

    def stop_camera(self):
        self.camera.stop()
        self.reset_recognition_state()
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.refresh_status()

    def reset_recognition_state(self):
        self.candidate_key = None
        self.candidate_count = 0
        self.confirmed_key = None

    # ---- recognition (called for every camera frame) -----------------------------
    def process_frame(self, frame):
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = self.engine.detect(gray)
        now = time.time()

        if not faces:
            self.candidate_key, self.candidate_count = None, 0
            if self.confirmed_key is not None and now - self.last_face_time > HOLD_SECONDS:
                self.confirmed_key = None
                if load_students():
                    self.set_result("-", "-", "Scanning...", "Look at the camera.", GREY)
            return frame

        self.last_face_time = now

        # No trained model yet -> only draw yellow boxes.
        if not self.engine.trained:
            for x, y, w, h in faces:
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 255), 2)
            return frame

        # The biggest face is identified; other faces just get a yellow box.
        box = faces[0]
        student_id, _confidence = self.engine.identify(gray, box)
        key = student_id if student_id else "UNKNOWN"
        x, y, w, h = box
        color = (0, 200, 0) if student_id else (0, 0, 255)
        cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)
        cv2.putText(frame, key, (x, max(y - 8, 15)), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)
        for other in faces[1:]:
            ox, oy, ow, oh = other
            cv2.rectangle(frame, (ox, oy), (ox + ow, oy + oh), (0, 255, 255), 2)

        # The same answer must appear in a few frames in a row before we trust it.
        if key == self.candidate_key:
            self.candidate_count += 1
        else:
            self.candidate_key, self.candidate_count = key, 1
        if self.candidate_count >= STABLE_FRAMES and key != self.confirmed_key:
            self.confirmed_key = key
            if key == "UNKNOWN":
                self.show_unknown()
            else:
                self.handle_recognized(key)
        return frame

    def show_unknown(self):
        # Attendance is NOT saved for unknown faces.
        self.set_result(
            "UNKNOWN STUDENT", "-", "Contact Teacher",
            "UNKNOWN STUDENT - Please Contact Teacher.\n"
            "The teacher must mark your attendance manually.", RED)

    def handle_recognized(self, student_id):
        student = load_students().get(student_id)
        name = student["name"] if student else student_id

        today = datetime.now().strftime(DATE_FORMAT)
        if self.marked_date != today:                 # new day -> forget yesterday
            self.marked_date, self.marked_ids = today, set()

        if student_id in self.marked_ids:
            status, message = "duplicate", "Attendance already marked for today."
        else:
            status, message = self.book.mark(student_id, name, "Face")
            if status in ("marked", "duplicate"):
                self.marked_ids.add(student_id)

        if status == "marked":
            self.set_result(name, student_id, "PRESENT",
                            "Attendance recorded successfully.", GREEN)
        elif status == "duplicate":
            self.set_result(name, student_id, "PRESENT",
                            f"{name} is already marked present today.", BLUE)
        else:
            self.set_result(name, student_id, "ERROR", message, RED)

    # ---- teacher area -----------------------------------------------------------
    def open_login(self):
        if self.dashboard is not None and self.dashboard.winfo_exists():
            self.dashboard.lift()
            return
        LoginWindow(self)

    def open_dashboard(self):
        """Called after a correct password. The camera is released for Add Student."""
        if self.camera.running:
            self.stop_camera()
        self.login_btn.config(state="disabled")
        self.start_btn.config(state="disabled")
        self.dashboard = TeacherDashboard(self)

    def on_dashboard_closed(self):
        self.dashboard = None
        self.login_btn.config(state="normal")
        self.start_btn.config(state="normal")
        self.refresh_status()

    # ---- closing ----------------------------------------------------------------
    def on_close(self):
        """Release the webcam and close every window."""
        self.camera.stop()
        if self.dashboard is not None and self.dashboard.winfo_exists():
            self.dashboard.close()
        self.root.destroy()


def main():
    root = tk.Tk()
    FaceAttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
