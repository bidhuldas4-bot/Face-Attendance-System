# Face Attendance System

A Windows desktop app (Tkinter + OpenCV + Excel) that marks student attendance with face recognition.

## Project folder

```text
face_attendance/
├── main.py                                  <- the program
├── requirements.txt                         <- packages to install
├── README.md                                <- this file
├── haarcascade_frontalface_default.xml      <- YOU place this here (see step 3)
│
├── known_faces/      (created automatically)  face photos, one folder per Student ID
│   └── ST001/1.jpg ... 20.jpg
├── data/             (created automatically)
│   ├── trainer.yml       trained face model
│   ├── labels.json       label number -> Student ID
│   └── students.json     Student ID -> name + label
└── attendance.xlsx   (created automatically)  one worksheet per date
```

## Installation (Windows, Command Prompt)

**1. Go to the project folder** (change the path to where YOU put it):

```cmd
cd "C:\Users\User\OneDrive\Microsoft Copilot Chat Files\Documents\face attence"
```

> Tip: a folder path with only English letters/numbers is safest for OpenCV.

**2. Create and activate a virtual environment:**

```cmd
python -m venv venv
venv\Scripts\activate
```

**3. Download the Haar cascade file into the same folder as `main.py`:**

```cmd
curl -L -o haarcascade_frontalface_default.xml https://raw.githubusercontent.com/opencv/opencv/master/data/haarcascades/haarcascade_frontalface_default.xml
```

PowerShell alternative:

```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/opencv/opencv/master/data/haarcascades/haarcascade_frontalface_default.xml" -OutFile "haarcascade_frontalface_default.xml"
```

**4. Install the packages:**

```cmd
pip install -r requirements.txt
```

**5. Check that face recognition is available** (must print `True`):

```cmd
python -c "import cv2; print(cv2.__version__, hasattr(cv2, 'face') and hasattr(cv2.face, 'LBPHFaceRecognizer_create'))"
```

If it prints `False` (or errors), you have the wrong OpenCV package. Only **one** OpenCV package may be installed, and it must be `opencv-contrib-python`:

```cmd
pip uninstall -y opencv-python opencv-python-headless opencv-contrib-python
pip install opencv-contrib-python
```

**6. Run:**

```cmd
python main.py
```

Next time, you only need: activate (`venv\Scripts\activate`) and `python main.py`.

## How to use

### Add the first student (teacher)
1. Click **Teacher Login** (top-right), enter the password (`1234`).
2. Click **ADD STUDENT**.
3. Type Student ID (e.g. `ST001`) and Name (e.g. `Rahul`).
4. Click **START CAMERA**, then **CAPTURE FACE**. Look at the camera and slowly turn your head a little. The window shows *Capturing face 1/20 ... 20/20* (it only captures when a face is detected).
5. Click **SAVE STUDENT**. Photos go to `known_faces/ST001/` and the model is trained automatically.

### Student attendance
1. On the main screen click **START CAMERA**.
2. Recognised student: the name, ID and **PRESENT** are shown and one row is added to Excel (Method = *Face*).
3. Already marked today: *"Rahul is already marked present today."* (no second row).
4. Unknown face: *UNKNOWN STUDENT - Please Contact Teacher.* Nothing is saved.

### Teacher manual attendance
Teacher Login → **MANUAL ATTENDANCE** → choose `ST001 - Rahul` → **MARK PRESENT**. Saved with Method = *Teacher*. Duplicates for the same day are blocked.

### Excel history
* `attendance.xlsx` has **one worksheet per date** (`29-09-2026`, `30-09-2026`, ...), created automatically.
* Columns: Student ID, Student Name, Time, Status, Method.
* In the app: Teacher Login → **VIEW ATTENDANCE HISTORY** → pick a date → **VIEW** (shows *Total Present*).
* **Close the Excel file** before marking attendance, otherwise Windows blocks saving.

## Settings (top of `main.py`)

| Setting | Meaning |
|---|---|
| `TEACHER_PASSWORD` | teacher password (only place it is stored) |
| `SAMPLES_NEEDED` | face photos per student (20) |
| `CONFIDENCE_LIMIT` | recognition strictness. Lower = stricter. Good students shown as UNKNOWN → raise (e.g. 80). Wrong people accepted → lower (e.g. 55) |

## Troubleshooting

| Message | Fix |
|---|---|
| haarcascade_frontalface_default.xml is missing | Do step 3 |
| OpenCV face module missing | Do step 5 (install `opencv-contrib-python`) |
| Camera could not be opened | Close Zoom/Teams/Camera app, check Windows camera privacy settings |
| Recognition is poor | Add the student in good, even light; make sure the face is fully visible; re-add the student (same ID replaces old photos) |

## Notes
* Capturing and recognising works best with the same lighting and camera position.
* On-video text shows the Student ID only (OpenCV cannot draw all languages); the full name appears in the panel below the camera.
* Face recognition is not perfect. Keep **Manual Attendance** as the fallback for the teacher.
