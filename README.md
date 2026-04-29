AI-Based Handwriting Recognition System 🖋️

🏆 Gold Award Winner** - Electronic Engineering Department FYP Innovation Competition (2025)

What is this?
I built this application for my Final Year Project (FYP) to solve a very simple problem: turning physical notes into digital text quickly. It is a real-time handwriting recognition desktop app. 
I was incredibly proud that this project won the Gold Award for innovation at my polytechnic. It was my first deep dive into bridging hardware, computer vision, and a user-friendly interface into one working, practical tool.

The Tech Stack
Language:Python
Interface:PyQt5 (I used this to keep the desktop GUI clean and straightforward)
Computer Vision:OpenCV (Handles all the heavy lifting for image prep—like noise reduction and thresholding—before the OCR even tries to read it)
Text Extraction:Tesseract OCR *(The core engine that actually maps and reads the characters)
Database:SQLite (Simple, lightweight local storage so users don't lose their scanned notes)

How It Works
* Live Text Conversion:It scans handwritten notes and instantly converts them into editable digital text.
* Under-the-Hood Image Cleanup:Before Tesseract even looks at the handwriting, the app runs an OpenCV pipeline to clean up background noise and fix the contrast. This makes the accuracy 'way'better.
* Built-In Note Storage:Everything you scan can be saved, searched, and managed right inside the app via the local database (`notes.db`).

⚙️ Want to run it locally?
1. Make sure you have Python 3.x installed.
2. Install the required libraries:
   `pip install PyQt5 opencv-python pytesseract`
3. *Important:* You need to have the Tesseract-OCR engine installed on your actual machine and added to your system PATH for it to work.
4. Launch the app:
   `python main.py`
