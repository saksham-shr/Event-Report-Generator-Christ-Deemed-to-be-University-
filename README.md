# Activity Report Generator

A **Flask-based web application** that automates the process of generating structured academic **Activity Reports**.  
Users can input event details, upload images, attendance lists, and generate professional **PDF reports** suitable for institutional use.

---

## 🚀 Features

- 📋 Collects structured event and department information  
- 🖼️ Upload flyers, approvals, and activity photos  
- 🧾 Accept attendance manually or through Excel files (`.xlsx`)  
- 📄 Automatically generate professional PDF reports using ReportLab  
- 🧠 Clean and responsive HTML5 + CSS3 interface (Jinja2 templates)  
- ☁️ Easy deployment on AWS EC2 or Elastic Beanstalk  

---

## 🏗️ Tech Stack

| Component | Technology |
|------------|-------------|
| Backend | Flask (Python) |
| Frontend | HTML5, CSS3 (Jinja2 Templates) |
| File Handling | Pandas, OpenPyXL |
| PDF Generation | ReportLab |
| Deployment | Gunicorn + Nginx (AWS EC2) |

---

## ⚙️ Installation & Setup

Follow these steps to set up the project locally:

```bash
git clone https://github.com/<your-username>/report-generator.git
cd report-generator

python -m venv venv
Activate the virtual environment:

For Windows:
venv\Scripts\activate

For Mac/Linux:
source venv/bin/activate
Install Dependencies
pip install -r requirements.txt

Run the Flask App
python app.py
