import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from report_logic import generate_report_pdf, ensure_image_resized
from io import BytesIO
import pandas as pd
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_IMAGE_EXTS = {'png', 'jpg', 'jpeg', 'gif'}
ALLOWED_EXCEL_EXTS = {'xlsx'}

def allowed_file(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def save_uploaded_file(fileobj, subfolder=None):
    if not fileobj:
        return None
    filename = secure_filename(fileobj.filename)
    if filename == '':
        return None
    # add timestamp to avoid collisions
    ts = int(time.time() * 1000)
    name = f"{ts}_{filename}"
    dest_dir = app.config['UPLOAD_FOLDER'] if not subfolder else os.path.join(app.config['UPLOAD_FOLDER'], subfolder)
    os.makedirs(dest_dir, exist_ok=True)
    path = os.path.join(dest_dir, name)
    fileobj.save(path)
    return path

def parse_attendance_excel(path):
    # Use pandas to read first column of first sheet
    try:
        df = pd.read_excel(path, engine='openpyxl')
        # if dataframe has no columns (rare), try to read as values
        if df.shape[1] == 0:
            return []
        # pick first column
        first_col = df.iloc[:, 0].dropna().astype(str).tolist()
        # If header-like values present e.g. "Name", drop it if all rows below are names
        if len(first_col) > 0 and first_col[0].strip().lower() in ['name', 'names', 'participant', 'participant name']:
            first_col = first_col[1:]
        # strip whitespace and filter empties
        cleaned = [s.strip() for s in first_col if s.strip()]
        return cleaned
    except Exception as e:
        print("Error parsing excel:", e)
        return []

def parse_attendance_text(text):
    if not text:
        return []
    # split by newlines or commas
    parts = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # if multiple names in line separated by commas
        if ',' in line:
            for p in line.split(','):
                p = p.strip()
                if p:
                    parts.append(p)
        else:
            parts.append(line)
    return parts

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            data = {}

            # Collect header & sections
            data['header'] = {
                'university': request.form.get('university', '').strip(),
                'school': request.form.get('school', '').strip(),
                'department': request.form.get('department', '').strip()
            }
            data['general_info'] = {
                'Type of Activity': request.form.get('Type of Activity', ''),
                'Title of the Activity': request.form.get('Title of the Activity', ''),
                'Date/s': request.form.get('Date/s', ''),
                'Time': request.form.get('Time', ''),
                'Venue': request.form.get('Venue', ''),
                'Collaboration/Sponsor (if any)': request.form.get('Collaboration/Sponsor (if any)', '')
            }
            data['speaker_details'] = {
                'Name': request.form.get('Name', ''),
                'Title/Position': request.form.get('Title/Position', ''),
                'Organization': request.form.get('Organization', ''),
                'Title of Presentation': request.form.get('Title of Presentation', '')
            }
            data['participants_profile'] = {
                'Type of Participants': request.form.get('Type of Participants', ''),
                'No. of Participants': request.form.get('No. of Participants', '')
            }
            data['synopsis'] = {
                'highlights': request.form.get('highlights', ''),
                'key_takeaways': request.form.get('key_takeaways', ''),
                'summary': request.form.get('summary', ''),
                'follow_up_plan': request.form.get('follow_up_plan', '')
            }
            data['report_prepared_by'] = {
                'Name': request.form.get('Name_Prepared', ''),
                'Designation/Title': request.form.get('Designation/Title_Prepared', '')
            }
            data['speaker_profile'] = {
                'profile_text': request.form.get('profile_text', '')
            }
            data['photos'] = {
                'caption': request.form.get('caption', '')
            }

            # Speaker image
            speaker_file = request.files.get('speaker_image')
            if speaker_file and allowed_file(speaker_file.filename, ALLOWED_IMAGE_EXTS):
                p = save_uploaded_file(speaker_file, subfolder='speaker')
                # resize and return path to resized image
                p2 = ensure_image_resized(p)
                data['speaker_profile']['image_path'] = p2
            else:
                data['speaker_profile']['image_path'] = None

            # Activity photos (up to 5)
            photo_paths = []
            for i in range(1, 6):
                f = request.files.get(f'photo_{i}')
                if f and allowed_file(f.filename, ALLOWED_IMAGE_EXTS):
                    p = save_uploaded_file(f, subfolder='photos')
                    p2 = ensure_image_resized(p)
                    photo_paths.append(p2)
            data['photos']['image_paths'] = photo_paths

            # Attendance: manual text + excel (excel overrides if provided)
            attendance_manual = request.form.get('attendance_text', '')
            attendance_list = parse_attendance_text(attendance_manual)

            attendance_excel = request.files.get('attendance_excel')
            if attendance_excel and allowed_file(attendance_excel.filename, ALLOWED_EXCEL_EXTS):
                excel_path = save_uploaded_file(attendance_excel, subfolder='attendance')
                parsed = parse_attendance_excel(excel_path)
                if parsed:
                    attendance_list = parsed  # excel takes precedence if parsed names exist

            data['attendance'] = attendance_list

            # Flyer, Approval, Impact - images only
            flyer = request.files.get('flyer')
            data['flyer_path'] = None
            if flyer and allowed_file(flyer.filename, ALLOWED_IMAGE_EXTS):
                p = save_uploaded_file(flyer, subfolder='attachments')
                data['flyer_path'] = ensure_image_resized(p)

            approval = request.files.get('approval_letter')
            data['approval_path'] = None
            if approval and allowed_file(approval.filename, ALLOWED_IMAGE_EXTS):
                p = save_uploaded_file(approval, subfolder='attachments')
                data['approval_path'] = ensure_image_resized(p)

            impact = request.files.get('impact_analysis_report')
            data['impact_path'] = None
            if impact and allowed_file(impact.filename, ALLOWED_IMAGE_EXTS):
                p = save_uploaded_file(impact, subfolder='attachments')
                data['impact_path'] = ensure_image_resized(p)

            # Feedback screenshots up to 5
            fb_paths = []
            for i in range(1, 6):
                f = request.files.get(f'feedback_ss_{i}')
                if f and allowed_file(f.filename, ALLOWED_IMAGE_EXTS):
                    p = save_uploaded_file(f, subfolder='feedback')
                    fb_paths.append(ensure_image_resized(p))
            data['feedback_screenshots'] = fb_paths

            # Generate PDF bytes
            pdf_bytes = generate_report_pdf(data)

            return send_file(
                BytesIO(pdf_bytes),
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f"{data['general_info'].get('Title of the Activity','Activity')}_Report.pdf"
            )

        except Exception as exc:
            print("Exception in POST:", exc)
            return render_template('index.html', error=str(exc))

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
