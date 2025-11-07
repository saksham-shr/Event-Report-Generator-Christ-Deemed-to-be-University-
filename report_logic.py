import os
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib import colors
from PIL import Image as PILImage

# Config
MAX_IMAGE_WIDTH_INCH = 6.6  # max width in inches for images in PDF
MAX_IMAGE_PX = 1200         # max pixel width we will resize images to (approx)
UPLOADS_DIR = "uploads"

styles = getSampleStyleSheet()
BASE_FONT = "Times-Roman"
BOLD_FONT = "Times-Bold"

# Paragraph styles
styles.add(ParagraphStyle(name='HeaderMain', fontName=BOLD_FONT, fontSize=16, alignment=1, spaceAfter=8))
styles.add(ParagraphStyle(name='HeaderSub', fontName=BASE_FONT, fontSize=10, alignment=1, spaceAfter=4))
styles.add(ParagraphStyle(name='ReportTitle', fontName=BOLD_FONT, fontSize=12, alignment=1, spaceAfter=12))
styles.add(ParagraphStyle(name='SectionTitle', fontName=BOLD_FONT, fontSize=10, alignment=0, spaceBefore=14, spaceAfter=6))
styles.add(ParagraphStyle(name='NormalText', fontName=BASE_FONT, fontSize=10, leading=14))
styles.add(ParagraphStyle(name='TableKey', fontName=BOLD_FONT, fontSize=10, leading=12))
styles.add(ParagraphStyle(name='TableValue', fontName=BASE_FONT, fontSize=10, leading=12))

def ensure_image_resized(path):
    """
    Resize image if it's very large. Save a resized copy with suffix _resized
    and return the resized path (or original path if small).
    """
    try:
        img = PILImage.open(path)
        w, h = img.size
        if w > MAX_IMAGE_PX:
            ratio = MAX_IMAGE_PX / float(w)
            new_w = int(w * ratio)
            new_h = int(h * ratio)
            img = img.resize((new_w, new_h), PILImage.LANCZOS)
            base, ext = os.path.splitext(path)
            new_path = f"{base}_resized{ext}"
            img.save(new_path, quality=85)
            return new_path
        else:
            return path
    except Exception as e:
        print("Image resize error:", e)
        return path

def make_table_from_dict(dct, colWidths=[2.2*inch, 4.8*inch]):
    """
    Create a Table flowable from a dictionary with consistent styling.
    """
    if not dct:
        return []
    data = []
    for k, v in dct.items():
        keyp = Paragraph(str(k), styles['TableKey'])
        valp = Paragraph(str(v) if v is not None else '', styles['TableValue'])
        data.append([keyp, valp])
    tbl = Table(data, colWidths=colWidths, hAlign='LEFT')
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.75, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTNAME', (0,0), (-1,-1), BASE_FONT),
        ('FONTNAME', (0,0), (0,-1), BOLD_FONT),
        ('LINEBEFORE', (0,0), (-1,-1), 0, colors.white),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    return [tbl, Spacer(1, 0.08*inch)]

def participant_table(names):
    """
    Create a table with a single column 'Participant Name' with each name on its own row.
    """
    if not names:
        return [Paragraph("No attendance records provided.", styles['NormalText']), Spacer(1, 0.08*inch)]
    data = [[Paragraph("Participant Name", styles['TableKey'])]]
    for n in names:
        data.append([Paragraph(n, styles['TableValue'])])
    tbl = Table(data, colWidths=[6.8*inch], hAlign='LEFT')
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.75, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f0f0f0")),
        ('ALIGN', (0,0), (-1,0), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), BOLD_FONT),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    return [tbl, Spacer(1, 0.12*inch)]

def image_flowable(path, max_width_inch=MAX_IMAGE_WIDTH_INCH):
    """
    Return a ReportLab Image flowable scaled to the given max width (keeps aspect ratio).
    """
    if not path or not os.path.exists(path):
        return Paragraph(f"[Image not found: {path}]", styles['NormalText'])
    try:
        # ReportLab uses points; 1 inch = 72 points
        max_w_pts = max_width_inch * 72
        # We can let ReportLab scale if we supply width only
        img = Image(path)
        iw, ih = img.drawWidth, img.drawHeight
        # if width already larger than allowed, scale
        if iw > max_w_pts:
            scale = max_w_pts / iw
            img.drawWidth = iw * scale
            img.drawHeight = ih * scale
        # center
        img.hAlign = 'CENTER'
        return img
    except Exception as e:
        print("Error creating image flowable:", e)
        return Paragraph(f"[Failed to load image: {path}]", styles['NormalText'])

def generate_report_pdf(data):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=0.7*inch, rightMargin=0.7*inch,
                            topMargin=0.7*inch, bottomMargin=0.7*inch)
    story = []

    # Header
    hdr = data.get('header', {})
    university = hdr.get('university', 'CHRIST (Deemed to be University), Bangalore')
    school = hdr.get('school', '')
    department = hdr.get('department', '')

    story.append(Paragraph(university, styles['HeaderMain']))  # 16pt bold centered
    if school:
        story.append(Paragraph(school, styles['HeaderSub']))
    if department:
        story.append(Paragraph(department, styles['HeaderSub']))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("Activity Report", styles['ReportTitle']))

    # General Info
    story.append(Paragraph("General Information", styles['SectionTitle']))
    story.extend(make_table_from_dict(data.get('general_info', {})))

    # Speaker details
    story.append(Paragraph("Speaker/Guest/Presenter Details", styles['SectionTitle']))
    story.extend(make_table_from_dict(data.get('speaker_details', {})))

    # Participants profile
    story.append(Paragraph("Participants Profile", styles['SectionTitle']))
    story.extend(make_table_from_dict(data.get('participants_profile', {})))

    # Synopsis
    story.append(Paragraph("Synopsis of the Activity (Description)", styles['SectionTitle']))
    synopsis = data.get('synopsis', {})
    synopsis_table = {
        "Highlights of the Activity (Description)": synopsis.get('highlights', ''),
        "Key Takeaways": synopsis.get('key_takeaways', ''),
        "Summary of the Activity": synopsis.get('summary', ''),
        "Follow-up plan": synopsis.get('follow_up_plan', '')
    }
    story.extend(make_table_from_dict(synopsis_table,))

    # Report prepared by
    story.append(Paragraph("Report Prepared By", styles['SectionTitle']))
    story.extend(make_table_from_dict(data.get('report_prepared_by', {})))

    # Speaker profile (text + image)
    sp = data.get('speaker_profile', {})
    if sp and (sp.get('profile_text') or sp.get('image_path')):
        story.append(Paragraph("Speaker Profile", styles['SectionTitle']))
        if sp.get('profile_text'):
            story.append(Paragraph(sp.get('profile_text', ''), styles['NormalText']))
            story.append(Spacer(1, 0.08*inch))
        if sp.get('image_path'):
            story.append(image_flowable(sp.get('image_path')))
            story.append(Spacer(1, 0.12*inch))

    # Photos of Activity
    photos = data.get('photos', {})
    photo_list = photos.get('image_paths', [])
    story.append(Paragraph("Photos of the Activity", styles['SectionTitle']))
    if photo_list:
        for p in photo_list:
            story.append(image_flowable(p))
            story.append(Spacer(1, 0.12*inch))
    else:
        story.append(Paragraph("No photos provided.", styles['NormalText']))
        story.append(Spacer(1, 0.08*inch))

    # Attendance List (neat table)
    story.append(Paragraph("Attendance List", styles['SectionTitle']))
    attendance = data.get('attendance', [])
    story.extend(participant_table(attendance))

    # Flyer of the Event
    story.append(Paragraph("Flyer of the Event", styles['SectionTitle']))
    flyer_path = data.get('flyer_path')
    if flyer_path:
        story.append(image_flowable(flyer_path))
        story.append(Spacer(1, 0.12*inch))
    else:
        story.append(Paragraph("No flyer uploaded.", styles['NormalText']))
        story.append(Spacer(1, 0.08*inch))

    # Approval Letter
    story.append(Paragraph("Approval Letter", styles['SectionTitle']))
    approval_path = data.get('approval_path')
    if approval_path:
        story.append(image_flowable(approval_path))
        story.append(Spacer(1, 0.12*inch))
    else:
        story.append(Paragraph("No approval letter uploaded.", styles['NormalText']))
        story.append(Spacer(1, 0.08*inch))

    # Feedback Screenshots (placed before Impact Analysis)
    story.append(Paragraph("Feedback Screenshots", styles['SectionTitle']))
    fb_list = data.get('feedback_screenshots', [])
    if fb_list:
        for p in fb_list:
            story.append(image_flowable(p))
            story.append(Spacer(1, 0.12*inch))
    else:
        story.append(Paragraph("No feedback screenshots uploaded.", styles['NormalText']))
        story.append(Spacer(1, 0.08*inch))

    # Impact Analysis Report
    story.append(Paragraph("Impact Analysis Report", styles['SectionTitle']))
    impact_path = data.get('impact_path')
    if impact_path:
        story.append(image_flowable(impact_path))
        story.append(Spacer(1, 0.12*inch))
    else:
        story.append(Paragraph("No impact analysis report uploaded.", styles['NormalText']))
        story.append(Spacer(1, 0.08*inch))

    # Build PDF
    doc.build(story)
    pdf = buf.getvalue()
    buf.close()
    return pdf
