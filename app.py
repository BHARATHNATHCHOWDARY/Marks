from flask import Flask, render_template, request, redirect, url_for, flash
import os
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)
app.secret_key = '9347611553' # Necessary for flash messages

allowed_emails = ['bharad008@rmkcet.ac.in', 'barathnathchowdary@gmail.com']

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']

        if email in allowed_emails and password == os.environ.get('REQUIRED_PASSWORD'):
            return redirect(url_for('upload'))
        else:
            flash('Invalid email or password.')
            return redirect(url_for('login'))

    return render_template('login.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)

        df = pd.read_excel(file_path)
        for index, row in df.iterrows():
            name = row['Name']
            maths_marks = row['Maths Marks']
            daa_marks = row['DAA Marks']
            physics_marks = row['Physics Marks']
            chemistry_marks = row['Chemistry Marks']
            cs_marks = row['Computer Science Marks']
            english_marks = row['English Marks']
            counselor_email = row['Counselor Email']
            college_reg_no = row['College Registration Number']
            percentage = row['Percentage']

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt=f"Report for {name}", ln=True, align='C')
            pdf.ln(10)
            pdf.cell(200, 10, txt=f"Maths Marks: {maths_marks}", ln=True)
            pdf.cell(200, 10, txt=f"DAA Marks: {daa_marks}", ln=True)
            pdf.cell(200, 10, txt=f"Physics Marks: {physics_marks}", ln=True)
            pdf.cell(200, 10, txt=f"Chemistry Marks: {chemistry_marks}", ln=True)
            pdf.cell(200, 10, txt=f"Computer Science Marks: {cs_marks}", ln=True)
            pdf.cell(200, 10, txt=f"English Marks: {english_marks}", ln=True)
            pdf.cell(200, 10, txt=f"Percentage: {percentage}", ln=True)
            pdf.cell(200, 10, txt=f"College Registration Number: {college_reg_no}", ln=True)

            pdf_file_path = os.path.join('reports', f"{name}_report.pdf")
            pdf.output(pdf_file_path)

            from_email = os.environ.get('SENDER_EMAIL')
            email_password = os.environ.get('EMAIL_PASSWORD')
            to_email = counselor_email
            subject = f"Report for {name}"
            body = f"Please find the report for {name} attached."
            send_email(from_email, to_email, subject, body, pdf_file_path, email_password)

        flash('Reports generated and sent successfully.')
        return redirect(url_for('upload'))

    return render_template('upload.html')

def send_email(from_email, to_email, subject, body, attachment, email_password):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    attachment_name = os.path.basename(attachment)
    with open(attachment, "rb") as attach_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attach_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {attachment_name}")
        msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, email_password)
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('reports', exist_ok=True)
    app.run(debug=True)
