from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file,
)  # Server side Code Writing and Backend processing
import os  # Path Operations
import json  # JSON Operations
import uuid  # Generate Unique Session ID
import zipfile  # Create reports Zip File
import time, threading, schedule  # Server Cleanup tools
import smtplib  # Mailer
from email.mime.multipart import MIMEMultipart  # Send Mail Features
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import datetime
from email import encoders  # To add encoding to mail attachments
import pandas as pd  # To Read Excel files
from reportlab.lib.pagesizes import letter  # To create a new document with a template
from reportlab.lib import colors  # For Colors
from reportlab.lib.styles import ParagraphStyle  # To add paragraph styling
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Image,
    Spacer,
)  # Reportlab Utility for customization

mail_id = ""  # Mail ID variable

mail_pass = ""  # Password Variable

receiver_address = ""  # Mail Receiver address

usn_start_var = ""  # Vaeiable for USN start

usn_end_var = ""  # Variable for  USN end

session_id = ""

# Initaiate Flask Application
app = Flask(__name__, static_url_path="/static", static_folder="static")


# Create Uploads directory if it doesnt already exist
uploads_dir = os.path.join(os.getcwd(), "uploads")
if not os.path.exists(uploads_dir):
    os.makedirs(uploads_dir)

xlsx_file_path = ""  # Global XLSX file path

data_frame = ""  # Global Data Frame Variable

data_pdf = ""  # Global Data_PDF variable

reports_dir = ""


# Function to clear the contents of both "reports" and "uploads" folders
def clear_folders():
    temp_zip_dir = os.path.join(os.getcwd(), "temp_zip")
    reports_dir = os.path.join(os.getcwd(), "reports")
    uploads_dir = os.path.join(os.getcwd(), "uploads")

    # Remove all files in the reports directory
    for root, dirs, files in os.walk(reports_dir, topdown=False):
        for file in files:
            os.remove(os.path.join(root, file))
        for dir in dirs:
            os.rmdir(os.path.join(root, dir))

    for root, dirs, files in os.walk(temp_zip_dir, topdown=False):
        for file in files:
            os.remove(os.path.join(root, file))
        for dir in dirs:
            os.rmdir(os.path.join(root, dir))

    # Remove all files in the uploads directory
    for file in os.listdir(uploads_dir):
        file_path = os.path.join(uploads_dir, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")


# Clears uploads, reports and temp_zip folder when program is run
clear_folders()


# Reportlab Function to create paragraph
def create_paragraph(text, style):
    return Paragraph(text, style)


# Route to send collection of zip file of reports
@app.route("/send-zip", methods=["POST"])
def send_zip():
    data = json.loads(request.data.decode("utf-8"))  # Unpack JSON stringified Data

    zip_path = data.get("zipFilePath")

    response = send_file(
        zip_path,
        as_attachment=True,
        mimetype="application/zip",
        download_name="Reports.zip",
    )

    # Send data
    return response


# Route to send mails when button is clicked
# Returns success or Failure of certain USN and all mails after that USN will not be sent
@app.route("/send-mail", methods=["GET"])
def send_mail():
    global mail_id, mail_pass, usn_start_var, usn_end_var, data_frame, data_pdf

    if mail_id == "" or mail_pass == "":
        return jsonify(
            {
                "error": "Empty Fields",
                "Message": "Please Set your mail ID and Password",
            }
        )

    if (
        len(data_frame.index[data_frame["USN"] == usn_start_var].tolist()) == 0
        or data_frame.index[data_frame["USN"] == usn_end_var].tolist() == 0
        or usn_start_var == ""
        or usn_end_var == ""
    ):
        return jsonify(
            {
                "error": "USN Error",
                "Message": "Please enter Valid USN range and try again",
            }
        )
    usn_start_integer = data_frame.index[data_frame["USN"] == usn_start_var].tolist()[0]
    usn_end_integer = data_frame.index[data_frame["USN"] == usn_end_var].tolist()[0]
    for i in range(usn_start_integer, usn_end_integer + 1):
        receiver_address = data_pdf[i + 1][3:4][0]

        mail_content = """Hello,
        Please find attached the report for your ward's 
        performance and attendance.
        Regards,
        JSSATEB 
        """
        try:
            message = MIMEMultipart()
            message["From"] = mail_id
            message["To"] = receiver_address
            message["Subject"] = f"Student Academic Report"

            # Attachments
            message.attach(MIMEText(mail_content, "plain"))
            attach_file_name = (
                f"reports/Reports - {session_id}/Report - {data_pdf[i + 1][1:2][0]}.pdf"
            )
            attach_file = open(attach_file_name, "rb")
            payload = MIMEBase("application", "pdf")
            payload.add_header(
                "Content-Disposition", f"attachment; filename={attach_file_name}"
            )  # Set the correct filename
            payload.set_payload((attach_file).read())
            encoders.encode_base64(payload)

            payload.add_header(
                "Content-Decomposition", "attachment", filename=attach_file_name
            )
            message.attach(payload)

            # Create SMTP session for sending the mail
            session = smtplib.SMTP("smtp.gmail.com", 587)  # Port 587 Gmail
            session.starttls()  # enable security
            session.login(mail_id, mail_pass)  # Session Login
            text = message.as_string()
            session.sendmail(mail_id, receiver_address, text)
            session.quit()

        except smtplib.SMTPAuthenticationError as e:
            error_message = str(e)
            if "Username and Password not accepted" in error_message:
                return jsonify(
                    {
                        "error": "Email & Password not set",
                        "Message": "Please update the mail ID and Password",
                    }
                )
        except FileNotFoundError as e:
            return jsonify(
                {
                    "error": "File Error",
                    "Message": f"File for USN {data_pdf[i + 1][1:2][0]} does not exist. Mails for files before this USN have been sent",
                }
            )
        except smtplib.SMTPRecipientsRefused as e:
            return jsonify(
                {
                    "error": "Mail Error",
                    "Message": f"Mail ID for USN {data_pdf[i + 1][1:2][0]} is incorrect/does not exist. All mails prior to this USN have been sent",
                }
            )
    return jsonify(
        {"operation": "Mail Successful", "Message": "All Mails Successfully sent"}
    )


# To make separate PDFS for all students.
def make_all_pdf(data_frame, usn_start_integer, usn_end_integer, data_pdf):
    for i in range(int(usn_start_integer), int(usn_end_integer) + 1):
        current_name = data_frame.iloc[i, 2]
        current_usn = data_frame.iloc[i, 1]
        make_pdf(
            f"reports/Reports - {session_id}/Report - {current_usn}.pdf",
            data_pdf,
            current_name,
            current_usn,
            i,
        )


# PDF Generation Starts here:
def make_pdf(filename, data_pdf, student_name, student_usn, current_stud):
    if usn_start_var == "" or usn_end_var == "":
        return jsonify(
            {
                "operation": "Incorrect USN format",
                "Message": "Enter USN numbers properly and try again.",
            }
        )

    doc = SimpleDocTemplate(
        filename, pagesize=letter, topMargin=0, leftMargin=50, bottomMargin=0
    )
    story = []

    top_heading_style = ParagraphStyle(
        name="Heading",
        fontWeight="bold",
        fontName="Helvetica",
        fontSize=20,
        textColor=colors.black,
        spaceAfter=20,
        alignment=1,
        leading=25,
    )

    main_heading_style = ParagraphStyle(
        name="Heading-2",
        fontWeight="bold",
        fontName="Helvetica",
        fontSize=16,
        textColor=colors.black,
        spaceAfter=20,
        alignment=0,
        leading=25,
    )

    normal_style = ParagraphStyle(
        name="Normal-Text",
        fontWeight="bold",
        fontName="Times-Roman",
        fontSize=14,
        textColor=colors.black,
        spaceAfter=20,
        alignment=0,
    )

    normal_table_style = TableStyle(
        [
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("FONTNAME", (0, 0), (-1, -1), "Times-Roman"),
            ("FONTSIZE", (0, 0), (-1, -1), 12),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ]
    )

    jss_logo = "./static/images/JSS_Logo.png"  # College Logo
    college_logo = Image(
        jss_logo, width=60, height=60
    )  # Adjusting width and height as needed

    college_header = "JSS ACADEMY OF TECHNICAL EDUCATION, BANGALORE"
    college_header_paragraph = create_paragraph(college_header, top_heading_style)

    table = [[college_logo, college_header_paragraph]]
    col_widths = [60, None]
    college_header_table = Table(table, colWidths=col_widths, rowHeights=[100])
    table_style = TableStyle(
        [
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ]
    )
    college_header_table.setStyle(table_style)
    story.append(college_header_table)

    # Add a heading paragraph
    heading_text = "<u>ACADEMIC REPORT:</u>"
    heading_paragraph = create_paragraph(heading_text, main_heading_style)
    story.append(heading_paragraph)

    stud_det_text = "<u>Student Details:</u>"
    stud_det_paragraph = create_paragraph(stud_det_text, normal_style)
    story.append(stud_det_paragraph)

    today = datetime.date.today().strftime("%d %B, %Y")

    student_details = {
        "Name:": student_name,
        "USN:": student_usn,
        "Date:": today,
    }
    stud_details = [var for var in student_details.items()]
    stud_details_table = Table([var for var in stud_details], hAlign="LEFT")
    stud_details_table.setStyle(normal_table_style)

    story.append(stud_details_table)
    spacer = Spacer(1, 15)
    story.append(spacer)

    header = data_pdf[:2]
    current_student_data = data_pdf[current_stud + 1]

    subjects = []

    sub1_empty = all(cell_value == " - " for cell_value in current_student_data[4:8])
    sub2_empty = all(cell_value == " - " for cell_value in current_student_data[8:12])
    sub3_empty = all(cell_value == " - " for cell_value in current_student_data[12:16])
    sub4_empty = all(cell_value == " - " for cell_value in current_student_data[16:20])
    sub5_empty = all(cell_value == " - " for cell_value in current_student_data[20:24])
    sub6_empty = all(cell_value == " - " for cell_value in current_student_data[24:28])
    sub7_empty = all(cell_value == " - " for cell_value in current_student_data[28:32])
    sub8_empty = all(cell_value == " - " for cell_value in current_student_data[32:36])

    # Adding tables only if they are not empty
    if not sub1_empty:
        sub7_data = [row[4:8] for row in header] + [current_student_data[4:8]]
        subjects.append(sub7_data)

    if not sub2_empty:
        sub8_data = [row[8:12] for row in header] + [current_student_data[8:12]]
        subjects.append(sub8_data)

    if not sub3_empty:
        sub7_data = [row[12:16] for row in header] + [current_student_data[12:16]]
        subjects.append(sub7_data)

    if not sub4_empty:
        sub8_data = [row[16:20] for row in header] + [current_student_data[16:20]]
        subjects.append(sub8_data)

    if not sub5_empty:
        sub7_data = [row[20:24] for row in header] + [current_student_data[20:24]]
        subjects.append(sub7_data)

    if not sub6_empty:
        sub8_data = [row[24:28] for row in header] + [current_student_data[24:28]]
        subjects.append(sub8_data)

    if not sub7_empty:
        sub7_data = [row[28:32] for row in header] + [current_student_data[28:32]]
        subjects.append(sub7_data)

    if not sub8_empty:
        sub8_data = [row[32:36] for row in header] + [current_student_data[32:36]]
        subjects.append(sub8_data)

    for i in range(len(subjects)):
        table = Table(subjects[i])
        style_report = TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.grey,
                ),  # Set Background for all columns of first row to grey
                (
                    "ALIGN",
                    (0, 0),
                    (-1, 0),
                    "CENTER",
                ),  # Center align all items in first row
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Times-Roman"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("FONTNAME", (0, 1), (-1, -1), "Times-Roman"),
                ("LEFTPADDING", (0, 0), (-1, 0), 5),
                ("SPAN", (0, 0), (3, 0)),
            ]
        )
        table.setStyle(style_report)
        story.append(table)
        story.append(spacer)
    story.append(spacer)

    remarks = f"<u>Remarks</u>: {data_pdf[current_stud + 1][36]}"
    remarks_para = create_paragraph(remarks, normal_style)
    story.append(remarks_para)

    sign_details = "S/D- HOD, Dean Academics & Principal, JSSATEB"
    sign_details_para = create_paragraph(sign_details, top_heading_style)
    story.append(sign_details_para)
    doc.build(story)


# Route to make a data frame, and start the report generation
# Returns a JSON and reports.zip
@app.route("/generate-report", methods=["POST"])
def make_data_frame():
    global usn_end_var, usn_start_var, data_frame, data_pdf, session_id
    post_error = False
    global xlsx_file_path

    if os.path.exists(f"./uploads/{session_id}_reports.xlsx"):
        xlsx_file_path = f"./uploads/{session_id}_reports.xlsx"
    elif os.path.exists(f"./uploads/{session_id}_reports.csv"):
        xlsx_file_path = f"./uploads/{session_id}_reports.csv"
    elif os.path.exists(f"./uploads/{session_id}_reports.xls"):
        xlsx_file_path = f"./uploads/{session_id}_reports.xls"
    else:
        return jsonify(
            {
                "error": "File Not Found",
                "Message": "Reports XLSX has not been Uploaded",
            }
        )

    data_frame = pd.read_excel(xlsx_file_path)
    data_pdf = [data_frame.columns.tolist()] + data_frame.values.tolist()
    data = json.loads(request.data.decode("utf-8"))  # Unpack JSON stringified Data
    usn_start_var = data.get("usnStartVar")
    usn_end_var = data.get("usnEndVar")

    usn_start_integer = 0
    usn_end_integer = 0
    try:
        # Using .iloc[] for integer-based indexing
        usn_start_integer = data_frame.index[
            data_frame["USN"] == usn_start_var
        ].tolist()[0]
        usn_end_integer = data_frame.index[data_frame["USN"] == usn_end_var].tolist()[0]
    except IndexError:
        post_error = True
        return jsonify(
            {"error": "USN Validity Error", "Message": "Please enter Valid USNs"}
        )
    if usn_end_integer < usn_start_integer:
        post_error = True
        return jsonify(
            {
                "error": "USN Format Error",
                "Message": "Please USNs in correct format",
            }
        )
    else:
        if (
            post_error == False
        ):  # Only execute second POST request if there are no POST errors in 1st one
            make_all_pdf(data_frame, usn_start_integer, usn_end_integer, data_pdf)
            zip_filepath = create_zip_reports(session_id)
            return jsonify(
                {
                    "operation": "PDF generated",
                    "Message": "PDFS have been successfully generated",
                    "zip_file_url": zip_filepath,
                }
            )
        else:
            return jsonify(
                {
                    "error": "POST Error",
                    "Message": "Need to fix USN range/Format Errors before generating PDFs",
                }
            )


# Function to Check Allowed Extension
# Returns true if extension is xlsx, csv or xls
def allowed_file(filename):
    allowed_extensions = {"xlsx", "xls", "csv"}
    _, file_extension = os.path.splitext(filename)
    return file_extension.lower()[1:] in allowed_extensions


# Homepage route
# This route will be fetched every time page is reloaded
@app.route("/", methods=["POST", "GET"])
def home():
    global mail_id, mail_pass, receiver_address, usn_end_var, usn_start_var, session_id, xlsx_file_path
    mail_id = ""
    mail_pass = ""
    receiver_address = ""

    # Generate a unique session ID for this device
    session_id = str(uuid.uuid4())

    usn_start_var = ""
    usn_end_var = ""

    xlsx_file_path = ""

    return render_template("index.html", session_id=session_id)


# Route to handle reports xlsx file uploads
@app.route("/upload-file", methods=["POST"])
def upload_file_handler():
    global reports_dir, session_id
    uploaded_file = request.files["file"]
    try:
        if uploaded_file.filename != "" and allowed_file(uploaded_file.filename):
            if uploaded_file.filename.split(".")[0] == "reports":
                # Save the uploaded file with the session ID as part of the filename
                filename = f"{session_id}_{uploaded_file.filename}"
                reports_dir = os.path.join(
                    os.getcwd(), f"reports/Reports - {session_id}"
                )
                if not os.path.exists(reports_dir):
                    os.makedirs(reports_dir)
                uploaded_file.save(os.path.join(uploads_dir, filename))
                # session["xlsx-file-path"] = os.path.join(uploads_dir, filename)
                return jsonify(
                    {
                        "message": "File uploaded successfully.",
                        "filename": filename,
                    }
                )
            else:
                return jsonify({"error": "File should be named 'reports'"})
        elif uploaded_file.filename == "":
            return jsonify({"error": "Select a File"})
        elif not allowed_file(uploaded_file.filename):
            return jsonify({"error": "File should be of type (CSV, XLS or XLSX)"})
    except Exception as e:
        return jsonify({"error": str(e)})


# Function to change mail ID and password for sender
@app.route("/change-mail", methods=["POST"])
def update_mail_credentials():
    data = json.loads(request.data.decode("utf-8"))  # Unpack JSON stringified Data
    new_mail = data.get("newMail")
    new_pass = data.get("newPass")
    if new_mail != "" and new_pass != "":
        # Login
        try:
            session = smtplib.SMTP("smtp.gmail.com", 587)
            session.starttls()
            session.login(new_mail, new_pass)
            session.quit()

            # Update the global mail_id and mail_pass variables
            global mail_id, mail_pass
            mail_id = new_mail
            mail_pass = new_pass

            return jsonify({"operation": "Mail ID and password have been set."})
        except smtplib.SMTPAuthenticationError as e:
            return jsonify({"error": "Incorrect Email/Password. Try Again."})
    else:
        return jsonify(
            {"error": "Empty Field: Please enter Mail ID and Password and try again"}
        )


# Function to create zip file for reports
def create_zip_reports(session_id):
    # Define the path to the user's reports directory
    user_reports_dir = os.path.join(os.getcwd(), f"reports/Reports - {session_id}")

    # Create a temporary directory to store the ZIP file
    temp_dir = os.path.join(os.getcwd(), f"temp_zip/{session_id}")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    # Create a ZIP file containing all the PDF reports
    zip_filename = f"{session_id}_reports.zip"
    zip_filepath = os.path.join(temp_dir, zip_filename)
    with zipfile.ZipFile(zip_filepath, "w", zipfile.ZIP_STORED) as zipf:
        for root, _, files in os.walk(user_reports_dir):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, user_reports_dir))

    return os.path.relpath(zip_filepath)


# Schedule the cleanup task to run every day at 2 AM
def schedule_cleanup_task():
    schedule.every().day.at("02:00").do(clear_folders)


# Every day at 2:00 am perform server cleanup
def start_cleanup_thread():
    while True:
        schedule.run_pending()
        time.sleep(1)


# Start the flask app
if __name__ == "__main__":
    cleanup_thread = threading.Thread(target=start_cleanup_thread)
    cleanup_thread.start()
    app.run(host="0.0.0.0", port=8000)
