import pandas as pd
import smtplib
import tkinter as tk
from tkinter import ttk, filedialog
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from ttkthemes import ThemedStyle

# Initialize the global variable email_column
email_column = []

# Function to send an email with an attachment
def send_email(sender_email, sender_password, recipient_email, subject, message, image_path=None, video_path=None):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # Add the message body
    msg.attach(MIMEText(message, 'plain'))

    # Add the image attachment if provided
    if image_path:
        with open(image_path, 'rb') as img_file:
            img = MIMEImage(img_file.read(), name='image.jpg')
        msg.attach(img)

    # Add the video attachment if provided
    if video_path:
        with open(video_path, 'rb') as video_file:
            video = MIMEBase('application', 'octet-stream')
            video.set_payload((video_file).read())
            encoders.encode_base64(video)
            video.add_header('Content-Disposition', 'attachment', filename='Internee-Task(Explaination).mp4')
        msg.attach(video)

    try:
        # Establish a secure connection with the SMTP server
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent successfully to {recipient_email}")

    except Exception as e:
        print(f"An error occurred while sending the email to {recipient_email}: {str(e)}")

    finally:
        server.quit()

# Function to load CSV file and extract email column
def load_csv():
    global email_column  # Declare email_column as global before assignment
    file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])
    if not file_path:
        return

    # Read the CSV file
    data = pd.read_csv(file_path)

    # Extract the email column
    email_column = data['emails'].tolist()  # Convert the email column to a list

# Function to initiate the email sending process
def send_emails():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    subject = subject_entry.get()
    message = message_entry.get('1.0', tk.END)

    # Load CSV and extract emails if not loaded already
    if not email_column:
        load_csv()

    # Iterate over each email and send the message
    for recipient_email in email_column:
        send_email(sender_email, sender_password, recipient_email, subject, message, image_path, video_path)

# Function to upload an image
def upload_image():
    file_path = filedialog.askopenfilename(filetypes=[('Image Files', '*.png;*.jpg;*.jpeg;*.gif')])
    if not file_path:
        return

    # Update the image path in the global variable
    global image_path
    image_path = file_path

# Function to upload a video
def upload_video():
    file_path = filedialog.askopenfilename(filetypes=[('Video Files', '*.mp4;*.avi;*.mkv')])
    if not file_path:
        return

    # Update the video path in the global variable
    global video_path
    video_path = file_path

# Create the main application window
app = tk.Tk()
app.title('Email Automation')

# Use the ThemedStyle to apply the "plastik" theme
style = ThemedStyle(app)
style.set_theme("plastik")

# Create labels and entry widgets for email details
tk.Label(app, text='Sender Email:').grid(row=0, column=0, padx=10, pady=5)
sender_email_entry = ttk.Entry(app)
sender_email_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(app, text='Sender Password:').grid(row=1, column=0, padx=10, pady=5)
sender_password_entry = ttk.Entry(app, show='*')
sender_password_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(app, text='Subject:').grid(row=2, column=0, padx=10, pady=5)
subject_entry = ttk.Entry(app)
subject_entry.grid(row=2, column=1, padx=10, pady=5)

tk.Label(app, text='Message:').grid(row=3, column=0, padx=10, pady=5)
message_entry = tk.Text(app, height=5, width=30)
message_entry.grid(row=3, column=1, padx=10, pady=5)

# Create the "Load CSV" button
load_button = ttk.Button(app, text='Upload CSV', command=load_csv)
load_button.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

# Create the "Upload Image" button
upload_image_button = ttk.Button(app, text='Upload Image', command=upload_image)
upload_image_button.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

# Create the "Upload Video" button
upload_video_button = ttk.Button(app, text='Upload Video', command=upload_video)
upload_video_button.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

# Create the "Send Email" button
send_button = ttk.Button(app, text='Send Emails', command=send_emails)
send_button.grid(row=7, column=0, columnspan=2, padx=10, pady=5)

# Set padding for all widgets
for child in app.winfo_children():
    child.grid_configure(padx=5, pady=5)

# Initialize the image and video path variables
image_path = None
video_path = None

# Run the main event loop
app.mainloop()
