# Imports
# For noise surpression
from audoai.noise_removal import NoiseRemovalClient

# For Speech Recognition, Speaker Diarization and Summary
import assemblyai as aai

# For PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# For Automated Emails
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os.path import basename

# Nosie Surpression
noise_removal = NoiseRemovalClient(api_key='API_KEY')
result = noise_removal.process('test_audio.mp3')
result.save('meeting-clean.mp3')

# Speech Recognition and Summary
aai.settings.api_key = "API_KEY"
num_attendees = 1 # Optional to increase accuracy of speaker diarization

config = aai.TranscriptionConfig(
  summarization = True,
  summary_model=aai.SummarizationModel.informative,
  summary_type=aai.SummarizationType.bullets,
  speaker_labels = True,
  speakers_expected = num_attendees
)

transcriber = aai.Transcriber()

# Speech Recognition & Conversion to Text
transcript = transcriber.transcribe("test_audio.mp3",config=config)

print(transcript.summary)

# Generate Meeting Summary and Store in PDF
def wrap_text(text, max_width, canvas):
    all_lines = []
    # Split the text by bullet points
    parts = text.split('-')
    for part in parts:
        if part:  # Check if the part is not empty
            words = part.split()
            current_line = ''
            for word in words:
                test_line = f"{current_line} {word}".strip()
                if canvas.stringWidth(test_line, 'Helvetica', 10) <= max_width:
                    current_line = test_line
                else:
                    all_lines.append(current_line)
                    current_line = word
            all_lines.append(current_line)  # Add the last line of the part
            all_lines.append("")  # Add an empty line to represent the bullet point break
    return all_lines

c = canvas.Canvas("Meeting_Summary.pdf", pagesize=letter)
width, height = letter
left_margin = 60
top_margin = 60
max_width = width - 2 * left_margin
line_height = 15

# Set bold font for the heading
c.setFont("Helvetica-Bold", 14)
c.drawString(left_margin, height - top_margin, "Meeting Summary")
y_position = height - top_margin - 2 * line_height  # Adjust starting position after heading

# Set regular font for the body text
c.setFont("Helvetica", 10)

lines = wrap_text(transcript.summary, max_width, c)
for line in lines:
    if line == "":  # Special handling for bullet point breaks
        y_position -= line_height * 0.5  # Adjust spacing for bullet points
        continue
    if y_position < line_height + top_margin:
        c.showPage()
        y_position = height - top_margin
    c.drawString(left_margin, y_position, line)
    y_position -= line_height

c.save()

# Speaker Diarization & PDF Generation
def wrap_text(text, max_width, canvas, y_position):
    # Split the text into words
    words = text.split()
    wrapped_text = ""
    line = ""
    for word in words:
        # Check if adding the next word exceeds the line width
        if canvas.stringWidth(line + word, "Helvetica", 10) < max_width:
            line += word + " "
        else:
            # If the line is too wide, wrap to the next line
            wrapped_text += line + "\n"
            line = word + " "  # Start a new line with the current word
    
    wrapped_text += line  # Add the last line
    return wrapped_text


c = canvas.Canvas("Meeting_Transcript.pdf", pagesize=letter)
width, height = letter

# Define margins
left_margin = 30
right_margin = width - 30
max_width = right_margin - left_margin

# Starting Y position, and the step for each new line
y_position = height - 50
line_height = 15

# Title or Header
c.setFont("Helvetica-Bold", 12)
c.drawString(left_margin, y_position, "Meeting Transcript")
y_position -= 2 * line_height

# Set font for the body text
c.setFont("Helvetica", 10)

# Iterate through the transcript utterances and add them to the PDF
for utterance in transcript.utterances:
    text = f"Speaker {utterance.speaker}: {utterance.text}"
    wrapped_text = wrap_text(text, max_width, c, y_position)
    for line in wrapped_text.split('\n'):
        c.drawString(left_margin, y_position, line)
        y_position -= line_height
        # Move to next page if there's not enough space
        if y_position < 50:
            c.showPage()
            y_position = height - 50
            c.setFont("Helvetica", 10)

c.save()

# Send Automated Emails
def send_email_with_pdfs(recipients, subject, body, pdf_paths, email_user, email_password, smtp_server="smtp.gmail.com", smtp_port=587):
    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = ", ".join(recipients)  # Join all recipient emails with a comma
    msg['Subject'] = subject
    
    # Add body to email
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach each PDF file
    for pdf_path in pdf_paths:
        with open(pdf_path, "rb") as file:
            part = MIMEApplication(
                file.read(),
                Name=basename(pdf_path)
            )
        part['Content-Disposition'] = f'attachment; filename="{basename(pdf_path)}"'
        msg.attach(part)
    
    # Log in to server and send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(email_user, email_password)
        server.send_message(msg)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Usage example
email_user = "meetly.ai.automated@gmail.com"  # Your email
email_password = "pssd wtcj krjr tuds"  # Your email password or app password
recipients = ["hgrover1904@gmail.com"]  # List of recipients
subject = "Meeting Summaries"
body = "Please find attached the meeting summary PDFs."
pdf_paths = ["Meeting_Summary.pdf", "Meeting_Transcript.pdf"]  # List of PDF file paths to attach

send_email_with_pdfs(recipients, subject, body, pdf_paths, email_user, email_password)