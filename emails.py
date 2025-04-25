import smtplib
import pandas as pd
from email.message import EmailMessage

# === CONFIGURATION ===
YOUR_EMAIL = "thushan.dev03@gmail.com"           # <- your Gmail
APP_PASSWORD = "vmnoemngkaytoyji"   # <- your Gmail app password
CV_FILE = "Thushan_Madarasinghe.pdf"       # <- your CV filename
EXCEL_FILE = "test.xlsx"  # <- Excel file with emails
COVER_LETTER = "Thushan_Madarasinghe_coverLetter.pdf"  # Ensure the correct file extension

# Load emails from Excel
data = pd.read_excel(EXCEL_FILE)
print(data.columns)

def send_email(to_email, recipient_name):
    msg = EmailMessage()
    msg['Subject'] = "Internship Opportunity – Computer Science Undergraduate | Full-Stack Development & AI"
    msg['From'] = YOUR_EMAIL
    msg['To'] = to_email

    # Plain-text fallback
    msg.set_content("This is a fallback message for email clients that do not support HTML.")

    # HTML Email Body with clickable buttons
    greeting = "Dear Sir/Madam,"
    html_body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>{greeting}</p>

        <p>
          I hope this message finds you well.<br><br>
          My name is <strong>Thushan Madarasinghe</strong>, and I am a second-year Computer Science undergraduate at the University of Westminster (IIT Sri Lanka). I am currently seeking internship opportunities where I can apply my skills in full-stack development, mobile app development, and AI-driven solutions to contribute meaningfully to innovative projects.
        </p>

        <p>
          I am particularly excited by opportunities to work in dynamic, technology-driven environments where impactful software solutions are built using modern engineering practices.
        </p>

        <p><strong>Some of my key projects include:</strong></p>
        <ul>
          <li><strong>Real-Time Ticketing System</strong> (Node.js, React.js, WebSockets): Built a scalable event ticketing platform using producer-consumer patterns and real-time data updates.</li>
          <li><strong>GoviShakthi</strong> – AI-Powered Product Recommendation App (MERN Stack, LLM-based Analysis): Leading the mobile app development and contributing to backend API integration to support farmers with intelligent market insights.</li>
          <li><strong>FinTrack</strong> – Personal Finance Tracker (MERN Stack): Developed a full-stack finance management application featuring secure authentication, data visualization, and real-time transaction tracking.</li>
          <li><strong>Plane Management System</strong> (Java): Built a robust flight seat reservation console application with search, ticketing, and reporting functionalities.</li>
        </ul>

        <p>
          I am proficient in technologies such as Java, JavaScript, React Native, React.js, Node.js, Express.js, MongoDB, MySQL and GIT. With leadership experience from university projects and extracurricular activities such as IEEE Computer Society, Richmond College IT Society and Richmond Live , I have honed strong teamwork and project management skills.
        </p>

        <p>
          I am eager to learn, adapt, and contribute to your organization’s success through my passion for building user-centered, high-quality solutions.
        </p>

        <p>
          Thank you for considering my application. I look forward to the possibility of contributing to your team.
        </p>

        <p>
          Best regards,<br>
          <strong>Thushan Madarasinghe</strong><br>
          📞 +94 70 392 1791<br>
          ✉️ thushan.dev03@gmail.com
        </p>

        <p>
          <a href="https://github.com/ThushanMadu" style="text-decoration:none;">
            <button style="padding: 8px 14px; margin: 5px; background-color: #333; color: white; border: none; border-radius: 5px;">GitHub</button>
          </a>
          <a href="https://thushanmadu.me" style="text-decoration:none;">
            <button style="padding: 8px 14px; margin: 5px; background-color: #007acc; color: white; border: none; border-radius: 5px;">Portfolio</button>
          </a>
          <a href="https://linkedin.com/in/thushan-madarasinghe-420810222" style="text-decoration:none;">
            <button style="padding: 8px 14px; margin: 5px; background-color: #0a66c2; color: white; border: none; border-radius: 5px;">LinkedIn</button>
          </a>
        </p>
      </body>
    </html>
    """

    msg.add_alternative(html_body, subtype='html')

    # Attach CV
    with open(CV_FILE, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=CV_FILE)

    # Attach Cover Letter
    with open(COVER_LETTER, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=f"{COVER_LETTER}.pdf")

    # Send the message
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(YOUR_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)

# === Sending loop ===
for index, row in data.iterrows():
    email = row["Email Address"]  # Replace "email" with the correct column name
    try:
        send_email(email, "Sir/Madam")
        print(f"✅ Sent to {email}")
    except Exception as e:
        print(f"❌ Failed to send to {email}: {e}")