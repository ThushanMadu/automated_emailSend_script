import os
import json
import smtplib
from email.message import EmailMessage

import pandas as pd


def load_env_file(env_path):
    """Load simple KEY=VALUE pairs from a .env file into os.environ."""
    if not os.path.exists(env_path):
        return

    with open(env_path, "r", encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


load_env_file(os.path.join(os.path.dirname(__file__), ".env"))

# === CONFIGURATION ===
YOUR_EMAIL = os.environ.get("YOUR_EMAIL", "")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

if not YOUR_EMAIL or not APP_PASSWORD:
    raise ValueError(
         "Missing YOUR_EMAIL or APP_PASSWORD environment variables. "
         "Set them in the process environment or in the local .env file."
    )

CV_FILE = "your_cv.pdf"
EXCEL_FILE = "test.xlsx"
CHECKPOINT_FILE = "email_checkpoint.json"

SUBJECT = " Inquiry Regarding Internship Opportunities – Computer Science Undergraduate"

HTML_BODY = f"""
<html>
  <body style="font-family: Arial, sans-serif; line-height: 1.6;">
    <p>Dear Sir/Madam,</p>

    <p>Hope you're having a good week.</p>

    <p>My name is Thushan Madarasinghe, a Computer Science undergraduate at the Informatics Institute of Technology (IIT), affiliated with the University of Westminster. I'm writing to express my strong interest in an internship opportunity with your esteemed organization.</p>

    <p>I'm genuinely impressed by your company's commitment to excellence and nurturing new talent. I'm eager to contribute to the impactful work you're doing, which feels like a fantastic next step for me.</p>

    <p>My studies have provided a solid foundation in Object-Oriented Programming, Algorithms, and Data Structures. I'm proficient in technologies including Java, React Native, Node.js, Express.js, React, MongoDB, SQL, and HTML/CSS, and I'm comfortable with Git for version control. I also have a good understanding of backend development, including APIs and database management.</p>

    <p>I enjoy tackling new technologies and figuring things out across the full stack, bringing different components together to build functional solutions.</p>

    <p>Key projects I've worked on include:</p>
    <ul>
      <li><strong>GoviShakthi:</strong> An AI-powered MERN stack app with LLM integration for product recommendations.</li>
      <li><strong>FinTrack:</strong> A personal finance tracker built with the MERN stack.</li>
      <li><strong>Real-Time Ticketing System:</strong> Developed using Node.js, React.js, and WebSockets.</li>
      <li><strong>Plane Management System:</strong> A project built using Java.</li>
    </ul>

    <p>My involvement with the IEEE Computer Society at university has also enhanced my communication, organization, and teamwork skills through various events.</p>

    <p>I'm eager to bring my energy, technical skills, and passion for learning to your team. Please find my CV attached for your review. I would be grateful for the chance to discuss how I could contribute. If there aren't any suitable openings right now, I'd be thankful if you'd keep my application in mind and let me know about any future opportunities that might come up.</p>

    <p>Thank you for your time and consideration.</p>

    <p>Sincerely,</p>

    <p><strong>Thushan Madarasinghe</strong><br>
    +94 70 392 1791<br>
    <a href="mailto:{YOUR_EMAIL}">{YOUR_EMAIL}</a><br>
    <a href="https://github.com/ThushanMadu">GitHub</a> |
    <a href="https://thushanmadu.me">Portfolio</a> |
    <a href="https://linkedin.com/in/thushan-madarasinghe-420810222">LinkedIn</a>
    </p>
  </body>
</html>
"""


def load_checkpoint():
    """Load progress from JSON checkpoint file."""
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}


def save_checkpoint(checkpoint):
    """Save progress to JSON checkpoint file."""
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(checkpoint, f, indent=2)


def send_email(to_email):
    """Sends an email to a single recipient."""
    msg = EmailMessage()
    msg['Subject'] = SUBJECT
    msg['From'] = YOUR_EMAIL
    msg['To'] = to_email

    msg.add_alternative(HTML_BODY, subtype='html')

    if not os.path.exists(CV_FILE):
        print(f"Error: CV file not found at {CV_FILE}")
        return False
    with open(CV_FILE, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(CV_FILE))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(YOUR_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"SMTP Error: {e}")
        return False


if __name__ == "__main__":
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: Excel file not found at {EXCEL_FILE}")
    else:
        try:
            data = pd.read_excel(EXCEL_FILE)

            if 'Status' not in data.columns:
                data['Status'] = ''

            if "Email Address" not in data.columns:
                print(f"Error: Column 'Email Address' not found in {EXCEL_FILE}. Please check the column name.")
            else:
                checkpoint = load_checkpoint()
                print(f"Loaded {len(data)} rows from {EXCEL_FILE}")
                
                if checkpoint:
                    print(f"Resuming from checkpoint ({len(checkpoint)} entries)")
                    for idx_str, status in checkpoint.items():
                        data.loc[int(idx_str), 'Status'] = status

                for index, row in data.iterrows():
                    idx_str = str(index)
                    email = row["Email Address"]
                    
                    if idx_str in checkpoint:
                        print(f"Skipping row {index} ({email}): Already in checkpoint.")
                        continue
                        
                    if pd.notna(email) and str(email).strip() != "":
                        if row['Status'] in ['Sent', 'Failed']:
                            print(f"Skipping row {index} ({email}): Status already '{row['Status']}'.")
                            continue
                            
                        email_str = str(email).strip()
                        print(f"Attempting to send to: {email_str}")
                        if send_email(email_str):
                            print(f"✅ Sent successfully to {email_str}")
                            data.loc[index, 'Status'] = 'Sent'
                            checkpoint[idx_str] = 'Sent'
                        else:
                            print(f"❌ Failed to send to {email_str}")
                            data.loc[index, 'Status'] = 'Failed'
                            checkpoint[idx_str] = 'Failed'
                        save_checkpoint(checkpoint)
                    else:
                        print(f"Skipping row {index}: No valid email address found.")

                try:
                    data.to_excel(EXCEL_FILE, index=False)
                    print(f"Processing complete. Status saved to {EXCEL_FILE}")
                    if os.path.exists(CHECKPOINT_FILE):
                        os.remove(CHECKPOINT_FILE)
                except Exception as e:
                    print(f"❌ Error saving Excel file: {e}")
                    print(f"Progress preserved in {CHECKPOINT_FILE}. Run again to resume.")

        except Exception as e:
            print(f"An error occurred while processing the Excel file: {e}")
