# Automated Email Send Script

A Python-based bulk email sending tool that reads recipient addresses from an Excel spreadsheet and sends personalized HTML emails with PDF attachments via Gmail SMTP.

## Features

- Reads email addresses from Excel files
- Sends personalized HTML-formatted emails
- Attaches PDF documents (CV/resume) automatically
- Tracks send status (Sent/Failed) in the Excel file
- Safe to re-run without sending duplicate emails
- Loads credentials and personal data from `.env` file
- JSON-based checkpointing for crash recovery

## Complete Setup Guide (A-Z)

### Step 1: Install Python

1. Download Python from [python.org](https://www.python.org/downloads/)
2. During installation on Windows, check **"Add Python to PATH"**
3. Verify installation:

```bash
python --version
```

### Step 2: Download the Project

1. Download or clone this repository:

```bash
git clone https://github.com/ThushanMadu/automated_emailSend_script.git
cd automated_emailSend_script
```

Or download the ZIP file and extract it to a folder.

### Step 3: Install Dependencies

```bash
pip install pandas openpyxl
```

### Step 4: Enable 2-Step Verification on Gmail

1. Go to [myaccount.google.com](https://myaccount.google.com/)
2. Click **Security** in the left sidebar
3. Under "How you sign in to Google", click **2-Step Verification**
4. Click **Get Started** and follow the setup steps
5. Verify your phone number when prompted
6. Click **Turn On**

### Step 5: Generate Gmail App Password

1. Go to [Google App Passwords](https://myaccount.google.com/apppasswords)
2. Sign in to your Google account if prompted
3. Click **Create** or **Select app**
4. In the app dropdown, select **Mail**
5. In the device dropdown, select **Other** and type a name (e.g., "Bulk Email Script")
6. Click **Generate**
7. Google will show a **16-character password** (e.g., `abcd efgh ijkl mnop`)
8. **Copy this password** - you'll need it in Step 6

**Note:** The password has spaces for readability. You can enter it with or without spaces in the `.env` file.

### Step 6: Configure the `.env` File

1. Copy the example file:

```bash
cp .env.example .env
```

2. Open `.env` in a text editor and fill in all fields:

```env
# SMTP Configuration
YOUR_EMAIL=yourgmail@gmail.com
# Gmail App Password (16-character app password from Google)
# Example format: abcd efgh ijkl mnop
APP_PASSWORD="your_app_password"

# Applicant Personal Information
APPLICANT_NAME="Your Full Name"
APPLICANT_PHONE="+1 000 000 0000"
APPLICANT_UNIVERSITY="Your University Name"
APPLICANT_MAJOR="Your Major/Degree"

# Links
GITHUB_URL="https://github.com/yourusername"
PORTFOLIO_URL="https://yourportfolio.com"
LINKEDIN_URL="https://linkedin.com/in/yourusername"

# File Configuration
CV_FILE="your_cv.pdf"
EXCEL_FILE="recipients.xlsx"

# Email Configuration
EMAIL_SUBJECT="Inquiry Regarding Internship Opportunities"
```

### Step 7: Prepare Your CV

1. Save your CV/resume as a **PDF file**
2. Place it in the project folder
3. Set the filename in `CV_FILE` in your `.env` file (e.g., `CV_FILE="my_cv.pdf"`)

### Step 8: Create the Recipients Excel Sheet

1. Create a new Excel file (`.xlsx`) using Excel, Google Sheets, or LibreOffice

2. In **Cell A1**, type exactly: `Email Address`

3. Add recipient emails below it (one per row):

|   | A |
|---|---|
| **1** | Email Address |
| **2** | hr@company1.com |
| **3** | careers@company2.com |
| **4** | jobs@company3.com |

4. **Important rules:**
   - The column header **must** be exactly `Email Address` (case-sensitive)
   - Do not add a `Status` column - the script creates it automatically
   - Each email should be in its own row
   - Leave empty cells blank - the script will skip them

5. Save the file as `.xlsx` in the project folder

6. Set the filename in `EXCEL_FILE` in your `.env` file (e.g., `EXCEL_FILE="recipients.xlsx"`)

**Backup Warning:** The script modifies your Excel file by adding a `Status` column. **Always keep a backup copy** of your original file before running.

### Step 9: Run the Script

```bash
python emails.py
```

**What to expect:**

- The script will print each email as it sends
- You'll see `✅ Sent successfully` or `❌ Failed to send`
- Progress is saved automatically after each email
- When finished, it prints `Processing complete. Status saved to recipients.xlsx`

### Step 10: Check Results

After running, open your Excel file. You'll see:

|   | A | B |
|---|---|---|
| **1** | Email Address | Status |
| **2** | hr@company1.com | Sent |
| **3** | careers@company2.com | Sent |
| **4** | jobs@company3.com | Failed |

**Re-running the script:**
- Already sent or failed emails are skipped automatically
- You can safely run the script again to continue where you left off
- Fix any failed entries and run again

## Configuration Reference

| Variable | Description | Required | Example |
|----------|-------------|----------|---------|
| `YOUR_EMAIL` | Gmail address to send from | Yes | `you@gmail.com` |
| `APP_PASSWORD` | Gmail App Password (16 chars) | Yes | `abcd efgh ijkl mnop` |
| `APPLICANT_NAME` | Your full name | Yes | `"John Doe"` |
| `APPLICANT_PHONE` | Your phone number | Yes | `"+1 555 123 4567"` |
| `APPLICANT_UNIVERSITY` | Your university name | Yes | `"MIT"` |
| `APPLICANT_MAJOR` | Your degree/major | Yes | `"Computer Science"` |
| `GITHUB_URL` | Link to your GitHub profile | Yes | `"https://github.com/user"` |
| `PORTFOLIO_URL` | Link to your portfolio website | Yes | `"https://yoursite.com"` |
| `LINKEDIN_URL` | Link to your LinkedIn profile | Yes | `"https://linkedin.com/in/user"` |
| `CV_FILE` | Filename of your CV PDF | Yes | `"my_cv.pdf"` |
| `EXCEL_FILE` | Filename of your email list | Yes | `"recipients.xlsx"` |
| `EMAIL_SUBJECT` | Subject line of the email | Yes | `"Internship Inquiry"` |

## Project Structure

```
automated_emailSend_script/
├── emails.py                    # Main script
├── recipients.xlsx              # Recipient email list (add yours)
├── your_cv.pdf                  # Your CV attachment (add yours)
├── .env                         # Your credentials (create from .env.example)
├── .env.example                 # Template for environment variables
├── email_checkpoint.json        # Auto-generated progress file (do not edit)
├── .gitignore
└── README.md
```

## Checkpointing & Recovery

The script uses `email_checkpoint.json` to track progress:

- **Crash recovery:** If the script crashes, just run it again. It resumes where it stopped
- **Safe interruption:** Press `Ctrl+C` to stop. Run again to resume
- **Duplicate prevention:** Already sent emails are never resent
- **Auto-cleanup:** The checkpoint file is deleted after the run completes and the Excel file is updated, even if some rows are marked `Failed`

## Security Notes

- Never share or commit your `.env` file
- Your Gmail App Password gives access to your email account
- The script uses SMTP SSL (port 465) for encrypted connections
- All personal data is stored in `.env`, not hardcoded in the script

## Troubleshooting

**"Missing YOUR_EMAIL or APP_PASSWORD"**
- Check that your `.env` file exists in the project folder
- Ensure both variables are set with no typos

**"Missing APPLICANT_NAME"**
- Ensure `APPLICANT_NAME` is set in your `.env` file

**"Column 'Email Address' not found"**
- The column header must be exactly `Email Address` (case-sensitive)
- Check for extra spaces before or after the header text

**"Error: CV file not found"**
- Make sure your CV PDF is in the project folder
- Verify `CV_FILE` in `.env` matches the actual filename

**SMTP authentication errors**
- Verify your App Password is correct (16 characters)
- Ensure 2-Step Verification is enabled on your Google account
- Check that `YOUR_EMAIL` matches the Gmail account you generated the App Password for

**"Excel file not found"**
- Verify `EXCEL_FILE` in `.env` matches your actual filename
- Make sure the file is in the same folder as `emails.py`

## License

MIT License - see [LICENSE](LICENSE) file for details

## Author

Thushan Madarasinghe

- GitHub: https://github.com/ThushanMadu
- Portfolio: https://thushanmadu.me
- LinkedIn: https://linkedin.com/in/thushan-madarasinghe-420810222
