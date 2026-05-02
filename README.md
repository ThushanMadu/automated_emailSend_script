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

## Prerequisites

- Python 3.7 or higher
- A Gmail account with 2-Factor Authentication enabled
- Gmail App Password (not your regular Gmail password)

## Setup

### 1. Clone the Repository

```bash
git clone https://github.com/ThushanMadu/automated_emailSend_script.git
cd automated_emailSend_script
```

### 2. Install Dependencies

```bash
pip install pandas openpyxl
```

### 3. Create `.env` File

Copy the example file and fill in your details:

```bash
cp .env.example .env
```

Edit `.env` with your information:

```env
# SMTP Configuration
YOUR_EMAIL=yourgmail@gmail.com
APP_PASSWORD=your16characterapppassword

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

**Important:** The `.env` file is gitignored to protect your credentials and personal data.

### 4. Generate Gmail App Password

1. Go to your Google Account settings
2. Navigate to Security > 2-Step Verification (must be enabled)
3. Go to App Passwords
4. Generate a new app password for "Mail"
5. Use this 16-character password in your `.env` file

### 5. Prepare Your Files

**Important:** The script modifies the Excel file in-place by adding and updating a `Status` column. **Always keep a backup copy of your original spreadsheet before running the script.**

Place these files in the project root:

- **Excel file**: Set `EXCEL_FILE` in `.env` (default: `recipients.xlsx`). Must contain a column named exactly `"Email Address"`
- **Your CV**: Add your CV/resume as a PDF file and set `CV_FILE` in `.env` (default: `your_cv.pdf`)

## Excel File Format

Your Excel file should have at least this column:

| Email Address |
|---------------|
| company1@email.com |
| company2@email.com |

The script automatically adds a `Status` column to track sent emails. **This modifies your original file** - always keep a backup copy before running.

## Usage

Run the script:

```bash
python emails.py
```

The script will:
1. Load all email addresses from the Excel file
2. Send emails one by one via Gmail SMTP
3. Save progress to a JSON checkpoint file after each email
4. Update the Excel file with send status (Sent/Failed) at the end
5. Skip already processed emails on re-runs

## Configuration

All configuration is done via the `.env` file:

| Variable | Description | Required |
|----------|-------------|----------|
| `YOUR_EMAIL` | Gmail address to send from | Yes |
| `APP_PASSWORD` | Gmail App Password | Yes |
| `APPLICANT_NAME` | Your full name | Yes |
| `APPLICANT_PHONE` | Your phone number | Yes |
| `APPLICANT_UNIVERSITY` | Your university name | Yes |
| `APPLICANT_MAJOR` | Your degree/major | Yes |
| `GITHUB_URL` | Link to your GitHub profile | Yes |
| `PORTFOLIO_URL` | Link to your portfolio website | Yes |
| `LINKEDIN_URL` | Link to your LinkedIn profile | Yes |
| `CV_FILE` | Filename of your CV PDF | Yes |
| `EXCEL_FILE` | Filename of your email list | Yes |
| `EMAIL_SUBJECT` | Subject line of the email | Yes |

## Project Structure

```
automated_emailSend_script/
├── emails.py                    # Main script
├── recipients.xlsx              # Recipient email list (add yours)
├── your_cv.pdf                  # Your CV attachment (add yours)
├── .env                         # Credentials (create from .env.example)
├── .env.example                 # Template for environment variables
├── email_checkpoint.json        # Auto-generated progress checkpoint
├── .gitignore
└── README.md
```

## Checkpointing

The script uses a JSON checkpoint file (`email_checkpoint.json`) to track progress. This allows you to safely interrupt the script (CTRL+C) or recover from a crash without resending emails to recipients who have already been contacted. The checkpoint is automatically deleted when all emails are sent successfully.

## Security Notes

- Never commit your `.env` file
- Keep your Gmail App Password secure
- The script uses SMTP SSL (port 465) for encrypted connections
- All personal data is stored in `.env`, not hardcoded in the script

## Troubleshooting

**"Missing YOUR_EMAIL or APP_PASSWORD"**
- Check that your `.env` file exists and contains both variables

**"Missing APPLICANT_NAME"**
- Ensure `APPLICANT_NAME` is set in your `.env` file

**"Column 'Email Address' not found"**
- Ensure your Excel column is named exactly `"Email Address"`

**SMTP authentication errors**
- Verify your App Password is correct
- Ensure 2-Factor Authentication is enabled on your Google account

## License

MIT License - see [LICENSE](LICENSE) file for details

## Author

Thushan Madarasinghe

- GitHub: https://github.com/ThushanMadu
- Portfolio: https://thushanmadu.me
- LinkedIn: https://linkedin.com/in/thushan-madarasinghe-420810222
