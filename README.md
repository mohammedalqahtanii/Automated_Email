## Automated Email Sender

A Python script that automates sending emails to a list of recipients from an Excel file, with support for attachments and message personalization.

### Features
* Reads recipient data (Name and Email) from an Excel file.
* Attaches a specified PDF file to each email.
* Automatically personalizes the message body for each recipient using their name.
* Includes a counter to track the progress of sent emails.
* Uses a random delay between messages to avoid spam filters.
* Validates rows and skips any with missing names or emails.

  ### How to Use:
  
    **1. Install Dependencies:**
          *pip install pandas
          *pip install openpyxl

  **2. Prepare Files:**

        * Create a `config.py` file in the same directory with your credentials:
          ```python
          # config.py
          EMAIL_ADDRESS = "your_email@gmail.com"
          EMAIL_PASSWORD = "your_app_password"
          ```
      * Ensure an `Emails_info.xlsx` file is in the same directory. It must contain two columns: `Company_Name` and `Email`.
      * Place the PDF file you want to attach in the same directory.

      
**3. Run the Script:**
    *python main.py



### Tech Stack
* Python
* Pandas
* Smtplib
  
