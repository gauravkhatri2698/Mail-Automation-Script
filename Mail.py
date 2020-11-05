import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = p.read_excel("Student.xlsx")
email_col = data.get("Email ID")
list_of_emails = list(email_col)
print(list_of_emails)

try:
    server = sm.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("gauravkhatri2698@gmail.com", "madangir@832")
    from_ = "gauravkhatri2698@gmail.com"
    to_ = list_of_emails
    message = MIMEMultipart("alternative")
    message['Subject'] = "This is just texting message"
    message['from'] = "gauravkhatri2698@gmail.com"

    html = '''

    <html>
    <head>

    </head>
    <body>
        <h1> My first automation program</h1>
        <h2> Learning in Progress </h2>
        <p> Paragraph </p>
        <button style="padding:20px; background:green; color:white">Verify</button> 
    </body>

    </html>

    '''
    text = MIMEText(html, "html")

    message.attach(text)

    server.sendmail(from_, to_, message.as_string())
    print("Message has been sent to the emails.")

except Exception as e:
    print(e)
    